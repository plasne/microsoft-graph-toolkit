/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BetaGraph, IGraph, Providers, createFromProvider, error, log } from '@microsoft/mgt-element';
import {
  HubConnection,
  HubConnectionBuilder,
  HubConnectionState,
  IHttpConnectionOptions,
  LogLevel
} from '@microsoft/signalr';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import type {
  Entity,
  Subscription,
  ChatMessage,
  Chat,
  AadUserConversationMember
} from '@microsoft/microsoft-graph-types';
import { GraphConfig } from './GraphConfig';
import { SubscriptionsCache, ComponentType } from './Caching/SubscriptionCache';
import { ProxySubscriptionCache } from './Caching/ProxySubscriptionCache';
import { Timer } from '../utils/Timer';
import { getOrGenerateGroupId } from './getOrGenerateGroupId';
import { v4 as uuid } from 'uuid';
import { MGTProxyOperations, ProxySubscription, RenewedProxySubscription } from './MGTProxyOperations';
import { MGTProxyTokenManager } from './MGTProxyTokenManager';

export const appSettings = {
  defaultSubscriptionLifetimeInMinutes: 10,
  renewalThreshold: 75, // The number of seconds before subscription expires it will be renewed
  renewalTimerInterval: 3, // The number of seconds between executions of the renewal timer
  useCanary: GraphConfig.useCanary
};

type ChangeTypes = 'created' | 'updated' | 'deleted';

interface Notification<T extends Entity> {
  subscriptionId: string;
  changeType: ChangeTypes;
  resource: string;
  resourceData: T & {
    id: string;
    '@odata.type': string;
    '@odata.id': string;
  };
  EncryptedContent: string;
}

type ReceivedNotification = Notification<Chat> | Notification<ChatMessage> | Notification<AadUserConversationMember>;

const isMessageNotification = (o: Notification<Entity>): o is Notification<ChatMessage> =>
  o.resource.includes('/messages(');
const isMembershipNotification = (o: Notification<Entity>): o is Notification<AadUserConversationMember> =>
  o.resource.includes('/members');

export class GraphNotificationUserClient {
  private readonly instanceId = uuid();
  private connection?: HubConnection = undefined;
  private renewalTimeout?: string;
  private renewalCount = 0;
  private wasConnected?: boolean | undefined;
  private userId = '';
  private lastNotificationUrl = '';
  private subscriptionId = '';

  private readonly proxyTokenManager: MGTProxyTokenManager = new MGTProxyTokenManager();
  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();
  private readonly proxySubscriptionCache: ProxySubscriptionCache = new ProxySubscriptionCache();
  private readonly timer = new Timer();
  private get graph() {
    return this._graph;
  }
  private get beta() {
    return BetaGraph.fromGraph(this._graph);
  }
  private get subscriptionGraph() {
    return GraphConfig.useCanary
      ? createFromProvider(Providers.globalProvider, GraphConfig.canarySubscriptionVersion, 'mgt-chat')
      : this.beta;
  }

  /**
   *
   */
  constructor(
    private readonly emitter: ThreadEventEmitter,
    private readonly _graph: IGraph
  ) {}

  /**
   * Removes any active timers that may exist to prevent memory leaks and perf issues.
   * Call this method when the component that depends an instance of this class is being removed from the DOM
   * i.e
   */
  public tearDown() {
    log('cleaning up graph user notification resources');
    if (this.renewalTimeout) this.timer.clearTimeout(this.renewalTimeout);
    this.timer.close();
  }

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  private readonly receiveNotificationMessage = (message: string) => {
    if (typeof message !== 'string') throw new Error('Expected string from receivenotificationmessageasync');

    const ackMessage: unknown = { StatusCode: '200' };
    const notification: ReceivedNotification = JSON.parse(message) as ReceivedNotification;
    // only process notifications for the current subscription
    if (this.subscriptionId && this.subscriptionId !== notification.subscriptionId) {
      log('Received notification for a different subscription', notification);
      return ackMessage;
    }

    log('received user notification message', notification);
    const emitter: ThreadEventEmitter | undefined = this.emitter;
    if (!notification.resourceData) throw new Error('Message did not contain resourceData');
    if (isMessageNotification(notification)) {
      this.processMessageNotification(notification, emitter);
    } else if (isMembershipNotification(notification)) {
      this.processMembershipNotification(notification, emitter);
    } else {
      this.processChatPropertiesNotification(notification, emitter);
    }
    // Need to return a status code string of 200 so that graph knows the message was received and doesn't re-send the notification
    return GraphConfig.ackAsString ? JSON.stringify(ackMessage) : ackMessage;
  };

  private processMessageNotification(notification: Notification<ChatMessage>, emitter: ThreadEventEmitter | undefined) {
    const message = notification.resourceData;

    switch (notification.changeType) {
      case 'created':
        emitter?.chatMessageReceived(message);
        return;
      case 'updated':
        emitter?.chatMessageEdited(message);
        return;
      case 'deleted':
        emitter?.chatMessageDeleted(message);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private processMembershipNotification(
    notification: Notification<AadUserConversationMember>,
    emitter: ThreadEventEmitter | undefined
  ) {
    const member = notification.resourceData;
    switch (notification.changeType) {
      case 'created':
        emitter?.participantAdded(member);
        return;
      case 'deleted':
        emitter?.participantRemoved(member);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private processChatPropertiesNotification(notification: Notification<Chat>, emitter: ThreadEventEmitter | undefined) {
    const chat = notification.resourceData;
    switch (notification.changeType) {
      case 'updated':
        emitter?.chatThreadPropertiesUpdated(chat);
        return;
      case 'deleted':
        emitter?.chatThreadDeleted(chat);
        return;
      default:
        throw new Error('Unknown change type');
    }
  }

  private readonly cacheSubscription = async (userId: string, subscriptionRecord: Subscription): Promise<void> => {
    log(subscriptionRecord);
    await this.subscriptionCache.cacheSubscription(userId, ComponentType.User, subscriptionRecord);
  };

  private readonly cacheProxySubscription = async (
    userId: string,
    proxySubscriptionRecord: ProxySubscription
  ): Promise<void> => {
    log(proxySubscriptionRecord);
    await this.proxySubscriptionCache.cacheProxySubscription(userId, ComponentType.User, proxySubscriptionRecord);
  };

  private proxySubscription: ProxySubscription | undefined;

  private async createSubscription(userId: string): Promise<Subscription | undefined> {
    const groupId = getOrGenerateGroupId(userId);
    log('Creating a new subscription with group Id:', groupId);
    const resourcePath = `/users/${userId}/chats/getAllmessages`;
    const changeTypes: ChangeTypes[] = ['created', 'updated', 'deleted'];

    // build subscription request
    const expirationDateTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    ).toISOString();
    const subscriptionDefinition: Subscription = {
      changeType: changeTypes.join(','),
      notificationUrl: `${GraphConfig.webSocketsPrefix}?groupId=${groupId}`,
      resource: resourcePath,
      expirationDateTime,
      includeResourceData: true,
      clientState: 'wsssecret'
    };
    log('subscribing to changes for ' + resourcePath);

    if (Providers.globalProvider.isWebProxyEnabled) {
      let proxySubscription = await this.getProxySubscription(this.userId);
      if (proxySubscription) {
        this.proxySubscription = proxySubscription;
      } else {
        proxySubscription = await this.createSubscriptionFromProxy(subscriptionDefinition);
        this.proxySubscription = proxySubscription;
      }

      if (this.proxySubscription?.subscription) {
        this.subscriptionId = this.proxySubscription.subscription.id!;
        await this.cacheProxySubscription(this.userId, this.proxySubscription);
      }
      return this.proxySubscription?.subscription;
    } else {
      const subscription = await this.createSubscriptionFromGraph(subscriptionDefinition);
      this.subscriptionId = subscription.id!;
      await this.cacheSubscription(this.userId, subscription);
      return subscription;
    }
  }

  private async performOperationTokenSafe(
    url: string,
    method: string,
    operationData: Subscription,
    accesstoken: string
  ): Promise<ProxySubscription | undefined> {
    let result = await MGTProxyOperations.PerformOperation(url, method, operationData, accesstoken);
    if (result === undefined) {
      // undefined is the code for 401, so we need to get a new token and try again
      const token = await this.proxyTokenManager.getProxyToken(true);
      result = await MGTProxyOperations.PerformOperation(url, method, operationData, token);
    }

    if (!this.isProxySubscriptionType(result)) {
      throw new Error('Failed to create/renew subscription');
    }

    return result;
  }

  private isProxySubscriptionType(obj: unknown): obj is ProxySubscription {
    return typeof obj === 'object' && obj !== null && 'subscription' in obj;
  }

  private async createSubscriptionFromProxy(
    subscriptionDefinition: Subscription
  ): Promise<ProxySubscription | undefined> {
    const token = await this.proxyTokenManager.getProxyToken();
    const proxySubscription: ProxySubscription | undefined = await this.performOperationTokenSafe(
      Providers.globalProvider.webProxyURL + GraphConfig.subscriptionEndpoint,
      'POST',
      subscriptionDefinition,
      token
    );
    log('Subscription created using web proxy.');
    return proxySubscription;
  }

  private async createSubscriptionFromGraph(subscriptionDefinition: Subscription): Promise<Subscription> {
    const subscriptionEndpoint = GraphConfig.subscriptionEndpoint;
    // send subscription POST to Graph
    const subscription: Subscription = (await this.subscriptionGraph
      .api(subscriptionEndpoint)
      .post(subscriptionDefinition)) as Subscription;
    if (!subscription?.notificationUrl) throw new Error('Subscription not created');
    log('Subscription created using graph api.');
    return subscription;
  }

  private async deleteCachedSubscriptions(userId: string) {
    try {
      log('Removing all user subscriptions from cache...');
      await this.subscriptionCache.deleteCachedSubscriptions(userId);
      this.subscriptionId = '';
      log('Successfully removed all user subscriptions from cache.');
    } catch (e) {
      error('Failed to remove all user subscriptions from cache.', e);
    }
  }

  private async deleteCachedProxySubscriptions(userId: string) {
    try {
      log('Removing all user subscriptions from cache...');
      await this.proxySubscriptionCache.deleteCachedProxySubscriptions(userId);
      this.subscriptionId = '';
      log('Successfully removed all proxy user subscriptions from cache.');
    } catch (e) {
      error('Failed to remove all proxy user subscriptions from cache.', e);
    }
  }

  private trySwitchToConnected() {
    if (this.wasConnected !== true) {
      log('The user will receive notifications from the user subscription.');
      this.wasConnected = true;
      this.emitter?.connected();
    }
  }

  private trySwitchToDisconnected(ignoreIfUndefined = false) {
    if (ignoreIfUndefined && this.wasConnected === undefined) return;
    if (this.wasConnected !== false) {
      log('The user will NOT receive notifications from the user subscription.');
      this.wasConnected = false;
      this.emitter?.disconnected();
    }
  }

  private readonly renewalSync = () => {
    void this.renewal();
  };

  private readonly renewal = async () => {
    let nextRenewalTimeInSec = appSettings.renewalTimerInterval;
    try {
      const currentUserId = this.userId;

      // if there is a current subscription or a webproxy suscription...
      let subscription;
      if (Providers.globalProvider.isWebProxyEnabled) {
        this.proxySubscription = await this.getProxySubscription(currentUserId);
        subscription = this.proxySubscription?.subscription;
      } else {
        subscription = await this.getSubscription(currentUserId);
      }
      if (subscription) {
        // attempt a renewal if necessary
        try {
          const expirationTime = new Date(subscription.expirationDateTime!);
          const diff = Math.round((expirationTime.getTime() - new Date().getTime()) / 1000);
          if (diff <= 0) {
            log(`Renewing user subscription ${subscription.id!} that has already expired...`);
            this.trySwitchToDisconnected(true);
            await this.renewSubscription(currentUserId, subscription);
            log(`Successfully renewed user subscription ${subscription.id!}.`);
          } else if (diff <= appSettings.renewalThreshold) {
            log(`Renewing user subscription ${subscription.id!} that will expire in ${diff} seconds...`);
            await this.renewSubscription(currentUserId, subscription);
            log(`Successfully renewed user subscription ${subscription.id!}.`);
          }
        } catch (e) {
          error(`Failed to renew user subscription ${subscription.id!}.`, e);
          await this.deleteCachedSubscriptions(currentUserId);
          subscription = undefined;
        }
      }

      // if there is no subscription, try to create one
      if (!subscription) {
        try {
          this.trySwitchToDisconnected(true);
          subscription = await this.createSubscription(currentUserId);
        } catch (e) {
          const err = e as { statusCode?: number; message: string };
          if (err.statusCode === 403 && err.message.indexOf('has reached its limit') > 0) {
            // if the limit is reached, back-off (NOTE: this should probably be a 429)
            nextRenewalTimeInSec = appSettings.renewalTimerInterval * 3;
            throw new Error(
              `Failed to create a new subscription due to a limitation; retrying in ${nextRenewalTimeInSec} seconds: ${err.message}.`
            );
          } else if (err.statusCode === 403 || err.statusCode === 402) {
            // permanent error, stop renewal
            error('Failed to create a new subscription due to a permanent condition; stopping renewals.', e);
            return; // exit without setting the next renewal timer
          } else {
            // transient error, retry
            throw new Error(
              `Failed to create a new subscription due to a transient condition; retrying in ${nextRenewalTimeInSec} seconds: ${err.message}.`
            );
          }
        }
      }

      if (!subscription) {
        throw new Error('Subscription not created');
      }

      // create or reconnect the SignalR connection
      // notificationUrl comes in the form of websockets:https://graph.microsoft.com/beta/subscriptions/notificationChannel/websockets/<Id>?groupid=<UserId>&sessionid=default
      // if <Id> changes, we need to create a new connection
      if (this.connection?.state === HubConnectionState.Connected) {
        await this.connection?.send('ping'); // ensure the connection is still alive
      }
      if (!this.connection) {
        log(`Creating a new SignalR connection for subscription ${subscription.id!}...`);
        this.trySwitchToDisconnected(true);
        this.lastNotificationUrl = subscription.notificationUrl!;
        await this.createSignalRConnection(subscription.notificationUrl!);
        log(`Successfully created a new SignalR connection for subscription ${subscription.id!}.`);
      } else if (this.connection.state !== HubConnectionState.Connected) {
        log(`Reconnecting SignalR connection for subscription ${subscription.id!}...`);
        this.trySwitchToDisconnected(true);
        await this.connection.start();
        log(`Successfully reconnected SignalR connection for subscription ${subscription.id!}.`);
      } else if (this.lastNotificationUrl !== subscription.notificationUrl) {
        log(`Updating SignalR connection for subscription ${subscription.id!} due to new notification URL...`);
        this.trySwitchToDisconnected(true);
        await this.closeSignalRConnection();
        this.lastNotificationUrl = subscription.notificationUrl!;
        await this.createSignalRConnection(subscription.notificationUrl!);
        log(`Successfully updated SignalR connection for subscription ${subscription.id!}.`);
      }

      // emit the new connection event if necessary
      this.trySwitchToConnected();
    } catch (e) {
      error('Error in user subscription connection process.', e);
      this.trySwitchToDisconnected();
    }
    this.renewalTimeout = this.timer.setTimeout(
      'renewal:' + this.instanceId,
      this.renewalSync,
      nextRenewalTimeInSec * 1000
    );
  };

  private async getSubscription(userId: string): Promise<Subscription | undefined> {
    const subscriptions = (await this.subscriptionCache.loadSubscriptions(userId))?.subscriptions || [];
    return subscriptions.length > 0 ? subscriptions[0] : undefined;
  }

  private async getProxySubscription(userId: string): Promise<ProxySubscription | undefined> {
    const proxySubscriptions =
      (await this.proxySubscriptionCache.loadProxySubscriptions(userId))?.proxySubscriptions || [];
    return proxySubscriptions.length > 0 ? proxySubscriptions[0] : undefined;
  }

  // this is used to create a unique session id for the web socket connection
  private getSessionId(): string {
    return uuid();
  }

  private readonly renewSubscription = async (userId: string, subscription: Subscription): Promise<void> => {
    this.renewalCount++;
    log(`Renewing Graph subscription for ChatList. RenewalCount: ${this.renewalCount}.`);

    const newExpirationTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );
    // PATCH /subscriptions/{id}
    const subscriptionId = subscription.id;
    const expirationDateTime = newExpirationTime.toISOString();

    if (Providers.globalProvider.isWebProxyEnabled) {
      const renewedSubscription = (await this.performOperationTokenSafe(
        `${Providers.globalProvider.webProxyURL}/subscriptions/${subscriptionId}`,
        'PATCH',
        { expirationDateTime },
        await this.proxyTokenManager.getProxyToken()
      )) as RenewedProxySubscription;
      if (this.proxySubscription && renewedSubscription) {
        this.proxySubscription.subscription = renewedSubscription.subscription!;
        await this.cacheProxySubscription(this.userId, this.proxySubscription);
      }
    } else {
      const renewedSubscription = (await this.graph.api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`).patch({
        expirationDateTime
      })) as Subscription;
      return this.cacheSubscription(userId, renewedSubscription);
    }
  };

  private readonly getAccessTokenForSignalRConnection = async () => {
    if (Providers.globalProvider.isWebProxyEnabled) {
      return this.proxySubscription?.negotiate?.accessToken ?? '';
    }
    return this.getToken();
  };

  private getNotificationUrlForSignalRConnection(notificationUrl: string) {
    if (Providers.globalProvider.isWebProxyEnabled) {
      return this.proxySubscription?.negotiate?.url ?? '';
    }
    return notificationUrl;
  }

  private async createSignalRConnection(notificationUrl: string) {
    const connectionOptions: IHttpConnectionOptions = {
      accessTokenFactory: () => this.getAccessTokenForSignalRConnection(),
      withCredentials: false
    };

    const connection = new HubConnectionBuilder()
      .withUrl(
        GraphConfig.adjustNotificationUrl(
          this.getNotificationUrlForSignalRConnection(notificationUrl),
          this.getSessionId()
        ),
        connectionOptions
      )
      .configureLogging(LogLevel.Information)
      .build();

    connection.on('receivenotificationmessageasync', this.receiveNotificationMessage);
    connection.on('EchoMessage', log);

    this.connection = connection;
    await connection.start();
  }

  public async closeSignalRConnection() {
    // stop the connection and set it to undefined so it will reconnect when next subscription is created.
    this.trySwitchToDisconnected();
    try {
      await this.connection?.stop();
    } catch (e) {
      error('Error closing a prior SignalR connection.', e);
    }
    this.connection = undefined;
  }

  public subscribeToUserNotifications(userId: string) {
    log(`User subscription with id: ${userId}`);
    this.wasConnected = undefined;
    this.userId = userId;
    this.renewalTimeout = this.timer.setTimeout('renewal:' + this.instanceId, this.renewalSync, 0);
  }
}
