/**
 * -------------------------------------------------------------------------------------------
 * Copyright (c) Microsoft Corporation.  All Rights Reserved.  Licensed under the MIT License.
 * See License in the project root for license information.
 * -------------------------------------------------------------------------------------------
 */

import { BetaGraph, IGraph, Providers, createFromProvider, error, log } from '@microsoft/mgt-element';
import { HubConnection, HubConnectionBuilder, IHttpConnectionOptions, LogLevel } from '@microsoft/signalr';
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
import { Timer } from '../utils/Timer';
import { getOrGenerateGroupId } from './getOrGenerateGroupId';
import { v4 as uuid } from 'uuid';

export const appSettings = {
  defaultSubscriptionLifetimeInMinutes: 10,
  renewalThreshold: 75, // The number of seconds before subscription expires it will be renewed
  renewalTimerInterval: 20, // The number of seconds between executions of the renewal timer
  removalThreshold: 1 * 60, // number of seconds after the last update of a subscription to consider in inactive
  removalTimerInterval: 1 * 60, // the number of seconds between executions of the timer to remove inactive subscriptions
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

export class GraphNotificationClient {
  private readonly instanceId = uuid();
  private connection?: HubConnection = undefined;
  private renewalTimeout?: string;
  private cleanupTimeout?: string;
  private renewalCount = 0;
  private chatId = '';
  private sessionId = '';
  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();
  private readonly timer = new Timer();
  private get graph() {
    return this._graph;
  }
  private get beta() {
    return this._graph ? BetaGraph.fromGraph(this._graph) : undefined;
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
    private readonly _graph: IGraph | undefined
  ) {
    // start the cleanup timer when we create the notification client.
    this.startCleanupTimer();
  }

  /**
   * Removes any active timers that may exist to prevent memory leaks and perf issues.
   * Call this method when the component that depends an instance of this class is being removed from the DOM
   */
  public tearDown() {
    log('cleaning up graph notification resources');
    if (this.cleanupTimeout) this.timer.clearTimeout(this.cleanupTimeout);
    if (this.renewalTimeout) this.timer.clearTimeout(this.renewalTimeout);
    this.timer.close();
  }

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  // TODO: understand if this is needed under the native model
  private readonly onReconnect = (connectionId: string | undefined) => {
    log(`Reconnected. ConnectionId: ${connectionId || 'undefined'}`);
    // void this.renewChatSubscriptions();
  };

  private readonly receiveNotificationMessage = (message: string) => {
    if (typeof message !== 'string') throw new Error('Expected string from receivenotificationmessageasync');

    const notification: ReceivedNotification = JSON.parse(message) as ReceivedNotification;
    log('received notification message', notification);
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
    const ackMessage: unknown = { StatusCode: '200' };
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

  private readonly cacheSubscription = async (subscriptionRecord: Subscription): Promise<void> => {
    log(subscriptionRecord);

    await this.subscriptionCache.cacheSubscription(this.chatId, ComponentType.Chat, subscriptionRecord);

    // only start timer once. undefined for renewalInterval is semaphore it has stopped.
    if (this.renewalTimeout === undefined) this.startRenewalTimer();
  };

  private async subscribeToResource(resourcePath: string, changeTypes: ChangeTypes[]) {
    try {
      // build subscription request
      const expirationDateTime = new Date(
        new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
      ).toISOString();
      const subscriptionDefinition: Subscription = {
        changeType: changeTypes.join(','),
        notificationUrl: `${GraphConfig.webSocketsPrefix}?groupId=${getOrGenerateGroupId(this.chatId)}`,
        resource: resourcePath,
        expirationDateTime,
        includeResourceData: true,
        clientState: 'wsssecret'
      };

      log('subscribing to changes for ' + resourcePath);
      const subscriptionEndpoint = GraphConfig.subscriptionEndpoint;
      const subscriptionGraph = this.subscriptionGraph;
      if (!subscriptionGraph) return;
      // send subscription POST to Graph
      const subscription: Subscription = (await subscriptionGraph
        .api(subscriptionEndpoint)
        .post(subscriptionDefinition)) as Subscription;
      if (!subscription?.notificationUrl) throw new Error('Subscription not created');
      log(subscription);

      const awaits: Promise<void>[] = [];
      // Cache the subscription in storage for re-hydration on page refreshes
      awaits.push(this.cacheSubscription(subscription));

      // create a connection to the web socket if one does not exist
      if (!this.connection) awaits.push(this.createSignalRConnection(subscription.notificationUrl));

      log('Invoked CreateSubscription');
      return Promise.all(awaits);
    } catch (e) {
      // this catches other unhandled exceptions such as when subscription.post fails
      return Promise.reject(e);
    }
  }

  private readonly startRenewalTimer = () => {
    if (this.renewalTimeout) this.timer.clearTimeout(this.renewalTimeout);
    this.renewalTimeout = this.timer.setTimeout(
      'renewal:' + this.instanceId,
      this.syncRenewalTimerWrapper,
      appSettings.renewalTimerInterval * 1000
    );
    log(`Start renewal timer . Id: ${this.renewalTimeout}`);
  };

  private readonly syncRenewalTimerWrapper = () => void this.renewalTimer();

  private readonly renewalTimer = async () => {
    try {
      log(`running subscription renewal timer for chatId: ${this.chatId} sessionId: ${this.sessionId}`);
      const subscriptions = (await this.subscriptionCache.loadSubscriptions(this.chatId))?.subscriptions || [];
      if (subscriptions.length === 0) {
        log(`No subscriptions found in session state. Stop renewal timer ${this.renewalTimeout}.`);
        if (this.renewalTimeout) this.timer.clearTimeout(this.renewalTimeout);
        return;
      }

      for (const subscription of subscriptions) {
        if (!subscription.expirationDateTime) continue;
        const expirationTime = new Date(subscription.expirationDateTime);
        const now = new Date();
        const diff = Math.round((expirationTime.getTime() - now.getTime()) / 1000);

        if (diff <= appSettings.renewalThreshold) {
          this.renewalCount++;
          log(`Renewing Graph subscription. RenewalCount: ${this.renewalCount}`);
          // stop interval to prevent new invokes until refresh is ready.
          if (this.renewalTimeout) this.timer.clearTimeout(this.renewalTimeout);
          this.renewalTimeout = undefined;
          await this.renewChatSubscriptions();
          // There is one subscription that need expiration, all subscriptions will be renewed
          break;
        }
      }
    } catch (e) {
      error(e);
    }
    this.renewalTimeout = this.timer.setTimeout(
      'renewal:' + this.instanceId,
      this.syncRenewalTimerWrapper,
      appSettings.renewalTimerInterval * 1000
    );
  };

  public renewChatSubscriptions = async () => {
    const expirationTime = new Date(
      new Date().getTime() + appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
    );

    const subscriptionCache = await this.subscriptionCache.loadSubscriptions(this.chatId);
    const awaits: Promise<unknown>[] = [];
    for (const subscription of subscriptionCache?.subscriptions || []) {
      if (!subscription.id) continue;
      // the renewSubscription method caches the updated subscription to track the new expiration time
      awaits.push(this.renewSubscription(subscription.id, expirationTime.toISOString()));
      log(`Invoked RenewSubscription ${subscription.id}`);
    }
    await Promise.all(awaits);
    if (!this.renewalTimeout) {
      this.renewalTimeout = this.timer.setTimeout(
        'renewal:' + this.instanceId,
        this.syncRenewalTimerWrapper,
        appSettings.renewalTimerInterval * 1000
      );
    }
  };

  public renewSubscription = async (subscriptionId: string, expirationDateTime: string): Promise<void> => {
    // PATCH /subscriptions/{id}
    try {
      const renewedSubscription = (await this.graph
        ?.api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`)
        .patch({
          expirationDateTime
        })) as Subscription | undefined;
      if (renewedSubscription) return this.cacheSubscription(renewedSubscription);
    } catch (e) {
      return Promise.reject(e);
    }
  };

  public async createSignalRConnection(notificationUrl: string) {
    const connectionOptions: IHttpConnectionOptions = {
      accessTokenFactory: this.getToken,
      withCredentials: false
    };
    const connection = new HubConnectionBuilder()
      .withUrl(GraphConfig.adjustNotificationUrl(notificationUrl, this.sessionId), connectionOptions)
      .withAutomaticReconnect()
      .configureLogging(LogLevel.Information)
      .build();

    connection.onreconnected(this.onReconnect);

    connection.on('receivenotificationmessageasync', this.receiveNotificationMessage);

    connection.on('EchoMessage', log);

    this.connection = connection;
    try {
      await connection.start();
      log(connection);
    } catch (e) {
      error('An error occurred connecting to the notification web socket', e);
    }
  }

  private async deleteSubscription(id: string) {
    try {
      await this.graph?.api(`${GraphConfig.subscriptionEndpoint}/${id}`).delete();
    } catch (e) {
      error(e);
    }
  }

  private async removeSubscriptions(subscriptions: Subscription[]): Promise<unknown[]> {
    const tasks: Promise<unknown>[] = [];
    for (const s of subscriptions) {
      // if there is no id or the subscription is expired, skip
      if (!s.id || (s.expirationDateTime && new Date(s.expirationDateTime) <= new Date())) continue;
      tasks.push(this.deleteSubscription(s.id));
    }
    return Promise.all(tasks);
  }

  private startCleanupTimer() {
    this.cleanupTimeout = this.timer.setTimeout(
      'cleanup:' + this.instanceId,
      this.cleanupTimerSync,
      appSettings.removalTimerInterval * 1000
    );
  }

  private readonly cleanupTimerSync = () => {
    void this.cleanupTimer();
  };

  private readonly cleanupTimer = async () => {
    try {
      log(`running cleanup timer`);

      // ensure there is a valid user
      const id = await Providers.getCacheId();
      if (!id) {
        return;
      }

      // get the inactive subs (requires a user since the db is per user)
      const offset = Math.min(
        appSettings.removalThreshold * 1000,
        appSettings.defaultSubscriptionLifetimeInMinutes * 60 * 1000
      );
      const threshold = new Date(new Date().getTime() - offset).toISOString();
      const inactiveSubs = await this.subscriptionCache.loadInactiveSubscriptions(threshold, ComponentType.Chat);

      // remove all subscriptions
      let tasks: Promise<unknown>[] = [];
      for (const inactive of inactiveSubs) {
        tasks.push(this.removeSubscriptions(inactive.subscriptions));
      }
      await Promise.all(tasks);

      // delete the cache entries
      tasks = [];
      for (const inactive of inactiveSubs) {
        tasks.push(this.subscriptionCache.deleteCachedSubscriptions(inactive.chatId));
      }
    } catch (e) {
      // EXAMPLE: a user does a sign-out so `loadInactiveSubscriptions` no longer has a valid user and results
      //    in an exception.
      error(e);
    } finally {
      this.startCleanupTimer();
    }
  };

  public async closeSignalRConnection() {
    // stop the connection and set it to undefined so it will reconnect when next subscription is created.
    await this.connection?.stop();
    this.connection = undefined;
  }

  public async subscribeToChatNotifications(chatId: string, sessionId: string) {
    this.chatId = chatId;
    this.sessionId = sessionId;
    // MGT uses a per-user cache, so no concerns of loading the cached data for another user.
    const cacheData = await this.subscriptionCache.loadSubscriptions(chatId);
    if (cacheData) {
      // check subscription validity & renew if all still valid otherwise recreate
      const someExpired = cacheData.subscriptions.some(
        s => s.expirationDateTime && new Date(s.expirationDateTime) <= new Date()
      );
      // for a given user + app + chatId + sessionId they only get one websocket and receive all notifications via that websocket.
      const webSocketUrl = cacheData.subscriptions.find(s => s.notificationUrl)?.notificationUrl;
      if (!someExpired && webSocketUrl) {
        // if we have a websocket url and all the subscriptions are valid, we can reuse the websocket and return before recreating subscriptions.
        await this.createSignalRConnection(webSocketUrl);
        await this.renewChatSubscriptions();
        return;
      } else if (someExpired) {
        // if some are expired, remove them and continue to recreate the subscription
        await this.removeSubscriptions(cacheData.subscriptions);
      }
      await this.subscriptionCache.deleteCachedSubscriptions(chatId);
    }
    const promises: Promise<unknown>[] = [];
    promises.push(this.subscribeToResource(`/chats/${chatId}/messages`, ['created', 'updated', 'deleted']));
    promises.push(this.subscribeToResource(`/chats/${chatId}/members`, ['created', 'deleted']));
    promises.push(this.subscribeToResource(`/chats/${chatId}`, ['updated', 'deleted']));
    await Promise.all(promises).catch((e: Error) => {
      this?.emitter.graphNotificationClientError(e);
    });
  }
}
