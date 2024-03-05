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
import { Timer } from '../utils/Timer';
import { getOrGenerateGroupId } from './getOrGenerateGroupId';
import { v4 as uuid } from 'uuid';

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

  private readonly subscriptionCache: SubscriptionsCache = new SubscriptionsCache();
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

  // private cacheToken = '';
  // private readonly getToken = async () => {
  //   if (this.cacheToken) {
  //     return this.cacheToken;
  //   }

  //   const token = await Providers.globalProvider.getAccessToken();
  //   if (!token) throw new Error('Could not retrieve token for user');

  //   const response = await fetch(`http://localhost:5201/token`, {
  //     method: 'GET',
  //     headers: {
  //       Authorization: `Bearer ${token}`,
  //       'Content-Type': 'application/json'
  //     }
  //   });

  //   if (!response.ok) {
  //     throw new Error(`HttpClient error: ${response.statusText}`);
  //   }

  //   this.cacheToken = await response.text();

  //   return this.cacheToken;
  //   // return token;
  // };

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

  private async createSubscription(userId: string): Promise<Subscription> {
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
      resource: resourcePath + '?model=B',
      expirationDateTime,
      includeResourceData: true,
      clientState: 'wsssecret'
    };

    log('subscribing to changes for ' + resourcePath);
    // const subscriptionEndpoint = GraphConfig.subscriptionEndpoint;
    // const subscription: Subscription = (await this.subscriptionGraph
    //   .api(subscriptionEndpoint)
    //   .post(subscriptionDefinition)) as Subscription;
    // send subscription POST to Graph
    // const statictoken =
    // 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IlkwNkg5N2NwUU90YWdlVTZvc0NTR0RneXljTzZUY2xmUVpFOEljT05nT3MiLCJhbGciOiJSUzI1NiIsIng1dCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSIsImtpZCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iOGZhOTk0My02ZTUwLTRkOWItYTgyMi02NjUzYWM2YTM1NGMvIiwiaWF0IjoxNzA5MjM3OTg3LCJuYmYiOjE3MDkyMzc5ODcsImV4cCI6MTcwOTI0MjU0NiwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhXQUFBQTRKdlZRc3lmYUNIeFNNODZTSjc1bEI1TDhSVDZSbENOVFFqZ0ZleUw3VngrSUVWUlZNWHV0OEdPY2pLRjdwR1IwZEJJVE9CVnlscDE0THNSTHlIV3hzeUloZDV4dWVTczMwb2NjdU41WGRBPSIsImFtciI6WyJwd2QiLCJyc2EiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiTUdUIENIQVQgTVMgR3JhcGggQVBJIiwiYXBwaWQiOiI1ZWYwMWZiMS1mYzAxLTQ5OTktYTkwZS0yNGRlMjFmMmFkMmYiLCJhcHBpZGFjciI6IjEiLCJkZXZpY2VpZCI6ImM3YTcwNTUzLTEyZjEtNDViYy05NWM0LTRjNTlhY2Q1MTgxYyIsImZhbWlseV9uYW1lIjoiQWRtaW4iLCJnaXZlbl9uYW1lIjoiRGxlZSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjcwLjExNS45OC4yIiwibmFtZSI6IkRMRUUgTVMgQWRtaW4iLCJvaWQiOiI4YzUyM2E2Ni01ZjFlLTQyMGUtOWRiNi1jNGQ4MmFjZDY2NzAiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDExODg1MEY2OCIsInJoIjoiMC5BWDBBUTVuNnVGQnVtMDJvSW1aVHJHbzFUQU1BQUFBQUFBQUF3QUFBQUFBQUFBQjlBTjQuIiwic2NwIjoiQm9va21hcmsuUmVhZC5BbGwgQ2FsZW5kYXJzLlJlYWQgQ2hhbm5lbC5SZWFkQmFzaWMuQWxsIENoYXQuQ3JlYXRlIENoYXQuUmVhZCBDaGF0LlJlYWRCYXNpYyBDaGF0LlJlYWRXcml0ZSBDaGF0LlJlYWRXcml0ZS5BbGwgQ2hhdE1lbWJlci5SZWFkV3JpdGUgQ2hhdE1lc3NhZ2UuU2VuZCBDb250YWN0cy5SZWFkIEV4dGVybmFsSXRlbS5SZWFkLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkIE1haWwuUmVhZEJhc2ljIG9wZW5pZCBQZW9wbGUuUmVhZCBQZW9wbGUuUmVhZC5BbGwgUHJlc2VuY2UuUmVhZC5BbGwgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBUZWFtLlJlYWRCYXNpYy5BbGwgVGVhbXNBcHBJbnN0YWxsYXRpb24uUmVhZEZvckNoYXQgVGVybVN0b3JlLlJlYWQuQWxsIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZEJhc2ljLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJpbmtub3dubnR3ayIsImttc2kiXSwic3ViIjoiQzlBa0VQTGxjYUN2QTZINklnMWhiMjFZeE9QRlRPSlh4SnNqc3E1bzR2NCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImI4ZmE5OTQzLTZlNTAtNGQ5Yi1hODIyLTY2NTNhYzZhMzU0YyIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AZGxlZW1zLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImFkbWluQGRsZWVtcy5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiI2MmNQUXhSdGkwcVlkNkZYd2lCV0FBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiLCJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX2FwcF9iaWxsX2lkIjoiNTdhMDgzMTktNmE5NS00M2QzLThlMzUtMzRhY2I4NDM2YTMzIiwieG1zX2FwcF9iaWxsX3N0YXRlIjoxLCJ4bXNfc3QiOnsic3ViIjoibXByU1BGcFRSSUptZ2tHTmJjR0Nob1QzYUpfNmoyRVlNWmFYV1hLempNdyJ9LCJ4bXNfdGNkdCI6MTYxMzczOTk5NX0.sGfNw0ElV9WpzIt4i7iY4ig0GSxy1FBTz9Vr_l4nth2IU8GdHK8ypL4nYNt1rNzg0nUe6YdqQW_iS9EGoELUB8MZE0KhVs0jVv3RX6QVSOJ3OUF2ALiKmgzZwh3meYVXas1RhR01tQbxzgCtjNvpDml2lQ4r76T-jY-1BV3IaNsMzB9sdhmFlF9c_uWLiFkY8P8HbrmaNZzkt_SzpWC3dFXHfZUytolOhVl6TT1Gb0jmfeqRqAF_l-MTIt_jrPHkk1T7s31vybcUOc7C-zrEPWAmJfiArbk5w5Jpjwi2ZPa8mVrErhMXjl0MBbwdy1ZfOjiH-KJBcZStk76lm3JtNA';

    // const token = await this.getToken();
    // const response = await fetch(`https://graph.microsoft.com/beta${subscriptionEndpoint}`, {
    //   method: 'POST',
    //   headers: {
    //     Authorization: `Bearer ${token}`,
    //     'Content-Type': 'application/json'
    //   },
    //   body: JSON.stringify(subscriptionDefinition)
    // });

    // if (!response.ok) {
    //   const wwwAuth = response.headers.get('Www-Authenticate');
    //   if (wwwAuth) {
    //     const parts = wwwAuth.split(' ')[1].split(',');
    //     log(parts);
    //   }
    //   throw new Error(`HttpClient error: ${response.statusText}`);
    // }

    const token = await Providers.globalProvider.getAccessTokenForScopes(
      'api://5ef01fb1-fc01-4999-a90e-24de21f2ad2f/access_as_user'
    );
    const response = await fetch(`http://localhost:5201/subscriptions`, {
      method: 'POST',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json'
      },
      body: JSON.stringify(subscriptionDefinition)
    });

    if (!response.ok) {
      throw new Error(`HttpClient error: ${response.statusText}`);
    }

    const subscription = (await response.json()) as Subscription;
    if (!subscription?.notificationUrl) throw new Error('Subscription not created');
    log(subscription);

    this.subscriptionId = subscription.id!;
    await this.cacheSubscription(userId, subscription);

    log('Subscription created.');

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

  private trySwitchToConnected() {
    if (!this.wasConnected) {
      log('The user can now receive notifications from the user subscription.');
      this.wasConnected = true;
      this.emitter?.connected();
    }
  }

  private trySwitchToDisconnected() {
    if (this.wasConnected) {
      log('The user can now receive notifications from the user subscription.');
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

      // if there is a current subscription...
      let subscription = await this.getSubscription(currentUserId);
      if (subscription) {
        // attempt a renewal if necessary
        try {
          const expirationTime = new Date(subscription.expirationDateTime!);
          const diff = Math.round((expirationTime.getTime() - new Date().getTime()) / 1000);
          if (diff <= 0) {
            log(`Renewing user subscription ${subscription.id!} that has already expired...`);
            this.trySwitchToDisconnected();
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
          this.trySwitchToDisconnected();
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

      // create or reconnect the SignalR connection
      // notificationUrl comes in the form of websockets:https://graph.microsoft.com/beta/subscriptions/notificationChannel/websockets/<Id>?groupid=<UserId>&sessionid=default
      // if <Id> changes, we need to create a new connection
      if (this.connection?.state === HubConnectionState.Connected) {
        await this.connection?.send('ping'); // ensure the connection is still alive
      }
      if (!this.connection) {
        log(`Creating a new SignalR connection for subscription ${subscription.id!}...`);
        this.trySwitchToDisconnected();
        this.lastNotificationUrl = subscription.notificationUrl!;
        await this.createSignalRConnection(subscription.notificationUrl!);
        log(`Successfully created a new SignalR connection for subscription ${subscription.id!}.`);
      } else if (this.connection.state !== HubConnectionState.Connected) {
        log(`Reconnecting SignalR connection for subscription ${subscription.id!}...`);
        this.trySwitchToDisconnected();
        await this.connection.start();
        log(`Successfully reconnected SignalR connection for subscription ${subscription.id!}.`);
      } else if (this.lastNotificationUrl !== subscription.notificationUrl) {
        log(`Updating SignalR connection for subscription ${subscription.id!} due to new notification URL...`);
        this.trySwitchToDisconnected();
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
    const renewedSubscription = (await this.graph.api(`${GraphConfig.subscriptionEndpoint}/${subscriptionId}`).patch({
      expirationDateTime
    })) as Subscription;
    return this.cacheSubscription(userId, renewedSubscription);
  };

  private readonly getToken = async () => {
    const token = await Providers.globalProvider.getAccessToken();
    if (!token) throw new Error('Could not retrieve token for user');
    return token;
  };

  private async createSignalRConnection(notificationUrl: string) {
    // const statictoken =
    // 'eyJ0eXAiOiJKV1QiLCJub25jZSI6IlkwNkg5N2NwUU90YWdlVTZvc0NTR0RneXljTzZUY2xmUVpFOEljT05nT3MiLCJhbGciOiJSUzI1NiIsIng1dCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSIsImtpZCI6IlhSdmtvOFA3QTNVYVdTblU3Yk05blQwTWpoQSJ9.eyJhdWQiOiIwMDAwMDAwMy0wMDAwLTAwMDAtYzAwMC0wMDAwMDAwMDAwMDAiLCJpc3MiOiJodHRwczovL3N0cy53aW5kb3dzLm5ldC9iOGZhOTk0My02ZTUwLTRkOWItYTgyMi02NjUzYWM2YTM1NGMvIiwiaWF0IjoxNzA5MjM3OTg3LCJuYmYiOjE3MDkyMzc5ODcsImV4cCI6MTcwOTI0MjU0NiwiYWNjdCI6MCwiYWNyIjoiMSIsImFpbyI6IkFWUUFxLzhXQUFBQTRKdlZRc3lmYUNIeFNNODZTSjc1bEI1TDhSVDZSbENOVFFqZ0ZleUw3VngrSUVWUlZNWHV0OEdPY2pLRjdwR1IwZEJJVE9CVnlscDE0THNSTHlIV3hzeUloZDV4dWVTczMwb2NjdU41WGRBPSIsImFtciI6WyJwd2QiLCJyc2EiLCJtZmEiXSwiYXBwX2Rpc3BsYXluYW1lIjoiTUdUIENIQVQgTVMgR3JhcGggQVBJIiwiYXBwaWQiOiI1ZWYwMWZiMS1mYzAxLTQ5OTktYTkwZS0yNGRlMjFmMmFkMmYiLCJhcHBpZGFjciI6IjEiLCJkZXZpY2VpZCI6ImM3YTcwNTUzLTEyZjEtNDViYy05NWM0LTRjNTlhY2Q1MTgxYyIsImZhbWlseV9uYW1lIjoiQWRtaW4iLCJnaXZlbl9uYW1lIjoiRGxlZSIsImlkdHlwIjoidXNlciIsImlwYWRkciI6IjcwLjExNS45OC4yIiwibmFtZSI6IkRMRUUgTVMgQWRtaW4iLCJvaWQiOiI4YzUyM2E2Ni01ZjFlLTQyMGUtOWRiNi1jNGQ4MmFjZDY2NzAiLCJwbGF0ZiI6IjMiLCJwdWlkIjoiMTAwMzIwMDExODg1MEY2OCIsInJoIjoiMC5BWDBBUTVuNnVGQnVtMDJvSW1aVHJHbzFUQU1BQUFBQUFBQUF3QUFBQUFBQUFBQjlBTjQuIiwic2NwIjoiQm9va21hcmsuUmVhZC5BbGwgQ2FsZW5kYXJzLlJlYWQgQ2hhbm5lbC5SZWFkQmFzaWMuQWxsIENoYXQuQ3JlYXRlIENoYXQuUmVhZCBDaGF0LlJlYWRCYXNpYyBDaGF0LlJlYWRXcml0ZSBDaGF0LlJlYWRXcml0ZS5BbGwgQ2hhdE1lbWJlci5SZWFkV3JpdGUgQ2hhdE1lc3NhZ2UuU2VuZCBDb250YWN0cy5SZWFkIEV4dGVybmFsSXRlbS5SZWFkLkFsbCBGaWxlcy5SZWFkV3JpdGUuQWxsIEdyb3VwLlJlYWRXcml0ZS5BbGwgTWFpbC5SZWFkIE1haWwuUmVhZEJhc2ljIG9wZW5pZCBQZW9wbGUuUmVhZCBQZW9wbGUuUmVhZC5BbGwgUHJlc2VuY2UuUmVhZC5BbGwgcHJvZmlsZSBTaXRlcy5SZWFkLkFsbCBTaXRlcy5SZWFkV3JpdGUuQWxsIFRhc2tzLlJlYWRXcml0ZSBUZWFtLlJlYWRCYXNpYy5BbGwgVGVhbXNBcHBJbnN0YWxsYXRpb24uUmVhZEZvckNoYXQgVGVybVN0b3JlLlJlYWQuQWxsIFVzZXIuUmVhZCBVc2VyLlJlYWQuQWxsIFVzZXIuUmVhZEJhc2ljLkFsbCBlbWFpbCIsInNpZ25pbl9zdGF0ZSI6WyJpbmtub3dubnR3ayIsImttc2kiXSwic3ViIjoiQzlBa0VQTGxjYUN2QTZINklnMWhiMjFZeE9QRlRPSlh4SnNqc3E1bzR2NCIsInRlbmFudF9yZWdpb25fc2NvcGUiOiJOQSIsInRpZCI6ImI4ZmE5OTQzLTZlNTAtNGQ5Yi1hODIyLTY2NTNhYzZhMzU0YyIsInVuaXF1ZV9uYW1lIjoiYWRtaW5AZGxlZW1zLm9ubWljcm9zb2Z0LmNvbSIsInVwbiI6ImFkbWluQGRsZWVtcy5vbm1pY3Jvc29mdC5jb20iLCJ1dGkiOiI2MmNQUXhSdGkwcVlkNkZYd2lCV0FBIiwidmVyIjoiMS4wIiwid2lkcyI6WyI2MmU5MDM5NC02OWY1LTQyMzctOTE5MC0wMTIxNzcxNDVlMTAiLCJiNzlmYmY0ZC0zZWY5LTQ2ODktODE0My03NmIxOTRlODU1MDkiXSwieG1zX2FwcF9iaWxsX2lkIjoiNTdhMDgzMTktNmE5NS00M2QzLThlMzUtMzRhY2I4NDM2YTMzIiwieG1zX2FwcF9iaWxsX3N0YXRlIjoxLCJ4bXNfc3QiOnsic3ViIjoibXByU1BGcFRSSUptZ2tHTmJjR0Nob1QzYUpfNmoyRVlNWmFYV1hLempNdyJ9LCJ4bXNfdGNkdCI6MTYxMzczOTk5NX0.sGfNw0ElV9WpzIt4i7iY4ig0GSxy1FBTz9Vr_l4nth2IU8GdHK8ypL4nYNt1rNzg0nUe6YdqQW_iS9EGoELUB8MZE0KhVs0jVv3RX6QVSOJ3OUF2ALiKmgzZwh3meYVXas1RhR01tQbxzgCtjNvpDml2lQ4r76T-jY-1BV3IaNsMzB9sdhmFlF9c_uWLiFkY8P8HbrmaNZzkt_SzpWC3dFXHfZUytolOhVl6TT1Gb0jmfeqRqAF_l-MTIt_jrPHkk1T7s31vybcUOc7C-zrEPWAmJfiArbk5w5Jpjwi2ZPa8mVrErhMXjl0MBbwdy1ZfOjiH-KJBcZStk76lm3JtNA';
    const connectionOptions: IHttpConnectionOptions = {
      accessTokenFactory: this.getToken,
      withCredentials: false
    };

    const connection = new HubConnectionBuilder()
      .withUrl(GraphConfig.adjustNotificationUrl(notificationUrl, this.getSessionId()), connectionOptions)
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
