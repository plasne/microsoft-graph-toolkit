import { MessageThreadProps, ErrorBarProps, Message } from '@azure/communication-react';
import { ActiveAccountChanged, IGraph, LoginChangedEvent, ProviderState, Providers } from '@microsoft/mgt-element';
import { GraphError } from '@microsoft/microsoft-graph-client';
import {
  ChatMessage,
  ChatRenamedEventMessageDetail,
  MembersAddedEventMessageDetail,
  MembersDeletedEventMessageDetail
} from '@microsoft/microsoft-graph-types';
import { produce } from 'immer';
import { currentUserId } from '../utils/currentUser';
import { graph } from '../utils/graph';
// TODO: MessageCache is added here for the purpose of following the convention of StatefulGraphChatClient. However, StatefulGraphChatListClient
//       is also leveraging the same cache and performing the same actions which would have resulted in race conditions against the same messages
//       in the cache. To avoid this, I have commented out the code. We should revisit this and determine if we need to use the cache.
// import { MessageCache } from './Caching/MessageCache';
import { GraphConfig } from './GraphConfig';
import { GraphNotificationUserClient } from './GraphNotificationUserClient';
import { ThreadEventEmitter } from './ThreadEventEmitter';
import { ChatThreadCollection, loadChat, loadChatThreads, loadChatThreadsByPage } from './graph.chat';
import { ChatMessageInfo, Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { error } from '@microsoft/mgt-element';
interface ODataType {
  '@odata.type': MessageEventType;
}
type MembersAddedEventDetail = ODataType &
  MembersAddedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.membersAddedEventMessageDetail';
  };
type MembersRemovedEventDetail = ODataType &
  MembersDeletedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.membersDeletedEventMessageDetail';
  };
type ChatRenamedEventDetail = ODataType &
  ChatRenamedEventMessageDetail & {
    '@odata.type': '#microsoft.graph.chatRenamedEventMessageDetail';
  };

type ChatMessageEvents = MembersAddedEventDetail | MembersRemovedEventDetail | ChatRenamedEventDetail;

// defines the type of the state object returned from the StatefulGraphChatListClient
export type GraphChatListClient = Pick<MessageThreadProps, 'userId'> & {
  status:
    | 'initial'
    | 'creating server connections'
    | 'subscribing to notifications'
    | 'loading messages'
    | 'no session id'
    | 'no messages'
    | 'ready'
    | 'error';
  chatThreads: GraphChat[];
  nextLink: string;
} & Pick<ErrorBarProps, 'activeErrorMessages'>;

interface StatefulClient<T> {
  /**
   * Get the current state of the client
   */
  getState(): T;
  /**
   * Register a callback to receive state updates
   *
   * @param handler Callback to receive state updates
   */
  onStateChange(handler: (state: T) => void): void;
  /**
   * Remove a callback from receiving state updates
   *
   * @param handler Callback to be unregistered
   */
  offStateChange(handler: (state: T) => void): void;
  /**
   * Register a callback to receive ChatList events
   *
   * @param handler Callback to receive ChatList events
   */
  onChatListEvent(handler: (event: ChatListEvent) => void): void;
  /**
   * Remove a callback to receive ChatList events
   *
   * @param handler Callback to be unregistered
   */
  offChatListEvent(handler: (event: ChatListEvent) => void): void;

  chatThreadsPerPage: number;

  loadMoreChatThreads(): void;
}

type MessageEventType =
  | '#microsoft.graph.membersAddedEventMessageDetail'
  | '#microsoft.graph.membersDeletedEventMessageDetail'
  | '#microsoft.graph.chatRenamedEventMessageDetail';

/**
 * Extended Message type with additional properties.
 */
export type GraphChatMessage = Message & {
  hasUnsupportedContent: boolean;
  rawChatUrl: string;
};

interface EventMessageDetail {
  chatDisplayName: string;
}

export interface ChatListEvent {
  type:
    | 'chatMessageReceived'
    | 'chatMessageDeleted'
    | 'chatMessageEdited'
    | 'chatRenamed'
    | 'memberAdded'
    | 'memberRemoved'
    | 'systemEvent';
  message: ChatMessage;
}

class StatefulGraphChatListClient implements StatefulClient<GraphChatListClient> {
  private readonly _notificationClient: GraphNotificationUserClient;
  private readonly _eventEmitter: ThreadEventEmitter;
  // private readonly _cache: MessageCache;
  private _stateSubscribers: ((state: GraphChatListClient) => void)[] = [];
  private _chatListEventSubscribers: ((state: ChatListEvent) => void)[] = [];
  private readonly _graph: IGraph;
  constructor(chatThreadsPerPage: number) {
    this.updateUserInfo();
    Providers.globalProvider.onStateChanged(this.onLoginStateChanged);
    Providers.globalProvider.onActiveAccountChanged(this.onActiveAccountChanged);
    this._eventEmitter = new ThreadEventEmitter();
    this.registerEventListeners();
    // this._cache = new MessageCache();
    this._graph = graph('mgt-chat', GraphConfig.version);
    this.chatThreadsPerPage = chatThreadsPerPage;
    this._notificationClient = new GraphNotificationUserClient(this._eventEmitter, this._graph);
  }

  /**
   * Provides the number of chat threads to display with each load more.
   */
  public chatThreadsPerPage: number;

  /**
   * Provides a method to clean up any resources being used internally when a consuming component is being removed from the DOM
   */
  public async tearDown() {
    await this._notificationClient.tearDown();
  }

  /**
   * Load more chat threads if applicable.
   */
  public loadMoreChatThreads(): void {
    const state = this.getState();

    if (state.nextLink === '') {
      return;
    }

    const filter = state.nextLink.split('?')[1];
    void loadChatThreadsByPage(this._graph, filter).then(
      chats => this.handleChatThreads(chats),
      e => error(e)
    );
  }

  /**
   * Register a callback to receive state updates
   *
   * @param {(state: GraphChatListClient) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public onStateChange(handler: (state: GraphChatListClient) => void): void {
    if (!this._stateSubscribers.includes(handler)) {
      this._stateSubscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving state updates
   *
   * @param {(state: GraphChatListClient) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public offStateChange(handler: (state: GraphChatListClient) => void): void {
    const index = this._stateSubscribers.indexOf(handler);
    if (index !== -1) {
      this._stateSubscribers = this._stateSubscribers.splice(index, 1);
    }
  }

  /**
   * Register a callback to receive ChatList events
   *
   * @param {(event: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public onChatListEvent(handler: (event: ChatListEvent) => void): void {
    if (!this._chatListEventSubscribers.includes(handler)) {
      this._chatListEventSubscribers.push(handler);
    }
  }

  /**
   * Unregister a callback from receiving ChatList events
   *
   * @param {(event: ChatListEvent) => void} handler
   * @memberof StatefulGraphChatListClient
   */
  public offChatListEvent(handler: (event: ChatListEvent) => void): void {
    const index = this._chatListEventSubscribers.indexOf(handler);
    if (index !== -1) {
      this._chatListEventSubscribers = this._chatListEventSubscribers.splice(index, 1);
    }
  }

  private readonly _initialState: GraphChatListClient = {
    status: 'initial',
    activeErrorMessages: [],
    userId: '',
    chatThreads: [],
    nextLink: ''
  };

  /**
   * State of the chat client with initial values set
   *
   * @private
   * @type {GraphChatListClient}
   * @memberof StatefulGraphChatListClient
   */
  private _state: GraphChatListClient = { ...this._initialState };

  /**
   * Calls each subscriber with the next state to be emitted
   *
   * @param recipe - a function which produces the next state to be emitted
   */
  private notifyStateChange(recipe: (draft: GraphChatListClient) => void) {
    this._state = produce(this._state, recipe);
    this._stateSubscribers.forEach(handler => handler(this._state));
  }

  /**
   * Handle ChatListEvent event types.
   */
  private notifyChatMessageEventChange(message: ChatListEvent) {
    this._chatListEventSubscribers.forEach(handler => handler(message));
    if (message.type === 'chatRenamed' && message.message.eventDetail) {
      const eventDetail = message.message.eventDetail as EventMessageDetail;
      this.notifyStateChange((draft: GraphChatListClient) => {
        const chatThread = draft.chatThreads.find(c => c.id === message.message.chatId);
        if (chatThread) {
          chatThread.topic = eventDetail.chatDisplayName;
        }
      });
    }
    if (message.type === 'chatMessageReceived') {
      const chatThread = this._state.chatThreads.find(c => c.id === message.message.chatId);
      if (chatThread) {
        const msgInfo = message.message as ChatMessageInfo;
        this.notifyStateChange((draft: GraphChatListClient) => {
          const draftChatThread = draft.chatThreads.find(c => c.id === chatThread.id);
          if (!draftChatThread) {
            Error('Unexpected state discrepancy: Chat thread not found in draft state');
            return;
          }
          draftChatThread.lastMessagePreview = msgInfo;
          // remove the updated chat thread from the list and add it to the top
          draft.chatThreads = draft.chatThreads.filter(c => c.id !== draftChatThread.id);
          draft.chatThreads.unshift(draftChatThread);
        });
      } else {
        // if the chat thread is not in the list, load it, add it to the top
        loadChat(this._graph, message.message.chatId as string).then(chat => {
          this.notifyStateChange((draft: GraphChatListClient) => {
            draft.chatThreads.unshift(chat);
          });
        }).catch(e => Error(e));
      }
    }
  }

  /*
   * Returns the type of events by checking the chat message.
   */
  private getSystemMessageType(message: ChatMessage) {
    // check if this is a SystemEvent
    if (message.messageType === 'systemEventMessage') {
      const eventDetail = message.eventDetail as ChatMessageEvents;
      switch (eventDetail['@odata.type']) {
        case '#microsoft.graph.membersAddedEventMessageDetail':
          return 'memberAdded';
        case '#microsoft.graph.membersDeletedEventMessageDetail':
          return 'memberRemoved';
        case '#microsoft.graph.chatRenamedEventMessageDetail':
          return 'chatRenamed';
        default:
          return 'systemEvent';
      }
    }

    return 'chatMessageReceived';
  }

  /*
   * Event handler to be called when a new message is received by the notification service
   */
  private readonly onMessageReceived = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: this.getSystemMessageType(message) });
    }
  };

  /*
   * Event handler to be called when a message deletion is received by the notification service
   */
  private readonly onMessageDeleted = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageDeleted' });

      // void this._cache.deleteMessage(message.chatId, message);
    }
  };

  /*
   * Event handler to be called when a message edit is received by the notification service
   */
  private readonly onMessageEdited = (message: ChatMessage) => {
    if (message.chatId) {
      this.notifyChatMessageEventChange({ message, type: 'chatMessageEdited' });

      // await this._cache.cacheMessage(message.chatId, message);
    }
  };

  /**
   * Return the current state of the chat client
   *
   * @return {{GraphChatListClient}
   * @memberof StatefulGraphChatListClient
   */
  public getState(): GraphChatListClient {
    return this._state;
  }

  /*
   * Event handler to be called when we need to load more chat threads.
   */
  private readonly handleChatThreads = (chatThreadCollection: ChatThreadCollection) => {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.nextLink = '';

      const nextLinkUrl = chatThreadCollection['@odata.nextLink'];
      if (nextLinkUrl && nextLinkUrl !== '') {
        draft.nextLink = nextLinkUrl;
      }
      const uniqeChatThreads = chatThreadCollection.value.filter(
        c => draft.chatThreads.findIndex(t => t.id === c.id) === -1
      );
      draft.chatThreads = draft.chatThreads.concat(uniqeChatThreads);
    });
  };

  /**
   * Update the state of the client when the Login state changes
   *
   * @private
   * @param {LoginChangedEvent} e The event that triggered the change
   * @memberof StatefulGraphChatListClient
   */
  private readonly onLoginStateChanged = (e: LoginChangedEvent) => {
    switch (e.detail) {
      case ProviderState.SignedIn:
        // update userId and displayName
        this.updateUserInfo();
        // load messages?
        // configure subscriptions
        // emit new state;
        if (this.userId) {
          void this.updateUserSubscription();

          if (this._graph !== undefined) {
            loadChatThreads(this._graph, this.chatThreadsPerPage).then(
              chats => this.handleChatThreads(chats),
              err => error(err)
            );
          }
        }
        return;
      case ProviderState.SignedOut:
        // clear userId
        // clear subscriptions
        // clear messages
        // emit new state
        return;
      case ProviderState.Loading:
      default:
        // do nothing for now
        return;
    }
  };

  private readonly onActiveAccountChanged = (e: ActiveAccountChanged) => {
    // eslint-disable-next-line @typescript-eslint/no-unsafe-member-access
    if (e.detail && this.userId !== e.detail?.id) {
      void this.handleAccountChange();
    }
  };

  private readonly handleAccountChange = async () => {
    this.clearCurrentUserMessages();
    // need to ensure that we close any existing connection if present
    await this._notificationClient?.closeSignalRConnection();

    this.updateUserInfo();
    // by updating the followed chat the notification client will reconnect to SignalR
    await this.updateUserSubscription();
  };

  private clearCurrentUserMessages() {
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.status = 'initial'; // no message?
    });
  }

  /**
   * Changes the current user ID value to the current value.
   */
  private updateUserInfo() {
    this.userId = currentUserId();
  }

  /**
   * Current User ID.
   */
  private _userId = '';

  /**
   * Returns the current User ID.
   */
  public get userId() {
    return this._userId;
  }

  /**
   * Sets the current User ID and updates the state value.
   */
  private set userId(userId: string) {
    if (this._userId === userId) {
      return;
    }
    this._userId = userId;
    this.notifyStateChange((draft: GraphChatListClient) => {
      draft.userId = userId;
    });
  }

  public get sessionId(): string {
    return 'default';
  }

  /**
   * A helper to co-ordinate the loading of a chat and its messages, and the subscription to notifications for that chat
   *
   * @private
   * @memberof StatefulGraphChatListClient
   */
  private async updateUserSubscription() {
    // avoid subscribing to a resource with an empty userId
    if (this.userId) {
      // reset state to initial
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'initial';
      });
      // Subscribe to notifications for messages
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'creating server connections';
      });
      try {
        // Prefer sequential promise resolving to catch loading message errors
        // TODO: in parallel promise resolving, find out how to trigger different
        // TODO: state for failed subscriptions in GraphChatClient.onSubscribeFailed
        const tasks: Promise<unknown>[] = [];
        // subscribing to notifications will trigger the chatMessageNotificationsSubscribed event
        // this client will then load the chat and messages when that event listener is called
        tasks.push(this._notificationClient.subscribeToUserNotifications(this._userId, this.sessionId));
        await Promise.all(tasks);
      } catch (e) {
        console.error('Failed to load chat data or subscribe to notications: ', e);
        if (e instanceof GraphError) {
          this.notifyStateChange((draft: GraphChatListClient) => {
            draft.status = 'no messages';
          });
        }
      }
    } else {
      this.notifyStateChange((draft: GraphChatListClient) => {
        draft.status = 'no session id';
      });
    }
  }

  /**
   * Register event listeners for chat events to be triggered from the notification service
   */
  private registerEventListeners() {
    this._eventEmitter.on('chatMessageReceived', (message: ChatMessage) => void this.onMessageReceived(message));
    this._eventEmitter.on('chatMessageDeleted', this.onMessageDeleted);
    this._eventEmitter.on('chatMessageEdited', (message: ChatMessage) => void this.onMessageEdited(message));
  }
}

export { StatefulGraphChatListClient };
