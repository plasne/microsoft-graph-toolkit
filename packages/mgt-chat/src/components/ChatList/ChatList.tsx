import React, { useCallback, useEffect, useState, useRef } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, Spinner, log, error } from '@microsoft/mgt-react';
import { makeStyles, Button, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { ChatMessage } from '@microsoft/microsoft-graph-types';
import {
  StatefulGraphChatListClient,
  GraphChatListClient,
  GraphChatThread
} from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { Error } from '../Error/Error';
import { LoadingMessagesErrorIcon } from '../Error/LoadingMessageErrorIcon';
import { CreateANewChat } from '../Error/CreateANewChat';
import { OpenTeamsLinkError } from '../Error/OpenTeams';
import { IChatListActions } from '../ChatListHeader/IChatListActions';

export interface IChatListProps {
  onSelected: (e: GraphChatThread) => void;
  onUnselected?: (e: GraphChatThread) => void;
  onLoaded?: (e: GraphChatThread[]) => void;
  onAllMessagesRead: (e: string[]) => void;
  buttonItems?: ChatListButtonItem[];
  chatThreadsPerPage: number;
  lastReadTimeInterval?: number;
  selectedChatId?: string;
  onMessageReceived?: (msg: ChatMessage) => void;
  onConnectionChanged?: (connected: boolean) => void;
}

const useStyles = makeStyles({
  chatList: {
    display: 'flex',
    flexDirection: 'column',
    height: '100%',
    ...shorthands.overflow('hidden'),
    paddingBlockEnd: '12px'
  },
  chatListItems: {
    height: '100%',
    visibility: 'visible',
    paddingRight: '2px'
  },
  scrollbox: {
    ...shorthands.overflow('auto'),
    visibility: 'hidden',
    '::-webkit-scrollbar': {
      width: '8px'
    },
    '::-webkit-scrollbar-track': {
      backgroundColor: 'transparent'
    },
    '::-webkit-scrollbar-thumb': {
      backgroundColor: '#a0a0a0'
    },
    ':hover': {
      visibility: 'visible'
    },
    ':focus': {
      visibility: 'visible'
    }
  },
  fullHeight: {
    height: '100%'
  },
  spinner: {
    justifyContent: 'center',
    display: 'flex',
    alignItems: 'center',
    height: '100%'
  },
  button: {
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%',
    ...shorthands.padding('0px'),
    ...shorthands.border('none')
  },
  error: {
    display: 'flex',
    justifyContent: 'center',
    height: '100%'
  },
  bottomWhitespace: {
    height: '80%',
    width: '100%'
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = ({
  lastReadTimeInterval = 30000, // default to 30 seconds
  selectedChatId,
  onMessageReceived,
  onAllMessagesRead,
  onLoaded,
  onConnectionChanged,
  onSelected,
  onUnselected,
  chatThreadsPerPage,
  ...props
}: MgtTemplateProps & IChatListProps & IChatListMenuItemsProps) => {
  const styles = useStyles();
  const [chatListClient] = useState<StatefulGraphChatListClient>(() => new StatefulGraphChatListClient());
  const [chatListActions] = useState<IChatListActions>(() => {
    return {
      markAllChatThreadsAsRead: () => chatListClient.markAllChatThreadsAsRead()
    };
  });
  const [initialLastReadTimeInterval, setInitialLastReadTimeInterval] = useState<number | undefined>();
  const [chatListState, setChatListState] = useState<GraphChatListClient | undefined>();

  useEffect(() => {
    chatListClient.onStateChange(setChatListState);
    return () => {
      chatListClient.offStateChange(setChatListState);
    };
  }, [chatListClient]);

  useEffect(() => {
    if (chatThreadsPerPage < 1) {
      error('chatThreadsPerPage must be greater than 0!');
      return;
    }

    // todo: implement a upperbound limit for chatThreadsPerPage
    chatListClient.chatThreadsPerPage = chatThreadsPerPage;
  }, [chatListClient, chatThreadsPerPage]);

  useEffect(() => {
    if (!selectedChatId) {
      chatListClient.clearSelectedChat();
    } else {
      chatListClient.setSelectedChatId(selectedChatId);
    }
  }, [chatListClient, selectedChatId]);

  // Store last read time in cache so that when the user comes back to the chat list,
  // we know what messages they are likely to have not read. This is not perfect because
  // the user could have read messages in another client (for instance, the Teams client).
  useEffect(() => {
    // setup timer only after we have a defined chatListClient
    if (initialLastReadTimeInterval) {
      error('lastReadTimeInterval can only be set once.');
      return;
    }

    if (lastReadTimeInterval < 1) {
      error('lastReadTimeInterval must be greater than 0!');
      return;
    }

    // todo: implement a upperbound limit for lastReadTimeInterval
    setInitialLastReadTimeInterval(lastReadTimeInterval);

    const timer = setInterval(() => {
      chatListClient.cacheLastReadTime('selected');
    }, lastReadTimeInterval);

    return () => {
      clearInterval(timer);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [chatListClient, lastReadTimeInterval]); // initialLastReadTimeInterval is used to ensure we only set interval once.

  useEffect(() => {
    if (!chatListState) {
      return;
    }

    // handle state changes
    if (chatListState.status === 'chat message received' && onMessageReceived && chatListState.chatMessage) {
      onMessageReceived(chatListState.chatMessage);
    }

    if (chatListState.status === 'chat selected' && onSelected && chatListState.internalSelectedChat) {
      onSelected(chatListState.internalSelectedChat);
    }

    if (chatListState.status === 'chat unselected' && onUnselected && chatListState.internalPrevSelectedChat) {
      onUnselected(chatListState.internalPrevSelectedChat);
    }

    if (chatListState.status === 'chats read' && onAllMessagesRead && chatListState.chatThreads) {
      onAllMessagesRead(chatListState.chatThreads.map(c => c.id!));
    }

    if (chatListState.status === 'chats loaded' && onLoaded) {
      onLoaded(chatListState?.chatThreads ?? []);
    }

    if (
      chatListState.status === 'chats loaded' &&
      chatListState.fireOnSelected &&
      onSelected &&
      chatListState.internalSelectedChat
    ) {
      onSelected(chatListState.internalSelectedChat);
    }

    if (chatListState.status === 'no chats' && onLoaded) {
      onLoaded([]);
    }

    if (chatListState.status === 'server connection established' && onConnectionChanged) {
      onConnectionChanged(true);
      void chatListClient.tryLoadChatThreads().catch(e => chatListClient.raiseFatalError(e as Error));
    }

    if (chatListState.status === 'server connection lost' && onConnectionChanged) {
      onConnectionChanged(false);
    }
  }, [
    chatListClient,
    chatListState,
    onLoaded,
    onMessageReceived,
    onSelected,
    onUnselected,
    onAllMessagesRead,
    onConnectionChanged
  ]);

  // this only runs once when the component is unmounted
  useEffect(() => {
    // tear down
    return () => {
      // log state of chatlistclient for debugging purposes
      log('ChatList unmounted.', chatListClient.getState());
      chatListClient.tearDown();
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []); // required for the teardown to only run once

  const onClickChatListItem = useCallback(
    (chat: GraphChatThread) => {
      if (chatListClient) {
        chatListClient.setSelectedChat(chat);
      }
    },
    [chatListClient]
  );

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;
  const chatListMenuItems = props.menuItems === undefined ? [] : props.menuItems;

  const isLoading = ['creating server connections', 'subscribing to notifications', 'loading chats'].includes(
    chatListState?.status ?? ''
  );

  const targetElementRef = useRef(null);

  useEffect(() => {
    // define the intersection observer callback
    const handleIntersection = (entries: IntersectionObserverEntry[]) => {
      for (const entry of entries) {
        // The element has come into view, you can perform your actions here
        if (entry.isIntersecting && chatListClient) {
          chatListClient.tryLoadMoreChatThreads().catch(e => error('Failed to load more chat threads.', e));
        }
      }
    };

    // create a new Intersection Observer instance
    const observer = new IntersectionObserver(handleIntersection, {
      root: null, // observing intersections with the viewport
      rootMargin: '0px',
      threshold: 0.03 // Callback is invoked when 3% of the target is visible
    });

    // start observing
    if (targetElementRef.current) {
      observer.observe(targetElementRef.current);
    }

    return () => observer.disconnect();
  }, [chatListClient, chatListState]);

  return (
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme} className={styles.fullHeight}>
        <div className={styles.chatList}>
          {chatListState?.status !== 'server connection lost' && chatListActions && (
            <ChatListHeader actions={chatListActions} buttonItems={chatListButtonItems} menuItems={chatListMenuItems} />
          )}
          {chatListState && chatListState.chatThreads.length > 0 ? (
            <>
              <div className={styles.scrollbox}>
                <div className={styles.chatListItems}>
                  {chatListState?.chatThreads.map(c => (
                    <Button className={styles.button} key={c.id} onClick={() => onClickChatListItem(c)}>
                      <ChatListItem
                        key={c.id}
                        chat={c}
                        userId={chatListState.userId}
                        isSelected={c.id === chatListState?.internalSelectedChat?.id}
                        isRead={c.isRead}
                      />
                    </Button>
                  ))}
                  {chatListState?.moreChatThreadsToLoad && (
                    <div ref={targetElementRef} className={styles.bottomWhitespace}>
                      &nbsp;
                    </div>
                  )}
                </div>
              </div>
            </>
          ) : (
            <>
              <div className={styles.error}>
                {isLoading && (
                  <div className={styles.spinner}>
                    <Spinner /> <br />
                    {chatListState?.status}
                  </div>
                )}
                {chatListState?.status === 'no chats' && (
                  <Error
                    icon={LoadingMessagesErrorIcon}
                    message="No threads were found for this user."
                    subheading={CreateANewChat}
                  ></Error>
                )}
                {chatListState?.status === 'server connection lost' && (
                  <Error message="We ran into a problem. Reconnecting..." subheading={OpenTeamsLinkError}></Error>
                )}
                {chatListState?.status === 'fatal error' && (
                  <Error message="We're sorry—we've run into an issue." subheading={OpenTeamsLinkError}></Error>
                )}
              </div>
            </>
          )}
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
