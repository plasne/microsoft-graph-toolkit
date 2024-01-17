import React, { SetStateAction, useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers, log } from '@microsoft/mgt-react';
import { makeStyles, Button, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import {
  StatefulGraphChatListClient,
  GraphChatListClient,
  ChatListEvent,
  GraphChatThread
} from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { ChatListMenuItem } from '../ChatListHeader/ChatListMenuItem';
import { LastReadCache } from '../../statefulClient/Caching/LastReadCache';

export interface IChatListItemProps {
  onSelected: (e: GraphChat) => void;
  onLoaded?: () => void;
  onAllMessagesRead: (e: string[]) => void;
  buttonItems?: ChatListButtonItem[];
  chatThreadsPerPage: number;
  lastReadTimeInterval?: number;
  onMessageReceived?: () => void;
}

const useStyles = makeStyles({
  headerContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  linkContainer: {
    display: 'flex',
    justifyContent: 'center',
    width: '100%',
    ...shorthands.padding('10px')
  },
  loadMore: {
    textDecorationLine: 'none',
    fontSize: '1.2em',
    fontWeight: 'bold',
    '&:hover': {
      textDecorationLine: 'none' // This removes the underline when hovering
    }
  },
  button: {
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%',
    ...shorthands.padding('0px'),
    ...shorthands.border('none')
  }
});

// this is a stub to move the logic here that should end up here.
export const ChatList = ({
  lastReadTimeInterval = 30000, // default to 30 seconds
  ...props
}: MgtTemplateProps & IChatListItemProps & IChatListMenuItemsProps) => {
  const styles = useStyles();
  const [chatListClient, setChatListClient] = useState<StatefulGraphChatListClient | undefined>();
  const [chatListState, setChatListState] = useState<GraphChatListClient | undefined>();
  const [menuItems, setMenuItems] = useState<ChatListMenuItem[]>(props.menuItems === undefined ? [] : props.menuItems);
  const [selectedItem, setSelectedItem] = useState<string>();
  const cache = new LastReadCache();

  // wait for provider to be ready before setting client and state
  useEffect(() => {
    const provider = Providers.globalProvider;
    provider.onStateChanged(evt => {
      if (evt.detail === ProviderState.SignedIn) {
        const client = new StatefulGraphChatListClient(props.chatThreadsPerPage);
        setChatListClient(client);
        setChatListState(client.getState());
      }
    });
  }, []);

  // Store last read time in cache so that when the user comes back to the chat list,
  // we know what messages they are likely to have not read. This is not perfect because
  // the user could have read messages in another client (for instance, the Teams client).
  useEffect(() => {
    const timer = setInterval(() => {
      if (selectedItem) {
        log(`caching the last-read timestamp of now to chat ID '${selectedItem}'...`);
        cache.cacheLastReadTime(selectedItem, new Date());
      }
    }, lastReadTimeInterval);

    return () => {
      clearInterval(timer);
    };
  }, [selectedItem]);

  useEffect(() => {
    // handles events emitted from the chat list client
    const handleChatListEvent = (event: ChatListEvent) => {
      if (event.type === 'chatMessageReceived') {
        if (props.onMessageReceived) {
          props.onMessageReceived();
        }
      }
    };

    if (chatListClient) {
      chatListClient.onStateChange(setChatListState);
      chatListClient.onStateChange(state => {
        if (state.status === 'chat threads loaded' && props.onLoaded) {
          const markAllAsRead = {
            displayText: 'Mark all as read',
            onClick: () => markAllThreadsAsRead()
          };
          // clone the menuItems array
          const updatedMenuItems = [...menuItems];
          updatedMenuItems.unshift(markAllAsRead);
          setMenuItems(updatedMenuItems);
          props.onLoaded();
        }
      });
      chatListClient.onChatListEvent(handleChatListEvent);
      return () => {
        chatListClient.offStateChange(setChatListState);
        chatListClient.offChatListEvent(handleChatListEvent);
        void chatListClient.tearDown();
      };
    }
  }, [chatListClient]);

  const markAllThreadsAsRead = () => {
    chatListClient?.markAllChatThreadsAsRead();
  };

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;

  // We need to have a function for "this" to work within the loadMoreChatThreads function, otherwise we get a undefined error.
  const loadMore = () => {
    chatListClient?.loadMoreChatThreads();
  };

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <div>
          <div className={styles.headerContainer}>
            <ChatListHeader buttonItems={chatListButtonItems} menuItems={menuItems} />
          </div>
          <div>
            {chatListState?.chatThreads.map(c => (
              <Button
                className={styles.button}
                key={c.id}
                onClick={() => {
                  // set selected state only once per click event
                  if (c.id && c.id !== selectedItem) {
                    setSelectedItem(c.id);
                    cache.cacheLastReadTime(c.id, new Date());
                    props.onSelected(c);
                  }
                }}
              >
                <ChatListItem
                  key={c.id}
                  chat={c}
                  myId={chatListState.userId}
                  isSelected={c.id === selectedItem}
                  isRead={c.isRead}
                />
              </Button>
            ))}
            {chatListState?.moreChatThreadsToLoad === true && (
              <div className={styles.linkContainer}>
                <Link onClick={loadMore} href="#" className={styles.loadMore}>
                  load more
                </Link>
              </div>
            )}
          </div>
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};
