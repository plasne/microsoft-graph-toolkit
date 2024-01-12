import React, { useEffect, useState } from 'react';
import { ChatListItem } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers, log } from '@microsoft/mgt-react';
import { makeStyles, Button, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { StatefulGraphChatListClient, GraphChatListClient } from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import { ChatListMenuItem } from '../ChatListHeader/ChatListMenuItem';
import { LastReadCache } from '../../statefulClient/Caching/LastReadCache';

export interface IChatListItemProps {
  onSelected: (e: GraphChat) => void;
  onAllMessagesRead: (e: string[]) => void;
  buttonItems?: ChatListButtonItem[];
  chatThreadsPerPage: number;
  lastReadTimeInterval?: number;
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
  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;
  const [menuItems, setMenuItems] = useState<ChatListMenuItem[]>(props.menuItems === undefined ? [] : props.menuItems);
  const [selectedItem, setSelectedItem] = useState<string>();
  const [readItems, setReadItems] = useState<string[]>([]);

  const cache = new LastReadCache();

  // We need to have a function for "this" to work within the loadMoreChatThreads function, otherwise we get a undefined error.
  const loadMore = () => {
    chatListClient?.loadMoreChatThreads();
  };

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
    if (chatListClient) {
      chatListClient.onStateChange(setChatListState);
      return () => {
        void chatListClient.tearDown();
        chatListClient.offStateChange(setChatListState);
      };
    }
  }, [chatListClient]);

  useEffect(() => {
    // when chat threads are updated, add markAllAsRead callback to menuItems
    const chatThreads = chatListState?.chatThreads;
    if (chatThreads && chatThreads.length > 0) {
      log("chatThreads updated, adding 'mark all as read' to menu...");
      const markAllAsRead = {
        displayText: 'Mark all as read',
        onClick: () => {
          var itemsMarkedAsRead = chatThreads
            .map(c => (c.id && c.id !== selectedItem ? c.id : ''))
            .filter(id => id !== '');
          // loop through readItems and cache the last read time
          if (selectedItem && readItems.includes(selectedItem)) {
            // add selected item to itemsMarkedAsRead
            itemsMarkedAsRead.push(selectedItem);
          }
          log(`marking all ${itemsMarkedAsRead?.length} chat threads as read...`);
          setReadItems(itemsMarkedAsRead);
          props.onAllMessagesRead(itemsMarkedAsRead);
          itemsMarkedAsRead.forEach(id => {
            cache.cacheLastReadTime(id, new Date());
          });
        }
      };
      // clone the menuItems array
      const updatedMenuItems = [...menuItems];
      const markAllAsReadIndex = menuItems.findIndex(m => m.displayText === markAllAsRead.displayText);
      if (markAllAsReadIndex === -1) {
        // adds read menu item only once
        log(`adding 'mark all as read' to menu...`);
        updatedMenuItems.unshift(markAllAsRead);
        setMenuItems(updatedMenuItems);
      } else {
        // adds latest updated chat threads to scope of mark as read menu item
        log(`updating 'mark all as read' menu...`);
        updatedMenuItems[markAllAsReadIndex] = markAllAsRead;
        setMenuItems(updatedMenuItems);
      }
    }
  }, [chatListState, selectedItem]);

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
                    props.onSelected(c);
                    if (!readItems.includes(c.id)) {
                      setReadItems([...readItems, c.id]);
                      cache.cacheLastReadTime(c.id, new Date());
                    }
                  }
                }}
              >
                <ChatListItem
                  key={c.id}
                  chat={c}
                  myId={chatListState.userId}
                  isSelected={c.id === selectedItem}
                  isRead={readItems.includes(c.id ?? '')}
                />
              </Button>
            ))}
            {chatListState?.nextLink !== '' && (
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

export default ChatList;
