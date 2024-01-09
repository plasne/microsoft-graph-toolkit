import React, { useEffect, useState } from 'react';
import { ChatListItem, IChatListItemInteractionProps } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps, ProviderState, Providers } from '@microsoft/mgt-react';
import { makeStyles, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { ChatMessageInfo } from '@microsoft/microsoft-graph-types';
import {
  ChatListEvent,
  StatefulGraphChatListClient,
  GraphChatListClient
} from '../../statefulClient/StatefulGraphChatListClient';
import { ChatListHeader } from '../ChatListHeader/ChatListHeader';
import { IChatListMenuItemsProps } from '../ChatListHeader/EllipsisMenu';
import { ChatListButtonItem } from '../ChatListHeader/ChatListButtonItem';
import ChatListMenuItem from '../ChatListHeader/ChatListMenuItem';

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
  }
});

interface EventMessageDetail {
  chatDisplayName: string;
}

// this is a stub to move the logic here that should end up here.
export const ChatList = (
  props: MgtTemplateProps &
    IChatListItemInteractionProps &
    IChatListMenuItemsProps & {
      buttonItems?: ChatListButtonItem[];
      chatThreadsPerPage: number;
    }
) => {
  const styles = useStyles();

  const [chatClient, setChatClient] = useState<StatefulGraphChatListClient | undefined>();
  const [chatState, setChatState] = useState<GraphChatListClient | undefined>();

  // wait for provider to be ready before setting client and state
  useEffect(() => {
    const provider = Providers.globalProvider;
    provider.onStateChanged(evt => {
      if (evt.detail === ProviderState.SignedIn) {
        const client = new StatefulGraphChatListClient(props.chatThreadsPerPage);
        setChatClient(client);
        setChatState(client.getState());
      }
    });
  }, []);

  // subscribe (from useGraphChatListClient)
  useEffect(() => {
    chatClient?.subscribeToUser('default');
  }, [chatClient]);

  // tear down (from useGraphChatListClient)
  useEffect(() => {
    return () => {
      void chatClient?.tearDown();
    };
  }, [chatClient]);

  const [menuItems, setMenuItems] = useState<ChatListMenuItem[]>(props.menuItems === undefined ? [] : props.menuItems);

  const onChatListEvent = (state: ChatListEvent) => {
    if (chatState !== undefined) {
      if (state.type === 'chatRenamed' && state.message.eventDetail) {
        let eventDetail = state.message.eventDetail as EventMessageDetail;
        let chatThread = chatState.chatThreads.find(c => c.id === state.message.chatId);
        if (chatThread) {
          chatThread.topic = eventDetail.chatDisplayName;
        }
      }

      if (state.type === 'chatMessageReceived') {
        let chatThread = chatState.chatThreads.find(c => c.id === state.message.chatId);
        if (chatThread) {
          let msgInfo = state.message as ChatMessageInfo;
          chatThread.lastMessagePreview = msgInfo;
        } else {
        }
      }
    }

    // TODO: implementation will happen later, right now, we just need to make sure messages are coming thru in console logs.
    console.log(state.type);
    console.log(state.message);
  };

  const loadMore = () => {
    chatClient?.loadMoreChatThreads();
  };

  useEffect(() => {
    if (chatClient) {
      chatClient.onChatListEvent(onChatListEvent);
      chatClient.onStateChange(setChatState);
      return () => {
        chatClient.offChatListEvent(onChatListEvent);
        chatClient.offStateChange(setChatState);
      };
    }
  }, [chatClient]);

  const chatListButtonItems = props.buttonItems === undefined ? [] : props.buttonItems;

  useEffect(() => {
    const markAllAsRead = {
      displayText: 'Mark all as read',
      onClick: () => {}
    };

    menuItems.unshift(markAllAsRead);
    setMenuItems(menuItems);
  }, []);

  return (
    // This is a temporary approach to render the chatlist items. This should be replaced.
    <FluentThemeProvider fluentTheme={FluentTheme}>
      <FluentProvider theme={webLightTheme}>
        <div>
          <div className={styles.headerContainer}>
            <ChatListHeader buttonItems={chatListButtonItems} menuItems={menuItems} />
          </div>
          <div>
            {chatState?.chatThreads.map(c => (
              <ChatListItem key={c.id} chat={c} myId={chatState?.userId} onSelected={props.onSelected} />
            ))}
            {chatState?.nextLink !== '' && (
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
