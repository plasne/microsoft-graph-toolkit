import React, { useEffect, useState } from 'react';
import { ChatListItem, IChatListItemInteractionProps } from '../ChatListItem/ChatListItem';
import { MgtTemplateProps } from '@microsoft/mgt-react';
import { makeStyles, Link, FluentProvider, shorthands, webLightTheme } from '@fluentui/react-components';
import { FluentThemeProvider } from '@azure/communication-react';
import { FluentTheme } from '@fluentui/react';
import { Chat as GraphChat } from '@microsoft/microsoft-graph-types';
import { ChatListEvent, StatefulGraphChatListClient } from '../../statefulClient/StatefulGraphChatListClient';
import { useGraphChatListClient } from '../../statefulClient/useGraphChatListClient';
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

// this is a stub to move the logic here that should end up here.
export const ChatList = (
  props: MgtTemplateProps &
    IChatListItemInteractionProps &
    IChatListMenuItemsProps & {
      buttonItems?: ChatListButtonItem[];
    }
) => {
  const { value } = props.dataContext as { value: GraphChat[] };
  const chats: GraphChat[] = value;

  const styles = useStyles();
  const chatClient: StatefulGraphChatListClient = useGraphChatListClient();
  const [chatState, setChatState] = useState(chatClient.getState());
  const [chatThreads, setChatThreads] = useState<GraphChat[]>(chats);
  const [menuItems, setMenuItems] = useState<ChatListMenuItem[]>(props.menuItems === undefined ? [] : props.menuItems);

  const onChatListEvent = (state: ChatListEvent) => {
    // TODO: implementation will happen later, right now, we just need to make sure messages are coming thru in console logs.

    console.log(state.type);
    console.log(state.message);
  };

  useEffect(() => {
    chatClient.onChatListEvent(onChatListEvent);
    chatClient.onStateChange(setChatState);
    return () => {
      chatClient.offChatListEvent(onChatListEvent);
      chatClient.offStateChange(setChatState);
    };
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
            {chatThreads.map(c => (
              <ChatListItem key={c.id} chat={c} myId={chatState.userId} onSelected={props.onSelected} />
            ))}
            <div className={styles.linkContainer}>
              <Link href="#" className={styles.loadMore}>
                load more
              </Link>
            </div>
          </div>
        </div>
      </FluentProvider>
    </FluentThemeProvider>
  );
};

export default ChatList;
