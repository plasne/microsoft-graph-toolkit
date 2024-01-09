import React, { useState, useEffect } from 'react';
import { makeStyles, mergeClasses, shorthands } from '@fluentui/react-components';
import {
  Chat,
  AadUserConversationMember,
  NullableOption,
  ChatMessageInfo,
  TeamworkApplicationIdentity
} from '@microsoft/microsoft-graph-types';
import { MgtTemplateProps, Person, PersonCardInteraction, PersonProps } from '@microsoft/mgt-react';
import { error } from '@microsoft/mgt-element';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';
import { rewriteEmojiContent } from '../../utils/rewriteEmojiContent';
import { convert } from 'html-to-text';

interface IMgtChatListItemProps {
  chat: Chat;
  myId: string | undefined;
  isSelected: boolean;
  isRead: boolean;
}

const useStyles = makeStyles({
  // highlight selection
  isSelected: {
    backgroundColor: '#e6f7ff'
  },

  isUnSelected: {
    backgroundColor: '#ffffff'
  },

  // highlight text
  isBold: {
    fontWeight: 'bold'
  },

  isNormal: {
    fontWeight: 'normal'
  },

  chatListItem: {
    display: 'flex',
    alignItems: 'center',
    width: '100%',
    paddingRight: '10px',
    paddingLeft: '10px'
  },

  profileImage: {
    ...shorthands.flex('0 0 auto'),
    marginRight: '10px',
    objectFit: 'cover',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },

  defaultProfileImage: {
    ...shorthands.borderRadius('50%'),
    objectFit: 'cover',
    display: 'flex',
    alignItems: 'center',
    justifyContent: 'center'
  },

  chatInfo: {
    flexGrow: 1,
    flexShrink: 1,
    minWidth: 0,
    ...shorthands.padding('5px')
  },

  chatTitle: {
    textAlign: 'left',
    ...shorthands.margin('0'),
    fontSize: '1em',
    color: '#333',
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: 'auto'
  },

  chatMessage: {
    fontSize: '0.9em',
    color: '#666',
    textAlign: 'left',
    ...shorthands.margin('0'),
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: 'auto'
  },

  chatTimestamp: {
    flexShrink: 0,
    textAlign: 'right',
    alignSelf: 'start',
    marginLeft: 'auto',
    paddingLeft: '10px',
    fontSize: '0.8em',
    color: '#999',
    whiteSpace: 'nowrap'
  },

  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});

/**
 * Regex to detect and replace image urls using graph requests to supply the image content
 */
const graphImageUrlRegex = /(<img[^>]+)/;

export const ChatListItem = ({ chat, myId, isSelected, isRead }: IMgtChatListItemProps) => {
  const styles = useStyles();

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

  const [read, setRead] = useState<boolean>(isRead);

  // when isSelected changes to true, setRead to true
  useEffect(() => {
    if (isSelected) {
      setRead(true);
    }
  }, [isSelected]);

  // Copied and modified from the sample ChatItem.tsx
  // Determines the title in the case of 1:1 and self chats
  // Self Chats are not possible, however, 1:1 chats with a bot will show no other members other than self.
  const inferTitle = (chatObj: Chat) => {
    if (myId && chatObj.chatType === 'oneOnOne' && chatObj.members) {
      const other = chatObj.members.find(m => (m as AadUserConversationMember).userId !== myId);
      const me = chatObj.members.find(m => (m as AadUserConversationMember).userId === myId);
      const application = chatObj.lastMessagePreview?.from?.application as TeamworkApplicationIdentity;
      // if there is no other member, return the application display name
      if (other) {
        return `${other?.displayName || (other as AadUserConversationMember)?.email || other?.id}`;
      } else if (application && me) {
        return `${application?.displayName}` || `${application?.id}`;
      }
    }
    if (chatObj.chatType === 'group' && chatObj.members) {
      const others = chatObj.members.filter(m => (m as AadUserConversationMember).userId !== myId);
      // if there are 3 or less members, display all members' first names
      if (chatObj.members.length <= 3) {
        return (
          chatObj.topic ||
          others.map(m => (m as AadUserConversationMember).displayName?.split(' ')[0]).join(', ') ||
          chatObj.chatType
        );
        // if there are more than 3 members, display the first 3 members' first names and a count of the remaining members
      } else if (chatObj.members.length > 3) {
        let firstThreeMembersSlice = others.slice(0, 3);
        let remainingMembersCount = chatObj.members.length - 3;
        let groupMembersString =
          firstThreeMembersSlice.map(m => (m as AadUserConversationMember).displayName?.split(' ')[0]).join(', ') +
          ' +' +
          remainingMembersCount;
        return chatObj.topic || groupMembersString || chatObj.chatType;
      }
    }
    return chatObj.topic || chatObj.chatType;
  };

  // Derives the timestamp to display
  const extractTimestamp = (timestamp: NullableOption<string> | undefined): string => {
    if (timestamp === undefined || timestamp === null) return '';
    const currentDate = new Date();
    const date = new Date(timestamp);

    const [month, day, year] = [date.getMonth(), date.getDate(), date.getFullYear()];
    const [currentMonth, currentDay, currentYear] = [
      currentDate.getMonth(),
      currentDate.getDate(),
      currentDate.getFullYear()
    ];

    // if the message was sent today, return the time
    if (currentDay === day && currentMonth === month && currentYear === year) {
      return date.toLocaleTimeString(navigator.language, { hour: 'numeric', minute: '2-digit' });
    }

    // if the message was sent in a previous year, include the year
    if (currentYear !== year) {
      return date.toLocaleDateString(navigator.language, { month: 'numeric', day: 'numeric', year: 'numeric' });
    }

    // otherwise, return the month and day
    return date.toLocaleDateString(navigator.language, { month: 'numeric', day: 'numeric' });
  };

  // Chooses the correct timestamp to display
  const determineCorrectTimestamp = (chatObj: Chat) => {
    let timestamp: Date | undefined;

    // lastMessageTime is the time of the last message sent in the chat
    // lastUpdatedTime is Date and time at which the chat was renamed or list of members were last changed.
    const lastMessageTime = new Date(chatObj.lastMessagePreview?.createdDateTime as string);
    const lastUpdatedTime = new Date(chatObj.lastUpdatedDateTime as string);

    if (lastMessageTime && lastUpdatedTime) {
      timestamp = new Date(Math.max(lastMessageTime.getTime(), lastUpdatedTime.getTime()));
    } else if (lastMessageTime) {
      timestamp = lastMessageTime;
    } else if (lastUpdatedTime) {
      timestamp = lastUpdatedTime;
    }

    return String(timestamp);
  };

  const getDefaultProfileImage = (chatObj: Chat) => {
    // define the JSX for FluentUI Icons + Styling
    const oneOnOneProfilePicture = <ChatListItemIcon chatType="oneOnOne" />;
    const GroupProfilePicture = <ChatListItemIcon chatType="group" />;

    const other = chatObj.members?.find(m => (m as AadUserConversationMember).userId !== myId);
    const otherAad = other as AadUserConversationMember;
    const application = chatObj.lastMessagePreview?.from?.application as TeamworkApplicationIdentity;
    let iconId: string | undefined;
    switch (true) {
      case chat.chatType === 'oneOnOne':
        if (!otherAad && application?.id) {
          iconId = application.id;
        } else {
          iconId = otherAad?.userId as string;
        }
        const Default = (props: MgtTemplateProps) => {
          return <div className={styles.defaultProfileImage}>{oneOnOneProfilePicture}</div>;
        };
        return (
          <Person
            className={styles.person}
            userId={iconId}
            avatarSize="small"
            showPresence={true}
            personCardInteraction={PersonCardInteraction.hover}
          >
            <Default template="no-data" />
          </Person>
        );
      case chat.chatType === 'group':
        return GroupProfilePicture;
      default:
        error(`Error: Unexpected chatType: ${chat.chatType}`);
        return oneOnOneProfilePicture;
    }
  };

  const enrichPreviewMessage = (previewMessage: NullableOption<ChatMessageInfo> | undefined) => {
    let previewString = '';
    let content = previewMessage?.body?.content as string;

    // handle null or undefined content
    if (!content) {
      if (previewMessage?.from?.user?.id === myId) {
        previewString = 'You: Sent a message';
      } else if (previewMessage?.from?.user?.displayName) {
        previewString = previewMessage?.from?.user?.displayName + ': Sent a message';
      } else if (previewMessage?.from?.application?.displayName) {
        previewString = previewMessage?.from?.application?.displayName + ': Sent a message';
      }
      return previewString;
    }

    // handle emojis
    content = rewriteEmojiContent(content);

    // handle images
    const imageMatch = content.match(graphImageUrlRegex);
    if (imageMatch) {
      content = 'Sent an image';
    }

    // convert html to text
    content = convert(content);

    // handle general chats from people and bots
    if (previewMessage?.from?.user?.id === myId) {
      previewString = 'You: ' + content;
    } else if (previewMessage?.from?.user?.displayName) {
      previewString = previewMessage?.from?.user?.displayName + ': ' + content;
    } else if (previewMessage?.from?.application?.displayName) {
      previewString = previewMessage?.from?.application?.displayName + ': ' + content;
    }

    // handle all events
    if (previewMessage?.eventDetail) {
      previewString = content as string;
    }

    return previewString;
  };

  const container = mergeClasses(
    styles.chatListItem,
    isSelected ? styles.isSelected : styles.isUnSelected,
    read ? styles.isNormal : styles.isBold
  );

  return (
    <div className={container}>
      <div className={styles.profileImage}>{getDefaultProfileImage(chat)}</div>
      <div className={styles.chatInfo}>
        <p className={styles.chatTitle}>{inferTitle(chat)}</p>
        <p className={styles.chatMessage}>{enrichPreviewMessage(chat.lastMessagePreview)}</p>
      </div>
      <div className={styles.chatTimestamp}>{extractTimestamp(determineCorrectTimestamp(chat))}</div>
    </div>
  );
};
