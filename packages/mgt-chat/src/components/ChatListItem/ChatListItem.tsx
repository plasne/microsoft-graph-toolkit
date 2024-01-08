import React from 'react';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
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

export interface IChatListItemInteractionProps {
  onSelected: (e: Chat) => void;
}

interface IMgtChatListItemProps {
  chat: Chat;
  myId: string | undefined;
}

const useStyles = makeStyles({
  chatListItem: {
    display: 'flex',
    flexDirection: 'row',
    alignItems: 'center',
    width: '100%',
    ...shorthands.padding('10px'),
    ...shorthands.borderBottom('1px solid #ccc')
  },
  profileImage: {
    flexGrow: 0,
    flexShrink: 0,
    flexBasis: 'auto',
    marginRight: '10px',
    objectFit: 'cover', // This ensures the image covers the area without stretching
    display: 'flex',
    alignItems: 'center', // This will vertically center the image
    justifyContent: 'center' // This will horizontally center the image
  },
  defaultProfileImage: {
    ...shorthands.borderRadius('50%'),
    objectFit: 'cover', // This ensures the image covers the area without stretching
    display: 'flex',
    alignItems: 'center', // This will vertically center the image
    justifyContent: 'center' // This will horizontally center the image
  },
  chatInfo: {
    flexGrow: 1,
    flexShrink: 1,
    flexBasis: 'auto',
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
    textAlign: 'left',
    ...shorthands.margin('0'),
    fontSize: '0.9em',
    color: '#666',
    textOverflow: 'ellipsis',
    ...shorthands.overflow('hidden'),
    whiteSpace: 'nowrap',
    width: 'auto'
  },
  chatTimestamp: {
    flexShrink: 0,
    flexBasis: 'auto',
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

export const ChatListItem = ({ chat, myId, onSelected }: IMgtChatListItemProps & IChatListItemInteractionProps) => {
  const styles = useStyles();

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

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
      return previewString;
    }

    // handle emojis
    content = rewriteEmojiContent(content);

    // handle images
    const imageMatch = content.match(graphImageUrlRegex);
    if (imageMatch) {
      content = 'Sent an image.';
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

  return (
    <Button
      className={styles.chatListItem}
      onClick={() => {
        onSelected(chat);
      }}
    >
      <div className={styles.profileImage}>{getDefaultProfileImage(chat)}</div>
      <div className={styles.chatInfo}>
        <h3 className={styles.chatTitle}>{inferTitle(chat)}</h3>
        <p className={styles.chatMessage}>{enrichPreviewMessage(chat.lastMessagePreview)}</p>
      </div>
      <div className={styles.chatTimestamp}>{extractTimestamp(determineCorrectTimestamp(chat))}</div>
    </Button>
  );
};
