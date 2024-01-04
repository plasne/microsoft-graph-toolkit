import React, { useCallback } from 'react';
import { makeStyles, shorthands, Button } from '@fluentui/react-components';
import {
  Chat,
  AadUserConversationMember,
  MembersAddedEventMessageDetail,
  NullableOption,
  ChatMessageInfo
} from '@microsoft/microsoft-graph-types';
import { Person, PersonCardInteraction, Spinner } from '@microsoft/mgt-react';
import { error } from '@microsoft/mgt-element';
import { Providers, ProviderState } from '@microsoft/mgt-element';
import { ChatListItemIcon } from '../ChatListItemIcon/ChatListItemIcon';

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
    justifyContent: 'space-between', // Add this if you want to push the timestamp to the end
    width: '100%',
    ...shorthands.padding('10px'),
    ...shorthands.borderBottom('1px solid #ccc')
  },
  profileImage: {
    flexGrow: 0,
    flexShrink: 0,
    flexBasis: 'auto',
    ...shorthands.borderRadius('50%'), // This will make it round
    marginRight: '10px',
    objectFit: 'cover', // This ensures the image covers the area without stretching
    display: 'flex',
    alignItems: 'center', // This will vertically center the image
    justifyContent: 'center' // This will horizontally center the image
  },
  chatInfo: {
    flexGrow: 1,
    flexShrink: 2,
    flexBasis: 'auto',
    minWidth: 0,
    alignSelf: 'left',
    alignItems: 'center',
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
    maxWidth: '300px',
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
    // maxWidth: '300px',
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
    color: '#999'
  },
  person: {
    '--person-avatar-size': '32px',
    '--person-alignment': 'center'
  }
});

export const ChatListItem = ({ chat, myId, onSelected }: IMgtChatListItemProps & IChatListItemInteractionProps) => {
  const styles = useStyles();

  // shortcut if no valid user
  if (!myId) {
    return <></>;
  }

  // Copied and modified from the sample ChatItem.tsx
  // Determines the title in the case of 1:1 and self chats
  const inferTitle = (chatObj: Chat) => {
    if (myId && chatObj.chatType === 'oneOnOne' && chatObj.members) {
      const other = chatObj.members.find(m => (m as AadUserConversationMember).userId !== myId);
      const me = chatObj.members.find(m => (m as AadUserConversationMember).userId === myId);
      return other
        ? `${other?.displayName || (other as AadUserConversationMember)?.email || other?.id}`
        : `${me?.displayName} (You)`;
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
  const determineCorrectTimestamp = (chat: Chat) => {
    let timestamp: Date | undefined;

    // lastMessageTime is the time of the last message sent in the chat
    // lastUpdatedTime is Date and time at which the chat was renamed or list of members were last changed.
    const lastMessageTimeString = chat.lastMessagePreview?.createdDateTime as string;
    const lastUpdatedTimeString = chat.lastUpdatedDateTime as string;

    const lastMessageTime = new Date(lastMessageTimeString);
    const lastUpdatedTime = new Date(lastUpdatedTimeString);

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

    switch (true) {
      case chat.chatType === 'oneOnOne':
        const other = chatObj.members?.find(m => (m as AadUserConversationMember).userId !== myId);
        const otherAad = other as AadUserConversationMember;
        if (!otherAad) {
          return oneOnOneProfilePicture;
        }
        return (
          <Person
            className={styles.person}
            userId={otherAad?.userId!}
            avatarSize="small"
            showPresence={true}
            personCardInteraction={PersonCardInteraction.hover}
          />
        );
      case chat.chatType === 'group':
        return GroupProfilePicture;
      default:
        error(`Error: Unexpected chatType: ${chat.chatType}`);
        return oneOnOneProfilePicture;
    }
  };

  const removeHtmlPTags = (str: string) => {
    return str.replace(/<\/?p>/g, '');
  };

  const enrichPreviewMessage = (previewMessage: NullableOption<ChatMessageInfo> | undefined) => {
    let previewString = '';

    // handle general chats from people and bots
    if (previewMessage?.from?.user?.id === myId) {
      previewString = 'You: ' + previewMessage?.body?.content;
    } else if (previewMessage?.from?.user?.displayName) {
      previewString = previewMessage?.from?.user?.displayName + ': ' + previewMessage?.body?.content;
    } else if (previewMessage?.from?.application?.displayName) {
      previewString = previewMessage?.from?.application?.displayName + ': ' + previewMessage?.body?.content;
    }

    // handle all events
    if (previewMessage?.eventDetail) {
      previewString = previewMessage?.body?.content as string;
    }

    return removeHtmlPTags(previewString);
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
