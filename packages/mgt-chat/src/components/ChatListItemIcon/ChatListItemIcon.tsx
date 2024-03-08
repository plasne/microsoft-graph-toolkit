import React from 'react';
import { Chat } from '@microsoft/microsoft-graph-types';
import { CalendarLtr24Regular, PeopleTeam24Regular, bundleIcon } from '@fluentui/react-icons';
import { error } from '@microsoft/mgt-element';
import { Circle } from '../Circle/Circle';
import { MgtTemplateProps } from '@microsoft/mgt-react';

const GroupIcon = bundleIcon(PeopleTeam24Regular, PeopleTeam24Regular);
const MeetingIcon = bundleIcon(CalendarLtr24Regular, CalendarLtr24Regular);

export const ChatListItemIcon = ({ chatType }: Chat & MgtTemplateProps): JSX.Element | null => {
  if (!chatType) return null;

  const iconColor = 'var(--colorBrandForeground2)';

  switch (chatType) {
    case 'meeting':
      return (
        <Circle>
          <MeetingIcon color={iconColor} />
        </Circle>
      );
    case 'group':
      return (
        <Circle>
          <GroupIcon color={iconColor} />
        </Circle>
      );
    default:
      error(`Attempted to render an icon for chat of type: ${chatType}`);
      return null;
  }
};
