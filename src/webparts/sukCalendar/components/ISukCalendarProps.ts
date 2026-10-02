import {
  AadHttpClient,
  SPHttpClient
} from '@microsoft/sp-http';

export type CalendarEventSource =
  | 'sharePointList'
  | 'outlookCalendar';

export interface ICalendarEvent {
  id: string;
  title: string;
  start: Date;
  end?: Date;
  description?: string;
  location?: string;
  category?: string;
  url?: string;
  allDay?: boolean;
}

export interface ISukCalendarProps {
  source: CalendarEventSource;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  graphClient?: AadHttpClient;
  listId?: string;
  titleField?: string;
  startField?: string;
  endField?: string;
  descriptionField?: string;
  locationField?: string;
  categoryField?: string;
  weekStartsOn: number;
  title: string;
  seeAllUrl?: string;
}
