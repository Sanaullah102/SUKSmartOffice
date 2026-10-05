import { SPHttpClient } from '@microsoft/sp-http';

export type ActivityDisplayStyle =
  | 'agenda'
  | 'cards'
  | 'timeline'
  | 'compact'
  | 'dateFocus';

export interface IActivityFieldMappings {
  title: string;
  startDate: string;
  endDate?: string;
  description?: string;
  location?: string;
  category?: string;
  link?: string;
  organizer?: string;
}

export interface ISukActivityDisplayProps {
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  listId?: string;
  mappings: IActivityFieldMappings;
  title: string;
  displayStyle: ActivityDisplayStyle;
  seeAllUrl?: string;
  addNewUrl?: string;
  showSeeAll: boolean;
  showAddButton: boolean;
  itemLimit: number;
  emptyMessage: string;
}
