import { SPHttpClient } from '@microsoft/sp-http';

export interface ISukAnnouncementItem {
  Id: number;
  title: string;
  description: string;
  date: string;
  category?: string;
}

export interface ISukAnnouncementsProps {
  webPartTitle: string;

  webAbsoluteUrl: string;

  spHttpClient: SPHttpClient;

  listId: string;

  titleField: string;

  descriptionField: string;

  dateField: string;

  categoryField?: string;

  itemLimit: number;

  clickMode: 'view' | 'edit';

  showSeeAll: boolean;
}