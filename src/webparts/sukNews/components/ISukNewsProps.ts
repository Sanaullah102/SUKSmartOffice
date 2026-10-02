import { SPHttpClient } from '@microsoft/sp-http';

export type NewsSourceType = 'newsPages' | 'list';

export type NewsDisplayStyle =
  | 'cardGrid'
  | 'splitCards'
  | 'editorial'
  | 'featured'
  | 'compact';

export interface INewsFieldMappings {
  title: string;
  description?: string;
  publishDate?: string;
  image?: string;
  link?: string;
  category?: string;
}

export interface ISukNewsProps {
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  sourceType: NewsSourceType;
  sourceId?: string;
  filterPromotedNews: boolean;
  fields: INewsFieldMappings;
  title: string;
  displayStyle: NewsDisplayStyle;
  itemLimit: number;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl?: string;
  emptyMessage: string;
  configurationError?: string;
}
