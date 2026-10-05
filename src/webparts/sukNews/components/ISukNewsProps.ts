import { SPHttpClient } from '@microsoft/sp-http';

export type NewsSourceType = 'newsPages' | 'list';

export type NewsDisplayStyle =
  | 'cardGrid'
  | 'splitCards'
  | 'editorial'
  | 'featured'
  | 'compact';

export type NewsPromotedFilter = 'all' | 'promotedOnly' | 'excludePromoted';
export type NewsSortOrder = 'newest' | 'oldest' | 'titleAsc' | 'titleDesc';
export type NewsKeywordMatch = 'any' | 'all';
export type NewsImageFit = 'cover' | 'contain' | 'fill' | 'none' | 'scale-down';
export type NewsImagePosition = 'center' | 'top' | 'bottom' | 'left' | 'right';

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
  hasPromotedState: boolean;
  promotedFilter: NewsPromotedFilter;
  includeKeywords: string;
  excludeKeywords: string;
  includedCategories: string;
  excludedCategories: string;
  keywordMatch: NewsKeywordMatch;
  filterTextIn: 'all' | 'title' | 'description' | 'category';
  publishedWithinDays: number;
  includeFutureDated: boolean;
  skipItems: number;
  sortOrder: NewsSortOrder;
  fields: INewsFieldMappings;
  title: string;
  displayStyle: NewsDisplayStyle;
  itemLimit: number;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl?: string;
  emptyMessage: string;
  imageFit: NewsImageFit;
  imagePosition: NewsImagePosition;
  configurationError?: string;
}
