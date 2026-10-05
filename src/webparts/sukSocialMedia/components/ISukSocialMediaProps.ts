export type SocialMediaStyle =
  | 'iconStrip'
  | 'socialCards'
  | 'featured'
  | 'embedFeed'
  | 'compactList';

export interface ISocialAccount {
  platform: string;
  label: string;
  url: string;
  handle?: string;
  description?: string;
  group?: string;
  embedUrl?: string;
  enabled?: boolean;
}

export interface ISukSocialMediaProps {
  title: string;
  description?: string;
  accountsJson: string;
  style: SocialMediaStyle;
  columns: number;
  showGroupTabs: boolean;
  allTabLabel: string;
  showDescriptions: boolean;
  showHandles: boolean;
  openInNewTab: boolean;
  embedHeight: number;
}
