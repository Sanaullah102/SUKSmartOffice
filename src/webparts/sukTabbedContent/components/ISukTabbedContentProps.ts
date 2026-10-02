import { SPHttpClient } from '@microsoft/sp-http';
import {
  BirthdayListService
} from '../../sukBirthdaySlider/services/BirthdayListService';

export type SukTabbedPanelType =
  | 'textLinks'
  | 'documentLibrary'
  | 'sukComponent';

export type SukTabbedComponentType =
  | 'announcements'
  | 'bannerSlider'
  | 'birthdaySlider'
  | 'calendar'
  | 'customLayout'
  | 'gallery'
  | 'quickLinks'
  | 'welcomeUser';

export interface ISukTabbedLink {
  text: string;
  url: string;
  description?: string;
}

export interface ISukDocumentTabSettings {
  libraryId: string;
  filterField?: string;
  filterValue?: string;
  itemLimit?: number;
  emptyMessage?: string;
}

export interface ISukTabbedPanel {
  id: string;
  title: string;
  type: SukTabbedPanelType;
  seeAllUrl?: string;
  seeAllText?: string;
  content?: string;
  links?: ISukTabbedLink[];
  documentLibrary?: ISukDocumentTabSettings;
  component?: SukTabbedComponentType;
  settings?: { [key: string]: unknown };
}

export interface ISukTabbedContentProps {
  tabs: ISukTabbedPanel[];
  heading: string;
  seeAllText: string;
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  birthdayListService?: BirthdayListService;
  configurationError?: string;
}
