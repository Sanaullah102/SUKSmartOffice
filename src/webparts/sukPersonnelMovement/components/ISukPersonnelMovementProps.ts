import { SPHttpClient } from '@microsoft/sp-http';

export type PersonnelMovementStyle =
  | 'table'
  | 'modernTable'
  | 'cards'
  | 'timeline'
  | 'compact';

export interface IPersonnelMovementMappings {
  name: string;
  grade?: string;
  jobTitle?: string;
  placement?: string;
  date?: string;
  category?: string;
  link?: string;
}

export interface ISukPersonnelMovementProps {
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  listId?: string;
  mappings: IPersonnelMovementMappings;
  title: string;
  style: PersonnelMovementStyle;
  itemLimit: number;
  placementLabel: string;
  placementValue: string;
  retirementLabel: string;
  retirementValue: string;
  showCategories: boolean;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl?: string;
  emptyMessage: string;
  configurationError?: string;
}
