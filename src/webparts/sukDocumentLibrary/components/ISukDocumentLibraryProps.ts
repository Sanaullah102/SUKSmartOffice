import { SPHttpClient } from '@microsoft/sp-http';

export type DocumentDisplayStyle = 'grid' | 'list' | 'table' | 'cards' | 'compact';
export type DocumentContentMode = 'files' | 'folders' | 'both';
export type FolderIconMode = 'default' | 'random';

export interface ISukDocumentLibraryProps {
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  libraryId?: string;
  libraryRootUrl?: string;
  title: string;
  description?: string;
  displayStyle: DocumentDisplayStyle;
  contentMode: DocumentContentMode;
  folderIconMode: FolderIconMode;
  itemLimit: number;
  showSearch: boolean;
  searchPlaceholder: string;
  searchButtonText: string;
  showBreadcrumbs: boolean;
  allowFolderNavigation: boolean;
  recursive: boolean;
  groupByField?: string;
  sortField: string;
  sortDirection: 'asc' | 'desc';
  titleField?: string;
  descriptionField?: string;
  showModifiedDate: boolean;
  showFileSize: boolean;
  showGroupTabs: boolean;
  allTabLabel: string;
  emptyMessage: string;
  configurationError?: string;
}
