import { SPHttpClient } from '@microsoft/sp-http';

export type GalleryContentMode = 'folders' | 'pictures';

export type GalleryTemplate =
  | 'feature'
  | 'editorial'
  | 'grid'
  | 'masonry'
  | 'filmstrip';

export interface IGalleryItem {
  id: number;
  name: string;
  title: string;
  description?: string;
  imageUrl: string;
  modified?: string;
  itemCount?: number;
  folderUrl?: string;
}

export interface ISukGalleryProps {
  webAbsoluteUrl: string;
  spHttpClient: SPHttpClient;
  libraryId: string;
  rootFolderUrl: string;
  folderUrl?: string;
  titleField?: string;
  descriptionField?: string;
  galleryTitle: string;
  contentMode: GalleryContentMode;
  itemLimit: number;
  template: GalleryTemplate;
  columns: number;
  showSeeAll: boolean;
  seeAllText: string;
  showCounts: boolean;
  showCaptions: boolean;
  enableLightbox: boolean;
  pictureBorderWidth: number;
  pictureBorderColor: string;
  pictureBorderRadius: number;
}
