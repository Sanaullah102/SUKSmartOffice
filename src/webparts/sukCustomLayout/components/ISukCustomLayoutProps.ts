export type SukModuleType =
  | 'empty'
  | 'text'
  | 'image'
  | 'hero'
  | 'slider'
  | 'quickLinks';

export interface ISukQuickLink {
  title: string;
  url: string;
  iconUrl?: string;
}

export interface ISukColumnConfig {
  heading?: string;

  moduleType: SukModuleType;

  width?: number;

  padding?: number;

  backgroundColor?: string;

  description?: string;

  imageUrl?: string;

  linkUrl?: string;

  buttonText?: string;

  /**
   * Supports:
   * - newline separated image URLs
   * - comma separated image URLs
   * - JSON string array
   * - string[]
   */
  sliderImages?: string | string[];

  /**
   * Supports:
   * - JSON array
   * - "Title|Url|IconUrl" one per line
   * - ISukQuickLink[]
   */
  quickLinks?: string | ISukQuickLink[];
}

export interface ISukCustomLayoutProps {
  columnCount: number;

  outerPadding?: number;

  columnGap?: number;

  borderRadius?: number;

  backgroundColor?: string;

  minHeight?: number;

  mobileBreakpoint?: number;

  verticalAlignment?: string;

  column1: ISukColumnConfig;

  column2?: ISukColumnConfig;

  column3?: ISukColumnConfig;
}