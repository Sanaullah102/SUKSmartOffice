export type QuickLinksHoverEffect =
  | 'none'
  | 'lift'
  | 'scale'
  | 'background'
  | 'border';

export type QuickLinksTextAlign =
  | 'left'
  | 'center'
  | 'right';

export type QuickLinksDisplayStyle =
  | 'tiles'
  | 'iconStrip'
  | 'serviceGrid'
  | 'list'
  | 'pills'
  | 'compact';

export interface IQuickLinkItem {
  id: string;
  text: string;
  description?: string;
  iconUrl: string;
  linkUrl: string;
  openInNewTab: boolean;

  overrideBackgroundColor: boolean;
  backgroundColor: string;

  overrideTextColor: boolean;
  textColor: string;
}

export interface ISukQuickLinksProps {
  items: IQuickLinkItem[];
  isEditMode: boolean;
  displayStyle: QuickLinksDisplayStyle;

  sectionTitle: string;
  showSectionTitle: boolean;
  showSeeAll: boolean;
  seeAllText: string;
  seeAllUrl: string;
  sectionBackgroundColor: string;
  sectionPadding: number;

  desktopColumns: number;
  tabletColumns: number;
  mobileColumns: number;
  gap: number;
  iconPosition: 'top' | 'left';
  showLinkBorder: boolean;

  cardMinHeight: number;
  cardPadding: number;
  cardBackgroundColor: string;
  cardBorderColor: string;
  cardBorderWidth: number;
  cardBorderRadius: number;
  cardShadow: boolean;

  iconSize: number;
  iconBoxSize: number;
  imageFit: 'contain' | 'cover';

  textSize: number;
  textColor: string;
  textWeight: number;
  textAlign: QuickLinksTextAlign;

  hoverEffect: QuickLinksHoverEffect;
  hoverBackgroundColor: string;
  hoverBorderColor: string;
  hoverScalePercent: number;
  hoverShadow: boolean;

  onItemsChanged: (items: IQuickLinkItem[]) => void;
  onGlobalColorsChanged: (
    backgroundColor: string,
    textColor: string
  ) => void;
}
