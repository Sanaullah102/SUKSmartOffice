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

export interface IQuickLinkItem {
  id: string;
  text: string;
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

  sectionTitle: string;
  showSectionTitle: boolean;
  sectionBackgroundColor: string;
  sectionPadding: number;

  desktopColumns: number;
  tabletColumns: number;
  mobileColumns: number;
  gap: number;

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
