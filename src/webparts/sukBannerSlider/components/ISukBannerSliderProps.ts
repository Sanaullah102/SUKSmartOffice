export type BannerTextPosition = 'left' | 'center' | 'right';
export type BannerSourceType = 'library' | 'list';

export interface IBannerSlide {
  id?: number;
  title: string;
  description?: string;
  imageUrl: string;
  linkUrl?: string;
  buttonText?: string;
  openInNewTab?: boolean;
  textPosition?: BannerTextPosition;
  overlayOpacity?: number;
  altText?: string;
}

export interface ISukBannerSliderProps {
  slides: IBannerSlide[];
  loading?: boolean;
  errorMessage?: string;

  autoplay?: boolean;
  interval?: number;
  height?: number;
  pauseOnHover?: boolean;
  showArrows?: boolean;
  showDots?: boolean;
  showTitle?: boolean;
  showDescription?: boolean;
  showButton?: boolean;
  defaultOverlayOpacity?: number;
}
