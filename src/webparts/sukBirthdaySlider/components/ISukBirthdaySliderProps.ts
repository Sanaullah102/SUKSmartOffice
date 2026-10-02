export type BirthdayCardStyle =
  | 'compact'
  | 'classic'
  | 'welcome';

export type BirthdayRange =
  | 'all'
  | 'today'
  | 'next7'
  | 'next10'
  | 'next30'
  | 'currentMonth';

export interface IBirthdayPerson {
  id: number;
  displayName: string;
  jobTitle?: string;
  department?: string;
  birthday?: string;
  birthdayLabel?: string;
  dateLabel?: string;
  daysUntil?: number;
  photoUrl?: string;
  message?: string;
}

export interface ISukBirthdaySliderProps {
  people: IBirthdayPerson[];
  loading?: boolean;
  errorMessage?: string;

  styleVariant?: BirthdayCardStyle;

  heading?: string;
  emptyMessage?: string;

  autoplay?: boolean;
  interval?: number;
  cardsPerView?: number;
  showArrows?: boolean;
  showDots?: boolean;

  showBirthdayDate?: boolean;
  showJobTitle?: boolean;
  showDepartment?: boolean;
  showMessage?: boolean;

  cardBackgroundColor?: string;
  accentColor?: string;
  headingColor?: string;
  borderRadius?: number;
}
