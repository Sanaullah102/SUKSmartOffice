import { IWelcomePerson } from '../../sukBirthdaySlider/services/BirthdayListService';

export interface ISukWelcomeUserProps {
  people: IWelcomePerson[];
  loading: boolean;
  errorMessage?: string;
  heading: string;
  emptyMessage: string;
  columns: number;
}
