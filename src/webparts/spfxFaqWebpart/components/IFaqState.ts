import { IFaqProp } from '../../../interface';

export interface IFaqState {
  originalData: IFaqProp[];
  actualData: IFaqProp[];
  BusinessCategory: any;
  isLoading: boolean;
  errorCause: string;
  selectedEntity: any;
  show: boolean;
  filterData: any;
  searchValue: string;
  filteredCategoryData: any;
  filteredQuestion: string;
  value: string;
  suggestions: any;
  actualCanvasContentHeight: number;
  actualCanvasWrapperHeight: number;
  actualAccordionHeight: number;
}