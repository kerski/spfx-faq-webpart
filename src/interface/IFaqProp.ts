export interface IFaqProp {
    Id?: number;
    Title?: string;
    Answer?: string;
    BusinessCategory?: string;
    Category?: string;
    CategorySortOrder?: number;
    QuestionSortOrder?: number;
    IsFullRow?: string;
    expandRow?: boolean;
    Modified?: Date;
    Link?: {
        Description?: string;
        Url?: string;
    }
  }
  