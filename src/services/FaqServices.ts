import { ServiceScope, ServiceKey } from '@microsoft/sp-core-library';
import { PageContext } from '@microsoft/sp-page-context';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { IFaqServices, IFaqProp} from '../interface';

export class FaqServices implements IFaqServices {


  public static readonly serviceKey: ServiceKey<IFaqServices> = ServiceKey.create<IFaqServices>('vrd:IFaqServices', FaqServices);
  private _spHttpClient: SPHttpClient;
  private _pageContext: PageContext;
  private _currentWebUrl: string;

  constructor(serviceScope: ServiceScope) {
    serviceScope.whenFinished(() => {
      this._spHttpClient = serviceScope.consume(SPHttpClient.serviceKey);
      this._pageContext = serviceScope.consume(PageContext.serviceKey);
      this._currentWebUrl = this._pageContext.web.absoluteUrl;
    });

  }


  public getFaq(listName): Promise<IFaqProp[]> {
    return new Promise<IFaqProp[]>((resolve: any) => {
      const ParentDetails = this.getFaqs(listName);
      resolve(ParentDetails);
    });
  }

  public async getMockFaq(): Promise<any> {
    const tempOrg = [{
      Id: 1,
      Title: "What is the HR Policy?",
      Answer: "There is no change in HR Policy",
      Category: "HR Policy",
      CategorySortOrder: 3,
      QuestionSortOrder: 3,
      Modified: '2020-03-27T11:07:21Z'
    },
    {
      Id: 2,
      Title: "What changes should I expect (or not) as an employee?",
      Answer: "For the immediate future, There is no change.",
      Category: "Top Questions",
      CategorySortOrder: 2,
      QuestionSortOrder: 2,
      Modified: '2020-03-27T11:07:21Z'
    },
    {
      Id: 3,
      Title: "What is the finance policy in the company?",
      Answer: "There is change in Finance Policy. ",
      Category: "Finance Policy",
      CategorySortOrder: 3,
      QuestionSortOrder: 1,
      Modified: '2020-03-27T11:07:21Z'
    }
    ];
    return tempOrg;
  }

  public async getFaqs(listName: string): Promise<IFaqProp[]> {
    try {
      const FaqProp:IFaqProp[] = [];
      let restUrl: string = this._currentWebUrl;
      //fix: load more than 100 items using top=5000
      //restUrl += "/_api/web/lists/getbytitle('" + listName + "')/items?$select=Id,Title,Answer,FAQCategory,CategorySortOrder,QuestionSortOrder,Modified,Link&$top=5000";
      restUrl += "/_api/web/lists/getbytitle('" + listName + "')/items?$select=Id,Title,Answer,FAQCategory,Modified,Link&$top=5000";
      return await this._spHttpClient.get(restUrl, SPHttpClient.configurations.v1,
        {
          headers: {
            "Accept": "application/json;odata=nometadata",
            "odata-version": "3.0"
          }
        })
        .then((response: SPHttpClientResponse) => {
          return response.json().then((responseFormatted: any) => {
            if (response.ok) {
              const collection = responseFormatted.value;
              console.log('hi');
              for (let i = 0; i < collection.length; i++) {
                FaqProp.push({
                  Id: collection[i].Id,
                  Title: collection[i].Title,
                  Answer: collection[i].Answer,
                  Category: collection[i].FAQCategory,
                  CategorySortOrder: 1,//collection[i].CategorySortOrder,
                  QuestionSortOrder: 1,//collection[i].QuestionSortOrder,
                  Modified: collection[i].Modified,
                  Link: collection[i].Link
                });
              }// end for
            }
            else {
              throw new Error(response.text().toString());
            }

            // Now Sort results by Category and then Question alphabetically
            FaqProp.sort((a, b) =>  a.Category.localeCompare(b.Category) || a.Title.localeCompare(b.Title) );              

            console.log(FaqProp);
            console.log('--return');
            return FaqProp;
          });
        });

    }
    catch (error) {
      console.log("Service API Error - " + error);
    }
  }
}
