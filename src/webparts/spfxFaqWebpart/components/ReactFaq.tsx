import * as React from 'react';
import { IReactFaqProps } from './IReactFaqProps';
import { IFaqProp, IFaqServices } from '../../../interface';
import { ServiceScope, Environment, EnvironmentType } from '@microsoft/sp-core-library';
import { FontAwesomeIcon } from '@fortawesome/react-fontawesome';
import * as fontawesome from '@fortawesome/free-solid-svg-icons';
import Autosuggest from 'react-autosuggest';
import { FaqServices } from '../../../services/FaqServices';
import ReactHtmlParser from 'react-html-parser';
import './reactAccordion.css';


import {
  Accordion,
  AccordionItem,
  AccordionItemHeading,
  AccordionItemButton,
  AccordionItemPanel,
} from 'react-accessible-accordion';


import './index.css';
import ErrorBoundary from './ErrorBoundary';



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

// FAQ Class
export default class ReactFaq extends React.Component<IReactFaqProps, IFaqState> {

  private faqServicesInstance: IFaqServices;

  constructor(props) {
    super(props);

    this.state = {
      originalData: [],
      actualData: [],
      BusinessCategory: [],
      isLoading: true,
      errorCause: "No Data",
      selectedEntity: [],
      show: false,
      filterData: [],
      searchValue: "",
      filteredCategoryData: [],
      filteredQuestion: '',
      value: '',
      suggestions: [],
      actualCanvasContentHeight: 0,
      actualCanvasWrapperHeight: 0,
      actualAccordionHeight: 0
    };
    try {
      let serviceScope: ServiceScope;
      serviceScope = this.props.ServiceScope;
      if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
        // Mapping to be used when webpart runs in SharePoint.
        this.faqServicesInstance = serviceScope.consume(FaqServices.serviceKey);
      }//end if
    } catch (error) {
        console.log(error);
    }
  }


  public onHandleChange = (event, value, FaqData) => {
    if (FaqData.length > 0 && event !== undefined) {
      if (value === "") {
        const FaqFilteredData = this.filterByValue(FaqData, value);
        this.setState({ originalData: FaqFilteredData });
      }
      else {
        this.setState({ originalData: this.state.actualData });
      }
    }
  }


  public onChange = (event, { newValue }) => {
    if (newValue !== "") {
      this.setState({
        value: newValue,
      });
    }
    else {

      this.setState({
        originalData: this.state.actualData,
      });
    }
  }

  // On Suggestion Selected
  public onSuggestionSelected = (FaqData, event) => {
    const FaqFilteredData = this.filterByValue(FaqData, event.currentTarget.innerText);
    if (FaqFilteredData) {
      if (FaqFilteredData.length > 0) {
        const autoSuggestTextbox = document.getElementById("txtSearchBox") as HTMLTextAreaElement;
        autoSuggestTextbox.value = event.currentTarget.innerText;
        autoSuggestTextbox.blur();
        const FaqId = FaqFilteredData[0].Id;
        const FaqCategory = FaqFilteredData[0].Category;
        const catData = [];
        catData.push(FaqCategory);
        this.setState({ filteredCategoryData: catData });
        const nodElem = 'acc-' + FaqCategory;
        const node = document.getElementsByClassName(nodElem);
        const chNode = node[0].children[0].children[0].children[0];
        const newAttr = document.createAttribute('aria-expanded');
        newAttr.value = 'true';
        chNode.setAttributeNode(newAttr);
        node[0].children[0].children[1].removeAttribute('hidden');
        const FaqNode = this.getFaqElement(FaqId);
        const txtNode = document.getElementById("txtSearchBox");
        const FaqEle = FaqNode[0];
        const newAttrII = document.createAttribute('aria-expanded');
        newAttrII.value = 'true';
        FaqEle.setAttributeNode(newAttrII);
        FaqEle.nextSibling.style.display = 'block';
        FaqEle.nextSibling.removeAttribute('class');
        if (FaqEle.previousElementSibling.previousSibling.classList !== undefined) {
          FaqEle.previousElementSibling.previousSibling.classList.add("hideDiv");
        }
        else {
          // IE11 does not implement classList on <svg>
          let appliedClasses = FaqEle.previousElementSibling.previousSibling.getAttribute("class") || "";
          appliedClasses = appliedClasses.split(" ").indexOf("hideDiv") === -1
            ? appliedClasses + " hideDiv"
            : appliedClasses;
          FaqEle.previousElementSibling.previousSibling.setAttribute('class', appliedClasses);
        }
        if (FaqEle.previousElementSibling.classList !== undefined) {
          FaqEle.previousElementSibling.classList.remove("hideDiv");
        }
        else {
          // IE11 does not implement classList on <svg>
          let appliedClassesII = FaqEle.previousElementSibling.getAttribute("class") || "";
          appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
            ? appliedClassesII.replace(" hideDiv", "")
            : appliedClassesII + " hideDiv";
          FaqEle.previousElementSibling.setAttribute('class', appliedClassesII);
        }

        const txtSibEle = txtNode.nextElementSibling;
        txtSibEle.classList.remove("react-autosuggest__suggestions-container--open");
        FaqEle.scrollIntoView({ behavior: 'smooth' });

        if (document.getElementsByClassName("mainContent") !== undefined && document.getElementsByClassName("mainContent").length > 0) {
          this.setFaqWebPartHeightDynamic();
        }

      }


    }
  }

    public onSuggestionsFetchRequested = ({ value }) => {
      this.setState({
        suggestions: this.getSuggestions(value)
      });
    }

    public onSuggestionsClearRequested = () => {
      const autoSuggestTextbox = document.getElementById("txtSearchBox") as HTMLTextAreaElement;
      if(autoSuggestTextbox.value === ""){
        autoSuggestTextbox.value = "";
        this.setState({
          suggestions: [],
          value: ""
        });
      }
    }

    // When suggestion is clicked, Autosuggest needs to populate the input
  // based on the clicked suggestion. Teach Autosuggest how to calculate the
  // input value for every given suggestion.
  public getSuggestionValue = (suggestion) => {
    if (suggestion.length < 0) {
      return "";
    }
    else {
      return suggestion.Title;
    }
  }

    public getSuggestions = (value) => {
      const inputValue = value.trim().toLowerCase();
      const inputLength = inputValue.length;
      return inputLength === 0 ? [] : this.state.actualData.filter(lang =>
        (lang.Title.toLowerCase().indexOf(inputValue) !== -1) ||
        (lang.Answer.toLowerCase().indexOf(inputValue) !== -1)
      );
    }

    public renderSuggestion = (suggestion) => {
      return (
        <div>
          {suggestion.Title}
        </div>
      );
    }

    public setNodeValues = () => {
      const SPCanvasFirstParent = (document.getElementsByClassName("mainContent") !== undefined && document.getElementsByClassName("mainContent").length > 0) ? document.getElementsByClassName("SPCanvas")[0].parentElement.offsetHeight : 0;
      const SPCanvasSecondParent = (document.getElementsByClassName("mainContent") !== undefined && document.getElementsByClassName("mainContent").length > 0) ? document.getElementsByClassName("SPCanvas")[0].parentElement.parentElement.offsetHeight : 0;
      this.setState({
        actualCanvasContentHeight: SPCanvasFirstParent,
        actualCanvasWrapperHeight: SPCanvasSecondParent
      }, this.dynamicHeight);
    }

    // Component Did Mount
    public async componentDidMount() {
      if (Environment.type === EnvironmentType.SharePoint || Environment.type === EnvironmentType.ClassicSharePoint) {
        this.loadFaq();
      }
      else {
        //await this.loadMockFaq();
      }
      this.setState({
        actualAccordionHeight: (document.getElementsByClassName("accordion") !== undefined && document.getElementsByClassName("accordion").length > 0) ? document.getElementsByClassName("accordion")[0].parentElement.offsetHeight : 0
      });
      const ua = window.navigator.userAgent;
      const trident = ua.indexOf('Trident/');

      if (trident > 0) {
        // IE 11 => return version number
        const rv = ua.indexOf('rv:');
        if ((parseInt(ua.substring(rv + 3, ua.indexOf('.', rv)), 10)) < 12) {
          document.getElementById("txtSearchBox").style.paddingTop = '3px';
        }
      }
    }

    public async loadFaq() {
      await this.faqServicesInstance.getFaq(this.props.listName).then((FaqData: IFaqProp[]) => {
        try {
          this.setState(
            {
              actualData: FaqData,
              originalData: FaqData
            }
          );

        }
        catch (error) {
          console.log("Error Occurred :" + error);
        }

      });
    }

      // Sort Data
      public categoryAndQuestionSorting = (Data): any => {
        const result = [];
        // Get Distinct category for sorting Category
        const distCate = this.distinct(Data, "Category");
        // Sort alphabetically
        distCate.sort((c, d) =>  c.Category > d.Category ? 1: -1);

        //Sorting the FQA as per CategorySortOrder
        distCate.forEach((distCateItem) => {
          Data.map((item) => {
            if (distCateItem.Category.toLowerCase() === item.Category.toLowerCase()) {
              result.push(item);
            }
          });
        });
        //Sorting the FQA as per Category and then Question (Alphabetically)
        result.sort((a, b) =>  a.Category.localeCompare(b.Category) || a.Title.localeCompare(b.Title) );
        return result;
      }

      // Set distinct
      public distinct(items, prop): any {
        const unique = [];
        const distinctItems = [];
        for (const item of items) {
          if (unique[item[prop]] === undefined) {
            distinctItems.push(item);
          }

          unique[item[prop]] = 0;
        }
        return distinctItems;
      }

      public filterByValue = (arrayData, value) :any => {
          return arrayData.filter(o =>
          this.includes(o.Title.toLowerCase(), value.toLowerCase()) || this.includes(o.Answer.toLowerCase(), value.toLowerCase())
        );
      }

      public getFaqElement = (FaqId) => {
        return Array.prototype.filter.call(
          document.getElementsByTagName('span'),
          (el) => el.getAttribute('data-id') === FaqId.toString()
        );
      }

      // Format Modified Date from SharePoint
      public formatDate = (ModifiedDate): any => {
        const monthNames = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec'];
        const dt = new Date(ModifiedDate);
        let hours = dt.getHours();
        const minutes = dt.getMinutes();
        const secs = dt.getSeconds();
        const ampm = hours >= 12 ? 'PM' : 'AM';
        hours = hours % 12;
        hours = hours ? hours : 12; // the hour '0' should be '12'
        const strTime = hours + ':' + minutes + ':' + secs + ' ' + ampm;

        return monthNames[dt.getMonth()] + " " + dt.getDate() + ", " + dt.getFullYear() + " " + strTime;
      }


      public loadMoreEvent(event: any): void {
        const clickedId = event.target.getAttribute('data-id');
        console.log('clicked - ' + clickedId + ' ' + event.target);
        if (event.target.nodeName === "SPAN") {
          if (event.target.nextElementSibling.classList.contains("hideDiv")) {
            event.target.nextElementSibling.classList.remove("hideDiv");
            try {
              if (event.currentTarget.children[0].classList !== undefined) {
                event.currentTarget.children[0].classList.add("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClasses = event.currentTarget.children[0].getAttribute("class") || "";
                appliedClasses = appliedClasses.split(" ").indexOf("hideDiv") === -1
                  ? appliedClasses + " hideDiv"
                  : appliedClasses;
                event.currentTarget.children[0].setAttribute('class', appliedClasses);
              }

              if (event.currentTarget.children[1].classList !== undefined) {
                event.currentTarget.children[1].classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.currentTarget.children[1].getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.currentTarget.children[1].setAttribute('class', appliedClassesII);
              }
              event.currentTarget.children[3].removeAttribute("style");
            }
            catch (e) { console.log(e);}
          }
          else {
            event.target.nextElementSibling.classList.add("hideDiv");
            try {
              if (event.currentTarget.children[1].classList !== undefined) {
                event.currentTarget.children[1].classList.add("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClasses = event.currentTarget.children[1].getAttribute("class") || "";
                appliedClasses = appliedClasses.split(" ").indexOf("hideDiv") === -1
                  ? appliedClasses + " hideDiv"
                  : appliedClasses;
                event.currentTarget.children[1].setAttribute('class', appliedClasses);
              }

              if (event.currentTarget.children[0].classList !== undefined) {
                event.currentTarget.children[0].classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.currentTarget.children[0].getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.currentTarget.children[0].setAttribute('class', appliedClassesII);
              }
              event.currentTarget.children[3].removeAttribute("style");
            }
            catch (e) { console.log(e);}
          }
        }
        else {
          if (event.target.nodeName === "path") {
            if (event.currentTarget.children[1] !== undefined) {
              event.currentTarget.children[1].classList.add("hideDiv");
              event.currentTarget.children[0].classList.add("hideDiv");
            }
            else {
              // IE11 does not implement classList on <svg>
              let appliedClasses = event.currentTarget.children[0].getAttribute("class") || "";
              appliedClasses = appliedClasses + " hideDiv";
              event.currentTarget.children[0].setAttribute('class', appliedClasses);
              let appliedClassesII = event.currentTarget.children[1].getAttribute("class") || "";
              appliedClassesII = appliedClassesII + " hideDiv";
              event.currentTarget.children[1].setAttribute('class', appliedClassesII);
            }
            if (event.target.parentElement.getAttribute('data-icon') === "plus-square") {
              event.target.parentElement.nextElementSibling.nextElementSibling.nextElementSibling.classList.remove("hideDiv");

              if (event.currentTarget.children[1].classList !== undefined) {
                event.currentTarget.children[1].classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.currentTarget.children[1].getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.currentTarget.children[1].setAttribute('class', appliedClassesII);
              }
            }
            else {
              event.target.parentElement.nextElementSibling.nextElementSibling.classList.add("hideDiv");
              event.target.parentElement.nextElementSibling.nextElementSibling.removeAttribute("style");
               if (event.currentTarget.children[0].classList !== undefined) {
                event.currentTarget.children[0].classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.currentTarget.children[0].getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.currentTarget.children[0].setAttribute('class', appliedClassesII);
              }
            }
          }
          else if (event.target.nodeName === "svg") {
           if (event.target.classList !== undefined) {
              event.target.classList.add("hideDiv");
            }
            else {
              // IE11 does not implement classList on <svg>
              let appliedClasses = event.target.getAttribute("class") || "";
              appliedClasses = appliedClasses + " hideDiv";
              event.target.setAttribute('class', appliedClasses);
            }
            //alert('path');
            if (event.target.getAttribute('data-icon') === "plus-square") {
              event.target.nextElementSibling.nextElementSibling.nextElementSibling.classList.remove("hideDiv");
              //event.target.nextElementSibling.classList.remove("hideDiv");

              if (event.target.nextElementSibling.classList !== undefined) {
                event.target.nextElementSibling.classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.target.nextElementSibling.getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.target.nextElementSibling.setAttribute('class', appliedClassesII);
              }
            }
            else {
              event.target.nextElementSibling.nextElementSibling.classList.add("hideDiv");
              event.target.nextElementSibling.nextElementSibling.removeAttribute("style");
              if (event.target.previousElementSibling.classList !== undefined) {
                event.target.previousElementSibling.classList.remove("hideDiv");
              }
              else {
                // IE11 does not implement classList on <svg>
                let appliedClassesII = event.target.previousElementSibling.getAttribute("class") || "";
                appliedClassesII = appliedClassesII.split(" ").indexOf("hideDiv") !== -1
                  ? appliedClassesII.replace(" hideDiv", "")
                  : appliedClassesII + " hideDiv";
                event.target.previousElementSibling.setAttribute('class', appliedClassesII);
              }
            }

          }
          else {
            if (event.target.getAttribute('data-icon') === "plus-square") {
              event.target.nextElementSibling.nextElementSibling.nextElementSibling.classList.remove("hideDiv");
              event.target.nextElementSibling.classList.remove("hideDiv");
              event.target.classList.add("hideDiv");
            }
            else {
              event.target.nextElementSibling.nextElementSibling.classList.add("hideDiv");
              event.target.previousElementSibling.classList.add("hideDiv");
              event.target.classList.add("hideDiv");
              event.target.removeAttribute("style");
            }
          }
        }
        if (document.getElementsByClassName("mainContent") !== undefined && document.getElementsByClassName("mainContent").length > 0) {
          this.setFaqWebPartHeightDynamic();
        }
      }

      public dynamicHeight = (): any => {
        const SPCanvasNode = document.getElementsByClassName("SPCanvas");
        const accordionNode = document.getElementsByClassName("accordion");
        if (SPCanvasNode.length > 0 && accordionNode.length > 0) {
          SPCanvasNode[0].parentElement.style.height = (this.state.actualCanvasContentHeight + (accordionNode[0].parentElement.offsetHeight - this.state.actualAccordionHeight)) + "px";
          SPCanvasNode[0].parentElement.parentElement.style.height = (this.state.actualCanvasWrapperHeight + (accordionNode[0].parentElement.offsetHeight - this.state.actualAccordionHeight)) + "px";
        }
      }

      public setFaqWebPartHeightDynamic = (): any => {
        if (this.state.actualCanvasContentHeight === 0) {
          this.setNodeValues();
        }
        else {
          this.dynamicHeight();
        }
      }

      // On Accordion Change
      public accordionOnchange = (): any => {
        if (document.getElementsByClassName("mainContent") !== undefined && document.getElementsByClassName("mainContent").length > 0) {
          this.setFaqWebPartHeightDynamic();
        }
      }

      // Includes
      public includes = (container, value): boolean => {
        let returnValue = false;
        const pos = container.indexOf(value);
        if (pos >= 0) {
          returnValue = true;
        }
        return returnValue;
      }

      /*
        Handle the rendering of the Link
      */
      public showLink = (linkObj?: any): React.ReactElement => {

        if(linkObj !== null){
            if(linkObj.Description !== null){
                return <div><a href={linkObj.Url} target="_blank" rel="noreferrer">{linkObj.Description}</a></div>
            }else //just show link
            {
                return <div><a href={linkObj.Url} target="_blank" rel="noreferrer">{linkObj.Url}</a></div>
            }//end if
        }
      }

  public render(): React.ReactElement<IReactFaqProps> {
    let uniqueBC = [];
    let FaqData = [];

    if (this.state.originalData.length > 0) {
      FaqData = this.categoryAndQuestionSorting(this.state.originalData);
      uniqueBC = this.distinct(FaqData, "BusinessCategory");
    }

    const { value, suggestions } = this.state;

    // Autosuggest will pass through all these props to the input.
    const inputProps = {
      placeholder: 'Search Frequently Asked Questions',
      value,
      onChange: this.onChange,
      id: 'txtSearchBox'
    };

    return (
      <div className={`container`}>

        <div className="FaqSearchBox" accept-charset="UTF-8">
          <Autosuggest
            suggestions={suggestions}
            onSuggestionsFetchRequested={this.onSuggestionsFetchRequested}
            onSuggestionsClearRequested={this.onSuggestionsClearRequested}
            getSuggestionValue={this.getSuggestionValue}
            renderSuggestion={this.renderSuggestion}
            onSuggestionSelected={this.onSuggestionSelected.bind(this, this.state.actualData)}
            inputProps={inputProps}
            focusInputOnSuggestionClick={false}

          />
        </div>
        <ErrorBoundary>

          <div className="clearBody">

            <Accordion allowMultipleExpanded={true} allowZeroExpanded={true} onChange={this.accordionOnchange.bind(this)} preExpanded={this.state.filteredCategoryData}
            >
              {uniqueBC.map((item) => (
                <div key={item.id}>
                  {this.distinct(FaqData, "Category").map((allCat) => (
                    <div key={`acc-${allCat.Category}`} className={`acc-${allCat.Category}`}>
                      <AccordionItem uuid={allCat.Category}>
                        <AccordionItemHeading>
                          <AccordionItemButton >
                            {allCat.Category}
                          </AccordionItemButton>
                        </AccordionItemHeading>
                        <AccordionItemPanel>
                          <div className="acc-item-panel">
                            {FaqData.filter(it => it.Category === allCat.Category).map((allFaq) => (

                              <div key={`acc-item-${allFaq.Id}`}
                                id={`acc-item-${allFaq.Id}`}
                                className="acc-item"
                                data-id={allFaq.Id}
                                onClick={
                                  event => this.loadMoreEvent(event)
                                }>
                                <FontAwesomeIcon icon={fontawesome.faPlusSquare} size="1x" data-id={allFaq.Id} className={"plusminusImg"} />
                                <FontAwesomeIcon icon={fontawesome.faMinusSquare} size="1x" data-id={allFaq.Id} className={"plusminusImg hideDiv"} />
                                <span id={`acc-span-text-${allFaq.Id}`} className="acc-span-text" data-id={allFaq.Id}>{allFaq.Title}</span>
                                <div className="hideDiv">
                                  <span className="acc-modified-text">Last Modified : {this.formatDate(allFaq.Modified)}</span>
                                  <div className="acc-answer">
                                    {ReactHtmlParser(allFaq.Answer)}
                                    {this.showLink(allFaq.Link)}
                                  </div>
                                </div>
                              </div>

                            ))}
                          </div>
                        </AccordionItemPanel>
                      </AccordionItem>
                    </div>
                  ))}
                </div>
              ))}
            </Accordion>
          </div>
        </ErrorBoundary>
      </div>
    );

  }
}
