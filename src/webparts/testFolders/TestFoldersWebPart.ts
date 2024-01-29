import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import type { IReadonlyTheme } from '@microsoft/sp-component-base';
import { escape } from '@microsoft/sp-lodash-subset';

import styles from './TestFoldersWebPart.module.scss';
import * as strings from 'TestFoldersWebPartStrings';
import { SPComponentLoader } from '@microsoft/sp-loader';

import {
  SPHttpClient,
  SPHttpClientResponse,
} from "@microsoft/sp-http";
import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { LogLevel, PnPLogging } from "@pnp/logging";
//import * as _ from 'lodash';

require("bootstrap");

export interface ISPLists {
  value: ISPList[];
}

export interface ISPList {
  Id : string;
  Title : string;
}

export interface ITestFoldersWebPartProps {
  description: string;
  dataResults: any[];
  siteTitle: string;
  completeFlag: boolean;
  URL:string;
  tenantURL: any;
  dcURL: string;   
  siteName: string;
  siteID: any;
  siteShortName: string;
  divisionName: string[];
  divisionTitle: string;
  teamName: string;
  libraryName: string;
  siteArray: any;
  dcDivisions: string[]; 
  folderArray: string[];
  subFolder1Array: any[];
  subFolder2Array: any[];
  subFolder3Array: any[];
  isPowerUser: boolean;
  //dataFlag : boolean;
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  //private _isDarkTheme: boolean = false;
  //private _environmentMessage: string = '';

  public async render(): Promise<void> {

    this.properties.URL = this.context.pageContext.web.absoluteUrl;
    this.properties.tenantURL = this.properties.URL.split('/',5);
    const siteSNArray : any[] = this.properties.URL.split('_',2);
    this.properties.siteShortName = siteSNArray[1];
    this.properties.siteTitle = this.context.pageContext.web.title;
    this.properties.siteID = this.context.pageContext.site.id;
    this.properties.siteArray = this.properties.siteTitle.split(" - ");
    this.properties.divisionTitle = this.properties.siteArray[0];
    this.properties.siteName = this.properties.siteArray[1];
    this.properties.completeFlag = false;
    this.properties.isPowerUser = true;

    const uri = this.context.pageContext.web.absoluteUrl;

    console.log("Sitetitle",this.properties.siteTitle);
    console.log("Sitename",this.properties.siteName);
    console.log("SiteID",this.properties.siteID);
    console.log("Division",this.properties.divisionTitle);
    console.log("URL",this.properties.URL);
    console.log("siteSNArray",siteSNArray);
    console.log("siteArray",this.properties.siteArray);
    console.log("tenantURL",this.properties.tenantURL);

    this.domElement.innerHTML = `
    <section class="${styles.testFolders} ${!!this.context.sdks.microsoftTeams ? styles.teams : ''}">
      <div class="${styles.welcome}"
        <div>siteTitle = ${this.properties.siteTitle}</div>
        <div>siteName = ${this.properties.siteName}</div>
        <div>siteShortName = ${this.properties.siteShortName}</div>
        <div>siteID = ${this.properties.siteID}</div>
        <div>divisionTitle = ${this.properties.divisionTitle}</div>
        <div>tenant URL = ${uri.split('/',3)[2]}</div>
      </div>

      <div class="row titleRow text-white rounded"> 
        ${
          this.properties.isPowerUser
            ? `<div class="col-1"><a href="${this.properties.dcURL}" title="open Document Centre" target="_blank"><h3 class="text-white"><i class="bi bi-collection"></i></h3></a></div>`
            : `<div class="col-1"><p>.</div>`
        }
        <div class="col-11">
          <div class="row">
            <div class="col-6 welcomeText">
              <h5 class="text-white">Welcome to ${escape(this.properties.siteTitle)} Documents
                ${[/*<mgt-person person-query="me"></mgt-person>*/]}
              </h5>
            </div>
            <div class="col-5 poweruserText">
              <h6>${
                this.properties.isPowerUser
                  ? `<i class="bi bi-person-circle"></i> (Power User)`
                  : ""
              }</h6>
            </div>
          </div>
        </div>
      </div>
      
      <div class="container v-scrollbar">
        <div class="row" style="height:50px" id="headerBar"></div>
        <div class="row">
          <div class="col-4 libraryContainer">
            <div class="d-flex mt-1 align-items-start">
              <div class="nav flex-column nav-pills me-3 libraryList" id="libraryTabs" role="tablist" aria-orientation="vertical"></div> 
            </div>
          </div>

          <div class="col-8 tab-content" id="v-pills-tabContent">
            <div class="tab-pane fade" id="policies" role="tabpanel" aria-labelledby="policies"> 
              <div class="row">
                <div class="col-auto" id="policiesFolders"></div>
                <div class="col" id="policiesFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="procedures" role="tabpanel" aria-labelledby="procedures">
              <div class="row">
                <div class="col-auto" id="proceduresFolders"></div>
                <div class="col" id="proceduresFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="guides" role="tabpanel" aria-labelledby="guides">
              <div class="row">
                <div class="col-6" id="guidesFolders"></div>
                <div class="col" id="guidesFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="forms" role="tabpanel" aria-labelledby="forms">
              <div class="row">
                <div class="col-auto" id="formsFolders"></div>
                <div class="col" id="formsFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="general" role="tabpanel" aria-labelledby="general">
              <div class="row">
                <div class="col-auto" id="generalFolders"></div>
                <div class="col" id="generalFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="management" role="tabpanel" aria-labelledby="management">
              <div class="row">
                <div class="col-auto" id="managementFolders"></div>
                <div class="col" id="managementFiles"></div>
              </div>               
            </div>
            <div class="tab-pane fade" id="custom" role="tabpanel" aria-labelledby="custom">
              <div class="row">
                <div class="col-auto" id="customFolders"></div>
                <div class="col" id="customFiles"></div>
              </div>               
            </div>                                                                  
          </div>
        </div>  
      </div>
    </section>`;

    //document.getElementById('policies_btn')?.addEventListener("click", (_e:Event) => this._getData('Policies',1,""));
    //document.getElementById('procedures_btn')?.addEventListener("click",(_e:Event) => this._getData('Procedures',2,""));
    //document.getElementById('guides_btn')?.addEventListener("click",(_e:Event) => this._getData('Guides',3,""));
    //document.getElementById('forms_btn')?.addEventListener("click",(_e:Event) => this._getData('Forms',4,""));
    //document.getElementById('general_btn')?.addEventListener("click",(_e:Event) => this._getData('General',5,""));
    await this.getLibraryTabs();
  }


  private async getLibraryTabs(): Promise<void> {
    const library = ["Policies", "Procedures", "Guides", "Forms", "General"];
    const listContainer: Element | null = this.domElement.querySelector("#libraryTabs");
    let html: string = "";

    try {
      // *** check if user is in Managers Security Group
      //await this.checkManagerPermission();
      //if (this.properties.isManager) {
      //  library.push("Management");
      //}

      for (let x = 0; x < library.length; x++) {

        let libraryFlag = await this.checkDataAsync(library[x],this.properties.siteName,"");
        console.log(library[x],libraryFlag);

        const dataTarget:string=library[x].toLowerCase();
        
        if(libraryFlag){
          if (this.properties.isPowerUser) {

          html += `<button class="btn libraryBtn nav-link text-left mb-1" id="${library[x]}_btn" data-bs-toggle="pill" data-bs-target="#${dataTarget}" type="button" role="tab" aria-controls="${library[x]}" aria-selected="true">
                        <div class="col-1 libraryUploadIcon">
                        <a href="${this.properties.dcURL}/${library[x]}/Forms/AllItems.aspx??FilterField1=DC%5FTeam&FilterValue1=${this.properties.siteName}&FilterType1=TaxonomyFieldType&viewid=a8a5315c%2Df2f1%2D457f%2Db45e%2D9540e01e1ae3" target="_blank">                         
                          <h3 class="text-white"><i class="bi bi-cloud-arrow-up"></i></h3>
                        </a>
                        </div>
                        <div class="col-11 libraryName">
                        <h6 class="libraryText">${library[x]}</h6>
                        </div>
                    </button>`;
          } else {
            html += `<button class="btn libraryBtn nav-link text-left mb-1" id="${library[x]}Tab" data-bs-toggle="pill" data-bs-target="#${library[x]}" type="button" role="tab" aria-controls="${library[x]}" aria-selected="true">${library[x]}</button>`;
          }
        }
      }

      if(listContainer){listContainer.innerHTML = html;}

      // *** get custom tabs from termstore and add library column
      await this.renderCustomTabsAsync();
      //await this.setLibraryListeners();
    } catch (err) {
      //await this.addError(this.properties.siteName,"getLibraryTabs",err.message);
    }
    return;
  }

  private async checkDataAsync(library:string,team:string,category:string):Promise<boolean> {
     
    let dcName : string = "";      
    let dataFlag : boolean = false;

    switch(this.properties.divisionTitle){
      case "Assessments":
        dcName = "asm_dc";
        break;
      case "Central":
        dcName = "cen_dc";
        break;
      case "Connect":
        dcName = "cnn_dc";
        break;
      case "Employability":
        dcName = "emp_dc";
        break;
      case "Health":
        dcName = "hea_dc";
        break;
    }

    await this.checkData(dcName,library,team,category)
      .then((response) => {
        console.log("Value",response.value.length);
        if(response.value.length>0){
          dataFlag = true; 
        }
      });
      return dataFlag;
  }
  
  private async checkData(dcName:string,library:string,team:string,category:string):Promise<ISPLists> {
    
    let requestUrl = '';
    
    if(category === ''){
      requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$filter=TaxCatchAll/Term eq '${team}'&$top=10`;
    }else{
      requestUrl=`https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$filter=TaxCatchAll/Term eq '${category}'&$top=10`;
    }
    return this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response : SPHttpClientResponse) => {
        return response.json();
      });   
  }

/*
  private async checkData(library:string,team:string,category:string):Promise<boolean> {
    
    let dcName : string = "";
    let dataFlag : boolean = false;
    
    switch(this.properties.divisionTitle){
      case "Assessments":
        dcName = "asm_dc";
        break;
      case "Central":
        dcName = "cen_dc";
        break;
      case "Connect":
        dcName = "cnn_dc";
        break;
      case "Employability":
        dcName = "emp_dc";
        break;
      case "Health":
        dcName = "hea_dc";
        break;
    }

    const requestUrl = `https://${this.properties.tenantURL[2]}/sites/${dcName}/_api/web/lists/GetByTitle('${library}')/items?$filter=TaxCatchAll/Term eq '${team}'&$top=10`;

    this.context.spHttpClient.get(requestUrl, SPHttpClient.configurations.v1)
      .then((response : SPHttpClientResponse) => {
        //if(response.ok){
          response.json().then((responseJSON:any) => {
            console.log("responseJSON",responseJSON.value.length);
            dataFlag = true;              
            //if(responseJSON.value.length>0){
            //  dataFlag=true; 
            //}
          });
          //}
      });
    return dataFlag;   
  }
*/

  // **** Function  : _renderCustomTabsAsync
  // **** Purpose   : set termID based on Divsion
  // **** Return    :
  // **** Author    : Jason Clark
  // **** Date      : Jan 2024
  private async renderCustomTabsAsync(): Promise<void> {
    const setID = "be84d0a6-e641-4f6d-830e-11e81f13e2f1";  // Maximus Custom Set Terms
    //const setID = "04f05939-559c-49ef-b69b-e903f79ee898";  // Munkieman Custom Set Terms
    let termID: string = "";
    let label: string = "";
  
    console.log("division",this.properties.divisionTitle);

    try {
      switch (this.properties.divisionTitle) {
        case "Assessments":
          termID = "90a0a9eb-bbcc-4693-9674-e56c4d41375f";  // Maximus Assessment Custom Terms
          //termID = "37618895-06ce-47c5-ba62-93176ee62020";  // Munkieman Assessment Custom Terms
          break;
        case "Central":
          termID = "471a563b-a4d9-4ce7-a8e6-4124562b3ace";  // Maximus Central Custom Terms
          //termID = "39713c5b-3b41-4b56-bbec-d42479a9059b";  // Munkieman Central Custom Terms
          break;
        case "Connect":
          termID = "3532f8fc-4ad2-415c-94ff-c5c7af559996";  // Maximus Connect Custom Terms
          break;
        case "Employability":
          termID = "feb3d3c8-d948-4d3e-b997-a2ea74653b3e";  // Maximus Employability Custom Terms
          break;
        case "Health":
          termID = "c9dfa3b6-c7c6-4e74-a738-0ffe54e1ff5c";  // Maximus Health Custom Terms
          break;
      }
  
      this.getCustomTabs(setID, termID)
      .then(async (terms) => {
        for (let x = 0; x < terms.value.length; x++) {
          label = terms.value[x].labels[0].name;
          termID = terms.value[x].id;
  
          if (label === this.properties.siteName) {
            await this.getCustomTabs(setID, termID).then(async (response) => {
              await this._renderCustomTabs(response.value);
              await this.setCustomLibraryListeners();
              await this.setLibraryListeners();
            });
          }
        }
      })
      .catch((err) => {console.log('renderCustomTabs error',err)});

    } catch (err) {
      //await this.addError(this.properties.siteName,"_renderCustomTabsAsync",err.message);
    }
    return;
  }
  
  // **** Function  : getCustomTabs
  // **** Purpose   : check termstore for the custom library tabs, for a given setID and termID.
  // **** Return    : response data as JSON from spHttpClient request
  // **** Author    : Jason Clark
  // **** Date      : Jan 2024
  private async getCustomTabs(setID: string, termID: string): Promise<any> {
    console.log("getCustomTabs",setID,termID);
    try {
      const groupID = "4660ef58-779c-4970-bcd7-51773916e8dd";  // Maximus Document Centre Terms
      // const groupID = "";  // Munkieman Document Centre Terms
      const url: string = `https://${this.properties.tenantURL[2]}/_api/v2.1/termStore/groups/${groupID}/sets/${setID}/terms/${termID}/children?select=id,labels`;
      console.log(url);

      return this.context.spHttpClient
        .get(url, SPHttpClient.configurations.v1)
        .then((response: SPHttpClientResponse) => {
          if (response.ok) {
            return response.json();
          }
        });
    } catch (err) {
      //await this.addError(this.properties.siteName,"getCustomTabs",err.message);
    }
  }
  
  // **** Function  : _renderCustomTabs
  // **** Purpose   : write the custom library tabs to the DOM
  // **** Return    : 
  // **** Author    : Jason Clark
  // **** Date      : Jan 2024
  private async _renderCustomTabs(items: any): Promise<void> {
    let html: string = "";
    let labelName: string = "";
    let labelID: string = "";
    //const libraryTabs : Element | null = this.domElement.querySelector("#libraryTabs");
    //const headerBar : Element | null = this.domElement.querySelector("#headerBar");

    try {
      //if(headerBar){
        this.domElement.querySelector("#headerBar")!.innerHTML = `<div class="col-3">
                      <img src="${require("./assets/Robot_Spin.gif")}" height="50" width="50"/>
                    </div>
                    <div class="col">
                      <div class="alert alert-info" role="alert">Checking for custom folders in Custom Library</div>
                    </div>`;
      //}

      for (let x = 0; x < items.length; x++) {
        labelName = items[x].labels[0].name;
        labelID = labelName.replace(/\s+/g, "");
  
        let libraryFlag = await this.checkDataAsync("Custom",this.properties.siteName,labelName);
        if(libraryFlag){
          if (this.properties.isPowerUser) {
            html += `<button class="btn libraryBtn nav-link mb-1" id="customTab" data-bs-toggle="pill" data-bs-target="#custom" type="button" role="tab" aria-controls="${labelID}" aria-selected="true">
                        <div class="col-1 libraryUploadIcon">
                        <a href="${this.properties.dcURL}/custom/Forms/AllItems.aspx?userFiltersInViewXml=1&FilterField1=DC%5FCategory&FilterValue1=${labelName}&FilterType1=TaxonomyFieldType&FilterField2=DC%5FTeam&FilterValue2=${this.properties.siteName}&FilterType2=TaxonomyFieldType&viewid=0d16333b%2Dc80f%2D438c%2Da14e%2Dff4bea492de9" target="_blank">
                          <h3 class="text-white"><i class="bi bi-cloud-arrow-up"></i></h3>
                        </a>
                        </div>
                        <div class="col-10 libraryName"><h6 id="customTabName" class="libraryText">${labelName}</h6></div>
                      </button>`;
          } else {
            html += `<button class="btn libraryBtn nav-link text-left mb-1" id="customTab" data-bs-toggle="pill" data-bs-target="#custom" type="button" role="tab" aria-controls="Custom" aria-selected="true"><h6 id="customTabName" class="libraryText">${labelName}</h6></button>`;
          }
        }
      }  
      //if(libraryTabs){
        this.domElement.querySelector("#libraryTabs")!.innerHTML += html;
      //}
      setTimeout(() => {
        //if(headerBar){
          this.domElement.querySelector("#headerBar")!.innerHTML = "";
        //}
      }, 3000);
    } catch (err) {
      //await this.addError(this.properties.siteName,"_renderCustomTabs",err.message);
    }
    return;
  }

  private async setCustomLibraryListeners(): Promise<void> {
    // *** event listeners for custom library category tabs
    try {
      const tabBtns = this.domElement.querySelectorAll("#customTabName");
      let timer:any;
      console.log("tabBtns",tabBtns);
      for (let x = 0; x < tabBtns.length; x++) {
        const customTabName = tabBtns[x].innerHTML;

        tabBtns[x]?.addEventListener("click", event => { 
          //if(event.detail===1){
            timer=setTimeout(async () => {
              await this.getData("Custom",7,customTabName);
            },400);
          //}
        });
  
        tabBtns[x]?.addEventListener("dblclick",event => {
          clearTimeout(timer);
        });        
        
        //tabBtns[x].addEventListener("click", async event => {
        //  await this.getData("Custom",7,customTabName);
        //});
      }      
    } catch (err) {
      //await this.addError(this.properties.siteName,"setCustomLibraryListeners",err.message);
    }     
    return;
  } 

  private async setLibraryListeners(): Promise<void> {
    try {
      let timer:any;

      // *** event listeners for main document libraries
      document
        .getElementById("Policies_btn")
        ?.addEventListener("click", async (_e: Event) => { 
          await this.getData("Policies",1,"");
      });
      document
        .getElementById("Procedures_btn")
        ?.addEventListener("click", async (_e: Event) => {
          await this.getData("Procedures",2,"");
      });

      let guideElem=document.getElementById("Guides_btn");

      guideElem?.addEventListener("click", async event => {
        if(event.detail===1){
          timer=setTimeout(() => {
            this.getData("Guides",3,"");
          },400);
        }
      });
      guideElem?.addEventListener("dblclick",event => {
        clearTimeout(timer);
      });

      document
        .getElementById("Forms_btn")
        ?.addEventListener("click", async (_e: Event) => {
          await this.getData("Forms",4,"");
        });
      document
        .getElementById("General_btn")
        ?.addEventListener("click", async (_e: Event) => {
          await this.getData("General",5,"");
        });
  
      // *** event listener for management library
      //if (this.properties.isManager) {
      //  document
      //    .getElementById("Management_btn")
      //    ?.addEventListener("click", async (_e: Event) => {
      //      await this.getData("Management", 6, "");
      //  });
      //}

      console.log("dataResults",this.properties.dataResults.length);

    } catch (err) {
      //await this.addError(this.properties.siteName, "setLibraryListeners", err.message);
    }  
    return;
  }  

  private async getData(libraryName:string,tabNum:number,category:string): Promise<void> {
    //alert(libraryName);
    
    const sp = spfi().using(SPFx(this.context)).using(PnPLogging(LogLevel.Warning));  
    this.properties.dataResults=[];
    //this.properties.folderArray=[];
    this.properties.dcDivisions=["asm","cen"];
    this.properties.divisionName=["Assessments","Central","Connect","Employability","Health"];

    const view : string = `<View>
                            <Query>
                              <Where>
                                <Or>
                                  <Eq>                
                                    <FieldRef Name="DC_Team"/>
                                    <Value Type="TaxonomyFieldType">${this.properties.siteName}</Value>
                                  </Eq>                                                                
                                  <Contains>                
                                    <FieldRef Name="DC_SharedWith"/>
                                    <Value Type='TaxonomyFieldTypeMulti'>${this.properties.siteName}</Value>
                                  </Contains>
                                </Or>                                  
                              </Where>
                              <OrderBy>
                                <FieldRef Name="DC_Folder" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder01" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder02" Ascending="TRUE" />
                                <FieldRef Name="DC_SubFolder03" Ascending="TRUE" />
                                <FieldRef Name="LinkFilename" Ascending="TRUE" />
                              </OrderBy>
                            </Query>                            
                          </View>`;
    console.log(view);

    this.properties.dcDivisions.forEach(async (site,index)=>{
      console.log(site,index);
      const dcTitle = site+"_dc";
      const tenant_uri = this.context.pageContext.web.absoluteUrl.split('/',3)[2];
      const webDC = Web([sp.web,`https://${tenant_uri}/sites/${dcTitle}/`]); 
      //const division = this.properties.divisionName[index];

      await webDC.lists.getByTitle(libraryName)
        .getItemsByCAMLQuery({ViewXml:view},"FieldValuesAsText/FileRef", "FieldValueAsText/FileLeafRef")
        .then(async (Results) => {
          if(Results.length>0){
            console.log(dcTitle+" Results");
            console.log(Results);
            await this.addToResults(Results).then(async ()=>{            
              await this._renderFolders(libraryName).then(async () => {
                await this.setFolderListeners(libraryName);
              });
             
              // const folderSet = new Set(Results.DC_Folder); 
              //let res = _.uniqWith(Results,_.isEqual);
              //const set = new Set();
              //const folders = Results.reduce((a: string | any[],c: any)=>{
              //  const str = JSON.stringify(c);
              //  if(!set.has(str)){
              //    set.add(str);
              //    return a.concat(c);
              //  }                
                //count[folderName]=(count[folderName] || 0)+1;
              //  return a;
              //}, [] )
              //console.log("folderCount",folders);
              //console.log("res",res);
              //console.log("folderset",folderSet);

            });        
          }else{
            //alert("No Data found for this Team in "+dcTitle);
          }    
        });
        //.catch(() => {console.log("error")});
    });
  }

  private async addToResults(results:any):Promise<void>{
    let count:number=0; 
    
    count=this.properties.dataResults.length;
    for(let x=0;x<results.length;x++){
      //const subfolderName = results[x].FieldValuesAsText.DC_x005f_SubFolder01;
      this.properties.dataResults[count+x]=results[x];
      //this.properties.subFolder1Array.push(subfolderName);
    }
    //const folderSet = new Set(this.properties.subFolder1Array);
    //console.log("folderSet",folderSet);
    //console.log("subfolderArray",this.properties.subFolder1Array);

    return;  //folderSet;
  }

  
  private async _renderFolders(libraryName:string): Promise<void>{

    console.log("dataResults length ",this.properties.dataResults.length);
    console.log("folder dataResults",this.properties.dataResults);

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFolders");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFolders");
    const guidesContainer : Element | null = this.domElement.querySelector("#guidesFolders");
    const formsContainer : Element | null = this.domElement.querySelector("#formsFolders");

    let folderHTML: string = "";
    //let subfolder1HTML : string = "";
    //let subfolder2HTML : string = "";
    //let subfolder3HTML : string = "";    
    //let folderHTMLEnd : string = "";

    //let subFolder1Flag : boolean;
    //let subFolder2Flag : boolean;
    //let subfolder3Flag : boolean;

    let folderName: string = "";
    let subFolderName1 : string = "";
    let subFolderName2 : string = "";
    let subFolderName3 : string = "";
    
    let folderPrev: string = "";
    let subFolderPrev1 : string = "";
    let subFolderPrev2 : string = "";
    let subFolderPrev3 : string = "";

    // *** arrays of folder id's for the Folder EventListeners
    this.properties.folderArray = [];
    this.properties.subFolder1Array = []; 
    this.properties.subFolder2Array = [];
    this.properties.subFolder3Array = [];

    // *** folder id's for event listeners on button click
    let folderNameID : string = ""; 
    let subFolderName1ID: string = "";
    let subFolderName2ID: string = "";
    let subFolderName3ID: string = "";

    //let folderID : string = ""; // accordion id's for folders
    //let subFolder01ID : string = "";
    //let subFolder02ID : string = "";

    let fcount:any=0;
    let sf1count:any=0;
    let sf2count:any=0;
    let sf3count:any=0;

    if(this.properties.dataResults.length > 0){

      switch (libraryName) {
        case "Policies":        
          if(policyContainer){policyContainer.innerHTML="";}
          break;
        case "Procedures":
          if(procedureContainer){procedureContainer.innerHTML="";}
          break;
        case "Guides":
          if(guidesContainer){guidesContainer.innerHTML="";}
          break;
        case "Forms":
          if(formsContainer){formsContainer.innerHTML="";}
          break;
        case "General":
          //this.domElement.querySelector("#generalFolders")!.innerHTML="";
          break;
        case "Management":
          //this.domElement.querySelector("#managementFolders")!.innerHTML="";
          break;
        case "Custom":
          //this.domElement.querySelector("#customFolders")!.innerHTML="";
          break;
      }

      //if(this.domElement.querySelector("#policiesFolders")!=null) {
      //  this.domElement.querySelector("#policiesFolders")!.innerHTML = "<h2>folder name</h2>";
      //}
      
      folderHTML=`<div class="d-grid gap-2 mx-auto">`;

      for(let x=0;x<this.properties.dataResults.length;x++){
        
        const fileStatus=this.properties.dataResults[x].FieldValuesAsText.OData__x005f_ModerationStatus
        //folderHTML="";
        //subFolder1Flag=false;
        //subFolder2Flag=false;

        if (fileStatus === "Approved" || (fileStatus === "Draft" && this.properties.isPowerUser)) {

          folderName = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_Folder;            
          subFolderName1 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder01;
          subFolderName2 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder02;
          subFolderName3 = this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder03;

          if(folderName.replace(/\s+/g, "")!==undefined){
            folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
          }else{
            folderNameID=folderName+"_"+x;
          }

          if(subFolderName1.replace(/\s+/g, "")!==undefined){
            subFolderName1ID=subFolderName1.replace(/\s+/g, "")+"_"+x;
          }else{
            subFolderName1ID=subFolderName1+"_"+x;
          }

          if(folderName !== null){

            if(folderName !== folderPrev){

              this.properties.folderArray.push(folderName,folderNameID);
              fcount = await this.fileCount(folderName,"","","");
              // *** Parent Folder ID
              //folderID = `accordion-Folder`;

              if(subFolderName1!==""){
                //folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" class="accordion-button mb-1 btn btn-primary text-left"><h6>${folderName}</h6></button>`;
                folderHTML+=`<button class="row btn btn-primary mb-1 ${styles.folderBtn}" type="button" data-bs-toggle="collapse" aria-pressed="true" data-bs-target="#accordion_${folderNameID}" aria-expanded="true">
                              <h5 class="accordion-header" id="folder_${folderNameID}">
                                <i class="col-1 text-left bi bi-folder2"></i>
                                <span class="col-2 text-left badge ms-bgColor-themePrimary"><i class="bi bi-file-earmark"></i> ${fcount}</span>                    
                                <a href="#" class="col text-white ms-2">${folderName}</a>
                              </h5>
                            </button>
                            <div class="collapse" id="accordion_${folderNameID}">
                                <div class="card card-body">
                                </div>
                            </div>`;                              
              }else{
                folderHTML+=`<button class="row btn btn-success mb-1 ${styles.folderBtn}" type="button" data-bs-toggle="button" aria-pressed="true">
                              <h5 class="accordion-header" id="folder_${folderNameID}">
                                <i class="col-1 text-left bi bi-folder2"></i>
                                <span class="col-2 text-left badge ms-bgColor-themePrimary"><i class="bi bi-file-earmark"></i> ${fcount}</span>                
                                <a href="#" class="col text-white ms-2">${folderName}</a>
                              </h5> 
                            </button>`;
              }
             
              console.log("folderName",folderName);
              console.log("subFolderName1",subFolderName1);
              folderPrev=folderName;
            }

            if(subFolderName1 !== ""){

              this.properties.subFolder1Array.push(subFolderName1,subFolderName1ID);
              sf1count = await this.fileCount(folderName,subFolderName1,"","");

              if(subFolderName1 !== subFolderPrev1){
                
                if(subFolderName2 !== ""){
                  folderHTML+=`<button class="btn btn-secondary ms-2 ${styles.folderBtn}" type="button" data-bs-toggle="collapse" aria-pressed="true" data-bs-target="#accordion_${subFolderName1ID}" aria-expanded="true">
                                <h5 class="accordion-header" id="SF1_${subFolderName1ID}">
                                <i class="bi bi-folder2"></i>
                                <span class="badge ms-bgColor-themePrimary">${sf1count}</span>                    
                                <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                </h5>
                                </button>
                                <div class="collapse" id="accordion_${subFolderName1ID}">
                                  <div class="card card-body">
                                  </div>
                                </div>`;         
                }else{
                  folderHTML+=`<button class="btn btn-dark ms-2 ${styles.folderBtn}" type="button" data-bs-toggle="button" aria-pressed="true">
                                <h5 class="accordion-header" id="SF1_${subFolderName1ID}">
                                <i class="bi bi-folder2"></i>
                                <span class="badge ms-bgColor-themePrimary">${sf1count}</span>                    
                                <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                </h5>
                              </button>`;
                              
                }  

                console.log("subFolderName1",subFolderName1);
                subFolderPrev1 = subFolderName1;
              }

              if(subFolderName2 !== ""){
                
                if(subFolderName2.replace(/\s+/g, "")!==undefined){
                  subFolderName2ID=subFolderName2.replace(/\s+/g, "")+"_"+x;
                }else{
                  subFolderName2ID=subFolderName2+"_"+x;
                }
                this.properties.subFolder2Array.push(subFolderName2,subFolderName2ID);
                sf2count = await this.fileCount(folderName,subFolderName1,subFolderName2,"");

                //subFolderName2 = this.properties.dataResults[x].DC_SubFolder02.Label;
                
                if(subFolderName2 !== subFolderPrev2){

                  if(subFolderName3 !== ""){
                    folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" class="ms-3 btn btn-info ${styles.folderBtn}">
                                  <h5 class="accordion-header" id="SF2_${subFolderName2ID}">
                                  <i class="bi bi-folder2"></i>
                                  <span class="badge ms-bgColor-themePrimary">${sf2count}</span>                     
                                  <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                  </h5>
                                  </button>`;
                  }else{
                    folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" class="ms-3 btn btn-warning ${styles.folderBtn}">
                                  <h5 class="accordion-header" id="SF2_${subFolderName2ID}">
                                  <i class="bi bi-folder2"></i>
                                  <span class="badge ms-bgColor-themePrimary">${sf2count}</span>                     
                                  <a href="#" class="text-white ms-1">${subFolderName2}</a>        
                                  </h5>
                                  </button>`;
                                  
                  }
                  
                  console.log("subFolderName2",subFolderName2);
                  subFolderPrev2 = subFolderName2;
                }

                if(subFolderName3 !== ""){

                  if(subFolderName3.replace(/\s+/g, "")!==undefined){
                    subFolderName3ID=subFolderName3.replace(/\s+/g, "")+"_"+x;
                  }else{
                    subFolderName3ID=subFolderName3+"_"+x;
                  }
                  this.properties.subFolder3Array.push(subFolderName3,subFolderName3ID);
                  sf3count = await this.fileCount(folderName,subFolderName1,subFolderName2,subFolderName3);

                  //subFolderName3 = this.properties.dataResults[x].DC_SubFolder03.Label;
                  
                  if(subFolderName3 !== subFolderPrev3){                    
                    folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" data-bs-toggle="collapse" class="ms-4 btn btn-danger ${styles.folderBtn}">
                                  <h5 class="accordion-header" id="SF3_${subFolderName3ID}">
                                  <i class="bi bi-folder2"></i>
                                  <span class="badge ms-bgColor-themePrimary">${sf3count}</span>                    
                                  <a href="#" class="text-white ms-1">${subFolderName3}</a>        
                                  </h5>
                                  </button>`;
                    console.log("subFolderName3",subFolderName3);
                    subFolderPrev3 = subFolderName3;
                  }
                }                 
              } 
            }            
          }
        }
      }  // *** end of for loop

      folderHTML+=`</div>`;

      console.log(folderHTML);

      switch (libraryName) {
        case "Policies":
          if(policyContainer){policyContainer.innerHTML=folderHTML;}
          break;
        case "Procedures":
          if(procedureContainer){procedureContainer.innerHTML=folderHTML;}
          break;
        case "Guides":
          if(guidesContainer){guidesContainer.innerHTML=folderHTML;}
          break;
        case "Forms":
          if(formsContainer){formsContainer.innerHTML=folderHTML;}
          //this.domElement.querySelector("#formsFolders")!.innerHTML=folderHTML;
          break;
        case "General":
          //this.domElement.querySelector("#generalFolders")!.innerHTML=folderHTML;
          break;
        case "Management":
          //this.domElement.querySelector("#managementFolders")!.innerHTML=folderHTML;
          break;
        case "Custom":
          //this.domElement.querySelector("#customFolders")!.innerHTML=folderHTML;
          break;
      }
    }
  }

  private async fileCount(folder:string,subfolder1:string,subfolder2:string,subfolder3:string): Promise<number>{

    let counter : number = 0;
    let fCounter : number = 0;
    let sf1Counter : number = 0;
    let sf2Counter : number = 0;
    let sf3Counter : number = 0;
    try{
      for (let c=0;c<this.properties.dataResults.length;c++) {
        const folderName = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_Folder;            
        const subFolderName1 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder01;
        const subFolderName2 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder02;
        const subFolderName3 = this.properties.dataResults[c].FieldValuesAsText.DC_x005f_SubFolder03;
    
        //console.log("folder",folder,"folderName",folderName,"SubFolder1",subFolderName1);

        if ( folderName === folder && subFolderName1 === '' ){ fCounter++; }
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === '' && subFolderName3 === ''){ sf1Counter++; } 
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === subfolder2 && subFolderName3 === ''){ sf2Counter++; } 
        if ( folderName === folder && subFolderName1 === subfolder1 && subFolderName2 === subfolder2 && subFolderName3 === subfolder3){ sf3Counter++; } 
      } 
      
      if(fCounter>0){counter = fCounter;}
      if(sf1Counter>0){counter = sf1Counter;}
      if(sf2Counter>0){counter = sf2Counter;}
      if(sf3Counter>0){counter = sf3Counter;}
    } catch (err) {
      //await this.addError(this.properties.siteName, "fileCount", err.message);
    }
    return counter;
  } 

  private async setFolderListeners(libraryName:string): Promise<void> {

    try {
      // *** event listeners for parent folders      
      if (this.properties.folderArray.length > 0) {

        const folderNameID = this.properties.folderArray.filter((value, index) => index % 2 > 0);
        //const folderName = this.properties.folderArray.filter((value, index) => index % 2 === 0);
        //const division = this.properties.folderArray.filter((value, index) => index % 3 === 3 - 3);

        //console.log("folderNameID ",folderNameID);
        //console.log("folderName ",folderName);
        //console.log("division ",division);

        for (let x = 0; x < folderNameID.length; x++) {

          document.getElementById("folder_" + folderNameID[x])
            ?.addEventListener("click", async (_e: Event) => {
              //await this.getFiles(libraryName,folderName[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#folder_" + folderNameID[x]);
              elem?.classList.add('active');
            });
        }
      }

      // *** event listeners for subfolders level 1      
      if (this.properties.subFolder1Array.length > 0) {

        const subFolder1NameID = this.properties.subFolder1Array.filter((value, index) => index % 2 > 0);
        //const subFolder1Name = this.properties.subFolder1Array.filter((value, index) => index % 2 === 0);

        for (let x = 0; x < subFolder1NameID.length; x++) {
          document.getElementById("SF1_" + subFolder1NameID[x])
            ?.addEventListener("click", async (_e: Event) => {
              //console.log("subfolder1name ",subFolder1Name[x]);
              //await this.getFiles(libraryName,subFolder1Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF1_" + subFolder1NameID[x]);
              elem?.classList.add('active');
            });
        }
      }

      // *** event listeners for subfolders level 2      
      if (this.properties.subFolder2Array.length > 0) {

        const subFolder2NameID = this.properties.subFolder2Array.filter((value, index) => index % 2 > 0);
        //const subFolder2Name = this.properties.subFolder2Array.filter((value, index) => index % 2 === 0);

        for (let x = 0; x < subFolder2NameID.length; x++) {          
          document.getElementById("SF2_" + subFolder2NameID[x])
            ?.addEventListener("click", async (_e: Event) => {
              //await this.getFiles(libraryName,subFolder2Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF2_" + subFolder2NameID[x]);
              elem?.classList.add('active');
            });
        }
      }      

      // *** event listeners for subfolders level 3      
      if (this.properties.subFolder3Array.length > 0) {

        const subFolder3NameID = this.properties.subFolder3Array.filter((value, index) => index % 2 > 0);
        //const subFolder3Name = this.properties.subFolder3Array.filter((value, index) => index % 2 === 0);

        for (let x = 0; x < subFolder3NameID.length; x++) {          
          document.getElementById("SF3_" + subFolder3NameID[x])
            ?.addEventListener("click", async (_e: Event) => {
              //await this.getFiles(libraryName,subFolder3Name[x]);

              if (document.querySelector("#" + libraryName + "Folders button.active") !== null) {
                document
                  .querySelector("#" + libraryName + "Folders button.active")
                  ?.classList.remove("active");
              }
        
              const elem = document.querySelector("#SF3_" + subFolder3NameID[x]);
              elem?.classList.add('active');
            });
        }
      } 

    } catch (err) {
      //await this.addError(this.properties.siteName, "setFolderListeners", err.message);
    }    
  }   

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");
            
    return this._getEnvironmentMessage().then(message => {
      //this._environmentMessage = message;
    });
  }

  private _getEnvironmentMessage(): Promise<string> {
    if (!!this.context.sdks.microsoftTeams) { // running in Teams, office.com or Outlook
      return this.context.sdks.microsoftTeams.teamsJs.app.getContext()
        .then(context => {
          let environmentMessage: string = '';
          switch (context.app.host.name) {
            case 'Office': // running in Office
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOffice : strings.AppOfficeEnvironment;
              break;
            case 'Outlook': // running in Outlook
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentOutlook : strings.AppOutlookEnvironment;
              break;
            case 'Teams': // running in Teams
            case 'TeamsModern':
              environmentMessage = this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentTeams : strings.AppTeamsTabEnvironment;
              break;
            default:
              environmentMessage = strings.UnknownEnvironment;
          }

          return environmentMessage;
        });
    }
    return Promise.resolve(this.context.isServedFromLocalhost ? strings.AppLocalEnvironmentSharePoint : strings.AppSharePointEnvironment);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    if (!currentTheme) {
      return;
    }

    //this._isDarkTheme = !!currentTheme.isInverted;
    const {
      semanticColors
    } = currentTheme;

    if (semanticColors) {
      this.domElement.style.setProperty('--bodyText', semanticColors.bodyText || null);
      this.domElement.style.setProperty('--link', semanticColors.link || null);
      this.domElement.style.setProperty('--linkHovered', semanticColors.linkHovered || null);
    }

  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
