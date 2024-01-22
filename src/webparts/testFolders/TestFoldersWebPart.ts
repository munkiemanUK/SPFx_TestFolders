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

import { spfi, SPFx } from "@pnp/sp";
import { Web } from "@pnp/sp/webs"; 
import "@pnp/sp/webs";
import "@pnp/sp/lists";
import "@pnp/sp/items";
import "@pnp/sp/items/get-all";
import { LogLevel, PnPLogging } from "@pnp/logging";
//import * as _ from 'lodash';

require("bootstrap");

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
}

export default class TestFoldersWebPart extends BaseClientSideWebPart<ITestFoldersWebPartProps> {

  private _isDarkTheme: boolean = false;
  private _environmentMessage: string = '';

  public render(): void {

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
      <div class="${styles.welcome}">
        <img alt="" src="${this._isDarkTheme ? require('./assets/welcome-dark.png') : require('./assets/welcome-light.png')}" class="${styles.welcomeImage}" />
        <h2>Well done, ${escape(this.context.pageContext.user.displayName)}!</h2>
        <div>${this._environmentMessage}</div>
        <div>Web part property value: <strong>${escape(this.properties.description)}</strong></div>
        <br/>
        <div>siteTitle = ${this.properties.siteTitle}</div>
        <div>siteName = ${this.properties.siteName}</div>
        <div>siteShortName = ${this.properties.siteShortName}</div>
        <div>siteID = ${this.properties.siteID}</div>
        <div>divisionTitle = ${this.properties.divisionTitle}</div>
        <div>tenant URL = ${uri.split('/',3)[2]}</div>
      </div>

      <div class="container v-scrollbar">
        <div class="row" style="height:50px" id="headerBar">
          ${[/*<div class="col-3"><img src="${require("./assets/Robot_Spin.gif")}" height="50" width="50"/></div>
            <div class="col" id="headerBar"></div>*/]}
          ${[/*<div class="col">${this.properties.siteTitle}</div>*/]}
        </div>
        <div class="row">
          <div class="d-flex align-items-start mt-1">
            <div class="nav flex-column nav-pills me-3" id="libraryTabs" role="tablist" aria-orientation="vertical">
              <button class="nav-link" id="policies_btn" data-bs-toggle="pill" data-bs-target="#policies" type="button" role="tab" aria-controls="v-pills-home" aria-selected="false">Policies</button>
              <button class="nav-link" id="procedures_btn" data-bs-toggle="pill" data-bs-target="#procedures" type="button" role="tab" aria-controls="v-pills-profile" aria-selected="false">Procedures</button>
              <button class="nav-link" id="guides_btn" data-bs-toggle="pill" data-bs-target="#guides" type="button" role="tab" aria-controls="v-pills-disabled" aria-selected="false">Guides</button>
              <button class="nav-link" id="forms_btn" data-bs-toggle="pill" data-bs-target="#forms" type="button" role="tab" aria-controls="v-pills-messages" aria-selected="false">Forms</button>
              <button class="nav-link" id="general_btn" data-bs-toggle="pill" data-bs-target="#general" type="button" role="tab" aria-controls="v-pills-settings" aria-selected="false">General</button>
            </div>

            <div class="tab-content" id="v-pills-tabContent">
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
                  <div class="col-auto" id="guidesFolders"></div>
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
      </div>
    </section>`;

    document.getElementById('policies_btn')?.addEventListener("click", (_e:Event) => this._getData('Policies',1,""));
    document.getElementById('procedures_btn')?.addEventListener("click",(_e:Event) => this._getData('Procedures',2,""));
    document.getElementById('guides_btn')?.addEventListener("click",(_e:Event) => this._getData('Guides',3,""));
    document.getElementById('forms_btn')?.addEventListener("click",(_e:Event) => this._getData('Forms',4,""));
    document.getElementById('general_btn')?.addEventListener("click",(_e:Event) => this._getData('General',5,""));
  }

  private  _getData(libraryName:string,tabNum:number,category:string): void {
    alert(libraryName);
    
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
              await this._renderFolders(tabNum,libraryName).then(async () => {
              //    this.setFolderListeners(division,libraryName);
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
            alert("No Data found for this Team in "+dcTitle);
          }    
        });
        //.catch(() => {console.log("error")});
    });
  }

  private async addToResults(results:any):Promise<void>{
    let count:number=0; 
    
    count=this.properties.dataResults.length;
    for(let x=0;x<results.length;x++){
      const subfolderName = results[x].FieldValuesAsText.DC_x005f_SubFolder01;
      this.properties.dataResults[count+x]=results[x];
      this.properties.subFolder1Array.push(subfolderName);
    }
    const folderSet = new Set(this.properties.subFolder1Array);
    console.log("folderSet",folderSet);
    console.log("subfolderArray",this.properties.subFolder1Array);

    return;  //folderSet;
  }

  
  private async _renderFolders(tabNum:number,libraryName:string): Promise<void>{

    console.log("dataResults length ",this.properties.dataResults.length);
    console.log("folder dataResults",this.properties.dataResults);

    const policyContainer : Element | null = this.domElement.querySelector("#policiesFolders");
    const procedureContainer : Element | null = this.domElement.querySelector("#proceduresFolders");
    const guidesContainer : Element | null = this.domElement.querySelector("#guidesFolders");
    const formsContainer : Element | null = this.domElement.querySelector("#formsFolders");

    let folderHTML: string = "";

    let folderName: string = "";
    let subFolderName1 : string = "";
    //let subFolderName2 : string = "";
    //let subFolderName3 : string = "";
    
    let folderPrev: string = "";
    let subFolderPrev1 : string = "";

    // *** arrays of folder id's for the Folder EventListeners
    this.properties.folderArray = [];
    this.properties.subFolder1Array = []; 
    this.properties.subFolder2Array = [];
    this.properties.subFolder3Array = [];
        
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
        console.log(fileStatus);
        
        if (fileStatus === "Approved" || (fileStatus === "Draft" && this.properties.isPowerUser)) {

          folderName = this.properties.dataResults[x].DC_Folder.Label;            

          if(folderName !== undefined ){

            if(folderName !== folderPrev){
              //this.properties.folderArray.push(folderName);
              folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" class="mb-1 btn btn-primary"><h6>${folderName}</h6></button>`;
              folderPrev=folderName;
            }

            if(this.properties.dataResults[x].FieldValuesAsText.DC_x005f_SubFolder01 !== ""){
              subFolderName1 = this.properties.dataResults[x].DC_SubFolder01.Label;
              
              if(subFolderName1 !== subFolderPrev1){
                folderHTML+=`<button type="button" data-bs-toggle="button" aria-pressed="true" class="ms-2 btn btn-secondary"><h6>${subFolderName1}</h6></button>`;
                console.log(subFolderName1);
                subFolderPrev1 = subFolderName1;
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

/*
  private async makeHTML(x:number,division:string,folderName:string,subFolderName1:string,subFolderName2:string,subFolderName3:string):Promise<string>{

    let fcount:any=0;
    let sf1count:any=0;
    let sf2count:any=0;
    let sf3count:any=0;

    let folderHTML: string = "";
    let folderHTMLEnd : string = "";
    
    // *** folder id's for event listeners on button click
    let folderNameID : string = ""; 
    let subFolderName1ID: string = "";
    let subFolderName2ID: string = "";
    let subFolderName3ID: string = "";

    // *** check if folderName has spaces or special characters and remove them for the ID.
    if(folderName.replace(/\s+/g, "")!==undefined){
      folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
    }else{
      folderNameID=folderName+"_"+x;
    }
    this.properties.folderArray.push(folderName,folderNameID);
    fcount = await this.fileCount(division,folderName);

    if(subFolderName1!==``){       
      folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        <button class="btn btn-primary accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF1-${x}" aria-expanded="true" aria-controls="collapseSF1-${x}">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1" id="${folderNameID}">${folderName}</a>
                          <span class="badge bg-secondary">${fcount}</span>                    
                        </button>
                      </h2>`;
    }else{
      folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1-${x}">
                          <i class="bi bi-folder2"></i>
                          <a href="#" class="text-white ms-1" id="${folderNameID}">${folderName}</a>
                          <span class="badge bg-secondary">${fcount}</span>                    
                        </button>
                      </h2>`;
    }
    
    if(subFolderName1 !== ''){
      //console.log("CHK subfolder1 ",subFolderName1);

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName1.replace(/\s+/g, "")!==undefined){
        subFolderName1ID=subFolderName1.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName1ID=subFolderName1+"_"+x;
      }
      this.properties.subFolder1Array.push(subFolderName1,subFolderName1ID);
      sf1count = await this.fileCount(division,subFolderName1);

      if(subFolderName2 !== ``){
        folderHTML+=`<div id="collapseSF1-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF1-${x}" data-bs-parent="#accordionPF-${x}">
                      <div class="accordion-body"> 
                        <div class="accordion" id="accordionSF1-${x}">                              
                          <div class="accordion-item">
                            <h2 class="accordion-header" id="SF1_${subFolderName1ID}">
                              <button class="btn btn-primary accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                <i class="bi bi-folder2"></i>
                                <a href="#" class="text-white ms-1" id="${subFolderName1ID}">${subFolderName1}</a>
                                <span class="badge bg-secondary">${sf1count}</span>                                        
                              </button>
                            </h2>`;
        folderHTMLEnd+=`</div></div></div></div>`;

      }else{
        folderHTML+=`<div id="collapseSF1-${x}" class="ms-1 accordion-collapse collapse" aria-labelledby="headingSF1" data-bs-parent="#accordionPF-${x}">
                      <div class="accordion-body">
                        <div class="accordion-item">
                          <h2 class="accordion-header" id="SF1_${subFolderName1ID}">
                            <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1">
                              <i class="bi bi-folder2"></i>
                              <a href="#" class="text-white ms-1" id="${subFolderName1ID}">${subFolderName1}</a>
                              <span class="badge bg-secondary">${sf1count}</span>                    
                            </button>
                          </h2>
                        </div>
                      </div>
                    </div>`;
      }               
    }

    if(subFolderName2 !== ''){
      //console.log("CHK subfolder2 ",subFolderName2);  

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName2.replace(/\s+/g, "")!==undefined){
        subFolderName2ID=subFolderName2.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName2ID=subFolderName2+"_"+x;
      }
      this.properties.subFolder2Array.push(subFolderName2,subFolderName2ID);
      sf2count = await this.fileCount(division,subFolderName2);

      if(subFolderName3 !==``){
        folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                      <div class="accordion-body">
                        <div class="accordion" id="accordionSF2-${x}">
                          <div class="accordion-item">
                            <h2 class="accordion-header" id="SF2_${subFolderName2ID}">
                              <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                <i class="bi bi-folder2"></i>
                                <a href="#" class="text-white ms-1">${subFolderName2}</a>
                                <span class="badge bg-secondary">${sf2count}</span>                    
                              </button>
                            </h2>`;
        folderHTMLEnd+=`</div></div></div></div>`;

      }else{
        folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                      <div class="accordion-body">
                        <div class="accordion-item">
                          <h2 class="accordion-header" id="SF2_${subFolderName2ID}">
                            <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF2-${x}">
                              <i class="bi bi-folder2"></i>
                              <a href="#" class="text-white ms-1">${subFolderName2}</a>
                              <span class="badge bg-secondary">${sf2count}</span>                    
                            </button>
                          </h2>
                        </div>
                      </div>
                    </div>`;
      }               
    }   

    if(subFolderName3 !== ''){

      // *** check if subfolderName has spaces or special characters and remove them for the ID.
      if(subFolderName3.replace(/\s+/g, "")!==undefined){
        subFolderName3ID=subFolderName3.replace(/\s+/g, "")+"_"+x;
      }else{
        subFolderName3ID=subFolderName3+"_"+x;
      }
      this.properties.subFolder3Array.push(subFolderName3,subFolderName3ID);
      sf3count = await this.fileCount(division,subFolderName3);

      folderHTML+=`<div id="collapseSF3-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="accordionSF2-${x}">
                    <div class="accordion-body">
                      <div class="accordion-item">
                        <h2 class="accordion-header" id="SF3_${subFolderName3ID}">
                          <button class="btn btn-primary" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF3-${x}">
                            <i class="bi bi-folder2"></i>
                            <a href="#" class="text-white ms-1" id="sf3ID">${subFolderName3}}</a>
                            <span class="badge bg-secondary">${sf3count}</span>                    
                          </button>
                        </h2>
                      </div>
                    </div>
                  </div>`;
    }
    folderHTML+=folderHTMLEnd;
    return folderHTML;
  }

  private async fileCount(division:string,folderName:string): Promise<number>{

    let counter : number = 0;
    for (let c=0;c<this.properties.dataResults.length;c++) {
      if(this.properties.dataResults[c].FieldValuesAsText.Knowledge_x005f_Division === division){
        if (this.properties.dataResults[c].FieldValuesAsText.Knowledge_x005f_Folder === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.Knowledge_x005f_SubFolder === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.Knowledge_x005f_SubFolder2 === folderName){
          counter++;
        } 
        if (this.properties.dataResults[c].FieldValuesAsText.Knowledge_x005f_SubFolder3 === folderName){
          counter++;
        }
      } 
    }
    return counter;
  } 

  private setFolderListeners(division:string,libraryName:string): void {
    console.log("setFolderListeners called");

    //this.properties.getFilesCallFlag = false;
    //console.log("setfolderlisteners callflag="+this.properties.getFilesCallFlag);

    try {
      //let division = this.properties.divisionName[num];

      // *** event listeners for parent folders      
      if (this.properties.folderArray.length > 0) {

        const folderNameID = this.properties.folderArray.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const folderName = this.properties.folderArray.filter(function (value, index, ar) {         
          return (index % 2 === 0);
        });

        for (let x = 0; x < folderNameID.length; x++) {

          document.getElementById("folder_" + folderNameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(division,folderName[x]);

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

        const subFolder1NameID = this.properties.subFolder1Array.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const subFolder1Name = this.properties.subFolder1Array.filter(function (value, index, ar) {
          return (index % 2 === 0);
        });

        for (let x = 0; x < subFolder1NameID.length; x++) {
          document.getElementById("SF1_" + subFolder1NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              console.log("subfolder1name ",subFolder1Name[x]);
              this.getFiles(division,subFolder1Name[x]);

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

        const subFolder2NameID = this.properties.subFolder2Array.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const subFolder2Name = this.properties.subFolder2Array.filter(function (value, index, ar) {
          return (index % 2 === 0);
        });

        for (let x = 0; x < subFolder2NameID.length; x++) {
          
          document.getElementById("SF2_" + subFolder2NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(division,subFolder2Name[x]);

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

        const subFolder3NameID = this.properties.subFolder3Array.filter(function (value, index, ar) {
          return (index % 2 > 0);
        });

        const subFolder3Name = this.properties.subFolder3Array.filter(function (value, index, ar) {
          return (index % 2 === 0);
        });

        for (let x = 0; x < subFolder3NameID.length; x++) {
          
          document.getElementById("SF3_" + subFolder3NameID[x])
            ?.addEventListener("click", (_e: Event) => {
              this.getFiles(division,subFolder3Name[x]);

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
      //console.log("setFolderListeners", err.message);
      //this.addError(this.properties.siteName, "setFolderListeners", err.message);
    }    
    console.log("setFolderListeners completed");
  }   

  private getFiles(division:string,folderName:string):void{
    alert("getFiles called");
  }

*/

  public async onInit(): Promise<void> {
    await super.onInit();
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css");
    SPComponentLoader.loadCss("https://cdn.jsdelivr.net/npm/bootstrap-icons@1.10.3/font/bootstrap-icons.css");
            
    return this._getEnvironmentMessage().then(message => {
      this._environmentMessage = message;
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

    this._isDarkTheme = !!currentTheme.isInverted;
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
