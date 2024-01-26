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


/*

        htmlString+=`<div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        ${folderHTML}
                      </h2>
                      <div id="collapseSF3" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="#accordionSF3">
                        <div class="accordion-body">
                          <div class="accordion" id="accordionSF3a">
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="headingSF3a">
                                ${subfolder1HTML}
                              </h2>
                              <div id="collapseSF3a" class="accordion-collapse collapse" aria-labelledby="headingSF3a" data-bs-parent="#accordionSF3a">
                                <div class="accordion-body">
                                  <div class="accordion" id="accordionSF3b">
                                    <div class="accordion-item">
                                      <h2 class="accordion-header" id="headingSF3b">
                                        ${subfolder2HTML}
                                      </h2>
                                      <div id="collapseSF3b" class="accordion-collapse collapse" aria-labelledby="headingSF3b" data-bs-parent="#accordionSF3b">
                                        <div class="accordion-body">
                                          ${subfolder3HTML}
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>`;                    

<div class="accordion" id="accordionExample">
  <div class="accordion-item">
    <h2 class="accordion-header" id="headingOne">
      <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseOne" aria-expanded="true" aria-controls="collapseOne">
        Accordion Item #1
      </button>
    </h2>
    <div id="collapseOne" class="accordion-collapse collapse show" aria-labelledby="headingOne" data-bs-parent="#accordionExample">
      <div class="accordion-body">
        <strong>This is the first item's accordion body.</strong> It is shown by default, until the collapse plugin adds the appropriate classes that we use to style each element. These classes control the overall appearance, as well as the showing and hiding via CSS transitions. You can modify any of this with custom CSS or overriding our default variables. It's also worth noting that just about any HTML can go within the <code>.accordion-body</code>, though the transition does limit overflow.
      </div>
    </div>
  </div>
</div>

                  <div class="accordion" id="accordionPF-${x}">
                    <div class="accordion-item">
                      <h2 class="accordion-header" id="folder_${folderNameID}">
                        <button class="accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3" aria-expanded="true" aria-controls="collapseSF3">
                        Folder Name
                        </button>
                      </h2>
                      <div id="collapseSF3" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="#accordionSF3">
                        <div class="accordion-body">
                          <div class="accordion" id="accordionSF3a">
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="headingSF3a">
                                <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3a" aria-expanded="false" aria-controls="collapseSF3a">
                                  Sub Folder 1
                                </button>
                              </h2>
                              <div id="collapseSF3a" class="accordion-collapse collapse" aria-labelledby="headingSF3a" data-bs-parent="#accordionSF3a">
                                <div class="accordion-body">
                                  <div class="accordion" id="accordionSF3b">
                                    <div class="accordion-item">
                                      <h2 class="accordion-header" id="headingSF3b">
                                        <button class="accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3b" aria-expanded="false" aria-controls="collapseSF3b">
                                          Sub Folder 2
                                        </button>
                                      </h2>
                                      <div id="collapseSF3b" class="accordion-collapse collapse" aria-labelledby="headingSF3b" data-bs-parent="#accordionSF3b">
                                        <div class="accordion-body">
                                          <button class="btn btn-info">Sub Folder 3</button>
                                        </div>
                                      </div>
                                    </div>
                                  </div>
                                </div>
                              </div>
                            </div>
                          </div>
                        </div>
                      </div>
                    </div>
                  </div>    
            

  // **** Function  : makeHTML
  // **** Purpose   : create and display the HTML string for each folder and subfolder, then display on DOM.
  // **** Return    :
  // **** Author    : Jason Clark
  // **** Date      : Nov 2023
  private async makeHTML(x:number,folderName:string,subFolderName1:string,subFolderName2:string,subFolderName3:string):Promise<string>{    
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

    try{
      // *** check if folderName has spaces or special characters and remove them for the ID.
      if(folderName.replace(/\s+/g, "")!==undefined){
        folderNameID=folderName.replace(/\s+/g, "")+"_"+x;
      }else{
        folderNameID=folderName+"_"+x;
      }
      this.properties.folderArray.push(folderName,folderNameID);
      fcount = await this.fileCount(folderName,"","","");

      if(subFolderName1!==``){       
        folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                      <div class="accordion-item">
                        <h2 class="accordion-header" id="folder_${folderNameID}">
                          <button class="btn btn-primary ms-1 accordion-button" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF1-${x}" aria-expanded="true" aria-controls="collapseSF1-${x}">
                            <i class="bi bi-folder2"></i>
                            <a href="#" class="text-white ms-2" id="${folderNameID}">${folderName}</a>
                            <span class="badge ms-bgColor-themePrimary">${fcount}</span>                    
                          </button>
                        </h2>`;
      }else{
        folderHTML+=`<div class="accordion" id="accordionPF-${x}">
                      <div class="accordion-item">
                        <h2 class="accordion-header" id="folder_${folderNameID}">
                          <button class="btn btn-primary ms-1" type="button" data-bs-toggle="collapse" aria-expanded="true" aria-controls="collapseSF1-${x}">
                            <i class="bi bi-folder2"></i>
                            <a href="#" class="text-white ms-1" id="${folderNameID}">${folderName}</a>
                            <span class="badge ms-bgColor-themePrimary">${fcount}</span>                    
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
        sf1count = await this.fileCount(folderName,subFolderName1,"","");

        if(subFolderName2 !== ``){
          folderHTML+=`<div id="collapseSF1-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF1-${x}" data-bs-parent="#accordionPF-${x}">
                        <div class="accordion-body"> 
                          <div class="accordion" id="accordionSF1-${x}">                              
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="SF1_${subFolderName1ID}">
                                <button class="btn btn-primary accordion-button ms-1" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF2-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                  <i class="bi bi-folder2"></i>
                                  <a href="#" class="text-white ms-2">${subFolderName1}</a>
                                  <span class="badge ms-bgColor-themePrimary">${sf1count}</span>                                        
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
                                <a href="#" class="text-white ms-1">${subFolderName1}</a>
                                <span class="badge ms-bgColor-themePrimary">${sf1count}</span>                    
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
        sf2count = await this.fileCount(folderName,subFolderName1,subFolderName2,"");

        if(subFolderName3 !==``){
          folderHTML+=`<div id="collapseSF2-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF2" data-bs-parent="accordionSF1-${x}">
                        <div class="accordion-body">
                          <div class="accordion" id="accordionSF2-${x}">
                            <div class="accordion-item">
                              <h2 class="accordion-header" id="SF2_${subFolderName2ID}">
                                <button class="btn btn-primary accordion-button collapsed" type="button" data-bs-toggle="collapse" data-bs-target="#collapseSF3-${x}" aria-expanded="false" aria-controls="collapseSF2-${x}">
                                  <i class="bi bi-folder2"></i>
                                  <a href="#" class="text-white ms-2">${subFolderName2}</a>
                                  <span class="badge ms-bgColor-themePrimary">${sf2count}</span>                    
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
                                <span class="badge ms-bgColor-themePrimary">${sf2count}</span>                    
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
        sf3count = await this.fileCount(folderName,subFolderName1,subFolderName2,subFolderName3);

        folderHTML+=`<div id="collapseSF3-${x}" class="accordion-collapse collapse" aria-labelledby="headingSF3" data-bs-parent="accordionSF2-${x}">
                      <div class="accordion-body">
                        <div class="accordion-item">
                          <h2 class="accordion-header" id="SF3_${subFolderName3ID}">
                            <button class="btn btn-primary ms-1" type="button" data-bs-toggle="collapse" aria-expanded="false" aria-controls="collapseSF3-${x}">
                              <i class="bi bi-folder2"></i>
                              <a href="#" class="text-white ms-1" id="sf3ID">${subFolderName3}</a>
                              <span class="badge ms-bgColor-themePrimary">${sf3count}</span>                    
                            </button>
                          </h2>
                        </div>
                      </div>
                    </div>`;
      }
      folderHTML+=folderHTMLEnd;
    } catch (err) {
      //await this.addError(this.properties.siteName, "makeHTML", err.message);
    }
    return folderHTML;
  }
*/
