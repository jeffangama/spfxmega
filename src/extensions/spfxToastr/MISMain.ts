import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { MISGlobalNavigation } from './MISGlobalNavigation';

export class MISMain {

  public listGUID: any;
  public navigationType: any; //'Menu' or 'Wiki'
  public termStoreGUID: any; // = _termStoreGUID,
  public termStoreName: any; // = _termStoreName,//Service de métadonnées gérées
  public homeBtnTitle: any;
  public homeBtnUrl: any;
  public isCentralResource: any;
  public spcontext: any;

  constructor(_spcontext: any) {
    this.spcontext = _spcontext;
    this.listGUID = null;
    this.navigationType = null;
    this.termStoreGUID = null;
    this.termStoreName = null;
    this.homeBtnTitle = null;
    this.homeBtnUrl = null;
    this.isCentralResource = null;///commm shall go to rouquet

    try {
      //Contain the webapp where mysites is configured
      var privateUrl = "https://private.monacoinformatiqueservice.mc";
      this.getMenuParameters();
      //SP.SOD.executeOrDelayUntilScriptLoaded(MIS.SiteCollection.execute, "sp.ui.pub.ribbon.js");
    } catch (ex) {
      //////console.log("$(document).ready" + ex);
      if (typeof console === "object") {
        ////console.log("$(document).ready" + ex);
      } else {
        alert("$(document).ready" + ex);
      }
    }
  }

  //public static init(spcontext: any, _termStoreName: any, _termStoreGUID: any) {
  private getMenuParameters(): void {
    ////console.log("getMenuParameters called");

    try {
      this.loadSessionStorage;
      if (this.termStoreGUID == undefined && this.termStoreName == undefined) {
        //var url = _spPageContextInfo.webAbsoluteUrl;
        var url = (this.spcontext.pageContext.web.absoluteUrl == "/") ? "" : this.spcontext.pageContext.web.absoluteUrl;
        //debugger
        this.getSettingsData(url).then(() => {
          //debugger
          ////console.log("then getSettingsData");
          ////console.log("isCentralresource value " + this.isCentralResource);
          ////console.log("termStoreName value " + this.termStoreName);

          //Insert the menu in the top navigation
          this.insertMenu();

          //if (this.isCentralResource == true && url != this.homeBtnUrl) {
          //  this.getSettingsData(this.homeBtnUrl); //Call GlobalNavigation settings data in the Central ressource url
          //}
        }); //Call GlobalNavigation settings data
      }
    } catch (e) {
      if (typeof console === "object") {
        ////console.log("getMenuParameters function:" + e);
      }
    }
  }
  /**
   * Load data from List GlobalNavigation URL and update the sessionStorage
   *
   */
  private getSettingsData(url): Promise<void> {

    return new Promise<void>((resolve, reject) => {

      ////console.log("gettingsData function called");

      try {
        var requestUri = url + "/_api/web/lists/getByTitle('GlobalNavigation')/Items?$filter=isActif eq 1";
        ////console.log(requestUri);
        // execute AJAX request synchrone
        //debugger;
        let thisGlobal = this;
        //debugger
        $.ajax({
          url: requestUri,
          type: "GET",
          headers: {
            "ACCEPT": "application/json;odata=verbose"
          },
          success: function (data) {
            ////debugger
            ////console.log(data);
            //debugger;
            //debugger
            if (data.d.results.length > 0) {

              thisGlobal.listGUID = data.d.results[0].ListGUID;
              thisGlobal.navigationType = data.d.results[0].NavigationType;
              thisGlobal.termStoreName = data.d.results[0].TermStoreName;
              ////console.log("1- set termstoreName with value " + data.d.results[0].TermStoreName)
              if (thisGlobal.termStoreName == null) {
                console.error("ERROR, the function getSettingsData did nto return the term store name from global nav list ! Did you deploy global nav list ?")
              }

              thisGlobal.termStoreGUID = data.d.results[0].TermStoreGUID;
              thisGlobal.homeBtnTitle = data.d.results[0].HomeBtnTitle;
              thisGlobal.homeBtnUrl = data.d.results[0].HomeBtnUrl;
              thisGlobal.isCentralResource = data.d.results[0].CentralResource;
              //Backup Data in session.Storage
              sessionStorage.setItem("listGUID", thisGlobal.listGUID);
              sessionStorage.setItem("navigationType", thisGlobal.navigationType);
              sessionStorage.setItem("termStoreName", thisGlobal.termStoreName);
              sessionStorage.setItem("termStoreGUID", thisGlobal.termStoreGUID);
              sessionStorage.setItem("homeBtnTitle", thisGlobal.homeBtnTitle);
              sessionStorage.setItem("homeBtnUrl", thisGlobal.homeBtnUrl);
              sessionStorage.setItem("isCentralResource", thisGlobal.isCentralResource);
              resolve();
            }else {
              throw new Error("Cant read settings from global navigation list");
            }
          },
          error: function (err) {
            //debugger
            if (typeof console === "object") {
              ////console.log("getSettingsData - ajax :" + err);
            }
            resolve();
          }
        });
      } catch (e) {
        if (typeof console === "object") {
          //console.log("getSettingsData function:" + e);
        }
        reject();
      }
    });
  }
  /**
   * Load data from sessionStorage
   *
   */
  private loadSessionStorage(): void {
    if (typeof (Storage) !== "undefined") {
      if (sessionStorage.termStoreGUID) {
        this.termStoreGUID = sessionStorage.getItem("termStoreGUID");
      }
      if (sessionStorage.termStoreName) {
        this.termStoreName = sessionStorage.getItem("termStoreName");
      }
      if (sessionStorage.navigationType) {
        this.navigationType = sessionStorage.getItem("navigationType");
      }
      if (sessionStorage.listGUID) {
        this.listGUID = sessionStorage.getItem("listGUID");
      }
      if (sessionStorage.homeBtnTitle) {
        this.homeBtnTitle = sessionStorage.getItem("homeBtnTitle");
      }
      if (sessionStorage.homeBtnUrl) {
        this.homeBtnUrl = sessionStorage.getItem("homeBtnUrl");
      }
      if (sessionStorage.isCentralResource) {
        this.isCentralResource = sessionStorage.getItem("isCentralResource");
      }
    } else {
      // Sorry! No Web Storage support..
      if (typeof console === "object") {
        ////console.log("Sorry! No Web Storage support..");
      }
    }
  }

  private insertMenu(): void {
    ////console.log("insertMenu called");
    var siteUrl = "";
    if (this.isCentralResource == "true") {
      siteUrl = this.homeBtnUrl;
    } else {
      siteUrl = (this.spcontext.pageContext.web.absoluteUrl == "/") ? "" : this.spcontext.pageContext.web.absoluteUrl;
    }

    //debugger
    switch (this.navigationType) {
      case "Menu":
        //Load the CSS Style sheet file of the project
        //$('head').append("<link rel='stylesheet' type='text/css' href='" + siteUrl + "/Style Library/MIS.GlobalNavigation/css/MIS.GlobalNavigation.css'> </link>");

        //Load the Global Navigation
        //debugger;
        let MISGlobalNavigationObject:MISGlobalNavigation = new MISGlobalNavigation(siteUrl);
        MISGlobalNavigationObject.init(this.spcontext, this.termStoreName, this.termStoreGUID);
        //init(termStoreName, termStoreGUID);//Global Navigation
        break;
      default: //Error in the configuration, nothing is done
    }
  }

  //});
  // }
}
