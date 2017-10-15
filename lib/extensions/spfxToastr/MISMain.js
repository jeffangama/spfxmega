"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var MISGlobalNavigation_1 = require("./MISGlobalNavigation");
var MISMain = (function () {
    function MISMain(_spcontext) {
        this.spcontext = _spcontext;
        this.listGUID = null;
        this.navigationType = null;
        this.termStoreGUID = null;
        this.termStoreName = null;
        this.homeBtnTitle = null;
        this.homeBtnUrl = null;
        this.isCentralResource = null;
        try {
            //Contain the webapp where mysites is configured
            var privateUrl = "https://private.monacoinformatiqueservice.mc";
            this.getMenuParameters();
            //SP.SOD.executeOrDelayUntilScriptLoaded(MIS.SiteCollection.execute, "sp.ui.pub.ribbon.js");
        }
        catch (ex) {
            //console.log("$(document).ready" + ex);
            if (typeof console === "object") {
                console.log("$(document).ready" + ex);
            }
            else {
                alert("$(document).ready" + ex);
            }
        }
    }
    //public static init(spcontext: any, _termStoreName: any, _termStoreGUID: any) {
    MISMain.prototype.getMenuParameters = function () {
        var _this = this;
        console.log("getMenuParameters called");
        try {
            this.loadSessionStorage;
            if (this.termStoreGUID == undefined && this.termStoreName == undefined) {
                //var url = _spPageContextInfo.webAbsoluteUrl;
                var url = (this.spcontext.pageContext.web.absoluteUrl == "/") ? "" : this.spcontext.pageContext.web.absoluteUrl;
                //debugger
                this.getSettingsData(url).then(function () {
                    //debugger
                    console.log("then getSettingsData");
                    console.log("isCentralresource value " + _this.isCentralResource);
                    console.log("termStoreName value " + _this.termStoreName);
                    //Insert the menu in the top navigation
                    _this.insertMenu();
                    //if (this.isCentralResource == true && url != this.homeBtnUrl) {
                    //  this.getSettingsData(this.homeBtnUrl); //Call GlobalNavigation settings data in the Central ressource url
                    //}
                }); //Call GlobalNavigation settings data
            }
        }
        catch (e) {
            if (typeof console === "object") {
                console.log("getMenuParameters function:" + e);
            }
        }
    };
    /**
     * Load data from List GlobalNavigation URL and update the sessionStorage
     *
     */
    MISMain.prototype.getSettingsData = function (url) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            console.log("gettingsData function called");
            try {
                var requestUri = url + "/_api/web/lists/getByTitle('GlobalNavigation')/Items?$filter=isActif eq 1";
                console.log(requestUri);
                // execute AJAX request synchrone
                var thisGlobal_1 = _this;
                //debugger
                $.ajax({
                    url: requestUri,
                    type: "GET",
                    headers: {
                        "ACCEPT": "application/json;odata=verbose"
                    },
                    success: function (data) {
                        ////debugger
                        console.log(data);
                        //debugger
                        if (data.d.results.length > 0) {
                            thisGlobal_1.listGUID = data.d.results[0].ListGUID;
                            thisGlobal_1.navigationType = data.d.results[0].NavigationType;
                            thisGlobal_1.termStoreName = data.d.results[0].TermStoreName;
                            console.log("1- set termstoreName with value " + data.d.results[0].TermStoreName);
                            if (thisGlobal_1.termStoreName == null) {
                                console.error("ERROR, the function getSettingsData did nto return the term store name from global nav list ! Did you deploy global nav list ?");
                            }
                            thisGlobal_1.termStoreGUID = data.d.results[0].TermStoreGUID;
                            thisGlobal_1.homeBtnTitle = data.d.results[0].HomeBtnTitle;
                            thisGlobal_1.homeBtnUrl = data.d.results[0].HomeBtnUrl;
                            thisGlobal_1.isCentralResource = data.d.results[0].CentralResource;
                            //Backup Data in session.Storage
                            sessionStorage.setItem("listGUID", thisGlobal_1.listGUID);
                            sessionStorage.setItem("navigationType", thisGlobal_1.navigationType);
                            sessionStorage.setItem("termStoreName", thisGlobal_1.termStoreName);
                            sessionStorage.setItem("termStoreGUID", thisGlobal_1.termStoreGUID);
                            sessionStorage.setItem("homeBtnTitle", thisGlobal_1.homeBtnTitle);
                            sessionStorage.setItem("homeBtnUrl", thisGlobal_1.homeBtnUrl);
                            sessionStorage.setItem("isCentralResource", thisGlobal_1.isCentralResource);
                            resolve();
                        }
                    },
                    error: function (err) {
                        //debugger
                        if (typeof console === "object") {
                            console.log("getSettingsData - ajax :" + err);
                        }
                        resolve();
                    }
                });
            }
            catch (e) {
                if (typeof console === "object") {
                    console.log("getSettingsData function:" + e);
                }
                resolve();
            }
        });
    };
    /**
     * Load data from sessionStorage
     *
     */
    MISMain.prototype.loadSessionStorage = function () {
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
        }
        else {
            // Sorry! No Web Storage support..
            if (typeof console === "object") {
                console.log("Sorry! No Web Storage support..");
            }
        }
    };
    MISMain.prototype.insertMenu = function () {
        console.log("insertMenu called");
        var siteUrl = "";
        if (this.isCentralResource == "true") {
            siteUrl = this.homeBtnUrl;
        }
        else {
            siteUrl = (this.spcontext.pageContext.web.absoluteUrl == "/") ? "" : this.spcontext.pageContext.web.absoluteUrl;
        }
        //debugger
        switch (this.navigationType) {
            case "Menu":
                //Load the CSS Style sheet file of the project
                //$('head').append("<link rel='stylesheet' type='text/css' href='" + siteUrl + "/Style Library/MIS.GlobalNavigation/css/MIS.GlobalNavigation.css'> </link>");
                //Load the Global Navigation
                //debugger
                var MISGlobalNavigationObject = new MISGlobalNavigation_1.MISGlobalNavigation();
                MISGlobalNavigationObject.init(this.spcontext, this.termStoreName, this.termStoreGUID);
                //init(termStoreName, termStoreGUID);//Global Navigation
                break;
            default: //Error in the configuration, nothing is done
        }
    };
    return MISMain;
}());
exports.MISMain = MISMain;

//# sourceMappingURL=MISMain.js.map
