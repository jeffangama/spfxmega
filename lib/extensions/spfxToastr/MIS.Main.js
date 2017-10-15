"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var MISMain = (function () {
    function MISMain() {
    }
    MISMain.init = function () {
        var listGUID, navigationType, //'Menu' or 'Wiki'
        termStoreGUID, termStoreName, //Service de métadonnées gérées
        homeBtnTitle, homeBtnUrl, isCentralResource;
        function getMenuParameters() {
            console.log("getMenuParameters called");
            try {
                loadSessionStorage();
                if (termStoreGUID == undefined && termStoreName == undefined) {
                    //var url = _spPageContextInfo.webAbsoluteUrl;
                    var url = (this.context.pageContext.web.absoluteUrl == "/") ? "" : this.context.pageContext.web.absoluteUrl;
                    getSettingsData(url); //Call GlobalNavigation settings data
                    if (isCentralResource == true && url != homeBtnUrl) {
                        getSettingsData(homeBtnUrl); //Call GlobalNavigation settings data in the Central ressource url
                    }
                }
                //Insert the menu in the top navigation
                insertMenu();
            }
            catch (e) {
                if (typeof console === "object") {
                    console.log("getMenuParameters function:" + e);
                }
            }
        }
        /**
         * Load data from List GlobalNavigation URL and update the sessionStorage
         *
         */
        function getSettingsData(url) {
            try {
                var requestUri = url + "/_api/web/lists/getByTitle('GlobalNavigation')/Items?$filter=isActif eq 1";
                // execute AJAX request synchrone
                $.ajax({
                    url: requestUri,
                    type: "GET",
                    headers: { "ACCEPT": "application/json;odata=verbose" },
                    async: false,
                    success: function (data) {
                        if (data.d.results.length > 0) {
                            listGUID = data.d.results[0].ListGUID;
                            navigationType = data.d.results[0].NavigationType;
                            termStoreName = data.d.results[0].TermStoreName;
                            termStoreGUID = data.d.results[0].TermStoreGUID;
                            homeBtnTitle = data.d.results[0].HomeBtnTitle;
                            homeBtnUrl = data.d.results[0].HomeBtnUrl;
                            isCentralResource = data.d.results[0].CentralResource;
                            //Backup Data in session.Storage
                            sessionStorage.setItem("listGUID", listGUID);
                            sessionStorage.setItem("navigationType", navigationType);
                            sessionStorage.setItem("termStoreName", termStoreName);
                            sessionStorage.setItem("termStoreGUID", termStoreGUID);
                            sessionStorage.setItem("homeBtnTitle", homeBtnTitle);
                            sessionStorage.setItem("homeBtnUrl", homeBtnUrl);
                            sessionStorage.setItem("isCentralResource", isCentralResource);
                        }
                    },
                    error: function (err) {
                        if (typeof console === "object") {
                            console.log("getSettingsData - ajax :" + err);
                        }
                    }
                });
            }
            catch (e) {
                if (typeof console === "object") {
                    console.log("getSettingsData function:" + e);
                }
            }
        }
        /**
         * Load data from sessionStorage
         *
         */
        function loadSessionStorage() {
            if (typeof (Storage) !== "undefined") {
                if (sessionStorage.termStoreGUID) {
                    termStoreGUID = sessionStorage.getItem("termStoreGUID");
                }
                if (sessionStorage.termStoreName) {
                    termStoreName = sessionStorage.getItem("termStoreName");
                }
                if (sessionStorage.navigationType) {
                    navigationType = sessionStorage.getItem("navigationType");
                }
                if (sessionStorage.listGUID) {
                    listGUID = sessionStorage.getItem("listGUID");
                }
                if (sessionStorage.homeBtnTitle) {
                    homeBtnTitle = sessionStorage.getItem("homeBtnTitle");
                }
                if (sessionStorage.homeBtnUrl) {
                    homeBtnUrl = sessionStorage.getItem("homeBtnUrl");
                }
                if (sessionStorage.isCentralResource) {
                    isCentralResource = sessionStorage.getItem("isCentralResource");
                }
            }
            else {
                // Sorry! No Web Storage support..
                if (typeof console === "object") {
                    console.log("Sorry! No Web Storage support..");
                }
            }
        }
        function insertMenu() {
            var siteUrl = "";
            if (isCentralResource == "true") {
                siteUrl = homeBtnUrl;
            }
            else {
                siteUrl = (_spPageContextInfo.siteServerRelativeUrl == "/") ? "" : _spPageContextInfo.siteServerRelativeUrl;
            }
            switch (navigationType) {
                case "Menu":
                    //Load the CSS Style sheet file of the project
                    $('head').append("<link rel='stylesheet' type='text/css' href='" + siteUrl + "/Style Library/MIS.GlobalNavigation/css/MIS.GlobalNavigation.css'> </link>");
                    //Load the Global Navigation
                    //MISGlobalNavigation.init('Service de métadonnées gérées', '1bdb0c2c-5ad5-4b03-9ee6-cd4e0a29a520');
                    //init(termStoreName, termStoreGUID);//Global Navigation
                    break;
                default: //Error in the configuration, nothing is done
            }
        }
        $(document).ready(function () {
            try {
                //Remove Sharepoint text if exist SP 2013
                $('#suiteBarLeft .ms-core-brandingText').remove();
                //Contain the webapp where mysites is configured
                var privateUrl = "https://private.monacoinformatiqueservice.mc";
                //Change the url of the top link "Sites"
                $("#Sites_BrandBar").attr('href', privateUrl + '/_layouts/15/MySite.aspx?MySiteRedirect=AllSites');
                //Change the url of the Tile "Sites"
                $("#O365_AppTile_Sites").attr('href', privateUrl + '/_layouts/15/MySite.aspx?MySiteRedirect=AllSites');
                //Change the url of the Tile "OneDrive"
                //$("#O365_AppTile_Documents").attr('href', privateUrl+'/_layouts/15/MySite.aspx?MySiteRedirect=AllDocuments');
                //Change the url of the Tile "Flux d'actualité"
                $("#O365_AppTile_Newsfeed").attr('href', privateUrl + '/_layouts/15/MySite.aspx');
                //20171002 - Change Sites link, Jeff ANGAMA
                //$('#ShellSites_BrandBar').remove(); //id du privee
                //$('#Sites_BrandBar').remove(); //id du public
                $('#ShellSites_BrandBar').attr('href', privateUrl + '/_layouts/15/MySite.aspx?MySiteRedirect=AllSites');
                //hide sharepoint text
                $('.o365cs-nav-appTitleLine').hide();
                $('.o365cs-nav-brandingText:first').text("");
                //To be sure
                //SP.SOD.registerSod("sp.js", "/" + _spPageContextInfo.layoutsUrl + "/sp.js");
                //SP.SOD.executeFunc('sp.js', false, function () { });
                ExecuteOrDelayUntilScriptLoaded(getMenuParameters, "sp.js");
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
        });
    };
    return MISMain;
}());
exports.default = MISMain;

//# sourceMappingURL=MIS.Main.js.map
