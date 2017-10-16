define("a861c815-e425-416d-9520-04bcdf557e27_0.0.1", ["@microsoft/decorators","@microsoft/sp-application-base","@microsoft/sp-loader","jquery"], function(__WEBPACK_EXTERNAL_MODULE_1__, __WEBPACK_EXTERNAL_MODULE_2__, __WEBPACK_EXTERNAL_MODULE_3__, __WEBPACK_EXTERNAL_MODULE_6__) { return /******/ (function(modules) { // webpackBootstrap
/******/ 	// The module cache
/******/ 	var installedModules = {};
/******/
/******/ 	// The require function
/******/ 	function __webpack_require__(moduleId) {
/******/
/******/ 		// Check if module is in cache
/******/ 		if(installedModules[moduleId]) {
/******/ 			return installedModules[moduleId].exports;
/******/ 		}
/******/ 		// Create a new module (and put it into the cache)
/******/ 		var module = installedModules[moduleId] = {
/******/ 			i: moduleId,
/******/ 			l: false,
/******/ 			exports: {}
/******/ 		};
/******/
/******/ 		// Execute the module function
/******/ 		modules[moduleId].call(module.exports, module, module.exports, __webpack_require__);
/******/
/******/ 		// Flag the module as loaded
/******/ 		module.l = true;
/******/
/******/ 		// Return the exports of the module
/******/ 		return module.exports;
/******/ 	}
/******/
/******/
/******/ 	// expose the modules object (__webpack_modules__)
/******/ 	__webpack_require__.m = modules;
/******/
/******/ 	// expose the module cache
/******/ 	__webpack_require__.c = installedModules;
/******/
/******/ 	// define getter function for harmony exports
/******/ 	__webpack_require__.d = function(exports, name, getter) {
/******/ 		if(!__webpack_require__.o(exports, name)) {
/******/ 			Object.defineProperty(exports, name, {
/******/ 				configurable: false,
/******/ 				enumerable: true,
/******/ 				get: getter
/******/ 			});
/******/ 		}
/******/ 	};
/******/
/******/ 	// getDefaultExport function for compatibility with non-harmony modules
/******/ 	__webpack_require__.n = function(module) {
/******/ 		var getter = module && module.__esModule ?
/******/ 			function getDefault() { return module['default']; } :
/******/ 			function getModuleExports() { return module; };
/******/ 		__webpack_require__.d(getter, 'a', getter);
/******/ 		return getter;
/******/ 	};
/******/
/******/ 	// Object.prototype.hasOwnProperty.call
/******/ 	__webpack_require__.o = function(object, property) { return Object.prototype.hasOwnProperty.call(object, property); };
/******/
/******/ 	// __webpack_public_path__
/******/ 	__webpack_require__.p = "";
/******/
/******/ 	// Load entry module and return exports
/******/ 	return __webpack_require__(__webpack_require__.s = 0);
/******/ })
/************************************************************************/
/******/ ([
/* 0 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

var __extends = (this && this.__extends) || (function () {
    var extendStatics = Object.setPrototypeOf ||
        ({ __proto__: [] } instanceof Array && function (d, b) { d.__proto__ = b; }) ||
        function (d, b) { for (var p in b) if (b.hasOwnProperty(p)) d[p] = b[p]; };
    return function (d, b) {
        extendStatics(d, b);
        function __() { this.constructor = d; }
        d.prototype = b === null ? Object.create(b) : (__.prototype = b.prototype, new __());
    };
})();
var __decorate = (this && this.__decorate) || function (decorators, target, key, desc) {
    var c = arguments.length, r = c < 3 ? target : desc === null ? desc = Object.getOwnPropertyDescriptor(target, key) : desc, d;
    if (typeof Reflect === "object" && typeof Reflect.decorate === "function") r = Reflect.decorate(decorators, target, key, desc);
    else for (var i = decorators.length - 1; i >= 0; i--) if (d = decorators[i]) r = (c < 3 ? d(r) : c > 3 ? d(target, key, r) : d(target, key)) || r;
    return c > 3 && r && Object.defineProperty(target, key, r), r;
};
Object.defineProperty(exports, "__esModule", { value: true });
var decorators_1 = __webpack_require__(1);
var sp_application_base_1 = __webpack_require__(2);
//Needed to reference external CSS files
var sp_loader_1 = __webpack_require__(3);
var MISMain_1 = __webpack_require__(4);
var $ = __webpack_require__(6);
var LOG_SOURCE = 'SpfxToastrApplicationCustomizer';
//SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');
/** A Custom Action which can be run during execution of a Client Side Application */
var SpfxToastrApplicationCustomizer = (function (_super) {
    __extends(SpfxToastrApplicationCustomizer, _super);
    function SpfxToastrApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SpfxToastrApplicationCustomizer.prototype.onInit = function () {
        //debugger;
        //Load CSS
        sp_loader_1.SPComponentLoader.loadCss(this.getSiteCollectionUrl() + "/Style%20Library/MIS.GlobalNavigation.Modern/css/MIS.GlobalNavigationModern.css");
        // Added to handle possible changes on the existence of placeholders.
        this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
        return Promise.resolve();
    };
    SpfxToastrApplicationCustomizer.prototype.buildMegaMenu = function () {
        $(document).ready(function () {
            //DO NOT REMOVE ! IT IS REQUIRED TO INIT JQUERY
        });
        var instanceMis = new MISMain_1.MISMain(this.context);
    };
    SpfxToastrApplicationCustomizer.prototype._renderPlaceHolders = function () {
        var _this = this;
        // Handling the top placeholder
        if (!this._topPlaceholder) {
            this._topPlaceholder =
                this.context.placeholderProvider.tryCreateContent(sp_application_base_1.PlaceholderName.Top, { onDispose: this._onDispose });
            // The extension should not assume that the expected placeholder is available.
            if (!this._topPlaceholder) {
                console.error('The expected placeholder (Top) was not found.');
                return;
            }
            if (this.properties) {
                var topString = this.properties.Top;
                if (!topString) {
                    topString = '(Top property was not defined.)';
                }
                if (this._topPlaceholder.domElement) {
                    this._topPlaceholder.domElement.innerHTML = "<div id=\"MEGAMENU\" class=\"MEGAMENU\"></div>";
                }
            }
        }
        this._loadSPJSOMScripts().then(function () {
            //debugger;
            _this.buildMegaMenu();
        });
    };
    SpfxToastrApplicationCustomizer.prototype._onDispose = function () {
        //console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
    };
    SpfxToastrApplicationCustomizer.prototype.getSiteCollectionUrl = function () {
        var baseUrl = window.location.protocol + "//" + window.location.host;
        var pathname = window.location.pathname;
        var siteCollectionDetector = "/sites/";
        if (pathname.indexOf(siteCollectionDetector) >= 0) {
            baseUrl += pathname.substring(0, pathname.indexOf("/", siteCollectionDetector.length));
        }
        return baseUrl;
    };
    SpfxToastrApplicationCustomizer.prototype._loadSPJSOMScripts = function () {
        var _this = this;
        return new Promise(function (resolve, reject) {
            var siteColUrl = _this.getSiteCollectionUrl();
            try {
                sp_loader_1.SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
                    globalExportsName: '$_global_init'
                })
                    .then(function () {
                    return sp_loader_1.SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
                        globalExportsName: 'Sys'
                    });
                })
                    .then(function () {
                    return sp_loader_1.SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
                        globalExportsName: 'SP'
                    });
                })
                    .then(function () {
                    return sp_loader_1.SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
                        globalExportsName: 'SP'
                    });
                })
                    .then(function () {
                    return sp_loader_1.SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
                        globalExportsName: 'SP'
                    });
                })
                    .then(function () {
                    //this.setState({ loadingScripts: false });
                    resolve();
                })
                    .catch(function (reason) {
                    resolve();
                    //this.setState({ loadingScripts: false, errors: [...this.state.errors, reason] });
                });
            }
            catch (error) {
                //this.setState({ loadingScripts: false, errors: [...this.state.errors, error] });
            }
        });
    };
    __decorate([
        decorators_1.override
    ], SpfxToastrApplicationCustomizer.prototype, "onInit", null);
    return SpfxToastrApplicationCustomizer;
}(sp_application_base_1.BaseApplicationCustomizer));
exports.default = SpfxToastrApplicationCustomizer;



/***/ }),
/* 1 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_1__;

/***/ }),
/* 2 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_2__;

/***/ }),
/* 3 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_3__;

/***/ }),
/* 4 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var MISGlobalNavigation_1 = __webpack_require__(5);
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
            //////console.log("$(document).ready" + ex);
            if (typeof console === "object") {
                ////console.log("$(document).ready" + ex);
            }
            else {
                alert("$(document).ready" + ex);
            }
        }
    }
    //public static init(spcontext: any, _termStoreName: any, _termStoreGUID: any) {
    MISMain.prototype.getMenuParameters = function () {
        ////console.log("getMenuParameters called");
        var _this = this;
        try {
            this.loadSessionStorage;
            if (this.termStoreGUID == undefined && this.termStoreName == undefined) {
                //var url = _spPageContextInfo.webAbsoluteUrl;
                var url = (this.spcontext.pageContext.web.absoluteUrl == "/") ? "" : this.spcontext.pageContext.web.absoluteUrl;
                //debugger
                this.getSettingsData(url).then(function () {
                    //debugger
                    ////console.log("then getSettingsData");
                    ////console.log("isCentralresource value " + this.isCentralResource);
                    ////console.log("termStoreName value " + this.termStoreName);
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
                ////console.log("getMenuParameters function:" + e);
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
            ////console.log("gettingsData function called");
            try {
                var requestUri = url + "/_api/web/lists/getByTitle('GlobalNavigation')/Items?$filter=isActif eq 1";
                ////console.log(requestUri);
                // execute AJAX request synchrone
                //debugger;
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
                        ////console.log(data);
                        //debugger;
                        //debugger
                        if (data.d.results.length > 0) {
                            thisGlobal_1.listGUID = data.d.results[0].ListGUID;
                            thisGlobal_1.navigationType = data.d.results[0].NavigationType;
                            thisGlobal_1.termStoreName = data.d.results[0].TermStoreName;
                            ////console.log("1- set termstoreName with value " + data.d.results[0].TermStoreName)
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
                        else {
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
            }
            catch (e) {
                if (typeof console === "object") {
                    //console.log("getSettingsData function:" + e);
                }
                reject();
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
                ////console.log("Sorry! No Web Storage support..");
            }
        }
    };
    MISMain.prototype.insertMenu = function () {
        ////console.log("insertMenu called");
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
                //debugger;
                var MISGlobalNavigationObject = new MISGlobalNavigation_1.MISGlobalNavigation(siteUrl);
                MISGlobalNavigationObject.init(this.spcontext, this.termStoreName, this.termStoreGUID);
                //init(termStoreName, termStoreGUID);//Global Navigation
                break;
            default: //Error in the configuration, nothing is done
        }
    };
    return MISMain;
}());
exports.MISMain = MISMain;



/***/ }),
/* 5 */
/***/ (function(module, exports, __webpack_require__) {

"use strict";

Object.defineProperty(exports, "__esModule", { value: true });
var HelperCSOM = (function () {
    function HelperCSOM() {
    }
    HelperCSOM.createDelegate = function (instance, method) {
        return function () {
            return method.apply(instance, arguments);
        };
    };
    return HelperCSOM;
}());
exports.HelperCSOM = HelperCSOM;
var MISGlobalNavigation = (function () {
    function MISGlobalNavigation(siteUrl) {
        //debugger;
        this.viewModelObj = new ViewModel(siteUrl);
        //sessionStorage.clear();
    }
    /**
    * Create the HTML dom of the menu - Only 3 level of menu
    *
    * @param {string} menuItem - represents a item of the menu (this.MenuItem)
    * @param {int} nLevel - Level of the menu: 0, 1 or 2
    */
    MISGlobalNavigation.prototype.printTree = function (menuItem, nLevel) {
        //console.log("printTree()");
        var nbChild = menuItem.children.length;
        switch (nLevel) {
            case 0://Top Level => Executive, Core Processes, Support or Social
                this.viewModelObj.htmlStr += "<div class='separator-megamenu mis-hidden'></div>";
                this.viewModelObj.htmlStr += "<div class='container-megamenu'>";
                //this.viewModelObj.htmlStr + "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
                this.viewModelObj.htmlStr += "<p class='menu-level1'>" + menuItem.name;
                this.viewModelObj.htmlStr += "<img class='icon-level1' align='right' src='" + this.viewModelObj.siteUrl + "/Style Library/MIS.GlobalNavigation/images/icon-" + menuItem.name.replace(/[~#%&*:<>?/\{|}.]/g, "-") + ".png'></img>";
                this.viewModelObj.htmlStr += "</p>";
                this.viewModelObj.htmlStr += "<ul class='subs mis-hidden' style='list-style: none;'>";
                for (var i = 0; i < nbChild; i++) {
                    this.printTree(menuItem.children[i], nLevel + 1);
                }
                this.viewModelObj.htmlStr += "</ul>";
                this.viewModelObj.htmlStr += "</div>";
                break;
            case 1:// Header sub Menu
                if (menuItem.url === undefined) {
                    this.viewModelObj.htmlStr += "<li><p class='menu-level2'>" + menuItem.name + "</p>";
                }
                else {
                    this.viewModelObj.htmlStr += "<li><a href='" + menuItem.url + "'>" +
                        menuItem.name +
                        "<span>" +
                        /*"<img class='callout' src='/Style Library/MIS.GlobalNavigation/images/callout.gif' />" +
                        menuItem.description +*/
                        "</span>" +
                        "</a>";
                }
                this.viewModelObj.htmlStr += "<ul style='list-style: none;'>";
                for (var i = 0; i < nbChild; i++) {
                    this.printTree(menuItem.children[i], nLevel + 1);
                }
                this.viewModelObj.htmlStr += "</ul>";
                this.viewModelObj.htmlStr += "</li>";
                break;
            case 2:// sub-sub Menu
                this.viewModelObj.htmlStr += "<li>";
                if (menuItem.url === undefined) {
                    //console.log("cas 2 display menuITem.name" + menuItem.name);
                    this.viewModelObj.htmlStr += menuItem.name;
                }
                else {
                    this.viewModelObj.htmlStr += "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
                }
                this.viewModelObj.htmlStr += "</li>";
                break;
            //We doesn't go deep in the Navigation because we display only 3 levels of navigation
            default://Should not go in this case
                this.viewModelObj.htmlStr += "<!--Level: " + nLevel + " <a href='" + menuItem.url + "'>" + menuItem.name + "</a>-->"; //For the debug
        }
    };
    /**
    * Build the tree
    *
    */
    MISGlobalNavigation.prototype.buildTree = function () {
        //console.log("buildTree()");
        this.treeFromArray2(this.viewModelObj.globalMenuItems); //Load the treeNav
        //console.log("entries in globalMenuItems  " + this.viewModelObj.globalMenuItems.length)
        for (var i = 0; i < this.viewModelObj.treeNav.length; i++) {
            this.printTree(this.viewModelObj.treeNav[i], 0);
        }
        //Add the last separator in the menu
        this.viewModelObj.htmlStr += "<div class='separator-megamenu mis-hidden'></div>";
        //BackUp viewMenuModel in the sessionStorage
        //debugger;
        var viewMenuModel_json = JSON.stringify(this.viewModelObj);
        sessionStorage.setItem("viewMenuModel", viewMenuModel_json);
    };
    ;
    /**
    * Loop on the Array globalMenuItems to create a Tree
    *
    */
    MISGlobalNavigation.prototype.treeFromArray = function (list, idAttr, parentAttr, childrenAttr) {
        //console.log("treeFromArray()");
        if (!idAttr)
            idAttr = '_id';
        if (!parentAttr)
            parentAttr = 'sonOf';
        if (!childrenAttr)
            childrenAttr = 'children';
        var treeList = [];
        var lookup = {};
        $.each(list, function (index, obj) {
            lookup[obj[idAttr]] = obj;
            obj[childrenAttr] = [];
        });
        $.each(list, function (index, obj) {
            if (obj[parentAttr] != null) {
                lookup[obj[parentAttr]][childrenAttr].push(obj);
            }
            else {
                treeList.push(obj);
            }
        });
        //Sort the treeList
        for (var i = 0; i < treeList.length; i++) {
            treeList[i] = this.sortTermsFromTree(treeList[i]);
        }
        this.viewModelObj.treeNav = treeList; //Save the tree in the global variable treeNav
    };
    ;
    MISGlobalNavigation.prototype.treeFromArray2 = function (list) {
        //console.log("treeFromArray2()");
        var idAttr = '_id';
        var parentAttr = 'sonOf';
        var childrenAttr = 'children';
        var treeList = [];
        var lookup = {};
        $.each(list, function (index, obj) {
            lookup[obj[idAttr]] = obj;
            obj[childrenAttr] = [];
        });
        $.each(list, function (index, obj) {
            if (obj[parentAttr] != null) {
                lookup[obj[parentAttr]][childrenAttr].push(obj);
            }
            else {
                treeList.push(obj);
            }
        });
        //Sort the treeList
        for (var i = 0; i < treeList.length; i++) {
            treeList[i] = this.sortTermsFromTree(treeList[i]);
        }
        this.viewModelObj.treeNav = treeList; //Save the tree in the global variable treeNav
    };
    ;
    /**
     * Sort children array of a term tree by a sort order
     *
     * @param {obj} tree The term tree
     * @return {obj} A sorted term tree
     */
    MISGlobalNavigation.prototype.sortTermsFromTree = function (tree) {
        //console.log("sortTermsFromTree()");
        // Check to see if the get_customSortOrder function is defined. If the term is actually a term collection,
        // there is nothing to sort.
        if (tree.children.length && tree.customSortOrder) {
            var sortOrder = null;
            if (tree.customSortOrder) {
                sortOrder = tree.customSortOrder;
            }
            // If not null, the custom sort order is a string of GUIDs, delimited by a :
            if (sortOrder) {
                sortOrder = sortOrder.split(':');
                tree.children.sort(function (a, b) {
                    var indexA = sortOrder.indexOf(a.guid);
                    var indexB = sortOrder.indexOf(b.guid);
                    if (indexA > indexB) {
                        return 1;
                    }
                    else if (indexA < indexB) {
                        return -1;
                    }
                    return 0;
                });
            }
            else {
                tree.children.sort(function (a, b) {
                    if (a.name > b.name) {
                        return 1;
                    }
                    else if (a.name < b.name) {
                        return -1;
                    }
                    return 0;
                });
            }
        }
        for (var i = 0; i < tree.children.length; i++) {
            tree.children[i] = this.sortTermsFromTree(tree.children[i]);
        }
        return tree;
    };
    ;
    /**
    * Generate the html dom of the navigation
    *
    */
    MISGlobalNavigation.prototype.displayMenu = function () {
        //console.log("displayMenu called");
        ////var homeBtnTitle, homeBtnUrl, isCentralResource;
        if (!sessionStorage.viewMenuModel) {
            this.buildTree();
        }
        //Add link to Home page
        //Add link to Home page
        var menuInnerHtml = "<div id='nav-container'>" +
            this.viewModelObj.htmlStr +
            "</div>";
        //console.log(menuInnerHtml);
        //Insert icon on the left of the SharePoint text (Top left corner) + all menu html
        //debugger;
        $('#MEGAMENU').append(menuInnerHtml);
        this.applyNavigationEvent();
    };
    ;
    //Event on the Menu
    MISGlobalNavigation.prototype.applyNavigationEvent = function () {
        this.bindHoverMenu();
        //Apply event on the menu
        //#homeBtn:hover  #nav-container
        // $("#MEGAMENU").on('click', function () {
        //   $("#MEGAMENU").toggleClass("mis-on");
        //   $('#nav-container').toggleClass("mis-on");
        //   //$("nav ul").toggleClass('mis-hidden');
        // });
        $("#MEGAMENU img")
            .mouseover(function () {
            var src = $(this).attr("src").replace("menu.png", "menuover.png");
            $(this).attr("src", src);
        })
            .mouseout(function () {
            var src = $(this).attr("src").replace("over.png", ".png");
            $(this).attr("src", src);
        });
    };
    ;
    MISGlobalNavigation.prototype.bindHoverMenu = function () {
        //get every hidden menus
        var hiddenMenu = $('.mis-hidden');
        //on hover menu level, show the menu, on hover off, hide the menu
        $("#nav-container").hover(function () {
            hiddenMenu.removeClass("mis-hidden");
            $("#nav-container").css('cursor', 'pointer');
        }, function () {
            hiddenMenu.addClass("mis-hidden");
        });
    };
    /**
    * Return a string name of the term's parent and null if no parent
    *
    * @param {string} sPath - Path of the term: "Level1;SubLevel1-1;Term"
    * @param {bool} isRoot - true if the term is root
    */
    MISGlobalNavigation.prototype.getParentTerm = function (sPath, isRoot) {
        var termParentName;
        if (isRoot)
            termParentName = null;
        else {
            var arrayTerm = sPath.split(";");
            var indParent = 0;
            if (arrayTerm.length > 1)
                indParent = arrayTerm.length - 2;
            termParentName = arrayTerm[indParent];
        }
        return termParentName;
    };
    ;
    /**
     * Load data from sessionStorage
     *
     */
    MISGlobalNavigation.prototype.loadSessionStorage = function (spcontext) {
        //debugger;
        if (typeof (Storage) !== "undefined") {
            // Code for localStorage/sessionStorage.
            if (sessionStorage.viewMenuModel) {
                var viewMenuModel_json = sessionStorage.getItem("viewMenuModel");
                this.viewModelObj = JSON.parse(viewMenuModel_json);
            }
            if (sessionStorage.isCentralResource) {
                this.viewModelObj.isCentralResource = (sessionStorage.getItem("isCentralResource") == "true") ? true : false;
            }
            if (sessionStorage.homeBtnUrl) {
                if (this.viewModelObj.isCentralResource) {
                    this.viewModelObj.siteUrl = sessionStorage.getItem("homeBtnUrl");
                }
                else {
                    this.viewModelObj.siteUrl = (spcontext.pageContext.web.absoluteUrl == "/") ? "" : spcontext.pageContext.web.absoluteUrl;
                }
            }
            if (sessionStorage.homeBtnTitle) {
                this.viewModelObj.homeBtnTitle = sessionStorage.getItem("homeBtnTitle");
                ;
            }
        }
        else {
            // Sorry! No Web Storage support..
            if (typeof console === "object") {
                //console.log("Sorry! No Web Storage support..");
            }
        }
    };
    ;
    /**
     * Initialisation to get information from the TermStore
     *
     * @param {string} termStoreName - Name of the term store service : 'Service de métadonnées gérées'
     * @param {object} callback - Callback function to call upon completion and pass termset into
     */
    MISGlobalNavigation.prototype.init = function (spcontext, termStoreName, termSetId) {
        //console.log("init called");
        //Load sessionStorage
        var _this = this;
        this.loadSessionStorage(spcontext);
        //Load all terms  if not already done.
        //if (!this.viewModelObj.globalMenuItems) {
        //console.log("no terms loading it");
        this.buildTermsModel(spcontext, termStoreName, termSetId).then(function () {
            //to do rendre asynchrone
            //setTimeout(()=>{}, 2000);
            _this.displayMenu();
        });
        //ko.applyBindings(this.viewModelObj);
        // SP.UI.Notify.removeNotification(nid);
        // }), Function.createDelegate(this, function (sender, args) {
        //    //console.log('The following error has occured while loading global navigation: ' + args.get_message());
        //  }));
        // }));
        //  }, 'core.js');
        // }
        // else {
        //  this.displayMenu();
        //}
    };
    ;
    MISGlobalNavigation.prototype.buildTermsModel = function (spcontext, termStoreName, termSetId) {
        var _this = this;
        //SP.SOD.executeOrDelayUntilScriptLoaded(function () {
        // 'use strict';
        // var nid = SP.UI.Notify.addNotification("<img src='/_layouts/15/images/loadingcirclests16.gif?rev=23' style='vertical-align:bottom; display:inline-block; margin-" + (document.documentElement.dir == "rtl" ? "left" : "right") + ":2px;' />&nbsp;<span style='vertical-align:top;'>Loading navigation...</span>", false);
        //  SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));
        // SP.SOD.executeFunc('sp.taxonomy.js', false, Function.createDelegate(this, function () {
        var globalThis = this;
        return new Promise(function (resolve, reject) {
            var context = new SP.ClientContext(spcontext.pageContext.web.absoluteUrl);
            //var context = SP.ClientContext.get_current();
            var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
            var termStore = taxonomySession.get_termStores().getByName(termStoreName);
            var termSet = termStore.getTermSet(termSetId);
            var terms = termSet.getAllTerms();
            context.load(terms);
            //context.executeQueryAsync(function name(sender, args) {
            context.executeQueryAsync(HelperCSOM.createDelegate(_this, function (sender, args) {
                var termsEnumerator = terms.getEnumerator();
                resolve();
                while (termsEnumerator.moveNext()) {
                    var currentTerm = termsEnumerator.get_current();
                    var termIsRoot = currentTerm.get_isRoot();
                    var termParentName = this.getParentTerm(currentTerm.get_pathOfTerm(), termIsRoot);
                    globalThis.viewModelObj.globalMenuItems.push(new MenuItem(currentTerm.get_name(), //title
                    currentTerm.get_localCustomProperties()['_Sys_Nav_SimpleLinkUrl'], //url
                    termIsRoot, //currentTerm.get_isRoot(),//isRoot
                    currentTerm.get_customSortOrder(), //customSortOrder
                    termParentName, //currentTerm.get_localCustomProperties()['sonOf']//sonOf
                    currentTerm.get_id().toString(), //guid of the term
                    currentTerm.get_localCustomProperties()['description']));
                    //console.log("Push in globalMenuItems " + currentTerm.get_name());
                    // }
                }
                //console.log("GlobalMenuItems" + globalThis.viewModelObj.globalMenuItems.length);
                //Call the logical function to display the menu
            }));
        });
    };
    return MISGlobalNavigation;
}());
exports.MISGlobalNavigation = MISGlobalNavigation;
/**
  * Represent an item of the menu
  *
  */
var MenuItem = (function () {
    function MenuItem(title, url, isRoot, customSortOrder, sonOf, guid, description) {
        this.name = title;
        this._id = title;
        this.guid = guid;
        this.url = url;
        this.isRoot = isRoot;
        this.customSortOrder = customSortOrder;
        this.sonOf = sonOf;
        this.children = [];
        this.description = description;
    }
    return MenuItem;
}());
exports.MenuItem = MenuItem;
var ViewModel = (function () {
    function ViewModel(siteUrl) {
        this.htmlStr = "<ul id='mis-nav'>";
        this.globalMenuItems = new Array();
        this.treeNav = new Array();
        this.siteUrl = siteUrl;
    }
    return ViewModel;
}());
exports.ViewModel = ViewModel;



/***/ }),
/* 6 */
/***/ (function(module, exports) {

module.exports = __WEBPACK_EXTERNAL_MODULE_6__;

/***/ })
/******/ ])});;
//# sourceMappingURL=spfx-toastr-application-customizer.js.map