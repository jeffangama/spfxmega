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
var decorators_1 = require("@microsoft/decorators");
var sp_core_library_1 = require("@microsoft/sp-core-library");
var sp_application_base_1 = require("@microsoft/sp-application-base");
//Needed to reference external CSS files
var sp_loader_1 = require("@microsoft/sp-loader");
var MISMain_1 = require("./MISMain");
var $ = require("jquery");
var LOG_SOURCE = 'SpfxToastrApplicationCustomizer';
require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');
sp_loader_1.SPComponentLoader.loadCss('https://samavangarde.sharepoint.com/sites/devjeff/Style%20Library/MIS.GlobalNavigation/css/MIS.GlobalNavigationModern.css');
//SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');
/** A Custom Action which can be run during execution of a Client Side Application */
var SpfxToastrApplicationCustomizer = (function (_super) {
    __extends(SpfxToastrApplicationCustomizer, _super);
    function SpfxToastrApplicationCustomizer() {
        return _super !== null && _super.apply(this, arguments) || this;
    }
    SpfxToastrApplicationCustomizer.prototype.onInit = function () {
        //Log.info(LOG_SOURCE, `Initialized v1.0.4 ${strings.Title}`);
        sp_core_library_1.Log.info(LOG_SOURCE, "After changedEvent");
        // Call render method for generating the HTML elements.
        //this._renderPlaceHolders();
        // Added to handle possible changes on the existence of placeholders.
        this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);
        sp_core_library_1.Log.info(LOG_SOURCE, "After _renderPlaceHolders");
        //Dialog.alert(`hello from SPFX`);
        return Promise.resolve();
    };
    SpfxToastrApplicationCustomizer.prototype.buildMegaMenu = function () {
        //$('.ms-bgColor-themeDark').text("tete");
        $(document).ready(function () {
            //alert("test");
        });
        var instanceMis = new MISMain_1.MISMain(this.context);
        // const context: SP.ClientContext = new SP.ClientContext(this.context.pageContext.web.absoluteUrl);
        // const web: SP.Web = context.get_web();
        // context.load(web);
        // context.executeQueryAsync(function name(sender, args) {
        //   //this.webpartTitle = web.get_title();
        //   //console.log(web.get_description());
        //   // let siteContent: string = `<div>
        //   //                            <h2>Title: ${web.get_title()}</h2>
        //   //                            <span>Description: ${web.get_description()}<span>
        //   //                          </div>`;
        //   //htmlContext.domElement.querySelector("#siteContent").innerHTML =siteContent;
        // },
        //   function (sender, args) {
        //     console.log(args.get_message());
        //   });
    };
    SpfxToastrApplicationCustomizer.prototype._renderPlaceHolders = function () {
        var _this = this;
        console.log('HelloWorldApplicationCustomizer._renderPlaceHolders()');
        console.log('Available placeholders: ', this.context.placeholderProvider.placeholderNames.map(function (name) { return sp_application_base_1.PlaceholderName[name]; }).join(', '));
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
                    // this._topPlaceholder.domElement.innerHTML = `
                    // <div id="MEGAMENU" class="">
                    //   <div class="ms-bgColor-themeDark ms-fontColor-white ">
                    //     This is my place holder
                    //     // <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>
                    //   </div>
                    // </div>`;
                }
            }
        }
        this._loadSPJSOMScripts().then(function () {
            _this.buildMegaMenu();
        });
    };
    SpfxToastrApplicationCustomizer.prototype._onDispose = function () {
        console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
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

//# sourceMappingURL=SpfxToastrApplicationCustomizer.js.map
