import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer,
  PlaceholderContent,
  PlaceholderName
} from '@microsoft/sp-application-base';
import { IWebPartContext } from '@microsoft/sp-webpart-base';
import * as strings from 'SpfxToastrApplicationCustomizerStrings';
//Needed to reference external CSS files
import { SPComponentLoader } from '@microsoft/sp-loader';

import { MISMain } from './MISMain';
import * as $ from 'jquery';

// import * as toastr from 'toastr';
// import styles from './SpfxToastr.module.scss';
//import { IToast, ToastService } from '../../services/toastService'; //loaded from the toastService barrel - temporarily disabled due to issue with WebPack
import { IToast } from '../../services/toastService/IToast';
import { ToastService } from '../../services/toastService/ToastService';

const LOG_SOURCE: string = 'SpfxToastrApplicationCustomizer';

require('sp-init');
require('microsoft-ajax');
require('sp-runtime');
require('sharepoint');

// SPComponentLoader.loadCss('https://samavangarde.sharepoint.com/sites/devjeff/Style%20Library/MIS.GlobalNavigation/css/MIS.GlobalNavigation.css')
/**
 * If your command set uses the ClientSideComponentProperties JSON input,
 * it will be deserialized into the BaseExtension.properties object.
 * You can define an interface to describe it.
 */
export interface ISpfxToastrApplicationCustomizerProperties {
  Top: string;
  Bottom: string;
}

//SPComponentLoader.loadCss('https://cdnjs.cloudflare.com/ajax/libs/toastr.js/latest/toastr.min.css');

/** A Custom Action which can be run during execution of a Client Side Application */
export default class SpfxToastrApplicationCustomizer
  extends BaseApplicationCustomizer<ISpfxToastrApplicationCustomizerProperties> {

  private toastsPromise: Promise<IToast[]>;

  private _headerPlaceholder: PlaceholderContent;
  private _topPlaceholder: PlaceholderContent | undefined;
  private _bottomPlaceholder: PlaceholderContent | undefined;

  @override
  public onInit(): Promise<void> {
    //Log.info(LOG_SOURCE, `Initialized v1.0.4 ${strings.Title}`);

    Log.info(LOG_SOURCE, `After changedEvent`);

    // Call render method for generating the HTML elements.
    //this._renderPlaceHolders();


    // Added to handle possible changes on the existence of placeholders.
    this.context.placeholderProvider.changedEvent.add(this, this._renderPlaceHolders);


    Log.info(LOG_SOURCE, `After _renderPlaceHolders`);
    //Dialog.alert(`hello from SPFX`);
    return Promise.resolve<void>();
  }

  private buildMegaMenu(): void {

    console.log("v1.5");
    $('.ms-bgColor-themeDark').text("tete");
    $('#suiteBarLeft .ms-core-brandingText').remove();

    $(document).ready(function () {
      //alert("test");
    });

    let instanceMis: MISMain = new MISMain(this.context);

    const context: SP.ClientContext = new SP.ClientContext(this.context.pageContext.web.absoluteUrl);
    const web: SP.Web = context.get_web();
    context.load(web);
    context.executeQueryAsync(function name(sender, args) {
      this.webpartTitle = web.get_title();
      console.log(web.get_description());
      let siteContent: string = `<div>
                                 <h2>Title: ${web.get_title()}</h2>
                                 <span>Description: ${web.get_description()}<span>
                               </div>`;
      //htmlContext.domElement.querySelector("#siteContent").innerHTML =siteContent;
    },
      function (sender, args) {
        console.log(args.get_message());
      });

  }
  private _renderPlaceHolders(): void {

    console.log('HelloWorldApplicationCustomizer._renderPlaceHolders()');
    console.log('Available placeholders: ',
      this.context.placeholderProvider.placeholderNames.map(name => PlaceholderName[name]).join(', '));

    // Handling the top placeholder
    if (!this._topPlaceholder) {
      this._topPlaceholder =
        this.context.placeholderProvider.tryCreateContent(
          PlaceholderName.Top,
          { onDispose: this._onDispose });

      // The extension should not assume that the expected placeholder is available.
      if (!this._topPlaceholder) {
        console.error('The expected placeholder (Top) was not found.');
        return;
      }

      if (this.properties) {
        let topString: string = this.properties.Top;
        if (!topString) {
          topString = '(Top property was not defined.)';
        }

        if (this._topPlaceholder.domElement) {
          this._topPlaceholder.domElement.innerHTML = `
        <div id="MEGAMENU" class="">
          <div class="ms-bgColor-themeDark ms-fontColor-white ">
            This is my place holder

            // <i class="ms-Icon ms-Icon--Info" aria-hidden="true"></i>
          </div>
        </div>`;
        }
      }
    }

    this._loadSPJSOMScripts().then(() => {

      this.buildMegaMenu();
    })
  }

  private _onDispose(): void {
    console.log('[HelloWorldApplicationCustomizer._onDispose] Disposed custom top and bottom placeholders.');
  }

  public getSiteCollectionUrl(): string {
    let baseUrl = window.location.protocol + "//" + window.location.host;
    const pathname = window.location.pathname;
    const siteCollectionDetector = "/sites/";
    if (pathname.indexOf(siteCollectionDetector) >= 0) {
      baseUrl += pathname.substring(0, pathname.indexOf("/", siteCollectionDetector.length));
    }
    return baseUrl;
  }

  private _loadSPJSOMScripts(): Promise<void> {
    return new Promise<void>((resolve, reject) => {
      const siteColUrl = this.getSiteCollectionUrl();
      try {
        SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/init.js', {
          globalExportsName: '$_global_init'
        })
          .then((): Promise<void> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/MicrosoftAjax.js', {
              globalExportsName: 'Sys'
            });
          })
          .then((): Promise<void> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.Runtime.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<void> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): Promise<void> => {
            return SPComponentLoader.loadScript(siteColUrl + '/_layouts/15/SP.taxonomy.js', {
              globalExportsName: 'SP'
            });
          })
          .then((): void => {
            //this.setState({ loadingScripts: false });
            resolve();
          })
          .catch((reason: any) => {
            resolve();
            //this.setState({ loadingScripts: false, errors: [...this.state.errors, reason] });
          });
      } catch (error) {
        //this.setState({ loadingScripts: false, errors: [...this.state.errors, error] });
      }
    });
  }
}

