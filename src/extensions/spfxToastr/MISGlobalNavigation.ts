
function createDelegate(instance: any, method: any): any {
  return function () {
    return method.apply(instance, arguments);
  }
}

export class MISGlobalNavigation {
  /**
  * Global Variable of the Menu
  *
  */

  //var treeNav = [];
  //var htmlStr = "<ul id='mis-nav' class='hidden'>";

  /**
  * Create the HTML dom of the menu - Only 3 level of menu
  *
  * @param {string} menuItem - represents a item of the menu (this.MenuItem)
  */

  public viewModelObj: ViewModel;

  constructor() {
    this.viewModelObj = new ViewModel();
    this.viewModelObj.globalMenuItems = new Array<any>();
    this.viewModelObj.treeNav = new Array<any>();
    sessionStorage.clear();
  }

  /**
  * Create the HTML dom of the menu - Only 3 level of menu
  *
  * @param {string} menuItem - represents a item of the menu (this.MenuItem)
  * @param {int} nLevel - Level of the menu: 0, 1 or 2
  */
  private printTree(menuItem: MenuItem, nLevel) {
    console.log("printTree()");
    this.viewModelObj.htmlStr += "<ul id='mis-nav' class='mis-hidden'>";
    var nbChild = menuItem.children.length;


    switch (nLevel) {
      case 0: //Top Level => Executive, Core Processes, Support or Social
        this.viewModelObj.htmlStr += "<div class='separator-megamenu'></div>";
        this.viewModelObj.htmlStr += "<div class='container-megamenu'>";
        //this.viewModelObj.htmlStr + "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
        this.viewModelObj.htmlStr += "<p class='menu-level1'>" + menuItem.name;
        this.viewModelObj.htmlStr += "<img class='icon-level1' align='right' src='" + this.viewModelObj.siteUrl + "/Style Library/MIS.GlobalNavigation/images/icon-" + menuItem.name.replace(/[~#%&*:<>?/\{|}.]/g, "-") + ".png'></img>";
        this.viewModelObj.htmlStr += "</p>";
        this.viewModelObj.htmlStr += "<ul class='subs' style='list-style: none;'>";

        for (var i = 0; i < nbChild; i++) {
          this.printTree(menuItem.children[i], nLevel + 1);
        }
        this.viewModelObj.htmlStr += "</ul>";
        this.viewModelObj.htmlStr += "</div>";
        break;
      case 1: // Header sub Menu
        if (menuItem.url === undefined) {
          this.viewModelObj.htmlStr += "<li><p class='menu-level2'>" + menuItem.name + "</p>";
        }
        else {
          this.viewModelObj.htmlStr += "<li><a href='" + menuItem.url + "' class='mis-tooltip'>" +
            menuItem.name +
            "<span>" +
            "<img class='callout' src='" + this.viewModelObj.siteUrl + "/Style Library/MIS.GlobalNavigation/images/callout.gif' />" +
            menuItem.description +
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
      case 2: // sub-sub Menu
        this.viewModelObj.htmlStr += "<li>";
        if (menuItem.url === undefined) {
          this.viewModelObj.htmlStr += menuItem.name;

        }
        else {

          this.viewModelObj.htmlStr += "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
        }
        this.viewModelObj.htmlStr += "</li>";
        break;
      //We doesn't go deep in the Navigation because we display only 3 levels of navigation
      default: //Should not go in this case
        this.viewModelObj.htmlStr += "<!--Level: " + nLevel + " <a href='" + menuItem.url + "'>" + menuItem.name + "</a>-->";//For the debug
    }
  }

  /**
  * Build the tree
  *
  */
  private buildTree() {
    console.log("buildTree()");
    this.treeFromArray2(this.viewModelObj.globalMenuItems);//Load the treeNav
    console.log("entries in globalMenuItems  " + this.viewModelObj.globalMenuItems.length)
    for (var i = 0; i < this.viewModelObj.treeNav.length; i++) {
      this.printTree(this.viewModelObj.treeNav[i], 0);
    }
    //Add the last separator in the menu
    this.viewModelObj.htmlStr += "<div class='separator-megamenu'></div>";


    //BackUp viewMenuModel in the sessionStorage
    var viewMenuModel_json = JSON.stringify(this.viewModelObj);
    sessionStorage.setItem("viewMenuModel", viewMenuModel_json);
  };

  /**
  * Loop on the Array globalMenuItems to create a Tree
  *
  */
  private treeFromArray(list, idAttr, parentAttr, childrenAttr) {
    console.log("treeFromArray()");
    if (!idAttr) idAttr = '_id';
    if (!parentAttr) parentAttr = 'sonOf';
    if (!childrenAttr) childrenAttr = 'children';
    var treeList = [];
    var lookup = {};
    $.each(list, function (index, obj) {
      lookup[obj[idAttr]] = obj;
      obj[childrenAttr] = [];
    });
    $.each(list, function (index, obj) {
      if (obj[parentAttr] != null) {
        lookup[obj[parentAttr]][childrenAttr].push(obj);
      } else {
        treeList.push(obj);
      }
    });
    //Sort the treeList
    for (var i = 0; i < treeList.length; i++) {
      treeList[i] = this.sortTermsFromTree(treeList[i]);
    }
    this.viewModelObj.treeNav = treeList; //Save the tree in the global variable treeNav

  };

  private treeFromArray2(list) {
    console.log("treeFromArray2()");
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
      } else {
        treeList.push(obj);
      }
    });
    //Sort the treeList
    for (var i = 0; i < treeList.length; i++) {
      treeList[i] = this.sortTermsFromTree(treeList[i]);
    }
    this.viewModelObj.treeNav = treeList; //Save the tree in the global variable treeNav

  };

  /**
   * Sort children array of a term tree by a sort order
   *
   * @param {obj} tree The term tree
   * @return {obj} A sorted term tree
   */

  private sortTermsFromTree(tree) {
    console.log("sortTermsFromTree()");
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
          } else if (indexA < indexB) {
            return -1;
          }

          return 0;
        });
      }
      // If null, terms are just sorted alphabetically
      else {
        tree.children.sort(function (a, b) {
          if (a.name > b.name) {
            return 1;
          } else if (a.name < b.name) {
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

  /**
  * Generate the html dom of the navigation
  *
  */
  private displayMenu() {
    console.log("displayMenu called");
    ////var homeBtnTitle, homeBtnUrl, isCentralResource;

    if (!sessionStorage.viewMenuModel) {
      this.buildTree();
    }

    //Add link to Home page


    //Add link to Home page
     var menuInnerHtml =
      "<div id='nav-container'>" +
      this.viewModelObj.htmlStr +
      "</div>";

    //Insert icon on the left of the SharePoint text (Top left corner) + all menu html
    $('#MEGAMENU').append(menuInnerHtml);
    this.applyNavigationEvent();
    ////If SharePoint Online
    //var isSPOnline = $("#O365_NavHeader div.o365cs-nav-leftAlign");
    //if (isSPOnline.length > 0) {
    //    var megaMenuHtml = "<div class='o365cs-nav-topItem o365cs-rsp-tw-hide o365cs-rsp-tn-hide' id='mis-GlobalNavigation' style='min-width: 155px;'></div>";
    //    $("#O365_NavHeader div.o365cs-nav-leftAlign").prepend(megaMenuHtml);

    //    //CSS Customisation
    //    $("#homeBtn").css("left", "10px");
    //    $("#homeBtn").css("top", "10px");
    //    $("#homeBtnWebcenter").css("left", "75px");
    //    $("#homeBtnWebcenter").css("top", "10px");

    //}



  };

  //Event on the Menu
  private applyNavigationEvent() {
    //Apply event on the menu
    //#homeBtn:hover  #nav-container
    $("#MEGAMENU").on('click', function () {
      $("#MEGAMENU").toggleClass("mis-on");
      $('#nav-container').toggleClass("mis-on");
      //$("nav ul").toggleClass('mis-hidden');
    });

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

  /**
  * Return a string name of the term's parent and null if no parent
  *
  * @param {string} sPath - Path of the term: "Level1;SubLevel1-1;Term"
  * @param {bool} isRoot - true if the term is root
  */
  private getParentTerm(sPath, isRoot) {
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

  /**
   * Load data from sessionStorage
   *
   */
  private loadSessionStorage(spcontext: any) {
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
        this.viewModelObj.homeBtnTitle = sessionStorage.getItem("homeBtnTitle");;
      }

    } else {
      // Sorry! No Web Storage support..
      if (typeof console === "object") {
        console.log("Sorry! No Web Storage support..");
      }
    }
  };

  /**
   * Initialisation to get information from the TermStore
   *
   * @param {string} termStoreName - Name of the term store service : 'Service de métadonnées gérées'
   * @param {object} callback - Callback function to call upon completion and pass termset into
   */
  public init(spcontext, termStoreName, termSetId) {
    console.log("init called");
    //Load sessionStorage


    this.loadSessionStorage(spcontext);

    //Load all terms  if not already done.
    //if (!this.viewModelObj.globalMenuItems) {
    console.log("no terms loading it");

    this.buildTermsModel(spcontext, termStoreName, termSetId).then(() => {
      //to do rendre asynchrone
      //setTimeout(()=>{}, 2000);
      this.displayMenu();
    })

    //ko.applyBindings(this.viewModelObj);
    // SP.UI.Notify.removeNotification(nid);
    // }), Function.createDelegate(this, function (sender, args) {
    //    console.log('The following error has occured while loading global navigation: ' + args.get_message());
    //  }));
    // }));
    //  }, 'core.js');
    // }
    // else {
    //  this.displayMenu();
    //}
  };

  private buildTermsModel(spcontext: any, termStoreName: any, termSetId: any): Promise<void> {
    //SP.SOD.executeOrDelayUntilScriptLoaded(function () {
    // 'use strict';
    // var nid = SP.UI.Notify.addNotification("<img src='/_layouts/15/images/loadingcirclests16.gif?rev=23' style='vertical-align:bottom; display:inline-block; margin-" + (document.documentElement.dir == "rtl" ? "left" : "right") + ":2px;' />&nbsp;<span style='vertical-align:top;'>Loading navigation...</span>", false);
    //  SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));
    // SP.SOD.executeFunc('sp.taxonomy.js', false, Function.createDelegate(this, function () {
    let globalThis = this;
    return new Promise<void>((resolve, reject) => {
      const context: SP.ClientContext = new SP.ClientContext(spcontext.pageContext.web.absoluteUrl);
      //var context = SP.ClientContext.get_current();
      var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
      var termStore = taxonomySession.get_termStores().getByName(termStoreName);
      var termSet = termStore.getTermSet(termSetId);
      var terms = termSet.getAllTerms();
      context.load(terms);
      //context.executeQueryAsync(function name(sender, args) {
      context.executeQueryAsync(createDelegate(this, function (sender, args) {
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
          console.log("Push in globalMenuItems " + currentTerm.get_name());
          // }
        }
        console.log("GlobalMenuItems" + globalThis.viewModelObj.globalMenuItems.length);
        //Call the logical function to display the menu
      }));


    });
  }
}

/**
  * Represent an item of the menu
  *
  */
export class MenuItem {

  name;
  _id;
  title;
  url;
  customSortOrder;
  isRoot;
  children;
  customSortOrde;
  sonOf;
  guid;
  description;

  constructor(title, url, isRoot, customSortOrder, sonOf, guid, description) {
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
}

export class ViewModel {
  public globalMenuItems: Array<any>; //Contain all the Navigation termset with set of Title, url, isRoot, parent, sonOf
  public treeNav: Array<any>;
  public htmlStr: any;
  public isCentralResource: boolean;
  public siteUrl: any;
  public homeBtnTitle: any;
}
