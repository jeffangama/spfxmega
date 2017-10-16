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

//# sourceMappingURL=MISGlobalNavigation.js.map
