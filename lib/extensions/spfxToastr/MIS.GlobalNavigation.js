"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var Helper = (function () {
    function Helper() {
        /**
         * Initialisation to get information from the TermStore
         *
         * @param {string} termStoreName - Name of the term store service : 'Service de métadonnées gérées'
         * @param {object} callback - Callback function to call upon completion and pass termset into
         */
        this.init = function (termStoreName, termSetId) {
            //Load sessionStorage
            this.loadSessionStorage();
            //Load all terms  if not already done.
            if (this.globalMenuItems.length <= 0) {
                // SP.SOD.executeOrDelayUntilScriptLoaded(function () {
                //   'use strict';
                //   var nid = SP.UI.Notify.addNotification("<img src='/_layouts/15/images/loadingcirclests16.gif?rev=23' style='vertical-align:bottom; display:inline-block; margin-" + (document.documentElement.dir == "rtl" ? "left" : "right") + ":2px;' />&nbsp;<span style='vertical-align:top;'>Loading navigation...</span>", false);
                //   SP.SOD.registerSod('sp.taxonomy.js', SP.Utilities.Utility.getLayoutsPageUrl('sp.taxonomy.js'));
                //   SP.SOD.executeFunc('sp.taxonomy.js', false, Function.createDelegate(this, function () {
                //     var context = SP.ClientContext.get_current();
                //     var taxonomySession = SP.Taxonomy.TaxonomySession.getTaxonomySession(context);
                //     var termStore = taxonomySession.get_termStores().getByName(termStoreName);
                //     var termSet = termStore.getTermSet(termSetId);
                //     var terms = termSet.getAllTerms();
                //     context.load(terms);
                //     context.executeQueryAsync(Function.createDelegate(this, function (sender, args) {
                //       var termsEnumerator = terms.getEnumerator();
                //       while (termsEnumerator.moveNext()) {
                //         var currentTerm = termsEnumerator.get_current();
                //         var termIsRoot = currentTerm.get_isRoot();
                //         var termParentName = this.getParentTerm(currentTerm.get_pathOfTerm(), termIsRoot);
                //         this.globalMenuItems.push(
                //           new this.MenuItem(currentTerm.get_name(),//title
                //             currentTerm.get_localCustomProperties()['_Sys_Nav_SimpleLinkUrl'],//url
                //             termIsRoot,//currentTerm.get_isRoot(),//isRoot
                //             currentTerm.get_customSortOrder(), //customSortOrder
                //             termParentName,//currentTerm.get_localCustomProperties()['sonOf']//sonOf
                //             currentTerm.get_id().toString(),//guid of the term
                //             currentTerm.get_localCustomProperties()['description']
                //           )
                //         );
                //       }
                //       //Call the logical function to display the menu
                //       this.displayMenu();
                //       //ko.applyBindings(this.viewModel);
                //       SP.UI.Notify.removeNotification(nid);
                //     }), Function.createDelegate(this, function (sender, args) {
                //       console.log('The following error has occured while loading global navigation: ' + args.get_message());
                //     }));
                //   }));
                // }, 'core.js');
            }
            else {
                this.displayMenu();
            }
        };
        this.ViewModel = new ViewModel();
    }
    Helper.prototype.printTreeOLD = function (menuItem) {
        this.ViewModel.htmlStr = "<ul id='mis-nav' class='mis-hidden'>";
        var nbChild = menuItem.children.length;
        if (nbChild > 0) {
            this.ViewModel.htmlStr + "<li class='has-sub'>";
            this.ViewModel.htmlStr + "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
            this.ViewModel.htmlStr + "<ul>";
            for (var i = 0; i < nbChild; i++) {
                this.printTree(menuItem.children[i], this.ViewModel.htmlStr);
            }
            this.ViewModel.htmlStr + "</ul>";
        }
        else {
            this.ViewModel.htmlStr + "<li>";
            this.ViewModel.htmlStr + "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
        }
        this.ViewModel.htmlStr + "</li>";
        console.log(this.ViewModel.htmlStr);
        return this.ViewModel.htmlStr;
    };
    /**
    * Create the HTML dom of the menu - Only 3 level of menu
    *
    * @param {string} menuItem - represents a item of the menu (this.MenuItem)
    * @param {int} nLevel - Level of the menu: 0, 1 or 2
    */
    Helper.prototype.printTree = function (menuItem, nLevel) {
        this.ViewModel.htmlStr = "<ul id='mis-nav' class='mis-hidden'>";
        var nbChild = menuItem.children.length;
        switch (nLevel) {
            case 0://Top Level => Executive, Core Processes, Support or Social
                this.ViewModel.htmlStr += "<div class='separator-megamenu'></div>";
                this.ViewModel.htmlStr += "<div class='container-megamenu'>";
                //ViewModel.htmlStr + "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
                this.ViewModel.htmlStr += "<p class='menu-level1'>" + menuItem.name;
                this.ViewModel.htmlStr += "<img class='icon-level1' align='right' src='" + this.ViewModel.siteUrl + "/Style Library/MIS.GlobalNavigation/images/icon-" + menuItem.name.replace(/[~#%&*:<>?/\{|}.]/g, "-") + ".png'></img>";
                this.ViewModel.htmlStr += "</p>";
                this.ViewModel.htmlStr += "<ul class='subs' style='list-style: none;'>";
                for (var i = 0; i < nbChild; i++) {
                    this.printTree(menuItem.children[i], nLevel + 1);
                }
                this.ViewModel.htmlStr += "</ul>";
                this.ViewModel.htmlStr += "</div>";
                break;
            case 1:// Header sub Menu
                if (menuItem.url === undefined) {
                    this.ViewModel.htmlStr += "<li><p class='menu-level2'>" + menuItem.name + "</p>";
                }
                else {
                    this.ViewModel.htmlStr += "<li><a href='" + menuItem.url + "' class='mis-tooltip'>" +
                        menuItem.name +
                        "<span>" +
                        "<img class='callout' src='" + this.ViewModel.siteUrl + "/Style Library/MIS.GlobalNavigation/images/callout.gif' />" +
                        menuItem.description +
                        "</span>" +
                        "</a>";
                }
                this.ViewModel.htmlStr += "<ul style='list-style: none;'>";
                for (var i = 0; i < nbChild; i++) {
                    this.printTree(menuItem.children[i], nLevel + 1);
                }
                this.ViewModel.htmlStr += "</ul>";
                this.ViewModel.htmlStr += "</li>";
                break;
            case 2:// sub-sub Menu
                this.ViewModel.htmlStr += "<li>";
                if (menuItem.url === undefined) {
                    this.ViewModel.htmlStr += menuItem.name;
                }
                else {
                    this.ViewModel.htmlStr += "<a href='" + menuItem.url + "'>" + menuItem.name + "</a>";
                }
                this.ViewModel.htmlStr += "</li>";
                break;
            //We doesn't go deep in the Navigation because we display only 3 levels of navigation
            default://Should not go in this case
                this.ViewModel.htmlStr += "<!--Level: " + nLevel + " <a href='" + menuItem.url + "'>" + menuItem.name + "</a>-->"; //For the debug
        }
    };
    /**
    * Build the tree
    *
    */
    Helper.prototype.buildTree = function () {
        this.treeFromArray2(this.ViewModel.globalMenuItems); //Load the treeNav
        //console.log(JSON.stringify(this.treeNav));
        for (var i = 0; i < this.ViewModel.treeNav.length; i++) {
            this.printTree(this.ViewModel.treeNav[i], 0);
        }
        //Add the last separator in the menu
        this.ViewModel.htmlStr += "<div class='separator-megamenu'></div>";
        //BackUp viewMenuModel in the sessionStorage
        var viewMenuModel_json = JSON.stringify(this.ViewModel);
        sessionStorage.setItem("viewMenuModel", viewMenuModel_json);
    };
    ;
    /**
    * Loop on the Array globalMenuItems to create a Tree
    *
    */
    Helper.prototype.treeFromArray = function (list, idAttr, parentAttr, childrenAttr) {
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
        this.ViewModel.treeNav = treeList; //Save the tree in the global variable treeNav
    };
    ;
    Helper.prototype.treeFromArray2 = function (list) {
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
        this.ViewModel.treeNav = treeList; //Save the tree in the global variable treeNav
    };
    ;
    /**
     * Sort children array of a term tree by a sort order
     *
     * @param {obj} tree The term tree
     * @return {obj} A sorted term tree
     */
    Helper.prototype.sortTermsFromTree = function (tree) {
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
    Helper.prototype.displayMenu = function () {
        ////var homeBtnTitle, homeBtnUrl, isCentralResource;
        if (!sessionStorage.viewMenuModel) {
            this.buildTree();
        }
        //Add link to Home page
        var menuInnerHtml = "<div id='homeBtnWebcenter'>" +
            "<a href='" + this.ViewModel.siteUrl + "' class='mis-tooltip-topnav s4-notdlg'>" +
            "<p>" + this.ViewModel.homeBtnTitle + "</p>" +
            "<img class='webcenter' src='" + this.ViewModel.siteUrl + "/Style Library/MIS.GlobalNavigation/images/home.png' height='30' alt='Webcenter'>" +
            "<span style='display: none !important'>" +
            "<img class='callout' src='" + this.ViewModel.siteUrl + "/Style Library/MIS.GlobalNavigation/images/callout.gif' />" +
            "Page d'accueil." +
            "</span>" +
            "</a></div>";
        //Add link to Home page
        menuInnerHtml += "<div id='homeBtn' class='s4-notdlg'>" +
            "<img class='HomeBtn' src='" + this.ViewModel.siteUrl + "/Style Library/MIS.GlobalNavigation/images/menu.png' height='30'>" +
            "<p>Menu</p>" +
            "<div id='nav-container'>" +
            this.ViewModel.htmlStr +
            "</div>" +
            "</div>";
        //Insert icon on the left of the SharePoint text (Top left corner) + all menu html
        $('body').append(menuInnerHtml);
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
        this.applyNavigationEvent();
    };
    ;
    //Event on the Menu
    Helper.prototype.applyNavigationEvent = function () {
        //Apply event on the menu
        //#homeBtn:hover  #nav-container
        $("#homeBtn").on('click', function () {
            $(this).toggleClass("mis-on");
            $('#nav-container').toggleClass("mis-on");
            $("nav ul").toggleClass('mis-hidden');
        });
        $("#homeBtn img")
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
    /**
    * Return a string name of the term's parent and null if no parent
    *
    * @param {string} sPath - Path of the term: "Level1;SubLevel1-1;Term"
    * @param {bool} isRoot - true if the term is root
    */
    Helper.prototype.getParentTerm = function (sPath, isRoot) {
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
    Helper.prototype.loadSessionStorage = function () {
        if (typeof (Storage) !== "undefined") {
            // Code for localStorage/sessionStorage.
            if (sessionStorage.viewMenuModel) {
                var viewMenuModel_json = sessionStorage.getItem("viewMenuModel");
                this.ViewModel = JSON.parse(viewMenuModel_json);
            }
            if (sessionStorage.isCentralResource) {
                this.ViewModel.isCentralResource = (sessionStorage.getItem("isCentralResource") == "true") ? true : false;
            }
            if (sessionStorage.homeBtnUrl) {
                if (this.ViewModel.isCentralResource) {
                    this.ViewModel.siteUrl = sessionStorage.getItem("homeBtnUrl");
                }
                else {
                    this.ViewModel.siteUrl = (_spPageContextInfo.siteServerRelativeUrl == "/") ? "" : _spPageContextInfo.siteServerRelativeUrl;
                }
            }
            if (sessionStorage.homeBtnTitle) {
                this.ViewModel.homeBtnTitle = sessionStorage.getItem("homeBtnTitle");
                ;
            }
        }
        else {
            // Sorry! No Web Storage support..
            if (typeof console === "object") {
                console.log("Sorry! No Web Storage support..");
            }
        }
    };
    ;
    return Helper;
}());
exports.Helper = Helper;
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
    function ViewModel() {
    }
    return ViewModel;
}());
exports.ViewModel = ViewModel;

//# sourceMappingURL=MIS.GlobalNavigation.js.map
