export declare class Helper {
    /**
    * Global Variable of the Menu
    *
    */
    /**
    * Create the HTML dom of the menu - Only 3 level of menu
    *
    * @param {string} menuItem - represents a item of the menu (this.MenuItem)
    */
    ViewModel: ViewModel;
    constructor();
    printTreeOLD(menuItem: MenuItem): any;
    /**
    * Create the HTML dom of the menu - Only 3 level of menu
    *
    * @param {string} menuItem - represents a item of the menu (this.MenuItem)
    * @param {int} nLevel - Level of the menu: 0, 1 or 2
    */
    printTree(menuItem: MenuItem, nLevel: any): void;
    /**
    * Build the tree
    *
    */
    private buildTree();
    /**
    * Loop on the Array globalMenuItems to create a Tree
    *
    */
    private treeFromArray(list, idAttr, parentAttr, childrenAttr);
    private treeFromArray2(list);
    /**
     * Sort children array of a term tree by a sort order
     *
     * @param {obj} tree The term tree
     * @return {obj} A sorted term tree
     */
    private sortTermsFromTree(tree);
    /**
    * Generate the html dom of the navigation
    *
    */
    private displayMenu();
    private applyNavigationEvent();
    /**
    * Return a string name of the term's parent and null if no parent
    *
    * @param {string} sPath - Path of the term: "Level1;SubLevel1-1;Term"
    * @param {bool} isRoot - true if the term is root
    */
    private getParentTerm(sPath, isRoot);
    /**
     * Load data from sessionStorage
     *
     */
    private loadSessionStorage();
    /**
     * Initialisation to get information from the TermStore
     *
     * @param {string} termStoreName - Name of the term store service : 'Service de métadonnées gérées'
     * @param {object} callback - Callback function to call upon completion and pass termset into
     */
    private init;
}
/**
  * Represent an item of the menu
  *
  */
export declare class MenuItem {
    name: any;
    _id: any;
    title: any;
    url: any;
    customSortOrder: any;
    isRoot: any;
    children: any;
    customSortOrde: any;
    sonOf: any;
    guid: any;
    description: any;
    constructor(title: any, url: any, isRoot: any, customSortOrder: any, sonOf: any, guid: any, description: any);
}
export declare class ViewModel {
    globalMenuItems: Array<any>;
    treeNav: Array<any>;
    htmlStr: any;
    isCentralResource: boolean;
    siteUrl: any;
    homeBtnTitle: any;
}
