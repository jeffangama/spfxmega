export declare class MISMain {
    listGUID: any;
    navigationType: any;
    termStoreGUID: any;
    termStoreName: any;
    homeBtnTitle: any;
    homeBtnUrl: any;
    isCentralResource: any;
    spcontext: any;
    constructor(_spcontext: any);
    private getMenuParameters();
    /**
     * Load data from List GlobalNavigation URL and update the sessionStorage
     *
     */
    private getSettingsData(url);
    /**
     * Load data from sessionStorage
     *
     */
    private loadSessionStorage();
    private insertMenu();
}
