"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var sp_http_1 = require("@microsoft/sp-http");
/** Returns items from the Toast list and caches the results */
var ToastService = (function () {
    function ToastService() {
    }
    //***********************
    //Public Methods
    //***********************
    /** Retrieves toasts that should be displayed for the given user*/
    ToastService.getToasts = function (spHttpClient, baseUrl) {
        var _this = this;
        return new Promise(function (resolve, reject) {
            _this.ensureToasts(spHttpClient, baseUrl)
                .then(function (toasts) {
                resolve(toasts);
            }).catch(function (error) {
                reject(error);
            });
        });
    };
    /** Stores the date/time a toast was acknowledged, used to control what shows on the next refresh
     * @param {number} id - The list ID of the toast to acknowledge
    */
    ToastService.acknowledgeToast = function (id) {
        var cachedData = ToastService.retrieveCache();
        // Check if the status already exists, and if so update it
        //  otherwise, add a new status for the id
        var index = ToastService.indexOfToastStatusById(id, cachedData.ToastStatuses);
        if (index >= 0) {
            cachedData.ToastStatuses[index].Ack = new Date();
        }
        else {
            cachedData.ToastStatuses.push({
                Id: id,
                Ack: new Date()
            });
        }
        ToastService.storeCache(cachedData);
    };
    //***********************
    //localStorage Management
    //***********************
    /** Rehydrates spfxToastr data from localStorage (or creates a new empty set) */
    ToastService.retrieveCache = function () {
        //Pull data from localStorage if available and we previously cached it
        var cachedData = localStorage ? JSON.parse(localStorage.getItem(ToastService.storageKey)) : undefined;
        if (cachedData) {
            cachedData.Loaded = new Date(cachedData.Loaded); //Rehydrate date from JSON (serializes to string)
        }
        else {
            //Initialize a new, empty object
            cachedData = {
                Toasts: [],
                ToastStatuses: []
            };
        }
        return cachedData;
    };
    /** Serializes spfxToastr data into localStorage */
    ToastService.storeCache = function (cachedData) {
        //Cache the data in localStorage when possible
        if (localStorage) {
            localStorage.setItem(ToastService.storageKey, JSON.stringify(cachedData));
        }
    };
    //***********************
    //Toast Retrieval
    //***********************
    /** Retrieves toasts from either the cache or the list depending on the cache's freshness */
    ToastService.ensureToasts = function (spHttpClient, baseUrl) {
        return new Promise(function (resolve, reject) {
            var cachedData = ToastService.retrieveCache();
            if (cachedData.Loaded) {
                //True Cache found, check if it is stale
                // anything older than 2 minutes will be considered stale
                var now = new Date();
                var staleTime = new Date(now.getTime() + -2 * 60000);
                if (cachedData.Loaded > staleTime && !ToastService.getFromListAlways) {
                    //console.log('Pulled toasts from localStorage');
                    resolve(ToastService.reduceToasts(cachedData));
                    return;
                }
            }
            if (window.spfxToastrLoadingData) {
                //Toasts are already being loaded! Briefly wait and try again
                window.setTimeout(function () {
                    ToastService.ensureToasts(spHttpClient, baseUrl)
                        .then(function (toasts) {
                        resolve(toasts);
                    });
                }, 100);
            }
            else {
                //Set a loading flag to prevent multiple data queries from firing
                //  this will be important should there be multiple consumers of the service on a single page
                window.spfxToastrLoadingData = true;
                //Toasts need to be loaded, so let's go get them!
                ToastService.getToastsFromList(spHttpClient, baseUrl)
                    .then(function (toasts) {
                    //console.log('Pulled toasts from the list');
                    cachedData.Toasts = toasts;
                    cachedData.Loaded = new Date(); //Reset the cache timeout
                    cachedData = ToastService.processCache(cachedData);
                    //Update the cache
                    ToastService.storeCache(cachedData);
                    //Clear the loading flag
                    window.spfxToastrLoadingData = false;
                    //Give them some toast!
                    resolve(ToastService.reduceToasts(cachedData));
                }).catch(function (error) {
                    reject(error);
                });
            }
        });
    };
    /** Pulls the active toast entries directly from the underlying list */
    ToastService.getToastsFromList = function (spHttpClient, baseUrl) {
        //Toasts are only shown during their scheduled window
        var now = new Date().toISOString();
        var filter = "(StartDate le datetime'" + now + "') and (EndDate ge datetime'" + now + "')";
        return spHttpClient.get(baseUrl + "/" + ToastService.apiEndPoint + "?$select=" + ToastService.select + "&$filter=" + filter + "&$orderby=" + ToastService.orderby, sp_http_1.SPHttpClient.configurations.v1)
            .then(function (response) {
            if (!response.ok) {
                //Failed requests don't automatically throw exceptions which
                // can be problematic for chained promises, so we throw one
                throw "Unable to get items: " + response.status + " (" + response.statusText + ")";
            }
            return response.json();
        })
            .then(function (results) {
            //Clean up extra properties
            // Even when your interface only defines certain properties, SP sends many
            // extra properties that you may or may not care about (we don't)
            // (this isn't strictly necessary but makes the cache much cleaner)
            var toasts = [];
            for (var _i = 0, _a = results.value; _i < _a.length; _i++) {
                var v = _a[_i];
                toasts.push({
                    Title: v.Title,
                    Id: v.Id,
                    Severity: v.Severity,
                    Frequency: v.Frequency,
                    Enabled: v.Enabled,
                    Message: v.Message
                });
            }
            return toasts;
        });
    };
    //***********************
    //Helper Functions
    //***********************
    /** Helper function to return the index of an IToastStatus object by the Id property */
    ToastService.indexOfToastStatusById = function (Id, toastStatuses) {
        for (var i = 0; i < toastStatuses.length; i++) {
            if (toastStatuses[i].Id == Id) {
                return i;
            }
        }
        return -1;
    };
    /** Helper function to clean up the toast statuses by removing old toasts */
    ToastService.processCache = function (cachedData) {
        //Setup a temporary array of Ids (makes the filtering easier)
        var activeIds = [];
        for (var _i = 0, _a = cachedData.Toasts; _i < _a.length; _i++) {
            var toast = _a[_i];
            activeIds.push(toast.Id);
        }
        //only keep the status info for toast that still matter (active)
        cachedData.ToastStatuses = cachedData.ToastStatuses.filter(function (value) {
            return activeIds.indexOf(value.Id) >= 0;
        });
        return cachedData;
    };
    /** Adjusts the toasts to display based on what the user has already acknowledged and the toast's frequency value*/
    ToastService.reduceToasts = function (cachedData) {
        return cachedData.Toasts.filter(function (toast) {
            if (!toast.Enabled) {
                //Disabled toasts are still queried so that their status isn't lost
                // however, they shouldn't be displayed
                return false;
            }
            var tsIndex = ToastService.indexOfToastStatusById(toast.Id, cachedData.ToastStatuses);
            if (tsIndex >= 0) {
                var lastShown = new Date(cachedData.ToastStatuses[tsIndex].Ack); //Likely needs to be rehyrdated from JSON
                switch (toast.Frequency) {
                    case 'Once':
                        //Already shown
                        return false;
                    case 'Always':
                        return true;
                    default:
                        //Default behavior is Once Per Day
                        var now = new Date();
                        if (now.getFullYear() !== lastShown.getFullYear()
                            || now.getMonth() !== lastShown.getMonth()
                            || now.getDay() !== lastShown.getDay()) {
                            //Last shown on a different day, so show it!
                            return true;
                        }
                        else {
                            //Already shown today
                            return false;
                        }
                }
            }
            else {
                //No previous status means it needs to be shown
                return true;
            }
        });
    };
    ToastService.storageKey = 'spfxToastr'; //Key used for localStorage
    ToastService.getFromListAlways = false; //Useful for testing
    //Breaking up the URL like this isn't necessary, but can be easier to update
    ToastService.apiEndPoint = "_api/web/lists/getbytitle('Toast')/items";
    ToastService.select = "Id,Title,Severity,Frequency,Enabled,Message";
    ToastService.orderby = "StartDate asc";
    return ToastService;
}());
exports.ToastService = ToastService;

//# sourceMappingURL=ToastService.js.map
