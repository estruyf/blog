///<reference path="typings/sharepoint/SharePoint.d.ts" />
///<reference path="typings/q/q.d.ts" />

/*

Author: Mikael Svenson - Puzzlepart 2016
Twitter: @mikaelsvenson

Description
-----------

Script which hooks into the query execution flow of a page using search web parts to inject custom query variables using JavaScript

The script requires jQuery to be loaded on the page, and then you can just attach this script on any page with script editor web part,
content editor web part, custom action or similar.

*/
"use strict";

interface SynonymValue {
    Title: string;
    Synonym: string;
    DoubleUsage: boolean;
}

import Q = require('q');
declare var Srch;
declare var Sys;
module mAdcOW.Search.VariableInjection {
    var _loading = false;
    var _userDefinedVariables = {};
    var _synonymTable = {};
    var _dataProviders = [];
    var _origExecuteQuery = Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQuery;
    var _origExecuteQueries = Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQueries;
    var _siteUrl: string = _spPageContextInfo.webAbsoluteUrl;

    // Function to load synonyms asynchronous - poor mans synonyms
    function loadSynonyms() {
        var defer = Q.defer();
        
        var urlSynonymsList: string = _siteUrl + "/_api/Web/Lists/getByTitle('Synonyms')/Items?$select=Title,Synonym,DoubleUsage";
        var req: XMLHttpRequest = new XMLHttpRequest();
        req.onreadystatechange = function() {
            if (this.readyState === 4) {
                if (this.status === 200) {
                    let data = JSON.parse(this.response);
                    if (typeof data.value !== 'undefined') {
                        let results: SynonymValue[] = data.value;
                        if (results.length) {
                            for (let i = 0; i < results.length; i++) {
                                let item = results[i];
                                if (item.DoubleUsage) {
                                    let synonyms: string[] = item.Synonym.split(',');
                                    // Set the default synonym
                                    _synonymTable[item.Title.toLowerCase()] = synonyms;
                                    // Loop over the list of synonyms
                                    let tmpSynonyms: string[] = synonyms;
                                    tmpSynonyms.push(item.Title.toLowerCase().trim());
                                    synonyms.forEach(s => {
                                        _synonymTable[s.toLowerCase().trim()] = tmpSynonyms.filter(function (fItem) { return fItem !== s });
                                    });
                                } else {
                                    // Set a single synonym
                                    _synonymTable[item.Title.toLowerCase()] = item.Synonym.split(',');
                                }
                            }
                        }
                    }
                    defer.resolve();
                }
                else if (this.status >= 400) {
                    console.error("getJSON failed, status: " + this.textStatus + ", error: " + this.error);
                    defer.reject(this.statusText);
                }
            }
        }
        req.open('GET', urlSynonymsList, true);
        req.setRequestHeader('Accept', 'application/json');
        req.send();
        return defer.promise;
    }

    // Function to inject synonyms at run-time
    function injectSynonyms(query: string, dataProvider) {
        // Remove complex query parts AND/OR/NOT/ANY/ALL/parenthasis/property queries/exclusions - can probably be improved            
        var cleanQuery: string = query.replace(/(-\w+)|(-"\w+.*?")|(-?\w+[:=<>]+\w+)|(-?\w+[:=<>]+".*?")|((\w+)?\(.*?\))|(AND)|(OR)|(NOT)/g, '');
        var queryParts: string[] = cleanQuery.match(/("[^"]+"|[^"\s]+)/g);
        
        if (queryParts) {
            for (var i = 0; i < queryParts.length; i++) {
                if (_synonymTable[queryParts[i]]) {
                    // Replace the current query part in the query with all the synonyms
                    query = query.replace(queryParts[i], String.format('({0} OR {1})', queryParts[i], _synonymTable[queryParts[i]].join(' OR ')));
                }
            }
        }
        
        // Add a custom action for the synonym query
        dataProvider.get_properties()["SynonymQuery"] = query;
    }

    // Sample function to load user variables asynchronous
    function loadUserVariables() {
        var defer = Q.defer();
        SP.SOD.executeFunc('sp.js', 'SP.ClientContext', () => {
            // Query user hidden list - not accessible via REST
            // If you want TERM guid's you need to mix and match the use of UserProfileManager and TermStore and cache client side
            var urlCurrentUser: string = _siteUrl + "/_vti_bin/listdata.svc/UserInformationList?$filter=Id eq " + _spPageContextInfo.userId;
            
            var req = new XMLHttpRequest();
            req.onreadystatechange = function() {
                if (this.readyState === 4) {
                    if (this.status === 200) {
                        var data = JSON.parse(this.response);
                        var user: SP.User = data['d']['results'][0];
                        for (var property in user) {
                            if (user.hasOwnProperty(property)) {
                                var val = user[property];
                                if (typeof val == "number") {
                                    console.log(property + " : " + val);
                                    _userDefinedVariables["mAdcOWUser." + property] = val;
                                } else if (typeof val == "string") {
                                    console.log(property + " : " + val);
                                    _userDefinedVariables["mAdcOWUser." + property] = val.split(/[\s,]+/);
                                }
                            }
                        }
                        defer.resolve();
                    }
                    else if (this.status >= 400) {
                        console.error("getJSON failed, status: " + this.textStatus + ", error: " + this.error);
                        defer.reject(this.statusText);
                    }
                }
            }
            req.open('GET', urlCurrentUser, true);
            req.setRequestHeader('Accept', 'application/json');
            req.send();
        });
        return defer.promise;
    }

    // Function to inject custom variables on page load
    function injectCustomQueryVariables() {
        var queryGroups = Srch.ScriptApplicationManager.get_current().queryGroups;
        for (var group in queryGroups) {
            if (queryGroups.hasOwnProperty(group)) {
                var dataProvider = queryGroups[group].dataProvider;
                var properties = dataProvider.get_properties();
                // add all user variables fetched and stored as mAdcOWUser.
                for (var prop in _userDefinedVariables) {
                    if (_userDefinedVariables.hasOwnProperty(prop)) {
                        properties[prop] = _userDefinedVariables[prop];
                    }
                }

                // add some custom variables for show
                dataProvider.get_properties()["awesomeness"] = "WOOOOOOT";
                dataProvider.get_properties()["moreawesomeness"] = ["foo", "bar"];

                // set hook for query time variables which can change
                dataProvider.add_queryIssuing((sender, e) => {
                    // code which should modify the current query based on context for each new query
                    injectSynonyms(e.queryState.k, sender);
                });

                _dataProviders.push(dataProvider);
            }
        }
    }

    function loadDataAndSearch() {
        if (!_loading) {
            _loading = true;
            // run all async code needed to pull in data for variables
            Q.all([loadSynonyms()/*, loadUserVariables()*/]).done(() => {
                // set loaded data as custom query variables
                injectCustomQueryVariables();
                
                // reset to original function
                Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQuery = _origExecuteQuery;
                Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQueries = _origExecuteQueries;
                
                // re-issue query for the search web parts
                for (var i = 0; i < _dataProviders.length; i++) {
                    // complete the intercepted event
                    _dataProviders[i].raiseResultReadyEvent(new Srch.ResultEventArgs(_dataProviders[i].get_initialQueryState()));
                    // re-issue query
                    _dataProviders[i].issueQuery();
                }
            });
        }
    }

    // Loader function to hook in client side custom query variables
    function hookCustomQueryVariables() {
        console.log("Hooking variable injection");

        // TODO: Check if we have cached data, if so, no need to intercept for async web parts
        // Override both executeQuery and executeQueries

        Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQuery = (query : Microsoft.SharePoint.Client.Search.Query.Query) => {
            loadDataAndSearch();
            return new SP.JsonObjectResult();
        }

        Microsoft.SharePoint.Client.Search.Query.SearchExecutor.prototype.executeQueries = (queryIds: string[], queries: Microsoft.SharePoint.Client.Search.Query.Query[], handleExceptions: boolean) => {
            loadDataAndSearch();
            return new SP.JsonObjectResult();
        }
    }

    ExecuteOrDelayUntilBodyLoaded(() => {
        Sys.Application.add_init(hookCustomQueryVariables);
    });
}