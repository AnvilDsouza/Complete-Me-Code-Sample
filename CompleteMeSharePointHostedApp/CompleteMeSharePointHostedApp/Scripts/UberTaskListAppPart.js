'use strict';

// Set the style of the client web part page to be consistent with the host web.
(function () {
    var hostUrl = '';
    if (document.URL.indexOf('?') != -1) {
        var params = document.URL.split('?')[1].split('&');
        for (var i = 0; i < params.length; i++) {
            var p = decodeURIComponent(params[i]);
            if (/^SPHostUrl=/i.test(p)) {
                hostUrl = p.split('=')[1];
                document.write('<link rel="stylesheet" href="' + hostUrl + '/_layouts/15/defaultcss.ashx" />');
                break;
            }
        }
    }
    if (hostUrl == '') {
        document.write('<link rel="stylesheet" href="/_layouts/15/1033/styles/themable/corev15.css" />');
    }
})();


var context;
var hostweburl;
var appweburl;
var web;
var factory;
var appContextSite;

// Load the required SharePoint libraries
$(document).ready(function () {
    //Get the URI decoded URLs.
    hostweburl =
        decodeURIComponent(
            getQueryStringParameter("SPHostUrl")
    );
    appweburl =
        decodeURIComponent(
            getQueryStringParameter("SPAppWebUrl")
    );

    // resources are in URLs in the form:
    // web_url/_layouts/15/resource
    var scriptbase = hostweburl + "/_layouts/15/";

    // Load the js files and continue to the successHandler
    $.getScript(scriptbase + "SP.Runtime.js",
        function () {
            $.getScript(scriptbase + "SP.js",
                function () { $.getScript(scriptbase + "SP.RequestExecutor.js", execCrossDomainRequest); }
                );
        }
     );
});

function execCrossDomainRequest() {
    context = new SP.ClientContext(appweburl);
    factory = new SP.ProxyWebRequestExecutorFactory(appweburl);
    context.set_webRequestExecutorFactory(factory);
    appContextSite = new SP.AppContextSite(context, hostweburl);

    web = appContextSite.get_web();
    context.load(web);

    var selectedListTitle = web.get_lists().getByTitle('Tasks');
    var camlQuery = SP.CamlQuery.createAllItemsQuery();
    this.listItemCollection = selectedListTitle.getItems(camlQuery);
    context.load(this.listItemCollection);

    context.executeQueryAsync(
        Function.createDelegate(this, function () {
            console.info("HostwebTitle: " + web.get_title());
            var ListEnumerator = this.listItemCollection.getEnumerator();
            while (ListEnumerator.moveNext()) {
                var currentItem = ListEnumerator.get_current();
                $('#message').text(currentItem.get_item('Title') + currentItem.get_item('ID'));
            }

        }),
        Function.createDelegate(this, function () {
            console.error('Houston we have a problem!');
        })
    );
}

// Function to retrieve a query string value.
// For production purposes you may want to use
// a library to handle the query string.
function getQueryStringParameter(paramToRetrieve) {
    var params =
        document.URL.split("?")[1].split("&");
    var strParams = "";
    for (var i = 0; i < params.length; i = i + 1) {
        var singleParam = params[i].split("=");
        if (singleParam[0] == paramToRetrieve)
            return singleParam[1];
    }
}