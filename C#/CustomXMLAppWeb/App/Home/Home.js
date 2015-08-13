/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
/// <reference path="../App.js" />

(function () {
    "use strict";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            $('#update-client').click(updateClient);
            $('#show-is-built-in').click(showXMLPartBuiltIn);
            $('#add-part').click(addCustomXMLPart);
            $('#delete-part').click(deleteCustomXMLPart);

        });
    };


    function updateClient() {
        // Get Custom XML Part object
        Office.context.document.customXmlParts.getByNamespaceAsync("TimeSummary", onGotXml);
    }


    function onGotXml(asyncResult) {

        // Get the client-level node
        if (asyncResult.value.length > 0)
            asyncResult.value[0].getNodesAsync("/ns0:TimeSummary[1]/ns0:Client[1]", onGotNode);
        else
            app.showNotification('CustomXML part TimeSummary is not found');

    }


    function onGotNode(asyncResult) {

        var clientName = document.getElementById("txtClientName").value;

        // Set client-node XML
        asyncResult.value[0].setXmlAsync("<Client xmlns='TimeSummary'>" + " " + clientName + "</Client>");

    }

    function addCustomXMLPart() {
        Office.context.document.customXmlParts.addAsync("<book xmlns='NewCustomXmlPart'>" +
            "<page number='1'>Hello</page>" +
            "<page number='2'>World!</page></book>", onAddPart)
    }

    function onAddPart(asyncResult) {
        if (asyncResult.result)
            app.showNotification(asyncResult.result.value);
    }

    function deleteCustomXMLPart() {
        Office.context.document.customXmlParts.getByNamespaceAsync("NewCustomXmlPart", deletePart);

    }

    function deletePart(asyncResult) {
        if (asyncResult.value.length > 0) {
            for (var i = 0; i < asyncResult.value.length; i++) {
                asyncResult.value[i].deleteAsync();
            }
        }
    }

    function showXMLPartBuiltIn() {
        Office.context.document.customXmlParts.getByIdAsync("{B5BE1450-13E1-4511-B22C-03F95E715B17}", function (result) {

            if (result.value != null) {
                var xmlPart = result.value;
                if (xmlPart.builtIn)
                    app.showNotification('CustomXML part is built in!');
                else
                    app.showNotification('CustomXML part is not built in');
            }
            else {
                app.showNotification('CustomXML part is not found in the document');
            }
        });
    }

})();
// *********************************************************
//
// Word-Add-in-Work-with-custom-XML-parts, https://github.com/OfficeDev/Word-Add-in-Work-with-custom-XML-parts
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************

