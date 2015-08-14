# Word-Add-in-Work-with-custom-XML-parts
This sample writes user-specified text to a custom XML part that is bound to a content control within a Word 2013 document.

**Description**

The Add-ins for Office: Work with custom XML parts sample opens a Word file, CustomXMLPartSample.docx, which already contains a custom XML part that has the namespace TimeSummary. When the user enters text in the Client Name text box and chooses Update Client, the cmdUpdateClient_onClick function is executed. This function uses the  getByNamespaceAsync method of the  CustomXmlParts object to identify the custom XML part whose namespace is TimeSummary. The function is passed a callback function, onGotXml, which uses the  getNodesAsync method of the  CustomXmlPart object to get the Client node of the XML part. That function, in turn, is passed another callback function, onGotNode, which uses the  setXmlAsync method of the  CustomXmlNode object to set the text in the content control that is bound to the client node of the Time Summary custom XML part to the new text.

Choosing the Is Part Built In button executes the showXMLPartBuiltIn function, which uses the  getByIdAsync method of the CustomXmlParts object to get the Time Summary part, and then it uses the  builtIn property of the CustomXmlPart object to test whether the part is built-in or custom.

Choosing the Add Part button executes the addCustomXMLPart function, which uses the  addAsync method of the CustomXmlParts object to add a new custom XML part named NewCustomXmlPart to the document. Choosing Delete Part executes the deletePart function, which uses the  deleteAysnc method of the CustomXmlPart object to delete the part.

**Prerequisites**

This sample requires the following:

* Word 2013.
* Visual Studio 2012; apps for Office project template.
* Internet Explorer 9, Internet Explorer 10, or Internet Explorer 11.
* Basic familiarity with JavaScript and HTML.

**Key components**

The Add-ins for Office: Work with custom XML parts sample is created by the CustomXMLApp solution, which contains the following projects and important files:

* The CustomXMLApp project, including the following files:
* CustomXMLApp.xml manifest file
* CustomXMLPartSample.docx file
* The CustomXMLAppWeb project, including the following files:
* Home.html file
* Home.js file

**Configure the sample**

No additional configuration is necessary.

**Build the sample**

Choose the F5 key in Visual Studio to build and deploy the app and open it in Word 2013.

**Run and test the sample**

1 Open the CustomXMLApp.sln file in Visual Studio.

2 Choose the F5 key in Visual Studio to build and deploy the app.

3 In the app task pane, enter text into the Client name box, and then choose Update client name.

4 Choose Is Custom part built in? to determine whether the Time Summary part of the document is built in.

5 Choose Add CustomXML part to add a new part to the document, and choose Delete CustomXML part(s) to delete the part or parts.

You can view a list of the custom XML parts in a document by opening the XML Mapping pane in Word (Developer tab).

**Troubleshooting**

If the app fails to respond as described, try reloading it. (In the task pane, choose the down arrow, and then choose the Reload button.)

**Change log**

* Third release.
* GitHub release: August 2015.

**Related content**

* [JavaScript API for Office](http://msdn.microsoft.com/en-us/library/fp142185(office.15).aspx)

