======================================================
Flynet Viewer Server Macro Examples Using Open XML
======================================================
This is a collection of examples of how to us Open XML with the Flynet Viewer Server Macro API to merge data from a mainframe host screen into
a Word document (.docx) or Excel spreadsheet (.xlsx).

These examples use the [Open XML SDK] (https://github.com/OfficeDev/Open-XML-SDK) and were built and tested with version 2.6.1.

The code for these examples was written and compiled using VS 2015 but should work with earlier versions.

These examples require version 2013JP (5.0.232 in the Admin Console) of [Flynet Viewer] (http://www.flynetviewer.com/).

The online help for Flynet Viewer server macros can be found [here] (http://flynetsoftware.com/help/html/servermacros/).

======================================================
Building the Examples
======================================================
Ensure you have a suitable version of Flynet Viewer installed.

Clone and build the [Open XML SDK] (https://github.com/OfficeDev/Open-XML-SDK). Copy DocumentFormat.OpenXml.dll and System.IO.Packaging.dll into the
Lib folder of the example.

Open FVTermSvrMacros.sln in VS and compile it, there is a post build event which will copy 3 dlls to C:\ProgramData\Inventu\FlowMacros\Insure\__Public\ServerMacros
so that they can be found by FVTerm.

If the copy step fails, 1st check that the folder exists and that you have permissions to write to it. If the copy has worked
in the past, then it may be that the dlls are in use by IIS. To resolve this restart IIS by opening a cmd prompt as Admin and typing iisreset.

======================================================
The Examples
======================================================
There are 3 macros in these examples:

1. LoginToAccountSummary - Can be run on the Start screen, log the user in and navigates to the Account Summary screen.

2. OpenXmlWordMergeMacro - Can be run on the Account Summary screen, merges data from Account Summary into a Word document (.docx).

3. OpenXmlExcelMergeMacro - Can be run on the Account Transactions screen (F11 on the Account Summary screen), merges data from the multi-row into an Excel
   document (.xlsx).

======================================================
Setup and Running the Examples
======================================================
The Files directory contains two files required by all the examples:
* Insure Script.xml - Copy to C:\Program Files\flynet\viewer\SimHostScripts if you don't already have it.
* insurerecog.xml - Copy to C:\Program Files\flynet\viewer\Definitions.

1.  Add an application to the FVTerm web.config file (in C:\inetpub\wwwroot\FVTerm by default). Take the next available
    application number and reference the insurerecog.xml file as in this example:

      &lt;add key="Application1" value="InsureRecog;insurerecog.xml;" /&gt;

2.  Use the Flynet Taskbar Control (run as Admin) to ensure you have the Simulated Host running and the Insure Script.xml
    loaded in to the simulator.
3.  Use the Flynet Admin Console to check that you have a host defined to connect to the simulator. A suitable host (called
    Insure) is created during install but it may have been removed or modified. You can use FVTerm to confirm that the host
    is configured correctly.
5.  Open FVTerm in a browser with http://localhost/fvterm/?Application=InsureRecog passing the Application parameter on the URL allows FVTerm to loaded
    the correct recognition file.
6.  Click on the Macro icon in the FVTerm toolbar (to the left of the Host Keys icon).
7.  Run the LoginToAccountSummary macro click Yes at the prompt and note how you are logged in and moved to the Account Summary screen.
8.  Click on the Macro icon again.
9.  Run the OpenXmlWordMergeMacro, Browse to wherever you downloaded the example and, from the Files folder, choose Word.docx, click Open and then OK.
10. A save dialog will open, save the file and compare it to the original, notice how the merge fields in the original have been replaced by data from
    the screen.
11. Hit F11 or click on the button at the bottom of the screen to move to the Account Transactions screen.
12. Click on the macros icon again.
13. Run the OpenXmlExcelMergeMacro, Browse to wherever you downloaded the example and, from the Files folder, choose Excel.xlsx, click Open and then OK.
14. A save dialog will open, save the file and compare it to the original, notice that the 2nd and subsequent rows have been overwritten with data from
    the multi-row. Notice also that the macro paged down to get all available rows and then paged back to leave the screen where it was.

======================================================
The Code
======================================================
All of the C# code for the examples can be found in the OpenXmlSvrMacros project.

* LoginToAccountSummaryMacro.cs - class for the macro that logs the user in and navigates to the Account Summary screen.
* MoreIndicator.cs - class that encapsulates the functionality of the more indicator on a multi-row.
* MultiRow.cs - class that encapsulates a multi-row.
* NumberingFormats - class that encapsulates the functionalty required to interact with number format in Excel via Open XML.
* OpenXmlExcelMergeMacro - class for the macro that populates an Excel spreadsheet from a multi-row.
* OpenXmlWordMergeMacro - class for the macro that merges data from a screen into a Word template.
* ScreenToExcelColumnConversion - class that encapsulates the mapping between a column in a multi-row and a column in Excel.
* SharedStrings - class that encapsulates the functionalty required to interact with shared strings in Word via Open XML.

======================================================
Working with merge fields in Word
======================================================
* ALT-F9 - will toggle the expanded display of merge fields on and off.
* CTRL-F9 - will insert a new merge field into the document.
* F9 - will refresh the fields in the document.

======================================================
Other Projects
======================================================
If you wish to develop more complex Open XML macros, you may find these other projects useful:

* [Open XML Power Tools] (https://github.com/OfficeDev/Open-Xml-PowerTools)
* [ClosedXML] (https://github.com/closedxml/closedxml)