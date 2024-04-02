# A-Number Processing (VBA)

`a_number_processing.bas` contains Visual Basic code that can be used as a macro in Excel for replacing A-numbers with UIDs. 

## Usage:

### Enabling Developer Tab

See: https://support.microsoft.com/en-us/office/show-the-developer-tab-e1192344-5e56-4d45-931b-e5fd9bea2d45

### Enabling the Macro

Assuming that the Developer tab has been enabled. To enable the macro:
1. In Excel, open the Visual Basic Editor:
    - go to the `Developer` tab.
    - click `Visual Basic` to open the editor.
1. In the Visual Basic Editor, import the file containing the macro:
    - go to `File >> Import File`.
    - select the file with the macros and click `Open`.
1. In the Visual Basic Editor, enable RegExp and Scripting packages: 
    - go to `Tools >> References...`.
    - tick the box for "Microsoft Scripting Runtime" and "Microsoft VBScript Regular Expressions 5.5" then click `OK`.

### Workflows

Assuming that the macro has been imported, or is otherwise available (see below for more discussion on this) Excel documents can be processed as follows:
1. Copy the target Excel document you wish to process.
1. Run the macro on the copied Excel document. There are atleast two ways in which the macro can be run:
        1. In the Visual Basic Editor, click the "run" button (a green triangle), select the macro, and click `Run`.
        2. In the Excel document, go to the `Developer` click `Macros`, select the macro and click `Run`.

If the Excel document that imported the macro remains open whilst other Excel documents are open, then the macro will be available to these other Excel documents.

Furthermore, suppose one created a blank Excel document with macros enabled, imported the relevant macros (as discussed above), and then saved the document. One could then simply open the Excel document that only contains the macros, alongside any Excel documents to be processed, and the macro would be available for use as discribed above.

## Design:

To be able be able to consistently replace A-numbers across multiple files, the program loads on start-up, and saves on successful termination, a dictionary from A-numbers to UID which it updates during running. The location and name of this file can be provided at program invokation. If the file is not present, then a program creates a new empty dictionary at start-up. This file is saved in a simple, custom, human-readable format, where each line of the text file represents a key-value pair that are separated by a colon.

Currently the serialization path is hardcoded but could easily be turned into a parameter of the macro.
