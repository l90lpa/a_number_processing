# A-Number Processing (VBA)

`a_number_processing.bas` contains Visual Basic code that can be used as a macro in Excel for replacing A-numbers with UIDs. 

## Usage:

Assuming that the Developer tab has been enabled. A workflow for processing Excel documents is to:
1. Copy the target Excel document you wish to process.
1. In Excel, import the Visual Basic code (a_number_processing.vba) through the Visial Basic Editor:
    - To open the Visual Basic Editor: `Developer >> Visual Basic`
    - To import the file from the Visual Basic Editor: `File >> Import File`
1. For each sheet you wish to process, run the macro, either through the Visual Basic Editor's "run" button. Or, back in the Excel document through the `Developer >> Macros`.

If the Excel document that imported the macro remains open whilst other Excel documents are open, then the macro will be available to these other Excel documents.

Furthermore, suppose one created a blank Excel document with macros enabled, imported the relevant macros (as discussed above), and then save the document. Then, then workflow above could be streamlined, by simply opening the Excel document containing the macros, alongside any Excel documents to be processed (and using the `Developer >> Macros` option to select and run the macro).

## Design:

To be able be able to consistently replace A-numbers across multiple files, the program loads on start-up, and saves on successful termination, a dictionary from A-numbers to UID which it updates during running. The location and name of this file can be provided at program invokation. If the file is not present, then a program creates a new empty dictionary at start-up. This file is saved in a simple, custom, human-readable format, where each line of the text file represents a key-value pair that are separated by a colon.

Currently the serialization path is hardcoded but could easily be turned into a parameter of the macro.