ExcelDropDown
=============

*ExcelDropDown.cls* provides a versatile drop down functionality for Microsoft Excel worksheets using VBA.


Usage
-----

Note: ExcelDropDown needs activated macros to work!

 Do the following to use ExcelDropDown in your Excel workbook.
 These steps are only required once per workbook.

   1. Define data validation for the desired cells
      (*Data* > *Data Validation*).
       - Choose *List* as validation criteria.
       - Choose a data source. You can also specify a name via "=MyName"
         (e.g. in order to use data on other worksheets). To define and
         and manage names, have a look at *Formula* > *Name Manager* in the
         Excel main window.
       - Uncheck *In-cell dropdown*.
       - You presumably want to deactivate *Input Message* and *Error Alert*.

   2. Activate the *Developer Tab* in the Excel Settings
      (<a href="http://www.addintools.com/documents/excel/how-to-add-developer-tab.html" target="_blank">How-To</a>).

   3. Press *Alt+F11* to run the VBA editor.
       - Add a reference (Menu *Tools* > *References*) to
         "Microsoft Forms 2.0 Object Library".
       - Import *ExcelDropDown.cls* and *ExcelMouseWheelSupport.bas* into the
         VBA project of your Excel workbook (Menu *File* > *Import File...*)
       - Put the following code into *ThisWorkbook*
         (replace occurrences of *Table1* with desired table name)
 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
           Dim WithEvents Table1DropDown As New ExcelDropDown

           Private Sub Workbook_Open()
               Set Table1DropDown = New ExcelDropDown

               ' Set desired options here
               Table1DropDown.ListScrollable = False

               ' At last set the target worksheet
               Table1DropDown.Worksheet = Worksheets("Table1")
           End Sub
 ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
       - If you want to support multiple sheets in your workbook, create a
         separate ExcelDropDown object for each sheet like shown above.
       - Save and reopen the workbook.
