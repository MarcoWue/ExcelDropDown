VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ExcelDropDown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' **
' ** ExcelDropDown.cls
'/**
'
' @brief        *ExcelDropDown.cls* provides a versatile drop down
'               functionality for Microsoft Excel worksheets using VBA.
'
' @version      1.0.1
'
' @author       MarcoWue
'               (<a href="https://github.com/MarcoWue" target="_blank">GitHub</a>)
'
' @copyright    Copyright (c) 2013 by MarcoWue.                         \n
'               This work is made available under the terms of the
'               Creative Commons Attribution 3.0 Unported License
'               (<a href="http://creativecommons.org/licenses/by/3.0/" target="_blank">CC BY 3.0</a>).
'
'
' Dependencies
' ------------
'
'   - ExcelMouseWheelSupport.bas
'   - Microsoft Forms 2.0 Object Library (Menu *Tools* > *References*)
'
'
' Usage
' -----
'
' @note     ExcelDropDown needs activated macros to work!
'
' Do the following to use ExcelDropDown in your Excel workbook.
' These steps are only required once per workbook.
'
'   1. Define data validation for the desired cells
'      (*Data* > *Data Validation*).
'       - Choose *List* as validation criteria.
'       - Choose a data source. You can also specify a name via "=MyName"
'         (e.g. in order to use data on other worksheets). To define and
'         and manage names, have a look at *Formula* > *Name Manager* in the
'         Excel main window.
'       - Uncheck *In-cell dropdown*.
'       - You presumably want to deactivate *Input Message* and *Error Alert*.
'
'   2. Activate the *Developer Tab* in the Excel Settings
'      (<a href="http://www.addintools.com/documents/excel/how-to-add-developer-tab.html" target="_blank">How-To</a>).
'
'   3. Press *Alt+F11* to run the VBA editor.
'       - Add a reference (Menu *Tools* > *References*) to
'         "Microsoft Forms 2.0 Object Library".
'       - Import *ExcelDropDown.cls* and *ExcelMouseWheelSupport.bas* into the
'         VBA project of your Excel workbook (Menu *File* > *Import File...*)
'       - Put the following code into *ThisWorkbook*
'         (replace occurrences of *Table1* with desired table name)
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'           Dim WithEvents Table1DropDown As New ExcelDropDown
'
'           Private Sub Workbook_Open()
'               Set Table1DropDown = New ExcelDropDown
'
'               '' Set desired options here
'               Table1DropDown.ListScrollable = False
'
'               '' At last set the target worksheet
'               Table1DropDown.Worksheet = Worksheets("Table1")
'           End Sub
' ~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
'       - If you want to support multiple sheets in your workbook, create a
'         separate ExcelDropDown object for each sheet like shown above.
'       - Save and reopen the workbook.
'
'
' Version History
' ---------------
'
'   - 1.0.1 (2013-09-18) *MarcoWue*
'       - Ported to Doxygen code documentation.
'
'   - 1.0.0 (2013-09-03) *MarcoWue*
'       - Initial release.
'
'**/


Option Explicit

Private Const CONTROL_NAME = "ExcelDropDown"
Private Const CONTROL_CLASS = "Forms.ComboBox.1"


Private WithEvents m_Sheet As Worksheet
Attribute m_Sheet.VB_VarHelpID = -1
Private WithEvents m_Book As Workbook
Attribute m_Book.VB_VarHelpID = -1
Private WithEvents m_App As Application
Attribute m_App.VB_VarHelpID = -1
Private m_OLEObject As OLEObject
Private WithEvents m_Control As ComboBox
Attribute m_Control.VB_VarHelpID = -1
Private m_TargetCell As Range

Private m_HookPaused As Boolean


Private m_AutoShow As Boolean
Private m_AutoDropDown As Boolean
Private m_ButtonInsideCell As Boolean

Private m_ListFixed As Boolean
Private m_ListRowCount As Long
Private m_ListScrollable As Boolean

Private m_ItemsFastProcessing As Boolean
Private m_ItemsSort As Boolean
Private m_ItemsRemoveDuplicate As Boolean
Private m_ItemsDuplicateCompare As VbCompareMethod
Private m_ItemsRemoveBlank As Boolean
Private m_ItemsAddFirstBlank As Boolean
Private m_ItemsTrim As Boolean


' Event ValidateDropDown
'
'' Is raised to validate if drop down should be shown in a specific cell.
''
'' @param[in]   Target          The new target cell where the drop down is
''                              going to be shown.
'' @param[out]  Cancel          Set to \c True to prevent the drop down from
''                              being shown.
'
Public Event ValidateDropDown(ByVal Target As Range, ByRef Cancel As Boolean)


' Event BeforeDropDown
'
'' Is raised immediately before drop down is shown.
''
'' If desired you can modify ComboBoxControl appearance or content here.
''
'' @param[in]   Target          The new target cell where the drop down is
''                              going to be shown.
'' @param[out]  Cancel          Set to \c True to prevent the drop down from
''                              being shown.
'
Public Event BeforeDropDown(ByVal Target As Range, ByRef Cancel As Boolean)


' Event AfterDropDown
'
'' Is raised immediately after drop down is shown.
''
'' @param[in]   Target          The cell where the drop down is currently
''                              shown at.
'
Public Event AfterDropDown(ByVal Target As Range)


' Event ExitDropDown
'
'' Is raised immediately after drop down is hidden.
''
'' @param[in]   Target          The cell where the drop down was shown at.
'
Public Event ExitDropDown(ByVal Target As Range)


' Property Worksheet
'
'' The worksheet ExcelDropDown is assigned to.
''
'' Setting this property causes initialization of the ComboBox control.     \n
'' To handle multiple worksheets, create a separate ExcelDropDown object
'' for each desired sheet.
'
Public Property Get Worksheet() As Worksheet
    Worksheet = m_Sheet
End Property
Public Property Let Worksheet(ByVal NewValue As Worksheet)
    On Error Resume Next
    If Not m_OLEObject Is Nothing Then
        m_Book.OLEObjects.Delete m_OLEObject
        Set m_OLEObject = Nothing
        Set m_Control = Nothing
    End If
    Set m_Sheet = NewValue
    Init
End Property


' Property ComboBoxControl
'
'' The ComboBox control ExcelDropDown uses (**read-only**).
''
'' You may want to change its appearance or content utilizing the
'' BeforeDropDown event.                                                    \n
'' Is \c Nothing before property Worksheet is set.
''
Public Property Get ComboBoxControl() As ComboBox
    Set ComboBoxControl = m_Control
End Property


' Property CurrentCell
'
'' The cell the drop down is currently shown in (**read-only**).
''
'' Is \c Nothing when no drop down is displayed at the moment.
'
Public Property Get CurrentCell() As Range
    Set CurrentCell = m_TargetCell
End Property


' Property AutoShow
'
'' Determines whether the drop down should be shown automatically.
''
'' If \c True, the drop down box is shown automatically whenever an
'' appropriate cell is selected in the worksheet.
'
Public Property Get AutoShow() As Boolean
    AutoShow = m_AutoShow
End Property
Public Property Let AutoShow(ByVal NewValue As Boolean)
    m_AutoShow = NewValue
End Property


' Property AutoDropDown
'
'' Determines whether the drop down should expanded automatically.
''
'' If \c True, it is is expanded automatically every time the drop down is
'' shown.
'
Public Property Get AutoDropDown() As Boolean
    AutoDropDown = m_AutoDropDown
End Property
Public Property Let AutoDropDown(ByVal NewValue As Boolean)
    m_AutoDropDown = NewValue
End Property


' Property ButtonInsideCell
'
'' Determines the drop down button position.
''
'' If \c True, the drop down button is shown inside the target cell.        \n
'' If \c False, it is shown besides of the cell.
'
Public Property Get ButtonInsideCell() As Boolean
    ButtonInsideCell = m_ButtonInsideCell
End Property
Public Property Let ButtonInsideCell(ByVal NewValue As Boolean)
    m_ButtonInsideCell = NewValue
End Property


' Property ListFixed
'
'' Determines the drop down behaviour.
''
'' If \c True, the user can only choose between the shown list items and is
'' not able to enter text into the drop down field.
'
Public Property Get ListFixed() As Boolean
    ListFixed = m_ListFixed
End Property
Public Property Let ListFixed(ByVal NewValue As Boolean)
    m_ListFixed = NewValue
End Property


' Property ListRowCount
'
'' Determines the length of the shown drop down list.
''
'' Is expressed as number of rows.
'
Public Property Get ListRowCount() As Long
    ListRowCount = m_ListRowCount
End Property
Public Property Let ListRowCount(ByVal NewValue As Long)
    m_ListRowCount = NewValue
End Property


' Property ListScrollable
'
'' Determines mouse wheel scrolling support of the drop down list.
''
'' If \c True, mouse wheel support is enabled by utilizing a so called hook
'' callback function which captures mouse activity on a low level basis.
''
'' @note    If you are encountering strange behaviour set this option to
''          \c False.
'
Public Property Get ListScrollable() As Boolean
    ListScrollable = m_ListScrollable
End Property
Public Property Let ListScrollable(ByVal NewValue As Boolean)
    If m_ListScrollable And Not NewValue Then _
        StopMouseWheelHook m_Control
    m_ListScrollable = NewValue
End Property


' Property ItemsFastProcessing
'
'' Determines if advanced list item processing is active.
''
'' If \c True, the drop down list items are displayed solely by faster
'' internal Excel routines utilizing the ListFillRange property of the
'' ComboBox control.                                                        \n
'' If \c False, additional processing of the items is done which may takes
'' some time with larger data amounts.
''
'' @note    Please set to \c False if you want to customize list items in
''          the BeforeDropDown event or if you have set one of the following
''          properties to \c True:
''              - ItemsSort
''              - ItemsRemoveDuplicate
''              - ItemsRemoveBlank
''              - ItemsAddFirstBlank
'
Public Property Get ItemsFastProcessing() As Boolean
    ItemsFastProcessing = m_ItemsFastProcessing
End Property
Public Property Let ItemsFastProcessing(ByVal NewValue As Boolean)
    m_ItemsFastProcessing = NewValue
End Property


' Property ItemsSort
'
'' Determines whether list items should be sorted alphabetically.
'
Public Property Get ItemsSort() As Boolean
    ItemsSort = m_ItemsFastProcessing
End Property
Public Property Let ItemsSort(ByVal NewValue As Boolean)
    m_ItemsSort = NewValue
End Property


' Property ItemsRemoveDuplicate
'
'' Determines whether duplicate list items should not be shown.
'
Public Property Get ItemsRemoveDuplicate() As Boolean
    ItemsRemoveDuplicate = m_ItemsRemoveDuplicate
End Property
Public Property Let ItemsRemoveDuplicate(ByVal NewValue As Boolean)
    m_ItemsRemoveDuplicate = NewValue
End Property


' Property ItemsDuplicateCompare
'
'' Determines whether ItemsRemoveDuplicate should be case sensitive.
'
Public Property Get ItemsDuplicateCompare() As VbCompareMethod
    ItemsDuplicateCompare = m_ItemsDuplicateCompare
End Property
Public Property Let ItemsDuplicateCompare( _
        ByVal NewValue As VbCompareMethod)
    m_ItemsDuplicateCompare = NewValue
End Property


' Property ItemsRemoveBlank
'
'' Determines whether empty list items should not be shown.
'
Public Property Get ItemsRemoveBlank() As Boolean
    ItemsRemoveBlank = m_ItemsRemoveBlank
End Property
Public Property Let ItemsRemoveBlank(ByVal NewValue As Boolean)
    m_ItemsRemoveBlank = NewValue
End Property


' Property ItemsAddFirstBlank
'
'' Determines whether a first blank item should be added to the drop down list.
''
'' If \c True, a first blank item will be added to the drop down list
'' regardless of the ItemsRemoveBlank property.
'
Public Property Get ItemsAddFirstBlank() As Boolean
    ItemsAddFirstBlank = m_ItemsAddFirstBlank
End Property
Public Property Let ItemsAddFirstBlank(ByVal NewValue As Boolean)
    m_ItemsAddFirstBlank = NewValue
End Property


' Property ItemsTrim
'
'' Determines whether leading and trailing spaces should be removed from items.
'
Public Property Get ItemsTrim() As Boolean
    ItemsTrim = m_ItemsTrim
End Property
Public Property Let ItemsTrim(ByVal NewValue As Boolean)
    m_ItemsTrim = NewValue
End Property


' Method Show()
'
'' Shows the drop down at the specified target cell.
''
'' This method usually has not to be called if AutoShow is set to \c True.  \n
'' Call this method to show the drop down manually. Please note that the
'' target cell must always have data validation enabled with a valid 'List'
'' criteria.
''
'' @param[in]   Target          The cell the drop down should be shown at.
'' @param[in]   NoAutoDropDown  If \c True, the AutoDropDown property will be
''                              ignored
''
'' @return  \c True if the drop down could be shown successfully, \c False
''          otherwise.
'
Public Function Show(ByVal Target As Range, _
            Optional ByVal NoAutoDropDown As Boolean = False) As Boolean
            
    Dim Cancel As Boolean
    Dim s As String
    Dim List() As String
    Dim i As Long
    Dim j As Long
    
    On Error GoTo EndSub
    
    Show = False
    If m_Control Is Nothing Then _
        Exit Function
    
    With m_Control
        ' Hide drop down on old position
        Hide
        
        ' Only allow single selected cells with appropriate data validation
        ' (unfortunately throws an error when no data validation is set)
        If (Target.Cells.Count = 1) And _
                (Target.Validation.Type = xlValidateList) Then
   
            ' Raise ValidateDropDown event
            Cancel = False
            RaiseEvent ValidateDropDown(Target, Cancel)
            If Cancel Then _
                Exit Function
            
            ' Retrieve data list reference and remove the leading '='
            s = Target.Validation.Formula1
            s = Right$(s, Len(s) - 1)
            
            On Error GoTo Err
            Application.EnableEvents = False
            
            ' Fill drop down list with data
            .ListFillRange = Range(s).Address(RowAbsolute:=True, _
                    ColumnAbsolute:=True, External:=True)

            ' Rewrite data to list to allow content modification for disabled
            ' ItemsFastProcessing option
            If Not m_ItemsFastProcessing Then
            
                ' Catch list items and trim if desired
                ReDim List(.ListCount)
                For i = 0 To .ListCount - 1
                    If m_ItemsTrim Then
                        List(i) = Trim$(.List(i))
                    Else
                        List(i) = .List(i)
                    End If
                Next i
                
                ' Disconnect data source
                .ListFillRange = ""
                .Clear
                
                ' Sort list if desired
                If m_ItemsSort Then _
                    QuicksortStrings List
                
                ' Add first blank item if desired
                ' (do not add when first line is already blank)
                If m_ItemsAddFirstBlank Then
                    Cancel = False
                    If Not m_ItemsRemoveBlank And (UBound(List) >= 0) Then _
                        Cancel = (Len(Trim$(List(0))) = 0)
                    If Not Cancel Then _
                        .AddItem ""
                End If
                
                ' Write back list items
                For i = LBound(List) To UBound(List)
                    s = List(i)
                    
                    ' Skip duplicates if desired
                    If m_ItemsRemoveDuplicate And i > 0 Then
                    
                        ' Ignore blank item if a first blank item was already
                        ' added
                        If m_ItemsAddFirstBlank Then
                            If (Len(Trim$(s)) = 0) Then _
                                GoTo NextItem
                        End If
                        
                        If m_ItemsSort Then
                            ' List is already sorted so we can identify
                            ' duplicates quickly
                            If StrComp(s, List(i - 1), _
                                    m_ItemsDuplicateCompare) = 0 Then _
                                GoTo NextItem
                        Else
                            ' Iterate through all previous items to identify
                            ' duplicates
                            For j = LBound(List) To i - 1
                                If StrComp(s, List(j), _
                                        m_ItemsDuplicateCompare) = 0 Then _
                                    GoTo NextItem
                            Next j
                        End If
                        
                    End If
                    
                    ' Ignore blank items if desired
                    If m_ItemsRemoveBlank Then
                        If Len(Trim$(s)) = 0 Then _
                            GoTo NextItem
                    End If
                    
                    .AddItem s
NextItem:
                Next i
            End If
            
            ' Set drop down text to target cell text
            s = Target
            .Text = s
            ListLookup s
            
            ' Adapt drop down appearance to target cell
            Set m_TargetCell = Target
            AdaptControl
            
            On Error Resume Next
            Application.EnableEvents = True
            
            ' Raise BeforeDropDown event
            Cancel = False
            RaiseEvent BeforeDropDown(Target, Cancel)
            If Cancel Then
                Set m_TargetCell = Nothing
                Exit Function
            End If
            
            ' Show, activate and eventually expand drop down
            .Visible = True
            .Activate
            If m_AutoDropDown And Not NoAutoDropDown Then _
                .DropDown
            
            ' Start scroll hook for mouse wheel support if desired
            StartHook
            
            ' Raise AfterDropDown event
            RaiseEvent AfterDropDown(Target)
            
            ' Set return value
            Show = True
        End If
    End With
    
    Exit Function
    
Err:
    Application.EnableEvents = True
EndSub:

End Function


' Method Hide()
'
'' Hides the currently shown drop down.
''
'' This method usually has not to be called by the user as the drop down is
'' always hidden automatically e.g. when the cell selection is changing.
'
Public Sub Hide()
    Dim OldTarget As Range
    
    On Error Resume Next
    
    If m_Control Is Nothing Then _
        Exit Sub
    
    Set OldTarget = m_TargetCell
    
    ' Hide drop down control on old position
    m_Control.Visible = False
    
    If Not OldTarget Is Nothing Then
        Set m_TargetCell = Nothing
        StopHook
        
        ' Raise ExitDropDown event
        RaiseEvent ExitDropDown(OldTarget)
        
        ' Clear
        m_Control.ListFillRange = ""
        m_Control.Clear
    End If
End Sub


' Contructor loading default property values
Private Sub Class_Initialize()
    m_AutoShow = True
    m_AutoDropDown = True
    m_ButtonInsideCell = False
    
    m_ListFixed = False
    m_ListRowCount = 20
    m_ListScrollable = True
    
    m_ItemsFastProcessing = False
    m_ItemsSort = False
    m_ItemsRemoveDuplicate = False
    m_ItemsDuplicateCompare = vbBinaryCompare
    m_ItemsRemoveBlank = True
    m_ItemsAddFirstBlank = False
    m_ItemsTrim = False
End Sub

' Destructor
Private Sub Class_Terminate()
    ' Be sure to have the scroll hook stopped
    StopHook
End Sub

' Is called when a worksheet is assigned to initialize the class.
Private Sub Init()
    On Error Resume Next
    
    If m_Sheet Is Nothing Then _
        Exit Sub
    
    Set m_Book = ThisWorkbook
    Set m_App = m_Sheet.Application
    
    ' Try to catch ComboBox OLE object if already existing
    ' (throws an error elsewise)
    Set m_OLEObject = m_Sheet.OLEObjects(CONTROL_NAME)
 
    ' Create ComboBox object if it could not be catched
    If m_OLEObject Is Nothing Then
        Set m_OLEObject = m_Sheet.OLEObjects.Add(ClassType:=CONTROL_CLASS)
        m_OLEObject.Name = CONTROL_NAME
    End If
    m_OLEObject.Visible = False
    
    ' Get ComboBox control from OLE object
    Set m_Control = m_OLEObject.Object
    m_Control.LinkedCell = ""
End Sub

' Adapts appearance of the ComboBox control to target cell.
Private Sub AdaptControl()
    If m_TargetCell Is Nothing Or m_Control Is Nothing Then _
        Exit Sub
        
    On Error Resume Next
    
    With m_Control
        .SpecialEffect = fmSpecialEffectFlat
        .Left = m_TargetCell.Left + 1
        .Top = m_TargetCell.Top + 1
        .Height = m_TargetCell.Height - 0.5
        .Width = m_TargetCell.Width - 0.5
        If Not m_ButtonInsideCell Then _
            .Width = .Width + 13
        
        If m_ListFixed Then
            .Style = fmStyleDropDownList
        Else
            .Style = fmStyleDropDownCombo
        End If

        .HideSelection = False
        .MatchEntry = fmMatchEntryComplete
        .AutoWordSelect = True
        
        .SelectionMargin = False
        .ListRows = m_ListRowCount
        
        ' Alignment does not work properly at the moment
        .TextAlign = fmTextAlignLeft
        'Select Case targetcell.HorizontalAlignment
        'Case xlCenter
        '    .TextAlign = fmTextAlignCenter
        'Case xlRight
        '    .TextAlign = fmTextAlignRight
        'Case Else
        '    .TextAlign = fmTextAlignLeft
        'End Select
    
        .BackStyle = fmBackStyleOpaque
        .BackColor = m_TargetCell.Interior.Color
        
        .ForeColor = m_TargetCell.Font.Color
        With .Font
            .Name = m_TargetCell.Font.Name
            .Size = m_TargetCell.Font.Size
            .Bold = m_TargetCell.Font.Bold
            .Italic = m_TargetCell.Font.Italic
            .Underline = (m_TargetCell.Font.Underline = xlUnderlineStyleSingle)
            .Strikethrough = m_TargetCell.Font.Strikethrough
        End With
    End With
End Sub

' Does an exact (case sensitive) lookup of the specified text in the drop down
' list and scrolls to this position
Private Sub ListLookup(Text As String)
    Dim i As Long

    On Error GoTo ExitSub
    
    With m_Control
        If .MatchFound Then
            ' A case insensitive match was found
            
            If Text <> .List(.ListIndex) Then
                ' Match is not exact
                
                ' Look for exact match from match position onwards
                For i = .ListIndex + 1 To .ListCount - 1
                    If Text = .List(i) Then
                        ' Exact match found
                        .ListIndex = i
                        Exit For
                    End If
                Next i
            End If
            
            ' Scroll to match
            .TopIndex = .ListIndex
        End If
    End With
    
ExitSub:
End Sub

' Sorts the string array List alphabetically via Quicksort algorithm
Private Sub QuicksortStrings(ByRef List() As String, _
            Optional ByVal StartIdx As Long = -1, _
            Optional ByVal StopIdx As Long = -1)
        
    ' Rewritten by MarcoWue. Original code is from Michael Ciurescu:
    '    http://www.vbforums.com/showthread.php?t=231925

    Dim iHigh As Long
    Dim iLow As Integer
    Dim sMid As String
    Dim sTmp As String

    If StartIdx = -1 Then _
        StartIdx = UBound(List)
    If StopIdx = -1 Then _
        StopIdx = UBound(List)

    iLow = StartIdx
    iHigh = StopIdx
    sMid = List((iLow + iHigh) / 2)
    
    Do
        While List(iLow) < sMid
            iLow = iLow + 1
        Wend
        While List(iHigh) > sMid
            iHigh = iHigh - 1
        Wend
        
        If iLow <= iHigh Then
            sTmp = List(iHigh)
            List(iHigh) = List(iLow)
            List(iLow) = sTmp
            
            iLow = iLow + 1
            iHigh = iHigh - 1
        End If
    Loop While iLow <= iHigh

    If iHigh > StartIdx Then _
        QuicksortStrings List, StartIdx, iHigh
    If iLow < StopIdx Then _
        QuicksortStrings List, iLow, StopIdx
End Sub

' Starts the scroll hook for mouse wheel support of the drop down list.
Private Sub StartHook(Optional ByVal ResumeHook As Boolean = False)
    If Not m_ListScrollable Or m_Control Is Nothing Then _
        Exit Sub
    If Not m_Book.ActiveSheet Is m_Sheet Then _
        Exit Sub
    If Not ResumeHook Or m_HookPaused Then _
        StartMouseWheelHook m_Control
    m_HookPaused = False
End Sub

' Stops the scroll hook
Private Sub StopHook(Optional ByVal PauseHook As Boolean = False)
    If Not m_ListScrollable Or m_Control Is Nothing Then _
        Exit Sub
    StopMouseWheelHook m_Control
    m_HookPaused = PauseHook
End Sub


' Write back drop down changes to the target cell
Private Sub m_Control_Change()
    On Error Resume Next
    If m_TargetCell Is Nothing Then _
        Exit Sub
    With m_Control
        m_TargetCell = .Text
        ListLookup .Text ' Scroll to match in drop down list
    End With
End Sub

' Emulate usual key press reactions
Private Sub m_Control_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, _
            ByVal Shift As Integer)
    
    If m_TargetCell Is Nothing Then _
        Exit Sub
        
    If KeyCode = 9 Then
        ' Tab
        Hide
        SendKeys "{TAB}"
        
    ElseIf KeyCode = 13 Then
        ' Enter
        Hide
        SendKeys "{ENTER}"
        
    ElseIf KeyCode = 27 Then
        ' ESC
        
        ' Pressing ESC, Excel lets the drop down control lose focus and puts
        ' the cursor into the underlaying cell (without drawing an active
        ' border around it). So what we want to do here is to correctly
        ' activate the underlaying cell and to hide the drop down control.
        m_TargetCell.Activate
        Hide
    End If
End Sub

' Automatically show drop down when cell selection changed
Private Sub m_Sheet_SelectionChange(ByVal Target As Range)
    If m_AutoShow Then _
        Show Target
End Sub

' Adapt currently shown drop down appearance to changes of worksheet
Private Sub m_Sheet_Change(ByVal Target As Range)
    AdaptControl
End Sub

' Pause/resume scroll hook for mouse wheel support
Private Sub m_Sheet_Deactivate()
    StopHook True
End Sub

Private Sub m_Sheet_Activate()
    StartHook True
End Sub

Private Sub m_App_WindowDeactivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If Wb Is m_Book Then _
        StopHook True
End Sub

Private Sub m_App_WindowActivate(ByVal Wb As Workbook, ByVal Wn As Window)
    If Wb Is m_Book Then _
        StartHook True
End Sub

Private Sub m_Book_Deactivate()
    StopHook True
End Sub

Private Sub m_Book_Activate()
    StartHook True
End Sub
