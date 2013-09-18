Attribute VB_Name = "ExcelMousewheelSupport"
' **
' ** ExcelMouseWheelSupport.bas
'/**
'
' @brief        *ExcelMouseWheelSupport.bas* adds mouse wheel scrolling
'               support to Microsoft Excel ActiveX Controls (e.g. the
'               ComboBox). This is achieved by low level mouse event
'               capturing via message hook.
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
' None.
'
'
' Usage
' -----
'
' @note     ExcelMouseWheelSupport needs activated macros to work!
'
' Call StartMouseWheelHook() when the specified control gets the focus.
' This will start the mouse event capturing. Call StopMouseWheelHook()
' after the control has lost focus.
'
' @note     After you started the hook, be sure to call StopMouseWheelHook()
'           whenever the Worksheet, Workbook or Window is being deactivated.
'
'
' Credits
' -------
'
' Thanks to Jaafar Tribak on whose work this code is based on. See
' <a href="http://www.mrexcel.com/forum/excel-questions/559658-combobox-scroll-down-enabled.html#post2765506" target="_blank">here</a>.
'
'
' Version History
' ---------------
'
'  - 1.0.1 (2013-09-18) *MarcoWue*
'       - Ported to Doxygen code documentation.
'
'  - 1.0.0 (2013-09-03) *MarcoWue*
'       - Initial release.
'
'**/


Option Explicit


'' Scroll speed (number of rows to scroll at once).
'' This constant can be changed to adjust the scrolling behaviour.
Private Const SCROLL_SPEED = 2

'' Class name of the Excel main window. May changes in further Excel versions.
Private Const MAINWINDOW_CLASSNAME = "XLMAIN"


Private Type POINTAPI
  X As Long
  Y As Long
End Type

Private Type MSLLHOOKSTRUCT
    pt As POINTAPI
    mousedata As Long
    flags As Long
    time As Long
    dwExtraInfo As Long
End Type

Private Declare Sub CopyMemory Lib "kernel32" _
    Alias "RtlMoveMemory" _
    (ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long)

Private Declare Function FindWindow Lib "user32.dll" _
    Alias "FindWindowA" _
    (ByVal lpClassName As String, _
    ByVal lpWindowName As String) As Long

Private Declare Function GetForegroundWindow Lib "user32" () As Long
    
Declare Function IsChild Lib "user32.dll" ( _
    ByVal hWndParent As Long, _
    ByVal hwnd As Long) As Long
    
Private Declare Function GetWindowLong Lib "user32" _
    Alias "GetWindowLongA" ( _
    ByVal hwnd As Long, _
    ByVal nIndex As Long) As Long

Private Declare Function SetWindowsHookEx Lib "user32" _
    Alias "SetWindowsHookExA" _
    (ByVal idHook As Long, _
    ByVal lpfn As Long, _
    ByVal hmod As Long, _
    ByVal dwThreadId As Long) As Long

Private Declare Function CallNextHookEx Lib "user32" _
    (ByVal hHook As Long, _
    ByVal nCode As Long, _
    ByVal wParam As Long, _
    lParam As Any) As Long

Private Declare Function UnhookWindowsHookEx Lib "user32" _
    (ByVal hHook As Long) As Long


Private Const HC_ACTION = 0
Private Const WH_MOUSE_LL = 14
Private Const WM_MOUSEWHEEL = &H20A
Private Const GWL_HINSTANCE = (-6)

Private HookObject As Object
Private hHook As Long
Private hMainWindow As Long


' Sub StartMouseWheelHook
'
'' Starts the hook to make the specified object scrollable with mouse wheel.
''
'' @param[in]   Obj             The object to be made scrollable.
'
Public Sub StartMouseWheelHook(ByVal Obj As Object)
    If Obj Is Nothing Then _
        Exit Sub
    
    Set HookObject = Obj
    hMainWindow = FindWindow(MAINWINDOW_CLASSNAME, Application.Caption)
    
    If hHook = 0 Then
        hHook = SetWindowsHookEx(WH_MOUSE_LL, AddressOf LowLevelMouseProc, _
                    GetWindowLong(hMainWindow, GWL_HINSTANCE), 0)
    End If
End Sub

' Sub StopMouseWheelHook
'
'' Stops the hook if the hook was started with the specified object.
''
'' @param[in]   Obj             The object on which the hook was started.
'
Public Sub StopMouseWheelHook(ByVal Obj As Object)
    If hHook = 0 Then _
        Exit Sub
    If Not Obj Is HookObject Then _
        Exit Sub
        
    UnhookWindowsHookEx hHook
    hHook = 0
    Set HookObject = Nothing
End Sub


' Hook callback function
Function LowLevelMouseProc(ByVal nCode As Long, ByVal wParam As Long, _
            ByVal lParam As Long) As Long
        
    Dim uParamStruct As MSLLHOOKSTRUCT
    
    On Error GoTo ExitProc
    
    ' Doing multiple if's here for performance reasons
    ' (VBA is always checking every And condition).
    
    If GetForegroundWindow = hMainWindow Then
        If nCode = HC_ACTION Then
            If wParam = WM_MOUSEWHEEL Then
            
                CopyMemory VarPtr(uParamStruct), lParam, LenB(uParamStruct)
                
                With HookObject
                    If uParamStruct.mousedata > 0 Then
                        .TopIndex = .TopIndex - SCROLL_SPEED
                    Else
                        .TopIndex = .TopIndex + SCROLL_SPEED
                    End If
                End With
                
                LowLevelMouseProc = -1
                Exit Function
                
            End If
        End If
    End If

ExitProc:
    LowLevelMouseProc = CallNextHookEx(hHook, nCode, wParam, ByVal lParam)
End Function
