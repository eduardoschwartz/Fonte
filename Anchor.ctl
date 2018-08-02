VERSION 5.00
Begin VB.UserControl Anchor 
   ClientHeight    =   495
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   495
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Anchor.ctx":0000
   PropertyPages   =   "Anchor.ctx":0BE2
   ScaleHeight     =   495
   ScaleWidth      =   495
   ToolboxBitmap   =   "Anchor.ctx":0BF2
End
Attribute VB_Name = "Anchor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'************************************************
'* Developers : Hamed Oveisi & Patrick de Groot *
'************************************************
'Advanced Anchor Control
'Version 4 (May 2014)
'With these features:
'- Resizes controls like in .NET and Delphi
'- Uses Freezing Technic to increase the speed of resizing Controls
'- Limits form width and/or height
'- Hooking into window / subclass to limit minimum and maximum form size
'  Do not use the rude END command while debugging, but close all forms!
'- MinWidth, MinHeight, MaxWidth, MaxHeight, UseWinHook properties
'- Keyboard shortcuts to property page to make achor assignments quick
'- MDI child form compatibility
'- Save form position and size between program runs
'- Plug & Play / Drop it on your form & Run
'- Fixes and optimalizations
'Developed By : Hamed Oveisi & Patrick de Groot
'Please leave your feedback and vote on PSC

Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function RedrawWindow Lib "user32" (ByVal hwnd As Long, lprcUpdate As Any, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Const WM_SETREDRAW = &HB
Private Const RDW_INVALIDATE = &H1
Private Const RDW_INTERNALPAINT = &H2
Private Const RDW_ALLCHILDREN = &H80
Private Const RDW_UPDATENOW = &H100
Private Const WM_PAINT = &HF

Private LastHeight As Long, LastWidth As Long    'Form last Height and Width values
Private WithEvents Frm As Form
Attribute Frm.VB_VarHelpID = -1
Private Tags() As String     'Saves Control Tag properties @RunTime
Attribute Tags.VB_VarUserMemId = 1073938447
Private InitWidth As Long, InitHeight As Long     'The initial form size at design time (usefull for MDI child forms)
Public SavePosition As Boolean  'Save last position of the form for the next time
Public SaveSize As Boolean
Public MinHeight As Long, MinWidth As Long    'Min Height and Width of Form so that form never get smaller than these values
Public MaxHeight As Long, MaxWidth As Long
Public UseWinHook As Boolean    'Use winproc hooking to optimize min/max size handling (not using it causes flickering when limiting size)
Public Event AfterResize(ByVal HeightChanged As Long, ByVal WidthChanged As Long)
Public Event BeforeResize()

Public Sub Freeze(ByVal ohWnd As Long)
'Freeze the Application
    SendMessage ohWnd, WM_SETREDRAW, False, 0
End Sub

Public Sub UnFreeze(ByVal ohWnd As Long, Optional ByVal ForceUpdate As Boolean = True)
'Unfreeze the application
    SendMessage ohWnd, WM_SETREDRAW, 1&, 0
    'SendMessage ohWnd, WM_PAINT, 1&, 0
    If ForceUpdate Then
        RedrawWindow ohWnd, ByVal 0&, ByVal 0&, RDW_INVALIDATE Or RDW_UPDATENOW Or RDW_ALLCHILDREN
    End If
End Sub

'DoResize checks the control Tag property and resize it. Use Tag this way: TTTT (T/F for True/False) for Left,Top,Right and Bottom anchor
Public Sub DoResize()
    On Error Resume Next
    Dim HeightChange As Long, WidthChange As Long
    Dim NewSize As Long
    Dim Tg As String
    Dim i As Integer
    Dim Ctrl As Control

    'Exit sub on Minimize
    If Frm.WindowState = vbMinimized Then Exit Sub
    'Freezing the Control Redraws makes the application faster in resizing controls.
    Freeze Frm.hwnd
    If Frm.Height < MinHeight Then Frm.Height = MinHeight
    If Frm.Width < MinWidth Then Frm.Width = MinWidth
    If MaxHeight > 0 And Frm.Height > MaxHeight Then Frm.Height = MaxHeight
    If MaxWidth > 0 And Frm.Width > MaxWidth Then Frm.Width = MaxWidth
    'Calculate the changes
    HeightChange = Frm.Height - LastHeight
    WidthChange = Frm.Width - LastWidth
    If HeightChange <> 0 Or WidthChange <> 0 Then
        RaiseEvent BeforeResize
        For Each Ctrl In Frm.Controls
            Tg = Tags(i)
            If LenB(Tg) = 0 Then GoTo Nxt
            If Right(Tg, 2) <> "FF" Then
                If Right(Tg, 1) = "T" Then
                    If Mid(Tg, 2, 1) = "T" Then
                        NewSize = Ctrl.Height + HeightChange
                        If NewSize > 0 Then
                            Ctrl.Height = NewSize
                        End If
                    Else
                        Ctrl.Top = Ctrl.Top + HeightChange
                    End If
                End If
                If Mid(Tg, 3, 1) = "T" Then
                    If Left(Tg, 1) = "T" Then
                        NewSize = Ctrl.Width + WidthChange
                        If NewSize > 0 Then
                            Ctrl.Width = Ctrl.Width + WidthChange
                        End If
                    Else
                        Ctrl.Left = Ctrl.Left + WidthChange
                    End If
                End If
            End If
Nxt:
            i = i + 1
        Next
        RaiseEvent AfterResize(HeightChange, WidthChange)
    End If
    'Save current form dimensions
    LastHeight = Frm.Height
    LastWidth = Frm.Width
    UnFreeze Frm.hwnd, True
End Sub

Private Sub Frm_Activate()
    SetFormMinMax
End Sub

Sub SetFormMinMax()
    Set CtrlParent = Frm
    MinWidthPix = MinWidth \ Screen.TwipsPerPixelX
    MinHeightPix = MinHeight \ Screen.TwipsPerPixelY
    MaxWidthPix = MaxWidth \ Screen.TwipsPerPixelX
    MaxHeightPix = MaxHeight \ Screen.TwipsPerPixelY
End Sub

Private Sub Frm_Load()
    If Ambient.UserMode Then
        DoInit
    End If
End Sub

Private Sub Frm_Resize()
    DoResize
End Sub

Private Sub UserControl_InitProperties()
    On Error Resume Next
    Set Frm = Extender.Parent
    Set CtrlParent = Extender.Parent
    UseWinHook = True
End Sub

Private Sub UserControl_Paint()
    If Not Ambient.UserMode Then
        InitWidth = UserControl.Extender.Parent.Width
        InitHeight = UserControl.Extender.Parent.Height
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "MinWidth", MinWidth, 0
    PropBag.WriteProperty "MinHeight", MinHeight, 0
    PropBag.WriteProperty "MaxWidth", MaxWidth, 0
    PropBag.WriteProperty "MaxHeight", MaxHeight, 0
    PropBag.WriteProperty "InitWidth", InitWidth, 0
    PropBag.WriteProperty "InitHeight", InitHeight, 0
    PropBag.WriteProperty "SavePosition", SavePosition, False
    PropBag.WriteProperty "SaveSize", SaveSize, False
    PropBag.WriteProperty "UseWinHook", UseWinHook, True
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    Set Frm = Extender.Parent
    Set CtrlParent = Extender.Parent
    MinWidth = PropBag.ReadProperty("MinWidth", 0)
    MinHeight = PropBag.ReadProperty("MinHeight", 0)
    MaxWidth = PropBag.ReadProperty("MaxWidth", 0)
    MaxHeight = PropBag.ReadProperty("MaxHeight", 0)
    InitWidth = PropBag.ReadProperty("InitWidth", 0)
    InitHeight = PropBag.ReadProperty("InitHeight", 0)
    SavePosition = PropBag.ReadProperty("SavePosition", False)
    SaveSize = PropBag.ReadProperty("SaveSize", False)
    UseWinHook = PropBag.ReadProperty("UseWinHook", True)
End Sub

Private Sub UserControl_Resize()
    Width = 480
    Height = 465
End Sub

Public Sub About()
Attribute About.VB_UserMemId = -552
    Dim Msg As String

    Msg = "Anchor Control for Visual Basic v4" & vbCr
    Msg = Msg & "Written by : Hamed Oveisi & Patrick de Groot"
    MsgBox Msg, vbInformation, "About Anchor"
End Sub

Private Sub UserControl_Terminate()
    If Not Frm Is Nothing Then
        If Frm.WindowState = vbNormal Then
            If SavePosition Then
                SaveSetting App.EXEName, Frm.Name, "Left", Frm.Left
                SaveSetting App.EXEName, Frm.Name, "Top", Frm.Top
            End If
            If SaveSize Then
                SaveSetting App.EXEName, Frm.Name, "Width", Frm.Width
                SaveSetting App.EXEName, Frm.Name, "Height", Frm.Height
            End If
        End If
    End If
    If UseWinHook Then
        Unhook
    End If
    Set Frm = Nothing
    Set CtrlParent = Nothing
End Sub

Public Sub DoInit()
    On Error Resume Next
    Dim i As Integer
    Dim Tg As String
    Dim Ctrl As Control
    
    If InitWidth > 0 Then Frm.Width = InitWidth
    If InitHeight > 0 Then Frm.Height = InitHeight
    If MinHeight < 1 Then MinHeight = Frm.Height
    If MinWidth < 1 Then MinWidth = Frm.Width
    ReDim Tags(Frm.Controls.Count)
    'Save Tag Properties of Controls so that Tag can be used in runtime for other purposes
    For Each Ctrl In Frm.Controls
        Tg = Ctrl.Tag
        If LenB(Tg) = 0 Then GoTo Nxt
        'Every anchor information ends with */
        If InStr(1, Tg, "*/") > 0 Then
            Tags(i) = Left(Tg, 4)
            'Eliminate anchors from Tag property so there is no dependency to the Tag property of the object @RunTime.
            Ctrl.Tag = Right(Tg, Len(Tg) - 6)
        End If
Nxt:
        i = i + 1
    Next
    LastHeight = Frm.Height
    LastWidth = Frm.Width
    If SavePosition Then
        'Restore last form position
        Frm.Left = GetSetting(App.EXEName, Frm.Name, "Left", Frm.Left)
        Frm.Top = GetSetting(App.EXEName, Frm.Name, "Top", Frm.Top)
    End If
    If SaveSize Then
        'Restore last form size
        Frm.Width = GetSetting(App.EXEName, Frm.Name, "Width", Frm.Width)
        Frm.Height = GetSetting(App.EXEName, Frm.Name, "Height", Frm.Height)
    End If
    SetFormMinMax
    If UseWinHook Then
        Hook
    End If
End Sub

