VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmRTF 
   Caption         =   "Form1"
   ClientHeight    =   5415
   ClientLeft      =   4455
   ClientTop       =   3240
   ClientWidth     =   9240
   LinkTopic       =   "Form1"
   ScaleHeight     =   5415
   ScaleWidth      =   9240
   Begin VB.CommandButton Command5 
      Caption         =   "justify "
      Height          =   375
      Left            =   4770
      TabIndex        =   5
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Left"
      Height          =   375
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "center"
      Height          =   375
      Left            =   1170
      TabIndex        =   2
      Top             =   240
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "right"
      Height          =   375
      Left            =   2100
      TabIndex        =   1
      Top             =   240
      Width           =   885
   End
   Begin VB.CommandButton Command4 
      Caption         =   "justify "
      Height          =   375
      Left            =   3060
      TabIndex        =   0
      Top             =   240
      Width           =   885
   End
   Begin RichTextLib.RichTextBox RichTextBox1 
      Height          =   2145
      Left            =   210
      TabIndex        =   4
      Top             =   1230
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   3784
      _Version        =   393217
      TextRTF         =   $"frmRTF.frx":0000
   End
End
Attribute VB_Name = "frmRTF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal HWND As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Const WM_USER = &H400
Const EM_SETTYPOGRAPHYOPTIONS = WM_USER + 202
Const TO_ADVANCEDTYPOGRAPHY = 1
Const EM_SETPARAFORMAT = WM_USER + 71
Private Const PFA_LEFT = 1
Private Const PFA_RIGHT = 2
Private Const PFA_CENTER = 3
Private Const PFA_JUSTIFY = &H4
Const MAX_TAB_STOPS = 32
Private Type PARAFORMAT2
    cbSize                     As Long
    dwMask                     As Long
    wNumbering                 As Integer
    wEffects                   As Integer
    dxStartIndent              As Long
    dxRightIndent              As Long
    dxOffset                   As Long
    wAlignment                 As Integer
    cTabCount                  As Integer
    rgxTabs(MAX_TAB_STOPS - 1) As Long
    dySpaceBefore              As Long
    dySpaceAfter               As Long
    dyLineSpacing              As Long
    sStyle                     As Integer
    bLineSpacingRule           As Byte
    bOutlineLevel              As Byte
    wShadingWeight             As Integer
    wShadingStyle              As Integer
    wNumberingStart            As Integer
    wNumberingStyle            As Integer
    wNumberingTab              As Integer
    wBorderSpace               As Integer
    wBorderWidth               As Integer
    wBorders                   As Integer
End Type
Public Enum ERECParagraphAlignmentConstants
   ercParaLeft = PFA_LEFT
   ercParaCentre = PFA_CENTER
   ercParaRight = PFA_RIGHT
   ercParaJustify = PFA_JUSTIFY
End Enum
Private Const PFM_ALIGNMENT = &H8&
 
Private Function SetAlignment(lHwnd As Long, ByVal eAlign As ERECParagraphAlignmentConstants)
    Dim tP2 As PARAFORMAT2
    Dim lR As Long
    tP2.dwMask = PFM_ALIGNMENT
    tP2.cbSize = Len(tP2)
    tP2.wAlignment = eAlign
    lR = SendMessageLong(lHwnd, EM_SETTYPOGRAPHYOPTIONS, TO_ADVANCEDTYPOGRAPHY, TO_ADVANCEDTYPOGRAPHY)
    lR = SendMessage(lHwnd, EM_SETPARAFORMAT, 0, tP2)
End Function



Private Sub Command1_Click()
SetAlignment RichTextBox1.HWND, ercParaLeft
End Sub

Private Sub Command2_Click()
SetAlignment RichTextBox1.HWND, ercParaCentre
End Sub

Private Sub Command3_Click()
SetAlignment RichTextBox1.HWND, ercParaRight
End Sub

Private Sub Command4_Click()
SetAlignment RichTextBox1.HWND, ercParaJustify
End Sub

