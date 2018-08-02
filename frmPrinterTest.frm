VERSION 5.00
Begin VB.Form frmPrinterTest 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleção de impressora padrão "
   ClientHeight    =   4845
   ClientLeft      =   4680
   ClientTop       =   3120
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4845
   ScaleWidth      =   4185
   Begin VB.TextBox txtLeft 
      Height          =   315
      Left            =   510
      TabIndex        =   9
      Text            =   "0"
      Top             =   3240
      Width           =   750
   End
   Begin VB.TextBox txtDown 
      Height          =   315
      Left            =   1650
      TabIndex        =   8
      Text            =   "0"
      Top             =   3780
      Width           =   750
   End
   Begin VB.TextBox txtRight 
      Height          =   315
      Left            =   2430
      TabIndex        =   7
      Text            =   "0"
      Top             =   3240
      Width           =   750
   End
   Begin VB.TextBox txtUp 
      Height          =   315
      Left            =   1650
      TabIndex        =   6
      Text            =   "0"
      Top             =   2700
      Width           =   750
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Sair"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1500
      TabIndex        =   4
      Top             =   4290
      Width           =   1215
   End
   Begin VB.ComboBox cboPrinters 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   225
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   1320
      Width           =   3735
   End
   Begin VB.TextBox txtDefault 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   225
      Locked          =   -1  'True
      TabIndex        =   0
      Top             =   480
      Width           =   3735
   End
   Begin VB.Image Image1 
      Height          =   600
      Left            =   1740
      Picture         =   "frmPrinterTest.frx":0000
      Stretch         =   -1  'True
      Top             =   3090
      Width           =   600
   End
   Begin VB.Label lbl2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Para alterar a impressora padrão, selecione uma da lista acima."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   555
      Left            =   390
      TabIndex        =   5
      Top             =   2130
      Width           =   3735
   End
   Begin VB.Label Label1 
      Caption         =   "Lista de Todas as Impressoras"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   1080
      Width           =   2895
   End
   Begin VB.Label lbl1 
      Caption         =   "Impressora padrão selecionada"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   2895
   End
End
Attribute VB_Name = "frmPrinterTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' ========================================================
' Prism_Printers by Jim Ryan, Prism Software Co 11/22/2009

' USE:
' If you are using software other than VB to print forms
' from your project (Example: Crystal Reports .DSR
' then you know that Crystal will only print to the
' system current default printer.  Using this class you
' can 1) set the default printer to a different printer,
' 2) print the Crystal report and 3) set the default
' printer back to it's original setting.

' TO USE:
' #1
' Under Project - References
' Select "Windows Script Host Object Model"
'
' #2
' Under Project - Add Class Module
' Select "PrismPrinters.cls
' ========================================================

' Set dp as a new PrismPrinters class
Dim dp As New PrismPrinters

Dim OldDefaultPrinter As String
Dim Loaded As Boolean

Private Sub cboPrinters_Click()
   If Loaded Then
      Me.MousePointer = vbHourglass
      If dp.SetDefaultPrinter(cboPrinters.Text) Then
         txtDefault.Text = cboPrinters.Text
      End If
      Me.MousePointer = vbNormal
   End If
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub Form_Load()
Dim prnX As Object, Sql As String, RdoAux As rdoResultset
Centraliza Me
   frmPrinterTest.show
   frmPrinterTest.Refresh
   
   ' Show the current Default printer
   txtDefault.Text = dp.GetDefaultPrinter
   OldDefaultPrinter = txtDefault.Text
   
   ' Create a list of printers in cboPrinter
   For Each prnX In Printers
      cboPrinters.AddItem prnX.DeviceName
   Next
   Set prnX = Nothing
   'cboPrinters.ListIndex = 0
   Loaded = True
   
    Sql = "select * from machines2 where computer='" & NomeDoComputador & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        txtUp.Text = !margin_top
        txtLeft.Text = !Margin_left
        txtDown.Text = !margin_bottom
        txtRight.Text = !margin_right
       .Close
    End With
   
   
   
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Dim Sql As String

Sql = "update machines2 set margin_top=" & Val(txtUp.Text) & ",margin_left=" & Val(txtLeft.Text) & ",margin_bottom=" & Val(txtDown.Text) & ",margin_right=" & Val(txtRight.Text) & " where computer='" & NomeDoComputador & "'"
cn.Execute Sql, rdExecDirect

nMargem_Top = Val(txtUp.Text)
nMargem_Left = Val(txtLeft.Text)
nMargem_Bottom = Val(txtDown.Text)
nMargem_Right = Val(txtRight.Text)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set dp = Nothing
   Set frmPrinterTest = Nothing
End Sub

Private Sub txtDown_GotFocus()
txtDown.SelStart = 0
txtDown.SelLength = Len(txtDown.Text)

End Sub

Private Sub txtDown_KeyPress(KeyAscii As Integer)
Tweak txtDown, KeyAscii, IntegerPositive
End Sub

Private Sub txtLeft_GotFocus()
txtLeft.SelStart = 0
txtLeft.SelLength = Len(txtLeft.Text)

End Sub

Private Sub txtLeft_KeyPress(KeyAscii As Integer)
Tweak txtLeft, KeyAscii, IntegerPositive
End Sub

Private Sub txtRight_GotFocus()
txtRight.SelStart = 0
txtRight.SelLength = Len(txtRight.Text)

End Sub

Private Sub txtRight_KeyPress(KeyAscii As Integer)
Tweak txtRight, KeyAscii, IntegerPositive
End Sub

Private Sub txtUp_GotFocus()
txtUp.SelStart = 0
txtUp.SelLength = Len(txtUp.Text)
End Sub

Private Sub txtUp_KeyPress(KeyAscii As Integer)
Tweak txtUp, KeyAscii, IntegerPositive
End Sub
