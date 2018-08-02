VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frm2vianotificacaoiss 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "2ª via de Notificação de ISS"
   ClientHeight    =   3270
   ClientLeft      =   4110
   ClientTop       =   2655
   ClientWidth     =   3525
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3270
   ScaleWidth      =   3525
   Begin VB.ListBox lstFiles 
      Height          =   1425
      Left            =   180
      TabIndex        =   7
      Top             =   1170
      Width           =   3165
   End
   Begin VB.FileListBox File1 
      Height          =   1455
      Left            =   4140
      TabIndex        =   6
      Top             =   1170
      Width           =   3120
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   1755
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   720
      Width           =   1050
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1755
      MaxLength       =   6
      TabIndex        =   1
      Top             =   225
      Width           =   1050
   End
   Begin prjChameleon.chameleonButton cmdRefresh 
      Height          =   300
      Left            =   2880
      TabIndex        =   4
      ToolTipText     =   "Consultar Assuntos por parte do nome"
      Top             =   225
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
      BTYPE           =   3
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frm2vianotificacaoiss.frx":0000
      PICN            =   "frm2vianotificacaoiss.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdOpen 
      Height          =   345
      Left            =   2115
      TabIndex        =   5
      ToolTipText     =   "Abrir o arquivo selecionado"
      Top             =   2790
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Abrir"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frm2vianotificacaoiss.frx":0176
      PICN            =   "frm2vianotificacaoiss.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      Caption         =   "Ano da notificação:"
      Height          =   240
      Left            =   225
      TabIndex        =   2
      Top             =   765
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Código do imóvel..:"
      Height          =   195
      Left            =   225
      TabIndex        =   0
      Top             =   270
      Width           =   1455
   End
End
Attribute VB_Name = "frm2vianotificacaoiss"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbAno_Click()
Dim x As Integer
lstFiles.Clear
File1.Path = "\\192.168.200.130\atualizagti\documentos\" & cmbAno.Text
File1.Pattern = "05*" & Format(txtCod.Text, "000000") & ".pdf"

For x = 0 To File1.ListCount - 1
    lstFiles.AddItem File1.List(x)
Next

End Sub

Private Sub cmdOpen_Click()
ShellExecute Me.hwnd, vbNullString, "\\192.168.200.130\atualizagti\documentos\" & cmbAno.Text & "\" & lstFiles.Text, vbNullString, vbNullString, vbNormalFocus
End Sub

Private Sub cmdRefresh_Click()
Dim Sql As String, RdoAux As rdoResultset, nAno As Integer, sDoc As String, x As Integer, bFind As Boolean

If Val(txtCod.Text) = 0 Then
    MsgBox "Digite o código do imóvel.", vbCritical, "Atenção!"
    Exit Sub
End If
cmbAno.Clear

Sql = "select * from documentopic WHERE (SUBSTRING(documento, 1, 2) = '05') AND (RIGHT(documento, 3) = 'PDF')"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sDoc = !Documento
        sDoc = Right(sDoc, 10)
        sDoc = Left(sDoc, 6)
        If sDoc = Format(txtCod.Text, "000000") Then
            nAno = Mid(!Documento, 3, 4)
            bFind = False
            For x = 0 To cmbAno.ListCount - 1
                If cmbAno.List(x) = CStr(nAno) Then
                    bFind = True
                    Exit For
                End If
            Next
            If Not bFind Then
                cmbAno.AddItem CStr(nAno)
            End If
        End If
       .MoveNext
    Loop
   .Close
End With

If cmbAno.ListCount > 0 Then
   cmbAno.ListIndex = 0
Else
    MsgBox "Código não encontrado!", vbInformation, "Atenção"
    txtCod.SetFocus
End If

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub txtCod_Change()
cmbAno.Clear
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    cmdRefresh_Click
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If
End Sub
