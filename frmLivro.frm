VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmLivro 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Livro da Divida Ativa"
   ClientHeight    =   2280
   ClientLeft      =   5355
   ClientTop       =   4020
   ClientWidth     =   6015
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2280
   ScaleWidth      =   6015
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      Left            =   1950
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   450
      Width           =   2265
   End
   Begin VB.ComboBox cmbAno 
      Appearance      =   0  'Flat
      Height          =   315
      ItemData        =   "frmLivro.frx":0000
      Left            =   1950
      List            =   "frmLivro.frx":0007
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   90
      Width           =   915
   End
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1950
      TabIndex        =   2
      Top             =   1260
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   503
      MouseIcon       =   "frmLivro.frx":0011
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Appearance      =   0
      MaxLength       =   10
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdRetorna 
      Height          =   315
      Left            =   4710
      TabIndex        =   10
      ToolTipText     =   "Cadastra a Área"
      Top             =   1860
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cadastrar"
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
      MICON           =   "frmLivro.frx":002D
      PICN            =   "frmLivro.frx":0049
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   3540
      TabIndex        =   11
      ToolTipText     =   "Sair da Tela"
      Top             =   1860
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Sair"
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
      MICON           =   "frmLivro.frx":00B7
      PICN            =   "frmLivro.frx":00D3
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblNumero 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   1980
      TabIndex        =   9
      Top             =   930
      Width           =   3990
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Número do Livro...........:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   8
      Top             =   930
      Width           =   1755
   End
   Begin VB.Label lblDataFim 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   4980
      TabIndex        =   7
      Top             =   1290
      Width           =   960
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Encerramento..:"
      Height          =   225
      Index           =   0
      Left            =   3135
      TabIndex        =   6
      Top             =   1290
      Width           =   1755
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Data de Abertura..........:"
      Height          =   225
      Index           =   5
      Left            =   105
      TabIndex        =   5
      Top             =   1305
      Width           =   1755
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Livro................:"
      Height          =   195
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   525
      Width           =   1755
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano do Livro.................:"
      Height          =   210
      Left            =   105
      TabIndex        =   3
      Top             =   165
      Width           =   1755
   End
End
Attribute VB_Name = "frmLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, bNovo As Boolean

Private Sub cmbAno_Click()
If cmbAno.ListIndex = -1 Then Exit Sub
Limpa
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
cmbTipo_Click
End Sub

Private Sub cmbTipo_Click()
Dim nLast As Integer, sNum As String
If cmbTipo.ListIndex = -1 Then Exit Sub
Limpa
Sql = "SELECT NUMERO,DATAABERTURA,DATAENCERRAMENTO FROM LIVRO WHERE ANO=" & Val(cmbAno.Text)
Sql = Sql & " AND CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
       bNovo = False
       Sql = "SELECT NUMEROOLD FROM GRADELIVRO WHERE ANO=" & Val(cmbAno.Text) & " AND "
       Sql = Sql & "CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex) & " ORDER BY NUMEROOLD"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       With RdoAux2
            If .RowCount > 0 Then
                Do Until .EOF
                    sNum = sNum & CStr(!NUMEROOLD) & ", "
                   .MoveNext
                Loop
                sNum = Chomp(sNum, chomp_righT, 1)
                lblNumero.Caption = sNum
            Else
                lblNumero.Caption = RdoAux!Numero
            End If
           .Close
       End With
       If Not IsNull(!DATAABERTURA) Then mskDataIni.Text = Format(!DATAABERTURA, "dd/mm/yyyy")
       If Not IsNull(!dataencerramento) Then lblDataFim.Caption = Format(!dataencerramento, "dd/mm/yyyy")
    Else
       bNovo = True
       Sql = "SELECT MAX(NUMERO) AS MAXIMO FROM LIVRO"
       Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
       If Not IsNull(RdoAux2!MAXIMO) Then
          nLast = RdoAux2!MAXIMO
       Else
          nLast = 0
       End If
       RdoAux2.Close
       lblNumero.Caption = nLast + 1
    End If
   .Close
End With

End Sub


Private Sub cmdRetorna_Click()
Dim nLastYear As Integer

If lblDataFim.Caption <> "" Then
    MsgBox "Não é possível modificar este livro.", vbCritical, "Atenção"
    Exit Sub
End If

Sql = "SELECT * FROM LIVRO WHERE NUMERO=" & Val(lblNumero.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        MsgBox "Livro já aberto.", vbExclamation, "Atenção"
        Exit Sub
    End If
End With

Sql = "SELECT MAX(ANO) AS MAXIMO FROM LIVRO WHERE CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If Not IsNull(!MAXIMO) Then
        nLastYear = !MAXIMO
       .Close
    Else
        nLastYear = Year(Now)
    End If
End With

If Val(cmbAno.Text) < 2003 Then
    MsgBox "Só é possível abrir um novo livro para anos subsequentes.", vbCritical, "atenção"
    Exit Sub
End If

If Val(cmbAno.Text) - nLastYear > 1 Then
    MsgBox "Só é possível abrir um novo livro para anos subsequentes.", vbCritical, "atenção"
    Exit Sub
End If

If Not IsDate(mskDataIni.Text) Then
    MsgBox "Data de Abertura inválida.", vbCritical, "atenção"
    Exit Sub
End If

If Year(CDate(mskDataIni.Text)) <> Val(cmbAno.Text) Then
    MsgBox "Data de Abertura tem que ser a do ano selecionado.", vbCritical, "atenção"
    Exit Sub
End If

If bNovo Then
   If MsgBox(" Deseja abrir este Livro?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
      Sql = "INSERT LIVRO (NUMERO,CODTIPO,ANO,DATAABERTURA) VALUES("
      Sql = Sql & Val(lblNumero.Caption) & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & ","
      Sql = Sql & Val(cmbAno.Text) & ",'" & Format(mskDataIni.Text, "mm/dd/yyyy") & "')"
      cn.Execute Sql, rdExecDirect
   End If
Else
   If MsgBox(" Deseja atualizar Livro?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
      Sql = "UPDATE LIVRO SET DATAABERTURA='" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' WHERE "
      Sql = Sql & "ANO=" & Val(cmbAno.Text) & " AND CODTIPO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
      cn.Execute Sql, rdExecDirect
   End If
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me

For x = 2010 To Year(Now) + 5
    cmbAno.AddItem CStr(x)
Next

cmbAno.Text = Year(Now)
Sql = "SELECT CODTIPO, DESCTIPO FROM TIPOLIVRO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCTIPO
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodTipo
       .MoveNext
    Loop
   .Close
End With
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0

End Sub

Private Sub Limpa()

lblNumero.Caption = ""
LimpaMascara mskDataIni
lblDataFim.Caption = ""
End Sub
