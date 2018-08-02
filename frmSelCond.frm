VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSelCond 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Seleção de Tipo de Imóvel"
   ClientHeight    =   2475
   ClientLeft      =   3195
   ClientTop       =   3465
   ClientWidth     =   5535
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2475
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtSubUnid 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   4230
      TabIndex        =   6
      Text            =   "0"
      Top             =   1350
      Width           =   765
   End
   Begin VB.TextBox txtUnid 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Enabled         =   0   'False
      Height          =   285
      Left            =   1530
      TabIndex        =   2
      Text            =   "0"
      Top             =   1350
      Width           =   765
   End
   Begin VB.ComboBox cmbCond 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   450
      Width           =   5325
   End
   Begin esMaskEdit.esMaskedEdit mskIC 
      Height          =   285
      Left            =   1500
      TabIndex        =   1
      Top             =   945
      Width           =   3945
      _ExtentX        =   6959
      _ExtentY        =   503
      MouseIcon       =   "frmSelCond.frx":0000
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
      MaxLength       =   18
      Mask            =   "#.##.####.#####.##"
      SelText         =   ""
      Text            =   "_.__.____._____.__"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   3060
      TabIndex        =   8
      ToolTipText     =   "Cancelar Edição"
      Top             =   1950
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmSelCond.frx":001C
      PICN            =   "frmSelCond.frx":0038
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdRetorna 
      Height          =   315
      Left            =   4140
      TabIndex        =   9
      ToolTipText     =   "Cadastra o Imóvel"
      Top             =   1950
      Width           =   1305
      _ExtentX        =   2302
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
      MICON           =   "frmSelCond.frx":0192
      PICN            =   "frmSelCond.frx":01AE
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da SubUnidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   0
      Left            =   2460
      TabIndex        =   7
      Top             =   1380
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Insc.Cadastral.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   1
      Left            =   150
      TabIndex        =   5
      Top             =   990
      Width           =   1305
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº da Unidade:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   225
      Index           =   7
      Left            =   150
      TabIndex        =   4
      Top             =   1380
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o Tipo de Imóvel"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   150
      TabIndex        =   3
      Top             =   150
      Width           =   3495
   End
End
Attribute VB_Name = "frmSelCond"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset
Dim Sql As String

Dim nDistrito As Integer
Dim nSetor As Integer
Dim nQuadra As Integer
Dim nLote As Long
Dim nSeq As Integer
Dim nUnidade As Integer
Dim nSubUnidade As Integer

Private Sub cmbCond_Click()

If cmbCond.ListIndex = 0 Then
     txtUnid.BackColor = Kde
     txtUnid.text = 0
     txtUnid.Enabled = False
     txtSubUnid.BackColor = Kde
     txtSubUnid.text = 0
     txtSubUnid.Enabled = False
     LimpaMascara mskIC
     mskIC.Enabled = True
Else
    Sql = "SELECT CD_DISTRITO,CD_SETOR,CD_QUADRA,CD_LOTE,CD_SEQ "
    Sql = Sql & "From CONDOMINIO "
    Sql = Sql & "WHERE CD_CODIGO=" & cmbCond.ItemData(cmbCond.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    With RdoAux
          nDistrito = !CD_DISTRITO
          nSetor = !CD_SETOR
          nQuadra = !CD_QUADRA
          nLote = !CD_LOTE
          nSeq = !CD_SEQ
          mskIC.text = nDistrito & "." & Format(nSetor, "00") & "." & Format(nQuadra, "0000") & "." & Format(nLote, "00000") & "." & Format(nSeq, "00")
         Sql = "SELECT MAX(UNIDADE) AS MAXIMO "
         Sql = Sql & "FROM CADIMOB "
         Sql = Sql & "WHERE DISTRITO=" & nDistrito & " AND "
         Sql = Sql & "SETOR=" & nSetor & " AND "
         Sql = Sql & "QUADRA=" & nQuadra & " AND "
         Sql = Sql & "LOTE=" & nLote & " AND "
         Sql = Sql & "SEQ=" & nSeq
         Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
          If RdoAux.RowCount > 0 Then
              If IsNull(RdoAux!MAXIMO) Then
                   txtUnid.text = 1
              Else
                   txtUnid.text = RdoAux!MAXIMO
              End If
          End If
    End With
    mskIC.Enabled = False
    txtUnid.BackColor = Branco
    txtUnid.Enabled = True
    txtSubUnid.BackColor = Branco
    txtSubUnid.Enabled = True
End If

End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdRetorna_Click()

nDistrito = Val(Left$(mskIC.text, 1))
nSetor = Val(Mid$(mskIC.text, 3, 2))
nQuadra = Val(Mid$(mskIC.text, 6, 4))
nLote = Val(Mid$(mskIC.text, 11, 5))
nSeq = Val(Mid$(mskIC.text, 17, 2))
nUnidade = Val(Mid$(mskIC.text, 20, 2))
nSubUnidade = Val(Mid$(mskIC.text, 23, 3))

If Len(mskIC.ClipText) < 14 Then
   MsgBox "Número cadastral incompleto.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
If Val(Left$(mskIC.text, 1)) = 0 Or Val(Left$(mskIC.text, 1)) > 3 Then
   MsgBox "Número de Distrito inválido.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
If nSetor > 0 Then
   Sql = "SELECT CODSETOR,CODSETOR FROM SETOR WHERE CODSETOR=" & nSetor & " AND CODDISTRITO=" & nDistrito
   Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   If RdoAux2.RowCount = 0 Then
      MsgBox "Setor não cadastrado ou não pertencente a este distrito.", vbExclamation, "Atenção"
      mskIC.SetFocus
      Exit Sub
   End If
   RdoAux2.Close
Else
   MsgBox "Número de Setor inválido.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
If nQuadra = 0 Then
   MsgBox "Número de Quadra inválida.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
If nLote = 0 Then
   MsgBox "Número de Lote inválido.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
If nSeq = 0 Then
   MsgBox "Número de Face inválido.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If
    
Sql = "SELECT CODREDUZIDO,DV FROM CADIMOB WHERE INATIVO=0 AND "
Sql = Sql & "DISTRITO=" & nDistrito & " AND "
Sql = Sql & "SETOR=" & nSetor & " AND "
Sql = Sql & "QUADRA=" & nQuadra & " AND "
Sql = Sql & "LOTE=" & nLote & " AND "
Sql = Sql & "SEQ=" & nSeq & " AND "
Sql = Sql & "UNIDADE=" & Val(txtUnid.text) & " AND "
Sql = Sql & "SUBUNIDADE=" & Val(txtSubUnid.text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    MsgBox "Número de Inscrição Cadastral já existente." & vbCrLf & "(Código Reduzido: " & Format(RdoAux!CODREDUZIDO, "00000") & ")", vbExclamation, "Atenção"
    Exit Sub
End If

Sql = "SELECT CODLOGR, CODAGRUPA From FACEQUADRA "
Sql = Sql & "WHERE CODDISTRITO=" & nDistrito & " AND "
Sql = Sql & "CODSETOR=" & nSetor & " AND "
Sql = Sql & "CODQUADRA=" & nQuadra & " AND "
Sql = Sql & "CODFACE=" & nSeq
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount = 0 Then
   MsgBox "Não existe Face de Quadra cadastrada para esta Inscrição Cadastral.", vbExclamation, "Atenção"
   mskIC.SetFocus
   Exit Sub
End If

If Val(txtUnid.text) > 0 Then
    Sql = "SELECT CD_NUMUNID From CONDOMINIO "
    Sql = Sql & "WHERE CD_CODIGO=" & cmbCond.ItemData(cmbCond.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If Val(txtUnid.text) > RdoAux!CD_NUMUNID Then
         MsgBox "Este condomínio possue apenas " & RdoAux!CD_NUMUNID & " Unidades.", vbExclamation, "Nº de Unidade inválido."
         txtUnid.SetFocus
         Exit Sub
    End If
End If

If Val(txtUnid.text) > 0 And Val(txtSubUnid.text) = 0 Then
         MsgBox "Digite o n de SubUnidades.", vbExclamation, "Nº de SubUnidade inválido."
         txtSubUnid.SetFocus
         Exit Sub
End If
    
If Val(txtUnid.text) > 0 Then
    Sql = "SELECT CD_SUBUNIDADES From CONDOMINIOUNIDADE "
    Sql = Sql & "WHERE CD_CODIGO=" & cmbCond.ItemData(cmbCond.ListIndex) & " AND "
    Sql = Sql & "CD_UNIDADE=" & Val(txtUnid.text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If Val(txtSubUnid.text) > RdoAux!CD_SUBUNIDADES Then
         MsgBox "Este condomínio possue apenas " & RdoAux!CD_SUBUNIDADES & " SubUnidades para esta Unidade.", vbExclamation, "Nº de SubUnidade inválido."
         txtSubUnid.SetFocus
         Exit Sub
    End If
End If

If cmbCond.ListIndex = 0 Then
    Sql = "SELECT CD_NOMECOND,CD_DISTRITO,CD_SETOR,CD_QUADRA,CD_LOTE,CD_SEQ "
    Sql = Sql & "From CONDOMINIO "
    Sql = Sql & "WHERE CD_DISTRITO=" & nDistrito & " AND CD_SETOR=" & nSetor & " AND CD_QUADRA=" & nQuadra & " AND CD_LOTE=" & nLote
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
    If RdoAux.RowCount > 0 Then
         MsgBox "Ja existe um Condominio com esta inscrição cadastral." & vbCrLf & "(" & RdoAux!CD_NOMECOND & ")", vbExclamation, "Atenção"
         Exit Sub
    End If
End If

With frmCadImob
    .lblIC.Caption = mskIC.text
    .lblSetor.Caption = Format(nSetor, "00")
    .txtQuadra.text = Format(nQuadra, "0000")
    .txtLote.text = Format(nLote, "00000")
    .txtSeq.text = Format(nSeq, "0")
    .lblUnid.Caption = Format(txtUnid.text, "00")
    .lblSubUnid.Caption = Format(txtSubUnid.text, "000")
    Sql = "SELECT CODDISTRITO,DESCDISTRITO FROM DISTRITO WHERE CODDISTRITO=" & nDistrito
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If RdoAux.RowCount > 0 Then
        .lblDist.Caption = Format(nDistrito, "00") & " - " & RdoAux!DescDistrito
    End If
    If cmbCond.ListIndex > 0 Then
         .lblCond.Caption = Format(cmbCond.ItemData(cmbCond.ListIndex), "0000") & " - " & cmbCond.text
   End If
End With
Unload Me
frmCadImob.SetFocus

End Sub

Private Sub Form_Load()
Centraliza Me

Sql = "SELECT CD_CODIGO, CD_NOMECOND "
Sql = Sql & "From CONDOMINIO WHERE CD_CODIGO<>999 ORDER BY CD_NOMECOND"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
cmbCond.AddItem "(Imóvel Normal)"
cmbCond.ItemData(cmbCond.NewIndex) = 999
With RdoAux
   Do Until .EOF
      cmbCond.AddItem !CD_NOMECOND
      cmbCond.ItemData(cmbCond.NewIndex) = !CD_CODIGO
     .MoveNext
   Loop
End With
If frmCadImob.lblCond.Caption = "Não Selecionado" Then
   cmbCond.ListIndex = 0
Else
   For x = 0 To cmbCond.ListCount - 1
       cmbCond.ListIndex = x
       If cmbCond.ItemData(cmbCond.ListIndex) = Val(Left$(frmCadImob.lblCond.Caption, 4)) Then
          Exit For
       End If
   Next
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
CodCond = cmbCond.ItemData(cmbCond.ListIndex)
NomeCond = cmbCond.text

End Sub


Private Sub mskIC_GotFocus()
mskIC.SelLength = Len(mskIC.text)
mskIC.SelStart = 0

bLibera = False
    
End Sub

Private Sub txtUnid_GotFocus()

txtUnid.SelStart = 0
txtUnid.SelLength = Len(txtUnid.text)

End Sub
