VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmProcessoArquivado 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Processos Arquivados"
   ClientHeight    =   5175
   ClientLeft      =   5895
   ClientTop       =   3855
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5175
   ScaleWidth      =   8475
   Begin VB.Frame frPBar 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   2655
      TabIndex        =   23
      Top             =   2205
      Visible         =   0   'False
      Width           =   2715
      Begin Tributacao.XP_ProgressBar PBar 
         Height          =   195
         Left            =   45
         TabIndex        =   24
         Top             =   45
         Width           =   2625
         _ExtentX        =   4630
         _ExtentY        =   344
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BrushStyle      =   0
         Color           =   16750899
      End
   End
   Begin prjChameleon.chameleonButton cmdRefresh 
      Height          =   330
      Left            =   5400
      TabIndex        =   11
      ToolTipText     =   "Carregar processos do ano selecionado"
      Top             =   4770
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   582
      BTYPE           =   14
      TX              =   ""
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   15658734
      BCOLO           =   15658734
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmProcessoArquivado.frx":0000
      PICN            =   "frmProcessoArquivado.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      Left            =   4455
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   4770
      Width           =   915
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   7155
      TabIndex        =   13
      ToolTipText     =   "Cancelar Edição"
      Top             =   4770
      Width           =   1185
      _ExtentX        =   2090
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
      MICON           =   "frmProcessoArquivado.frx":012E
      PICN            =   "frmProcessoArquivado.frx":014A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   5895
      TabIndex        =   12
      ToolTipText     =   "Gravar os Dados"
      Top             =   4770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Gravar"
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
      MICON           =   "frmProcessoArquivado.frx":02A4
      PICN            =   "frmProcessoArquivado.frx":02C0
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdAlterar 
      Height          =   315
      Left            =   1350
      TabIndex        =   8
      ToolTipText     =   "Alterar a descrição do documento"
      Top             =   4770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Editar"
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
      MICON           =   "frmProcessoArquivado.frx":0665
      PICN            =   "frmProcessoArquivado.frx":0681
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   2610
      TabIndex        =   9
      ToolTipText     =   "Remover o documento selecionado"
      Top             =   4770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "E&xcluir"
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
      MICON           =   "frmProcessoArquivado.frx":07DB
      PICN            =   "frmProcessoArquivado.frx":07F7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdNovo 
      Height          =   315
      Left            =   90
      TabIndex        =   7
      ToolTipText     =   "Incluir novo documento"
      Top             =   4770
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Novo"
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
      MICON           =   "frmProcessoArquivado.frx":0899
      PICN            =   "frmProcessoArquivado.frx":08B5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   4650
      Left            =   0
      TabIndex        =   14
      Top             =   0
      Width           =   8430
      Begin vbAcceleratorSGrid6.vbalGrid grdMain 
         Height          =   4470
         Left            =   45
         TabIndex        =   6
         Top             =   135
         Width           =   8295
         _ExtentX        =   14631
         _ExtentY        =   7885
         NoHorizontalGridLines=   -1  'True
         BackgroundPictureHeight=   0
         BackgroundPictureWidth=   0
         BackColor       =   16777215
         HighlightForeColor=   8388608
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BorderStyle     =   0
         DisableIcons    =   -1  'True
         GroupBoxHintText=   "Arraste as colunas que deseja agrupar"
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Height          =   4650
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   8430
      Begin VB.TextBox txtProcesso 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1620
         TabIndex        =   0
         Top             =   225
         Width           =   1365
      End
      Begin VB.TextBox txtAnexo 
         Appearance      =   0  'Flat
         Height          =   870
         Left            =   135
         MaxLength       =   1000
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   3645
         Width           =   8115
      End
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   1500
         Left            =   135
         MaxLength       =   5000
         MultiLine       =   -1  'True
         TabIndex        =   4
         Top             =   1800
         Width           =   8115
      End
      Begin VB.ComboBox cmbAssunto 
         Height          =   315
         Left            =   1620
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   630
         Width           =   6630
      End
      Begin esMaskEdit.esMaskedEdit mskDataAbe 
         Height          =   285
         Left            =   1620
         TabIndex        =   2
         Top             =   1080
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmProcessoArquivado.frx":0A0F
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
      Begin esMaskEdit.esMaskedEdit mskDataArq 
         Height          =   285
         Left            =   5130
         TabIndex        =   3
         Top             =   1080
         Width           =   1140
         _ExtentX        =   2011
         _ExtentY        =   503
         MouseIcon       =   "frmProcessoArquivado.frx":0A2B
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
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Anexos..:"
         Height          =   240
         Index           =   4
         Left            =   225
         TabIndex        =   22
         Top             =   3375
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Observação..:"
         Height          =   240
         Index           =   3
         Left            =   180
         TabIndex        =   21
         Top             =   1530
         Width           =   1140
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Arquivado...:"
         Height          =   240
         Index           =   2
         Left            =   3735
         TabIndex        =   20
         Top             =   1125
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Abertura......:"
         Height          =   240
         Index           =   1
         Left            =   180
         TabIndex        =   19
         Top             =   1125
         Width           =   1320
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto...............:"
         Height          =   240
         Index           =   0
         Left            =   180
         TabIndex        =   18
         Top             =   675
         Width           =   1320
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Processo...:"
         Height          =   240
         Left            =   180
         TabIndex        =   17
         Top             =   270
         Width           =   1320
      End
   End
   Begin VB.Label lblAno 
      BackStyle       =   0  'Transparent
      Caption         =   "Ano..:"
      Height          =   240
      Left            =   3960
      TabIndex        =   15
      Top             =   4815
      Width           =   465
   End
End
Attribute VB_Name = "frmProcessoArquivado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String
Dim nAnoProc As Integer, nNumProc As Long, sNumProcesso As String

Private Sub cmdRefresh_Click()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long

Ocupado
grdMain.Redraw = False
grdMain.Clear
grdMain.Redraw = True

Sql = "SELECT processoarquivado.ano, processoarquivado.numero, processoarquivado.dataabertura, processoarquivado.dataarquiva, processoarquivado.assunto, "
Sql = Sql & "processoarquivado.obs , processoarquivado.anexo, assunto.NOME FROM processoarquivado INNER JOIN assunto ON processoarquivado.assunto = assunto.CODIGO"
Sql = Sql & " where ano=" & Val(cmbAno.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If cGetInputState() <> 0 Then DoEvents
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, Format(!Numero, "000000") & "/" & !Ano, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 2, !NOME
        grdMain.CellDetails grdMain.Rows, 3, Format(!DATAABERTURA, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 4, Format(!DATAARQUIVA, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 5, SubNull(!OBS)
        grdMain.CellDetails grdMain.Rows, 6, SubNull(!ANEXO)
        grdMain.CellDetails grdMain.Rows, 7, "N"
       .MoveNext
    Loop
   .Close
End With

frPBar.Visible = True
Me.Refresh
Sql = "SELECT processogti.ANO, processogti.NUMERO, processogti.DATAENTRADA, processogti.DATAARQUIVA, assunto.NOME, cidadao.nomecidadao, "
Sql = Sql & "centrocusto.DESCRICAO AS centrocusto FROM processogti INNER JOIN assunto ON processogti.CODASSUNTO = assunto.CODIGO LEFT OUTER JOIN "
Sql = Sql & "centrocusto ON processogti.CENTROCUSTO = centrocusto.CODIGO LEFT OUTER JOIN cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "Where (processogti.Ano = " & cmbAno.Text & ") And (processogti.DATAARQUIVA Is Not Null)"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount: nPos = 1
    grdMain.Redraw = False
    Do Until .EOF
        If nPos Mod 20 = 0 Then CallPb nPos, nTot
        If cGetInputState() <> 0 Then DoEvents
        grdMain.AddRow
        grdMain.CellDetails grdMain.Rows, 1, Format(!Numero, "000000") & "/" & !Ano, DT_CENTER
        grdMain.CellDetails grdMain.Rows, 2, !NOME
        grdMain.CellDetails grdMain.Rows, 3, Format(!DATAENTRADA, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 4, Format(!DATAARQUIVA, "dd/mm/yyyy"), DT_CENTER
        grdMain.CellDetails grdMain.Rows, 7, "S"
        grdMain.CellDetails grdMain.Rows, 8, IIf(IsNull(!NOMECIDADAO), SubNull(!CENTROCUSTO), !NOMECIDADAO)
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
grdMain.Redraw = True
frPBar.Visible = False
Me.Refresh

Liberado

End Sub

Private Sub Form_Load()
Dim x As Integer, Sql As String, RdoAux As rdoResultset
Centraliza Me
Eventos "INICIAR"
GridHeader
For x = 1970 To Year(Now)
    cmbAno.AddItem x
Next
cmbAno.ListIndex = cmbAno.ListCount - 1

Sql = "SELECT CODIGO,NOME,ATIVO FROM ASSUNTO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbAssunto.AddItem !NOME
        cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdAlterar_Click()
    Dim nRow As Integer
    
    With grdMain
        If .Rows < 1 Then
           MsgBox "Não existem Registros.", vbCritical, "Atenção"
           Exit Sub
        End If
        nRow = .SelectedRow
        If nRow = 0 Then nRow = 1
        If .CellText(nRow, 7) = "S" Then
            MsgBox "Apenas processos antigos podem ser editados por esta tela.", vbCritical, "Atenção"
        Else
            Eventos "INCLUIR"
            Evento = "Alterar"
            Le
            txtProcesso.Locked = True
            txtProcesso.BackColor = Kde
        End If
    End With
End Sub

Private Sub cmdCancel_Click()
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()

Dim nRow As Integer

With grdMain
    If .Rows < 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    nRow = .SelectedRow
    If nRow = 0 Then nRow = 1
    If .CellText(nRow, 7) = "S" Then
        MsgBox "Apenas processos antigos podem ser excluidos por esta tela.", vbCritical, "Atenção"
    Else
        sNumProcesso = .CellText(nRow, 1)
        nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
        nNumProc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))
        
        If MsgBox("Excluir este processo?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
            Sql = "DELETE FROM PROCESSOARQUIVADO WHERE ANO=" & nAno & " AND NUMERO=" & nNumProc
            cn.Execute Sql, rdExecDirect
            .RemoveRow (nRow)
        End If
    End If
End With

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset

sNumProcesso = txtProcesso.Text
If Trim$(sNumProcesso) = "" Then
    MsgBox "Nº de Processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If InStr(1, sNumProcesso, "/", vbBinaryCompare) = 0 Then
    MsgBox "Nº do processo inválido. Formato deve ser: Nº do Processo/Ano.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsNumeric(Right$(sNumProcesso, 4)) Then
    MsgBox "Nº do processo inválido. O ano deve ter 4 digitos.", vbCritical, "Atenção"
    Exit Sub
End If

If IsNumeric(Right$(sNumProcesso, 5)) Then
    MsgBox "Nº do processo inválido. O ano deve ter 4 digitos.", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsNumeric(Left$(sNumProcesso, 1)) Then
    MsgBox "Nº do processo inválido.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbAssunto.ListIndex = -1 Then
    MsgBox "Selecione o assunto", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataAbe.Text) Then
    MsgBox "Informe a data de abertura", vbCritical, "Atenção"
    Exit Sub
End If

If Not IsDate(mskDataArq.Text) Then
    MsgBox "Informe a data de arquivo", vbCritical, "Atenção"
    Exit Sub
End If

Grava
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    txtProcesso.Locked = False
    txtProcesso.BackColor = Branco
End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmProcessoArquivado
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   Frame2.Visible = False
   Frame1.Visible = True
   cmbAno.Visible = True
   lblAno.Visible = True
   cmdRefresh.Visible = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmProcessoArquivado
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Branco
          Ct.Enabled = True
       End If
   Next
   Frame2.Visible = True
   Frame1.Visible = False
   cmbAno.Visible = False
   lblAno.Visible = False
   cmdRefresh.Visible = False
End If

End Sub

Private Sub GridHeader()

With grdMain
    .HeaderFlat = True
    .HeaderHeight = 18
    .DefaultRowHeight = 17
    .GridFillLineColor = vbWhite
    .RowMode = True
    .GridLines = True
    .GridLineMode = ecgGridFillControl
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
        
    .AddColumn "NumProc", "Nº Processo", ecgHdrTextALignCentre, , 80
    .AddColumn "Assunto", "Assunto", ecgHdrTextALignLeft, , 290
    .AddColumn "DTABE", "Dt.Abertura", ecgHdrTextALignCentre, , 80
    .AddColumn "DTARQ", "Dt.Arquivo", ecgHdrTextALignCentre, , 80
    .AddColumn "Obs", "Observação", ecgHdrTextALignLeft, , 150
    .AddColumn "Anexo", "Anexo", ecgHdrTextALignLeft, , 150
    .AddColumn "Novo", "Novo", ecgHdrTextALignLeft, , 50
    .AddColumn "Requerente", "Requerente", ecgHdrTextALignLeft, , 250
    .ColumnVisible(5) = False
    .ColumnVisible(6) = False
    .ColumnVisible(7) = False
End With

End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC00000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.Value = (nVal * 100) / nTot
Else
   PBar.Value = 100
End If
Me.Refresh
If cGetInputState() <> 0 Then DoEvents
End Sub

Private Sub grdMain_ColumnClick(ByVal lCol As Long)

Dim sTag As String
Dim iSortIndex As Long
      
   With grdMain.SortObject
      
      ' This demo allows grouping.  When a column is clicked
      ' for sorting, we only want to remove any grouped rows:
      .ClearNongrouped
      
      ' See if this column is already in the sort object:
      iSortIndex = .IndexOf(lCol)
      If (iSortIndex = 0) Then
         ' If not, we add it:
         iSortIndex = .Count + 1
         .SortColumn(iSortIndex) = lCol
      End If
   
      ' Determine which sort order to apply:
      sTag = grdMain.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdMain.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdMain.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdMain.Sort
   Screen.MousePointer = vbDefault

End Sub

Private Sub Le()
Dim x As Integer

With grdMain
    txtProcesso.Text = .CellText(.SelectedRow, 1)
    mskDataAbe.Text = .CellText(.SelectedRow, 3)
    mskDataArq.Text = .CellText(.SelectedRow, 4)
    txtObs.Text = .CellText(.SelectedRow, 5)
    txtAnexo.Text = .CellText(.SelectedRow, 6)
    For x = 0 To cmbAssunto.ListCount - 1
        If cmbAssunto.List(x) = .CellText(.SelectedRow, 2) Then
            cmbAssunto.ListIndex = x
            Exit For
        End If
    Next
End With

End Sub

Private Sub Limpa()
txtProcesso.Text = ""
cmbAssunto.ListIndex = -1
LimpaMascara mskDataAbe
LimpaMascara mskDataArq
txtObs.Text = ""
txtAnexo.Text = ""
End Sub

Private Sub Grava()
Dim nRow As Integer
sNumProcesso = txtProcesso.Text
nAno = Val(Mid(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) + 1, 4))
nNumProc = Val(Left$(Trim$(sNumProcesso), InStr(1, Trim$(sNumProcesso), "/", vbBinaryCompare) - 1))

If Evento = "Novo" Then
    Sql = "INSERT PROCESSOARQUIVADO(ANO,NUMERO,DATAABERTURA,DATAARQUIVA,ASSUNTO,OBS,ANEXO) values("
    Sql = Sql & nAno & "," & nNumProc & ",'" & Format(mskDataAbe.Text, "mm/dd/yyyy") & "','"
    Sql = Sql & Format(mskDataArq.Text, "mm/dd/yyyy") & "'," & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ",'"
    Sql = Sql & Mask(txtObs.Text) & "','" & Mask(txtAnexo.Text) & "')"
Else
    With grdMain
        nRow = .SelectedRow
        .CellDetails nRow, 2, cmbAssunto.Text
        .CellDetails nRow, 3, mskDataAbe.Text, DT_CENTER
        .CellDetails nRow, 4, mskDataArq.Text, DT_CENTER
        .CellDetails nRow, 5, txtObs.Text
        .CellDetails nRow, 6, txtAnexo.Text
    End With
    Sql = "UPDATE PROCESSOARQUIVADO SET DATAABERTURA='" & Format(mskDataAbe.Text, "mm/dd/yyyy") & "',DATAARQUIVA='" & Format(mskDataArq.Text, "mm/dd/yyyy") & "',"
    Sql = Sql & "ASSUNTO=" & cmbAssunto.ItemData(cmbAssunto.ListIndex) & ",OBS='" & Mask(txtObs.Text) & "',ANEXO='" & Mask(txtAnexo.Text) & "'"
    Sql = Sql & " WHERE ANO = " & nAno & " AND NUMERO=" & nNumProc
End If
cn.Execute Sql, rdExecDirect

End Sub
