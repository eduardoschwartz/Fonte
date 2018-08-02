VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Object = "{DE8CE233-DD83-481D-844C-C07B96589D3A}#1.1#0"; "vbalSGrid6.ocx"
Begin VB.Form frmCPCnsProcesso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de processos (Compras)"
   ClientHeight    =   5700
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5700
   ScaleWidth      =   8145
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Campos de Pesquisa"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   2445
      Left            =   45
      TabIndex        =   1
      Top             =   45
      Width           =   8025
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   5535
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   20
         TabStop         =   0   'False
         Top             =   2025
         Visible         =   0   'False
         Width           =   765
      End
      Begin VB.TextBox txtAno2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5040
         MaxLength       =   4
         TabIndex        =   15
         Top             =   330
         Width           =   795
      End
      Begin VB.TextBox txtAno1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3180
         MaxLength       =   4
         TabIndex        =   14
         Top             =   330
         Width           =   795
      End
      Begin VB.TextBox txtNome 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         TabIndex        =   13
         Top             =   660
         Width           =   4545
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         TabIndex        =   12
         Top             =   1680
         Width           =   4125
      End
      Begin VB.ComboBox cmbLocal 
         Height          =   315
         Left            =   1290
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   990
         Width           =   4575
      End
      Begin VB.ComboBox cmbFisico 
         Height          =   315
         ItemData        =   "frmCPCnsProcesso.frx":0000
         Left            =   6930
         List            =   "frmCPCnsProcesso.frx":000D
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   630
         Width           =   1005
      End
      Begin VB.ComboBox cmbAssunto 
         Height          =   315
         Left            =   1290
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1350
         Width           =   6195
      End
      Begin VB.TextBox txtAssunto 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         TabIndex        =   7
         Top             =   1350
         Visible         =   0   'False
         Width           =   6165
      End
      Begin VB.ComboBox cmbInterno 
         Height          =   315
         ItemData        =   "frmCPCnsProcesso.frx":0024
         Left            =   6930
         List            =   "frmCPCnsProcesso.frx":0031
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   990
         Width           =   1005
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1290
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2025
         Width           =   3180
      End
      Begin VB.ListBox lstNomeLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00C0FFFF&
         Height          =   1590
         ItemData        =   "frmCPCnsProcesso.frx":0048
         Left            =   1305
         List            =   "frmCPCnsProcesso.frx":004A
         TabIndex        =   3
         Top             =   720
         Visible         =   0   'False
         Width           =   3705
      End
      Begin VB.TextBox txtNumImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5310
         TabIndex        =   2
         Top             =   2025
         Width           =   525
      End
      Begin prjChameleon.chameleonButton cmdNomeLogr 
         Height          =   270
         Left            =   4545
         TabIndex        =   4
         ToolTipText     =   "Pesquisar endere�o"
         Top             =   2025
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "..."
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCPCnsProcesso.frx":004C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskDataEntrada 
         Height          =   285
         Left            =   6930
         TabIndex        =   10
         Top             =   330
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         MouseIcon       =   "frmCPCnsProcesso.frx":0068
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
         Mask            =   "##/##/####"
         SelText         =   ""
         Text            =   "__/__/____"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdFiltrar 
         Default         =   -1  'True
         Height          =   315
         Left            =   5925
         TabIndex        =   16
         ToolTipText     =   "Consulta processos baseados no filtro selecionado"
         Top             =   1995
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Filtrar"
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
         MICON           =   "frmCPCnsProcesso.frx":0084
         PICN            =   "frmCPCnsProcesso.frx":00A0
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdLimpar 
         Height          =   315
         Left            =   6975
         TabIndex        =   17
         ToolTipText     =   "Limpar campos de pesquisa"
         Top             =   1995
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "Limpar"
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
         MICON           =   "frmCPCnsProcesso.frx":027A
         PICN            =   "frmCPCnsProcesso.frx":0296
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskNumProc 
         Height          =   285
         Left            =   1290
         TabIndex        =   18
         Top             =   330
         Width           =   915
         _ExtentX        =   1614
         _ExtentY        =   503
         MouseIcon       =   "frmCPCnsProcesso.frx":03F0
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
         MaxLength       =   8
         Mask            =   "######-#"
         SelText         =   ""
         Text            =   "______-_"
         HideSelection   =   -1  'True
      End
      Begin prjChameleon.chameleonButton cmdAbc 
         Height          =   270
         Left            =   7515
         TabIndex        =   19
         ToolTipText     =   "Alternar entre Lista e Texto"
         Top             =   1350
         Width           =   390
         _ExtentX        =   688
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "Abc"
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
         FCOL            =   192
         FCOLO           =   192
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCPCnsProcesso.frx":040C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Assunto..........:"
         Height          =   225
         Index           =   3
         Left            =   120
         TabIndex        =   32
         Top             =   1380
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano Final....:"
         Height          =   225
         Index           =   2
         Left            =   4080
         TabIndex        =   31
         Top             =   390
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ano Inicial.:"
         Height          =   225
         Index           =   1
         Left            =   2280
         TabIndex        =   30
         Top             =   390
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Requerente....:"
         Height          =   225
         Index           =   0
         Left            =   120
         TabIndex        =   29
         Top             =   690
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Complemento..:"
         Height          =   225
         Index           =   4
         Left            =   90
         TabIndex        =   28
         Top             =   1740
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N� Processo...:"
         Height          =   225
         Index           =   5
         Left            =   120
         TabIndex        =   27
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor/Depto...:"
         Height          =   225
         Index           =   7
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   1125
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Dt.Entrada..:"
         Height          =   225
         Index           =   9
         Left            =   5940
         TabIndex        =   25
         Top             =   360
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "F�sico.........:"
         Height          =   225
         Index           =   6
         Left            =   5940
         TabIndex        =   24
         Top             =   675
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Interno.......:"
         Height          =   225
         Index           =   8
         Left            =   5940
         TabIndex        =   23
         Top             =   1035
         Width           =   945
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "End.Ocorr�n....:"
         Height          =   225
         Index           =   10
         Left            =   90
         TabIndex        =   22
         Top             =   2100
         Width           =   1230
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "N�.:"
         Height          =   225
         Index           =   11
         Left            =   4995
         TabIndex        =   21
         Top             =   2070
         Width           =   375
      End
   End
   Begin MSComctlLib.ProgressBar PBar 
      Height          =   195
      Left            =   75
      TabIndex        =   0
      Top             =   5385
      Width           =   2175
      _ExtentX        =   3836
      _ExtentY        =   344
      _Version        =   393216
      Appearance      =   0
      Scrolling       =   1
   End
   Begin prjChameleon.chameleonButton cmdAbrir 
      Height          =   345
      Left            =   5685
      TabIndex        =   33
      ToolTipText     =   "Abrir processo selecionado"
      Top             =   5295
      Width           =   1155
      _ExtentX        =   2037
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
      MICON           =   "frmCPCnsProcesso.frx":0428
      PICN            =   "frmCPCnsProcesso.frx":0444
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   6885
      TabIndex        =   34
      ToolTipText     =   "Sair da Tela"
      Top             =   5295
      Width           =   1155
      _ExtentX        =   2037
      _ExtentY        =   609
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
      MICON           =   "frmCPCnsProcesso.frx":04CB
      PICN            =   "frmCPCnsProcesso.frx":04E7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin vbAcceleratorSGrid6.vbalGrid grdProc 
      Height          =   2715
      Left            =   75
      TabIndex        =   35
      Top             =   2505
      Width           =   7965
      _ExtentX        =   14049
      _ExtentY        =   4789
      BackgroundPictureHeight=   0
      BackgroundPictureWidth=   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      HeaderDragReorderColumns=   0   'False
      HeaderHotTrack  =   0   'False
      HeaderFlat      =   -1  'True
      BorderStyle     =   0
      ScrollBarStyle  =   1
      Editable        =   -1  'True
      DisableIcons    =   -1  'True
   End
   Begin VB.Label lblTot 
      BackStyle       =   0  'Transparent
      Caption         =   "0 processos localizados"
      Height          =   255
      Left            =   2505
      TabIndex        =   36
      Top             =   5385
      Width           =   2895
   End
End
Attribute VB_Name = "frmCPCnsProcesso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String

Private Sub cmdAbc_Click()
If txtAssunto.Visible = True Then
    txtAssunto.Text = ""
    txtAssunto.Visible = False
    cmbAssunto.Visible = True
    cmdAbc.Caption = "Abc"
Else
    cmbAssunto.ListIndex = 0
    cmbAssunto.Visible = False
    txtAssunto.Visible = True
    cmdAbc.Caption = "->"
End If
End Sub

Private Sub cmdAbrir_Click()
AnoProcessoCP = 0: CodProcessoCP = 0
If grdProc.SelectedRow > 0 Then
    AnoProcessoCP = grdProc.cell(grdProc.SelectedRow, 1).Text
    CodProcessoCP = Val(Left$(grdProc.cell(grdProc.SelectedRow, 2).Text, Len(grdProc.cell(grdProc.SelectedRow, 2).Text) - 1))
    frmCPProcesso.show
    frmCPProcesso.ZOrder 0
    Me.Hide
Else
    MsgBox "Selecione um processo.", vbExclamation, "Aten��o"
End If

End Sub

Private Sub cmdFiltrar_Click()
Dim nAno1 As Integer, nAno2 As Integer
Dim bNumProc As Boolean, bAno As Boolean, bReq As Boolean, bAss As Boolean, bCompl As Boolean, bCC As Boolean, bEntrada As Boolean, bFis As Boolean, bInt As Boolean, bLogr As Boolean, bNum As Boolean
Dim xId As Long, nNumRec As Long
If mskNumProc.ClipText <> "" Then
   If Right$(mskNumProc.ClipText, 1) <> RetornaDVProcesso(Left$(mskNumProc.ClipText, Len(mskNumProc.ClipText) - 1)) Then
      MsgBox "N�mero de processo inv�lido.", vbExclamation, "Aten��o"
      mskNumProc.SetFocus
      Exit Sub
   Else
      mskNumProc.Text = Format(mskNumProc.ClipText, "000000-0")
   End If
End If

If Val(txtAno1.Text) < 1950 And Val(txtAno1.Text) <> 0 Then
    MsgBox "Ano inicial invalido.", vbExclamation, "Aten��o"
    Exit Sub
End If

If Val(txtAno2.Text) < 1950 And Val(txtAno2.Text) <> 0 Then
    MsgBox "Ano final invalido.", vbExclamation, "Aten��o"
    Exit Sub
End If

If mskDataEntrada.ClipText <> "" Then
    If Not IsDate(mskDataEntrada.Text) Then
        MsgBox "Data de entrada inv�lida.", vbExclamation, "Aten��o"
        Exit Sub
    End If
End If

PBar.Value = 0
lblTot.Caption = "0 processos localizados."
grdProc.Clear
bNumProc = False: bAno = False: bReq = False: bAss = False: bCompl = False: bEntrada = False: bFis = False: bLogr = False

nAno1 = Val(txtAno1.Text): nAno2 = Val(txtAno2.Text)
If nAno1 = 0 Then nAno1 = 1920
If nAno2 = 0 Then nAno2 = Year(Now)

If Val(mskNumProc.Text) > 0 Then bNumProc = True
If Val(txtAno1.Text) > 0 Or Val(txtAno2.Text) > 0 Then bAno = True
If Trim$(txtNome.Text) <> "" Then bReq = True
If cmbAssunto.ListIndex > 0 Or txtAssunto.Text <> "" Then bAss = True
If cmbLocal.ListIndex > 0 Then bCC = True
If Trim$(txtCompl.Text) <> "" Then bCompl = True
If IsDate(mskDataEntrada.Text) Then bEntrada = True
If cmbFisico.ListIndex > 0 Then bFis = True
If cmbInterno.ListIndex > 0 Then bInt = True
If Val(txtCodLogr.Text) > 0 Then bLogr = True
If Val(txtNumImovel.Text) > 0 Then bNum = True

If bNumProc = False And bAno = False And bReq = False And bAss = False And bCompl = False And bCC = False And bEntrada = False And bFis = False And bInt = False And bLogr = False Then
    MsgBox "Selecione ao menos 1 crit�rio.", vbExclamation, "Aten��o"
    Exit Sub
End If

Ocupado
If cGetInputState() <> 0 Then DoEvents
Sql = "SELECT CPPROCESSOGTI.ANO, CPPROCESSOGTI.NUMERO, CPPROCESSOGTI.CODASSUNTO, cpassunto.NOME AS DESCASSUNTO, CPPROCESSOGTI.COMPLEMENTO, CPPROCESSOGTI.OBSERVACAO, "
Sql = Sql & "CPPROCESSOGTI.FISICO, CPPROCESSOGTI.INTERNO, CPPROCESSOGTI.DATAENTRADA, CPPROCESSOGTI.DATAREATIVA, CPPROCESSOGTI.DATACANCEL, CPPROCESSOGTI.DATAARQUIVA,"
Sql = Sql & "CPPROCESSOGTI.DATASUSPENSO, CPPROCESSOGTI.CODCIDADAO, CPPROCESSOGTI.RESPONSAVEL, cidadao.nomecidadao, CPPROCESSOGTI.CENTROCUSTO, cpcentrocusto.DESCRICAO,"
Sql = Sql & "CPPROCESSOEND.CODLOGR, CPPROCESSOEND.NUMERO AS NUMIMOVEL FROM CPPROCESSOGTI INNER JOIN cpassunto ON CPPROCESSOGTI.CODASSUNTO = cpassunto.CODIGO LEFT OUTER JOIN "
Sql = Sql & "CPPROCESSOEND ON CPPROCESSOGTI.ANO = CPPROCESSOEND.ANO AND CPPROCESSOGTI.NUMERO = CPPROCESSOEND.NUMPROCESSO LEFT OUTER JOIN cpcentrocusto ON CPPROCESSOGTI.CENTROCUSTO = cpcentrocusto.CODIGO LEFT OUTER JOIN "
Sql = Sql & "cidadao ON CPPROCESSOGTI.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "WHERE CPPROCESSOGTI.ANO BETWEEN " & nAno1 & " AND " & nAno2
If bNumProc Then
    Sql = Sql & " AND CPPROCESSOGTI.NUMERO=" & Val(Left$(mskNumProc.ClipText, Len(mskNumProc.ClipText) - 1))
End If
If bReq Then
    Sql = Sql & " AND NOMECIDADAO LIKE '%" & Mask(txtNome.Text) & "%'"
End If
If bAss Then
    If txtAssunto.Visible = True Then
        Sql = Sql & " AND CPASSUNTO.NOME LIKE '%" & txtAssunto.Text & "%' "
    Else
        Sql = Sql & " AND CPPROCESSOGTI.CODASSUNTO=" & cmbAssunto.ItemData(cmbAssunto.ListIndex)
    End If
End If
If bCompl Then
    Sql = Sql & " AND CPPROCESSOGTI.COMPLEMENTO LIKE '%" & Mask(txtCompl.Text) & "%'"
End If
If bCC Then
    Sql = Sql & " AND CPPROCESSOGTI.CENTROCUSTO=" & cmbLocal.ItemData(cmbLocal.ListIndex)
End If
If bEntrada Then
    Sql = Sql & " AND CPPROCESSOGTI.DATAENTRADA='" & Format(mskDataEntrada.Text, "mm/dd/yyyy") & "'"
End If
If bFis Then
    Sql = Sql & " AND CPPROCESSOGTI.FISICO=" & IIf(cmbFisico.ListIndex = 1, 1, 0)
End If
If bInt Then
    Sql = Sql & " AND CPPROCESSOGTI.INTERNO=" & IIf(cmbInterno.ListIndex = 1, 1, 0)
End If
If bLogr Then
    Sql = Sql & " AND CPPROCESSOEND.CODLOGR=" & Val(txtCodLogr.Text)
End If
If bNum Then
    Sql = Sql & " AND NUMIMOVEL=" & Val(txtNumImovel.Text)
End If

Sql = Sql & " ORDER BY CPPROCESSOGTI.ANO,CPPROCESSOGTI.NUMERO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
grdProc.Redraw = False
With RdoAux
    nNumRec = .RowCount
    xId = 1
    Do Until .EOF
        If xId Mod 10 = 0 Then
           CallPb xId, nNumRec
        End If
        grdProc.AddRow
        grdProc.CellDetails grdProc.Rows, 1, !Ano, DT_CENTER
        grdProc.CellDetails grdProc.Rows, 2, Format(!Numero & RetornaDVProcesso(CLng(!Numero)), "000000-0"), DT_CENTER
        grdProc.CellDetails grdProc.Rows, 3, IIf(IsNull(!NOMECIDADAO), SubNull(!DESCRICAO), !NOMECIDADAO), DT_LEFT
        grdProc.CellDetails grdProc.Rows, 4, !COMPLEMENTO, DT_LEFT
        grdProc.CellDetails grdProc.Rows, 5, Format(!DATAENTRADA, "dd/mm/yyyy"), DT_CENTER
        If Not IsNull(!DataCancel) Then
            grdProc.CellDetails grdProc.Rows, 6, Format(!DataCancel, "dd/mm/yyyy"), DT_CENTER
        Else
            grdProc.CellDetails grdProc.Rows, 6, "--------", DT_CENTER
        End If
        If Not IsNull(!DATAARQUIVA) Then
            grdProc.CellDetails grdProc.Rows, 7, Format(!DATAARQUIVA, "dd/mm/yyyy"), DT_CENTER
       Else
           grdProc.CellDetails grdProc.Rows, 7, "--------", DT_CENTER
       End If
       grdProc.CellDetails grdProc.Rows, 8, IIf(!FISICO, "S", "N"), DT_CENTER
       grdProc.CellDetails grdProc.Rows, 9, IIf(!INTERNO, "S", "N"), DT_CENTER
       xId = xId + 1
     .MoveNext
    Loop
   .Close
End With
CallPb xId, nNumRec
grdProc.Redraw = True
If grdProc.Rows = 0 Then MsgBox "Nenhum item coincidente.", vbInformation, "Aten��o"
lblTot.Caption = nNumRec & " processos localizados."
Liberado

End Sub

Private Sub cmdLimpar_Click()
LimpaMascara mskNumProc
txtAno1.Text = ""
txtAno2.Text = ""
LimpaMascara mskDataEntrada
txtNome.Text = ""
cmbLocal.ListIndex = 0
cmbAssunto.ListIndex = 0
txtCompl.Text = ""
txtCodLogr = "0"
txtNomeLogr.Text = ""
txtNumImovel.Text = ""
End Sub

Private Sub cmdNomeLogr_Click()
txtNomeLogr_KeyPress (vbKeyReturn)
End Sub

Private Sub cmdSair_Click()

Me.Hide
End Sub

Private Sub Form_Activate()
mskNumProc.SetFocus
End Sub

Private Sub Form_Load()
Centraliza Me
CarregaCombo
GridHeader
cmbFisico.ListIndex = 0: cmbInterno.ListIndex = 0

End Sub

Private Sub lvProc_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvProc.SortKey = ColumnHeader.Position - 1
lvProc.Sorted = True
lvProc.SortOrder = lvwAscending
End Sub

Private Sub grdProc_ColumnClick(ByVal lCol As Long)
Dim sTag As String
Dim iSortIndex As Long
      
   With grdProc.SortObject
      
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
      sTag = grdProc.ColumnTag(lCol)
      If (sTag = "") Then
         sTag = "DESC"
         .SortOrder(iSortIndex) = CCLOrderAscending
      Else
         sTag = ""
         .SortOrder(iSortIndex) = CCLOrderDescending
      End If
      grdProc.ColumnTag(lCol) = sTag
      
      ' Set the type of sorting:
      .SortType(iSortIndex) = grdProc.ColumnSortType(lCol)
   End With
   
   ' Do the sort:
   Screen.MousePointer = vbHourglass
   grdProc.Sort
   Screen.MousePointer = vbDefault

End Sub

Private Sub mskDataEntrada_GotFocus()
mskDataEntrada.SelStart = 0
mskDataEntrada.SelLength = Len(mskDataEntrada.Text)
End Sub

Private Sub mskNumProc_GotFocus()
mskNumProc.SelStart = 0
mskNumProc.SelLength = Len(mskNumProc.Text)
End Sub

Private Sub mskNumProc_LostFocus()
On Error Resume Next
If mskNumProc.ClipText <> "" Then
   If Right$(mskNumProc.ClipText, 1) <> RetornaDVProcesso(Left$(mskNumProc.ClipText, Len(mskNumProc.ClipText) - 1)) Then
      MsgBox "N�mero de processo inv�lido.", vbExclamation, "Aten��o"
      mskNumProc.SetFocus
   Else
      mskNumProc.Text = Format(mskNumProc.ClipText, "000000-0")
   End If
End If
End Sub

Private Sub txtAno1_KeyPress(KeyAscii As Integer)
Tweak txtAno1, KeyAscii, IntegerPositive
End Sub

Private Sub txtAno2_KeyPress(KeyAscii As Integer)
Tweak txtAno2, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaCombo()
Dim sText As String

cmbAssunto.Clear
cmbAssunto.AddItem "{Todos}"
Sql = "SELECT CODIGO,NOME FROM ASSUNTO ORDER BY NOME"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbAssunto.AddItem !NOME
       cmbAssunto.ItemData(cmbAssunto.NewIndex) = !Codigo
      .MoveNext
    Loop
    cmbAssunto.ListIndex = 0
   .Close
End With

cmbLocal.Clear
cmbLocal.AddItem "{Todos}"
Sql = "SELECT CODIGO,DESCRICAO FROM CPCENTROCUSTO WHERE SUBSTRING(DESCRICAO,1,1)<>'.'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       sText = !DESCRICAO
       cmbLocal.AddItem sText
       cmbLocal.ItemData(cmbLocal.NewIndex) = !Codigo
      .MoveNext
    Loop
    cmbLocal.ListIndex = 0
   .Close
End With

End Sub

Private Sub GridHeader()
With grdProc
    .GridFillLineColor = vbWhite
    .Editable = False
    .GridLines = True
    .HighlightBackColor = Marrom
    .HighlightForeColor = vbWhite
    .RowMode = True
    .DefaultRowHeight = 17
    .AddColumn "kAno", "Ano", ecgHdrTextALignCentre, , 40
    .AddColumn "kNum", "Numero", ecgHdrTextALignLeft, , 60
    .AddColumn "kReq", "Requerente", ecgHdrTextALignLeft, , 210
    .AddColumn "kAssu", "Assunto", ecgHdrTextALignLeft, , 200
    .AddColumn "kEnt", "Dt.Entrada", ecgHdrTextALignCentre, , 80
    .AddColumn "kCan", "Dt.Cancel", ecgHdrTextALignCentre, , 80
    .AddColumn "kArq", "Dt.Arquiva", ecgHdrTextALignCentre, , 80
    .AddColumn "kFis", "F�sico", ecgHdrTextALignCentre, , 60
    .AddColumn "kInt", "Interno", ecgHdrTextALignCentre, , 60
End With

End Sub

Private Sub CallPb(nPos As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents
If nTotal = 0 Then Exit Sub
If ((nPos * 100) / nTotal) <= 100 Then
   PBar.Value = (nPos * 100) / nTotal
Else
   PBar.Value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

Exit Sub
Erro:
MsgBox Err.Description
End Sub


Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.Text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr)
End Sub

Private Sub txtNomeLogr_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeLog.Clear
   If txtNomeLogr.Text <> "" Then
      Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
      Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
      Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO,DATAOFIC,"
      Sql = Sql & "NUMOFIC FROM vwLOGRADOURO "
      Sql = Sql & "WHERE NOMELOGRADOURO LIKE '%" & Trim$(txtNomeLogr) & "%' "
      Sql = Sql & "ORDER BY NOMELOGRADOURO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeLog.AddItem Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CODLOGRADOURO
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder (0)
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro n�o encontrado.", vbInformation, "Aten��o"
             lstNomeLog.Visible = False
             txtNomeLogr.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Text = 0
End If

End Sub

Private Sub lstNomeLog_DblClick()
If lstNomeLog.ListIndex > -1 Then
   txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
   txtCodLogr_LostFocus
   lstNomeLog.Visible = False
   txtNumImovel.SetFocus
End If

End Sub

Private Sub txtCodLogr_LostFocus()
If Val(txtCodLogr.Text) > 0 Then
   Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
   Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
   Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
   Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & Val(txtCodLogr.Text)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          txtNomeLogr.Text = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
       Else
          txtNomeLogr.Text = ""
          MsgBox "Logradouro n�o cadastrado.", vbExclamation, "Aten��o"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub


