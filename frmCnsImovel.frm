VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmCnsImovel 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consulta de Imóvel"
   ClientHeight    =   6180
   ClientLeft      =   2190
   ClientTop       =   2280
   ClientWidth     =   8775
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6180
   ScaleWidth      =   8775
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Selecione um ou mais Parâmetros"
      ForeColor       =   &H00000080&
      Height          =   3045
      Left            =   10
      TabIndex        =   12
      Top             =   -15
      Width           =   8745
      Begin VB.TextBox txtNomeBairro 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2115
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1980
         Width           =   2820
      End
      Begin VB.TextBox txtCompl 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   40
         TabIndex        =   9
         Top             =   2670
         Width           =   5670
      End
      Begin VB.ComboBox cmbCond 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2130
         TabIndex        =   8
         Top             =   2310
         Width           =   5685
      End
      Begin VB.TextBox txtLote 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7080
         TabIndex        =   7
         Top             =   1980
         Width           =   705
      End
      Begin VB.TextBox txtQuadra 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5760
         TabIndex        =   6
         Top             =   1980
         Width           =   735
      End
      Begin esMaskEdit.esMaskedEdit mskNumInsc 
         Height          =   300
         Left            =   2130
         TabIndex        =   1
         Top             =   615
         Width           =   4185
         _ExtentX        =   7382
         _ExtentY        =   529
         MouseIcon       =   "frmCnsImovel.frx":0000
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
         MaxLength       =   25
         Mask            =   "9.99.9999.99999.99.99.999"
         SelText         =   ""
         Text            =   "_.__.____._____.__.__.___"
         HideSelection   =   -1  'True
      End
      Begin VB.TextBox txtCodLogr 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   6360
         Locked          =   -1  'True
         MaxLength       =   50
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1290
         Width           =   765
      End
      Begin VB.TextBox txtNomeLogr 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1260
         Width           =   4170
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         MaxLength       =   9
         TabIndex        =   0
         Top             =   300
         Width           =   1035
      End
      Begin VB.TextBox txtNumImovel 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2130
         TabIndex        =   3
         Top             =   1620
         Width           =   705
      End
      Begin VB.TextBox txtProp 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   2130
         Locked          =   -1  'True
         TabIndex        =   10
         Top             =   960
         Width           =   4170
      End
      Begin prjChameleon.chameleonButton cmdBuscaProp 
         Height          =   285
         Left            =   6360
         TabIndex        =   18
         ToolTipText     =   "Busca Proprietário"
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsImovel.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdDelProp 
         Height          =   285
         Left            =   6720
         TabIndex        =   19
         ToolTipText     =   "Limpa Campo Proprietário"
         Top             =   960
         Width           =   315
         _ExtentX        =   556
         _ExtentY        =   503
         BTYPE           =   14
         TX              =   "X"
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmCnsImovel.frx":0038
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdConsultar 
         Height          =   315
         Left            =   7365
         TabIndex        =   26
         ToolTipText     =   "Seleciona o Imóvel"
         Top             =   720
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         BTYPE           =   3
         TX              =   "&Selecionar"
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
         MICON           =   "frmCnsImovel.frx":0054
         PICN            =   "frmCnsImovel.frx":0070
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPesq 
         Height          =   345
         Left            =   7380
         TabIndex        =   27
         ToolTipText     =   "Pesquisar"
         Top             =   315
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   "C&onsultar"
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
         MICON           =   "frmCnsImovel.frx":00DE
         PICN            =   "frmCnsImovel.frx":00FA
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdCancel 
         Cancel          =   -1  'True
         Height          =   315
         Left            =   7365
         TabIndex        =   28
         ToolTipText     =   "Cancelar Edição"
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
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
         MICON           =   "frmCnsImovel.frx":0254
         PICN            =   "frmCnsImovel.frx":0270
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
         Left            =   7380
         TabIndex        =   29
         ToolTipText     =   "Limpar campos de pesquisa"
         Top             =   1440
         Width           =   1200
         _ExtentX        =   2117
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
         MICON           =   "frmCnsImovel.frx":03CA
         PICN            =   "frmCnsImovel.frx":03E6
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.ListBox lstNomeLog 
         BackColor       =   &H00C0FFFF&
         Height          =   1620
         ItemData        =   "frmCnsImovel.frx":0540
         Left            =   2130
         List            =   "frmCnsImovel.frx":0542
         TabIndex        =   20
         Top             =   300
         Visible         =   0   'False
         Width           =   5010
      End
      Begin VB.ComboBox cmbBairro 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   2385
         TabIndex        =   5
         Top             =   1950
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.ListBox lstNomeBairro 
         BackColor       =   &H00C0FFFF&
         Height          =   1620
         ItemData        =   "frmCnsImovel.frx":0544
         Left            =   2115
         List            =   "frmCnsImovel.frx":054B
         TabIndex        =   31
         Top             =   1350
         Visible         =   0   'False
         Width           =   2850
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço complementar..:"
         Height          =   225
         Index           =   9
         Left            =   135
         TabIndex        =   30
         Top             =   2700
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Condomínio......:"
         Height          =   225
         Index           =   8
         Left            =   150
         TabIndex        =   25
         Top             =   2370
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         Height          =   225
         Index           =   7
         Left            =   6630
         TabIndex        =   24
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra:"
         Height          =   225
         Index           =   6
         Left            =   5100
         TabIndex        =   23
         Top             =   2010
         Width           =   615
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro/Loteamento...........:"
         Height          =   225
         Index           =   3
         Left            =   150
         TabIndex        =   22
         Top             =   2010
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Imóvel....................:"
         Height          =   225
         Index           =   2
         Left            =   150
         TabIndex        =   17
         Top             =   1650
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código Reduzido.............:"
         Height          =   225
         Index           =   0
         Left            =   150
         TabIndex        =   16
         Top             =   330
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome do Logradouro.......:"
         Height          =   225
         Index           =   1
         Left            =   150
         TabIndex        =   15
         Top             =   1320
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário/Compromis....:"
         Height          =   225
         Index           =   4
         Left            =   150
         TabIndex        =   14
         Top             =   990
         Width           =   1905
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Incrição Cadastral..:"
         Height          =   225
         Index           =   5
         Left            =   150
         TabIndex        =   13
         Top             =   675
         Width           =   1905
      End
   End
   Begin MSComctlLib.ListView lvImovel 
      Height          =   3075
      Left            =   0
      TabIndex        =   11
      Top             =   3105
      Width           =   8745
      _ExtentX        =   15425
      _ExtentY        =   5424
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12640511
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   11
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Código"
         Object.Width           =   1693
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Alignment       =   2
         SubItemIndex    =   1
         Text            =   "ATIVO"
         Object.Width           =   1252
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Nº de Insc."
         Object.Width           =   3881
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Proprietário Principal"
         Object.Width           =   4233
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "CPF/CNPJ/RG"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Nome do Logradouro"
         Object.Width           =   4304
      EndProperty
      BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   6
         Text            =   "Nº Log."
         Object.Width           =   1411
      EndProperty
      BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   7
         Text            =   "Complemento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(9) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   8
         Text            =   "Bairro/Loteamento"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(10) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   9
         Text            =   "Quadra"
         Object.Width           =   1305
      EndProperty
      BeginProperty ColumnHeader(11) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   10
         Text            =   "Lote"
         Object.Width           =   1305
      EndProperty
   End
End
Attribute VB_Name = "frmCnsImovel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String, bCod As Boolean, bDist As Boolean, bSetor As Boolean
Dim bQuadra As Boolean, bLote As Boolean, bSeq As Boolean, bUnidade As Boolean, bSubUnidade As Boolean
Dim bNome As Boolean, bLog As Boolean, bNum As Boolean, bBairro As Boolean, bQuadras As Boolean, bLotes As Boolean, bCond As Boolean, bCompl As Boolean

Private Sub cmdBuscaProp_Click()

Dim frm As Object
Set frm = frmCnsCidadao
frm.sForm = Me.Name
frmCnsCidadao.show

End Sub

Private Sub cmdCancel_Click()
   'Unload frmCnsImovel
   frmCnsImovel.Hide
End Sub

Private Sub cmdConsultar_Click()
Dim Achou As Boolean

If lvImovel.ListItems.Count > 0 Then
   CodImovel = lvImovel.SelectedItem.Text
   modLg "Consulta de imóvel: " & Val(Left$(CodImovel, 7)) & " - " & lvImovel.SelectedItem.SubItems(3)
   If sForm = "CI" Then
        Achou = False
        For X = 0 To Forms.Count - 1
            If Forms(X).Name = "frmCadImob" Then
                 Achou = True
                 Exit For
            End If
        Next
        If Not Achou Then
             frmCadImob.show
             frmCadImob.ZOrder 0
             frmCadImob.SetFocus
        End If
   ElseIf sForm = "DD" Then
        frmDesmembra.SetFocus
   ElseIf sForm = "UN" Then
        frmUnifica.SetFocus
   ElseIf sForm = "DI" Then
        frmDebitoImob.SetFocus
   ElseIf sForm = "CR" Then
        frmCancelReparc.SetFocus
   ElseIf sForm = "EI" Then
        frm2ViaLaser.SetFocus
   ElseIf sForm = "ALUGUEL" Then
        frmManAluguel.SetFocus
   ElseIf sForm = "2VIA" Then
        frmEmissao2Via.SetFocus
   ElseIf sForm = "2VIAE" Then
        frmEmissao2ViaEspecial.SetFocus
   ElseIf sForm = "frmRequerIPTU" Then
        frmRequerIPTU.SetFocus
   ElseIf sForm = "frmDeclaraIsento" Then
        frmDeclaraIsento.SetFocus
   End If
   'Unload frmCnsImovel
   frmCnsImovel.Hide
Else
   MsgBox "Selecione o imóvel que deseja consultar.", vbExclamation, "Atenção"
End If

End Sub


Private Sub cmdDelProp_Click()
txtProp.Text = ""
End Sub

Private Sub cmdLimpar_Click()
txtCod.Text = ""
LimpaMascara mskNumInsc
txtProp.Text = ""
txtNomeLogr.Text = ""
txtCodLogr.Text = ""
txtNumImovel.Text = ""
cmbBairro.ListIndex = -1
txtQuadra.Text = ""
txtLote.Text = ""
cmbCond.ListIndex = -1
End Sub

Private Sub cmdPesq_Click()

Dim itmX As ListItem
Dim z As Long, sCep As String

If txtCod.Text = "" And txtProp.Text = "" And Val(txtCodLogr.Text) = 0 And Val(txtNumImovel.Text) = 0 And mskNumInsc.ClipText = "" And txtNomeBairro.Tag = "" And txtQuadra.Text = "" And txtLote.Text = "" And cmbCond.ListIndex < 1 And txtCompl.Text = "" Then
    MsgBox "Favor selecionar ao menos um critério para busca.", vbExclamation, "Atenção"
    Exit Sub
End If

Screen.MousePointer = vbHourglass
Ocupado

z = SendMessage(lvImovel.hwnd, LVM_DELETEALLITEMS, 0, 0)

bCod = False
bDist = False
bSetor = False
bQuadra = False
bLote = False
bSeq = False
bUnidade = False
bSubUnidade = False
bNome = False
bLog = False
bNum = False
bBairro = False
bQuadras = False
bLotes = False
bCond = False
bCompl = False

If Val(txtCod.Text) > 0 Then bCod = True
If Val(Left$(mskNumInsc.Text, 1)) > 0 Then bDist = True
If Val(Mid$(mskNumInsc.Text, 3, 2)) > 0 Then bSetor = True
If Val(Mid$(mskNumInsc.Text, 6, 4)) > 0 Then bQuadra = True
If Val(Mid$(mskNumInsc.Text, 11, 5)) > 0 Then bLote = True
If Val(Mid$(mskNumInsc.Text, 17, 2)) > 0 Then bSeq = True
If Val(Mid$(mskNumInsc.Text, 20, 2)) > 0 Then bUnidade = True
If Val(Mid$(mskNumInsc.Text, 23, 3)) > 0 Then bSubUnidade = True
If txtProp.Text <> "" Then bNome = True
If Val(txtCodLogr.Text) > 0 Then bLog = True
If txtNumImovel.Text <> "" Then bNum = True
If txtNomeBairro.Tag <> "" And txtNomeBairro.Tag <> "0" Then bBairro = True
If txtQuadra.Text <> "" Then bQuadras = True
If txtLote.Text <> "" Then bLotes = True
If cmbCond.ListIndex > 0 Then bCond = True
If txtCompl.Text <> "" Then bCompl = True

Sql = "SELECT     CADIMOB.CODREDUZIDO,INATIVO, CADIMOB.DV, CADIMOB.DISTRITO, CADIMOB.SETOR,FACEQUADRA.CODLOGR, CADIMOB.QUADRA, CADIMOB.LOTE, CADIMOB.SEQ, "
Sql = Sql & "                      CADIMOB.LI_CODBAIRRO,LI_COMPL, CADIMOB.LI_QUADRAS, CADIMOB.LI_LOTES, CADIMOB.UNIDADE, CADIMOB.SUBUNIDADE, PROPRIETARIO.CODCIDADAO, "
Sql = Sql & "                      CIDADAO.NOMECIDADAO, CIDADAO.RG, CIDADAO.ORGAO, CIDADAO.CPF, CIDADAO.CNPJ, CADIMOB.LI_NUM,CADIMOB.LI_COMPL, LOGRADOURO.NOMELOGRADOURO, "
Sql = Sql & "                      TIPOLOGRADOURO.AbrevTipoLog , TITLOGRADOURO.AbrevTitLog, BAIRRO.DescBairro "
Sql = Sql & " FROM         BAIRRO RIGHT OUTER JOIN "
Sql = Sql & "                      CADIMOB ON BAIRRO.SIGLAUF = CADIMOB.LI_UF AND BAIRRO.CODCIDADE = CADIMOB.LI_CODCIDADE AND "
Sql = Sql & "                      BAIRRO.CODBAIRRO = CADIMOB.LI_CODBAIRRO LEFT OUTER JOIN "
Sql = Sql & "                      TITLOGRADOURO RIGHT OUTER JOIN "
Sql = Sql & "                      FACEQUADRA INNER JOIN "
Sql = Sql & "                      LOGRADOURO ON FACEQUADRA.CODLOGR = LOGRADOURO.CODLOGRADOURO LEFT OUTER JOIN "
Sql = Sql & "                      TIPOLOGRADOURO ON LOGRADOURO.CODTIPOLOG = TIPOLOGRADOURO.CODTIPOLOG ON "
Sql = Sql & "                      TITLOGRADOURO.CODTITLOG = LOGRADOURO.CODTITLOG ON CADIMOB.DISTRITO = FACEQUADRA.CODDISTRITO AND "
Sql = Sql & "                      CADIMOB.SETOR = FACEQUADRA.CODSETOR AND CADIMOB.QUADRA = FACEQUADRA.CODQUADRA AND "
Sql = Sql & "                      CADIMOB.SEQ = FACEQUADRA.CODFACE FULL OUTER JOIN "
Sql = Sql & "                      PROPRIETARIO ON CADIMOB.CODREDUZIDO = PROPRIETARIO.CODREDUZIDO FULL OUTER JOIN "
Sql = Sql & "                      CIDADAO ON PROPRIETARIO.CODCIDADAO = CIDADAO.CODCIDADAO "
Sql = Sql & "WHERE  "
Sql = Sql & "PROPRIETARIO.TIPOPROP='P' AND PRINCIPAL=1 AND "
If bCod Then
   Sql = Sql & "CADIMOB.CODREDUZIDO=" & Val(txtCod.Text) & " AND "
End If
If bDist Then
   Sql = Sql & "CADIMOB.DISTRITO=" & Val(Left$(mskNumInsc.Text, 1)) & " AND "
End If
If bSetor Then
   Sql = Sql & "CADIMOB.SETOR=" & Val(Mid$(mskNumInsc.Text, 3, 2)) & " AND "
End If
If bQuadra Then
   Sql = Sql & "CADIMOB.QUADRA=" & Val(Mid$(mskNumInsc.Text, 6, 4)) & " AND "
End If
If bLote Then
   Sql = Sql & "CADIMOB.LOTE=" & Val(Mid$(mskNumInsc.Text, 11, 5)) & " AND "
End If
If bSeq Then
   Sql = Sql & "CADIMOB.SEQ=" & Val(Mid$(mskNumInsc.Text, 17, 2)) & " AND "
End If
If bUnidade Then
   Sql = Sql & "CADIMOB.UNIDADE=" & Val(Mid$(mskNumInsc.Text, 20, 2)) & " AND "
End If
If bSubUnidade Then
   Sql = Sql & "CADIMOB.SUBUNIDADE=" & Val(Mid$(mskNumInsc.Text, 23, 3)) & " AND "
End If
If bNum Then
   Sql = Sql & "CADIMOB.LI_NUM=" & Val(txtNumImovel.Text) & " AND "
End If
If bCompl Then
   Sql = Sql & "CADIMOB.LI_COMPL='" & Mask(txtCompl.Text) & "' AND "
End If
If bNome Then
    Sql = Sql & "PROPRIETARIO.CODCIDADAO=" & Val(Left$(txtProp.Text, 6)) & "  AND  "
   'Sql = Sql & "CIDADAO.NOMECIDADAO LIKE '" & Mid(txtProp.text, 10, Len(txtProp.text) - 9) & "%' AND PROPRIETARIO.TIPOPROP='P' AND  "
End If
If bLog Then
   Sql = Sql & "FACEQUADRA.CODLOGR=" & Val(txtCodLogr.Text) & " AND "
End If
If bBairro Then
   Sql = Sql & "CADIMOB.LI_CODBAIRRO=" & Val(txtNomeBairro.Tag) & " AND "
End If
If bQuadras Then
   Sql = Sql & "LTRIM(CADIMOB.LI_QUADRAS)='" & Mask(txtQuadra.Text) & "' AND "
End If
If bLotes Then
   Sql = Sql & "CADIMOB.LI_LOTES='" & txtLote.Text & "' AND "
End If
If bCond Then
   Sql = Sql & "CADIMOB.CODCONDOMINIO='" & cmbCond.ItemData(cmbCond.ListIndex) & "' AND "
End If
Sql = Left$(Sql, Len(Sql) - 5)

On Error Resume Next
RdoAux.Close
On Error GoTo 0

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurReadOnly)
If RdoAux.RowCount > 0 Then
    If RdoAux.RowCount > 5000 Then
        MsgBox "A consulta retornou muitos dados e não pode ser carregada, cancelando consulta!", vbCritical, "ERRO"
        Liberado
        Exit Sub
    End If
   With RdoAux
'   Open sPathBin & "\JARDIMSTOANTONIO.TXT" For Output As #1
       Do Until .EOF
          
          If Not IsNull(!CODREDUZIDO) Then
          Set itmX = lvImovel.ListItems.Add(, "C" & Format(!CODREDUZIDO, "0000000") & Format(!CodCidadao, "0000000"), Format(!CODREDUZIDO, "0000000") & "-" & CStr(!DV))
          itmX.SubItems(1) = IIf(!Inativo, "N", "S")
          itmX.SubItems(2) = !Distrito & "." & Format(!Setor, "00") & "." & Format(!Quadra, "0000") & "." & Format(!Lote, "00000") & "." & Format(!Seq, "00") & "." & Format(!Unidade, "00") & "." & Format(!SubUnidade, "000")
          itmX.SubItems(3) = SubNull(!nomecidadao)
          If Not IsNull(!CPF) Then
             itmX.SubItems(4) = SubNull(!CPF)
          ElseIf Not IsNull(!Cnpj) Then
             itmX.SubItems(4) = SubNull(!Cnpj)
          ElseIf Not IsNull(!rg) Then
             itmX.SubItems(4) = SubNull(!rg)
          Else
             itmX.SubItems(4) = ""
          End If
          itmX.SubItems(5) = Trim$(SubNull(!AbrevTipoLog)) & " " & IIf(IsNull(!AbrevTitLog), "", Trim$(SubNull(!AbrevTitLog)) & " ") & Trim$(SubNull(!NomeLogradouro))
          itmX.SubItems(6) = SubNull(!Li_Num)
          itmX.SubItems(7) = SubNull(!Li_Compl)
          
          sCep = RetornaCEP(!CodLogr, !Li_Num)
          'itmX.SubItems(8) = RetornaBairro(RetornaNumero(sCep)).Nome
          
          itmX.SubItems(8) = SubNull(!DescBairro)
          itmX.SubItems(9) = SubNull(!Li_Quadras)
          itmX.SubItems(10) = SubNull(!Li_Lotes)
         
            
 '               sLayout = "00"
  '              ax = !CODREDUZIDO & "," & !NOMECIDADAO & "," & itmX.SubItems(5) & "," & itmX.SubItems(6) & "," & itmX.SubItems(9) & "," & itmX.SubItems(10)
   '             Print #1, ax
            
         
         
         .MoveNext
         End If
       Loop
    '   Close #1
      .Close
   End With
Else
   MsgBox "Não existem imóveis com estes parâmetros.", vbExclamation, "Atenção"
End If
Liberado
Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
Ocupado
CodImovel = ""

Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO INNER JOIN CIDADE ON BAIRRO.SIGLAUF = CIDADE.SIGLAUF AND BAIRRO.CODCIDADE = CIDADE.CODCIDADE INNER JOIN UF ON CIDADE.SIGLAUF = UF.SIGLAUF WHERE (UF.SIGLAUF = 'SP') AND (DESCCIDADE = 'JABOTICABAL') AND (CODBAIRRO <> 999) ORDER BY DESCBAIRRO; "
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
   cmbBairro.AddItem ""
    Do Until .EOF
       cmbBairro.AddItem !DescBairro
       cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
      .MoveNext
    Loop
   .Close
End With

cmbCond.AddItem ""
cmbCond.ItemData(cmbCond.NewIndex) = 999
Sql = "SELECT CD_CODIGO,CD_NOMECOND FROM CONDOMINIO  WHERE CD_CODIGO<>999  ORDER BY CD_NOMECOND"
Set RdoAux = cn.OpenResultset(Sql, rdOpenForwardOnly, rdConcurReadOnly)
With RdoAux
    Do Until .EOF
       cmbCond.AddItem !cd_nomecond
       cmbCond.ItemData(cmbCond.NewIndex) = !CD_CODIGO
      .MoveNext
    Loop
   .Close
End With
cmbCond.ListIndex = 0
Centraliza Me
Liberado
End Sub

Private Sub lstNomeBairro_DblClick()
If lstNomeBairro.ListIndex > -1 Then
    txtNomeBairro.Text = lstNomeBairro.Text
   txtNomeBairro.Tag = lstNomeBairro.ItemData(lstNomeBairro.ListIndex)
   lstNomeBairro.Visible = False
   txtQuadra.SetFocus
End If

End Sub

Private Sub lstNomeBairro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeBairro.ListIndex > -1 Then
       txtNomeBairro.Text = lstNomeBairro.Text
       txtNomeBairro.Tag = lstNomeBairro.ItemData(lstNomeBairro.ListIndex)
       lstNomeBairro.Visible = False
       txtQuadra.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeBairro.Visible = False
   txtNomeBairro.SetFocus
End If

End Sub

Private Sub lstNomeBairro_LostFocus()
lstNomeBairro.Visible = False
End Sub

Private Sub lstNomeLog_LostFocus()
lstNomeLog.Visible = False
End Sub

Private Sub lvImovel_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
lvImovel.SortKey = ColumnHeader.Position - 1
lvImovel.Sorted = True
lvImovel.SortOrder = lvwAscending
End Sub

Private Sub lvImovel_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = vbRightButton Then
    If MsgBox("Deseja gerar arquivo texto ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
        GeraTexto
    End If
End If
End Sub

Private Sub GeraTexto()
Dim fName As String, sTexto As String

fName = App.Path & "\ListaImovel.txt"
    
    ' The function we need for saving.
    Dim FileId As Integer
    Dim X As Integer
    Dim sIdx As Integer
    sIdx = lvImovel.ColumnHeaders.Count - 1
    FileId = FreeFile
    On Error Resume Next
    Open fName For Output As #FileId
    For i = 1 To lvImovel.ListItems.Count
        sTexto = "'" & lvImovel.ListItems.Item(i).Text & "',"
        For X = 1 To sIdx
            sTexto = sTexto & lvImovel.ListItems.Item(i).SubItems(X) & "','"
        Next
        sTexto = Left$(sTexto, Len(sTexto) - 2)
        Print #FileId, sTexto
    Next
    Close #FileId
    
MsgBox "Gerado Arquivo: " & App.Path & "\ListaImovel.txt", vbInformation, "Informação"
    
End Sub

Private Sub txtCod_GotFocus()
txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)
End Sub

Private Sub mskNumInsc_GotFocus()
mskNumInsc.SelStart = 0
mskNumInsc.SelLength = Len(mskNumInsc.Text)
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   cmdPesq_Click
   cmdConsultar_Click
Else
   Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtNomeBairro_Change()
If Trim$(txtNomeBairro) = "" Then
   txtNomeBairro.Tag = "0"
End If

End Sub

Private Sub txtNomeBairro_GotFocus()
txtNomeBairro.SelStart = 0
txtNomeBairro.SelLength = Len(txtNomeBairro.Text)

End Sub

Private Sub txtNomeBairro_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   lstNomeBairro.Clear
   If txtNomeBairro.Text <> "" Then
      Sql = "SELECT CODBAIRRO,DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='SP' AND CODCIDADE=413 AND DESCBAIRRO LIKE '%" & Trim$(txtNomeBairro) & "%' "
      Sql = Sql & "ORDER BY DESCBAIRRO"
      Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
      With RdoAux
          If .RowCount > 0 Then
             Do Until .EOF
                lstNomeBairro.AddItem !DescBairro
                lstNomeBairro.ItemData(lstNomeBairro.NewIndex) = !CodBairro
               .MoveNext
             Loop
             lstNomeBairro.Visible = True
             lstNomeBairro.ZOrder (0)
             lstNomeBairro.ListIndex = 0
             lstNomeBairro.SetFocus
          Else
             MsgBox "Bairro não encontrado.", vbInformation, "Atenção"
             lstNomeBairro.Visible = False
             txtNomeBairro.SetFocus
          End If
      End With
   End If
Else
   txtCodLogr.Tag = "0"
End If

End Sub

Private Sub txtNomeLogr_Change()
If Trim$(txtNomeLogr) = "" Then
   txtCodLogr.Text = 0
End If
End Sub

Private Sub txtNomeLogr_GotFocus()
txtNomeLogr.SelStart = 0
txtNomeLogr.SelLength = Len(txtNomeLogr.Text)
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
                lstNomeLog.ItemData(lstNomeLog.NewIndex) = !CodLogradouro
               .MoveNext
             Loop
             lstNomeLog.Visible = True
             lstNomeLog.ZOrder (0)
             lstNomeLog.ListIndex = 0
             lstNomeLog.SetFocus
          Else
             MsgBox "Logradouro não encontrado.", vbInformation, "Atenção"
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

Private Sub lstNomeLog_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    If lstNomeLog.ListIndex > -1 Then
       txtCodLogr.Text = lstNomeLog.ItemData(lstNomeLog.ListIndex)
       txtCodLogr_LostFocus
       lstNomeLog.Visible = False
       txtNumImovel.SetFocus
    End If
ElseIf KeyAscii = vbKeyEscape Then
   lstNomeLog.Visible = False
   txtNomeLogr.SetFocus
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
          MsgBox "Logradouro não cadastrado.", vbExclamation, "Atenção"
          txtCodLogr.SetFocus
       End If
      .Close
   End With
End If

End Sub

