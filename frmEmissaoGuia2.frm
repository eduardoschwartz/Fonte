VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmEmissaoGuia2 
   Appearance      =   0  'Flat
   BackColor       =   &H00FBFBE3&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Composição da guia"
   ClientHeight    =   4680
   ClientLeft      =   10905
   ClientTop       =   6765
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   9075
   ShowInTaskbar   =   0   'False
   Begin Tributacao.jcFrames jcFrames4 
      Height          =   1995
      Left            =   4290
      Top             =   1920
      Width           =   4725
      _ExtentX        =   8334
      _ExtentY        =   3519
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Observação"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   1665
         Left            =   60
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   270
         Width           =   4605
      End
   End
   Begin Tributacao.jcFrames jcFrames3 
      Height          =   2715
      Left            =   30
      Top             =   1920
      Width           =   4215
      _ExtentX        =   7435
      _ExtentY        =   4789
      FrameColor      =   12829635
      Style           =   0
      Caption         =   "Composição do lançamento"
      TextColor       =   128
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin prjChameleon.chameleonButton cmdQtde 
         Height          =   240
         Left            =   3210
         TabIndex        =   12
         ToolTipText     =   "Altera a Qtde do Tributo"
         Top             =   0
         Width           =   630
         _ExtentX        =   1111
         _ExtentY        =   423
         BTYPE           =   3
         TX              =   "Qtde"
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
         MICON           =   "frmEmissaoGuia2.frx":0000
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin MSComctlLib.ListView lvTrib 
         Height          =   2340
         Left            =   60
         TabIndex        =   11
         Top             =   300
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   4128
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   16777215
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Desc.Tributo"
            Object.Width           =   4047
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Qtde"
            Object.Width           =   1059
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Valor"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Descrição Completa"
            Object.Width           =   6068
         EndProperty
      End
   End
   Begin Tributacao.jcFrames jcFrames2 
      Height          =   885
      Left            =   30
      Top             =   1020
      Width           =   8985
      _ExtentX        =   15849
      _ExtentY        =   1561
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   0
      ColorTo         =   0
      Begin VB.ComboBox cmbTabela 
         Height          =   315
         ItemData        =   "frmEmissaoGuia2.frx":001C
         Left            =   7950
         List            =   "frmEmissaoGuia2.frx":001E
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   120
         Width           =   975
      End
      Begin VB.CheckBox chkUnica 
         Appearance      =   0  'Flat
         BackColor       =   &H00FBFBE3&
         Caption         =   "Parcela única"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5115
         TabIndex        =   6
         Top             =   165
         Width           =   1620
      End
      Begin VB.TextBox txtQtdeParc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   1215
         MaxLength       =   6
         TabIndex        =   3
         Top             =   120
         Width           =   585
      End
      Begin VB.ComboBox cmbAnoTabela 
         Height          =   315
         ItemData        =   "frmEmissaoGuia2.frx":0020
         Left            =   7950
         List            =   "frmEmissaoGuia2.frx":0022
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox txtAbate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6000
         TabIndex        =   9
         Text            =   "0,00"
         Top             =   480
         Width           =   885
      End
      Begin prjChameleon.chameleonButton cmdAddData 
         Height          =   270
         Left            =   4365
         TabIndex        =   5
         ToolTipText     =   "Editar Datas de Vencimento"
         Top             =   135
         Width           =   345
         _ExtentX        =   609
         _ExtentY        =   476
         BTYPE           =   3
         TX              =   "+"
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   12582912
         FCOLO           =   12582912
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmEmissaoGuia2.frx":0024
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin esMaskEdit.esMaskedEdit mskDataVencimento 
         Height          =   285
         Left            =   3225
         TabIndex        =   4
         Top             =   120
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   16777215
         MouseIcon       =   "frmEmissaoGuia2.frx":0040
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
      Begin esMaskEdit.esMaskedEdit mskDataInicio 
         Height          =   285
         Left            =   3225
         TabIndex        =   8
         Top             =   480
         Width           =   1095
         _ExtentX        =   1931
         _ExtentY        =   503
         BackColor       =   16777215
         MouseIcon       =   "frmEmissaoGuia2.frx":005C
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
         Locked          =   -1  'True
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Data 1º Vencto.:"
         Height          =   225
         Index           =   15
         Left            =   1965
         TabIndex        =   27
         Top             =   165
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Qtde Parc.....:"
         Height          =   225
         Index           =   14
         Left            =   90
         TabIndex        =   26
         Top             =   165
         Width           =   1065
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Exercício.:"
         Height          =   225
         Index           =   17
         Left            =   7080
         TabIndex        =   25
         Top             =   180
         Width           =   825
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Cálculo proporcional a partir da data de.....:"
         Height          =   225
         Index           =   16
         Left            =   90
         TabIndex        =   24
         Top             =   525
         Width           =   3105
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tabela.....:"
         Height          =   240
         Index           =   0
         Left            =   7080
         TabIndex        =   23
         Top             =   525
         Width           =   825
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Abatimento em NF:"
         ForeColor       =   &H00000000&
         Height          =   210
         Left            =   4500
         TabIndex        =   22
         Top             =   525
         Width           =   1425
      End
   End
   Begin Tributacao.jcFrames jcFrames1 
      Height          =   915
      Left            =   60
      Top             =   60
      Width           =   8955
      _ExtentX        =   15796
      _ExtentY        =   1614
      FrameColor      =   12829635
      Style           =   0
      RoundedCornerTxtBox=   -1  'True
      Caption         =   ""
      TextColor       =   13579779
      Alignment       =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColorFrom       =   16514019
      ColorTo         =   0
      Begin VB.ComboBox cmbTipoGuia 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   90
         Width           =   2745
      End
      Begin VB.ComboBox cmbLanc 
         Height          =   315
         Left            =   1230
         Style           =   2  'Dropdown List
         TabIndex        =   1
         Top             =   480
         Width           =   4905
      End
      Begin VB.ListBox lstAtividade 
         Appearance      =   0  'Flat
         Height          =   705
         Left            =   6210
         Style           =   1  'Checkbox
         TabIndex        =   2
         Top             =   90
         Width           =   2655
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo de guia..:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   150
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Lançamento..:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   20
         Top             =   540
         Width           =   1095
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade..:"
         Height          =   195
         Index           =   2
         Left            =   5310
         TabIndex        =   19
         Top             =   150
         Width           =   855
      End
   End
   Begin prjChameleon.chameleonButton cmdBaixa 
      Height          =   345
      Left            =   7830
      TabIndex        =   14
      ToolTipText     =   "Gera as guias informadas"
      Top             =   4200
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "&Próximo"
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
      MICON           =   "frmEmissaoGuia2.frx":0078
      PICN            =   "frmEmissaoGuia2.frx":0094
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   1
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Única:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4320
      TabIndex        =   18
      Top             =   4110
      Width           =   1005
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5385
      TabIndex        =   17
      Top             =   4110
      Width           =   975
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Total.:"
      ForeColor       =   &H00800000&
      Height          =   210
      Left            =   4320
      TabIndex        =   16
      Top             =   4380
      Width           =   1005
   End
   Begin VB.Label lblTotal 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "0,00"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   5385
      TabIndex        =   15
      Top             =   4380
      Width           =   975
   End
End
Attribute VB_Name = "frmEmissaoGuia2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmbTipoGuia_Click()

If cmbTipoGuia.ListIndex = -1 Then Exit Sub

If cmbTipoGuia.ListIndex = 1 Then
    CarregaLancamento 6
ElseIf cmbTipoGuia.ListIndex = 2 Then
    CarregaLancamento 14
ElseIf cmbTipoGuia.ListIndex = 3 Then
    CarregaLancamento 13
Else
    CarregaLancamento 0
End If

End Sub

Private Sub Form_Activate()
txtQtdeParc.SetFocus
End Sub

Private Sub Form_Load()
Dim x As Integer

Me.Top = frmEmissaoGuia.Top + 2500
Me.Left = frmEmissaoGuia.Left + 1000

cmbTipoGuia.AddItem "(Lançamentos diversos)"
cmbTipoGuia.AddItem "Taxa de Licença"
cmbTipoGuia.AddItem "ISS Fixo"
cmbTipoGuia.AddItem "Vigilância Sanitária"
cmbTipoGuia.ListIndex = 0

For x = 1994 To Year(Now)
    cmbTabela.AddItem x
Next
cmbTabela.Text = Year(Now)

For x = 2011 To Year(Now)
    cmbAnoTabela.AddItem x
Next
cmbAnoTabela.Text = Year(Now)

End Sub

Private Sub CarregaLancamento(nCodigo As Integer)
Dim Sql As String, RdoAux As rdoResultset

cmbLanc.Clear
Sql = "select codlancamento, descreduz from lancamento "
If nCodigo = 6 Or nCodigo = 13 Or nCodigo = 14 Then
    Sql = Sql & "where codlancamento=" & nCodigo
Else
    Sql = Sql & "where codlancamento not in (2,3,5,6,8,12,13,14,20,21,30) "
End If
Sql = Sql & "order by descreduz"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbLanc.AddItem !descreduz
        cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With

If cmbLanc.ListCount > 0 Then cmbLanc.ListIndex = 0

End Sub

Private Sub mskDataInicio_GotFocus()
mskDataInicio.SelStart = 0
mskDataInicio.SelLength = Len(mskDataInicio.Text)

End Sub

Private Sub mskDataVencimento_GotFocus()
mskDataVencimento.SelStart = 0
mskDataVencimento.SelLength = Len(mskDataVencimento.Text)
End Sub

Private Sub txtQtdeParc_GotFocus()
txtQtdeParc.SelStart = 0
txtQtdeParc.SelLength = Len(txtQtdeParc.Text)

End Sub

Private Sub txtQtdeParc_KeyPress(KeyAscii As Integer)
Tweak txtQtdeParc, KeyAscii, IntegerPositive
End Sub
