VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmUnificacao 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Unificação de Imóveis"
   ClientHeight    =   5475
   ClientLeft      =   3765
   ClientTop       =   3390
   ClientWidth     =   9075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5475
   ScaleWidth      =   9075
   Begin VB.Frame pnlWait 
      BackColor       =   &H000000C0&
      BorderStyle     =   0  'None
      Height          =   825
      Left            =   1380
      TabIndex        =   48
      Top             =   2400
      Visible         =   0   'False
      Width           =   6795
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Gravando Dados...Aguarde !!!"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   345
         Left            =   0
         TabIndex        =   49
         Top             =   240
         Width           =   6705
      End
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Height          =   315
      Left            =   120
      TabIndex        =   15
      ToolTipText     =   "Executar Unificação"
      Top             =   4710
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Unificar"
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
      MICON           =   "frmUnificacao.frx":0000
      PICN            =   "frmUnificacao.frx":001C
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
      Left            =   120
      TabIndex        =   16
      ToolTipText     =   "Sair da Tela"
      Top             =   5070
      Width           =   1245
      _ExtentX        =   2196
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
      MICON           =   "frmUnificacao.frx":00BB
      PICN            =   "frmUnificacao.frx":00D7
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
      Height          =   1725
      Left            =   1770
      TabIndex        =   46
      Top             =   3720
      Width           =   7275
      Begin MSComctlLib.ListView lvArea 
         Height          =   1545
         Left            =   60
         TabIndex        =   47
         Top             =   120
         Width           =   7155
         _ExtentX        =   12621
         _ExtentY        =   2725
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         HotTracking     =   -1  'True
         _Version        =   393217
         ForeColor       =   8388672
         BackColor       =   16777215
         Appearance      =   0
         NumItems        =   8
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Lote"
            Object.Width           =   1765
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Seq"
            Object.Width           =   882
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Tp"
            Object.Width           =   742
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Área"
            Object.Width           =   1744
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Ano"
            Object.Width           =   1942
         EndProperty
         BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   5
            Text            =   "Uso"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(7) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   6
            Text            =   "Tipo"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(8) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   7
            Text            =   "Categoria"
            Object.Width           =   2540
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTestada 
      Height          =   2145
      Left            =   120
      TabIndex        =   44
      Top             =   2460
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   3784
      _Version        =   393216
      Rows            =   1
      FixedCols       =   0
      BackColorBkg    =   15658734
      Appearance      =   0
      FormatString    =   "^Face   |> Metros     "
   End
   Begin VB.Frame Pnl2 
      BackColor       =   &H00EEEEEE&
      Height          =   2175
      Left            =   1770
      TabIndex        =   33
      Top             =   1530
      Width           =   7275
      Begin VB.ComboBox cmbCat 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   570
         Width           =   2175
      End
      Begin VB.ComboBox cmbSit 
         Height          =   315
         Left            =   5010
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   960
         Width           =   2175
      End
      Begin VB.ComboBox cmbUso 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   600
         Width           =   2325
      End
      Begin VB.ComboBox cmbBen 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   990
         Width           =   2325
      End
      Begin VB.ComboBox cmbTop 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   1380
         Width           =   2325
      End
      Begin VB.ComboBox cmbPed 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   1770
         Width           =   2325
      End
      Begin VB.TextBox txtFracao 
         Height          =   315
         Left            =   5010
         TabIndex        =   14
         Text            =   "0,00"
         Top             =   1380
         Width           =   1575
      End
      Begin VB.ComboBox cmbProp 
         Height          =   315
         Left            =   1350
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   210
         Width           =   5835
      End
      Begin VB.Label lblAreaTerreno 
         BackStyle       =   0  'Transparent
         Caption         =   "0,00"
         ForeColor       =   &H00000080&
         Height          =   165
         Left            =   5040
         TabIndex        =   43
         Top             =   1830
         Width           =   1275
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Categ.Propr....:"
         Height          =   195
         Left            =   3780
         TabIndex        =   42
         Top             =   690
         Width           =   1095
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Situação..........:"
         Height          =   195
         Left            =   3780
         TabIndex        =   41
         Top             =   1065
         Width           =   1185
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Área Terreno...:"
         Height          =   195
         Left            =   3780
         TabIndex        =   40
         Top             =   1830
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Uso do Terreno:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   39
         Top             =   690
         Width           =   1185
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Benfeitoria........:"
         Height          =   195
         Left            =   90
         TabIndex        =   38
         Top             =   1065
         Width           =   1185
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Topografia........:"
         Height          =   195
         Left            =   90
         TabIndex        =   37
         Top             =   1455
         Width           =   1185
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Pedologia.........:"
         Height          =   195
         Left            =   90
         TabIndex        =   36
         Top             =   1830
         Width           =   1185
      End
      Begin VB.Label Label15 
         BackStyle       =   0  'Transparent
         Caption         =   "Fração Ideal....:"
         Height          =   195
         Left            =   3780
         TabIndex        =   35
         Top             =   1455
         Width           =   1185
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Proprietário......:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   34
         Top             =   300
         Width           =   1185
      End
   End
   Begin VB.Frame Pnl1 
      BackColor       =   &H00EEEEEE&
      Height          =   1515
      Left            =   1770
      TabIndex        =   18
      Top             =   0
      Width           =   7275
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1020
         Width           =   4455
      End
      Begin VB.ComboBox cmbNumero 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   630
         Width           =   1035
      End
      Begin VB.ComboBox cmbQuadra 
         Height          =   315
         Left            =   2580
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   180
         Width           =   1035
      End
      Begin VB.ComboBox cmbFace 
         Height          =   315
         Left            =   6120
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   180
         Width           =   1035
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº.:"
         Height          =   195
         Left            =   5610
         TabIndex        =   32
         Top             =   690
         Width           =   405
      End
      Begin VB.Label lblEnd 
         BackStyle       =   0  'Transparent
         Caption         =   "00765 - AVENIDA JARDIM BOTANICO"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   900
         TabIndex        =   31
         Top             =   690
         Width           =   4755
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro...:"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   30
         Top             =   1080
         Width           =   705
      End
      Begin VB.Label lblDist 
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   690
         TabIndex        =   29
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label17 
         BackStyle       =   0  'Transparent
         Caption         =   "Distrito:"
         Height          =   195
         Left            =   90
         TabIndex        =   28
         Top             =   240
         Width           =   525
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Setor:"
         Height          =   195
         Left            =   1020
         TabIndex        =   27
         Top             =   240
         Width           =   435
      End
      Begin VB.Label lblSetor 
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   1530
         TabIndex        =   26
         Top             =   240
         Width           =   225
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Quadra:"
         Height          =   195
         Left            =   1920
         TabIndex        =   25
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Lote:"
         Height          =   165
         Left            =   3900
         TabIndex        =   24
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Face:"
         Height          =   195
         Left            =   5610
         TabIndex        =   23
         Top             =   240
         Width           =   435
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Cep:"
         Height          =   225
         Left            =   5610
         TabIndex        =   22
         Top             =   1050
         Width           =   435
      End
      Begin VB.Label lblCEP 
         Alignment       =   1  'Right Justify
         BackStyle       =   0  'Transparent
         Caption         =   "14870-000"
         ForeColor       =   &H00000080&
         Height          =   255
         Left            =   6120
         TabIndex        =   21
         Top             =   1050
         Width           =   765
      End
      Begin VB.Label lblLote 
         BackStyle       =   0  'Transparent
         Caption         =   "01"
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   4410
         TabIndex        =   20
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Rua.....:"
         Height          =   195
         Index           =   1
         Left            =   90
         TabIndex        =   19
         Top             =   690
         Width           =   735
      End
   End
   Begin VB.ListBox lstImovel 
      Appearance      =   0  'Flat
      Height          =   1395
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   1575
   End
   Begin prjChameleon.chameleonButton cmdAddImovel 
      Height          =   285
      Left            =   930
      TabIndex        =   1
      ToolTipText     =   "Adicionar Imóvel"
      Top             =   1800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
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
      MICON           =   "frmUnificacao.frx":0145
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDelImovel 
      Height          =   285
      Left            =   1290
      TabIndex        =   2
      ToolTipText     =   "Remover Imóvel"
      Top             =   1800
      Width           =   315
      _ExtentX        =   556
      _ExtentY        =   503
      BTYPE           =   3
      TX              =   "-"
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
      MICON           =   "frmUnificacao.frx":0161
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
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      Caption         =   "Testadas"
      ForeColor       =   &H00FFFFFF&
      Height          =   225
      Left            =   120
      TabIndex        =   45
      Top             =   2220
      Width           =   1545
   End
   Begin VB.Label Label27 
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione os Imóveis:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   90
      TabIndex        =   17
      Top             =   120
      Width           =   1665
   End
End
Attribute VB_Name = "frmUnificacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type tImovel
    nCodReduz As Long
    nDistrito As Integer
    nSetor As Integer
    nQuadra As Integer
    nLote As Integer
    nSeq As Integer
    nNumero As Integer
    nCodBairro As Integer
    sDescBairro As String
    nCodUsoTerreno As Integer
    sDescUsoTerreno As String
    nCodBenfeitoria As Integer
    sDescBenfeitoria As String
    nCodTopografia As Integer
    sDescTopografia As String
    nCodCatProp As Integer
    sDescCatProp As String
    nCodSituacao As Integer
    sDescSituacao As String
    nCodPedologia As Integer
    sDescPedologia As String
    nAreaTerreno As Double
    nCodProprietario As Long
    sNomeProprietario As String
End Type

Dim aImovel() As tImovel
Dim RdoAux As rdoResultset, Sql As String
Dim nCodReduz As Long

Private Sub cmbFace_Click()

If cmbFace.ListIndex = -1 Then
    lblEnd.Caption = ""
Else
    Sql = "SELECT facequadra.codlogr, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG, vwLOGRADOURO.NOMELOGRADOURO "
    Sql = Sql & "FROM facequadra INNER JOIN vwLOGRADOURO ON facequadra.codlogr = vwLOGRADOURO.CODLOGRADOURO WHERE "
    Sql = Sql & "CODDISTRITO=" & Val(lblDist.Caption) & " AND CODSETOR=" & Val(lblSetor.Caption) & " AND "
    Sql = Sql & "CODQUADRA=" & Val(cmbQuadra.Text) & " AND CODFACE=" & Val(cmbFace.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            lblEnd.Caption = Format(!CodLogr, "0000") & " - " & Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
        Else
            MsgBox "Logradouro não cadastrado para esta face.", vbCritical, "Atenção"
        End If
       .Close
    End With
    If cmbNumero.ListIndex > -1 Then
        cmbNumero_Click
    End If
End If

End Sub

Private Sub cmbNumero_Click()
If lblEnd.Caption <> "" And cmbNumero.ListIndex > -1 Then
    lblCep.Caption = RetornaCEP(CLng(Left(lblEnd.Caption, 4)), Val(cmbNumero.Text))
End If
End Sub

Private Sub cmbQuadra_Click()

lblEnd.Caption = "": lblCep.Caption = ""
If cmbQuadra.ListIndex = -1 Then
    lblLote.Caption = "00000"
Else
    Sql = "SELECT MAX(LOTE) AS ULTIMOLOTE FROM CADIMOB WHERE "
    Sql = Sql & "DISTRITO=" & Val(lblDist.Caption) & " AND SETOR=" & Val(lblSetor.Caption) & " AND QUADRA=" & Val(cmbQuadra.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    lblLote.Caption = Format(RdoAux!ULTIMOLOTE + 1, "00000")
    RdoAux.Close
End If

cmbFace.Clear
Sql = "SELECT CODFACE FROM FACEQUADRA WHERE CODDISTRITO=" & Val(lblDist.Caption) & " AND CODSETOR=" & Val(lblSetor.Caption) & " AND CODQUADRA=" & Val(cmbQuadra.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbFace.AddItem Format(!CODFACE, "000")
       .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub cmdAddImovel_Click()
Dim z As Variant, x As Integer, bAchou As Boolean

z = InputBox("Digite o código do imóvel.", "Adicionar Imóveis")
If Val(z) = 0 Then Exit Sub
If Val(z) < 100000 And IsNumeric(z) Then
    For x = 0 To lstImovel.ListCount - 1
        If lstImovel.List(x) = Val(z) Then
            MsgBox "Imóvel já incluido na lista.", vbExclamation, "Atenção"
            Exit Sub
        End If
    Next
    If AdicionaImovel(CLng(z)) Then
        lstImovel.AddItem Format(Val(z), "000000")
    End If
Else
    MsgBox "Código inválido.", vbCritical, "Atenção"
End If

End Sub

Private Function AdicionaImovel(nCodImovel As Long) As Boolean
Dim RdoAux2 As rdoResultset


AdicionaImovel = False
Sql = "SELECT cadimob.codreduzido, cadimob.distrito, cadimob.setor, cadimob.quadra, cadimob.lote, cadimob.seq, cadimob.inativo, cadimob.li_num,"
Sql = Sql & "cadimob.li_codbairro, bairro.descbairro,dt_areaterreno, cadimob.dt_codusoterreno, usoterreno.descusoterreno, cadimob.dt_codbenf, benfeitoria.descbenfeitoria,"
Sql = Sql & "cadimob.dt_codtopog, topografia.desctopografia, cadimob.dt_codcategprop, categprop.desccategprop, cadimob.dt_codsituacao, situacao.descsituacao,"
Sql = Sql & "cadimob.Dt_CodPedol , pedologia.DescPedologia, cadimob.Dt_FracaoIdeal , cidadao.codcidadao,Cidadao.nomecidadao FROM cadimob INNER JOIN bairro ON cadimob.li_uf = bairro.siglauf AND "
Sql = Sql & "cadimob.li_codcidade = bairro.codcidade AND cadimob.li_codbairro = bairro.codbairro INNER JOIN usoterreno ON cadimob.dt_codusoterreno = usoterreno.codusoterreno INNER JOIN "
Sql = Sql & "benfeitoria ON cadimob.dt_codbenf = benfeitoria.codbenfeitoria INNER JOIN topografia ON cadimob.dt_codtopog = topografia.codtopografia INNER JOIN "
Sql = Sql & "pedologia ON cadimob.dt_codpedol = pedologia.codpedologia INNER JOIN categprop ON cadimob.dt_codcategprop = categprop.codcategprop INNER JOIN "
Sql = Sql & "situacao ON cadimob.dt_codsituacao = situacao.codsituacao INNER JOIN proprietario ON cadimob.codreduzido = proprietario.codreduzido INNER JOIN "
Sql = Sql & "cidadao ON proprietario.codcidadao = cidadao.codcidadao Where cadimob.CODREDUZIDO = " & nCodImovel
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Imóvel não cadastrado.", vbCritical, "Atenção"
    Else
        If !Inativo = 1 Then
            MsgBox "Este imóvel esta inativo.", vbCritical, "Atenção"
        Else
            If lstImovel.ListCount > 0 Then
                bAchou = False
                For x = 1 To UBound(aImovel)
                    If !Distrito <> aImovel(x).nDistrito Or !Setor <> aImovel(x).nSetor Then
                        bAchou = True
                        Exit For
                    End If
                Next
            End If
            If bAchou Then
                MsgBox "Os imóveis tem que pertencer ao mesmo distrito e setor.", vbCritical, "Atenção"
            Else
                Sql = "SELECT * FROM DEBITOAUTOMATICO WHERE CODREDUZ=" & nCodImovel
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    If .RowCount > 0 Then
                        MsgBox "Este imóvel não pode ser unificado pois possue cadastro no débito automático.", vbCritical, "Atenção"
                        AdicionaImovel = False
                       .Close
                       Exit Function
                    End If
                End With
            
                ReDim Preserve aImovel(UBound(aImovel) + 1)
                aImovel(UBound(aImovel)).nCodReduz = !CODREDUZIDO
                aImovel(UBound(aImovel)).nDistrito = !Distrito
                aImovel(UBound(aImovel)).nSetor = !Setor
                aImovel(UBound(aImovel)).nQuadra = !Quadra
                aImovel(UBound(aImovel)).nNumero = !Li_Num
                aImovel(UBound(aImovel)).nCodBairro = !Li_CodBairro
                aImovel(UBound(aImovel)).sDescBairro = !DescBairro
                aImovel(UBound(aImovel)).nCodUsoTerreno = !Dt_CodUsoTerreno
                aImovel(UBound(aImovel)).sDescUsoTerreno = !DescUsoTerreno
                aImovel(UBound(aImovel)).nCodBenfeitoria = !Dt_CodBenf
                aImovel(UBound(aImovel)).sDescBenfeitoria = !DescBenfeitoria
                aImovel(UBound(aImovel)).nCodTopografia = !Dt_CodTopog
                aImovel(UBound(aImovel)).sDescTopografia = !DescTopografia
                aImovel(UBound(aImovel)).nCodCatProp = !Dt_CodCategProp
                aImovel(UBound(aImovel)).sDescCatProp = !DescCategProp
                aImovel(UBound(aImovel)).nCodSituacao = !Dt_CodSituacao
                aImovel(UBound(aImovel)).sDescSituacao = !DescSituacao
                aImovel(UBound(aImovel)).nCodPedologia = !Dt_CodPedol
                aImovel(UBound(aImovel)).sDescPedologia = !DescPedologia
                aImovel(UBound(aImovel)).nAreaTerreno = !Dt_AreaTerreno
                aImovel(UBound(aImovel)).nCodProprietario = !CodCidadao
                aImovel(UBound(aImovel)).sNomeProprietario = !NomeCidadao
                CarregaCampos
                AdicionaImovel = True
            End If
        End If
    End If
   .Close
End With
End Function

Private Sub CarregaCampos()
Dim x As Integer, Y As Integer, bAchou As Boolean, nAreaTerreno As Double, itmX As ListItem
LimpaCampos
If UBound(aImovel) = 0 Then Exit Sub

lblDist.Caption = Format(aImovel(1).nDistrito, "00")
lblSetor.Caption = Format(aImovel(1).nSetor, "00")

For x = 1 To UBound(aImovel)
    bAchou = False
    For Y = 0 To cmbQuadra.ListCount - 1
        If aImovel(x).nQuadra = Val(cmbQuadra.List(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbQuadra.AddItem Format(aImovel(x).nQuadra, "0000")
    End If
    bAchou = False
    For Y = 0 To cmbNumero.ListCount - 1
        If aImovel(x).nNumero = Val(cmbNumero.List(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbNumero.AddItem aImovel(x).nNumero
    End If
    bAchou = False
    For Y = 0 To cmbBairro.ListCount - 1
        If aImovel(x).nCodBairro = Val(cmbBairro.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbBairro.AddItem aImovel(x).sDescBairro
        cmbBairro.ItemData(cmbBairro.NewIndex) = aImovel(x).nCodBairro
    End If
   'USO TERRENO
    bAchou = False
    For Y = 0 To cmbUso.ListCount - 1
        If aImovel(x).nCodUsoTerreno = Val(cmbUso.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbUso.AddItem aImovel(x).sDescUsoTerreno
        cmbUso.ItemData(cmbUso.NewIndex) = aImovel(x).nCodUsoTerreno
    End If
   'BENFEITORIA
    bAchou = False
    For Y = 0 To cmbBen.ListCount - 1
        If aImovel(x).nCodBenfeitoria = Val(cmbBen.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbBen.AddItem aImovel(x).sDescBenfeitoria
        cmbBen.ItemData(cmbBen.NewIndex) = aImovel(x).nCodBenfeitoria
    End If
   'TOPOGRAFIA
    bAchou = False
    For Y = 0 To cmbTop.ListCount - 1
        If aImovel(x).nCodTopografia = Val(cmbTop.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbTop.AddItem aImovel(x).sDescTopografia
        cmbTop.ItemData(cmbTop.NewIndex) = aImovel(x).nCodTopografia
    End If
   'CATEG.PROP.
    bAchou = False
    For Y = 0 To cmbCat.ListCount - 1
        If aImovel(x).nCodCatProp = Val(cmbCat.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbCat.AddItem aImovel(x).sDescCatProp
        cmbCat.ItemData(cmbCat.NewIndex) = aImovel(x).nCodCatProp
    End If
   'SITUACAO
    bAchou = False
    For Y = 0 To cmbSit.ListCount - 1
        If aImovel(x).nCodSituacao = Val(cmbSit.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbSit.AddItem aImovel(x).sDescSituacao
        cmbSit.ItemData(cmbSit.NewIndex) = aImovel(x).nCodSituacao
    End If
   'PEDOLOGIA
    bAchou = False
    For Y = 0 To cmbPed.ListCount - 1
        If aImovel(x).nCodPedologia = Val(cmbPed.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbPed.AddItem aImovel(x).sDescPedologia
        cmbPed.ItemData(cmbPed.NewIndex) = aImovel(x).nCodPedologia
    End If
   'PROPRIETARIO
    bAchou = False
    For Y = 0 To cmbProp.ListCount - 1
        If aImovel(x).nCodProprietario = Val(cmbProp.ItemData(Y)) Then
            bAchou = True
            Exit For
        End If
    Next
    If Not bAchou Then
        cmbProp.AddItem aImovel(x).sNomeProprietario
        cmbProp.ItemData(cmbProp.NewIndex) = aImovel(x).nCodProprietario
    End If
    'TESTADAS
    Sql = "SELECT codreduzido, numface, areatestada From Testada Where CODREDUZIDO =" & aImovel(x).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            bAchou = False
            For Y = 0 To grdTestada.Rows - 1
                If Val(grdTestada.TextMatrix(Y, 0)) = !NUMFACE Then
                    bAchou = True
                    Exit For
                End If
            Next
            If bAchou Then
               grdTestada.TextMatrix(Y, 1) = FormatNumber(CDbl(grdTestada.TextMatrix(Y, 1)) + !AREATESTADA, 2)
            Else
               grdTestada.AddItem Format(!NUMFACE, "00") & Chr(9) & FormatNumber(!AREATESTADA, 2)
            End If
           .MoveNext
        Loop
       .Close
    End With
    'AREAS
    Sql = "SELECT AREAS.SEQAREA,AREAS.TIPOAREA,AREAS.DATAAPROVA,AREAS.AREACONSTR,"
    Sql = Sql & "AREAS.USOCONSTR,USOCONSTR.DESCUSOCONSTR,AREAS.TIPOCONSTR,TIPOCONSTR.DESCTIPOCONSTR,"
    Sql = Sql & "AREAS.CATCONSTR,CATEGCONSTR.DESCCATEGCONSTR FROM AREAS INNER JOIN USOCONSTR ON "
    Sql = Sql & "AREAS.USOCONSTR = USOCONSTR.CODUSOCONSTR INNER JOIN TIPOCONSTR ON "
    Sql = Sql & "AREAS.TIPOCONSTR = TIPOCONSTR.CODTIPOCONSTR INNER JOIN CATEGCONSTR ON "
    Sql = Sql & "AREAS.CATCONSTR = CATEGCONSTR.CODCATEGCONSTR "
    Sql = Sql & "WHERE CODREDUZIDO=" & aImovel(x).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
          Do Until .EOF
               Set itmX = lvArea.ListItems.Add(, , Format(aImovel(x).nCodReduz, "000000"))
                itmX.SubItems(1) = Format(!SEQAREA, "00")
                itmX.SubItems(2) = !TIPOAREA
                itmX.SubItems(3) = FormatNumber(!AREACONSTR, 2) & " m²"
                itmX.SubItems(4) = Format(!DATAAPROVA, "dd/mm/yyyy")
                itmX.SubItems(5) = Format(!USOCONSTR, "00") & " - " & !descusoconstr
                itmX.SubItems(6) = Format(!TIPOCONSTR, "00") & " - " & !DESCTIPOCONSTR
                itmX.SubItems(7) = Format(!CATCONSTR, "00") & " - " & !desccategconstr
              .MoveNext
          Loop
         .Close
    End With
    
Next

If cmbQuadra.ListCount = 1 Then cmbQuadra.ListIndex = 0
If cmbNumero.ListCount = 1 Then cmbNumero.ListIndex = 0
If cmbBairro.ListCount = 1 Then cmbBairro.ListIndex = 0
If cmbUso.ListCount = 1 Then cmbUso.ListIndex = 0
If cmbBen.ListCount = 1 Then cmbBen.ListIndex = 0
If cmbTop.ListCount = 1 Then cmbTop.ListIndex = 0
If cmbCat.ListCount = 1 Then cmbCat.ListIndex = 0
If cmbSit.ListCount = 1 Then cmbSit.ListIndex = 0
If cmbPed.ListCount = 1 Then cmbPed.ListIndex = 0
If cmbProp.ListCount = 1 Then cmbProp.ListIndex = 0

'AREA TERRENO
nAreaTerreno = 0
For Y = 1 To UBound(aImovel)
    nAreaTerreno = nAreaTerreno + aImovel(Y).nAreaTerreno
Next
lblAreaTerreno.Caption = FormatNumber(nAreaTerreno, 2)

End Sub

Private Sub LimpaCampos()
lblDist.Caption = "00"
lblSetor.Caption = "00"
lblLote.Caption = ""
cmbQuadra.Clear
cmbNumero.Clear
cmbBairro.Clear
cmbUso.Clear
cmbBen.Clear
cmbTop.Clear
cmbPed.Clear
cmbCat.Clear
cmbSit.Clear
cmbProp.Clear
cmbFace.Clear
lblEnd.Caption = ""
lblCep.Caption = ""
txtFracao.Text = "0,00"
lblAreaTerreno.Caption = "0,00"
grdTestada.Rows = 1
Inicio:
For Y = 1 To lvArea.ListItems.Count
    lvArea.ListItems.Remove (Y)
    GoTo Inicio
Next

End Sub

Private Sub cmdDelImovel_Click()
Dim x As Integer
If lstImovel.ListIndex = -1 Then
    MsgBox "Selecione o imóvel a ser removido da lista.", vbCritical, "Atenção"
Else
    lstImovel.RemoveItem (lstImovel.ListIndex)
    Inicio
    For x = 0 To lstImovel.ListCount - 1
        AdicionaImovel lstImovel.List(x)
    Next
    CarregaCampos
End If
End Sub

Private Sub cmdExec_Click()
Dim Achou As Boolean

If lstImovel.ListCount < 2 Then
    MsgBox "Selecione ao menos 2 imóveis.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbQuadra.ListIndex = -1 Or cmbFace.ListIndex = -1 Then
    MsgBox "Selecione quadra e face.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbBairro.ListIndex = -1 Or cmbNumero.ListIndex = -1 Then
    MsgBox "Selecione bairro e número.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbProp.ListIndex = -1 Then
    MsgBox "Selecione o proprietário.", vbCritical, "Atenção"
    Exit Sub
End If

If cmbUso.ListIndex = -1 Or cmbBen.ListIndex = -1 Or cmbTop.ListIndex = -1 Or cmbPed.ListIndex = -1 Or cmbCat.ListIndex = -1 Or cmbSit.ListIndex = -1 Then
    MsgBox "Selecione todos os dados do terreno.", vbCritical, "Atenção"
    Exit Sub
End If

'Area
If lvArea.ListItems.Count > 0 Then
    Achou = False
    For x = 1 To lvArea.ListItems.Count
          If lvArea.ListItems(x).Checked = True Then
               Achou = True
          End If
    Next
    
    If Not Achou Then
       MsgBox "Selecione a Área Principal.", vbCritical, "Erro de Validação."
       Exit Sub
    End If
End If

'testada
Achou = False
For x = 1 To grdTestada.Rows - 1
    If Val(grdTestada.TextMatrix(x, 0)) = Val(cmbFace.Text) Then
        Achou = True
        Exit For
    End If
Next
If Not Achou Then
    MsgBox "A face selecionada não possue nenhuma testada.", vbCritical, "Atenção"
    Exit Sub
End If

If MsgBox("Deseja executar a unificação destes imóveis ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    cmdAddImovel.Enabled = False
    cmdDelImovel.Enabled = False
    cmdExec.Enabled = False
    cmdSair.Enabled = False
    Pnl1.Enabled = False
    Pnl2.Enabled = False
    pnlWait.Visible = True
    Ocupado
    If cGetInputState() <> 0 Then DoEvents
    Grava
    For x = 1 To UBound(aImovel)
        TransferenciaDivida aImovel(x).nCodReduz, nCodReduz
        'STATUS TRANSFERIDO
        Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & aImovel(x).nCodReduz & " AND STATUSLANC=3"
        cn.Execute Sql, rdExecDirect
    Next
    Liberado
    If cGetInputState() <> 0 Then DoEvents
    pnlWait.Visible = False
    MsgBox "Unificação executada com sucesso." & vbCrLf & "Criado o imóvel: " & nCodReduz, vbInformation, "INFORMAÇÃO"
    Unload Me
Else
    MsgBox "Gravação cancelada.", vbCritical, "Cancelado"
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Inicio
End Sub

Private Sub Inicio()
ReDim aImovel(0)
LimpaCampos
End Sub

Private Sub lvArea_ItemCheck(ByVal Item As MSComctlLib.ListItem)
Dim x As Integer

If Item.ListSubItems(2).Text = "C" Then
     Item.Checked = False
     MsgBox "Apenas áreas principais podem ser selecionadas.", vbExclamation, "Atenção"
     Exit Sub
End If

With lvArea
    For x = 1 To .ListItems.Count
          If .ListItems(x).Index <> Item.Index Then
               .ListItems(x).Checked = False
          End If
    Next
End With

End Sub

Private Sub Grava()
Dim nLote As Integer, sData As String, x As Integer, nSeq As Integer, sHist As String, Y As Integer

'Verificamos novamente o ultimo lote e codreduzido para evitar duplicação
Sql = "SELECT MAX(LOTE) AS ULTIMOLOTE FROM CADIMOB WHERE "
Sql = Sql & "DISTRITO=" & Val(lblDist.Caption) & " AND SETOR=" & Val(lblSetor.Caption) & " AND QUADRA=" & Val(cmbQuadra.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
nLote = RdoAux!ULTIMOLOTE + 1
RdoAux.Close

Sql = "SELECT MAX(CODREDUZIDO) AS ULTIMOCODREDUZ FROM CADIMOB WHERE CODREDUZIDO<40000"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
nCodReduz = RdoAux!ULTIMOCODREDUZ + 1
RdoAux.Close

'*******GRAVA IMOVEL**********************************************

Sql = "INSERT CADIMOB(CODREDUZIDO,DV,CODCONDOMINIO,DISTRITO,SETOR,QUADRA,LOTE,SEQ,UNIDADE,SUBUNIDADE,LI_NUM,LI_COMPL,"
Sql = Sql & "LI_UF,LI_CODCIDADE,LI_CODBAIRRO,LI_QUADRAS,LI_LOTES,DT_AREATERRENO,DT_CODUSOTERRENO,DT_CODBENF,DT_CODTOPOG,"
Sql = Sql & "DT_CODCATEGPROP,DT_CODSITUACAO,DT_CODPEDOL,DT_NUMAGUA,DT_FRACAOIDEAL,DC_QTDEEDIF,DC_QTDEPAV,EE_TIPOEND,INATIVO,RESIDEIMOVEL) values("
Sql = Sql & nCodReduz & "," & RetornaDVCodReduzido(nCodReduz) & "," & 999 & ","
Sql = Sql & Val(lblDist.Caption) & "," & Val(lblSetor.Caption) & "," & Val(cmbQuadra.Text) & "," & nLote & ","
Sql = Sql & Val(cmbFace.Text) & "," & 0 & "," & 0 & "," & Val(cmbNumero.Text) & ",'"
Sql = Sql & "" & "','" & "SP" & "'," & 413 & "," & cmbBairro.ItemData(cmbBairro.ListIndex) & ",'"
Sql = Sql & "" & "','" & "" & "'," & Virg2Ponto(RemovePonto(lblAreaTerreno.Caption)) & "," & cmbUso.ItemData(cmbUso.ListIndex) & ","
Sql = Sql & cmbBen.ItemData(cmbBen.ListIndex) & "," & cmbTop.ItemData(cmbTop.ListIndex) & ","
Sql = Sql & cmbCat.ItemData(cmbCat.ListIndex) & "," & cmbSit.ItemData(cmbSit.ListIndex) & ","
Sql = Sql & cmbPed.ItemData(cmbPed.ListIndex) & "," & "Null" & "," & Virg2Ponto(txtFracao.Text) & "," & 0 & ",0,0,0,1)"
cn.Execute Sql, rdExecDirect

'*******GRAVA PROPRIETARIO *******************************************
Sql = "INSERT PROPRIETARIO (CODREDUZIDO,CODCIDADAO,TIPOPROP,PRINCIPAL) VALUES("
Sql = Sql & nCodReduz & "," & cmbProp.ItemData(cmbProp.ListIndex) & ",'"
Sql = Sql & "P" & "'," & 1 & ")"
cn.Execute Sql, rdExecDirect
AtualizaPropDuplicado nCodReduz, cmbProp.ItemData(cmbProp.ListIndex)

'*******GRAVA TESTADA *******************************************
GravaTestada:
For x = 1 To grdTestada.Rows - 1
    Sql = "INSERT TESTADA(CODREDUZIDO,NUMFACE,AREATESTADA) VALUES("
    Sql = Sql & nCodReduz & "," & Val(grdTestada.TextMatrix(x, 0)) & "," & Virg2Ponto(RemovePonto(grdTestada.TextMatrix(x, 1))) & ")"
    cn.Execute Sql, rdExecDirect
Next

'*******GRAVA AREA *******************************************
For x = 1 To lvArea.ListItems.Count
    sData = lvArea.ListItems(x).ListSubItems(4).Text
    Sql = "INSERT AREAS (CODREDUZIDO,SEQAREA,TIPOAREA,DATAAPROVA,AREACONSTR,USOCONSTR,TIPOCONSTR,CATCONSTR) VALUES(" & nCodReduz & "," & x & ",'" & IIf(lvArea.ListItems(x).Checked, "P", "C") & "',"
    Sql = Sql & IIf(IsDate(sData), "'" & Format(sData, "mm/dd/yyyy") & "'", "Null") & "," & Virg2Ponto(RemovePonto(Left$(lvArea.ListItems(x).ListSubItems(3).Text, Len(lvArea.ListItems(x).ListSubItems(3).Text) - 3))) & ","
    Sql = Sql & Val(Left$(lvArea.ListItems(x).ListSubItems(5).Text, 2)) & "," & Val(Left$(lvArea.ListItems(x).ListSubItems(6).Text, 2)) & "," & Val(Left$(lvArea.ListItems(x).ListSubItems(7).Text, 2)) & ")"
    cn.Execute Sql, rdExecDirect
Next

'*******INATIVA OS IMÓVEIS ANTIGOS*******
For x = 1 To UBound(aImovel)
    Sql = "UPDATE CADIMOB SET INATIVO=1 WHERE CODREDUZIDO=" & aImovel(x).nCodReduz
    cn.Execute Sql, rdExecDirect
Next

'*******GRAVA HISTÓRICO DOS LOTES********
'NOVO LOTE
Sql = "SELECT CODREDUZIDO,SEQ FROM HISTORICO WHERE CODREDUZIDO=" & nCodReduz
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
If RdoAux.RowCount > 0 Then
    nSeq = RdoAux!Seq + 1
Else
   nSeq = 1
End If
sHist = "O imóvel foi criado a partir da Unificação dos Imóveis: "
For x = 1 To UBound(aImovel)
    sHist = sHist & Format(aImovel(x).nCodReduz, "000000") & ", "
Next
sHist = Left$(sHist, Len(sHist) - 2)
Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST,DATAHIST2) VALUES("
Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & sHist & "','" & Format(Now, "mm/dd/yyyy") & "')"
cn.Execute Sql, rdExecDirect

'LOTES ANTIGOS

For x = 1 To UBound(aImovel)
    Sql = "SELECT max(SEQ) as MAXIMO FROM HISTORICO WHERE CODREDUZIDO=" & aImovel(x).nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset)
    If Not IsNull(RdoAux!maximo) Then
        nSeq = RdoAux!maximo + 1
    Else
       nSeq = 1
    End If
    sHist = "O imóvel foi unificado com o(s) imóvel(is): "
    For Y = 0 To lstImovel.ListCount - 1
        If lstImovel.List(Y) <> aImovel(x).nCodReduz Then
            sHist = sHist & Format(lstImovel.List(Y), "000000") & ", "
        End If
    Next
    sHist = Left$(sHist, Len(sHist) - 2)
    sHist = sHist & " e criou o imóvel " & Format(nCodReduz, "000000")
    
    Sql = "INSERT HISTORICO (CODREDUZIDO,SEQ,DATAHIST,DESCHIST) VALUES("
    Sql = Sql & aImovel(x).nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sHist) & "')"
    cn.Execute Sql, rdExecDirect
Next

End Sub

Public Sub TransferenciaDivida(nCodAntigo As Long, nCodNovo As Long)
Dim nAno As Integer, nLanc As Integer, nSeq As Integer, nParc As Integer, nCompl As Integer
Dim nSeq2 As Integer, RdoAux2 As rdoResultset, RdoS As rdoResultset

'CARREGA DEBITOPARCELA ANTIGO MENOS OS PARCELAMENTOS
Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodAntigo & " AND CODLANCAMENTO<>20"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        nAno = !AnoExercicio
        nLanc = !CodLancamento
        nSeq = !SeqLancamento
        nParc = !NumParcela
        nCompl = !CODCOMPLEMENTO
        'BUSCA NOVA SEQUENCIA
        Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodNovo & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If IsNull(!maximo) Then
                nSeq2 = 0
            Else
                nSeq2 = !maximo + 1
            End If
           .Close
        End With
        'GRAVA DEBITOPARCELA NOVO
'        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,NUMPROCESSO,USUARIO) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !statuslanc & ",'"
'        Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',1," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & ","
'        Sql = Sql & Val(SubNull(!numcertidao)) & "," & IIf(IsDate(!datainscricao), "'" & Format(!datainscricao, "mm/dd/yyyy") & "'", "Null") & "," & IIf(IsDate(!dataajuiza), "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'", "Null") & ",'" & !numprocesso & "','" & Left$(NomeDeLogin, 25) & "')"
        Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
        Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,NUMPROCESSO,USERID) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !statuslanc & ",'"
        Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "',1," & Val(SubNull(!numerolivro)) & "," & Val(SubNull(!paginalivro)) & ","
        Sql = Sql & Val(SubNull(!numcertidao)) & "," & IIf(IsDate(!datainscricao), "'" & Format(!datainscricao, "mm/dd/yyyy") & "'", "Null") & "," & IIf(IsDate(!dataajuiza), "'" & Format(!dataajuiza, "mm/dd/yyyy") & "'", "Null") & ",'" & !numprocesso & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
        cn.Execute Sql, rdExecDirect
        'BUSCA DEBITOTRIBUTO ANTIGO
        Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & nCodAntigo & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            Do Until .EOF
                'GRAVA DEBITOTRIBUTO NOVO
                Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !CodTributo & "," & Virg2Ponto(CStr(!ValorTributo)) & ")"
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With
        'ATUALIZA PARCELADOCUMENTO
        Sql = "UPDATE PARCELADOCUMENTO SET CODREDUZIDO=" & nCodNovo & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodAntigo & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        'ATUALIZA DEBITOPAGO
        Sql = "UPDATE DEBITOPAGO SET CODREDUZIDO=" & nCodNovo & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodAntigo & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
        'ATUALIZA OBS
        Sql = "UPDATE obsparcela SET CODREDUZIDO=" & nCodNovo & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodAntigo & " AND "
        Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
        Sql = Sql & "CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

'CARREGA SEQ DE PARCELAMENTOS DO DEBITOPARCELA ANTIGO (SMAR)
Sql = "SELECT DISTINCT SEQLANCAMENTO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodAntigo & " AND CODLANCAMENTO=20 ORDER BY SEQLANCAMENTO"
Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoS
    Do Until .EOF
        nSeq = !SeqLancamento
        'BUSCA NOVA SEQUENCIA PARA PARCELAMENTO
        Sql = "SELECT MAX(SEQLANCAMENTO) AS MAXIMO FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodNovo & " AND CODLANCAMENTO=20"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            If IsNull(!maximo) Then
                nSeq2 = 0
            Else
                nSeq2 = !maximo + 1
            End If
           .Close
        End With
        'CARREGA PARCELAMENTO DEBITOPARCELA ANTIGO
        Sql = "SELECT * FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodAntigo & " AND CODLANCAMENTO=20 AND SEQLANCAMENTO=" & nSeq
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            Do Until .EOF
                nAno = !AnoExercicio
                nLanc = !CodLancamento
                nParc = !NumParcela
                nCompl = !CODCOMPLEMENTO
                'GRAVA DEBITOPARCELA NOVO
'                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
'                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,VALORJUROS,NUMPROCESSO,USUARIO) VALUES("
'                Sql = Sql & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !statuslanc & ",'"
'                Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "'," & !CODMOEDA & "," & IIf(IsNull(!numerolivro), 0, !numerolivro) & ","
'                Sql = Sql & IIf(IsNull(!paginalivro), 0, !paginalivro) & "," & IIf(IsNull(!numcertidao), 0, !numcertidao) & ",'" & Format(!datainscricao, "mm/dd/yyyy") & "','" & Format(!dataajuiza, "mm/dd/yyyy") & "',"
'                Sql = Sql & IIf(IsNull(Virg2Ponto(CStr(!ValorJuros))), 0, Virg2Ponto(CStr(!ValorJuros))) & ",'" & SubNull(!numprocesso) & "','" & Left$(NomeDeLogin, 25) & "')"
                Sql = "INSERT DEBITOPARCELA (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,STATUSLANC,"
                Sql = Sql & "DATAVENCIMENTO,DATADEBASE,CODMOEDA,NUMEROLIVRO,PAGINALIVRO,NUMCERTIDAO,DATAINSCRICAO,DATAAJUIZA,VALORJUROS,NUMPROCESSO,USERID) VALUES("
                Sql = Sql & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !statuslanc & ",'"
                Sql = Sql & Format(!DataVencimento, "mm/dd/yyyy") & "','" & Format(!DATADEBASE, "mm/dd/yyyy") & "'," & !CODMOEDA & "," & IIf(IsNull(!numerolivro), 0, !numerolivro) & ","
                Sql = Sql & IIf(IsNull(!paginalivro), 0, !paginalivro) & "," & IIf(IsNull(!numcertidao), 0, !numcertidao) & ",'" & Format(!datainscricao, "mm/dd/yyyy") & "','" & Format(!dataajuiza, "mm/dd/yyyy") & "',"
                Sql = Sql & IIf(IsNull(!ValorJuros), 0, Virg2Ponto(CStr(SubNull(!ValorJuros)))) & ",'" & SubNull(!numprocesso) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
                cn.Execute Sql, rdExecDirect
                'BUSCA DEBITOTRIBUTO ANTIGO
                Sql = "SELECT * FROM DEBITOTRIBUTO WHERE CODREDUZIDO=" & nCodAntigo & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux2
                    Do Until .EOF
                        'GRAVA DEBITOTRIBUTO NOVO
                        Sql = "INSERT DEBITOTRIBUTO (CODREDUZIDO,ANOEXERCICIO,CODLANCAMENTO,SEQLANCAMENTO,NUMPARCELA,CODCOMPLEMENTO,CODTRIBUTO,"
                        Sql = Sql & "VALORTRIBUTO) VALUES(" & nCodNovo & "," & nAno & "," & nLanc & "," & nSeq2 & "," & nParc & "," & nCompl & "," & !CodTributo & "," & Virg2Ponto(CStr(!ValorTributo)) & ")"
                        cn.Execute Sql, rdExecDirect
                       .MoveNext
                    Loop
                   .Close
                End With
                'ATUALIZA PARCELADOCUMENTO
                Sql = "UPDATE PARCELADOCUMENTO SET CODREDUZIDO=" & nCodNovo & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodAntigo & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
                'ATUALIZA DEBITOPAGO
                Sql = "UPDATE DEBITOPAGO SET CODREDUZIDO=" & nCodNovo & ",SEQLANCAMENTO=" & nSeq2 & " WHERE CODREDUZIDO=" & nCodAntigo & " AND "
                Sql = Sql & "ANOEXERCICIO=" & nAno & " AND CODLANCAMENTO=" & nLanc & " AND SEQLANCAMENTO=" & nSeq & " AND NUMPARCELA=" & nParc & " AND "
                Sql = Sql & "CODCOMPLEMENTO=" & nCompl
                cn.Execute Sql, rdExecDirect
               .MoveNext
            Loop
           .Close
        End With
        'ATUALIZA REPARCTMP
        Sql = "UPDATE REPARCTMP SET CODREDUZD=" & nCodNovo & " ,CODSEQD=" & nSeq2 & ", CODREDUZO=" & nCodNovo & " WHERE CODREDUZD=" & nCodAntigo & " AND "
        Sql = Sql & "CODLANCD=" & nLanc & " AND CODSEQD=" & nSeq
        cn.Execute Sql, rdExecDirect
        'ATUALIZA REPARC2TMP
        Sql = "UPDATE REPARC2TMP SET CODREDUZ=" & nCodNovo & " ,CODREDUZ2=" & nCodNovo & " ,CODSEQ=" & nSeq2 & " WHERE CODREDUZ=" & nCodAntigo & " AND CODSEQ=" & nSeq
        cn.Execute Sql, rdExecDirect
        'ATUALIZA DESTINOREPARC
        Sql = "UPDATE DESTINOREPARC SET CODREDUZIDO=" & nCodNovo & ",NUMSEQUENCIA=" & nSeq2 & " WHERE "
        Sql = Sql & "CODREDUZIDO=" & nCodAntigo & " AND NUMSEQUENCIA=" & nSeq
        Sql = Sql & " AND CODCOMPLEMENTO=" & nCompl
        cn.Execute Sql, rdExecDirect
       .MoveNext
    Loop
   .Close
End With

'ATUALIZA ORIGEMREPARC
Sql = "UPDATE ORIGEMREPARC SET CODREDUZIDO=" & nCodNovo & " WHERE CODREDUZIDO=" & nCodAntigo
cn.Execute Sql, rdExecDirect
        
'ATUALIZA PROCESSOREPARC
Sql = "UPDATE PROCESSOREPARC SET CODIGORESP=" & nCodNovo & " WHERE CODIGORESP=" & nCodAntigo
cn.Execute Sql, rdExecDirect

Sql = "UPDATE DEBITOPARCELA SET STATUSLANC=13 WHERE CODREDUZIDO=" & nCodAntigo & " AND STATUSLANC=3"
cn.Execute Sql, rdExecDirect

End Sub

