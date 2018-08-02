VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmRequerIPTU 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Requerimento para isenção de IPTU"
   ClientHeight    =   4740
   ClientLeft      =   4125
   ClientTop       =   3045
   ClientWidth     =   7710
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4740
   ScaleWidth      =   7710
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   315
      Left            =   6570
      TabIndex        =   14
      ToolTipText     =   "Imprimir Requerimento"
      Top             =   4365
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmRequerIPTU.frx":0000
      PICN            =   "frmRequerIPTU.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.CheckBox chkObs 
      Caption         =   "Incluir observação sobre impossibilidade física de deslocamento"
      Height          =   240
      Left            =   90
      TabIndex        =   13
      Top             =   4410
      Width           =   4965
   End
   Begin VB.Frame Frame4 
      Caption         =   "Observação"
      Height          =   1230
      Left            =   90
      TabIndex        =   43
      Top             =   3060
      Width           =   7530
      Begin VB.TextBox txtObs 
         Appearance      =   0  'Flat
         Height          =   915
         Left            =   90
         MaxLength       =   5000
         MultiLine       =   -1  'True
         TabIndex        =   12
         Top             =   225
         Width           =   7350
      End
   End
   Begin VB.Frame Frame3 
      Height          =   465
      Left            =   3555
      TabIndex        =   42
      Top             =   0
      Width           =   4065
      Begin VB.OptionButton OptTipo 
         Caption         =   "Renovação da Isenção"
         Height          =   195
         Index           =   1
         Left            =   1890
         TabIndex        =   3
         Top             =   180
         Width           =   2085
      End
      Begin VB.OptionButton OptTipo 
         Caption         =   "Isenção do IPTU"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   2
         Top             =   180
         Value           =   -1  'True
         Width           =   1635
      End
   End
   Begin VB.Frame Frame2 
      Height          =   465
      Left            =   90
      TabIndex        =   41
      Top             =   0
      Width           =   3435
      Begin VB.OptionButton Opt 
         Caption         =   "Pessoa Física"
         Height          =   195
         Index           =   0
         Left            =   90
         TabIndex        =   0
         Top             =   180
         Value           =   -1  'True
         Width           =   1410
      End
      Begin VB.OptionButton Opt 
         Caption         =   "Pessoa Jurídica"
         Height          =   195
         Index           =   1
         Left            =   1620
         TabIndex        =   1
         Top             =   180
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1140
      Left            =   90
      TabIndex        =   32
      Top             =   1935
      Width           =   7530
      Begin prjChameleon.chameleonButton cmdCnsImovel2 
         Height          =   270
         Left            =   1080
         TabIndex        =   11
         ToolTipText     =   "Consulta Cidadão"
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   476
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRequerIPTU.frx":0176
         PICN            =   "frmRequerIPTU.frx":0192
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
         Caption         =   "N°..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   14
         Left            =   6210
         TabIndex        =   40
         Top             =   495
         Width           =   345
      End
      Begin VB.Label Label1 
         Caption         =   "Bairro..........:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   13
         Left            =   90
         TabIndex        =   39
         Top             =   765
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço....:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   12
         Left            =   90
         TabIndex        =   38
         Top             =   495
         Width           =   975
      End
      Begin VB.Label lblBairroImovel 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1080
         TabIndex        =   37
         Top             =   765
         Width           =   4980
      End
      Begin VB.Label lblNumImovel 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   6615
         TabIndex        =   36
         Top             =   495
         Width           =   750
      End
      Begin VB.Label lblEndImovel 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1080
         TabIndex        =   35
         Top             =   495
         Width           =   4980
      End
      Begin VB.Label lblCodImovel 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1665
         TabIndex        =   34
         Top             =   225
         Width           =   750
      End
      Begin VB.Label Label1 
         Caption         =   "Cód. Imóvel.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   11
         Left            =   90
         TabIndex        =   33
         Top             =   225
         Width           =   1020
      End
   End
   Begin VB.Frame frF 
      Height          =   1545
      Left            =   90
      TabIndex        =   15
      Top             =   405
      Width           =   7530
      Begin VB.TextBox txtNumProc1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2655
         TabIndex        =   5
         Top             =   1125
         Width           =   1230
      End
      Begin prjChameleon.chameleonButton cmdCnsImovel 
         Height          =   270
         Left            =   1080
         TabIndex        =   4
         ToolTipText     =   "Consulta Cidadão"
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   476
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRequerIPTU.frx":02EC
         PICN            =   "frmRequerIPTU.frx":0308
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
         Caption         =   "Processo de Avaliação Social Nº..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   90
         TabIndex        =   30
         Top             =   1170
         Width           =   2505
      End
      Begin VB.Label lblRequerente 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1620
         TabIndex        =   23
         Top             =   225
         Width           =   5835
      End
      Begin VB.Label Label1 
         Caption         =   "Requerente.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   90
         TabIndex        =   22
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Nº de RG..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   90
         TabIndex        =   21
         Top             =   540
         Width           =   885
      End
      Begin VB.Label lblRG 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1035
         TabIndex        =   20
         Top             =   540
         Width           =   1830
      End
      Begin VB.Label Label1 
         Caption         =   "Nº do CPF..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   3015
         TabIndex        =   19
         Top             =   540
         Width           =   975
      End
      Begin VB.Label lblCPF 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4005
         TabIndex        =   18
         Top             =   540
         Width           =   3450
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   90
         TabIndex        =   17
         Top             =   855
         Width           =   885
      End
      Begin VB.Label lblEndereco 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   270
         Left            =   1035
         TabIndex        =   16
         Top             =   855
         Width           =   6420
      End
   End
   Begin VB.Frame frJ 
      Height          =   1545
      Left            =   90
      TabIndex        =   24
      Top             =   405
      Visible         =   0   'False
      Width           =   7530
      Begin VB.TextBox txtRepresentante 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         TabIndex        =   6
         Top             =   495
         Width           =   3660
      End
      Begin VB.TextBox txtNumProc2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1755
         TabIndex        =   10
         Top             =   1125
         Width           =   1230
      End
      Begin VB.TextBox txtEnd 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   285
         Left            =   3240
         TabIndex        =   9
         Top             =   810
         Width           =   4110
      End
      Begin VB.TextBox txtCPF 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   540
         TabIndex        =   8
         Top             =   810
         Width           =   1545
      End
      Begin VB.TextBox txtRG 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5535
         TabIndex        =   7
         Top             =   495
         Width           =   1815
      End
      Begin prjChameleon.chameleonButton cmdCnsEmpresa 
         Height          =   270
         Left            =   1080
         TabIndex        =   45
         ToolTipText     =   "Consulta Cidadão"
         Top             =   180
         Width           =   465
         _ExtentX        =   820
         _ExtentY        =   476
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
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmRequerIPTU.frx":0462
         PICN            =   "frmRequerIPTU.frx":047E
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label lblCNPJ 
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   5580
         TabIndex        =   47
         Top             =   225
         Width           =   1815
      End
      Begin VB.Label lblRazao 
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   1620
         TabIndex        =   46
         Top             =   225
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Representan.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   15
         Left            =   90
         TabIndex        =   44
         Top             =   540
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Processo Anterior Nº..:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   90
         TabIndex        =   31
         Top             =   1170
         Width           =   1695
      End
      Begin VB.Label Label1 
         Caption         =   "CNPJ.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   4950
         TabIndex        =   29
         Top             =   225
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Requerente.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   7
         Left            =   90
         TabIndex        =   28
         Top             =   225
         Width           =   1020
      End
      Begin VB.Label Label1 
         Caption         =   "Nº RG:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   4950
         TabIndex        =   27
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "CPF.:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   90
         TabIndex        =   26
         Top             =   855
         Width           =   480
      End
      Begin VB.Label Label1 
         Caption         =   "Endereço....:"
         BeginProperty Font 
            Name            =   "Microsoft Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   2205
         TabIndex        =   25
         Top             =   855
         Width           =   1020
      End
   End
End
Attribute VB_Name = "frmRequerIPTU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCnsEmpresa_Click()
Set frm = frmCidadao
frm.sForm = Me.Name & "2"
frm.show
frm.ZOrder 0
End Sub

Private Sub cmdCnsImovel_Click()
Set frm = frmCidadao
frm.sForm = Me.Name
frm.show
frm.ZOrder 0
End Sub

Private Sub cmdCnsImovel2_Click()
sForm = Me.Name
frmCnsImovel.show
frmCnsImovel.ZOrder 0
End Sub

Private Sub cmdPrint_Click()
Dim Sql As String

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

If txtObs.Text <> "" Then
    Sql = "INSERT REPORTTMP(USUARIO,MEMO1) VALUES('" & NomeDeLogin & "','Obs: " & Mask(txtObs.Text) & "')"
    cn.Execute Sql, rdExecDirect
End If

frmReport.ShowReport2 "REQUERIPTU", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM REPORTTMP WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
    lblCodImovel.Caption = Format(Val(Left$(CodImovel, 7)), "000000")
    CodImovel = 0
    LeImovel
End If

If Opt(1).Value = True Then Exit Sub
If Val(lblRequerente.Tag) = 0 Then Exit Sub
If Val(lblRequerente.Tag) > 500000 Then
    Le
Else
    MsgBox "Código de cidadão inválido.", vbExclamation, "Atenção"
    Limpa
End If

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub Opt_Click(Index As Integer)
If Index = 0 Then
    frF.Visible = True
    frJ.Visible = False
Else
    frJ.Visible = True
    frF.Visible = False
End If
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset
Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & Val(lblRequerente.Tag)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        lblRG.Caption = SubNull(!rg) & " " & SubNull(!ORGAO)
        lblCPF.Caption = SubNull(!CPF)
        If lblCPF.Caption <> "" Then
            lblCPF.Caption = Format(RdoAux!CPF, "00#\.###\.###-##")
        End If
        If lblCPF.Caption = "" Then
            If Not IsNull(!Cnpj) Then
                lblCPF.Caption = Format(!Cnpj, "0#\.###\.###/####-##")
            End If
        End If
        lblEndereco.Caption = SubNull(!Endereco) & " ," & SubNull(!NUMIMOVEL)
        If Not IsNull(!DescBairro) Then
            lblEndereco.Caption = lblEndereco.Caption & " - " & !DescBairro
'        Else
'            If Not IsNull(!NOMEBairro) Then
'                lblEndereco.Caption = lblEndereco.Caption & " - " & !NOMEBairro
'            End If
        End If
        If Not IsNull(!Complemento) Or !Complemento <> "" Then
            lblEndereco.Caption = lblEndereco.Caption & " " & !Complemento
        End If
        
'        If Not IsNull(!NomeCidade) Then
'            lblEndereco.Caption = lblEndereco.Caption & " - na cidade de " & !NomeCidade & "/" & SubNull(!NOMEUF)
'        Else
            If Not IsNull(!descCidade) Then
                lblEndereco.Caption = lblEndereco.Caption & " - na cidade de " & !descCidade & "/" & SubNull(!siglaUF)
            End If
 '       End If
       
       .MoveNext
    Loop
   .Close
End With


End Sub

Private Sub Limpa()
lblRequerente.Caption = ""
lblRG.Caption = ""
lblCPF.Caption = ""
lblEndereco.Caption = ""
End Sub

Private Sub LeImovel()
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT * FROM vwFULLIMOVEL2 WHERE CODREDUZIDO=" & Val(lblCodImovel.Caption)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    lblEndImovel.Caption = SubNull(!Logradouro)
    lblNumImovel.Caption = SubNull(!Li_Num)
    lblBairroImovel.Caption = SubNull(!DescBairro)
    .Close
End With

End Sub
