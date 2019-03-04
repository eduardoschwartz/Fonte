VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAlvara 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alvará de Funcionamento"
   ClientHeight    =   5685
   ClientLeft      =   6150
   ClientTop       =   3225
   ClientWidth     =   6990
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5685
   ScaleWidth      =   6990
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   585
      MaxLength       =   200
      TabIndex        =   54
      Top             =   1890
      Width           =   6270
   End
   Begin VB.TextBox txtObs 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   630
      MaxLength       =   400
      TabIndex        =   53
      Top             =   4860
      Width           =   6270
   End
   Begin VB.CheckBox chkPrefeito 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Ocultar assinatura do Prefeito"
      Height          =   195
      Left            =   135
      TabIndex        =   37
      Top             =   5310
      Width           =   2625
   End
   Begin VB.CheckBox chkTipoAlvara 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Alvará provisório"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   1
      Top             =   1620
      Width           =   1815
   End
   Begin VB.ComboBox cmbAss 
      Height          =   315
      ItemData        =   "frmAlvara.frx":0000
      Left            =   1260
      List            =   "frmAlvara.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   35
      Top             =   4455
      Width           =   5730
   End
   Begin VB.TextBox txtCodigo 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1110
      MaxLength       =   6
      TabIndex        =   0
      Top             =   120
      Width           =   1155
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5850
      TabIndex        =   41
      ToolTipText     =   "Sair da Tela"
      Top             =   5295
      Width           =   1125
      _ExtentX        =   1984
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
      MICON           =   "frmAlvara.frx":0004
      PICN            =   "frmAlvara.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4635
      TabIndex        =   39
      ToolTipText     =   "Cancelar Edição"
      Top             =   5295
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmAlvara.frx":008E
      PICN            =   "frmAlvara.frx":00AA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskCNPJ 
      Height          =   240
      Left            =   1110
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   450
      Width           =   2145
      _ExtentX        =   3784
      _ExtentY        =   423
      BackColor       =   15658734
      ForeColor       =   8388608
      MouseIcon       =   "frmAlvara.frx":0204
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
      BackStyle       =   0
      BorderStyle     =   0
      MaxLength       =   18
      Mask            =   "99.999.999/9999-99"
      SelText         =   ""
      Text            =   "__.___.___/____-__"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskCPF 
      Height          =   240
      Left            =   4140
      TabIndex        =   44
      Top             =   450
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   423
      BackColor       =   15658734
      MouseIcon       =   "frmAlvara.frx":0220
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
      BorderStyle     =   0
      MaxLength       =   14
      Mask            =   "999.999.999-99"
      SelText         =   ""
      Text            =   "___.___.___-__"
      HideSelection   =   -1  'True
   End
   Begin VB.Frame frP 
      BackColor       =   &H00EEEEEE&
      Height          =   2175
      Left            =   45
      TabIndex        =   17
      Top             =   2250
      Width           =   6765
      Begin VB.TextBox txtAlvara 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1125
         MaxLength       =   4
         TabIndex        =   18
         Top             =   225
         Width           =   705
      End
      Begin VB.TextBox txtProcesso 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   5430
         TabIndex        =   20
         Top             =   255
         Width           =   1155
      End
      Begin VB.ListBox lstDoc 
         Appearance      =   0  'Flat
         Height          =   1155
         ItemData        =   "frmAlvara.frx":023C
         Left            =   135
         List            =   "frmAlvara.frx":024F
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   855
         Width           =   6465
      End
      Begin VB.Label lblAnoAlvara 
         BackStyle       =   0  'Transparent
         Caption         =   "/2000"
         Height          =   225
         Left            =   1890
         TabIndex        =   51
         Top             =   270
         Width           =   1005
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Alvará nº.....:"
         Height          =   225
         Left            =   135
         TabIndex        =   23
         Top             =   285
         Width           =   1005
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Processo nº.....:"
         Height          =   225
         Index           =   0
         Left            =   4230
         TabIndex        =   21
         Top             =   315
         Width           =   1155
      End
      Begin VB.Label Label14 
         BackStyle       =   0  'Transparent
         Caption         =   "Lista de documentos:"
         Height          =   225
         Left            =   135
         TabIndex        =   19
         Top             =   615
         Width           =   1605
      End
   End
   Begin VB.Frame frN 
      BackColor       =   &H00EEEEEE&
      Height          =   2175
      Left            =   90
      TabIndex        =   24
      Top             =   2250
      Width           =   6765
      Begin VB.CheckBox chkBombon 
         Caption         =   "Bombonieri"
         Height          =   240
         Left            =   3420
         TabIndex        =   27
         Top             =   270
         Width           =   1275
      End
      Begin VB.CheckBox chk24Hrs 
         Caption         =   "24 Hrs."
         Height          =   240
         Left            =   2520
         TabIndex        =   26
         Top             =   270
         Width           =   825
      End
      Begin VB.ComboBox cmbDataAlvara 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAlvara.frx":0312
         Left            =   1170
         List            =   "frmAlvara.frx":0314
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1665
         Width           =   5505
      End
      Begin VB.ComboBox cmbTipo 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "frmAlvara.frx":0316
         Left            =   1170
         List            =   "frmAlvara.frx":0329
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   585
         Width           =   5505
      End
      Begin VB.TextBox txtProcesso2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1200
         TabIndex        =   25
         Top             =   225
         Width           =   1155
      End
      Begin esMaskEdit.esMaskedEdit mskDataBomb 
         Height          =   285
         Left            =   2385
         TabIndex        =   29
         Top             =   945
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         MouseIcon       =   "frmAlvara.frx":03C1
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
      Begin esMaskEdit.esMaskedEdit mskDataVS 
         Height          =   285
         Left            =   2385
         TabIndex        =   30
         Top             =   1305
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         MouseIcon       =   "frmAlvara.frx":03DD
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
      Begin esMaskEdit.esMaskedEdit mskDataSaaej 
         Height          =   285
         Left            =   5670
         TabIndex        =   31
         Top             =   945
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         MouseIcon       =   "frmAlvara.frx":03F9
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
      Begin esMaskEdit.esMaskedEdit mskDataCETESB 
         Height          =   285
         Left            =   5670
         TabIndex        =   32
         Top             =   1305
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   503
         MouseIcon       =   "frmAlvara.frx":0415
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
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Data alvará com SAAEJ.....:"
         Height          =   225
         Index           =   5
         Left            =   3555
         TabIndex        =   48
         Top             =   990
         Width           =   2100
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Data alvará com CETESB..:"
         Height          =   225
         Index           =   4
         Left            =   3555
         TabIndex        =   47
         Top             =   1350
         Width           =   2100
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Data alvará...:"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   42
         Top             =   1710
         Width           =   1005
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Data alvará com vig.sanitária..:"
         Height          =   225
         Index           =   3
         Left            =   135
         TabIndex        =   40
         Top             =   1350
         Width           =   2280
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Data alvará com bombeiro......:"
         Height          =   225
         Index           =   2
         Left            =   135
         TabIndex        =   38
         Top             =   990
         Width           =   2280
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Tipo alvará...:"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   36
         Top             =   630
         Width           =   1005
      End
      Begin VB.Label Label13 
         BackStyle       =   0  'Transparent
         Caption         =   "Processo nº..:"
         Height          =   225
         Index           =   1
         Left            =   135
         TabIndex        =   34
         Top             =   285
         Width           =   1155
      End
   End
   Begin VB.Label lblIE 
      BackStyle       =   0  'Transparent
      Caption         =   "IEst....:"
      Height          =   225
      Left            =   5280
      TabIndex        =   59
      Top             =   1590
      Width           =   1665
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "IEst....:"
      Height          =   225
      Left            =   4710
      TabIndex        =   58
      Top             =   1590
      Width           =   825
   End
   Begin VB.Label Ponto 
      BackStyle       =   0  'Transparent
      Caption         =   "Ponto/Agen...:"
      Height          =   225
      Index           =   1
      Left            =   120
      TabIndex        =   57
      Top             =   5880
      Width           =   1005
   End
   Begin VB.Label lblPontoAgencia 
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1110
      TabIndex        =   56
      Top             =   5880
      Width           =   2565
   End
   Begin VB.Label Label16 
      Caption         =   "Data:"
      Height          =   195
      Left            =   90
      TabIndex        =   55
      Top             =   1935
      Width           =   420
   End
   Begin VB.Label Label15 
      Caption         =   "Obs:"
      Height          =   195
      Left            =   135
      TabIndex        =   52
      Top             =   4905
      Width           =   420
   End
   Begin VB.Label lblCompl 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4500
      TabIndex        =   50
      Top             =   810
      Width           =   765
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Compl.:"
      Height          =   225
      Left            =   3915
      TabIndex        =   49
      Top             =   810
      Width           =   600
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Cidade......:"
      Height          =   225
      Left            =   2010
      TabIndex        =   46
      Top             =   1590
      Width           =   1005
   End
   Begin VB.Label lblCidade 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2850
      TabIndex        =   45
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "CPF.......:"
      Height          =   225
      Left            =   3420
      TabIndex        =   43
      Top             =   495
      Width           =   735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Assinatura....:"
      Height          =   225
      Index           =   0
      Left            =   225
      TabIndex        =   16
      Top             =   4500
      Width           =   1005
   End
   Begin VB.Label lblAtividade 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1110
      TabIndex        =   14
      Top             =   1380
      Width           =   5715
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Atividade....:"
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   1380
      Width           =   1005
   End
   Begin VB.Label lblCEP 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   4500
      TabIndex        =   12
      Top             =   1110
      Width           =   2325
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "CEP..........:"
      Height          =   225
      Left            =   3570
      TabIndex        =   11
      Top             =   1110
      Width           =   915
   End
   Begin VB.Label lblBairro 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1110
      TabIndex        =   10
      Top             =   1080
      Width           =   2325
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Bairro.........:"
      Height          =   225
      Left            =   120
      TabIndex        =   9
      Top             =   1080
      Width           =   1005
   End
   Begin VB.Label lblNum 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   6060
      TabIndex        =   8
      Top             =   780
      Width           =   765
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Nº...:"
      Height          =   225
      Left            =   5580
      TabIndex        =   7
      Top             =   780
      Width           =   495
   End
   Begin VB.Label lblEndereco 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   1110
      TabIndex        =   6
      Top             =   780
      Width           =   2565
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Endereço...:"
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   5
      Top             =   780
      Width           =   1005
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "CNPJ.........:"
      Height          =   225
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   1005
   End
   Begin VB.Label lblNome 
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2370
      TabIndex        =   3
      Top             =   180
      Width           =   4635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Contribuinte:"
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   180
      Width           =   1005
   End
End
Attribute VB_Name = "frmAlvara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim nOldAlvara As Long, sHorario As String
Public sControle As String

Private Sub chkTipoAlvara_Click()
Dim Sql As String, RdoAux As rdoResultset

If chkTipoAlvara.value = 0 Then
    frN.Visible = True
    frP.Visible = False
Else
    Sql = "select * from parametros where nomeparam='SEQALVARA'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        txtAlvara.Text = !valparam
        nOldAlvara = !valparam
        .Close
    End With
    frN.Visible = False
    frP.Visible = True
    cmbTipo.ListIndex = -1
End If
End Sub

Private Sub cmdPrint_Click()
Dim z As Integer, bAchou As Boolean, bAchou2 As Boolean, Sql As String, sTexto1 As String, nCodReduz As Long, nSeq As Integer
Dim qd As New rdoQuery, RdoAux As rdoResultset, sValidade As String, sDoc As String

If chkTipoAlvara.value = vbChecked Then
    If NomeDeLogin <> "RITA" And NomeDeLogin <> "DANIELAR" And NomeDeLogin <> "SCHWARTZ" Then
        MsgBox "Você não tem permissão para emitir alvará provisório.", vbCritical, "ERRO"
        Exit Sub
    End If
End If

'If cmbDataAlvara.ListIndex = 0 Then
'    MsgBox "Data do Alvará inválida,", vbCritical, "ERRO"
'    Exit Sub
'End If


If lblNome.Caption = "" Then
    MsgBox "Selecione o contribuinte", vbCritical, "ERRO"
    Exit Sub
End If

If chkTipoAlvara.value = 1 Then
    If txtAlvara.Text = "" Then
        MsgBox "Digite o número do alvará", vbCritical, "ERRO"
        Exit Sub
    End If
    If txtProcesso.Text = "" Then
        MsgBox "Digite o número do processo", vbCritical, "ERRO"
        Exit Sub
    End If
    
    bAchou = False
    For z = 0 To lstDoc.ListCount - 1
        If lstDoc.Selected(z) = True Then
            bAchou = True
        End If
    Next
    
    bAchou2 = False
    If Not bAchou Then
        MsgBox "Selecione ao menos um documento.", vbCritical, "Atenção"
    Else
        nOldAlvara = nOldAlvara + 1
'        Sql = "update parametros set valparam='" & CStr(nOldAlvara) & "' where nomeparam='SEQALVARA'"
'        cn.Execute Sql, rdExecDirect
        frmReport.ShowReport "ALVARAPROVISORIO", frmMdi.HWND, Me.HWND
       ' frmReport.ShowReport "ALVARAPROVISORIOVICE", frmMdi.HWND, Me.HWND
        txtAlvara.Text = nOldAlvara
        Limpa
        txtCodigo.Text = ""
    End If
Else
    If cmbAss.ListIndex = -1 Then
        MsgBox "Selecione uma assinatura.", vbExclamation, "Atenção"
        Exit Sub
    End If
    If txtProcesso2.Text = "" Then
        MsgBox "Digite o número do processo", vbCritical, "ERRO"
        Exit Sub
    End If
    If cmbDataAlvara.ListIndex = -1 Then
        MsgBox "Selecione a data do alvará.", vbExclamation, "Atenção"
        Exit Sub
    End If
    
    nCodReduz = Val(txtCodigo.Text)
    sTexto1 = "Emisão de Alvará de Funcionamento."
    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 0
    Else
        nSeq = RdoAux!maximo + 1
    End If
                
    Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
    Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sTexto1) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
    cn.Execute Sql, rdExecDirect

    
    
    
    
    
    Set qd.ActiveConnection = cn
    On Error Resume Next
    RdoAux.Close
    On Error GoTo 0
    qd.Sql = "{ Call spALVARA2(?) }"
    qd(0) = Val(txtCodigo.Text)
    Set RdoAux = qd.OpenResultset(rdOpenKeyset)
    If RdoAux!Tipo <> 1 And RdoAux!Tipo <> 3 Then
        MsgBox "Solicitação inválida", vbCritical, "ERRO"
        Exit Sub
    Else
        sHorario = SubNull(RdoAux!Horario)
        sDoc = SubNull(RdoAux!Cnpj)
        If sDoc = "" Then
            sDoc = SubNull(RdoAux!CPF)
            If sDoc <> "" Then
                sDoc = Format(RdoAux!CPF, "000\.000\.000-00")
            End If
        Else
            sDoc = Format(RdoAux!Cnpj, "00\.000\.000/0000-00")
        End If
    End If
    RdoAux.Close
    sValidade = "30/06/2019"
    
    If cmbDataAlvara.ListIndex = 1 Then
        Sql = "select * from debitoparcela where codreduzido=" & Val(txtCodigo.Text) & " and anoexercicio=" & Year(Now) & " and "
        Sql = Sql & "codlancamento=6 and numparcela>0 and statuslanc=3"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            MsgBox "Existem parcelas não pagas de taxa de licença no exercício atual.", vbCritical, "Atenção"
            Exit Sub
        End If
    End If
    
    'Sql = "select * from parametros where nomeparam='SEQALVARA'"
    Sql = "select max(numero) as maximo from alvara_funcionamento where ano=" & Year(Now)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If RdoAux!maximo = Null Then
            nodalvara = 1
        Else
            nOldAlvara = !maximo + 1
        End If
        .Close
    End With
    
    sControle = Format(nOldAlvara, "00000") & Format(Year(Now), "0000") & "/" & Format(Val(txtCodigo.Text), "000000") & "-AF"
    Sql = " insert alvara_funcionamento(ano,numero,controle,codigo,razao_social,documento,endereco,bairro,atividade,horario,validade,data_gravada) values("
    Sql = Sql & Year(Now) & "," & nOldAlvara & ",'" & sControle & "'," & Val(txtCodigo.Text) & ",'" & Mask(lblNome.Caption) & "','" & sDoc & "','" & Mask(lblEndereco.Caption) & "','" & Mask(lblBairro.Caption) & "','" & lblAtividade.Caption & "','" & sHorario & "','" & Format(sValidade, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
    cn.Execute Sql, rdExecDirect
    
'    Sql = "update parametros set valparam='" & nOldAlvara + 1 & "' where nomeparam='SEQALVARA'"
'    cn.Execute Sql, rdExecDirect
    
    If txtData.Text = "" Then
        frmReport.ShowReport2 "ALVARA", frmMdi.HWND, Me.HWND
        'frmReport.ShowReport2 "ALVARAVICE", frmMdi.HWND, Me.HWND
    Else
        frmReport.ShowReport2 "ALVARASEMDATA", frmMdi.HWND, Me.HWND
        'frmReport.ShowReport2 "ALVARASEMDATAVICE", frmMdi.HWND, Me.HWND
    End If
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset

Centraliza Me
frN.Visible = True
frP.Visible = False
cmbAss.AddItem "(Sem Assinatura)"
Sql = "select * from assinatura WHERE usuario<>'NOBODY'  order by nome"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        If UCase(Left(!Nome, 5)) <> "RITA " And UCase(Left(!Nome, 5)) <> "DANIE" Then
        Else
        cmbAss.AddItem !Nome
        End If
       .MoveNext
    Loop
   .Close
End With

nOldAlvara = 0
lblAnoAlvara.Caption = "/" & CStr(Year(Now))
'cmbDataAlvara.AddItem "*** DESATIVADO **** Alvará até 30 de junho de " & Year(Now)
cmbDataAlvara.AddItem "Alvará até 30 de junho de " & Year(Now)
cmbDataAlvara.AddItem "Alvará até 30 de junho de " & Year(Now) + 1
cmbDataAlvara.ListIndex = 1
End Sub

Private Sub txtAlvara_KeyPress(KeyAscii As Integer)
Tweak txtAlvara, KeyAscii, IntegerPositive
End Sub

Private Sub txtAlvara_LostFocus()
Dim Sql As String

If Val(txtAlvara.Text) = 0 Then
    txtAlvara.Text = nOldAlvara
    Exit Sub
End If

If Val(txtAlvara.Text) <> nOldAlvara Then
    nOldAlvara = Val(txtAlvara.Text)
'    Sql = "update parametros set valparam='" & CStr(nOldAlvara) & "' where nomeparam='SEQALVARA'"
'    cn.Execute Sql, rdExecDirect
End If

End Sub

Private Sub txtCodigo_Change()
If lblNome.Caption <> "" Then Limpa
End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    Carrega
Else
    Tweak txtCodigo, KeyAscii, IntegerPositive
End If
End Sub

Public Sub Carrega()
Dim Sql As String, RdoAux As rdoResultset
Limpa
If Val(txtCodigo.Text) = 0 Then Exit Sub

Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & Val(txtCodigo.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then
        MsgBox "Código não cadastrado", vbExclamation, "Atenção"
       .Close
        Exit Sub
    Else
        lblNome.Caption = !razaosocial
        mskCNPJ.Text = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
        mskCPF.Text = Format(Trim(!CPF), "000\.000\.000-00")
        If Not IsNull(!NomeLogradouro) Then
            lblEndereco.Caption = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro)
        Else
            lblEndereco.Caption = SubNull(!NomeLogr)
        End If
        lblCompl.Caption = !Complemento
        lblNum.Caption = !Numero
        lblBairro.Caption = !DescBairro
        lblCidade.Caption = SubNull(!descCidade) & " - " & SubNull(!SiglaUF)
        lblCEP.Caption = Format(!Cep, "00000-000")
        lblAtividade.Caption = !ativextenso
        lblPontoAgencia.Caption = SubNull(!ponto_agencia)
        lblIE.Caption = IIf(IsNull(!INSCESTADUAL), "ISENTO", !INSCESTADUAL)
    End If
   .Close
End With


End Sub

Private Sub Limpa()
lblNome.Caption = ""
LimpaMascara mskCNPJ
lblEndereco.Caption = ""
lblPontoAgencia.Caption = ""
lblNum.Caption = ""
lblBairro.Caption = ""
lblAtividade.Caption = ""
txtProcesso.Text = ""
txtProcesso2.Text = ""
End Sub

Private Sub txtCodigo_LostFocus()
Carrega
End Sub

