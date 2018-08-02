VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{9DC93C3A-4153-440A-88A7-A10AEDA3BAAA}#3.5#0"; "vbalDTab6.ocx"
Begin VB.Form frmSecurity 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Atribuição de Níveis de Segurança do Sistema"
   ClientHeight    =   6450
   ClientLeft      =   6705
   ClientTop       =   1080
   ClientWidth     =   9150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6450
   ScaleWidth      =   9150
   ShowInTaskbar   =   0   'False
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   1380
      Left            =   30
      TabIndex        =   6
      Top             =   7500
      Width           =   5310
      _ExtentX        =   9366
      _ExtentY        =   2434
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      ForeColorFixed  =   0
      BackColorSel    =   16777215
      ForeColorSel    =   -2147483642
      FocusRect       =   0
      FormatString    =   "CodTela |Nome Tela                  |CodEvento|Evento            |X    "
   End
   Begin MSComctlLib.ImageList ilsIcons 
      Left            =   6930
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":0000
            Key             =   "TTF"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":015C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":02BC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imlSecurity 
      Left            =   360
      Top             =   3030
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":0418
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmSecurity.frx":0734
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   675
      Left            =   0
      TabIndex        =   1
      Top             =   5760
      Width           =   3045
      Begin prjChameleon.chameleonButton cmdAlterar 
         Height          =   345
         Left            =   2160
         TabIndex        =   11
         ToolTipText     =   "Editar Registro"
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
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
         MICON           =   "frmSecurity.frx":0A50
         PICN            =   "frmSecurity.frx":0A6C
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
         Height          =   345
         Left            =   2580
         TabIndex        =   7
         ToolTipText     =   "Sair da Tela"
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
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
         MICON           =   "frmSecurity.frx":0BC6
         PICN            =   "frmSecurity.frx":0BE2
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
         Height          =   345
         Left            =   2580
         TabIndex        =   9
         ToolTipText     =   "Gravar os Dados"
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSecurity.frx":0C50
         PICN            =   "frmSecurity.frx":0C6C
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
         Height          =   345
         Left            =   2160
         TabIndex        =   10
         ToolTipText     =   "Cancelar Edição"
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BTYPE           =   14
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
         COLTYPE         =   1
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmSecurity.frx":1011
         PICN            =   "frmSecurity.frx":102D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdAddAll 
         Height          =   345
         Left            =   570
         TabIndex        =   12
         ToolTipText     =   "Adicionar Acesso Total"
         Top             =   210
         Visible         =   0   'False
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MICON           =   "frmSecurity.frx":1187
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdRemoveAll 
         Height          =   345
         Left            =   1005
         TabIndex        =   13
         ToolTipText     =   "Remover todos os Acessos"
         Top             =   210
         Width           =   375
         _ExtentX        =   661
         _ExtentY        =   609
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
         MICON           =   "frmSecurity.frx":11A3
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdHerdaGrupo 
         Height          =   345
         Left            =   570
         TabIndex        =   8
         ToolTipText     =   "Herdar Atributos do Grupo"
         Top             =   210
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   609
         BTYPE           =   3
         TX              =   ""
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
         MICON           =   "frmSecurity.frx":11BF
         PICN            =   "frmSecurity.frx":11DB
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
   End
   Begin MSComctlLib.TreeView tvUser 
      Height          =   5745
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   3045
      _ExtentX        =   5371
      _ExtentY        =   10134
      _Version        =   393217
      Style           =   7
      FullRowSelect   =   -1  'True
      Appearance      =   1
   End
   Begin MSFlexGridLib.MSFlexGrid grdObj 
      Height          =   4905
      Left            =   -390
      TabIndex        =   15
      Top             =   1050
      Width           =   6015
      _ExtentX        =   10610
      _ExtentY        =   8652
      _Version        =   393216
      Rows            =   1
      Cols            =   7
      FormatString    =   "Objeto                                         |^Select |^Insert |^Update |^Delete |^Exec |^Cg "
   End
   Begin MSComctlLib.TreeView tvOpt 
      Height          =   4725
      Left            =   1290
      TabIndex        =   16
      Top             =   360
      Width           =   5985
      _ExtentX        =   10557
      _ExtentY        =   8334
      _Version        =   393217
      Indentation     =   647
      Style           =   7
      Appearance      =   1
   End
   Begin vbalDTab6.vbalDTabControl TabSeg 
      Height          =   5955
      Left            =   3030
      TabIndex        =   14
      Top             =   480
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   10504
      TabAlign        =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty SelectedFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   15658734
   End
   Begin VB.Label lblGrupo 
      Height          =   225
      Left            =   8040
      TabIndex        =   5
      Top             =   90
      Visible         =   0   'False
      Width           =   945
   End
   Begin VB.Label lblTag 
      Height          =   255
      Left            =   3990
      TabIndex        =   4
      Top             =   60
      Visible         =   0   'False
      Width           =   405
   End
   Begin VB.Label lblUser 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   285
      Left            =   4620
      TabIndex        =   3
      Top             =   60
      Width           =   3615
   End
   Begin VB.Label lblTipo 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   285
      Left            =   3120
      TabIndex        =   2
      Top             =   60
      Width           =   1455
   End
End
Attribute VB_Name = "frmSecurity"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim RdoS As rdoResultset
Dim RdoAux2 As rdoResultset
Dim Sql As String
Dim Evento As String

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
    cmdCancel.Visible = False
    cmdGravar.Visible = False
    cmdAlterar.Visible = True
    cmdHerdaGrupo.Visible = True
    cmdSair.Visible = True
    tvUser.Enabled = True
ElseIf Tipo = "INCLUIR" Then
    cmdCancel.Visible = True
    cmdGravar.Visible = True
    cmdAlterar.Visible = False
    cmdHerdaGrupo.Visible = False
    cmdSair.Visible = False
    tvUser.Enabled = False
End If

End Sub

Private Sub cmdAddAll_Click()

If tvUser.SelectedItem.Children > 0 Then
     If MsgBox("Atribuir acesso total ao Grupo " & lblUser.Caption & "?" & vbCrLf & vbCrLf & "Atenção!!! Esta operação garantirá o acesso completo e incondicional  ao Sistema de Tributação.", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
          grdTemp.Rows = 1
          Evento = "AddAll"
          Limpa
          Ocupado
          Screen.MousePointer = vbHourglass
          For x = 1 To tvOpt.Nodes.Count
              If tvOpt.Nodes(x).Children = 0 Then
                 tvOpt_NodeClick tvOpt.Nodes(x)
              End If
          Next
          Grava
          Liberado
          Screen.MousePointer = vbDefault
     End If
Else
     If MsgBox("Atribuir acesso total ao Usuário " & lblUser.Caption & "?" & vbCrLf & vbCrLf & "Atenção!!! Esta operação garantirá o acesso completo e incondicional ao Sistema de Tributação.", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
          grdTemp.Rows = 1
          Evento = "AddAll"
          Limpa
          Ocupado
          Screen.MousePointer = vbHourglass
          For x = 1 To tvOpt.Nodes.Count
              If tvOpt.Nodes(x).Children = 0 Then
                 tvOpt_NodeClick tvOpt.Nodes(x)
              End If
          Next
          Grava
          Liberado
          Screen.MousePointer = vbDefault
     End If
End If

End Sub

Private Sub cmdAlterar_Click()
Evento = "INCLUIR"
Eventos "INCLUIR"
End Sub

Private Sub cmdCancel_Click()
Evento = ""
Eventos "INICIAR"
tvUser_NodeClick tvUser.SelectedItem
End Sub

Private Sub cmdGravar_Click()
If MsgBox("Deseja gravar as alterações ?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
   Ocupado
   Screen.MousePointer = vbHourglass
   Grava
   Screen.MousePointer = vbDefault
   Liberado
   MsgBox "Autorização gravada com sucesso." & vbCrLf & vbCrLf & "AS ALTERAÇÕES SÓ TERÃO EFEITO NO PRÓXIMO LOGIN DO SISTEMA DE TRIBUTAÇÃO", vbInformation, "Informação"
End If
Evento = ""
Eventos "INICIAR"
'tvUser_NodeClick tvUser.SelectedItem

End Sub


Private Sub cmdHerdaGrupo_Click()

Dim sNomeGrupo As String
Dim sCodUser As String
Ocupado
If tvUser.SelectedItem.Children = 0 Then
    If MsgBox("TODOS os atributos de segurança concedidos a este usuário serão apagados e substituidos pelos atributos de segurança do Grupo." & vbCrLf & vbCrLf & "Voce deseja Continuar ???", vbQuestion + vbYesNo, "Leia com Atenção !!!") = vbYes Then
        sCodUser = lblUser.Caption
        sNomeGrupo = tvUser.SelectedItem.Parent.Text
        Sql = "DELETE FROM SEG_USERACESS WHERE NOMEUSUARIO='" & sCodUser & "'"
        cn.Execute Sql, rdExecDirect
        Sql = "INSERT SEG_USERACESS SELECT '" & sCodUser & "',CODTELA,CODEVENTO "
        Sql = Sql & "FROM SEG_GRUPOACESSO WHERE NOMEGRUPO='" & sNomeGrupo & "'"
        cn.Execute Sql, rdExecDirect
        CarregaGrid "U"
        Limpa
        Le
        'GRAVA PERMISSAO NO SQL SERVER
        With grdObj
             For x = 1 To .Rows - 1
                If cGetInputState() <> 0 Then DoEvents
                GravaAtribSqlServer lblUser.Caption, .TextMatrix(x, 0), IIf(Val(.TextMatrix(x, 1)) > 0, True, False), IIf(Val(.TextMatrix(x, 3)) > 0, True, False), IIf(Val(.TextMatrix(x, 2)) > 0, True, False), IIf(Val(.TextMatrix(x, 4)) > 0, True, False), IIf(Val(.TextMatrix(x, 5)) > 0, True, False)
             Next
        End With
        Conecta NomeDeLogin, sWd, "Tributacao"
        DefaultAccess sCodUser
    End If
    MsgBox "Permissões herdadas com sucesso.", vbInformation, "Atenção"
Else
   MsgBox "Um Grupo não pode herdar direitos dele mesmo.", vbExclamation, "Atenção"
End If
Liberado
End Sub

Private Sub cmdRemoveAll_Click()

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
Ocupado
Evento = "RemoveAll"

For x = 1 To tvOpt.Nodes.Count
    DoEvents
    If tvOpt.Nodes(x).Children = 0 Then
       tvOpt_NodeClick tvOpt.Nodes(x)
       If tvOpt.Nodes(x).Image = 1 Then
          tvOpt_NodeClick tvOpt.Nodes(x)
       End If
    End If
Next
Liberado
cmdGravar_Click
End Sub

Private Sub cmdSair_Click()
Unload Me

End Sub

Private Sub Command1_Click()



'grant select,update,insert,delete on analise2 to gtisys


End Sub

Private Sub Form_Load()
Ocupado

Dim c As cTab
With TabSeg
    .ShowCloseButton = False
    .AllowScroll = False
    Set c = .Tabs.Add("Tab1", , "Acesso a Objeto")
    c.Panel = tvOpt
    Set c = .Tabs.Add("Tab2", , "Base de Dados")
    c.Panel = grdObj
End With

Screen.MousePointer = vbHourglass
Centraliza Me
On Error GoTo Erro
'Set oSQLServer = New SQLDMO.SQLServer
'oSQLServer.LoginTimeout = 20

'With oSQLServer
'    .Disconnect
'    .LoginTimeout = 20
'    .Connect "192.168.15.160", LastUser, UserPwd
'End With

BuildTvUser
Evento = ""
MontaTvOpt
MontaBD
Eventos "INICIAR"

tvUser.Nodes(1).Selected = True
tvUser_NodeClick tvUser.Nodes(1)
Screen.MousePointer = vbDefault

Exit Sub
Erro:

MsgBox Err.Description
Screen.MousePointer = vbDefault
Resume Next

Liberado

End Sub

Private Sub MontaBD()
On Error Resume Next
For x = 1 To cn.rdoTables.Count
    If cn.rdoTables(x).Type = "TABLE" Then
        grdObj.AddItem cn.rdoTables(x).Name
    ElseIf cn.rdoTables(x).Type = "VIEW" Then
        If UCase(Left(cn.rdoTables(x).Name, 2)) = "VW" Then
            grdObj.AddItem cn.rdoTables(x).Name
        End If
    End If
Next

grdObj.AddItem "spDADOSDEUMIMOVEL"
grdObj.AddItem "spEXTRATO"
grdObj.AddItem "spEXTRATONEW"
grdObj.AddItem "spGRAVABAIXATMP"
grdObj.AddItem "spGRAVAMOBILIARIO"
grdObj.AddItem "spGRAVAPARAMPARCELA"
grdObj.AddItem "spGRAVAPROCESSO"
grdObj.AddItem "spRELDEVEDOR"
grdObj.AddItem "spRELDEVEDORREPARCELAMENTO"

End Sub

Private Sub MontaTvOpt()
On Error GoTo Erro
Dim x As Integer
Dim NodX As Object

With tvOpt
    .ImageList = ilsIcons
    Sql = "SELECT CODTELA,NOMEFORM,NOMETELA  FROM SEG_TELASISTEMA "
    Sql = Sql & "ORDER BY NOMETELA"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    Do Until RdoAux.EOF
          Set NodX = .Nodes.Add(, , RdoAux!NOMETELA & Format(RdoAux!CODTELA, "000"), RdoAux!NOMETELA, 2)
          .Nodes(RdoAux!NOMETELA & Format(RdoAux!CODTELA, "000")).Bold = True
          Sql = "SELECT DISTINCT  SEG_EVENTOACESSO.CODEVENTO AS CODIGOEVENTO,"
          Sql = Sql & "SEG_EVENTO.DESCEVENTO,SEG_EVENTOACESSO.CODTELA "
          Sql = Sql & "FROM SEG_EVENTOACESSO INNER JOIN  SEG_EVENTO ON "
          Sql = Sql & "SEG_EVENTOACESSO.CODEVENTO = SEG_EVENTO.CODEVENTO "
          Sql = Sql & "Where CODTELA =" & RdoAux!CODTELA
          Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
          If RdoS.RowCount > 0 Then
               Do Until RdoS.EOF
                  Set NodX = .Nodes.Add(RdoAux!NOMETELA & Format(RdoAux!CODTELA, "000"), tvwChild, RdoAux!NOMETELA & RdoS!DESCEVENTO, RdoS!DESCEVENTO, 3)
                 .Nodes(RdoAux!NOMETELA & RdoS!DESCEVENTO).Tag = RdoS!CODIGOEVENTO
                  RdoS.MoveNext
               Loop
          Else
                Set NodX = .Nodes.Add(RdoAux!NOMETELA & Format(RdoAux!CODTELA, "000"), tvwChild, RdoAux!NOMETELA & "Abrir Tela", "Abrir Tela", 3)
                .Nodes(RdoAux!NOMETELA & "Abrir Tela").Tag = 1
          End If
          RdoS.Close
          RdoAux.MoveNext
    Loop
    RdoAux.Close
End With

For x = 1 To tvOpt.Nodes.Count
   tvOpt.Nodes(x).EnsureVisible
Next
tvOpt.Nodes(1).Selected = True

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub BuildTvUser()
On Error GoTo Erro

Dim NodX As Node
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset
Dim sGrupo As String, sUsuario As String

tvUser.ImageList = imlSecurity

Sql = "SELECT DISTINCT GRUPO FROM USUARIO WHERE GRUPO IS NOT NULL AND GRUPO<>''"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sGrupo = !grupo
        Set NodX = tvUser.Nodes.Add(, , sGrupo, sGrupo, 1)
        tvUser.Nodes(sGrupo).Bold = True
        Sql = "SELECT NOMELOGIN FROM USUARIO where GRUPO='" & sGrupo & "' and ativo=1 order by NOMELOGIN"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
             Do Until .EOF
                Set NodX = tvUser.Nodes.Add(sGrupo, tvwChild, sGrupo & !NomeLogin, !NomeLogin, 2)
                .MoveNext
             Loop
            .Close
        End With
       .MoveNext
    Loop
   .Close
End With
  
For x = 1 To tvUser.Nodes.Count
   tvUser.Nodes(x).EnsureVisible
Next

Exit Sub
Erro:
MsgBox Err.Description
Resume Next

End Sub

Private Sub Form_Unload(Cancel As Integer)
Set oSQLServer = Nothing

End Sub

Private Sub tvOpt_NodeClick(ByVal Node As MSComctlLib.Node)

Dim nCodTela As Integer
Dim nCodEvento As Integer
Dim sAtrib As String
Dim x As Integer, y As Integer, z As Integer, Achou As Boolean

If Evento = "" Then Exit Sub
If Node.Children > 0 Then Exit Sub

nCodTela = Val(Right$(Node.Parent.Key, 3))
nCodEvento = Val(tvOpt.Nodes(Node.Key).Tag)
If nCodEvento = 0 Then nCodEvento = 1

If Node.Image = 3 Then
     Node.Image = 1
     Achou = False
     For z = 1 To grdTemp.Rows - 1
        If grdTemp.TextMatrix(z, 0) = nCodTela And grdTemp.TextMatrix(z, 2) = nCodEvento Then
            Achou = True
            Exit For
        End If
     Next
     If Not Achou Then
        grdTemp.AddItem nCodTela & Chr(9) & tvOpt.Nodes(Node.Parent.Key).Text & Chr(9) & nCodEvento & Chr(9) & tvOpt.Nodes(Node.Key).Text & Chr(9) & "I"
     Else
        grdTemp.TextMatrix(z, 4) = "I"
     End If
     Sql = "SELECT NOMEOBJETO,ATRIBSEG FROM SEG_EVENTOACESSO WHERE "
     Sql = Sql & "CODTELA=" & nCodTela & " AND "
     Sql = Sql & "CODEVENTO=" & nCodEvento
     Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
     With RdoAux
        Do Until .EOF
            sAtrib = RdoAux!ATRIBSEG
            For x = 1 To grdObj.Rows - 1
                If UCase$(grdObj.TextMatrix(x, 0)) = UCase$(!NOMEOBJETO) Then
'                If grdObj.TextMatrix(x, 0) = !NOMEOBJETO Then
                    grdObj.TextMatrix(x, 6) = "X"
                    If InStr(1, sAtrib, "S", vbBinaryCompare) > 0 Then
                         grdObj.TextMatrix(x, 1) = Val(grdObj.TextMatrix(x, 1)) + 1
                    End If
                    If InStr(1, sAtrib, "I", vbBinaryCompare) > 0 Then
                         grdObj.TextMatrix(x, 2) = Val(grdObj.TextMatrix(x, 2)) + 1
                    End If
                    If InStr(1, sAtrib, "U", vbBinaryCompare) > 0 Then
                         grdObj.TextMatrix(x, 3) = Val(grdObj.TextMatrix(x, 3)) + 1
                    End If
                    If InStr(1, sAtrib, "D", vbBinaryCompare) > 0 Then
                         grdObj.TextMatrix(x, 4) = Val(grdObj.TextMatrix(x, 4)) + 1
                    End If
                    If InStr(1, sAtrib, "E", vbBinaryCompare) > 0 Then
                         grdObj.TextMatrix(x, 5) = Val(grdObj.TextMatrix(x, 5)) + 1
                    End If
                    Exit For
                End If
            Next
           .MoveNext
        Loop
     End With
Else
     Node.Image = 3
Inicio:
     For x = 1 To grdTemp.Rows - 1
        If grdTemp.TextMatrix(x, 0) = nCodTela And grdTemp.TextMatrix(x, 2) = nCodEvento Then
            Sql = "SELECT NOMEOBJETO,ATRIBSEG FROM SEG_EVENTOACESSO WHERE "
            Sql = Sql & "CODTELA=" & nCodTela & " AND "
            Sql = Sql & "CODEVENTO=" & nCodEvento
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
            With RdoAux
               Do Until .EOF
                   sAtrib = RdoAux!ATRIBSEG
                   For y = 1 To grdObj.Rows - 1
                       'If Mid(grdObj.TextMatrix(Y, 0), InStr(1, grdObj.TextMatrix(Y, 0), "..", vbBinaryCompare) + 2, Len(grdObj.TextMatrix(Y, 0)) - InStr(1, grdObj.TextMatrix(Y, 0), "..", vbBinaryCompare) + 2) = !NOMEOBJETO Then
                       If grdObj.TextMatrix(y, 0) = !NOMEOBJETO Then
                           grdObj.TextMatrix(x, 6) = "X"
                           If InStr(1, sAtrib, "S", vbBinaryCompare) > 0 Then
                                grdObj.TextMatrix(y, 1) = Val(grdObj.TextMatrix(y, 1)) - 1
                           End If
                           If InStr(1, sAtrib, "I", vbBinaryCompare) > 0 Then
                                grdObj.TextMatrix(y, 2) = Val(grdObj.TextMatrix(y, 2)) - 1
                           End If
                           If InStr(1, sAtrib, "U", vbBinaryCompare) > 0 Then
                                grdObj.TextMatrix(y, 3) = Val(grdObj.TextMatrix(y, 3)) - 1
                           End If
                           If InStr(1, sAtrib, "D", vbBinaryCompare) > 0 Then
                                grdObj.TextMatrix(y, 4) = Val(grdObj.TextMatrix(y, 4)) - 1
                           End If
                           If InStr(1, sAtrib, "E", vbBinaryCompare) > 0 Then
                                grdObj.TextMatrix(y, 5) = Val(grdObj.TextMatrix(y, 5)) - 1
                           End If
                           Exit For
                       End If
                   Next
                  .MoveNext
               Loop
            End With
            For z = 1 To grdTemp.Rows - 1
               If grdTemp.TextMatrix(z, 0) = nCodTela And grdTemp.TextMatrix(z, 2) = nCodEvento Then
                  Exit For
               End If
            Next
            grdTemp.TextMatrix(z, 4) = "D"
        End If
     Next
End If

End Sub

Private Sub tvUser_NodeClick(ByVal Node As MSComctlLib.Node)

Ocupado
Screen.MousePointer = vbHourglass
With tvUser
    If .SelectedItem.Bold = True Then
       lblTipo.Caption = "GRUPO:"
       lblGrupo.Caption = tvUser.SelectedItem.Text
       lblUser.Caption = tvUser.SelectedItem.Text
       CarregaGrid "G"
    Else
       lblTipo.Caption = "USUARIO:"
       lblGrupo.Caption = tvUser.SelectedItem.Parent
       lblUser.Caption = tvUser.SelectedItem.Text
       CarregaGrid "U"
    End If
End With
Limpa
Le
'MontaAcesso
Liberado
Screen.MousePointer = vbDefault

End Sub

Private Sub Limpa()

For x = 1 To tvOpt.Nodes.Count
      If tvOpt.Nodes(x).Children = 0 Then
           tvOpt.Nodes(x).Image = 3
      End If
Next

End Sub

Private Sub CarregaGrid(sTipo As String)

grdTemp.Rows = 1
For x = 1 To grdObj.Rows - 1
    For y = 1 To 6
        grdObj.TextMatrix(x, y) = ""
    Next
Next


If sTipo = "G" Then
    Sql = "SELECT SEG_GRUPOACESSO.CODTELA,SEG_TELASISTEMA.NOMETELA,SEG_GRUPOACESSO.CODEVENTO,"
    Sql = Sql & "SEG_EVENTO.DESCEVENTO FROM SEG_GRUPOACESSO INNER JOIN SEG_TELASISTEMA ON SEG_GRUPOACESSO.CODTELA = SEG_TELASISTEMA.CODTELA Inner Join "
    Sql = Sql & "SEG_EVENTO ON SEG_GRUPOACESSO.CODEVENTO = SEG_EVENTO.CODEVENTO WHERE NOMEGRUPO = '" & lblUser.Caption & "'"
Else
    Sql = "SELECT SEG_USERACESS.CODTELA, SEG_TELASISTEMA.NOMETELA, SEG_USERACESS.CODEVENTO,SEG_EVENTO.DESCEVENTO "
    Sql = Sql & "FROM SEG_USERACESS INNER JOIN SEG_TELASISTEMA ON SEG_USERACESS.CODTELA = SEG_TELASISTEMA.CODTELA INNER Join "
    Sql = Sql & "SEG_EVENTO ON SEG_USERACESS.CODEVENTO = SEG_EVENTO.CODEVENTO "
    Sql = Sql & "WHERE NOMEUSUARIO = '" & lblUser.Caption & "'"
End If
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    Do Until .EOF
        grdTemp.AddItem !CODTELA & Chr(9) & !NOMETELA & Chr(9) & !CODEVENTO & Chr(9) & !DESCEVENTO & Chr(9) & "N"
        Sql = "SELECT NOMEOBJETO,ATRIBSEG FROM SEG_EVENTOACESSO WHERE CODTELA = " & !CODTELA & " AND CODEVENTO = " & !CODEVENTO
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
        With RdoAux2
            Do Until .EOF
               sAtrib = !ATRIBSEG
               For x = 1 To grdObj.Rows - 1
                   If UCase$(grdObj.TextMatrix(x, 0)) = UCase$(!NOMEOBJETO) Then
                   'If grdObj.TextMatrix(x, 0) = !NOMEOBJETO Then
                      If InStr(1, sAtrib, "S", vbBinaryCompare) > 0 Then
                          grdObj.TextMatrix(x, 1) = Val(grdObj.TextMatrix(x, 1)) + 1
                      End If
                      If InStr(1, sAtrib, "I", vbBinaryCompare) > 0 Then
                          grdObj.TextMatrix(x, 2) = Val(grdObj.TextMatrix(x, 2)) + 1
                      End If
                      If InStr(1, sAtrib, "U", vbBinaryCompare) > 0 Then
                          grdObj.TextMatrix(x, 3) = Val(grdObj.TextMatrix(x, 3)) + 1
                      End If
                      If InStr(1, sAtrib, "D", vbBinaryCompare) > 0 Then
                          grdObj.TextMatrix(x, 4) = Val(grdObj.TextMatrix(x, 4)) + 1
                      End If
                      If InStr(1, sAtrib, "E", vbBinaryCompare) > 0 Then
                          grdObj.TextMatrix(x, 5) = Val(grdObj.TextMatrix(x, 5)) + 1
                      End If
                      Exit For
                   End If
               Next
              .MoveNext
            Loop
        End With
       .MoveNext
    Loop
End With

End Sub

Private Sub Grava()
On Error Resume Next
Dim z As Integer, aUser() As String

'MATRIX PARA USUARIOS DO GRUPO
ReDim aUser(0)

For z = 1 To tvUser.Nodes.Count
    If tvUser.Nodes(z).Parent = tvUser.SelectedItem.Child.Parent Then
       ReDim Preserve aUser(UBound(aUser) + 1)
       aUser(UBound(aUser)) = tvUser.Nodes(z).Text
    End If
Next

For x = 1 To grdTemp.Rows - 1
    If grdTemp.TextMatrix(x, 4) = "N" Then GoTo PROXIMO
    
    If UCase$(lblTipo.Caption) = "GRUPO:" Then
       If grdTemp.TextMatrix(x, 4) = "D" Then
          'APAGA O ACESSO
           Sql = "DELETE FROM SEG_GRUPOACESSO WHERE NOMEGRUPO='" & lblUser.Caption & "' AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
           cn.Execute Sql, rdExecDirect
          'APAGA OS ACESSOS DOS USUARIOS DESTE GRUPO
           For z = 1 To UBound(aUser)
               Sql = "DELETE FROM SEG_USERACESS WHERE NOMEUSUARIO='" & aUser(z) & "' AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
               cn.Execute Sql, rdExecDirect
           Next
       Else
          'GRAVA OS NOVOS ACESSOS DO GRUPO
           Sql = "SELECT CODTELA,CODEVENTO FROM SEG_GRUPOACESSO WHERE NOMEGRUPO='" & lblUser.Caption & "'  AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
                If .RowCount = 0 Then
                    Sql = "INSERT SEG_GRUPOACESSO (NOMEGRUPO,CODTELA,CODEVENTO) VALUES('"
                    Sql = Sql & lblUser.Caption & "'," & grdTemp.TextMatrix(x, 0) & "," & grdTemp.TextMatrix(x, 2) & ")"
                    cn.Execute Sql, rdExecDirect
                End If
               .Close
           End With
          'GRAVA OS NOVOS ACESSOS AOS USUARIOS DO GRUPO
           For z = 1 To UBound(aUser)
               Sql = "SELECT CODTELA,CODEVENTO FROM SEG_USERACESS WHERE NOMEUSUARIO='" & aUser(z) & "' AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                    If .RowCount = 0 Then
                       Sql = "INSERT SEG_USERACESS (NOMEUSUARIO,CODTELA,CODEVENTO) VALUES('"
                       Sql = Sql & aUser(z) & "'," & grdTemp.TextMatrix(x, 0) & "," & grdTemp.TextMatrix(x, 2) & ")"
                       cn.Execute Sql, rdExecDirect
                    End If
                   .Close
               End With
           Next
      End If
    
    Else 'SE FOR USUARIO
      
       If grdTemp.TextMatrix(x, 4) = "D" Then
          'APAGA O ACESSO
           Sql = "DELETE FROM SEG_USERACESS WHERE NOMEUSUARIO='" & lblUser.Caption & "' AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
           cn.Execute Sql, rdExecDirect
       Else
          'GRAVA OS NOVOS ACESSOS DO USUÁRIO
           Sql = "SELECT CODTELA,CODEVENTO FROM SEG_USERACESS WHERE NOMEUSUARIO='" & lblUser.Caption & "' AND CODTELA=" & grdTemp.TextMatrix(x, 0) & " AND CODEVENTO=" & grdTemp.TextMatrix(x, 2)
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
                If .RowCount = 0 Then
                   Sql = "INSERT SEG_USERACESS (NOMEUSUARIO,CODTELA,CODEVENTO) VALUES('"
                   Sql = Sql & lblUser.Caption & "'," & grdTemp.TextMatrix(x, 0) & "," & grdTemp.TextMatrix(x, 2) & ")"
                   cn.Execute Sql, rdExecDirect
                End If
               .Close
           End With
       End If
    End If

PROXIMO:
Next

'GRAVA PERMISSAO NO SQL SERVER
With grdObj
     For x = 1 To .Rows - 1
        If .TextMatrix(x, 6) = "X" Then
 '          GravaAtribSqlServer lblUser.Caption, .TextMatrix(x, 0), IIf(Val(.TextMatrix(x, 1)) > 0, True, False), IIf(Val(.TextMatrix(x, 3)) > 0, True, False), IIf(Val(.TextMatrix(x, 2)) > 0, True, False), IIf(Val(.TextMatrix(x, 4)) > 0, True, False), IIf(Val(.TextMatrix(x, 5)) > 0, True, False)
           'SE FOR GRUPO ATRIBUI A TODOS OS MEMBROS DA MATRIX
           If UCase$(lblTipo.Caption) = "GRUPO:" Then
              For z = 1 To UBound(aUser)
'                  GravaAtribSqlServer aUser(z), .TextMatrix(x, 0), IIf(Val(.TextMatrix(x, 1)) > 0, True, False), IIf(Val(.TextMatrix(x, 3)) > 0, True, False), IIf(Val(.TextMatrix(x, 2)) > 0, True, False), IIf(Val(.TextMatrix(x, 4)) > 0, True, False), IIf(Val(.TextMatrix(x, 5)) > 0, True, False)
              Next
           End If
        End If
     Next
End With

End Sub

Private Sub GravaAtribSqlServer(sUser As String, sObj As String, bSelect As Boolean, bUpdate As Boolean, bInsert As Boolean, bDelete As Boolean, bExec As Boolean)

Dim sAtrib As String, sBase As String
On Error Resume Next

sBase = Left(sObj, InStr(1, sObj, "..", vbBinaryCompare) - 1)

Sql = "REVOKE ALL ON " & sObj & " TO " & sUser
cn.Execute Sql, rdExecDirect

If bSelect Then
     sAtrib = "SELECT,"
End If
If bUpdate Then
     sAtrib = sAtrib & "UPDATE,"
End If
If bInsert Then
     sAtrib = sAtrib & "INSERT,"
End If
If bDelete Then
     sAtrib = sAtrib & "DELETE,"
End If
If bExec Then
     sAtrib = sAtrib & "EXEC,"
End If
'On Error Resume Next
If sAtrib <> "" Then
   sAtrib = Left$(sAtrib, Len(sAtrib) - 1)
   Sql = "USE Tributacao"
   cn.Execute Sql, rdExecDirect
   Sql = "GRANT " & sAtrib & " ON " & sObj & " TO " & sUser
   cn.Execute Sql, rdExecDirect
End If

End Sub

Private Sub Le()
Dim sNomeForm As String
Dim nCodTela As Integer
Dim sEvento As String

For x = 1 To grdTemp.Rows - 1
       Sql = "SELECT NOMEFORM,NOMETELA FROM SEG_TELASISTEMA WHERE CODTELA=" & grdTemp.TextMatrix(x, 0)
       Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
       sNomeForm = RdoAux!NOMETELA
       RdoAux.Close
       nCodTela = grdTemp.TextMatrix(x, 0)
       sEvento = grdTemp.TextMatrix(x, 3)
       For y = 1 To tvOpt.Nodes.Count
            If tvOpt.Nodes(y).Key = sNomeForm & sEvento Then
                  tvOpt.Nodes(y).Image = 1
                 Exit For
            End If
       Next
Next

End Sub


