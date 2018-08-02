VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSisObras 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Integração com o sistema SisobraPref"
   ClientHeight    =   3090
   ClientLeft      =   3825
   ClientTop       =   3975
   ClientWidth     =   7725
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3090
   ScaleWidth      =   7725
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Data de emissão de alvará"
      ForeColor       =   &H00800000&
      Height          =   690
      Left            =   90
      TabIndex        =   14
      Top             =   1755
      Width           =   5235
      Begin VB.CheckBox chkData 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Todos"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   315
         Value           =   1  'Checked
         Width           =   870
      End
      Begin esMaskEdit.esMaskedEdit mskDataIni 
         Height          =   285
         Left            =   2070
         TabIndex        =   16
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmSisObras.frx":0000
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
      Begin esMaskEdit.esMaskedEdit mskDataFim 
         Height          =   285
         Left            =   4050
         TabIndex        =   17
         Top             =   270
         Width           =   1005
         _ExtentX        =   1773
         _ExtentY        =   503
         MouseIcon       =   "frmSisObras.frx":001C
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
      Begin VB.Label lblVenc 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Inicial:"
         Height          =   195
         Index           =   1
         Left            =   1170
         TabIndex        =   19
         Top             =   315
         Width           =   840
      End
      Begin VB.Label lblVenc 
         BackStyle       =   0  'Transparent
         Caption         =   "Data Final:"
         Height          =   195
         Index           =   0
         Left            =   3240
         TabIndex        =   18
         Top             =   315
         Width           =   840
      End
   End
   Begin VB.TextBox txtPwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4635
      PasswordChar    =   "*"
      TabIndex        =   3
      Text            =   "masterkey"
      Top             =   855
      Width           =   1905
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   1575
      TabIndex        =   2
      Text            =   "sysdba"
      Top             =   855
      Width           =   1905
   End
   Begin VB.TextBox txtSep 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2925
      MaxLength       =   1
      TabIndex        =   6
      Text            =   "@"
      Top             =   2655
      Width           =   510
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   1
      Top             =   495
      Width           =   4965
   End
   Begin VB.TextBox txtServidor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1575
      TabIndex        =   0
      Top             =   135
      Width           =   4965
   End
   Begin prjChameleon.chameleonButton cmdConnect 
      Default         =   -1  'True
      Height          =   615
      Left            =   6660
      TabIndex        =   4
      ToolTipText     =   "Conectar ao sistema Sisobra"
      Top             =   135
      Width           =   945
      _ExtentX        =   1667
      _ExtentY        =   1085
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   16777215
      MPTR            =   1
      MICON           =   "frmSisObras.frx":0038
      PICN            =   "frmSisObras.frx":0054
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdTxt 
      Height          =   360
      Left            =   180
      TabIndex        =   5
      ToolTipText     =   "Gerar em arquivo TXT"
      Top             =   2610
      Width           =   1755
      _ExtentX        =   3096
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Gerar emTXT"
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
      MICON           =   "frmSisObras.frx":036E
      PICN            =   "frmSisObras.frx":038A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   4590
      TabIndex        =   13
      Top             =   2700
      Width           =   2940
      _ExtentX        =   5186
      _ExtentY        =   423
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
      Color           =   16777215
      Scrolling       =   1
      ShowText        =   -1  'True
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha...:"
      Height          =   195
      Index           =   3
      Left            =   3825
      TabIndex        =   12
      Top             =   900
      Width           =   870
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Usuário.:"
      Height          =   195
      Index           =   2
      Left            =   90
      TabIndex        =   11
      Top             =   900
      Width           =   1500
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Separador:"
      Height          =   195
      Left            =   2025
      TabIndex        =   10
      Top             =   2700
      Width           =   825
   End
   Begin VB.Label lblMsg 
      BackStyle       =   0  'Transparent
      Caption         =   "Não Conectado"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   240
      Left            =   90
      TabIndex        =   9
      Top             =   1215
      Width           =   6855
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   45
      X2              =   7575
      Y1              =   1620
      Y2              =   1635
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Servidor.:"
      Height          =   195
      Index           =   0
      Left            =   90
      TabIndex        =   8
      Top             =   180
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Banco de Dados..:"
      Height          =   195
      Index           =   1
      Left            =   90
      TabIndex        =   7
      Top             =   540
      Width           =   1500
   End
End
Attribute VB_Name = "frmSisObras"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type OBRA
    CS_TIP_RSP As String
    ID_RESPONSAVEL As String
    NM_RESPONSAVEL As String
    DT_ATUALIZACAO As String
    ID_ALVARA As String
    ID_HABITESE As String
    NM_OBRA As String
    NM_ENDERECO As String
    NU_ENDERECO As String
    VL_AREA_ACRES As String
    VL_AREA_CONSTR As String
    VL_AREA_DEMOL As String
    DT_INICIO As String
    DT_EMISSAO_ALVARA As String
    DT_FIM As String
    DT_ALVARA As String
    VL_AREA_REFORMA As String
    CS_TIP_HABITESE As String
    VL_AREA As String
    USUARIO_HABITESE As String
    USUARIO As String
    ID_OBRA As String
    ID_RESPONSAVEL_SOL As String
    DT_HABITESE As String
    DT_EMISSAO_HABITESE As String
    CS_SITUACAO As String
    CS_CLASSE As String
    ID_ALVARA_ANTERIOR As String
    NU_PROCESSO As String
    NM_ENGENHEIRO As String
    CREA_ENGENHEIRO As String
    NU_ART_OBRA As String
    NU_IND_FISCAL As String
    NM_ENGENHEIRO_PROJ As String
    CREA_ENGENHEIRO_PROJ As String
    NU_ART_PROJETO As String
    TX_ESPEC_OBRA As String
    TX_OBS As String
End Type

Dim adoConn As New ADODB.Connection, aObra() As OBRA





Private Sub cmdConnect_Click()
Dim LoginDSN As String

On Error GoTo Erro:

Ocupado
lblMsg.Caption = "Tentando estabelecer uma conexão...Aguarde..."
lblMsg.ForeColor = &H800000
Me.Refresh

'adoConn.Open "provider=LCPI.IBProvider;data source=" & txtServidor.Text & ":" & txtDB.Text, "sysdba", "masterkey"

adoConn.ConnectionString = _
"DRIVER=Firebird/InterBase(r) driver;UID=" & txtUser.Text & ";PWD=" & txtPwd.Text & ";DBNAME=" & txtServidor.Text & "\" & txtDB.Text
adoConn.Open

LoginDSN = "SisObra"
'Set cnIB = en.OpenConnection(dsname:=LoginDSN, _
         Prompt:=rdDriverNoPrompt, _
        Connect:="DRIVER=Firebird/InterBase(r) driver;SERVER=" & txtServidor.Text & ";DATABASE=" & txtServidor.Text & ":" & txtDB.Text & ";UID=" & txtUser.Text & ";PWD=" & txtPwd & "")


'If cnIB.rdoTables.Count > 0 Then

If adoConn.State = 1 Then
    lblMsg.Caption = "Sistema conectado com sucesso."
    lblMsg.ForeColor = &H8000&
    cmdConnect.Enabled = False
    txtServidor.Locked = True
    txtDB.Locked = True
    txtServidor.BackColor = Kde
    txtDB.BackColor = Kde
    Me.Refresh
    Liberado
Else
    Liberado
    lblMsg.Caption = "Erro na Conexão !!!"
    lblMsg.ForeColor = &HC0&
    Me.Refresh
    MsgBox Err.Description
End If

Exit Sub
Erro:
Liberado
lblMsg.Caption = "Erro na Conexão !!!"
lblMsg.ForeColor = &HC0&
Me.Refresh
MsgBox Err.Description

End Sub

Private Sub cmdTxt_Click()

If adoConn.State = 0 Then
    MsgBox "Não esta conectado!", vbCritical, "Atenção"
    Exit Sub
End If

If txtSep.Text = "" Then
    MsgBox "Selecione um separador de campos!", vbCritical, "Atenção"
    Exit Sub
End If

GeraTxt


End Sub



Private Sub Form_Load()
Centraliza Me
lblMsg.Caption = "Não Conectado"
lblMsg.ForeColor = &HC0&
txtServidor.Text = GetSetting("GTI", "OPTION", "IBSERVER")
txtDB.Text = GetSetting("GTI", "OPTION", "IBDB")
'txtUser.text = GetSetting("GTI", "OPTION", "IBUSER")

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If adoConn.State = 1 Then
    adoConn.Close
End If
End Sub

Private Sub txtDB_LostFocus()
SaveSetting "GTI", "OPTION", "IBDB", txtDB.Text
End Sub

Private Sub txtUser_LostFocus()
SaveSetting "GTI", "OPTION", "IBUSER", txtUser.Text
End Sub

Private Sub txtServidor_LostFocus()
SaveSetting "GTI", "OPTION", "IBSERVER", txtServidor.Text
End Sub
 
Private Sub GeraTxt()
Dim ax As String, ind As Integer, sResp As String, sSep As String, Sql As String
Dim rst As New ADODB.Recordset, rst2 As New ADODB.Recordset
Ocupado
DoEvents
ReDim aObra(0)
PBar.Value = 0
sSep = txtSep.Text
adoConn.BeginTrans
With rst
    Sql = "SELECT * FROM OBRA WHERE NM_OBRA IS NOT NULL "
    If chkData.Value = vbUnchecked Then
        Sql = Sql & " AND DT_ALVARA BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' AND '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
    End If
    Sql = Sql & " ORDER BY ID_OBRA"
   .Source = Sql
   .ActiveConnection = adoConn
   .CursorType = adOpenKeyset
   .Open
    Me.MousePointer = vbHourglass
    Do Until .EOF
        ind = UBound(aObra) + 1
        If ind Mod 30 = 0 Then
           CallPb CLng(ind), CLng(.RecordCount)
        End If
        ReDim Preserve aObra(ind)
        aObra(ind).ID_ALVARA = SubNull(!ID_ALVARA)
        aObra(ind).NM_OBRA = SubNull(!NM_OBRA)
        aObra(ind).NM_ENDERECO = SubNull(!NM_ENDERECO)
        aObra(ind).NU_ENDERECO = SubNull(!NU_ENDERECO)
        aObra(ind).VL_AREA_CONSTR = SubNull(!VL_AREA_CONSTR)
        aObra(ind).VL_AREA_ACRES = SubNull(!VL_AREA_ACRES)
        aObra(ind).ID_ALVARA = SubNull(!ID_ALVARA)
        aObra(ind).DT_INICIO = SubNull(!DT_INICIO)
        aObra(ind).DT_EMISSAO_ALVARA = SubNull(!DT_EMISSAO_ALVARA)
        aObra(ind).DT_FIM = SubNull(!DT_FIM)
        aObra(ind).DT_ALVARA = SubNull(!DT_ALVARA)
        aObra(ind).VL_AREA_REFORMA = SubNull(!VL_AREA_REFOR)
        aObra(ind).USUARIO = SubNull(!USUARIO)
        aObra(ind).ID_OBRA = SubNull(!ID_OBRA)
        aObra(ind).ID_RESPONSAVEL_SOL = SubNull(!ID_RESPONSAVEL_SOL)
        aObra(ind).CS_SITUACAO = SubNull(!CS_SITUACAO)
        aObra(ind).CS_CLASSE = SubNull(!CS_CLASSE)
        aObra(ind).ID_ALVARA_ANTERIOR = SubNull(!ID_ALVARA_ANTERIOR)
        aObra(ind).NU_PROCESSO = SubNull(!NU_PROCESSO)
        aObra(ind).NM_ENGENHEIRO = SubNull(!NM_ENGENHEIRO)
        aObra(ind).CREA_ENGENHEIRO = SubNull(!CREA_ENGENHEIRO)
        aObra(ind).NU_ART_OBRA = SubNull(!NU_ART_OBRA)
        aObra(ind).NU_IND_FISCAL = SubNull(!NU_IND_FISCAL)
        aObra(ind).NM_ENGENHEIRO_PROJ = SubNull(!NM_ENGENHEIRO_PROJ)
        aObra(ind).CREA_ENGENHEIRO_PROJ = SubNull(!CREA_ENGENHEIRO_PROJ)
        aObra(ind).NU_ART_PROJETO = SubNull(!NU_ART_PROJETO)
        aObra(ind).TX_ESPEC_OBRA = SubNull(!TX_ESPEC_OBRA)
        aObra(ind).TX_OBS = SubNull(!TX_OBS)
        DoEvents
        With rst2
           .Source = "SELECT ID_RESPONSAVEL FROM VINCULO WHERE ID_OBRA=" & rst!ID_OBRA
           .ActiveConnection = adoConn
           .Open
            sResp = !ID_RESPONSAVEL
           .Close
           
           .Source = "SELECT * FROM RESPONSAVEL WHERE ID_RESPONSAVEL=" & sResp
           .ActiveConnection = adoConn
           .Open
           aObra(ind).CS_TIP_RSP = SubNull(!CS_TIP_RESP)
           aObra(ind).ID_RESPONSAVEL = SubNull(!ID_RESPONSAVEL)
           aObra(ind).NM_RESPONSAVEL = SubNull(!NM_RESPONSAVEL)
           aObra(ind).DT_ATUALIZACAO = SubNull(!DT_ATUALIZACAO)
           .Close
        
           .Source = "SELECT * FROM HABITESE WHERE ID_OBRA=" & aObra(ind).ID_OBRA
           .ActiveConnection = adoConn
           .Open
           aObra(ind).ID_HABITESE = SubNull(!ID_HABITESE)
           aObra(ind).CS_TIP_HABITESE = SubNull(!CS_TIP_HABITESE)
           aObra(ind).USUARIO_HABITESE = SubNull(!USUARIO_HABITESE)
           aObra(ind).VL_AREA = SubNull(!Area)
           aObra(ind).DT_HABITESE = SubNull(!DT_HABITESE)
           aObra(ind).DT_EMISSAO_HABITESE = Format(SubNull(!DT_EMISSAO_HABITESE), "dd/mm/yyyy")
           .Close
        
        End With
        
       .MoveNext
    Loop
   .Close
End With
FINALIZA:
adoConn.CommitTrans
Liberado
Open sPathBin & "\RELSISOBRA.CSV" For Output As #1
ax = "Tipo resp" & sSep & "CPF/CNPJ" & sSep & "Nome responsável" & sSep & "Data Atualização" & sSep & "No Alvará" & sSep & "No Habite-se" & sSep & "Nome obra" & sSep
ax = ax & "Endereço" & sSep & "Numero" & sSep & "Área acres." & sSep & "Área constr." & sSep & "Área demolida" & sSep & "Data início" & sSep & "Data emissão alvará" & sSep & "Data final" & sSep
ax = ax & "Data alvará" & sSep & "Área reformada" & sSep & "Tipo habite-se" & sSep & "Área" & sSep & "Usuário habite-se" & sSep & "Usuário" & sSep & "No Obra" & sSep
ax = ax & "No responsável" & sSep & "Data habite-se" & sSep & "Data emissão habite-se" & sSep & "Situação" & sSep & "Classe" & sSep & "Alvará anterior" & sSep
ax = ax & "No Processo" & sSep & "Nome engenheiro" & sSep & "CREA Engenheiro" & sSep & "Num Art Obra" & sSep & "Num  Ind Fiscal" & sSep & "Nome Engenheiro Proj" & sSep
ax = ax & "CREA Engenheiro Proj" & sSep & "Num Art Proj" & sSep & "Espec Obra" & sSep & "Observações"
Print #1, ax
For ind = 1 To UBound(aObra)
    With aObra(ind)
        ax = .CS_TIP_RSP & sSep & .ID_RESPONSAVEL & sSep & .NM_RESPONSAVEL & sSep & .DT_ATUALIZACAO & sSep & .ID_ALVARA & sSep & .ID_HABITESE & sSep & .NM_OBRA & sSep
        ax = ax & .NM_ENDERECO & sSep & .NU_ENDERECO & sSep & .VL_AREA_ACRES & sSep & .VL_AREA_CONSTR & sSep & .VL_AREA_DEMOL & sSep & .DT_INICIO & sSep & .DT_EMISSAO_ALVARA & sSep & .DT_FIM & sSep
        ax = ax & .DT_ALVARA & sSep & .VL_AREA_REFORMA & sSep & .CS_TIP_HABITESE & sSep & .VL_AREA & sSep & .USUARIO_HABITESE & sSep & .USUARIO & sSep & .ID_OBRA & sSep
        ax = ax & .ID_RESPONSAVEL & sSep & .DT_HABITESE & sSep & .DT_EMISSAO_HABITESE & sSep & .CS_SITUACAO & sSep & .CS_CLASSE & sSep & .ID_ALVARA_ANTERIOR & sSep
        ax = ax & .NU_PROCESSO & sSep & .NM_ENGENHEIRO & sSep & .CREA_ENGENHEIRO & sSep & .NU_ART_OBRA & sSep & .NU_IND_FISCAL & sSep & .NM_ENGENHEIRO_PROJ & sSep
        ax = ax & .CREA_ENGENHEIRO_PROJ & sSep & .NU_ART_PROJETO & sSep & .TX_ESPEC_OBRA & sSep & .TX_OBS
    End With
    Print #1, ax
Next
Close #1
Liberado
PBar.Value = 0
MsgBox "O arquivo foi gerado em: " & sPathBin & "\RELSISOBRA.CSV", vbInformation, "Atenção"
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

