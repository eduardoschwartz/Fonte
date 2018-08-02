VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmConversorSisObraPref 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Conversor do Sistema SisObraPref"
   ClientHeight    =   5415
   ClientLeft      =   3600
   ClientTop       =   2670
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5415
   ScaleWidth      =   7830
   Begin VB.ListBox lstLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      ForeColor       =   &H0080FFFF&
      Height          =   3150
      Left            =   90
      TabIndex        =   10
      Top             =   2160
      Width           =   7620
   End
   Begin VB.TextBox txtServidor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      TabIndex        =   4
      Text            =   "localhost"
      Top             =   180
      Width           =   4965
   End
   Begin VB.TextBox txtDB 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1620
      TabIndex        =   3
      Text            =   "C:\Program Files\Dataprev\SisobraPref\Data\Sisobra.idb"
      Top             =   540
      Width           =   4965
   End
   Begin VB.TextBox txtUser 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      Left            =   1620
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "sysdba"
      Top             =   900
      Visible         =   0   'False
      Width           =   1905
   End
   Begin VB.TextBox txtPwd 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   4680
      Locked          =   -1  'True
      PasswordChar    =   "*"
      TabIndex        =   1
      Text            =   "masterkey"
      Top             =   900
      Visible         =   0   'False
      Width           =   1905
   End
   Begin prjChameleon.chameleonButton cmdConnect 
      Height          =   615
      Left            =   6705
      TabIndex        =   0
      ToolTipText     =   "Conectar ao sistema Sisobra"
      Top             =   270
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
      MICON           =   "frmConversorSisObraPref.frx":0000
      PICN            =   "frmConversorSisObraPref.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdOK 
      Default         =   -1  'True
      Height          =   345
      Left            =   90
      TabIndex        =   11
      ToolTipText     =   "Conectar ao sistema Sisobra"
      Top             =   1710
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Iniciar Conversão"
      ENAB            =   0   'False
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
      MICON           =   "frmConversorSisObraPref.frx":0336
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   2
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin Tributacao.XP_ProgressBar PBar 
      Height          =   240
      Left            =   3915
      TabIndex        =   12
      Top             =   1800
      Width           =   3795
      _ExtentX        =   6694
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
      Caption         =   "Banco de Dados..:"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   1
      Left            =   135
      TabIndex        =   9
      Top             =   585
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Servidor.:"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   0
      Left            =   135
      TabIndex        =   8
      Top             =   225
      Width           =   1500
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      BorderWidth     =   2
      X1              =   90
      X2              =   7620
      Y1              =   1530
      Y2              =   1545
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
      ForeColor       =   &H000000FF&
      Height          =   240
      Left            =   135
      TabIndex        =   7
      Top             =   1215
      Width           =   6855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Usuário.:"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   2
      Left            =   135
      TabIndex        =   6
      Top             =   945
      Visible         =   0   'False
      Width           =   1500
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha...:"
      ForeColor       =   &H00C0FFFF&
      Height          =   195
      Index           =   3
      Left            =   3870
      TabIndex        =   5
      Top             =   945
      Visible         =   0   'False
      Width           =   870
   End
End
Attribute VB_Name = "frmConversorSisObraPref"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim cnIB As New rdoConnection

Private Sub cmdConnect_Click()
Dim LoginDSN As String

On Error GoTo Erro
LoginDSN = "SisObra"
Set cnIB = en.OpenConnection(dsname:=LoginDSN, _
         Prompt:=rdDriverNoPrompt, _
        Connect:="DRIVER=Firebird/InterBase(r) driver;SERVER=localhost;DATABASE=" & txtServidor.Text & ":" & txtDB.Text & ";UID=" & txtUser.Text & ";PWD=" & txtPwd & "")


If cnIB.rdoTables.Count > 0 Then
    lblMsg.Caption = "Sistema conectado com sucesso."
    lblMsg.ForeColor = &HFF00&
    cmdConnect.Enabled = False
    txtServidor.Locked = True
    txtDB.Locked = True
    txtServidor.BackColor = Kde
    txtDB.BackColor = Kde
    cmdOK.Enabled = True
    lstLog.Clear
    Me.Refresh
    Liberado
Else
    Liberado
    lblMsg.Caption = "Erro na Conexão !!!"
    lblMsg.ForeColor = &HFF&
    cmdOK.Enabled = False
    Me.Refresh
    MsgBox Err.Description
End If

Exit Sub
Erro:
Liberado
lblMsg.Caption = "Erro na Conexão !!!"
lblMsg.ForeColor = &HFF&
cmdOK.Enabled = False
Me.Refresh
MsgBox Err.Description


End Sub

Private Sub cmdOK_Click()
Dim Sql As String, RdoIB As rdoResultset, nPos As Long, nTotal As Long
Ocupado
cmdOK.Enabled = False

Lg "Deleting table [so_obra]"

Sql = "IF EXISTS (SELECT 1 From INFORMATION_SCHEMA.Tables WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='SO_OBRA') "
Sql = Sql & "DROP TABLE SO_OBRA"
cn.Execute Sql, rdExecDirect

Lg "Creating table [so_obra]"
Sql = "CREATE TABLE so_obra(id_obra int PRIMARY KEY not null,id_alvara VARCHAR(15),dt_alvara SMALLDATETIME,nm_obra VARCHAR(55),id_tipo_lgdr varchar(10),"
Sql = Sql & "nm_endereco varchar(55),nu_endereco varchar(10),compl_endereco varchar(10),nm_bairro varchar(20),nu_cep varchar(8),"
Sql = Sql & "id_municipio varchar(6),nm_cidade varchar(30),cs_uf char(2),te_email varchar(60),id_tipolgdr_corr varchar(10),"
Sql = Sql & "nm_endereco_corr varchar(55),nu_endereco_corr varchar(10),nm_bairro_corr varchar(20),compl_endereco_corr varchar(10),"
Sql = Sql & "nu_cep_corr varchar(8),id_municipio_corr varchar(6),nu_telefone varchar(14),nu_fax varchar(14),dt_atualizacao smalldatetime,"
Sql = Sql & "in_transmitido varchar(1),dt_transmissao smalldatetime,cs_tipo_tx varchar(1),cs_situacao varchar(70),cs_classe varchar(10),"
Sql = Sql & "id_alvara_anterior varchar(15),dt_inicio smalldatetime,dt_fim smalldatetime,nu_art_projeto varchar(20),nu_art_obra varchar(20),"
Sql = Sql & "nm_engenheiro varchar(55),nu_ind_fiscal varchar(15),nu_processo varchar(25),tx_espec_obra varchar(4096),tx_obs varchar(4096),"
Sql = Sql & "cs_uf_corr char(2),cs_tip_resp_sol char(1),id_responsavel_sol varchar(14),nm_responsavel_sol varchar(55),id_tipo_lgdr_sol varchar(10),"
Sql = Sql & "nm_endereco_sol varchar(55),nu_endereco_sol varchar(10),compl_endereco_sol varchar(10),nm_bairro_sol varchar(20),"
Sql = Sql & "nu_cep_sol varchar(8),id_municipio_sol varchar(6),nm_cidade_sol varchar(30),cs_uf_sol char(2),cs_tip_ocup_constr int,"
Sql = Sql & "cs_tip_constr_constr int,vl_area_constr float,cs_tip_ocup_demol int,cs_tip_constr_demol int,vl_area_demol float,"
Sql = Sql & "cs_tip_ocup_acres int,cs_tip_constr_acres int ,vl_area_acres float,vl_area_existente float,cs_tip_ocup_refor int,"
Sql = Sql & "cs_tip_constr_refor int,vl_area_refor float,nm_engenheiro_proj varchar(55),crea_engenheiro varchar(20),crea_engenheiro_proj varchar(20),"
Sql = Sql & "usuario varchar(8),dt_emissao_alvara smalldatetime,nu_pavimento smallint,nu_unidade smallint,nu_ddd_tel varchar(4),nu_ddd_fax varchar(4) )"
cn.Execute Sql, rdExecDirect

Lg "Granting permission to table [so_obra]"
Sql = "GRANT SELECT,UPDATE,INSERT,DELETE ON SO_OBRA TO PUBLIC"
cn.Execute Sql, rdExecDirect

Lg "Converting Table [so_obra] to Sql Server"

Sql = "SELECT * FROM OBRA"
Set RdoIB = cnIB.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoIB
    nPos = 1: nTotal = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "INSERT SO_OBRA(ID_OBRA,ID_ALVARA,DT_ALVARA,NM_OBRA,ID_TIPO_LGDR,NM_ENDERECO,NU_ENDERECO,COMPL_ENDERECO,NM_BAIRRO,NU_CEP,ID_MUNICIPIO,"
        Sql = Sql & "NM_CIDADE,CS_UF,TE_EMAIL,ID_TIPOLGDR_CORR,NM_ENDERECO_CORR,NU_ENDERECO_CORR,NM_BAIRRO_CORR,COMPL_ENDERECO_CORR,NU_CEP_CORR,"
        Sql = Sql & "ID_MUNICIPIO_CORR,NU_TELEFONE,NU_FAX,DT_ATUALIZACAO,IN_TRANSMITIDO,DT_TRANSMISSAO,CS_TIPO_TX,CS_SITUACAO,CS_CLASSE,ID_ALVARA_ANTERIOR,"
        Sql = Sql & "DT_INICIO,DT_FIM,NU_ART_PROJETO,NU_ART_OBRA,NM_ENGENHEIRO,NU_IND_FISCAL,NU_PROCESSO,TX_ESPEC_OBRA,TX_OBS,CS_UF_CORR,CS_TIP_RESP_SOL,"
        Sql = Sql & "ID_RESPONSAVEL_SOL,NM_RESPONSAVEL_SOL,ID_TIPO_LGDR_SOL,NM_ENDERECO_SOL,NU_ENDERECO_SOL,COMPL_ENDERECO_SOL,NM_BAIRRO_SOL,NU_CEP_SOL,"
        Sql = Sql & "ID_MUNICIPIO_SOL,NM_CIDADE_SOL,CS_UF_SOL,CS_TIP_OCUP_CONSTR,CS_TIP_CONSTR_CONSTR,VL_AREA_CONSTR,CS_TIP_OCUP_DEMOL,CS_TIP_CONSTR_DEMOL,"
        Sql = Sql & "VL_AREA_DEMOL,CS_TIP_OCUP_ACRES,CS_TIP_CONSTR_ACRES,VL_AREA_ACRES,VL_AREA_EXISTENTE,CS_TIP_OCUP_REFOR,CS_TIP_CONSTR_REFOR,VL_AREA_REFOR,"
        Sql = Sql & "NM_ENGENHEIRO_PROJ,CREA_ENGENHEIRO,CREA_ENGENHEIRO_PROJ,USUARIO,DT_EMISSAO_ALVARA,NU_PAVIMENTO,NU_UNIDADE,NU_DDD_TEL,NU_DDD_FAX"
        Sql = Sql & ") VALUES("
        Sql = Sql & !ID_OBRA & ",'" & !ID_ALVARA & "','" & Format(!DT_ALVARA, "mm/dd/yyyy") & "','" & Mask(!NM_OBRA) & "','" & SubNull(!ID_TIPO_LGDR) & "','"
        Sql = Sql & Mask(SubNull(!NM_ENDERECO)) & "','" & !NU_ENDERECO & "','" & Mask(SubNull(!COMPL_ENDERECO)) & "','" & Mask(SubNull(!NM_BAIRRO)) & "','"
        Sql = Sql & Mask(SubNull(!NU_CEP)) & "','" & Mask(SubNull(!ID_MUNICIPIO)) & "','" & Mask(SubNull(!NM_CIDADE)) & "','" & Mask(SubNull(!CS_UF)) & "','"
        Sql = Sql & Mask(SubNull(!TE_EMAIL)) & "','" & Mask(SubNull(!ID_TIPO_LGDR_CORR)) & "','" & Mask(SubNull(!NM_ENDERECO_CORR)) & "','" & Mask(SubNull(!NU_ENDERECO_CORR)) & "','"
        Sql = Sql & Mask(SubNull(!NM_BAIRRO_CORR)) & "','" & Mask(SubNull(!COMPL_ENDERECO_CORR)) & "','" & Mask(SubNull(!NU_CEP_CORR)) & "','" & Mask(SubNull(!ID_MUNICIPIO_CORR)) & "','"
        Sql = Sql & SubNull(!NU_TELEFONE) & "','" & SubNull(!NU_FAX) & "','" & Format(!DT_ATUALIZACAO, "mm/dd/yyyy") & "','" & SubNull(!IN_TRANSMITIDO) & "','" & Format(!DT_TRANSMISSAO, "mm/dd/yyyy") & "','"
        Sql = Sql & SubNull(!CS_TIPO_TX) & "','" & SubNull(!CS_SITUACAO) & "','" & SubNull(!CS_CLASSE) & "','" & SubNull(!ID_ALVARA_ANTERIOR) & "','" & Format(!DT_INICIO, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(!DT_FIM, "mm/dd/yyyy") & "','" & SubNull(!NU_ART_PROJETO) & "','" & SubNull(!NU_ART_OBRA) & "','" & Mask(SubNull(!NM_ENGENHEIRO)) & "','" & SubNull(!NU_IND_FISCAL) & "','"
        Sql = Sql & Mask(SubNull(!NU_PROCESSO)) & "','" & Mask(SubNull(!TX_ESPEC_OBRA)) & "','" & Mask(SubNull(!TX_OBS)) & "','" & SubNull(!cs_uf_corr) & "','"
        Sql = Sql & Mask(SubNull(!CS_TIP_RESP_SOL)) & "','" & Mask(SubNull(!ID_RESPONSAVEL_SOL)) & "','" & Mask(SubNull(!NM_RESPONSAVEL_SOL)) & "','" & Mask(SubNull(!ID_TIPO_LGDR_SOL)) & "','"
        Sql = Sql & Mask(SubNull(!NM_ENDERECO_SOL)) & "','" & Mask(SubNull(!NU_ENDERECO_SOL)) & "','" & Mask(SubNull(!COMPL_ENDERECO_SOL)) & "','" & Mask(SubNull(!NM_BAIRRO_SOL)) & "','"
        Sql = Sql & Mask(SubNull(!NU_CEP_SOL)) & "','" & SubNull(!ID_MUNICIPIO_SOL) & "','" & Mask(SubNull(!NM_CIDADE_SOL)) & "','" & SubNull(!CS_UF_SOL) & "',"
        Sql = Sql & Val(SubNull(!CS_TIP_OCUP_CONSTR)) & "," & Val(SubNull(!CS_TIP_CONSTR_CONSTR)) & "," & Virg2Ponto((!VL_AREA_CONSTR)) & "," & Val(SubNull(!CS_TIP_OCUP_DEMOL)) & ","
        Sql = Sql & Val(SubNull(!CS_TIP_CONSTR_DEMOL)) & "," & IIf(IsNull(!VL_AREA_DEMOL), "Null", Virg2Ponto(SubNull(!VL_AREA_DEMOL))) & "," & Val(SubNull(!CS_TIP_OCUP_ACRES)) & "," & Val(SubNull(!CS_TIP_CONSTR_ACRES)) & ","
        Sql = Sql & Virg2Ponto((!VL_AREA_ACRES)) & "," & IIf(IsNull(!VL_AREA_EXISTENTE), "Null", Virg2Ponto(SubNull(!VL_AREA_EXISTENTE))) & "," & Val(SubNull(!CS_TIP_OCUP_REFOR)) & "," & Val(SubNull(!CS_TIP_CONSTR_REFOR)) & ","
        Sql = Sql & Virg2Ponto((!VL_AREA_REFOR)) & ",'" & Mask(SubNull(!NM_ENGENHEIRO_PROJ)) & "','" & Mask(SubNull(!CREA_ENGENHEIRO)) & "','" & Mask(SubNull(!CREA_ENGENHEIRO_PROJ)) & "','"
        Sql = Sql & Mask(SubNull(!USUARIO)) & "','" & Format(!DT_EMISSAO_ALVARA, "mm/dd/yyyy") & "'," & Val(SubNull(!NU_PAVIMENTO)) & "," & Val(SubNull(!NU_UNIDADE)) & ",'"
        Sql = Sql & Mask(SubNull(!NU_DDD_TEL)) & "','" & Mask(SubNull(!NU_DDD_FAX))
        Sql = Sql & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Lg "Setting null values"
Sql = "UPDATE SO_OBRA SET ID_TIPO_LGDR=NULL WHERE ID_TIPO_LGDR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET COMPL_ENDERECO=NULL WHERE COMPL_ENDERECO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_BAIRRO=NULL WHERE NM_BAIRRO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_CEP=NULL WHERE NU_CEP=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET TE_EMAIL=NULL WHERE TE_EMAIL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_TIPOLGDR_CORR=NULL WHERE ID_TIPOLGDR_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_ENDERECO_CORR=NULL WHERE NM_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_ENDERECO_CORR=NULL WHERE NU_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_BAIRRO_CORR=NULL WHERE NM_BAIRRO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET COMPL_ENDERECO_CORR=NULL WHERE COMPL_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_CEP_CORR=NULL WHERE NU_CEP_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_MUNICIPIO_CORR=NULL WHERE ID_MUNICIPIO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_TELEFONE=NULL WHERE NU_TELEFONE=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_FAX=NULL WHERE NU_FAX=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET IN_TRANSMITIDO=NULL WHERE IN_TRANSMITIDO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET DT_TRANSMISSAO=NULL WHERE DT_TRANSMISSAO='01/01/1900'"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_TIPO_TX=NULL WHERE CS_TIPO_TX=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_SITUACAO=NULL WHERE CS_SITUACAO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_CLASSE=NULL WHERE CS_CLASSE=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_ALVARA_ANTERIOR=NULL WHERE ID_ALVARA_ANTERIOR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET DT_INICIO=NULL WHERE DT_INICIO='01/01/1900'"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET DT_FIM=NULL WHERE DT_FIM='01/01/1900'"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_ART_PROJETO=NULL WHERE NU_ART_PROJETO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_ART_OBRA=NULL WHERE NU_ART_OBRA=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_ENGENHEIRO=NULL WHERE NM_ENGENHEIRO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_IND_FISCAL=NULL WHERE NU_IND_FISCAL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_PROCESSO=NULL WHERE NU_PROCESSO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET TX_ESPEC_OBRA=NULL WHERE TX_ESPEC_OBRA=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET TX_OBS=NULL WHERE TX_OBS=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_UF_CORR=NULL WHERE CS_UF_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_TIP_RESP_SOL=NULL WHERE CS_TIP_RESP_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_RESPONSAVEL_SOL=NULL WHERE ID_RESPONSAVEL_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_RESPONSAVEL_SOL=NULL WHERE NM_RESPONSAVEL_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_TIPO_LGDR_SOL=NULL WHERE ID_TIPO_LGDR_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_ENDERECO_SOL=NULL WHERE NM_ENDERECO_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_ENDERECO_SOL=NULL WHERE NU_ENDERECO_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET COMPL_ENDERECO_SOL=NULL WHERE COMPL_ENDERECO_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_BAIRRO_SOL=NULL WHERE NM_BAIRRO_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_CEP_SOL=NULL WHERE NU_CEP_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET ID_MUNICIPIO_SOL=NULL WHERE ID_MUNICIPIO_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_CIDADE_SOL=NULL WHERE NM_CIDADE_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CS_UF_SOL=NULL WHERE CS_UF_SOL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NM_ENGENHEIRO_PROJ=NULL WHERE NM_ENGENHEIRO_PROJ=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CREA_ENGENHEIRO=NULL WHERE CREA_ENGENHEIRO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET CREA_ENGENHEIRO_PROJ=NULL WHERE CREA_ENGENHEIRO_PROJ=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_DDD_TEL=NULL WHERE NU_DDD_TEL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_OBRA SET NU_DDD_FAX=NULL WHERE NU_DDD_FAX=''"
cn.Execute Sql, rdExecDirect

HABITESE:
Lg "Deleting table [so_habitese]"

Sql = "IF EXISTS (SELECT 1 From INFORMATION_SCHEMA.Tables WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='SO_HABITESE') "
Sql = Sql & "DROP TABLE SO_HABITESE"
cn.Execute Sql, rdExecDirect

Lg "Creating table [so_habitese]"
Sql = "CREATE TABLE so_habitese(id_habitese varchar(15) PRIMARY KEY not null,dt_habitese smalldatetime,id_obra int,area float,sequencial_habitese smallint,"
Sql = Sql & "cs_tip_habitese varchar(1),usuario_habitese varchar(8),dt_emissao_habitese smalldatetime,tx_obs varchar(4096))"
cn.Execute Sql, rdExecDirect

Lg "Granting permission to table [so_habitese]"
Sql = "GRANT SELECT,UPDATE,INSERT,DELETE ON SO_HABITESE TO PUBLIC"
cn.Execute Sql, rdExecDirect

Lg "Converting Table [so_habitese] to Sql Server"

Sql = "SELECT * FROM HABITESE"
Set RdoIB = cnIB.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoIB
    nPos = 1: nTotal = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "INSERT SO_HABITESE(ID_HABITESE,DT_HABITESE,ID_OBRA,AREA,SEQUENCIAL_HABITESE,CS_TIP_HABITESE,USUARIO_HABITESE,DT_EMISSAO_HABITESE,TX_OBS"
        Sql = Sql & ") VALUES('"
        Sql = Sql & !ID_HABITESE & "','" & Format(!DT_HABITESE, "mm/dd/yyyy") & "'," & Val(SubNull(!ID_OBRA)) & "," & Virg2Ponto(!Area) & ","
        Sql = Sql & Val(SubNull(!SEQUENCIAL_HABITESE)) & "," & Val(SubNull(!CS_TIP_HABITESE)) & ",'" & Mask(SubNull(!USUARIO_HABITESE)) & "','" & Format(!DT_EMISSAO_HABITESE, "mm/dd/yyyy") & "','"
        Sql = Sql & Mask(SubNull(!TX_OBS)) & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

VINCULO:
Lg "Deleting table [so_vinculo]"

Sql = "IF EXISTS (SELECT 1 From INFORMATION_SCHEMA.Tables WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='SO_VINCULO') "
Sql = Sql & "DROP TABLE SO_VINCULO"
cn.Execute Sql, rdExecDirect

Lg "Creating table [so_vinculo]"
Sql = "CREATE TABLE so_vinculo(id_responsavel varchar(14) not null,id_obra int not null,cs_tipo_responsavel char(1),dt_inicio_periodo smalldatetime,"
Sql = Sql & "dt_fim_periodo smalldatetime,tp_qualific_inss varchar(2),dt_atualizacao smalldatetime,seq_vinc smallint,in_transmitido varchar(1), "
Sql = Sql & "CONSTRAINT [PK_so_vinculo] PRIMARY KEY CLUSTERED ([id_responsavel] ASC,[ID_OBRA] Asc) WITH (PAD_INDEX  = OFF, IGNORE_DUP_KEY = OFF) ON [PRIMARY]) ON [PRIMARY]"
cn.Execute Sql, rdExecDirect

Lg "Granting permission to table [so_vinculo]"
Sql = "GRANT SELECT,UPDATE,INSERT,DELETE ON SO_VINCULO TO PUBLIC"
cn.Execute Sql, rdExecDirect

Lg "Converting Table [so_vinculo] to Sql Server"

Sql = "SELECT * FROM VINCULO"
Set RdoIB = cnIB.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoIB
    nPos = 1: nTotal = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "INSERT SO_VINCULO(ID_RESPONSAVEL,ID_OBRA,CS_TIPO_RESPONSAVEL,DT_INICIO_PERIODO,DT_FIM_PERIODO,TP_QUALIFIC_INSS,DT_ATUALIZACAO,SEQ_VINC,IN_TRANSMITIDO"
        Sql = Sql & ") VALUES('"
        Sql = Sql & !ID_RESPONSAVEL & "'," & !ID_OBRA & ",'" & SubNull(!CS_TIPO_RESPONSAVEL) & "','" & Format(!DT_INICIO_PERIODO, "mm/dd/yyyy") & "','"
        Sql = Sql & Format(!DT_FIM_PERIODO, "mm/dd/yyyy") & "','" & SubNull(!TP_QUALIFIC_INSS) & "','" & Format(!DT_ATUALIZACAO, "mm/dd/yyyy") & "',"
        Sql = Sql & !SEQ_VINC & ",'" & SubNull(!IN_TRANSMITIDO) & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Lg "Setting null values"
Sql = "UPDATE SO_VINCULO SET IN_TRANSMITIDO=NULL WHERE IN_TRANSMITIDO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_VINCULO SET DT_FIM_PERIODO=NULL WHERE DT_FIM_PERIODO='01/01/1900'"
cn.Execute Sql, rdExecDirect

CEP:
Lg "Deleting table [so_cep]"

Sql = "IF EXISTS (SELECT 1 From INFORMATION_SCHEMA.Tables WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='SO_CEP') "
Sql = Sql & "DROP TABLE SO_CEP"
cn.Execute Sql, rdExecDirect

Lg "Creating table [so_cep]"
Sql = "CREATE TABLE so_cep(nu_cep varchar(8) PRIMARY KEY not null,id_tipo_logr varchar(10),te_descricao_cep varchar(70),nm_bairro varchar(60),"
Sql = Sql & "sg_uf varchar(2),id_muni_prev varchar(5),id_paf varchar(8))"
cn.Execute Sql, rdExecDirect

Lg "Granting permission to table [so_cep]"
Sql = "GRANT SELECT,UPDATE,INSERT,DELETE ON SO_CEP TO PUBLIC"
cn.Execute Sql, rdExecDirect

Lg "Converting Table [so_cep] to Sql Server"

Sql = "SELECT * FROM TAB_CEP WHERE ID_MUNI_PREV='21244'"
Set RdoIB = cnIB.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoIB
    nPos = 1: nTotal = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "INSERT SO_CEP(NU_CEP,ID_TIPO_LOGR,TE_DESCRICAO_CEP,NM_BAIRRO,SG_UF,ID_MUNI_PREV,ID_PAF"
        Sql = Sql & ") VALUES('"
        Sql = Sql & !NU_CEP & "','" & SubNull(!ID_TIPO_LOGR) & "','" & Mask(SubNull(!TE_DESCRICAO_CEP)) & "','" & SubNull(!NM_BAIRRO) & "','"
        Sql = Sql & SubNull(!SG_UF) & "','" & SubNull(!ID_MUNI_PREV) & "','" & SubNull(!ID_PAF) & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

RESP:
Lg "Deleting table [so_responsavel]"

Sql = "IF EXISTS (SELECT 1 From INFORMATION_SCHEMA.Tables WHERE TABLE_TYPE='BASE TABLE' AND TABLE_NAME='SO_RESPONSAVEL') "
Sql = Sql & "DROP TABLE SO_RESPONSAVEL"
cn.Execute Sql, rdExecDirect

Lg "Creating table [so_responsavel]"
Sql = "CREATE TABLE so_responsavel(cs_tip_resp char(1),id_responsavel varchar(14) PRIMARY KEY not null,nm_responsavel varchar(55),id_tipo_lgdr varchar(10),nm_endereco varchar(55),"
Sql = Sql & "nu_endereco varchar(10),compl_endereco varchar(10),nm_bairro varchar(20),nu_cep varchar(8),id_municipio varchar(6),nm_cidade varchar(30),"
Sql = Sql & "cs_uf char(2),id_tipo_lgdr_corr varchar(10),nm_endereco_corr varchar(55),nu_endereco_corr varchar(10),compl_endereco_corr varchar(10),"
Sql = Sql & "nm_bairro_corr varchar(20),nu_cep_corr varchar(8),id_municipio_corr varchar(6),nm_cidade_corr varchar(30),nu_telefone varchar(14),"
Sql = Sql & "nu_fax varchar(14),te_email varchar(60),dt_atualizacao smalldatetime,cs_uf_corr char(2),nu_ddd_tel varchar(4),nu_ddd_fax varchar(4))"
cn.Execute Sql, rdExecDirect

Lg "Granting permission to table [so_responsavel]"
Sql = "GRANT SELECT,UPDATE,INSERT,DELETE ON SO_RESPONSAVEL TO PUBLIC"
cn.Execute Sql, rdExecDirect

Lg "Converting Table [so_responsavel] to Sql Server"

Sql = "SELECT * FROM RESPONSAVEL"
Set RdoIB = cnIB.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoIB
    nPos = 1: nTotal = .RowCount
    Do Until .EOF
        If nPos Mod 10 = 0 Then
            CallPb nPos, nTotal
        End If
        Sql = "INSERT SO_RESPONSAVEL(CS_TIP_RESP,ID_RESPONSAVEL,NM_RESPONSAVEL,ID_TIPO_LGDR,NM_ENDERECO,NU_ENDERECO,COMPL_ENDERECO,NM_BAIRRO,NU_CEP,"
        Sql = Sql & "ID_MUNICIPIO,NM_CIDADE,CS_UF,ID_TIPO_LGDR_CORR,NM_ENDERECO_CORR,NU_ENDERECO_CORR,COMPL_ENDERECO_CORR,NM_BAIRRO_CORR,NU_CEP_CORR,"
        Sql = Sql & "ID_MUNICIPIO_CORR,NM_CIDADE_CORR,NU_TELEFONE,NU_FAX,TE_EMAIL,DT_ATUALIZACAO,cs_uf_corr,NU_DDD_TEL,NU_DDD_FAX"
        Sql = Sql & ") VALUES('"
        Sql = Sql & !CS_TIP_RESP & "','" & !ID_RESPONSAVEL & "','" & Mask(SubNull(!NM_RESPONSAVEL)) & "','" & SubNull(!ID_TIPO_LGDR) & "','" & Mask(SubNull(!NM_ENDERECO)) & "','"
        Sql = Sql & SubNull(!NU_ENDERECO) & "','" & Mask(SubNull(!COMPL_ENDERECO)) & "','" & Mask(SubNull(!NM_BAIRRO)) & "','" & Mask(SubNull(!NU_CEP)) & "','" & Mask(SubNull(!ID_MUNICIPIO)) & "','"
        Sql = Sql & Mask(SubNull(!NM_CIDADE)) & "','" & Mask(SubNull(!CS_UF)) & "','" & Mask(SubNull(!ID_TIPO_LGDR_CORR)) & "','" & Mask(SubNull(!NM_ENDERECO_CORR)) & "','"
        Sql = Sql & Mask(SubNull(!NU_ENDERECO_CORR)) & "','" & Mask(SubNull(!COMPL_ENDERECO_CORR)) & "','" & Mask(SubNull(!NM_BAIRRO_CORR)) & "','" & Mask(SubNull(!NU_CEP_CORR)) & "','"
        Sql = Sql & Mask(SubNull(!ID_MUNICIPIO_CORR)) & "','" & Mask(SubNull(!NM_CIDADE_CORR)) & "','" & Mask(SubNull(!NU_TELEFONE)) & "','" & Mask(SubNull(!NU_FAX)) & "','"
        Sql = Sql & Mask(SubNull(!TE_EMAIL)) & "','" & Format(!DT_ATUALIZACAO, "mm/dd/yyyy") & "','" & Mask(SubNull(!cs_uf_corr)) & "','" & Mask(SubNull(!NU_DDD_TEL)) & "','"
        Sql = Sql & Mask(SubNull(!NU_DDD_FAX))
        Sql = Sql & "')"
        cn.Execute Sql, rdExecDirect
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Lg "Setting null values"
Sql = "UPDATE SO_RESPONSAVEL SET ID_TIPO_LGDR=NULL WHERE ID_TIPO_LGDR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET COMPL_ENDERECO=NULL WHERE COMPL_ENDERECO=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET ID_TIPO_LGDR_CORR=NULL WHERE ID_TIPO_LGDR_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NM_ENDERECO_CORR=NULL WHERE NM_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_ENDERECO_CORR=NULL WHERE NU_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET COMPL_ENDERECO_CORR=NULL WHERE COMPL_ENDERECO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NM_BAIRRO_CORR=NULL WHERE NM_BAIRRO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_CEP_CORR=NULL WHERE NU_CEP_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET ID_MUNICIPIO_CORR=NULL WHERE ID_MUNICIPIO_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NM_CIDADE_CORR=NULL WHERE NM_CIDADE_CORR=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_TELEFONE=NULL WHERE NU_TELEFONE=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_FAX=NULL WHERE NU_FAX=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET TE_EMAIL=NULL WHERE TE_EMAIL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET cs_uf_corr=NULL WHERE cs_uf_corr=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_DDD_TEL=NULL WHERE NU_DDD_TEL=''"
cn.Execute Sql, rdExecDirect
Sql = "UPDATE SO_RESPONSAVEL SET NU_DDD_FAX=NULL WHERE NU_DDD_FAX=''"
cn.Execute Sql, rdExecDirect

PBar.Value = 0: PBar.Color = &HFFFFFF
Lg ""
Lg "Conversion of SisObraPref to SqlServer is completed."

Liberado

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Sub Lg(sTexto As String)
lstLog.AddItem sTexto
lstLog.ListIndex = lstLog.ListCount - 1
DoEvents
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
cnIB.Close
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
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


