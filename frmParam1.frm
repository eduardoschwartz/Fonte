VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmParam1 
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4035
   ClientLeft      =   3015
   ClientTop       =   1935
   ClientWidth     =   5595
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4035
   ScaleWidth      =   5595
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   3225
      TabIndex        =   19
      ToolTipText     =   "Gravar os Dados"
      Top             =   3615
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   14
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParam1.frx":0000
      PICN            =   "frmParam1.frx":001C
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
      Left            =   4350
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   3600
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParam1.frx":03C1
      PICN            =   "frmParam1.frx":03DD
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
      Left            =   30
      TabIndex        =   15
      ToolTipText     =   "Novo Registro"
      Top             =   3600
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParam1.frx":0537
      PICN            =   "frmParam1.frx":0553
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
      Left            =   1080
      TabIndex        =   16
      ToolTipText     =   "Editar Registro"
      Top             =   3600
      Width           =   1035
      _ExtentX        =   1826
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmParam1.frx":06AD
      PICN            =   "frmParam1.frx":06C9
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
      Left            =   2130
      TabIndex        =   17
      ToolTipText     =   "Excluir Registro"
      Top             =   3600
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Excluir"
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
      MICON           =   "frmParam1.frx":0823
      PICN            =   "frmParam1.frx":083F
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
      Left            =   4320
      TabIndex        =   18
      ToolTipText     =   "Sair da Tela"
      Top             =   3630
      Width           =   1035
      _ExtentX        =   1826
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
      MICON           =   "frmParam1.frx":08E1
      PICN            =   "frmParam1.frx":08FD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2520
      Width           =   5580
      Begin VB.TextBox txtAbrev 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4470
         TabIndex        =   3
         TabStop         =   0   'False
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   255
         Width           =   1005
      End
      Begin VB.TextBox txtDesc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   5
         Top             =   570
         Width           =   4005
      End
      Begin VB.Label lblAbrev 
         BackStyle       =   0  'Transparent
         Caption         =   "Abreviatura.........:"
         Height          =   195
         Left            =   3150
         TabIndex        =   13
         Top             =   285
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   135
         TabIndex        =   6
         Top             =   300
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Descrição...........:"
         Height          =   195
         Index           =   11
         Left            =   135
         TabIndex        =   4
         Top             =   630
         Width           =   1275
      End
   End
   Begin VB.Frame Fra 
      BackColor       =   &H00EEEEEE&
      Height          =   990
      Index           =   1
      Left            =   0
      TabIndex        =   7
      Top             =   2520
      Width           =   5580
      Begin VB.TextBox txtValorUFIR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         MaxLength       =   50
         TabIndex        =   9
         Top             =   570
         Width           =   1305
      End
      Begin VB.TextBox txtAnoUFIR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1485
         TabIndex        =   8
         TabStop         =   0   'False
         Top             =   255
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "VALOR UFIR.....:"
         Height          =   195
         Index           =   2
         Left            =   135
         TabIndex        =   11
         Top             =   630
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "ANO UFIR..........:"
         Height          =   195
         Index           =   1
         Left            =   135
         TabIndex        =   10
         Top             =   300
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdUFIR 
      Height          =   2505
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   4419
      _Version        =   393216
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "^Ano UFIR             |>Valor UFIR                          "
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   2505
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5610
      _ExtentX        =   9895
      _ExtentY        =   4419
      _Version        =   393216
      FixedCols       =   0
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   "Código       |<Descricão                                                                            "
   End
End
Attribute VB_Name = "frmParam1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String, bResize As Boolean
Dim Evento As String
Dim sRet As String
'Dim evCnBF As Integer, evCnCC As Integer, evCnCP As Integer
'Dim evCnTC As Integer, evCnUF As Integer
'Dim evCnUC As Integer, evCnUT As Integer, evCnMO As Integer
'
'Dim evNewCC As Integer, evNewCP As Integer
'Dim evNewBF As Integer, evNewTC As Integer, evNewUF As Integer
'Dim evNewUC As Integer, evNewUT As Integer, evNewMO As Integer
'
'Dim evEditCC As Integer, evEditCP As Integer
'Dim evEditBF As Integer, evEditTC As Integer, evEditUF As Integer
'Dim evEditUC As Integer, evEditUT As Integer, evEditMO As Integer
'
'Dim evDelCC As Integer, evDelCP As Integer
'Dim evDelBF As Integer, evDelTC As Integer, evDelUF As Integer
'Dim evDelUC As Integer, evDelUT As Integer, evDelMO As Integer
'
'Dim bEvCnCC As Boolean, bEvCnCP As Boolean
'Dim bEvCnBF As Boolean, bEvCnTC As Boolean, bEvCnUF As Boolean
'Dim bEvCnUC As Boolean, bEvCnUT As Boolean, bEvCnMO As Boolean
'
'Dim bEvNewCC As Boolean, bEvNewCP As Boolean
'Dim bEvNewBF As Boolean, bEvNewTC As Boolean, bEvNewUF As Boolean
'Dim bEvNewUC As Boolean, bEvNewUT As Boolean, bEvNewMO As Boolean
'
'Dim bEvEditCC As Boolean, bEvEditCP As Boolean
'Dim bEvEditBF As Boolean, bEvEditTC As Boolean, bEvEditUF As Boolean
'Dim bEvEditUC As Boolean, bEvEditUT As Boolean, bEvEditMO As Boolean
'
'Dim bEvDelCC As Boolean, bEvDelCP As Boolean
'Dim bEvDelBF As Boolean, bEvDelTC As Boolean, bEvDelUF As Boolean
'Dim bEvDelUC As Boolean, bEvDelUT As Boolean, bEvDelMO As Boolean

Private Sub cmdAlterar_Click()
    
    If txtCod.text = "" And sParamForm <> "UFIR" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    If sParamForm <> "UFIR" Then
         sOldDesc = txtDesc.text
    Else
        sOldDesc = txtValorUFIR.text
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"

End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub cmdExcluir_Click()
    If txtCod.text = "" And sParamForm <> "UFIR" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir este Registro ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
        Select Case sParamForm
             Case "BENF"
                 Sql = "SELECT CODREDUZIDO,DT_CODBENF FROM CADIMOB WHERE DT_CODBENF=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM FATORBENFEITORIA WHERE CODBENFEITORIA=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                        Sql = "DELETE FROM BENFEITORIA WHERE CODBENFEITORIA=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "CATC"
                 Sql = "SELECT CODREDUZIDO,CATCONSTR FROM AREAS WHERE CATCONSTR=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM FATORCATEG WHERE CODCATEG=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                        Sql = "DELETE FROM CATEGCONSTR WHERE CODCATEGCONSTR=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "CATP"
                 Sql = "SELECT CODREDUZIDO,DT_CODCATEGPROP FROM CADIMOB WHERE DT_CODCATEGPROP=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM CATEGPROP WHERE CODCATEGPROP=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "TIPC"
                 Sql = "SELECT CODREDUZIDO,TIPOCONSTR FROM AREAS WHERE TIPOCONSTR=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM FATORCATEG WHERE CODTIPO=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                        Sql = "DELETE FROM TIPOCONSTR WHERE CODTIPOCONSTR=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "USOC"
                 Sql = "SELECT CODREDUZIDO,USOCONSTR FROM AREAS WHERE USOCONSTR=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM FATORCATEG WHERE CODUSO=" & Val(txtCod.text)
                        cn.Execute Sql, rdExecDirect
                        Sql = "DELETE FROM USOCONSTR WHERE CODUSOCONSTR=" & Val(txtCod.text)
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "USOT"
                 Sql = "SELECT CODREDUZIDO,DT_CODUSOTERRENO FROM CADIMOB WHERE DT_CODUSOTERRENO=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM USOTERRENO WHERE CODUSOTERRENO=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum imóvel.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "MOED"
                 Sql = "SELECT CODUSO,CODMOEDA FROM FATORCATEG WHERE CODMOEDA=" & Val(txtCod.text)
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                    If .RowCount = 0 Then
                        Sql = "DELETE FROM MOEDA WHERE CODMOEDA=" & txtCod.text
                        cn.Execute Sql, rdExecDirect
                       .Close
                    Else
                        MsgBox "Não é possível excluir o registro pois esta sendo utilizado por algum Fator Categoria.", vbExclamation, "Atenção"
                       .Close
                    End If
                 End With
             Case "UFIR"
                 Sql = "DELETE FROM UFIR WHERE ANOUFIR=" & Val(txtAnoUFIR.text)
                 cn.Execute Sql, rdExecDirect
        End Select
       Select Case sParamForm
            Case "UFIR"
                    Log Form, Me.Caption, Exclusão, "Excluído UFIR " & txtAnoUFIR.text & "-" & txtValorUFIR.text
            Case Else
                    Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.text, "000") & "-" & txtDesc.text
        End Select
       Limpa
       CarregaLista
       Le
    End If

End Sub

Private Sub cmdGravar_Click()

Select Case sParamForm
        Case "UFIR"
            If Val(txtAnoUFIR.text) < 1980 Or Val(txtAnoUFIR.text) > 2010 Then
                 MsgBox "Ano de UFIR inválido.", vbExclamation, "Atenção"
                 txtAnoUFIR.SetFocus
                 Exit Sub
            End If
            If CDbl(txtValorUFIR.text) <= 0 Then
                 MsgBox "Favor Digitar o Valor da UFIR.", vbExclamation, "Atenção"
                 txtValorUFIR.SetFocus
                 Exit Sub
            End If
            If Evento = "Novo" Then
                Sql = "SELECT VALORUFIR FROM UFIR WHERE ANOUFIR=" & Val(txtAnoUFIR.text)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                        If .RowCount > 0 Then
                             MsgBox "Já existe uma UFIR Cadastrada para este Ano.", vbExclamation, "Atenção"
                             txtAnoUFIR.SetFocus
                             Exit Sub
                        End If
                       .Close
                End With
            End If
        Case "MOED"
            If txtDesc.text = "" Then
               MsgBox "Favor digitar a Descrição.", vbExclamation, "Atenção"
               txtDesc.SetFocus
               Exit Sub
            End If
            If txtAbrev.text = "" Then
               MsgBox "Favor digitar a Abreviatura.", vbExclamation, "Atenção"
               txtAbrev.SetFocus
               Exit Sub
            End If
        Case Else
            If txtDesc.text = "" Then
               MsgBox "Favor digitar a Descrição.", vbExclamation, "Atenção"
               txtDesc.SetFocus
               Exit Sub
            End If
End Select

Grava
Eventos "INICIAR"

End Sub

Private Sub cmdNovo_Click()
On Error Resume Next
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
    txtAnoUFIR.Enabled = True
    txtAnoUFIR.BackColor = Branco
    txtAnoUFIR.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
Liberado
bResize = True
End Sub

Private Sub Form_Load()

Ocupado

Select Case sParamForm
       Case "BENF"
               Me.Caption = "Tabela de Benfeitorias"
       Case "CATC"
               Me.Caption = "Categoria da Construção"
       Case "CATP"
               Me.Caption = "Categoria da Propriedade"
       Case "TIPC"
               Me.Caption = "Tipo de Construção"
       Case "USOC"
               Me.Caption = "Uso da Construção"
       Case "USOT"
               Me.Caption = "Uso do Terreno"
       Case "MOED"
               Me.Caption = "Tabela de Moedas"
       Case "UFIR"
               Me.Caption = "Tabela de UFIR"
End Select

Centraliza Me

sRet = RetEventUserForm(Me.Name)

Select Case sParamForm
       Case "UFIR"
            Fra(0).Visible = False
            Fra(1).Visible = True
            grdUFIR.Visible = True
            grdMain.Visible = False
            grdUFIR.Rows = 1
            lblAbrev.Visible = False
            txtAbrev.Visible = False
       Case "MOED"
            Fra(0).Visible = True
            Fra(1).Visible = False
            grdUFIR.Visible = False
            grdMain.Visible = True
            grdMain.Rows = 1
            lblAbrev.Visible = True
            txtAbrev.Visible = True
       Case Else
            Fra(0).Visible = True
            Fra(1).Visible = False
            grdUFIR.Visible = False
            grdMain.Visible = True
            grdMain.Rows = 1
            lblAbrev.Visible = False
            txtAbrev.Visible = False
End Select

CarregaLista
Le

Eventos "INICIAR"
End Sub

Private Sub grdMain_Click()
If grdMain.Row > 0 Then
     txtCod.text = grdMain.TextMatrix(grdMain.Row, 0)
     txtDesc.text = grdMain.TextMatrix(grdMain.Row, 1)
     If sParamForm = "MOED" Then
          Sql = "SELECT  ABREVMOEDA FROM MOEDA WHERE CODMOEDA=" & Val(txtCod.text)
          Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
          With RdoAux
                txtAbrev.text = !ABREVMOEDA
               .Close
          End With
     End If
 End If
End Sub

Private Sub grdMain_RowColChange()
grdMain_Click
End Sub

Private Sub grdUFIR_RowColChange()
If grdUFIR.Row > 0 Then
     txtAnoUFIR.text = grdUFIR.TextMatrix(grdUFIR.Row, 0)
     txtValorUFIR.text = grdUFIR.TextMatrix(grdUFIR.Row, 1)
 End If

End Sub

Private Sub Grava()
Dim x As Integer
Dim MaxCod As Integer
Dim qd As New rdoQuery

On Error Resume Next
RdoAux.Close
On Error GoTo Fim
Set qd.ActiveConnection = cn

Select Case sParamForm
    Case "BENF"
        Sql = "SELECT MAX(CODBENFEITORIA) AS MAXIMO FROM BENFEITORIA WHERE CODBENFEITORIA<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT BENFEITORIA (CODBENFEITORIA,DESCBENFEITORIA) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE BENFEITORIA SET DESCBENFEITORIA='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODBENFEITORIA=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "CATC"
        Sql = "SELECT MAX(CODCATEGCONSTR) AS MAXIMO FROM CATEGCONSTR WHERE CODCATEGCONSTR<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT CATEGCONSTR (CODCATEGCONSTR,DESCCATEGCONSTR) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE CATEGCONSTR SET DESCCATEGCONSTR='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODCATEGCONSTR=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "CATP"
        Sql = "SELECT MAX(CODCATEGPROP) AS MAXIMO FROM CATEGPROP WHERE CODCATEGPROP<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT CATEGPROP (CODCATEGPROP,DescCategProp) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE CATEGPROP SET DescCategProp='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODCATEGPROP=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "TIPC"
        Sql = "SELECT MAX(CODTIPOCONSTR) AS MAXIMO FROM TIPOCONSTR WHERE CODTIPOCONSTR<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT TIPOCONSTR (CODTIPOCONSTR,DESCTIPOCONSTR) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE TIPOCONSTR SET DESCTIPOCONSTR='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODTIPOCONSTR=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "USOC"
        Sql = "SELECT MAX(CODUSOCONSTR) AS MAXIMO FROM USOCONSTR WHERE CODUSOCONSTR<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT USOCONSTR (CODUSOCONSTR,DESCTIPOCONSTR) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE USOCONSTR SET DESCTIPOCONSTR='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODUSOCONSTR=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "USOT"
        Sql = "SELECT MAX(CODUSOTERRENO) AS MAXIMO FROM USOTERRENO WHERE CODUSOTERRENO<999"
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux!MAXIMO) Then
            MaxCod = 1
        Else
            MaxCod = RdoAux!MAXIMO + 1
        End If
        RdoAux.Close
        If Evento = "Novo" Then
            Sql = "INSERT USOTERRENO (CODUSOTERRENO,DescUsoTerreno) VALUES("
            Sql = Sql & MaxCod & ",'" & Mask(txtDesc.text) & "')"
        Else
            Sql = "UPDATE USOTERRENO SET DescUsoTerreno='" & Mask(txtDesc.text) & "' WHERE "
            Sql = Sql & "CODUSOTERRENO=" & Val(txtCod.text)
        End If
        cn.Execute Sql, rdExecDirect
     Case "UFIR"
        If Evento = "Novo" Then
            Sql = "INSERT UFIR (ANOUFIR,VALORUFIR) VALUES("
            Sql = Sql & txtAnoUFIR.text & "," & Virg2Ponto(txtValorUFIR) & ")"
        Else
            Sql = "UPDATE UFIR SET VALORUFIR=" & Virg2Ponto(txtValorUFIR) & " WHERE "
            Sql = Sql & "ANOUFIR=" & txtAnoUFIR.text
        End If
        cn.Execute Sql, rdExecDirect
End Select

If Evento = "Novo" Then
     grdMain.AddItem MaxCod & Chr(9) & txtDesc.text
Else
    grdMain.TextMatrix(grdMain.Row, 1) = txtDesc.text
End If
Evento = ""
Exit Sub

FimUfir:
If Evento = "Novo" Then
     grdUFIR.AddItem txtAnoUFIR.text & Chr(9) & Format(txtValorUFIR.text, "#0.0000")
Else
    grdUFIR.TextMatrix(grdUFIR.Row, 1) = Format(txtValorUFIR.text, "#0.0000")
End If
Evento = ""
      
Exit Sub

Fim:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub Eventos(Tipo As String)

Dim Ct As Control

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   For Each Ct In frmParam1
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   grdMain.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmParam1
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = vbWhite
          Ct.Enabled = True
       End If
   Next
   txtCod.BackColor = Kde
   txtCod.Locked = True
   grdMain.Enabled = False
   If sParamForm <> "UFIR" Then
        txtDesc.SetFocus
   Else
        txtAnoUFIR.Enabled = False
        txtAnoUFIR.BackColor = Kde
        txtValorUFIR.SetFocus
   End If
End If

FormHagana sParamForm

End Sub

Private Sub Le()

Select Case sParamForm
        Case "UFIR"
            If grdUFIR.Row = 0 Then Exit Sub
            txtAnoUFIR.text = grdUFIR.TextMatrix(grdUFIR.Row, 0)
            txtValorUFIR.text = grdUFIR.TextMatrix(grdUFIR.Row, 1)
        Case "MOED"
            If grdMain.Row = 0 Then Exit Sub
            txtCod.text = grdMain.TextMatrix(grdMain.Row, 0)
            txtDesc.text = grdMain.TextMatrix(grdMain.Row, 1)
            Sql = "SELECT  ABREVMOEDA FROM MOEDA WHERE CODMOEDA=" & Val(txtCod.text)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                  txtAbrev.text = !ABREVMOEDA
                 .Close
            End With
        Case Else
            If grdMain.Row = 0 Then Exit Sub
            txtCod.text = grdMain.TextMatrix(grdMain.Row, 0)
            txtDesc.text = grdMain.TextMatrix(grdMain.Row, 1)
End Select

End Sub

Private Sub Limpa()
Select Case sParamForm
        Case "UFIR"
            txtAnoUFIR.text = ""
            txtValorUFIR.text = ""
        Case Else
            txtCod.text = ""
            txtDesc.text = ""
            If txtAbrev.Visible = True Then txtAbrev.text = ""
End Select
End Sub

Private Sub CarregaLista()

Select Case sParamForm
     Case "BENF"
           Sql = "Select CODBENFEITORIA,DESCBENFEITORIA From BENFEITORIA WHERE CODBENFEITORIA<>999"
           Sql = Sql & "ORDER BY DESCBENFEITORIA"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODBENFEITORIA & Chr(9) & !DescBenfeitoria
                  .MoveNext
                Loop
               .Close
            End With
     Case "CATC"
           Sql = "Select CODCATEGCONSTR,DESCCATEGCONSTR From CATEGCONSTR WHERE CODCATEGCONSTR<>999"
           Sql = Sql & "ORDER BY DESCCATEGCONSTR"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODCATEGCONSTR & Chr(9) & !DESCCATEGCONSTR
                  .MoveNext
                Loop
               .Close
            End With
     Case "CATP"
           Sql = "Select CODCATEGPROP,DESCCATEGPROP From CATEGPROP WHERE CODCATEGPROP<>999"
           Sql = Sql & "ORDER BY DESCCATEGPROP"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODCATEGPROP & Chr(9) & !DescCategProp
                  .MoveNext
                Loop
               .Close
            End With
     Case "TIPC"
           Sql = "Select CODTIPOCONSTR,DESCTIPOCONSTR From TIPOCONSTR WHERE CODTIPOCONSTR<>999"
           Sql = Sql & "ORDER BY DESCTIPOCONSTR"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODTIPOCONSTR & Chr(9) & !DESCTIPOCONSTR
                  .MoveNext
                Loop
               .Close
            End With
     Case "USOC"
           Sql = "Select CODUSOCONSTR,DESCUSOCONSTR From USOCONSTR WHERE CODUSOCONSTR<>999"
           Sql = Sql & "ORDER BY DESCUSOCONSTR"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODUSOCONSTR & Chr(9) & !DESCUSOCONSTR
                  .MoveNext
                Loop
               .Close
            End With
     Case "USOT"
           Sql = "Select CODUSOTERRENO,DESCUSOTERRENO From USOTERRENO WHERE CODUSOTERRENO<>999"
           Sql = Sql & "ORDER BY DESCUSOTERRENO"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODUSOTERRENO & Chr(9) & !DescUsoTerreno
                  .MoveNext
                Loop
               .Close
            End With
     Case "MOED"
           Sql = "Select CODMOEDA,DESCMOEDA,ABREVMOEDA FROM MOEDA "
           Sql = Sql & "ORDER BY DESCMOEDA"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdMain.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdMain.AddItem !CODMOEDA & Chr(9) & !DESCMOEDA
                  .MoveNext
                Loop
               .Close
            End With
     Case "UFIR"
           Sql = "Select ANOUFIR,VALORUFIR FROM UFIR "
           Sql = Sql & "ORDER BY ANOUFIR"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            grdUFIR.Rows = 1
            With RdoAux
               .MoveFirst
                Do Until .EOF
                   grdUFIR.AddItem !ANOUFIR & Chr(9) & Format(!VALORUFIR, "#0.0000")
                  .MoveNext
                Loop
               .Close
            End With
End Select
End Sub

Private Sub FormHagana(sTela As String)

If NomeDeLogin = "SCHWARTZ" Then Exit Sub


evNewBF = 9: evNewCC = 13: evNewCP = 17: evNewTC = 29: evNewUC = 37
evNewUT = 41: evNewMO = 45: evNewUF = 77

evEditBF = 10: evEditCC = 13: evEditCP = 18: evEditTC = 30: evEditUC = 38
evEditUT = 42: evEditMO = 46: evEditUF = 78

evDelBF = 11: evDelCC = 15: evDelCP = 19: evDelTC = 31: evDelUC = 39
evDelUT = 43: evDelMO = 47: evDelUF = 79

bEvNewBF = False: bEvNewCC = False: bEvNewCP = False: bEvNewTC = False: bEvNewUC = False
bEvNewUT = False: bEvNewMO = False: bEvNewUF = False

bEvEditBF = False: bEvEditCC = False: bEvEditCP = False: bEvEditTC = False: bEvEditUC = False
bEvEditUT = False: bEvEditMO = False: bEvEditUF = False

bEvDelBF = False: bEvDelCC = False: bEvDelCP = False: bEvDelTC = False: bEvDelUC = False
bEvDelUT = False: bEvDelMO = False: bEvDelUF = False

If InStr(1, sRet, Format(evNewBF, "000"), vbBinaryCompare) > 0 Then bEvNewBF = True
If InStr(1, sRet, Format(evNewCC, "000"), vbBinaryCompare) > 0 Then bEvNewCC = True
If InStr(1, sRet, Format(evNewCP, "000"), vbBinaryCompare) > 0 Then bEvNewCP = True
If InStr(1, sRet, Format(evNewTC, "000"), vbBinaryCompare) > 0 Then bEvNewTC = True
If InStr(1, sRet, Format(evNewUC, "000"), vbBinaryCompare) > 0 Then bEvNewUC = True
If InStr(1, sRet, Format(evNewUT, "000"), vbBinaryCompare) > 0 Then bEvNewUT = True
If InStr(1, sRet, Format(evNewMO, "000"), vbBinaryCompare) > 0 Then bEvNewMO = True
If InStr(1, sRet, Format(evNewUF, "000"), vbBinaryCompare) > 0 Then bEvNewUF = True

If InStr(1, sRet, Format(evEditBF, "000"), vbBinaryCompare) > 0 Then bEvEditBF = True
If InStr(1, sRet, Format(evEditCC, "000"), vbBinaryCompare) > 0 Then bEvEditCC = True
If InStr(1, sRet, Format(evEditCP, "000"), vbBinaryCompare) > 0 Then bEvEditCP = True
If InStr(1, sRet, Format(evEditTC, "000"), vbBinaryCompare) > 0 Then bEvEditTC = True
If InStr(1, sRet, Format(evEditUC, "000"), vbBinaryCompare) > 0 Then bEvEditUC = True
If InStr(1, sRet, Format(evEditUT, "000"), vbBinaryCompare) > 0 Then bEvEditUT = True
If InStr(1, sRet, Format(evEditMO, "000"), vbBinaryCompare) > 0 Then bEvEditMO = True
If InStr(1, sRet, Format(evEditUF, "000"), vbBinaryCompare) > 0 Then bEvEditUF = True

If InStr(1, sRet, Format(evDelBF, "000"), vbBinaryCompare) > 0 Then bEvDelBF = True
If InStr(1, sRet, Format(evDelCC, "000"), vbBinaryCompare) > 0 Then bEvDelCC = True
If InStr(1, sRet, Format(evDelCP, "000"), vbBinaryCompare) > 0 Then bEvDelCP = True
If InStr(1, sRet, Format(evDelTC, "000"), vbBinaryCompare) > 0 Then bEvDelTC = True
If InStr(1, sRet, Format(evDelUC, "000"), vbBinaryCompare) > 0 Then bEvDelUC = True
If InStr(1, sRet, Format(evDelUT, "000"), vbBinaryCompare) > 0 Then bEvDelUT = True
If InStr(1, sRet, Format(evDelMO, "000"), vbBinaryCompare) > 0 Then bEvDelMO = True
If InStr(1, sRet, Format(evDelUF, "000"), vbBinaryCompare) > 0 Then bEvDelUF = True

Select Case sTela
          Case "BENF"
                cmdNovo.Enabled = bEvNewBF
                cmdAlterar.Enabled = bEvEditBF
                cmdExcluir.Enabled = bEvDelBF
          Case "CATC"
                cmdNovo.Enabled = bEvNewCC
                cmdAlterar.Enabled = bEvEditCC
                cmdExcluir.Enabled = bEvDelCC
          Case "CATP"
                cmdNovo.Enabled = bEvNewCP
                cmdAlterar.Enabled = bEvEditCP
                cmdExcluir.Enabled = bEvDelCP
          Case "TIPC"
                cmdNovo.Enabled = bEvNewTC
                cmdAlterar.Enabled = bEvEditTC
                cmdExcluir.Enabled = bEvDelTC
          Case "USOC"
                cmdNovo.Enabled = bEvNewUC
                cmdAlterar.Enabled = bEvEditUC
                cmdExcluir.Enabled = bEvDelUC
          Case "USOT"
                cmdNovo.Enabled = bEvNewUT
                cmdAlterar.Enabled = bEvEditUT
                cmdExcluir.Enabled = bEvDelUT
          Case "MOED"
                cmdNovo.Enabled = bEvNewMO
                cmdAlterar.Enabled = bEvEditMO
                cmdExcluir.Enabled = bEvDelMO
          Case "UFIR"
                cmdNovo.Enabled = bEvNewUF
                cmdAlterar.Enabled = bEvEditUF
                cmdExcluir.Enabled = bEvDelUF
End Select

End Sub

Private Sub txtAnoUFIR_KeyPress(KeyAscii As Integer)
Tweak txtAnoUFIR, KeyAscii, IntegerPositive
End Sub

Private Sub txtValorUFIR_KeyPress(KeyAscii As Integer)
Tweak txtValorUFIR, KeyAscii, DecimalPositive, 4
End Sub
