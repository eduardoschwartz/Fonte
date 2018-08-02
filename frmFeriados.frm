VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmFeriados 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Feriados "
   ClientHeight    =   5130
   ClientLeft      =   5145
   ClientTop       =   2775
   ClientWidth     =   6450
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5130
   ScaleWidth      =   6450
   Begin VB.Frame frButtons 
      BackColor       =   &H00EEEEEE&
      Height          =   510
      Left            =   45
      TabIndex        =   14
      Top             =   4590
      Width           =   6405
      Begin prjChameleon.chameleonButton cmdSair 
         Height          =   315
         Left            =   5325
         TabIndex        =   15
         ToolTipText     =   "Sair da Tela"
         Top             =   135
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmFeriados.frx":0000
         PICN            =   "frmFeriados.frx":001C
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
         Height          =   315
         Left            =   5295
         TabIndex        =   16
         ToolTipText     =   "Gravar os Dados"
         Top             =   135
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   556
         BTYPE           =   3
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   12632256
         BCOLO           =   12632256
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmFeriados.frx":008A
         PICN            =   "frmFeriados.frx":00A6
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
         Left            =   4245
         TabIndex        =   17
         ToolTipText     =   "Cancelar Edição"
         Top             =   135
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmFeriados.frx":044B
         PICN            =   "frmFeriados.frx":0467
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
         Left            =   2145
         TabIndex        =   18
         ToolTipText     =   "Excluir Registro"
         Top             =   135
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmFeriados.frx":05C1
         PICN            =   "frmFeriados.frx":05DD
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
         Left            =   1095
         TabIndex        =   19
         ToolTipText     =   "Editar Registro"
         Top             =   135
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmFeriados.frx":067F
         PICN            =   "frmFeriados.frx":069B
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
         Left            =   45
         TabIndex        =   20
         ToolTipText     =   "Novo Registro"
         Top             =   135
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
         MCOL            =   13026246
         MPTR            =   1
         MICON           =   "frmFeriados.frx":07F5
         PICN            =   "frmFeriados.frx":0811
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
   Begin VB.Frame frAll 
      Height          =   3795
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   6435
      Begin VB.ListBox lstAll 
         Appearance      =   0  'Flat
         Height          =   3540
         Left            =   45
         TabIndex        =   13
         Top             =   165
         Width           =   6330
      End
   End
   Begin VB.ComboBox cmbFeriado 
      Height          =   315
      Left            =   2670
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   360
      Width           =   3615
   End
   Begin VB.TextBox txtObs 
      Height          =   855
      Left            =   2670
      MaxLength       =   40
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1530
      Width           =   3585
   End
   Begin MSFlexGridLib.MSFlexGrid grdFer 
      Height          =   1380
      Left            =   0
      TabIndex        =   1
      Top             =   2790
      Width           =   6435
      _ExtentX        =   11351
      _ExtentY        =   2434
      _Version        =   393216
      Rows            =   1
      Cols            =   4
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Cod       |^Data            |<Descrição                             |<Observação                            "
   End
   Begin MSComCtl2.MonthView Mv 
      Height          =   2370
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   15658734
      Appearance      =   1
      StartOfWeek     =   53346305
      TitleBackColor  =   192
      TitleForeColor  =   12648447
      CurrentDate     =   37439
   End
   Begin prjChameleon.chameleonButton cmdAddFeriado 
      Height          =   315
      Left            =   2670
      TabIndex        =   7
      ToolTipText     =   "Adiciona Feriado"
      Top             =   750
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Adicionar"
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
      MICON           =   "frmFeriados.frx":096B
      PICN            =   "frmFeriados.frx":0987
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdDelFeriado 
      Height          =   315
      Left            =   3870
      TabIndex        =   8
      ToolTipText     =   "Remove Feriado"
      Top             =   750
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Remover"
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
      MICON           =   "frmFeriados.frx":0AE1
      PICN            =   "frmFeriados.frx":0AFD
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdLookAll 
      Height          =   315
      Left            =   4440
      TabIndex        =   9
      ToolTipText     =   "Consulta os Feriados Cadastrados"
      Top             =   2430
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "Consul&tar Todos"
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
      MICON           =   "frmFeriados.frx":0C57
      PICN            =   "frmFeriados.frx":0C73
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancelAll 
      Height          =   330
      Left            =   5100
      TabIndex        =   10
      ToolTipText     =   "Cancela Seleção"
      Top             =   4215
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
      BTYPE           =   3
      TX              =   "C&ancelar"
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
      MICON           =   "frmFeriados.frx":0DCD
      PICN            =   "frmFeriados.frx":0DE9
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSelAll 
      Height          =   330
      Left            =   3660
      TabIndex        =   11
      ToolTipText     =   "Retorna o Registro Selecionado"
      Top             =   4215
      Visible         =   0   'False
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   582
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
      MICON           =   "frmFeriados.frx":0F43
      PICN            =   "frmFeriados.frx":0F5F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label3 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "Observação:"
      Height          =   225
      Left            =   2700
      TabIndex        =   6
      Top             =   1290
      Width           =   1575
   End
   Begin VB.Label Label2 
      BackColor       =   &H00EEEEEE&
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione o Feriado:"
      Height          =   225
      Left            =   2700
      TabIndex        =   4
      Top             =   90
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Feriados Cadastrados no Mês"
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   60
      TabIndex        =   2
      Top             =   2520
      Width           =   2325
   End
End
Attribute VB_Name = "frmFeriados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   grdFer.Enabled = True
   cmbFeriado.Enabled = True
   cmbFeriado.BackColor = vbWhite
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   grdFer.Enabled = False
   If Evento <> "Novo" Then
      cmbFeriado.Enabled = False
      cmbFeriado.BackColor = Kde
   Else
      cmbFeriado.Enabled = True
      cmbFeriado.BackColor = vbWhite
   End If
End If

FormHagana

End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False

End Sub

Private Sub cmdAddFeriado_Click()
Dim sN As String
Dim nCod As Integer

sN = InputBox("Digite o Nome do Novo Feriado.", "Inclusão de Feriados")

If sN <> "" Then
   Sql = "SELECT MAX(CODFERIADO) AS MAXIMO FROM FERIADO"
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If IsNull(!MAXIMO) Then
          nCod = 1
       Else
          nCod = !MAXIMO + 1
       End If
      .Close
   End With
   Sql = "INSERT FERIADO (CODFERIADO,NOMEFERIADO) VALUES("
   Sql = Sql & nCod & ",'" & sN & "')"
   cn.Execute Sql, rdExecDirect
   cmbFeriado.AddItem sN
   cmbFeriado.ItemData(cmbFeriado.NewIndex) = nCod
   cmbFeriado.ListIndex = cmbFeriado.ListCount - 1
End If

End Sub

Private Sub cmdAlterar_Click()
    If grdFer.Rows = 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""

End Sub

Private Sub cmdCancelAll_Click()
cmdSelAll.Visible = False
cmdCancelAll.Visible = False

frAll.Visible = False
frButtons.Enabled = True
End Sub

Private Sub cmdDelFeriado_Click()

If cmbFeriado.ListIndex = -1 Then
   MsgBox "Selecione o feriado a ser excluído.", vbExclamation, "atenção"
   cmbFeriado.SetFocus
   Exit Sub
End If

If MsgBox("Excluir o Feriado " & cmbFeriado.text & " ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
   Sql = "SELECT CODFERIADO FROM FERIADODEF WHERE CODFERIADO=" & cmbFeriado.ItemData(cmbFeriado.ListIndex)
   Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
   With RdoAux
       If .RowCount > 0 Then
          MsgBox "Não é possível excluir este feriado pois ja existem Datas Cadastradas com este feriado.", vbExclamation, "Atenção"
         .Close
          Exit Sub
       End If
      .Close
   End With

   Sql = "DELETE FROM FERIADO WHERE CODFERIADO=" & cmbFeriado.ItemData(cmbFeriado.ListIndex)
   cn.Execute Sql, rdExecDirect
   cmbFeriado.RemoveItem (cmbFeriado.ListIndex)
   If cmbFeriado.ListCount > 0 Then cmbFeriado.ListIndex = 0
End If

End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Erro

    If grdFer.Rows = 1 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir este Feriado ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM FERIADODEF WHERE DIA=" & Mv.Day & " AND MES=" & Mv.Month & " AND ANO=" & Mv.Year & " AND CODFERIADO=" & grdFer.TextMatrix(grdFer.Row, 0)
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído Feriado " & grdFer.TextMatrix(grdFer.Row, 2) & " do dia " & grdFer.TextMatrix(grdFer.Row, 1)
       Limpa
       CarregaLista
       Le
    End If
    
Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next

End Sub

Private Sub cmdGravar_Click()

If cmbFeriado.ListIndex = -1 Then
   MsgBox "Selecione o Feriado.", vbExclamation, "Atenção"
   Exit Sub
End If
Grava
Evento = ""
Eventos "INICIAR"

End Sub

Private Sub cmdLookAll_Click()

lstAll.Clear
Sql = "SELECT DIA, MES, ANO, FERIADODEF.CODFERIADO,"
Sql = Sql & "NOMEFERIADO,OBSERVACAO FROM FERIADODEF INNER JOIN "
Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
Sql = Sql & " ORDER BY DIA,MES,ANO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       lstAll.AddItem Format(!DIA, "00") & "/" & Format(!Mes, "00") & "/" & Format(!Ano, "0000") & " - " & !NOMEFERIADO
      .MoveNext
    Loop
End With
cmdSelAll.Visible = True
cmdCancelAll.Visible = True
frAll.Visible = True
frButtons.Enabled = False
End Sub

Private Sub cmdNovo_Click()
   Limpa
   Evento = "Novo"
   Eventos "INCLUIR"
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub cmdSelAll_Click()
Dim sData As Date
If lstAll.text = "" Then Exit Sub
sData = CDate(Left$(lstAll.text, 10))

cmdSelAll.Visible = False
cmdCancelAll.Visible = False

Mv.Day = Day(sData)
Mv.Month = Month(sData)
Mv.Year = Year(sData)
Mv_DateClick (sData)
frAll.Visible = False
frButtons.Enabled = True
End Sub

Private Sub Form_Load()
sRet = RetEventUserForm(Me.Name)
frAll.Visible = False
Eventos "INICIAR"
Sql = "SELECT CODFERIADO,NOMEFERIADO FROM FERIADO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       cmbFeriado.AddItem !NOMEFERIADO
       cmbFeriado.ItemData(cmbFeriado.NewIndex) = !CODFERIADO
      .MoveNext
    Loop
   .Close
End With
Centraliza Me
Mv.Day = Day(Now)
Mv.Month = Month(Now)
Mv.Year = Year(Now)
CarregaLista
End Sub

Private Sub Le()
Limpa
If grdFer.Rows = 1 Then Exit Sub
cmbFeriado.text = grdFer.TextMatrix(grdFer.Row, 2)
txtObs.text = grdFer.TextMatrix(grdFer.Row, 3)

End Sub

Private Sub Limpa()
txtObs.text = ""
End Sub

Private Sub CarregaLista()
grdFer.Rows = 1

Sql = "SELECT DIA, MES, ANO, FERIADODEF.CODFERIADO,"
Sql = Sql & "NOMEFERIADO,OBSERVACAO FROM FERIADODEF INNER JOIN "
Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
Sql = Sql & "WHERE  MES = " & Mv.Month & " AND ANO =" & Mv.Year
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       grdFer.AddItem !CODFERIADO & Chr(9) & Format(!DIA, "00") & "/" & Format(!Mes, "00") & "/" & Format(!Ano, "0000") & Chr(9) & !NOMEFERIADO & Chr(9) & SubNull(!OBSERVACAO)
      .MoveNext
    Loop
   .Close
End With
Le
End Sub

Private Sub Grava()
On Error GoTo Erro
If Evento = "Novo" Then
    Sql = "INSERT FERIADODEF(DIA,MES,ANO,CODFERIADO,OBSERVACAO) VALUES("
    Sql = Sql & Mv.Day & "," & Mv.Month & "," & Mv.Year & "," & cmbFeriado.ItemData(cmbFeriado.ListIndex) & ",'" & txtObs.text & "')"
    cn.Execute Sql, rdExecDirect
Else
    Sql = "UPDATE FERIADODEF SET OBSERVACAO='" & Mask(txtObs.text) & "' WHERE "
    Sql = Sql & "DIA=" & Mv.Day & " AND MES=" & Mv.Month & " AND ANO=" & Mv.Year & " AND CODFERIADO=" & cmbFeriado.ItemData(cmbFeriado.ListIndex)
    cn.Execute Sql, rdExecDirect
End If


If Evento = "Novo" Then
   grdFer.AddItem cmbFeriado.ItemData(cmbFeriado.ListIndex) & Chr(9) & Mv.Value & Chr(9) & cmbFeriado.text & Chr(9) & txtObs.text
   grdFer.Row = grdFer.Rows - 1
   grdFer.ColSel = 3
   Log Form, Me.Caption, Inclusão, "Inserido Feriado " & cmbFeriado.text & " na data " & Mv.Value
 ElseIf Evento = "Alterar" Then
   grdFer.TextMatrix(grdFer.Row, 3) = txtObs.text
   Log Form, Me.Caption, Alteração, "Alterado Feriado " & cmbFeriado.text & " na data " & Mv.Value
End If

Exit Sub
Erro:
MsgBox "feriado já cadastrado", vbCritical, "Erro"

End Sub

Private Sub grdFer_Click()
Le
End Sub

Private Sub Mv_Click()
CarregaLista
End Sub

Private Sub Mv_DateClick(ByVal DateClicked As Date)
CarregaLista
End Sub

