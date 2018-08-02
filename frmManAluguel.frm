VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmManAluguel 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Manutenção de Aluguéis"
   ClientHeight    =   5025
   ClientLeft      =   6930
   ClientTop       =   3795
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5025
   ScaleWidth      =   7500
   Begin VB.ComboBox cmbTipo 
      Height          =   315
      ItemData        =   "frmManAluguel.frx":0000
      Left            =   1455
      List            =   "frmManAluguel.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   22
      Top             =   90
      Width           =   5835
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   6
      TabIndex        =   11
      Top             =   2925
      Width           =   1080
   End
   Begin VB.TextBox txtValor 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   12
      TabIndex        =   10
      Top             =   3495
      Width           =   1080
   End
   Begin VB.TextBox txtDesc 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   40
      TabIndex        =   9
      Top             =   4125
      Width           =   5400
   End
   Begin VB.TextBox txtPercMulta 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1995
      MaxLength       =   10
      TabIndex        =   8
      Top             =   3810
      Width           =   1080
   End
   Begin VB.TextBox txtPercJuros 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   5040
      MaxLength       =   10
      TabIndex        =   7
      Top             =   3810
      Width           =   1080
   End
   Begin prjChameleon.chameleonButton cmdCnsImovel 
      Height          =   300
      Left            =   3120
      TabIndex        =   12
      ToolTipText     =   "Consulta Imóvel"
      Top             =   2910
      Width           =   405
      _ExtentX        =   714
      _ExtentY        =   529
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmManAluguel.frx":0004
      PICN            =   "frmManAluguel.frx":0020
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskDataVencto 
      Height          =   285
      Left            =   5040
      TabIndex        =   13
      Top             =   3495
      Width           =   1080
      _ExtentX        =   1905
      _ExtentY        =   503
      MouseIcon       =   "frmManAluguel.frx":017A
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
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   6360
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   4590
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
      MICON           =   "frmManAluguel.frx":0196
      PICN            =   "frmManAluguel.frx":01B2
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
      Left            =   6345
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   4590
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
      MICON           =   "frmManAluguel.frx":030C
      PICN            =   "frmManAluguel.frx":0328
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
      Left            =   5265
      TabIndex        =   2
      ToolTipText     =   "Gravar os Dados"
      Top             =   4590
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmManAluguel.frx":0396
      PICN            =   "frmManAluguel.frx":03B2
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
      Left            =   2250
      TabIndex        =   3
      ToolTipText     =   "Excluir Registro"
      Top             =   4590
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
      MICON           =   "frmManAluguel.frx":0757
      PICN            =   "frmManAluguel.frx":0773
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
      Left            =   1170
      TabIndex        =   4
      ToolTipText     =   "Editar Registro"
      Top             =   4590
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
      MICON           =   "frmManAluguel.frx":0815
      PICN            =   "frmManAluguel.frx":0831
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
      Left            =   90
      TabIndex        =   5
      ToolTipText     =   "Novo Registro"
      Top             =   4590
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
      MICON           =   "frmManAluguel.frx":098B
      PICN            =   "frmManAluguel.frx":09A7
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdMain 
      Height          =   2235
      Left            =   15
      TabIndex        =   0
      Top             =   525
      Width           =   7470
      _ExtentX        =   13176
      _ExtentY        =   3942
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmManAluguel.frx":0B01
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Tipo de Aluguel..:"
      Height          =   195
      Index           =   1
      Left            =   120
      TabIndex        =   23
      Top             =   135
      Width           =   1320
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Código Reduzido/I.M...:"
      Height          =   225
      Index           =   0
      Left            =   210
      TabIndex        =   21
      Top             =   2970
      Width           =   1785
   End
   Begin VB.Label lblNome 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   2010
      TabIndex        =   20
      Top             =   3255
      Width           =   5445
   End
   Begin VB.Label lblRS 
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Inquilino........:"
      Height          =   225
      Left            =   210
      TabIndex        =   19
      Top             =   3255
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor Aluguel Anual......:"
      Height          =   225
      Left            =   210
      TabIndex        =   18
      Top             =   3555
      Width           =   1695
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Data 1º Vencimento.....:"
      Height          =   225
      Left            =   3255
      TabIndex        =   17
      Top             =   3540
      Width           =   1695
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Descrição Resumida....:"
      Height          =   225
      Left            =   210
      TabIndex        =   16
      Top             =   4170
      Width           =   1695
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentual de Multa.....:"
      Height          =   225
      Left            =   210
      TabIndex        =   15
      Top             =   3870
      Width           =   1695
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Percentual de Juros.....:"
      Height          =   225
      Left            =   3255
      TabIndex        =   14
      Top             =   3870
      Width           =   1695
   End
End
Attribute VB_Name = "frmManAluguel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public WithEvents m_cMenuContrib As cPopupMenu
Attribute m_cMenuContrib.VB_VarHelpID = -1
Dim RdoAux As rdoResultset, Sql As String, Evento As String

Private Sub cmbTipo_Click()
Limpa
txtCod.Text = ""
If cmbTipo.ListIndex = -1 Then Exit Sub
grdMain.Rows = 1
Sql = "SELECT CODREDUZIDO,NOME,DESCRICAO FROM MANUTENCAOALUGUEL WHERE CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddItem Format(!CODREDUZIDO, "000000") & Chr(9) & !Nome & Chr(9) & SubNull(!descricao)
       .MoveNext
    Loop
End With

If grdMain.Rows > 1 Then
    grdMain_Click
End If

End Sub

Private Sub cmdAlterar_Click()

If lblNome.Caption = "" Then
    MsgBox "Selecione o inquilino", vbExclamation, "Atenção"
    Exit Sub
End If

Evento = "Alterar"
Eventos "INCLUIR"
txtValor.SetFocus
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
Evento = ""
End Sub

Private Sub cmdCnsImovel_Click()
If Evento <> "Novo" Then Exit Sub

lIndex = m_cMenuContrib.ShowPopupMenu(cmdCnsImovel.Left, cmdCnsImovel.Top, cmdCnsImovel.Left, cmdCnsImovel.Top, Me.ScaleWidth - cmdCnsImovel.Left - cmdCnsImovel.Width, cmdCnsImovel.Top + cmdCnsImovel.Height, False)
End Sub

Private Sub cmdExcluir_Click()
If lblNome.Caption = "" Then
    MsgBox "Selecione o inquilino", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Excluir este inquilino ?", vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    Sql = "DELETE FROM MANUTENCAOALUGUEL WHERE CODREDUZIDO=" & Val(grdMain.TextMatrix(grdMain.Row, 0)) & " AND "
    Sql = Sql & "CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
    cn.Execute Sql, rdExecDirect
    Limpa
    If grdMain.Rows > 2 Then
        grdMain.RemoveItem grdMain.Row
        grdMain.Row = 1
        grdMain_Click
    Else
        grdMain.Rows = 1
    End If
End If

End Sub

Private Sub cmdGravar_Click()
Dim df As Integer
If lblNome.Caption = "" Then
   MsgBox "Selecione o Inquilino.", vbExclamation, "Atenção"
   Exit Sub
End If

If Val(txtPercJuros.Text) = 0 Then txtPercJuros.Text = 0
If Val(txtPercMulta.Text) = 0 Then txtPercMulta.Text = 0

If Val(txtPercJuros.Text) > 90 Then
    MsgBox "Juros inválidos.", vbExclamation, "atenção"
    Exit Sub
End If

If Val(txtPercMulta.Text) > 90 Then
    MsgBox "Multa inválida.", vbExclamation, "atenção"
    Exit Sub
End If

If Val(txtValor.Text) = 0 Then
   MsgBox "Digite o valor do aluguel anual.", vbExclamation, "Atenção"
   Exit Sub
End If

If Not IsDate(mskDataVencto.Text) Then
    MsgBox "Data de vencimento inválida.", vbExclamation, "atenção"
    Exit Sub
End If

df = ValidaFeriado(CDate(mskDataVencto.Text))
If df = 1 Then
    If MsgBox("Data do 1º Vencimento cai no Domingo." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencto.Text = Format(RetornaDiaUtil(CDate(mskDataVencto.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 2 Then
    If MsgBox("Data do 1º Vencimento cai no sábado." & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
        mskDataVencto.Text = Format(RetornaDiaUtil(CDate(mskDataVencto.Text)), "dd/mm/yyyy")
    Else
        Exit Sub
    End If
ElseIf df = 3 Then
    Sql = "SELECT NOMEFERIADO FROM FERIADODEF INNER JOIN "
    Sql = Sql & "FERIADO ON FERIADODEF.CODFERIADO = FERIADO.CODFERIADO "
    Sql = Sql & " Where DIA = " & Day(CDate(mskDataVencto.Text))
    Sql = Sql & " AND MES=" & Month(CDate(mskDataVencto.Text)) & " AND ANO=" & Year(CDate(mskDataVencto.Text))
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            If MsgBox("Data do 1º Vencimento cai no Feriado (" & !NOMEFERIADO & ")" & vbCrLf & "Próximo Dia Util é " & RetornaDiaUtil(CDate(mskDataVencto.Text)) & vbCrLf & vbCrLf & "Aceitar esta Data?", vbQuestion + vbYesNo, "Atenção") = vbYes Then
                mskDataVencto.Text = RetornaDiaUtil(CDate(mskDataVencto.Text))
            Else
                Exit Sub
            End If
          .Close
        End If
    End With
End If

If Evento = "Novo" Then
    Sql = "SELECT * FROM MANUTENCAOALUGUEL WHERE CODREDUZIDO=" & Val(txtCod.Text) & " AND CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            MsgBox "Inquilino ja cadastrado neste tipo de aluguel.", vbExclamation, "atenção"
            Exit Sub
        Else
            Sql = "INSERT MANUTENCAOALUGUEL(CODREDUZIDO,CODLANCAMENTO,NOME,VALORTOTAL,DATAVENCTO,MULTA,JUROS,DESCRICAO) VALUES("
            Sql = Sql & Val(txtCod.Text) & "," & cmbTipo.ItemData(cmbTipo.ListIndex) & ",'" & Left(Mask(lblNome.Caption), 50) & "'," & Virg2Ponto(txtValor.Text) & ",'"
            Sql = Sql & Format(mskDataVencto.Text, "mm/dd/yyyy") & "'," & Virg2Ponto(txtPercMulta.Text) & ","
            Sql = Sql & Virg2Ponto(txtPercJuros.Text) & ",'" & Mask(txtDesc.Text) & "')"
            cn.Execute Sql, rdExecDirect
            grdMain.AddItem Format(txtCod.Text, "000000") & Chr(9) & lblNome.Caption & Chr(9) & Mask(txtDesc.Text)
        End If
       .Close
    End With
Else
    Sql = "UPDATE MANUTENCAOALUGUEL SET VALORTOTAL=" & Virg2Ponto(RemovePonto(txtValor.Text)) & " ,DATAVENCTO='" & Format(mskDataVencto.Text, "mm/dd/yyyy") & "' ,"
    Sql = Sql & "MULTA=" & Virg2Ponto(txtPercMulta.Text) & " ,JUROS=" & Virg2Ponto(txtPercJuros.Text) & " ,DESCRICAO='" & Mask(txtDesc.Text) & "' "
    Sql = Sql & "Where CODREDUZIDO = " & Val(grdMain.TextMatrix(grdMain.Row, 0)) & " And CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
    cn.Execute Sql, rdExecDirect
    grdMain.TextMatrix(grdMain.Row, 2) = Mask(txtDesc.Text)
End If

Eventos "INICIAR"
Evento = ""
End Sub


Private Sub cmdNovo_Click()
Evento = "Novo"
Eventos "INCLUIR"
Limpa
txtCod.Text = ""
txtCod.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Activate()
If Val(CodImovel) > 0 Then
     txtCod.Text = Val(Left$(CodImovel, 7))
     CodImovel = 0
     txtCod_LostFocus
Else
    If Val(CodEmpresa) > 0 Then
         txtCod.Text = Val(Left$(CodEmpresa, 7))
         CodEmpresa = 0
         txtCod_LostFocus
    Else
        If Val(CodCidadao) > 0 Then
             Unload frmCnsCidadao
             If cGetInputState() <> 0 Then DoEvents
             txtCod.Text = Val(CodCidadao)
             CodCidadao = 0
             txtCod_LostFocus
        End If
    End If
End If

End Sub

Private Sub Form_Load()
Centraliza Me
MontaMenu
Sql = "SELECT TIPOALUGUEL.CODLANCAMENTO,LANCAMENTO.DESCFULL FROM TIPOALUGUEL INNER JOIN "
Sql = Sql & "LANCAMENTO ON TIPOALUGUEL.CODLANCAMENTO = LANCAMENTO.CODLANCAMENTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbTipo.AddItem !DESCFULL
        cmbTipo.ItemData(cmbTipo.NewIndex) = !CodLancamento
       .MoveNext
    Loop
   .Close
End With
If cmbTipo.ListCount > 0 Then cmbTipo.ListIndex = 0
Eventos "INICIAR"
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Set m_cMenuContrib = Nothing
End Sub

Private Sub grdMain_Click()
Limpa
With grdMain
    If .Rows = 1 Then Exit Sub
    Sql = "SELECT * FROM MANUTENCAOALUGUEL WHERE CODREDUZIDO=" & Val(.TextMatrix(.Row, 0)) & " AND "
    Sql = Sql & "CODLANCAMENTO=" & cmbTipo.ItemData(cmbTipo.ListIndex)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        txtCod.Text = !CODREDUZIDO
        lblNome.Caption = !Nome
        txtValor.Text = FormatNumber(!ValorTotal, 2)
        mskDataVencto.Text = Format(!DataVencto, "dd/mm/yyyy")
        txtPercJuros.Text = FormatNumber(!Juros, 2)
        txtPercMulta.Text = FormatNumber(!Multa, 2)
        txtDesc.Text = SubNull(!descricao)
    End With
End With

End Sub

Private Sub m_cMenuContrib_Click(ItemNumber As Long)
Select Case m_cMenuContrib.ItemKey(ItemNumber)
    Case "mnuMob"
        sFormMob = "ALUGUEL"
        frmCnsMob.show
        frmCnsMob.ZOrder 0
    Case "mnuImob"
        sForm = "ALUGUEL"
        frmCnsImovel.show
        frmCnsImovel.ZOrder 0
    Case "mnuOutros"
        Set frm = frmCnsCidadao
        frm.sForm = "ALUGUEL"
        frm.show
        frm.ZOrder 0
End Select

End Sub

Private Sub txtPercJuros_KeyPress(KeyAscii As Integer)
Tweak txtPercJuros, KeyAscii, DecimalPositive
End Sub

Private Sub txtPercMulta_KeyPress(KeyAscii As Integer)
Tweak txtPercMulta, KeyAscii, DecimalPositive
End Sub

Private Sub txtValor_KeyPress(KeyAscii As Integer)
Tweak txtValor, KeyAscii, DecimalPositive
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
   For Each Ct In frmManAluguel
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Then
           Ct.BackColor = Kde
           Ct.Locked = True
       End If
   Next
   cmbTipo.Locked = False
   cmbTipo.BackColor = Branco
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   For Each Ct In frmManAluguel
       If TypeOf Ct Is TextBox Or TypeOf Ct Is esMaskedEdit Then
          Ct.BackColor = Branco
          Ct.Locked = False
       End If
   Next
   cmbTipo.Locked = True
   cmbTipo.BackColor = Kde
   If Evento = "Alterar" Then
      txtCod.Locked = False
      txtCod.BackColor = Kde
   End If
End If

End Sub

Private Sub txtCod_GotFocus()

txtCod.SelStart = 0
txtCod.SelLength = Len(txtCod.Text)

End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = 13 Then
    KeyAscii = 0
    txtCod_LostFocus
    Exit Sub
End If

Tweak txtCod, KeyAscii, IntegerPositive

End Sub

Private Sub txtCod_LostFocus()
Dim nCodImovel As Long
Limpa
If Val(txtCod.Text) = 0 Then Exit Sub
nCodImovel = Val(txtCod.Text)

Sql = "SELECT NOMECIDADAO,INATIVO FROM vwCONSULTAIMOVELPROP WHERE CODREDUZIDO=" & nCodImovel
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount > 0 Then
        If !Inativo = 1 Then
           MsgBox "Este imóvel encontra-se inativo.", vbExclamation, "Atenção"
           Exit Sub
        End If
        lblNome.Caption = !nomecidadao
    Else
        Sql = "SELECT CODIGOMOB,INSCESTADUAL,RAZAOSOCIAL,ABREVTIPOLOG,ABREVTITLOG,NOMELOGRADOURO,NUMERO,CODBAIRRO,CEP,COMPLEMENTO,DATAENCERRAMENTO "
        Sql = Sql & " FROM vwCNSMOBILIARIO WHERE CODIGOMOB=" & nCodImovel
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
               If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
                  MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
                  Exit Sub
               End If
              'suspenção
               Sql = "SELECT CODTIPOEVENTO,DATAPROCEVENTO FROM MOBILIARIOEVENTO WHERE CODMOBILIARIO=" & txtCod.Text
               Sql = Sql & " ORDER BY DATAEVENTO DESC"
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux2
                   If .RowCount > 0 Then
                       If !CODTIPOEVENTO = 2 Then
                           MsgBox "Esta empresa esta SUSPENSA", vbExclamation, "Atenção"
                           Exit Sub
                       End If
                   End If
                  .Close
               End With
               lblNome.Caption = !RazaoSocial
            Else
               Sql = "SELECT CIDADAO.CODCIDADAO,CIDADAO.NOMECIDADAO,CIDADAO.CPF, CIDADAO.CNPJ, CIDADAO.CODLOGRADOURO,vwLOGRADOURO.ABREVTIPOLOG,"
               Sql = Sql & "vwLOGRADOURO.ABREVTITLOG,vwLOGRADOURO.NOMELOGRADOURO,CIDADAO.NUMIMOVEL, CIDADAO.COMPLEMENTO,CIDADAO.CODBAIRRO, BAIRRO.DESCBAIRRO,"
               Sql = Sql & "CIDADAO.CODCIDADE, CIDADE.DESCCIDADE,CIDADAO.SIGLAUF, UF.DESCUF, CIDADAO.CEP,CIDADAO.NOMELOGRADOURO AS RUA2 "
               Sql = Sql & "FROM vwLOGRADOURO RIGHT OUTER JOIN CIDADAO ON vwLOGRADOURO.CODLOGRADOURO = CIDADAO.CODLOGRADOURO "
               Sql = Sql & "LEFT OUTER JOIN CIDADE INNER JOIN BAIRRO ON CIDADE.SIGLAUF = BAIRRO.SIGLAUF AND CIDADE.CODCIDADE = BAIRRO.CODCIDADE INNER JOIN "
               Sql = Sql & "UF ON CIDADE.SIGLAUF = UF.SIGLAUF ON CIDADAO.SIGLAUF = BAIRRO.SIGLAUF AND CIDADAO.CODCIDADE = BAIRRO.CODCIDADE AND CIDADAO.CODBAIRRO = BAIRRO.CODBAIRRO WHERE CODCIDADAO=" & Val(txtCod.Text)
               Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               With RdoAux
                   If .RowCount > 0 Then
                       lblNome.Caption = !nomecidadao
                   Else
                       MsgBox "Código não cadastrado.", vbCritical, "Atenção"
                   End If
                  .Close
               End With
            End If
           .Close
        End With
    End If
End With
End Sub

Private Sub Limpa()

lblNome.Caption = ""
LimpaMascara mskDataVencto
txtValor.Text = ""
txtPercJuros.Text = ""
txtPercMulta.Text = ""
txtDesc.Text = ""

End Sub

Private Sub MontaMenu()

   Set m_cMenuContrib = New cPopupMenu
   With m_cMenuContrib
      .hwndOwner = Me.hwnd
      .GradientHighlight = True
      
      i = .AddItem("Mobiliário", "", 1, , , , , "mnuMob")
      .OwnerDraw(i) = True
      i = .AddItem("Imobiliário", "", 1, , , , , "mnuImob")
      .OwnerDraw(i) = True
      i = .AddItem("Outros", "", 1, , , , , "mnuOutros")
      .OwnerDraw(i) = True
   End With
   
End Sub


