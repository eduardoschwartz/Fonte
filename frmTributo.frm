VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmTributo 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Tributos"
   ClientHeight    =   5370
   ClientLeft      =   3675
   ClientTop       =   2370
   ClientWidth     =   10950
   Icon            =   "frmTributo.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5370
   ScaleWidth      =   10950
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5760
      TabIndex        =   33
      ToolTipText     =   "Sair da Tela"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":014A
      PICN            =   "frmTributo.frx":0166
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
      Left            =   5760
      TabIndex        =   28
      ToolTipText     =   "Cancelar Edição"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":01D4
      PICN            =   "frmTributo.frx":01F0
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
      Left            =   120
      TabIndex        =   29
      ToolTipText     =   "Novo Registro"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":034A
      PICN            =   "frmTributo.frx":0366
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
      TabIndex        =   30
      ToolTipText     =   "Editar Registro"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":04C0
      PICN            =   "frmTributo.frx":04DC
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
      Left            =   2220
      TabIndex        =   31
      ToolTipText     =   "Excluir Registro"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":0636
      PICN            =   "frmTributo.frx":0652
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
      Left            =   4710
      TabIndex        =   32
      ToolTipText     =   "Gravar os Dados"
      Top             =   4920
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
      MICON           =   "frmTributo.frx":06F4
      PICN            =   "frmTributo.frx":0710
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00EEEEEE&
      Height          =   1785
      Left            =   6870
      TabIndex        =   19
      Top             =   3510
      Width           =   4035
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   8
         Left            =   3090
         TabIndex        =   13
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   7
         Left            =   3090
         TabIndex        =   12
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   6
         Left            =   3090
         TabIndex        =   11
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   5
         Left            =   3090
         TabIndex        =   10
         Top             =   210
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   4
         Left            =   1140
         TabIndex        =   9
         Top             =   1200
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   3
         Left            =   1140
         TabIndex        =   8
         Top             =   870
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   2
         Left            =   1140
         TabIndex        =   7
         Top             =   540
         Width           =   855
      End
      Begin VB.TextBox txtFicha 
         Appearance      =   0  'Flat
         Height          =   285
         Index           =   1
         Left            =   1140
         TabIndex        =   6
         Top             =   210
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A.J./Encar...:"
         Height          =   195
         Index           =   9
         Left            =   2100
         TabIndex        =   27
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "A.J./Jur.Mul.:"
         Height          =   195
         Index           =   8
         Left            =   2100
         TabIndex        =   26
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Ajuizado.......:"
         Height          =   195
         Index           =   7
         Left            =   2100
         TabIndex        =   25
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "D.A./Encar..:"
         Height          =   195
         Index           =   6
         Left            =   2100
         TabIndex        =   24
         Top             =   255
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "D.A./Jur.Mul:"
         Height          =   195
         Index           =   5
         Left            =   90
         TabIndex        =   23
         Top             =   1245
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Divida Ativa.:"
         Height          =   195
         Index           =   4
         Left            =   90
         TabIndex        =   22
         Top             =   915
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Juros/Multa..:"
         Height          =   195
         Index           =   3
         Left            =   90
         TabIndex        =   21
         Top             =   585
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº de Ficha..:"
         Height          =   195
         Index           =   2
         Left            =   90
         TabIndex        =   20
         Top             =   255
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   1245
      Left            =   60
      TabIndex        =   14
      Top             =   3510
      Width           =   6795
      Begin VB.CheckBox chkMulta 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Multa"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5820
         TabIndex        =   3
         Top             =   210
         Width           =   780
      End
      Begin VB.CheckBox chkJuros 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Juros"
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   4950
         TabIndex        =   2
         Top             =   210
         Width           =   690
      End
      Begin VB.CheckBox chkDA 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Divida Ativa"
         Height          =   225
         Left            =   2820
         TabIndex        =   1
         Top             =   225
         Width           =   1200
      End
      Begin VB.TextBox txtDescR 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   25
         TabIndex        =   4
         Top             =   540
         Width           =   5265
      End
      Begin VB.TextBox txtCod 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         TabIndex        =   0
         TabStop         =   0   'False
         Top             =   210
         Width           =   1005
      End
      Begin VB.TextBox txtDescC 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1410
         MaxLength       =   100
         TabIndex        =   5
         Top             =   870
         Width           =   5265
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Resumida..:"
         Height          =   195
         Index           =   11
         Left            =   60
         TabIndex        =   17
         Top             =   585
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código................:"
         Height          =   195
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   255
         Width           =   1275
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Desc.Completa...:"
         Height          =   195
         Index           =   1
         Left            =   60
         TabIndex        =   15
         Top             =   930
         Width           =   1275
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdTrib 
      Height          =   3420
      Left            =   60
      TabIndex        =   18
      Top             =   30
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   6033
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      BackColorFixed  =   15658734
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   $"frmTributo.frx":0AB5
   End
End
Attribute VB_Name = "frmTributo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sOldDesc As String
Dim RdoAux As rdoResultset
Dim Sql As String
Dim Evento As String
Dim m_SortColumn As Integer
Dim sRet As String
Dim evEdit As Integer, evNew As Integer, evDel As Integer, evEsp As Integer
Dim bEdit As Boolean, bNew As Boolean, bDel As Boolean, bEsp As Boolean

Private Sub cmdAlterar_Click()
    If Val(txtCod.Text) = 0 Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
'    If Val(txtCod.text) < 149 Then
'       Evento = "DA"
'       Eventos "INCLUIR"
'       Exit Sub
'    End If
    sOldDesc = txtDescR.Text
    Eventos "INCLUIR"
    Evento = "Alterar"
End Sub

Private Sub cmdCancel_Click()
    Le
    Eventos "INICIAR"
    Evento = ""
End Sub

Private Sub cmdExcluir_Click()
On Error GoTo Erro
    If txtCod.Text = "" Then
       MsgBox "Não existem Registros.", vbCritical, "Atenção"
       Exit Sub
    End If
    If Val(txtCod.Text) < 149 Then
       MsgBox "Tributos menores que 149 só podem ser excluidos pelo Administrador do Sistema.", vbCritical, "Atenção"
       Exit Sub
    End If
    
    If MsgBox("Excluir este Tributo ?", vbQuestion + vbYesNoCancel, "Atenção") = vbYes Then
       Sql = "DELETE FROM TRIBUTO WHERE CODTRIBUTO=" & txtCod.Text
       cn.Execute Sql, rdExecDirect
       Log Form, Me.Caption, Exclusão, "Excluído registro " & Format(txtCod.Text, "000") & "-" & txtDescR.Text
       Limpa
       CarregaLista
       Le
    End If
Exit Sub
Erro:
MsgBox Err.Description

End Sub

Private Sub cmdGravar_Click()
    If txtDescR.Text = "" Then
       MsgBox "Favor digitar a Descrição Resumida.", vbExclamation, "Atenção"
       txtDescR.SetFocus
       Exit Sub
    End If
    If txtDescC.Text = "" Then
       MsgBox "Favor digitar a Descrição Completa.", vbExclamation, "Atenção"
       txtDescC.SetFocus
       Exit Sub
    End If
    Grava
    Eventos "INICIAR"
End Sub


Private Sub cmdNovo_Click()
    Limpa
    Eventos "INCLUIR"
    Evento = "Novo"
End Sub

Private Sub cmdSair_Click()
'MudaFicha
   Unload Me
End Sub


Private Sub Form_Activate()
Liberado
End Sub

Private Sub Form_Load()

Me.Left = frmMdi.ScaleWidth / 2 - Me.Width / 2
Me.Top = frmMdi.ScaleHeight / 2 - Me.Height / 2
Centraliza Me
sRet = RetEventUserForm(Me.Name)
grdTrib.Rows = 1
CarregaLista
Le

Eventos "INICIAR"

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
   For Each Ct In frmTributo
       If TypeOf Ct Is TextBox Then
          Ct.BackColor = Kde
          Ct.Enabled = False
       End If
   Next
   chkDA.Enabled = False
   chkJuros.Enabled = False
   chkMulta.Enabled = False
   grdTrib.Enabled = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   If Evento = "DA" Then
        grdTrib.Enabled = False
        chkDA.Enabled = True
        chkJuros.Enabled = True
        chkMulta.Enabled = True
   Else
        For Each Ct In frmTributo
            If TypeOf Ct Is TextBox Then
               Ct.BackColor = Branco
               Ct.Enabled = True
            End If
        Next
        chkDA.Enabled = True
        chkJuros.Enabled = True
        chkMulta.Enabled = True
        txtCod.BackColor = Kde
        txtCod.Locked = True
        grdTrib.Enabled = False
        txtDescR.SetFocus
  End If
End If

FormHagana

End Sub

Private Sub Le()

If grdTrib.Row = 0 Then Exit Sub
txtCod.Text = grdTrib.TextMatrix(grdTrib.Row, 0)
txtDescR.Text = grdTrib.TextMatrix(grdTrib.Row, 2)
txtDescC.Text = grdTrib.TextMatrix(grdTrib.Row, 1)
chkDA.value = Val(grdTrib.TextMatrix(grdTrib.Row, 3))
chkJuros.value = Val(grdTrib.TextMatrix(grdTrib.Row, 4))
chkMulta.value = Val(grdTrib.TextMatrix(grdTrib.Row, 5))
txtFicha(1).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 6))
txtFicha(2).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 7))
txtFicha(3).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 8))
txtFicha(4).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 9))
txtFicha(5).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 10))
txtFicha(6).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 11))
txtFicha(7).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 12))
txtFicha(8).Text = Val(grdTrib.TextMatrix(grdTrib.Row, 13))

End Sub

Private Sub Limpa()
txtCod.Text = ""
txtDescR.Text = ""
txtDescC.Text = ""
chkDA.value = vbUnchecked
chkJuros.value = vbUnchecked
chkMulta.value = vbUnchecked
txtFicha(1).Text = "": txtFicha(2).Text = "": txtFicha(3).Text = "": txtFicha(4).Text = ""
txtFicha(5).Text = "": txtFicha(6).Text = "": txtFicha(7).Text = "": txtFicha(8).Text = ""

End Sub

Private Sub CarregaLista()

Sql = "Select CODTRIBUTO,DESCTRIBUTO,ABREVTRIBUTO,DA,JUROS,MULTA,FICHA,FICHAJRMULTA,FICHADIVIDA,FICHADAJRMUL,FICHADAENCA,FICHAAJUIZA,FICHAAJJRMUL,FICHAAJENCA "
Sql = Sql & "From TRIBUTO WHERE CODTRIBUTO <>129 ORDER BY DESCTRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

grdTrib.Rows = 1
With RdoAux
   .MoveFirst
    Do Until .EOF
       grdTrib.AddItem !CodTributo & Chr(9) & UCase$(!desctributo) & Chr(9) & UCase$(!ABREVTRIBUTO) & Chr(9) & IIf(!DA = True, 1, 0) & Chr(9) & IIf(!Juros = True, 1, 0) & Chr(9) & IIf(!Multa = True, 1, 0) & Chr(9) & _
       Val(SubNull(!Ficha)) & Chr(9) & Val(SubNull(!FichaJrMulta)) & Chr(9) & Val(SubNull(!FichaDivida)) & Chr(9) & Val(SubNull(!FichaDaJrMul)) & Chr(9) & Val(SubNull(!FichaDaEnca)) & Chr(9) & Val(SubNull(!FichaAjuiza)) & Chr(9) & _
       Val(SubNull(!FichaAjJrMul)) & Chr(9) & Val(SubNull(!FichaAjEnca))
      .MoveNext
    Loop
   .Close
End With

End Sub

Private Sub Grava()

Dim MaxCod As Integer

Sql = "SELECT MAX(CODTRIBUTO) AS MAXIMO FROM TRIBUTO"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If IsNull(RdoAux!maximo) Then
    MaxCod = 1
Else
    MaxCod = RdoAux!maximo + 1
End If
RdoAux.Close

If Evento = "Novo" Then
    Sql = "INSERT TRIBUTO(CODTRIBUTO,DESCTRIBUTO,ABREVTRIBUTO,DA,JUROS,Multa,FICHA,FICHAJRMULTA,FICHADIVIDA,FICHADAJRMUL,FICHADAENCA,FICHAAJUIZA,FICHAAJJRMUL,FICHAAJENCA) VALUES("
    Sql = Sql & MaxCod & ",'" & Mask(txtDescC.Text) & "','" & Left(Mask(txtDescR.Text), 25) & "',"
    Sql = Sql & chkDA.value & "," & chkJuros.value & "," & chkMulta.value & "," & Val(txtFicha(1).Text) & ","
    Sql = Sql & Val(txtFicha(2).Text) & "," & Val(txtFicha(3).Text) & "," & Val(txtFicha(4).Text) & "," & Val(txtFicha(5).Text) & "," & Val(txtFicha(6).Text) & "," & Val(txtFicha(7).Text) & "," & Val(txtFicha(8).Text) & ")"
Else
    Sql = "UPDATE TRIBUTO SET DESCTRIBUTO='" & Mask(txtDescC.Text) & "',ABREVTRIBUTO='" & Mask(txtDescR.Text) & "',"
    Sql = Sql & "DA=" & chkDA.value & ",JUROS=" & chkJuros.value & ",MULTA=" & chkMulta.value & ",FICHA=" & Val(txtFicha(1).Text) & ","
    Sql = Sql & "FICHAJRMULTA=" & Val(txtFicha(2).Text) & ",FICHADIVIDA=" & Val(txtFicha(3).Text) & ",FICHADAJRMUL=" & Val(txtFicha(4).Text) & ",FICHADAENCA=" & Val(txtFicha(5).Text)
    Sql = Sql & " ,FICHAAJUIZA=" & Val(txtFicha(6).Text) & ",FICHAAJJRMUL=" & Val(txtFicha(7).Text) & ",FICHAAJENCA=" & Val(txtFicha(8).Text)
    Sql = Sql & " WHERE CODTRIBUTO=" & Val(txtCod.Text)
End If
cn.Execute Sql, rdExecDirect

If Evento = "Novo" Then
   grdTrib.AddItem MaxCod & Chr(9) & txtDescC.Text & Chr(9) & txtDescR.Text & Chr(9) & chkDA.value & Chr(9) & chkJuros.value & Chr(9) & chkMulta & Chr(9) & txtFicha(1).Text & Chr(9) & txtFicha(2).Text & Chr(9) & txtFicha(3).Text & Chr(9) & txtFicha(4).Text & _
   Chr(9) & txtFicha(5).Text & Chr(9) & txtFicha(6).Text & Chr(9) & txtFicha(7).Text & Chr(9) & txtFicha(8).Text
   txtCod.Text = MaxCod
   grdTrib.Row = grdTrib.Rows - 1
   grdTrib.ColSel = 1
   Log Form, Me.Caption, Inclusão, "Inserido registro " & Format(MaxCod, "000") & "-" & txtDescR.Text
 ElseIf Evento = "Alterar" Or Evento = "DA" Then
   grdTrib.TextMatrix(grdTrib.Row, 2) = txtDescR.Text
   grdTrib.TextMatrix(grdTrib.Row, 1) = txtDescC.Text
   grdTrib.TextMatrix(grdTrib.Row, 3) = chkDA.value
   grdTrib.TextMatrix(grdTrib.Row, 4) = chkJuros.value
   grdTrib.TextMatrix(grdTrib.Row, 5) = chkMulta.value
   grdTrib.TextMatrix(grdTrib.Row, 6) = txtFicha(1).Text
   grdTrib.TextMatrix(grdTrib.Row, 7) = txtFicha(2).Text
   grdTrib.TextMatrix(grdTrib.Row, 8) = txtFicha(3).Text
   grdTrib.TextMatrix(grdTrib.Row, 9) = txtFicha(4).Text
   grdTrib.TextMatrix(grdTrib.Row, 10) = txtFicha(5).Text
   grdTrib.TextMatrix(grdTrib.Row, 11) = txtFicha(6).Text
   grdTrib.TextMatrix(grdTrib.Row, 12) = txtFicha(7).Text
   grdTrib.TextMatrix(grdTrib.Row, 13) = txtFicha(8).Text
   Log Form, Me.Caption, Alteração, "Alterado registro " & Format(txtCod.Text, "000") & " de " & sOldDesc & " para " & txtDescR.Text
End If
      
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
   KeyAscii = 0
   If cmdNovo.Visible = True Then
      cmdNovo_Click
   Else
      cmdGravar_Click
   End If
End If
End Sub


Private Sub grdTrib_Click()

    ' See if the user clicked row 0.
    If grdTrib.MouseRow > 0 Then Exit Sub

    ' See if this is the same column.
    If Val(grdTrib.MouseCol) = Val(m_SortColumn) Then
        ' This is the current sort column.
        ' Change the sort order and the column title.
        m_SortAscending = Not m_SortAscending
        If m_SortAscending Then
            grdTrib.TextMatrix(0, m_SortColumn) = _
                 grdTrib.TextMatrix(0, m_SortColumn)
        Else
            grdTrib.TextMatrix(0, m_SortColumn) = _
                 grdTrib.TextMatrix(0, m_SortColumn)
        End If
    Else
        ' This is a new sort column.
        ' Restore the previous sorting column's name.
        If m_SortColumn >= 0 Then
            grdTrib.TextMatrix(0, m_SortColumn) = _
                grdTrib.TextMatrix(0, m_SortColumn)
        End If

        ' Save the new sort column.
        m_SortColumn = grdTrib.MouseCol

        ' Sort using the new column.
        m_SortAscending = True
        grdTrib.TextMatrix(0, m_SortColumn) = _
             grdTrib.TextMatrix(0, m_SortColumn)
    End If

    grdTrib.Row = 1
    grdTrib.RowSel = grdTrib.Rows - 1
    grdTrib.col = m_SortColumn

    If m_SortAscending Then
        Select Case m_SortColumn
            Case 0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
                grdTrib.Sort = flexSortNumericAscending
            Case Else
                grdTrib.Sort = flexSortStringAscending
        End Select
    Else
        Select Case m_SortColumn
            Case 0, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13
                grdTrib.Sort = flexSortNumericDescending
            Case Else
                grdTrib.Sort = flexSortStringDescending
        End Select
    End If

End Sub

Private Sub grdTrib_RowColChange()
Le
End Sub

Private Sub FormHagana()

evNew = 2
evEdit = 3
evDel = 4
evEsp = 11

If InStr(1, sRet, Format(evNew, "000"), vbBinaryCompare) > 0 Then bNew = True
If InStr(1, sRet, Format(evEdit, "000"), vbBinaryCompare) > 0 Then bEdit = True
If InStr(1, sRet, Format(evDel, "000"), vbBinaryCompare) > 0 Then bDel = True
If InStr(1, sRet, Format(evEsp, "000"), vbBinaryCompare) > 0 Then bEsp = True

If Not bNew Then cmdNovo.Enabled = False
If Not bEdit Then cmdAlterar.Enabled = False
If Not bDel Then cmdExcluir.Enabled = False
If Not bEsp Then
    txtFicha(1).Locked = True
    txtFicha(2).Locked = True
    txtFicha(3).Locked = True
    txtFicha(4).Locked = True
    txtFicha(5).Locked = True
    txtFicha(6).Locked = True
    txtFicha(7).Locked = True
    txtFicha(8).Locked = True
Else
    txtFicha(1).Locked = False
    txtFicha(2).Locked = False
    txtFicha(3).Locked = False
    txtFicha(4).Locked = False
    txtFicha(5).Locked = False
    txtFicha(6).Locked = False
    txtFicha(7).Locked = False
    txtFicha(8).Locked = False
End If

End Sub

Private Sub txtFicha_KeyPress(Index As Integer, KeyAscii As Integer)
Tweak txtFicha(Index), KeyAscii, IntegerPositive
End Sub
Private Sub MudaFicha()
Dim Sql As String, RdoAux As rdoResultset, nPos As Long, nTot As Long, RdoAux2 As rdoResultset, x As Integer
Dim nCodReduz As Long, nFicha As Long, sNatureza As String, nFichaNova As Integer

If NomeDeLogin <> "SCHWARTZ" Then Exit Sub
'Exit Sub
nPos = 1

Sql = "SELECT * from tributo order by codtributo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount
    Do Until .EOF
        nFicha = !Ficha
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set ficha=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
On Error Resume Next
        nFicha = !FichaJrMulta
        If nFicha = 148 Then nFicha = 156
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichajrmulta=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaDivida
        If nFicha = 148 Then nFicha = 156
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichadivida=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaDaJrMul
        If nFicha = 148 Then nFicha = 156
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichadajrmul=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaDaEnca
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichadaenca=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaAjuiza
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichaajuiza=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaAjJrMul
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichaajjrmul=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nFicha = !FichaAjEnca
        Sql = "select ficha from fichacontabil where ficha_old=" & nFicha
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nFichaNova = RdoAux2!Ficha
        RdoAux2.Close
        Sql = "update tributo set fichaajenca=" & nFichaNova & " where codtributo=" & !CodTributo
        cn.Execute Sql, rdExecDirect
        
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With
    
    
proximo:

MsgBox "fim"

End Sub

Private Sub AtualizaFicha()
Dim Sql As String, RdoAux As rdoResultset, sVinculo_cod As String, sVinculo_nome As String
Dim sNatureza_cod As String, sNatureza_nome As String, sVinculo As String, sNatureza As String, nPos As Long

Sql = "select * from fichacontabil order by seq"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        sNatureza = !natureza_old
        nPos = InStr(1, sNatureza, "-", vbBinaryCompare)
        sNatureza_cod = Left(sNatureza, nPos - 2)
        sNatureza_nome = Mid(sNatureza, nPos + 2, Len(sNatureza) - nPos)
        
        sVinculo = !vinculo_old
        nPos = InStr(1, sVinculo, "-", vbBinaryCompare)
        sVinculo_cod = Left(sVinculo, nPos - 2)
        sVinculo_nome = Mid(sVinculo, nPos + 2, Len(sVinculo) - nPos)
        
        Sql = "update fichacontabil set natureza='" & sNatureza_cod & "',natureza_desc='" & sNatureza_nome & "',vinculo='" & sVinculo_cod & "',vinculo_desc='" & sVinculo_nome & "' where seq=" & !Seq
        cn.Execute Sql, rdExecDirect
        
       .MoveNext
    Loop
   .Close
End With

MsgBox "fim"


End Sub
