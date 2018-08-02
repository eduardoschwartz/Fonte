VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmAutorizaNota 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Autorização de Talão de Nota Fiscal"
   ClientHeight    =   5235
   ClientLeft      =   3225
   ClientTop       =   2835
   ClientWidth     =   8130
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5235
   ScaleWidth      =   8130
   Begin VB.TextBox txtRazao 
      Appearance      =   0  'Flat
      BackColor       =   &H00EEEEEE&
      ForeColor       =   &H00C00000&
      Height          =   285
      Left            =   60
      Locked          =   -1  'True
      MaxLength       =   100
      TabIndex        =   23
      Top             =   420
      Width           =   8010
   End
   Begin VB.TextBox txtCod 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   1770
      MaxLength       =   6
      TabIndex        =   0
      Top             =   90
      Width           =   990
   End
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   6930
      TabIndex        =   12
      ToolTipText     =   "Sair da tela"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAutorizaNota.frx":0000
      PICN            =   "frmAutorizaNota.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame PnNota 
      BackColor       =   &H00EEEEEE&
      Height          =   945
      Left            =   90
      TabIndex        =   15
      Top             =   3750
      Width           =   7965
      Begin VB.TextBox txtSerie 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   2700
         MaxLength       =   10
         TabIndex        =   3
         Top             =   150
         Width           =   990
      End
      Begin VB.TextBox txtInicio 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   4740
         MaxLength       =   10
         TabIndex        =   4
         Top             =   150
         Width           =   990
      End
      Begin VB.TextBox txtFinal 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6630
         MaxLength       =   10
         TabIndex        =   5
         Top             =   150
         Width           =   990
      End
      Begin VB.TextBox txtAut 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   1260
         Locked          =   -1  'True
         MaxLength       =   25
         TabIndex        =   6
         Top             =   510
         Width           =   1920
      End
      Begin VB.CheckBox chkCancel 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Caption         =   "Cancelado.:"
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   6420
         TabIndex        =   8
         Top             =   540
         Width           =   1215
      End
      Begin esMaskEdit.esMaskedEdit mskDataNota 
         Height          =   285
         Left            =   4740
         TabIndex        =   7
         Top             =   510
         Width           =   990
         _ExtentX        =   1746
         _ExtentY        =   503
         MouseIcon       =   "frmAutorizaNota.frx":0176
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
      Begin prjChameleon.chameleonButton cmdManual 
         Height          =   270
         Left            =   3195
         TabIndex        =   24
         ToolTipText     =   "Alternar entre Autorização manual e automática"
         Top             =   540
         Width           =   330
         _ExtentX        =   582
         _ExtentY        =   476
         BTYPE           =   5
         TX              =   "..."
         ENAB            =   -1  'True
         BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmAutorizaNota.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   -1  'True
         VALUE           =   0   'False
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Sequência....:"
         Height          =   225
         Left            =   180
         TabIndex        =   21
         Top             =   210
         Width           =   1065
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Série......:"
         Height          =   225
         Left            =   1950
         TabIndex        =   20
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Inicial....:"
         Height          =   225
         Left            =   3810
         TabIndex        =   19
         Top             =   210
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº Final..:"
         Height          =   225
         Left            =   5850
         TabIndex        =   18
         Top             =   210
         Width           =   795
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Autorização...:"
         Height          =   225
         Left            =   150
         TabIndex        =   17
         Top             =   570
         Width           =   1005
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "Data aut....:"
         Height          =   225
         Left            =   3810
         TabIndex        =   16
         Top             =   570
         Width           =   885
      End
      Begin VB.Label lblSeqNota 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1260
         TabIndex        =   2
         Top             =   210
         Width           =   555
      End
   End
   Begin MSFlexGridLib.MSFlexGrid grdNF 
      Height          =   2955
      Left            =   60
      TabIndex        =   1
      Top             =   780
      Width           =   8025
      _ExtentX        =   14155
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   1
      Cols            =   8
      FixedCols       =   0
      BackColorSel    =   192
      ForeColorSel    =   16777215
      BackColorBkg    =   15658734
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "^Seq     |<Série    |^Núm.Inicial     |^Núm.Final     |<Nº de Autorização      |^Cancel |^Data Autoriz. |<Usuario              "
   End
   Begin prjChameleon.chameleonButton cmdExcluir 
      Height          =   315
      Left            =   5760
      TabIndex        =   11
      ToolTipText     =   "Excluir Nota Fiscal"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAutorizaNota.frx":01AE
      PICN            =   "frmAutorizaNota.frx":01CA
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
      Left            =   4590
      TabIndex        =   10
      ToolTipText     =   "Alterar Nota Fiscal"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Alterar"
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
      MICON           =   "frmAutorizaNota.frx":026C
      PICN            =   "frmAutorizaNota.frx":0288
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
      Left            =   3420
      TabIndex        =   9
      ToolTipText     =   "Nova Nota Fiscal"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAutorizaNota.frx":03E2
      PICN            =   "frmAutorizaNota.frx":03FE
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
      Left            =   5760
      TabIndex        =   13
      ToolTipText     =   "Gravar Nota"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAutorizaNota.frx":0558
      PICN            =   "frmAutorizaNota.frx":0574
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancelar 
      Height          =   315
      Left            =   6930
      TabIndex        =   14
      ToolTipText     =   "Cancelar Edição"
      Top             =   4830
      Width           =   1095
      _ExtentX        =   1931
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
      MICON           =   "frmAutorizaNota.frx":0919
      PICN            =   "frmAutorizaNota.frx":0935
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
      BackStyle       =   0  'Transparent
      Caption         =   "Código da Empresa..:"
      Height          =   225
      Left            =   180
      TabIndex        =   22
      Top             =   150
      Width           =   1575
   End
End
Attribute VB_Name = "frmAutorizaNota"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String, Sql As String, RdoAux As rdoResultset, bISSEletro As Boolean

Private Sub cmdAlterar_Click()
If Val(txtCod.text) = 0 Then
    MsgBox "Selecione uma empresa.", vbExclamation, "Atenção"
    Exit Sub
End If


If Val(lblSeqNota.Caption) = 0 Then Exit Sub
Evento = "Alterar"
Eventos "INCLUIR"

End Sub

Private Sub cmdCancelar_Click()

Eventos "INICIAR"
Evento = ""
If grdNF.Rows > 2 Then
    grdNF.Row = 1
    grdNF_RowColChange
End If

End Sub

Private Sub cmdExcluir_Click()
If Val(txtCod.text) = 0 Then
    MsgBox "Selecione uma empresa.", vbExclamation, "Atenção"
    Exit Sub
End If


If txtSerie.text = "" Then Exit Sub
End Sub

Private Sub cmdGravar_Click()
Dim sData As String

If mskDataNota.ClipText = "" Then
    MsgBox "Data inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If IsDate(mskDataNota.text) Then
    sData = mskDataNota.text
Else
    If mskDataNota.ClipText <> "" Then
        MsgBox "Data inválida.", vbExclamation, "Atenção"
        Exit Sub
    End If
    sData = ""
End If

If CDate(mskDataNota.text) < CDate("01/01/1970") Or CDate(mskDataNota.text) > CDate("01/01/2020") Then
    MsgBox "Data inválida.", vbExclamation, "Atenção"
    Exit Sub
End If

If bISSEletro And txtAut.text = "" Then
    MsgBox "Digite o processo.", vbExclamation, "Atenção"
    Exit Sub
End If

If Not bISSEletro Then
'    If Val(txtFinal.text) <= Val(txtInicio.text) Then
'        MsgBox "Nº Final tem que ser maior que nº inicial.", vbExclamation, "Atenção"
'        Exit Sub
'    End If
    
    If txtSerie.text = "" Or txtInicio.text = "" Or txtFinal.text = "" Then
        MsgBox "Digite nº de série, nº inicial e nº final.", vbExclamation, "Atenção"
        Exit Sub
    End If
End If

If Evento = "Novo" Then
    If Not bISSEletro Then
        If cmdManual.Value = False Then
            Sql = "SELECT VALPARAM FROM PARAMETROS WHERE NOMEPARAM='SEQNOT'"
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                txtAut.text = Format(!VALPARAM + 1, "000000")
               .Close
            End With
                        
            Sql = "UPDATE PARAMETROS SET VALPARAM=" & txtAut.text & " WHERE NOMEPARAM='SEQNOT'"
            cn.Execute Sql, rdExecDirect
        Else
            If Trim(txtAut.text) = "" Then
                MsgBox "Digite a autenticação.", vbExclamation, "Atenção"
                Exit Sub
            End If
        End If
        txtInicio.text = Format(txtInicio.text, "00000000")
        txtFinal.text = Format(txtFinal.text, "00000000")
    End If
    grdNF.AddItem lblSeqNota.Caption & Chr(9) & txtSerie.text & Chr(9) & txtInicio.text & Chr(9) & txtFinal.text & Chr(9) & txtAut.text & Chr(9) & IIf(chkCancel.Value = vbChecked, "Sim", "") & Chr(9) & sData & Chr(9) & NomeDeLogin
    Sql = "INSERT MOBILIARIONF(CODIGOMOB,SEQ,SERIE,NUMINI,NUMFIM,NUMAUT,DATAAUT,CANCEL,USUARIO) VALUES(" & Val(txtCod.text) & "," & lblSeqNota.Caption & ",'"
    Sql = Sql & Mask(txtSerie.text) & "'," & Val(txtInicio.text) & "," & Val(txtFinal.text) & ",'" & Mask(txtAut.text) & "','" & Format(mskDataNota.text, "mm/dd/yyyy") & "',"
    Sql = Sql & IIf(chkCancel.Value = vbChecked, 1, 0) & ",'" & NomeDeLogin & "')"
    cn.Execute Sql, rdExecDirect
Else
    With grdNF
        .TextMatrix(.Row, 1) = txtSerie.text
        .TextMatrix(.Row, 2) = Format(txtInicio.text, "00000000")
        .TextMatrix(.Row, 3) = Format(txtFinal.text, "00000000")
        .TextMatrix(.Row, 4) = txtAut.text
        .TextMatrix(.Row, 5) = IIf(chkCancel.Value = vbChecked, "Sim", "")
        .TextMatrix(.Row, 6) = sData
        .TextMatrix(.Row, 7) = NomeDeLogin
        Sql = "UPDATE MOBILIARIONF SET SERIE='" & Mask(txtSerie.text) & "',NUMINI=" & Format(txtInicio.text, "00000000") & ",NUMFIM=" & Format(txtFinal.text, "00000000") & ","
        Sql = Sql & "NUMAUT='" & Mask(txtAut.text) & "',DATAAUT='" & Format(mskDataNota.text, "mm/dd/yyyy") & "',CANCEL=" & IIf(chkCancel.Value = vbChecked, 1, 0) & ","
        Sql = Sql & "USUARIO='" & NomeDeLogin & "' WHERE CODIGOMOB=" & Val(txtCod.text) & " AND SEQ=" & lblSeqNota.Caption
        cn.Execute Sql, rdExecDirect
    End With
End If

Eventos "INICIAR"
End Sub

Private Sub cmdManual_Click()
If Not bISSEletro Then
    If cmdManual.Value = False Then
        txtAut.text = "Definição Automática"
        txtAut.Locked = True
    Else
        txtAut.text = ""
        txtAut.Locked = False
    End If
End If
End Sub

Private Sub cmdNovo_Click()
If Val(txtCod.text) = 0 Then
    MsgBox "Selecione uma empresa.", vbExclamation, "Atenção"
    Exit Sub
End If

bISSEletro = False
If MsgBox("Autorizar para ISS Eletrônico através de Processo?", vbQuestion + vbYesNo + vbDefaultButton2, "Tipo de Autorização") = vbYes Then
    bISSEletro = True
End If

Limpa
Evento = "Novo"
Eventos "INCLUIR"


If grdNF.Rows = 1 Then
    lblSeqNota.Caption = "0001"
Else
    lblSeqNota.Caption = Format(Val(grdNF.TextMatrix(grdNF.Rows - 1, 0)) + 1, "0000")
End If

If Not bISSEletro Then
    txtAut.text = "Definição Automática"
    txtSerie.SetFocus
Else
    txtAut.SetFocus
End If

End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
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
   cmdCancelar.Visible = False
   For Each Ct In frmAutorizaNota
       If TypeOf Ct Is TextBox Or TypeOf Ct Is CheckBox Then
         Ct.BackColor = Kde
         Ct.Enabled = False
       End If
   Next
   mskDataNota.Enabled = False
   mskDataNota.BackColor = Kde
   txtCod.Enabled = True
   txtCod.BackColor = Branco
   txtRazao.Enabled = True
   grdNF.Enabled = True
   txtAut.Locked = True
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancelar.Visible = True
   grdNF.Enabled = False
   If Not bISSEletro Then
      For Each Ct In frmAutorizaNota
          If TypeOf Ct Is TextBox Or TypeOf Ct Is CheckBox Then
             Ct.BackColor = vbWhite
             Ct.Enabled = True
          End If
      Next
      If Evento = "Novo" Then
         txtAut.Locked = True
      Else
         txtAut.Locked = False
      End If
   Else
      txtAut.Locked = False
      txtAut.Enabled = True
      txtAut.BackColor = vbWhite
      chkCancel.Enabled = True
   End If
   mskDataNota.Enabled = True
   mskDataNota.BackColor = Branco
End If

End Sub

Private Sub Limpa()
txtSerie.text = ""
txtInicio.text = ""
txtFinal.text = ""
txtAut.text = ""
lblSeqNota.Caption = ""
LimpaMascara mskDataNota
chkCancel.Value = vbUnchecked
cmdManual.Value = False
End Sub

Private Sub grdNF_Click()
grdNF_RowColChange
End Sub

Private Sub grdNF_RowColChange()
If grdNF.Rows = 1 Then Exit Sub
If grdNF.Row = 0 Then Exit Sub
With grdNF
    lblSeqNota.Caption = .TextMatrix(.Row, 0)
    txtSerie.text = .TextMatrix(.Row, 1)
    txtInicio.text = .TextMatrix(.Row, 2)
    txtFinal.text = .TextMatrix(.Row, 3)
    txtAut.text = .TextMatrix(.Row, 4)
    chkCancel.Value = IIf(.TextMatrix(.Row, 5) = "", vbUnchecked, vbChecked)
    mskDataNota.text = .TextMatrix(.Row, 6)
End With

End Sub

Private Sub mskDataNota_GotFocus()
mskDataNota.SetFocus
End Sub

Private Sub txtCod_Change()
If txtRazao.text <> "" Then txtRazao.text = ""
If grdNF.Rows > 1 Then grdNF.Rows = 1
Limpa
End Sub

Private Sub txtCod_KeyPress(KeyAscii As Integer)

If KeyAscii = vbKeyReturn Then
    If Val(txtCod.text) = 0 Then Exit Sub
    Sql = "SELECT RAZAOSOCIAL FROM MOBILIARIO WHERE CODIGOMOB=" & Val(txtCod.text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
    With RdoAux
        If .RowCount > 0 Then
            txtRazao.text = !RAZAOSOCIAL
            CarregaNF
            grdNF_RowColChange
        Else
            txtRazao.text = ""
            MsgBox "Empresa não cadastrada.", vbCritical, "Atenção"
        End If
       .Close
    End With
Else
    Tweak txtCod, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtCod_LostFocus()
If Val(txtCod.text) > 0 Then txtCod_KeyPress (vbKeyReturn)
End Sub

Private Sub txtFinal_KeyPress(KeyAscii As Integer)
Tweak txtFinal, KeyAscii, IntegerPositive
End Sub

Private Sub txtInicio_KeyPress(KeyAscii As Integer)
Tweak txtInicio, KeyAscii, IntegerPositive
End Sub

Private Sub CarregaNF()

grdNF.Rows = 1
Sql = "SELECT SEQ,SERIE,NUMINI,NUMFIM,NUMAUT,DATAAUT,CANCEL,USUARIO FROM MOBILIARIONF  "
Sql = Sql & "Where CODIGOMOB = " & Val(txtCod.text) & " ORDER BY SEQ "
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdNF.AddItem Format(!Seq, "0000") & Chr(9) & SubNull(!Serie) & Chr(9) & Format(!NUMINI, "00000000") & _
        Chr(9) & Format(!NUMFIM, "00000000") & Chr(9) & !NUMAUT & Chr(9) & IIf(!Cancel, "Sim", "") & Chr(9) & IIf(IsNull(!DATAAUT), "", Format(!DATAAUT, "dd/mm/yyyy")) & Chr(9) & SubNull(!USUARIO)
       .MoveNext
    Loop
   .Close
End With

End Sub
