VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmPeriodoSN 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Períodos S.Nac."
   ClientHeight    =   3960
   ClientLeft      =   11265
   ClientTop       =   3900
   ClientWidth     =   4185
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3960
   ScaleWidth      =   4185
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1875
      TabIndex        =   1
      Top             =   2745
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      MouseIcon       =   "frmPeriodoSN.frx":0000
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
   Begin prjChameleon.chameleonButton cmdGravar 
      Height          =   315
      Left            =   2355
      TabIndex        =   7
      ToolTipText     =   "Gravar os Dados"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   14
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":001C
      PICN            =   "frmPeriodoSN.frx":0038
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
      Left            =   1815
      TabIndex        =   6
      ToolTipText     =   "Cancelar Edição"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":03DD
      PICN            =   "frmPeriodoSN.frx":03F9
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
      Left            =   225
      TabIndex        =   3
      ToolTipText     =   "Novo Registro"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":0553
      PICN            =   "frmPeriodoSN.frx":056F
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
      Left            =   765
      TabIndex        =   4
      ToolTipText     =   "Editar Registro"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":06C9
      PICN            =   "frmPeriodoSN.frx":06E5
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
      Left            =   1305
      TabIndex        =   5
      ToolTipText     =   "Excluir Registro"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   ""
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
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":083F
      PICN            =   "frmPeriodoSN.frx":085B
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
      Height          =   2490
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   4065
      _ExtentX        =   7170
      _ExtentY        =   4392
      _Version        =   393216
      Rows            =   1
      Cols            =   3
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorSel    =   192
      ForeColorSel    =   16777215
      GridColor       =   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      FormatString    =   "^Data Início        |^Data Final          |Usuário           "
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   1875
      TabIndex        =   2
      Top             =   3135
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      MouseIcon       =   "frmPeriodoSN.frx":08FD
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
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   270
      Left            =   2340
      TabIndex        =   8
      ToolTipText     =   "Sair da Tela"
      Top             =   3555
      Width           =   495
      _ExtentX        =   873
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
      MCOL            =   13026246
      MPTR            =   1
      MICON           =   "frmPeriodoSN.frx":0919
      PICN            =   "frmPeriodoSN.frx":0935
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final........:"
      Height          =   225
      Left            =   585
      TabIndex        =   10
      Top             =   3165
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial......:"
      Height          =   225
      Left            =   585
      TabIndex        =   9
      Top             =   2775
      Width           =   1215
   End
End
Attribute VB_Name = "frmPeriodoSN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Evento As String

Private Sub cmdAlterar_Click()
If Not IsDate(mskDataIni.Text) Then
    MsgBox "Selecione um período.", vbExclamation, "Atenção"
    Exit Sub
End If
Eventos "INCLUIR"
Evento = "Alterar"
mskDataIni.SetFocus
End Sub

Private Sub cmdCancel_Click()
Eventos "INICIAR"
End Sub

Private Sub cmdExcluir_Click()
If Not IsDate(mskDataIni.Text) Then
    MsgBox "Selecione um período.", vbExclamation, "Atenção"
    Exit Sub
End If

If MsgBox("Excluir este período?", vbQuestion + vbYesNo, "Confirmação") = vbNo Then Exit Sub

Sql = "DELETE FROM PERIODOSN WHERE CODIGO=" & Val(frmCadMob.txtCodEmpresa.Text) & " AND DATAINI='" & Format(mskDataIni.Text, "mm/dd/yyyy") & "'"
cn.Execute Sql, rdExecDirect

CarregaLista

End Sub

Private Sub cmdGravar_Click()
Dim x As Integer

If Not IsDate(mskDataIni.Text) Then
    MsgBox "Digite a Data inicial.", vbCritical, "Atenção"
    Exit Sub
End If

If IsDate(mskDataFim.Text) Then
    If CDate(mskDataIni.Text) >= CDate(mskDataFim.Text) Then
        MsgBox "Data inicial tem que ser menor que a data final.", vbCritical, "Atenção"
        Exit Sub
    End If
End If

With grdMain
        
    Dim nConta As Integer
    nConta = 0
    If mskDataFim.ClipText = "" Then
        For x = 1 To .Rows - 1
            If Not IsDate(.TextMatrix(x, 1)) Then
                nConta = nConta + 1
            End If
        Next
    End If
    
    If nConta > 1 Then
        MsgBox "Apenas um dos períodos pode permanecer sem data final.", vbCritical, "Atenção"
        Exit Sub
    End If
    
    If Evento = "Novo" And .Rows > 2 Then
        If CDate(mskDataIni.Text) < CDate(.TextMatrix(.Rows - 1, 0)) Then
            MsgBox "Data inicial tem que ser maior ou igual a data final do último período.", vbCritical, "Atenção"
            Exit Sub
        End If
    End If

    If Evento = "Alterar" And .Rows > 2 Then
        If .Row > 1 Then
            If CDate(mskDataIni.Text) < CDate(.TextMatrix(.Row - 1, 1)) Then
                MsgBox "Data inicial tem que ser maior ou igual a data final do último período.", vbCritical, "Atenção"
                Exit Sub
            End If
        End If
        If .Row < .Rows - 1 And mskDataFim.ClipText = "" Then
            MsgBox "Apenas o último período pode permanecer sem data final.", vbCritical, "Atenção"
            Exit Sub
        End If
    End If

 '   If .Rows > 1 And Evento = "Novo" Then
 '       If Not IsDate(.TextMatrix(.Rows - 1, 1)) Then
 '           MsgBox "Apenas o último período pode permanecer sem data final.", vbCritical, "Atenção"
 '           Exit Sub
 '       End If
 '   End If

End With

If Evento = "Novo" Then
    grdMain.AddItem mskDataIni.Text & Chr(9) & mskDataFim.Text
Else
    grdMain.TextMatrix(grdMain.Row, 0) = mskDataIni.Text
    grdMain.TextMatrix(grdMain.Row, 1) = mskDataFim.Text
End If

Sql = "DELETE FROM PERIODOSN WHERE CODIGO=" & Val(frmCadMob.txtCodEmpresa.Text)
cn.Execute Sql, rdExecDirect

With grdMain
    For x = 1 To .Rows - 1
        Sql = "INSERT PERIODOSN(CODIGO,DATAINI,DATAFIM,USUARIO) VALUES(" & Val(frmCadMob.txtCodEmpresa.Text) & ",'" & Format(.TextMatrix(x, 0), "mm/dd/yyyy") & "',"
        If IsDate(.TextMatrix(x, 1)) Then
            Sql = Sql & "'" & Format(.TextMatrix(x, 1), "mm/dd/yyyy") & "','" & NomeDeLogin & "')"
        Else
            Sql = Sql & "Null" & ",'" & NomeDeLogin & "')"
        End If
        cn.Execute Sql, rdExecDirect
    Next
End With
CarregaLista
Eventos "INICIAR"
End Sub

Private Sub cmdNovo_Click()
Eventos "INCLUIR"
Evento = "Novo"
Limpa
mskDataIni.SetFocus
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
Me.Top = Me.Top + 800
Me.BackColor = Kde
Eventos "INICIAR"
If NomeDeLogin = "LUIZH" Or NomeDeLogin = "LEANDRO" Or NomeDeLogin = "SCHWARTZ" Then
    cmdNovo.Enabled = True
    cmdAlterar.Enabled = True
    cmdExcluir.Enabled = True
    cmdGravar.Enabled = True
    cmdCancel.Enabled = True
End If
CarregaLista
End Sub

Private Sub Eventos(Tipo As String)

If Tipo = "INICIAR" Then
   cmdNovo.Visible = True
   cmdAlterar.Visible = True
   cmdExcluir.Visible = True
   cmdSair.Visible = True
   cmdGravar.Visible = False
   cmdCancel.Visible = False
   mskDataIni.Locked = True
   mskDataFim.Locked = True
   mskDataIni.BackColor = Kde
   mskDataFim.BackColor = Kde
ElseIf Tipo = "INCLUIR" Then
   cmdNovo.Visible = False
   cmdAlterar.Visible = False
   cmdExcluir.Visible = False
   cmdSair.Visible = False
   cmdGravar.Visible = True
   cmdCancel.Visible = True
   mskDataIni.Locked = False
   mskDataFim.Locked = False
   mskDataIni.BackColor = Branco
   mskDataFim.BackColor = Branco
End If

End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String
grdMain.Rows = 1
Limpa
Sql = "SELECT * FROM optante_simples WHERE CODIGO=" & Val(frmCadMob.txtCodEmpresa.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        grdMain.AddItem Format(!Data_Inicio, "dd/mm/yyyy") & Chr(9) & IIf(IsNull(!Data_Final), "", Format(!Data_Final, "dd/mm/yyyy"))
       .MoveNext
    Loop
   .Close
End With
If grdMain.Rows > 1 Then grdMain_RowColChange

End Sub

Private Sub Limpa()
LimpaMascara mskDataIni
LimpaMascara mskDataFim
End Sub

Private Sub grdMain_Click()
If grdMain.Rows = 1 Then Exit Sub
If grdMain.Row = 0 Then Exit Sub
grdMain_RowColChange
End Sub

Private Sub grdMain_RowColChange()
Limpa
If grdMain.Rows = 1 Then Exit Sub
If grdMain.Row = 0 Then Exit Sub
mskDataIni.Text = grdMain.TextMatrix(grdMain.Row, 0)
mskDataFim.Text = grdMain.TextMatrix(grdMain.Row, 1)
End Sub
