VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSenhaPre 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Pré-Atendimento"
   ClientHeight    =   4335
   ClientLeft      =   8190
   ClientTop       =   4860
   ClientWidth     =   3255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4335
   ScaleWidth      =   3255
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   5
      Left            =   1845
      TabIndex        =   21
      Text            =   "0"
      Top             =   6570
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   4
      Left            =   1845
      TabIndex        =   18
      Text            =   "0"
      Top             =   6255
      Width           =   690
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2655
      Top             =   3825
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   3
      Left            =   1845
      TabIndex        =   13
      Text            =   "0"
      Top             =   5940
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   2
      Left            =   1845
      TabIndex        =   11
      Text            =   "0"
      Top             =   5625
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   1
      Left            =   1845
      TabIndex        =   9
      Text            =   "0"
      Top             =   5310
      Width           =   690
   End
   Begin VB.TextBox txtSenha 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   285
      Index           =   0
      Left            =   1845
      TabIndex        =   6
      Text            =   "0"
      Top             =   4995
      Width           =   690
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   4365
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   1260
      Width           =   645
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   0
      Left            =   90
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   180
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREFEITURA"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   1
      Left            =   90
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   765
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "2ª VIA IPTU"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":031A
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   2
      Left            =   90
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1350
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREFERENCIAL"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":0634
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   3
      Left            =   90
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1935
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "SENHA PAT"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":094E
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdConfig 
      Height          =   330
      Left            =   855
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3825
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      BTYPE           =   14
      TX              =   "Configuração"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   12640511
      FCOLO           =   12640511
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":0C68
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   -1  'True
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   4
      Left            =   90
      TabIndex        =   17
      TabStop         =   0   'False
      Top             =   2520
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "REFIS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":0F82
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   5
      Left            =   90
      TabIndex        =   20
      TabStop         =   0   'False
      Top             =   3105
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREF. REFIS"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   4210752
      BCOLO           =   4210752
      FCOL            =   8454143
      FCOLO           =   8454143
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaPre.frx":129C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Mov.Econom.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   6
      Left            =   495
      TabIndex        =   22
      Top             =   6570
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Refis..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   5
      Left            =   495
      TabIndex        =   19
      Top             =   6255
      Width           =   1230
   End
   Begin VB.Label lblBanda 
      Height          =   240
      Left            =   4230
      TabIndex        =   16
      Top             =   765
      Width           =   1410
   End
   Begin VB.Label lblSenha 
      Height          =   240
      Left            =   4230
      TabIndex        =   15
      Top             =   405
      Width           =   1410
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Senha PAT..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   4
      Left            =   495
      TabIndex        =   14
      Top             =   5940
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Preferêncial.:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   3
      Left            =   495
      TabIndex        =   12
      Top             =   5625
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prefeitura 2..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   2
      Left            =   495
      TabIndex        =   10
      Top             =   5310
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Prefeitura 1..:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   195
      Index           =   1
      Left            =   495
      TabIndex        =   8
      Top             =   4995
      Width           =   1230
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Próxima Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFC0&
      Height          =   195
      Index           =   0
      Left            =   810
      TabIndex        =   7
      Top             =   4635
      Width           =   1590
   End
End
Attribute VB_Name = "frmSenhaPre"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdConfig_Click()
If cmdConfig.value = False Then
    Me.Height = 4905
Else
    Me.Height = 7545
End If
End Sub

Private Sub cmdSenha_Click(Index As Integer)
Dim RdoAux As rdoResultset, Sql As String, nSenha As Integer, nBanda As Integer
Me.Enabled = False

Ocupado
nBanda = Index + 1

If nBanda = 1 Then
    nSenha = Val(txtSenha(0).Text)
ElseIf nBanda = 2 Then
    nSenha = Val(txtSenha(1).Text)
ElseIf nBanda = 3 Then
    nSenha = Val(txtSenha(2).Text)
ElseIf nBanda = 4 Then
    nSenha = Val(txtSenha(3).Text)
ElseIf nBanda = 5 Then
    nSenha = Val(txtSenha(4).Text)
ElseIf nBanda = 6 Then
    nSenha = Val(txtSenha(5).Text)
End If

Sql = "SELECT * FROM SSPAC WHERE DATAENTRADA='" & Format(Now, "mm/dd/yyyy") & "' AND SENHA=" & nSenha
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
If RdoAux.RowCount = 0 Then
    Sql = "INSERT SSPAC(DATAENTRADA,HORAENTRADA,SENHA,BANDA,MONITOR) VALUES('" & Format(Now, "mm/dd/yyyy") & "','"
    Sql = Sql & Format(Now, "hh:mm:ss") & "'," & nSenha & "," & nBanda & ",0)"
    cn.Execute Sql, rdExecDirect
    RdoAux.Close
End If

lblSenha.Caption = Format(nSenha, "000")
lblBanda.Caption = Left(cmdSenha(Index).Caption, InStr(1, cmdSenha(Index).Caption, "(") - 2)

txtSenha(Index).Text = Val(txtSenha(Index).Text) + 1
On Error Resume Next
Text1.SetFocus

On Error GoTo Erro
Open "Lpt1" For Output As #1
Print #1, Spc(2); "========================================="
Print #1, Spc(6); "PREFEITURA MUNICIPAL DE JABOTICABAL"
Print #1, Spc(2); "Sistema Pratico de Atendimento ao Cidadao"
Print #1, Spc(2); "========================================="
Print #1, Spc(8); "Data:" & Format(Now, "dd/mm/yyyy"); Spc(2); "Hora:" & Format(Now, "hh:mm:ss")
Print #1, Spc(12); Chr(27) & Chr(69) & Chr(27) & Chr(14) + "Senha:" & Format(nSenha, "000")
Print #1, Spc(12); Chr(27) & Chr(70) & Chr(27) & Chr(14) + lblBanda.Caption
Print #1, Chr(20)
Print #1, Spc(2); "POR FAVOR AGUARDE."
Print #1, Chr(10) & Chr(13)
Print #1, Chr(10)
Print #1, Chr(10)
Print #1, Chr(10)
Close #1
Liberado

Me.Enabled = True
Exit Sub
Erro:
Liberado
Me.Enabled = True
MsgBox "Impressora não conectada.", vbCritical, "Atenção"

End Sub

Private Sub cmdSenha_MouseOut(Index As Integer)
On Error Resume Next
Text1.SetFocus
End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset, nSenha As Integer
Ocupado

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 1
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 1
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(0).Text = nSenha

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 2
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 400
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(1).Text = nSenha

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 3
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 500
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(2).Text = nSenha

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 4
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 600
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(3).Text = nSenha

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 5
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 700
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(4).Text = nSenha

Sql = "SELECT MAX(SENHA) AS MAXIMO FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & 6
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurRowVer)
With RdoAux
    If IsNull(!maximo) Then
        nSenha = 900
    Else
        nSenha = !maximo + 1
    End If
   .Close
End With
txtSenha(5).Text = nSenha


Le
Liberado
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, dData As Date, x As Integer
Dim aCount(5) As Integer
dData = Now
cmdSenha(0).Caption = "PREFEITURA (0)"
cmdSenha(1).Caption = "2ª VIA IPTU (0)"
cmdSenha(2).Caption = "PREFERENCIAL (0)"
cmdSenha(3).Caption = "SENHA PAT (0)"
cmdSenha(4).Caption = "REFIS (0)"
cmdSenha(5).Caption = "PREF. REFIS (0)"

Sql = "SELECT * FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(dData) & " AND MONTH(DATAENTRADA)=" & Month(dData) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(dData) & " AND DATACHAMADA IS NULL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    
    Do Until .EOF
        aCount(!BANDA - 1) = (aCount(!BANDA - 1)) + 1
       Select Case !BANDA - 1
            Case 0
                cmdSenha(!BANDA - 1).Caption = "PREFEITURA (" & aCount(!BANDA - 1) & ")"
            Case 1
                cmdSenha(!BANDA - 1).Caption = "2ª VIA IPTU (" & aCount(!BANDA - 1) & ")"
            Case 2
                cmdSenha(!BANDA - 1).Caption = "PREFERENCIAL (" & aCount(!BANDA - 1) & ")"
            Case 3
                cmdSenha(!BANDA - 1).Caption = "SENHA PAT (" & aCount(!BANDA - 1) & ")"
            Case 4
                cmdSenha(!BANDA - 1).Caption = "REFIS (" & aCount(!BANDA - 1) & ")"
            Case 5
                cmdSenha(!BANDA - 1).Caption = "PREF. REFIS (" & aCount(!BANDA - 1) & ")"
        End Select
        .MoveNext
    Loop
   .Close
End With
'Liberado
End Sub

Private Sub Timer1_Timer()
Le
End Sub
