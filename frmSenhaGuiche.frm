VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSenhaGuiche 
   BackColor       =   &H00404040&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Guiche Nº "
   ClientHeight    =   5205
   ClientLeft      =   6090
   ClientTop       =   2220
   ClientWidth     =   3135
   Icon            =   "frmSenhaGuiche.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5205
   ScaleWidth      =   3135
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   2205
      Top             =   5445
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   495
      TabIndex        =   0
      Text            =   "Text1"
      Top             =   5535
      Width           =   1095
   End
   Begin prjChameleon.chameleonButton cmdSenha 
      Height          =   510
      Index           =   0
      Left            =   180
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   150
      Width           =   2760
      _ExtentX        =   4868
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":030A
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
      Left            =   180
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   735
      Width           =   2760
      _ExtentX        =   4868
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":0624
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
      Left            =   180
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   1320
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "PREFERÊNCIAL"
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":093E
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
      Left            =   180
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   1905
      Width           =   2760
      _ExtentX        =   4868
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":0C58
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
      Index           =   4
      Left            =   180
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   2475
      Width           =   2760
      _ExtentX        =   4868
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":0F72
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdHist 
      Height          =   330
      Left            =   2565
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   4725
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   582
      BTYPE           =   14
      TX              =   "Hist"
      ENAB            =   -1  'True
      BeginProperty FONT {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
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
      FCOL            =   65535
      FCOLO           =   65535
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":128C
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
      Left            =   180
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   3060
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   900
      BTYPE           =   14
      TX              =   "TRANSP.UNIV."
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
      FCOL            =   65280
      FCOLO           =   65280
      MCOL            =   16777215
      MPTR            =   99
      MICON           =   "frmSenhaGuiche.frx":15A6
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblTempo 
      BackStyle       =   0  'Transparent
      Caption         =   "17:32 MIN"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   11
      Top             =   4770
      Width           =   1275
   End
   Begin VB.Label lblHora 
      BackStyle       =   0  'Transparent
      Caption         =   "12:38"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   10
      Top             =   4455
      Width           =   1455
   End
   Begin VB.Label lblBanda 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "PREFEITURA1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   240
      Left            =   675
      TabIndex        =   9
      Top             =   4140
      Width           =   1995
   End
   Begin VB.Label lblSenha 
      BackStyle       =   0  'Transparent
      Caption         =   "251"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   285
      Left            =   1260
      TabIndex        =   8
      Top             =   3780
      Width           =   1050
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "ESPERA.:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   3
      Left            =   90
      TabIndex        =   7
      Top             =   4770
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "HORA......:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   2
      Left            =   90
      TabIndex        =   6
      Top             =   4455
      Width           =   1230
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "SENHA...:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFC0&
      Height          =   285
      Index           =   1
      Left            =   90
      TabIndex        =   5
      Top             =   3825
      Width           =   1230
   End
End
Attribute VB_Name = "frmSenhaGuiche"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdHist_Click()
frmHistSenha.show vbModal
End Sub

Private Sub cmdSenha_Click(Index As Integer)
Dim Sql As String, RdoAux As rdoResultset, nMinute As Long, nSeconds As Long, nSecond As Long
Dim Hora1 As Date, Hora2 As Date
Dim Daycount As Long, SecondsRemaining As Integer, HourCount As Integer
Dim MinutesCount As Integer, SecondsCount As Integer
If nGuiche = 0 Or nGuiche > 12 Then
    MsgBox "Voce não pode acessar o sistema de senhas!", vbCritical, "Acesso Negado"
    Exit Sub
End If

Ocupado
On Error GoTo Erro
Sql = "SELECT * FROM SSPAC WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now) & " AND "
Sql = Sql & "DAY(DATAENTRADA)=" & Day(Now) & " AND BANDA=" & Index + 1 & " AND DATACHAMADA IS NULL "
Sql = Sql & "ORDER BY SENHA"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If .RowCount = 0 Then GoTo fim
    lblSenha.Caption = Format(!SENHA, "000")
    lblBanda.Caption = cmdSenha(Index).Caption
    lblHora.Caption = !HORAENTRADA
    
    Hora1 = !HORAENTRADA
    Hora2 = Format(Now, "hh:mm:ss")
    
    nSeconds = DateDiff("s", Hora1, Hora2)
    
    Daycount = nSeconds \ (86400)
    SecondsRemaining = nSeconds - (Daycount * (86400))
    HourCount = Abs(SecondsRemaining \ (60 * 60))
    SecondsRemaining = Abs(SecondsRemaining Mod (60 * 60))
    MinutesCount = Abs(SecondsRemaining \ 60)
    SecondsCount = Abs(SecondsRemaining Mod 60)
    
    lblTempo.Caption = Format(HourCount, "00") & ":" & Format(MinutesCount, "00") & ":" & Format(SecondsCount, "00")
    
    Sql = "UPDATE SSPAC SET DATACHAMADA='" & Format(Now, "mm/dd/yyyy") & "',HORACHAMADA='" & Format(Now, "hh:mm:ss") & "',GUICHE=" & nGuiche & ","
    Sql = Sql & "ATENDENTE='" & NomeDeLogin & "',ESPERA='" & lblTempo.Caption & "',MONITOR=0 "
    Sql = Sql & " WHERE YEAR(DATAENTRADA)=" & Year(Now) & " AND MONTH(DATAENTRADA)=" & Month(Now)
    Sql = Sql & " AND DAY(DATAENTRADA)=" & Day(Now) & " AND SENHA=" & !SENHA
    cn.Execute Sql, rdExecDirect
    
   .Close
End With
fim:
Liberado
Text1.SetFocus
Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
    MsgBox rdoErrors(x).Description
Next
End Sub

Private Sub cmdSenha_MouseOut(Index As Integer)
On Error Resume Next
If Me.Enabled = True Then
    Text1.SetFocus
End If
End Sub

Private Sub Le()
Dim Sql As String, RdoAux As rdoResultset, dData As Date, x As Integer
Dim aCount(5) As Integer
dData = Now
'Ocupado
On Error Resume Next
For x = 0 To 5
    cmdSenha(x).Enabled = False
    aCount(x) = 0
Next

cmdSenha(0).Caption = "PREFEITURA (0)"
cmdSenha(1).Caption = "2ª VIA IPTU (0)"
cmdSenha(2).Caption = "PREFERENCIAL (0)"
cmdSenha(3).Caption = "SENHA PAT (0)"
cmdSenha(4).Caption = "REFIS (0)"
cmdSenha(5).Caption = "TRANSP.UNIV. (0)"

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
                cmdSenha(!BANDA - 1).Caption = "TRANSP.UNIV. (" & aCount(!BANDA - 1) & ")"
        End Select
        cmdSenha(!BANDA - 1).Enabled = True
        .MoveNext
    Loop
   .Close
End With
'Liberado
End Sub

Private Sub Form_Load()

Me.Caption = Me.Caption & Format(nGuiche, "00")
Limpa
Le
End Sub

Private Sub Timer1_Timer()
Me.Refresh
Le
End Sub

Private Sub Limpa()
lblSenha.Caption = ""
lblBanda.Caption = ""
lblHora.Caption = ""
lblTempo.Caption = ""
End Sub


