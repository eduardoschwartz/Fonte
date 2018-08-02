VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmReportMob1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Relatório dos Devedores do ISS Eletrônico"
   ClientHeight    =   1560
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1560
   ScaleWidth      =   5400
   Begin VB.OptionButton Opt 
      Caption         =   "Detalhado"
      Height          =   285
      Index           =   1
      Left            =   3780
      TabIndex        =   6
      Top             =   315
      Width           =   1230
   End
   Begin VB.OptionButton Opt 
      Caption         =   "Resumido"
      Height          =   285
      Index           =   0
      Left            =   2430
      TabIndex        =   5
      Top             =   315
      Value           =   -1  'True
      Width           =   1230
   End
   Begin VB.TextBox txtAno 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   990
      TabIndex        =   4
      Top             =   270
      Width           =   705
   End
   Begin MSComCtl2.UpDown UDAno 
      Height          =   285
      Left            =   1710
      TabIndex        =   3
      Top             =   270
      Width           =   255
      _ExtentX        =   450
      _ExtentY        =   503
      _Version        =   393216
      Value           =   2010
      BuddyControl    =   "txtAno"
      BuddyDispid     =   196610
      OrigLeft        =   1800
      OrigTop         =   270
      OrigRight       =   2055
      OrigBottom      =   555
      Max             =   2010
      Min             =   2008
      SyncBuddy       =   -1  'True
      BuddyProperty   =   65547
      Enabled         =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   360
      Left            =   3915
      TabIndex        =   0
      ToolTipText     =   "Imprimir Relatório"
      Top             =   1035
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   635
      BTYPE           =   3
      TX              =   "Imprimir"
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
      MICON           =   "frmReportMob1.frx":0000
      PICN            =   "frmReportMob1.frx":001C
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
      Left            =   270
      TabIndex        =   1
      Top             =   1125
      Width           =   3390
      _ExtentX        =   5980
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
      Caption         =   "Ano..:"
      Height          =   195
      Left            =   270
      TabIndex        =   2
      Top             =   315
      Width           =   645
   End
End
Attribute VB_Name = "frmReportMob1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type ISSELETRO
    nCodReduz As Long
    sMes1 As String
    sMes2 As String
    sMes3 As String
    sMes4 As String
    sMes5 As String
    sMes6 As String
    sMes7 As String
    sMes8 As String
    sMes9 As String
    sMes10 As String
    sMes11 As String
    sMes12 As String
End Type

Private Sub cmdPrint_Click()
Dim sNomeArq As String, FF1 As Integer, Sql As String, RdoAux As rdoResultset, m As Integer, c As Integer
Dim nAno As Integer, aISS() As ISSELETRO, nPos As Integer, nTot As Long

If Val(txtAno.Text) < 2008 Or Val(txtAno.Text) > Year(Now) Then
    MsgBox "Ano inválido", vbExclamation, "Atenção"
    Exit Sub
End If

ReDim aISS(0)
nAno = 2009

Ocupado

Sql = "SELECT DISTINCT identificaprestador From nfisseletro WHERE (tiponota = 1) AND "
Sql = Sql & "(identificaprestador BETWEEN 100000 AND 300000) AND (anoref = " & nAno & ") ORDER BY IDENTIFICAPRESTADOR"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        ReDim Preserve aISS(UBound(aISS) + 1)
        nPos = UBound(aISS)
        aISS(nPos).nCodReduz = !IdentificaPrestador
        aISS(nPos).sMes1 = "0"
        aISS(nPos).sMes2 = "0"
        aISS(nPos).sMes3 = "0"
        aISS(nPos).sMes4 = "0"
        aISS(nPos).sMes5 = "0"
        aISS(nPos).sMes6 = "0"
        aISS(nPos).sMes7 = "0"
        aISS(nPos).sMes8 = "0"
        aISS(nPos).sMes9 = "0"
        aISS(nPos).sMes10 = "0"
        aISS(nPos).sMes11 = "0"
        aISS(nPos).sMes12 = "0"
       .MoveNext
    Loop
   .Close
End With

PBar.Color = vbRed
nTot = UBound(aISS)
'For m = 1 To UBound(aISS)
For m = 1 To 50
    CallPb CLng(m), nTot
    Sql = "SELECT * From nfisseletro2 WHERE identificaprestador = " & aISS(m).nCodReduz & " and tiponota = 1 AND "
    Sql = Sql & "anoref = " & nAno & " ORDER BY MESREF"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        Do Until .EOF
            For c = 1 To UBound(aISS)
                If aISS(m).nCodReduz = aISS(c).nCodReduz Then
                    Select Case !MesRef
                        Case 1
                            aISS(c).sMes1 = CDbl(aISS(c).sMes1) + !ValorTotal
                    End Select
                End If
            Next
           .MoveNext
        Loop
       .Close
    End With
Next
PBar.Color = vbWhite
Liberado

MsgBox "Relatório disponível em " & sPathBin & "\REPORTMOB2.TXT"

End Sub

Private Sub CallPb(nPosF As Long, nTotal As Long)
On Error GoTo Erro
If cGetInputState() <> 0 Then DoEvents

If ((nPosF * 100) / nTotal) <= 100 Then
   PBar.Value = (nPosF * 100) / nTotal
Else
   PBar.Value = 100
End If

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub Form_Load()
UDAno.Max = Year(Now)
txtAno.Text = Year(Now)
Centraliza Me
End Sub

