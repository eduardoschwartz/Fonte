VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCalculo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cálculo geral de IPTU"
   ClientHeight    =   3570
   ClientLeft      =   4350
   ClientTop       =   3330
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3570
   ScaleWidth      =   7485
   Begin VB.Frame Frame1 
      Height          =   2445
      Left            =   90
      TabIndex        =   4
      Top             =   600
      Width           =   7305
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor c/Isenção"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Index           =   7
         Left            =   5520
         TabIndex        =   27
         Top             =   210
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Valor s/Isenção"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Index           =   6
         Left            =   3540
         TabIndex        =   26
         Top             =   210
         Width           =   1605
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "Qtde."
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004080&
         Height          =   225
         Index           =   5
         Left            =   2430
         TabIndex        =   25
         Top             =   210
         Width           =   645
      End
      Begin VB.Label lblValorTot 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5250
         TabIndex        =   24
         Top             =   2040
         Width           =   1845
      End
      Begin VB.Label lblValorIP 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5250
         TabIndex        =   23
         Top             =   1590
         Width           =   1845
      End
      Begin VB.Label lblValorIA 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5250
         TabIndex        =   22
         Top             =   1230
         Width           =   1845
      End
      Begin VB.Label lblValorIM 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5250
         TabIndex        =   21
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label lblValorIN 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   225
         Left            =   5250
         TabIndex        =   20
         Top             =   540
         Width           =   1845
      End
      Begin VB.Label lblValorTotFull 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3300
         TabIndex        =   19
         Top             =   2040
         Width           =   1845
      End
      Begin VB.Label lblValorIPFull 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3300
         TabIndex        =   18
         Top             =   1590
         Width           =   1845
      End
      Begin VB.Label lblValorIAFull 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3300
         TabIndex        =   17
         Top             =   1230
         Width           =   1845
      End
      Begin VB.Label lblValorIMFull 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3300
         TabIndex        =   16
         Top             =   870
         Width           =   1845
      End
      Begin VB.Label lblValorINFull 
         Alignment       =   1  'Right Justify
         Caption         =   "R$ 0,00"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3300
         TabIndex        =   15
         Top             =   540
         Width           =   1845
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00800080&
         X1              =   150
         X2              =   7260
         Y1              =   1920
         Y2              =   1920
      End
      Begin VB.Label lblQtdeTot 
         Alignment       =   2  'Center
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   2400
         TabIndex        =   14
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label lblQtdeIP 
         Alignment       =   2  'Center
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   2400
         TabIndex        =   13
         Top             =   1590
         Width           =   615
      End
      Begin VB.Label lblQtdeIA 
         Alignment       =   2  'Center
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   2400
         TabIndex        =   12
         Top             =   1230
         Width           =   615
      End
      Begin VB.Label lblQtdeIM 
         Alignment       =   2  'Center
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   2400
         TabIndex        =   11
         Top             =   870
         Width           =   615
      End
      Begin VB.Label lblQtdeIN 
         Alignment       =   2  'Center
         Caption         =   "00000"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00004000&
         Height          =   225
         Left            =   2400
         TabIndex        =   10
         Top             =   540
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Isentos processo...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   180
         TabIndex        =   9
         Top             =   1590
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Total de Imóveis...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   2040
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Isentos por área...:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   180
         TabIndex        =   7
         Top             =   1230
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Imunidade total....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   6
         Top             =   870
         Width           =   2115
      End
      Begin VB.Label Label2 
         Caption         =   "Imóveis normais....:"
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   5
         Top             =   540
         Width           =   2115
      End
   End
   Begin VB.ComboBox cmbAno 
      Height          =   315
      ItemData        =   "frmCalculo.frx":0000
      Left            =   1620
      List            =   "frmCalculo.frx":0002
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   180
      Width           =   1125
   End
   Begin prjChameleon.chameleonButton cmdExec 
      Default         =   -1  'True
      Height          =   345
      Left            =   5700
      TabIndex        =   0
      ToolTipText     =   "Executar cálculo"
      Top             =   3150
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Executar Cálculo"
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
      MICON           =   "frmCalculo.frx":0004
      PICN            =   "frmCalculo.frx":0020
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
      Left            =   120
      TabIndex        =   3
      Top             =   3180
      Width           =   3645
      _ExtentX        =   6429
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
      Caption         =   "Ano de Cálculo....:"
      Height          =   225
      Left            =   150
      TabIndex        =   1
      Top             =   240
      Width           =   1425
   End
End
Attribute VB_Name = "frmCalculo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aParc() As Date, sVenctoUnica2 As String, sVenctoUnica3 As String

Private Sub cmdExec_Click()

If MsgBox("Deseja executar o cálculo geral de IPTU para o ano de " & cmbAno.Text, vbQuestion + vbYesNo, "Confirmação") = vbYes Then
    cmdExec.Enabled = False
    CarregaTela
    EfetuaCalculo
    cmdExec.Enabled = True
End If

End Sub

Private Sub Form_Load()
Dim x As Integer

Centraliza Me
PBar.Color = vbWhite
For x = 2010 To Year(Now) + 1
    cmbAno.AddItem x
Next

cmbAno.ListIndex = cmbAno.ListCount - 1

End Sub

Private Sub EfetuaCalculo()
Dim Sql As String, RdoAux As rdoResultset, qd As New rdoQuery, RdoAux2 As rdoResultset, nAno As Integer
Dim nQtdeIN As Long, nQtdeIM As Long, nQtdeIA As Long, nQtdeIP As Long, nQtdeTot As Long
Dim nValorIN As Double, nValorIM As Double, nValorIA As Double, nValorIP As Double, nValorTot As Double, nNumParc As Integer
Dim nValorINFull As Double, nValorIMFull As Double, nValorIAFull As Double, nValorIPFull As Double, nValorTotFull As Double
Dim nCodReduz As Long, nPos As Long, nTot As Long, nTipoIsencao As Integer, nPercIsencao As Double, nLastDoc As Long
Dim nVVT As Double, nVVP As Double, nVVI As Double, nValorFinal As Double, nValorFinalFull As Double, nAgrupamento As Double
Dim nValorParcela As Double, nValorUnica As Double, nValorUnica2 As Double, nValorUnica3 As Double, x As Integer
Dim nValorExp As Double, nCodTributo As Integer, bDocInc1 As Boolean, bDocInc2 As Boolean

Set qd.ActiveConnection = cn
bDocInc1 = False
bDocInc2 = False
nAno = Val(cmbAno.Text)

Sql = "delete from isentoipturel"
cn.Execute Sql, rdExecDirect

If NomeDeLogin = "SCHWARTZ" And nAno = 2018 Then
    Sql = "DELETE FROM LASERIPTU WHERE ANO=" & nAno
    'cn.Execute Sql, rdExecDirect
End If

Open sPathBin & "\DEBITOPARCELA.TXT" For Output As #1
Open sPathBin & "\DEBITOTRIBUTO.TXT" For Output As #2
Open sPathBin & "\PARCELADOCUMENTO.TXT" For Output As #3
Open sPathBin & "\NUMDOCUMENTO.TXT" For Output As #4

nLastDoc = 15787960
nValorExp = 0

Sql = "select codreduzido from cadimob where inativo=0 order by codreduzido"
'Sql = "select codreduzido from cadimob where codreduzido<2000"
'Sql = "select codigo as codreduzido from table1 order by codigo"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    nTot = .RowCount: nPos = 1
    Do Until .EOF
        nCodReduz = !CODREDUZIDO
        
'        If nLastDoc > 15556090 Then
'            nLastDoc = 15567114
'        End If
        
      ' If nCodReduz = 1886 Then MsgBox "teste"
        If nPos Mod 100 = 0 Then
            CallPb nPos, nTot
            GoSub Tela
        End If
        
        On Error Resume Next
        RdoAux2.Close
        On Error GoTo 0
        qd.Sql = "{ Call spCalculo(?,?) }"
        qd(0) = nCodReduz
        qd(1) = Val(cmbAno.Text)
        Set RdoAux2 = qd.OpenResultset(rdOpenKeyset)
        With RdoAux2
            nTipoIsencao = !tipoisencao
            If IsNull(!vvt) Then GoTo proximo
            nVVT = !vvt
            nVVP = !VVP
            nVVI = !VVI
            nValorFinal = !valorfinal
            nValorFinalFull = !valorfinalfull
            nPercIsencao = !percisencao
            nAgrupamento = !Agrupamento
           'totais
            If nTipoIsencao = 1 Then
                nQtdeIM = nQtdeIM + 1
                nValorIM = nValorIM + FormatNumber(nValorFinal, 2)
                nValorIMFull = nValorIMFull + FormatNumber(nValorFinalFull, 2)
            ElseIf nTipoIsencao = 3 Then
                nQtdeIP = nQtdeIP + 1
                nValorIP = nValorIP + FormatNumber(nValorFinal, 2)
                nValorIPFull = nValorIPFull + FormatNumber(nValorFinalFull, 2)
            ElseIf nTipoIsencao = 2 Then
                nQtdeIA = nQtdeIA + 1
                nValorIA = nValorIA + FormatNumber(nValorFinal, 2)
                nValorIAFull = nValorIAFull + FormatNumber(nValorFinalFull, 2)
            Else
                nQtdeIN = nQtdeIN + 1
                nValorIN = nValorIN + FormatNumber(nValorFinal, 2)
                nValorINFull = nValorINFull + FormatNumber(nValorFinalFull, 2)
            End If
            
            If nTipoIsencao > 0 Then
                Sql = "insert isentoipturel (codreduzido,vvt,vvp,vvi,areat,areac,valoriptu,tipoisencao) values(" & nCodReduz & ","
                Sql = Sql & Virg2Ponto(CStr(nVVT)) & "," & Virg2Ponto(CStr(nVVP)) & "," & Virg2Ponto(CStr(nVVI)) & "," & Virg2Ponto(CStr(!AreaTerreno)) & ","
                Sql = Sql & Virg2Ponto(CStr(!AreaPredial)) & "," & Virg2Ponto(CStr(nValorFinalFull)) & "," & nTipoIsencao & ")"
                cn.Execute Sql, rdExecDirect
                DoEvents
            End If
           
           GoTo proximo
            nQtdeTot = nQtdeTot + 1
            nValorTot = nValorTot + FormatNumber(nValorFinal, 2)
            nValorTotFull = nValorTotFull + FormatNumber(nValorFinalFull, 2)
         
            If nVVP = 0 Then
                nCodTributo = 2
            Else
                nCodTributo = 1
            End If
         
            '** NOVA ROTINA DE QTDE DE PARCELAS ***
            If nValorFinal > 0 And nValorFinal <= 10 Then
                nNumParc = 1
            ElseIf nValorFinal > 10 And nValorFinal <= 20 Then nNumParc = 1
            ElseIf nValorFinal > 20 And nValorFinal <= 30 Then nNumParc = 2
            ElseIf nValorFinal > 30 And nValorFinal <= 40 Then nNumParc = 3
            ElseIf nValorFinal > 40 And nValorFinal <= 50 Then nNumParc = 4
            ElseIf nValorFinal > 50 And nValorFinal <= 60 Then nNumParc = 5
            ElseIf nValorFinal > 60 And nValorFinal <= 70 Then nNumParc = 6
            ElseIf nValorFinal > 70 And nValorFinal <= 80 Then nNumParc = 7
            ElseIf nValorFinal > 80 And nValorFinal <= 90 Then nNumParc = 8
            ElseIf nValorFinal > 90 And nValorFinal <= 100 Then nNumParc = 9
            Else
                nNumParc = 12
            End If
            '**************************************
            
            'nValorUnica = Round(nValorFinal - (nValorFinal * 0.06), 2)
            nValorUnica = !ValorUnica
            nValorUnica2 = !ValorUnica2
            nValorUnica3 = !ValorUnica3
            nValorParcela = Round(nValorFinal / nNumParc, 2)
           ' If nTipoIsencao > 0 Then MsgBox "teste"

            
            'GRAVA TABELA LASERIPTU
            If NomeDeLogin = "SCHWARTZ" And (nTipoIsencao = 0 Or (nTipoIsencao = 3 And nPercIsencao <> 100)) Then
                Sql = "INSERT LASERIPTU (ANO,CODREDUZIDO,VVT,VVC,VVI,IMPOSTOPREDIAL,IMPOSTOTERRITORIAL,NATUREZA,AREACONSTRUCAO,"
                Sql = Sql & "TESTADAPRINC,VALORTOTALPARC,VALORTOTALUNICA,VALORTOTALUNICA2,VALORTOTALUNICA3,QTDEPARC,TXEXPPARC,TXEXPUNICA,AREATERRENO,FATORCAT,FATORPED,FATORSIT,"
                Sql = Sql & "FATORPRO,FATORTOP,FATORDIS,FATORGLE,AGRUPAMENTO,FRACAOIDEAL,ALIQUOTA) VALUES("
                Sql = Sql & nAno & "," & nCodReduz & "," & Virg2Ponto(CStr(!vvt)) & "," & Virg2Ponto(CStr(!VVP)) & ","
                Sql = Sql & Virg2Ponto(CStr(!VVI)) & "," & Virg2Ponto(CStr(!ValorIPTU)) & "," & Virg2Ponto(CStr(!ValorITU)) & ",'"
                Sql = Sql & !Natureza & "'," & Virg2Ponto(CStr(!AreaPredial)) & "," & Virg2Ponto(CStr(!TESTADAPRINC)) & ","
                Sql = Sql & Virg2Ponto(CStr(!VALORPARCELA)) & "," & Virg2Ponto(CStr(!ValorUnica)) & "," & Virg2Ponto(CStr(!ValorUnica2)) & "," & Virg2Ponto(CStr(!ValorUnica3)) & "," & !qtdeparc & ","
                Sql = Sql & 0 & "," & 0 & "," & Virg2Ponto(CStr(!AreaTerreno)) & ","
                Sql = Sql & Virg2Ponto(CStr(!fcat)) & "," & Virg2Ponto(CStr(!fped)) & "," & Virg2Ponto(CStr(!fsit)) & "," & Virg2Ponto(CStr(!fpro)) & ","
                Sql = Sql & Virg2Ponto(CStr(!ftop)) & "," & Virg2Ponto(CStr(IIf(IsNull(!fdis), "0,00", !fdis))) & "," & Virg2Ponto(CStr(!fgle)) & "," & Virg2Ponto(CStr(!valorAgrupamento)) & ","
                Sql = Sql & Virg2Ponto(CStr(!FRACAO)) & "," & Virg2Ponto(CStr(!Aliquota * 100)) & ")"
                'cn.Execute Sql, rdExecDirect
            
            End If
            
           .Close
        End With
        
        If nTipoIsencao = 0 Or (nTipoIsencao = 3 And nPercIsencao <> 100) Then
            For x = 0 To nNumParc
                DoEvents
                'GRAVA NA TABELA DEBITOPARCELA
                ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
                ax = ax & 18 & "," & Format(aParc(x), "mm/dd/yyyy") & "," & Format("01/01/" & CStr(nAno), "mm/dd/yyyy") & ","
                ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & ","
                ax = ax & Null & "," & 0
                Print #1, ax
                'GRAVA NA TABELA DEBITO TRIBUTO
                ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & x & "," & 0 & ","
                ax = ax & nCodTributo & "," & Virg2Ponto(IIf(x = 0, Round(nValorUnica, 2), Round(nValorParcela, 2))) & ","
                ax = ax & 0 & "," & 0 & "," & 0
                Print #2, ax
                'GRAVA NA TABELA NUMDOCUMENTO
                nLastDoc = nLastDoc + 1
                ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & Virg2Ponto(CStr(nValorExp))
                Print #4, ax
                'GRAVA NA TABELA PARCELADOCUMENTO
                ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & ","
                ax = ax & x & "," & 0 & "," & nLastDoc & "," & "0" & "," & "0"
                Print #3, ax
            Next
            '******************************
            '****** COMPLEM. 91 ***********
            '******************************
            'GRAVA NA TABELA DEBITOPARCELA
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 91 & ","
            ax = ax & 18 & "," & Format(sVenctoUnica2, "mm/dd/yyyy") & "," & Format("01/01/" & CStr(nAno), "mm/dd/yyyy") & ","
            ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & "," & Null & "," & 0
            Print #1, ax
            'GRAVA NA TABELA DEBITO TRIBUTO
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 91 & ","
            ax = ax & nCodTributo & "," & Virg2Ponto(Round(nValorUnica2, 2)) & "," & 0 & "," & 0 & "," & 0
            Print #2, ax
            'GRAVA NA TABELA NUMDOCUMENTO
            nLastDoc = nLastDoc + 1
            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & 0
            Print #4, ax
            'GRAVA NA TABELA PARCELADOCUMENTO
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 91 & "," & nLastDoc & "," & "0" & "," & "0"
            Print #3, ax
            '******************************
            '****** COMPLEM. 91 ***********
            '******************************
            'GRAVA NA TABELA DEBITOPARCELA
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 92 & ","
            ax = ax & 18 & "," & Format(sVenctoUnica3, "mm/dd/yyyy") & "," & Format("01/01/" & CStr(nAno), "mm/dd/yyyy") & ","
            ax = ax & 1 & "," & 0 & "," & 0 & "," & 0 & "," & Null & "," & Null & "," & 0
            Print #1, ax
            'GRAVA NA TABELA DEBITO TRIBUTO
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 92 & ","
            ax = ax & nCodTributo & "," & Virg2Ponto(Round(nValorUnica3, 2)) & "," & 0 & "," & 0 & "," & 0
            Print #2, ax
            'GRAVA NA TABELA NUMDOCUMENTO
            nLastDoc = nLastDoc + 1
            ax = nLastDoc & "," & Format(Now, "mm/dd/yyyy") & "," & 0 & "," & Null & "," & 0 & "," & 0
            Print #4, ax
            'GRAVA NA TABELA PARCELADOCUMENTO
            ax = nCodReduz & "," & nAno & "," & 1 & "," & 0 & "," & 0 & "," & 92 & "," & nLastDoc & "," & "0" & "," & "0"
            Print #3, ax
        End If
        
proximo:
        nPos = nPos + 1
       .MoveNext
    Loop
   .Close
End With

Close #4
Close #3
Close #2
Close #1

GoSub Tela

MsgBox "Cálculo finalizado.", vbInformation, "Atenção"

PBar.Color = vbWhite
PBar.value = 0
Exit Sub

Tela:
lblQtdeIN.Caption = Format(nQtdeIN, "00000")
lblValorIN.Caption = Format(nValorIN, "R$ #0.00")
lblValorINFull.Caption = Format(nValorINFull, "R$ #0.00")
lblQtdeIA.Caption = Format(nQtdeIA, "00000")
lblValorIA.Caption = Format(nValorIA, "R$ #0.00")
lblValorIAFull.Caption = Format(nValorIAFull, "R$ #0.00")
lblQtdeIP.Caption = Format(nQtdeIP, "00000")
lblValorIP.Caption = Format(nValorIP, "R$ #0.00")
lblValorIPFull.Caption = Format(nValorIPFull, "R$ #0.00")
lblQtdeIM.Caption = Format(nQtdeIM, "00000")
lblValorIM.Caption = Format(nValorIM, "R$ #0.00")
lblValorIMFull.Caption = Format(nValorIMFull, "R$ #0.00")
lblQtdeTot.Caption = Format(nQtdeTot, "00000")
lblValorTot.Caption = Format(nValorTot, "R$ #0.00")
lblValorTotFull.Caption = Format(nValorTotFull, "R$ #0.00")
Return

End Sub

Private Sub CarregaTela()

Sql = "SELECT ANO,QTDEPARCELA,PARCELAUNICA,DESCONTOUNICA,VENCUNICA,VENCUNICA2,VENCUNICA3,VENC01,VENC02,VENC03,VENC04,VENC05,"
Sql = Sql & "VENC06,VENC07,VENC08,VENC09,VENC10,VENC11,VENC12 FROM PARAMPARCELA WHERE CODTIPO=1 AND ANO=" & Val(cmbAno.Text)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     If .RowCount = 0 Then GoTo fim
     ReDim aParc(!qtdeparcela)
     Do Until .EOF
        If Not IsNull(!vencunica) Then aParc(0) = Format(!vencunica, "dd/mm/yyyy")
        If Not IsNull(!venc01) Then aParc(1) = Format(!venc01, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc02) Then aParc(2) = Format(!venc02, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc03) Then aParc(3) = Format(!venc03, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc04) Then aParc(4) = Format(!venc04, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc05) Then aParc(5) = Format(!venc05, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc06) Then aParc(6) = Format(!venc06, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc07) Then aParc(7) = Format(!venc07, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc08) Then aParc(8) = Format(!venc08, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc09) Then aParc(9) = Format(!venc09, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc10) Then aParc(10) = Format(!venc10, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc11) Then aParc(11) = Format(!venc11, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!venc12) Then aParc(12) = Format(!venc12, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!vencunica2) Then sVenctoUnica2 = Format(!vencunica2, "dd/mm/yyyy") Else Exit Do
        If Not IsNull(!vencunica3) Then sVenctoUnica3 = Format(!vencunica3, "dd/mm/yyyy") Else Exit Do
        x = x + 1
       .MoveNext
     Loop
    .Close
End With
fim:
End Sub

Private Sub CallPb(nVal As Long, nTot As Long)
If nVal > 0 Then
    PBar.Color = &HC0C000
Else
    PBar.Color = vbWhite
End If
If ((nVal * 100) / nTot) <= 100 Then
   PBar.value = (nVal * 100) / nTot
Else
   PBar.value = 100
End If

Me.Refresh
If cGetInputState() <> 0 Then DoEvents

End Sub

