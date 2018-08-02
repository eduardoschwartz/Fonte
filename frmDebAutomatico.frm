VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmDebAutomatico 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Retorno de Débito Automático"
   ClientHeight    =   5625
   ClientLeft      =   4230
   ClientTop       =   3225
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5625
   ScaleWidth      =   7485
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   0
      Left            =   4620
      TabIndex        =   0
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   90
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDebAutomatico.frx":0000
      PICN            =   "frmDebAutomatico.frx":001C
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
      Left            =   6270
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   5070
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDebAutomatico.frx":08C9
      PICN            =   "frmDebAutomatico.frx":08E5
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComCtl2.MonthView Mv 
      Height          =   2370
      Left            =   90
      TabIndex        =   2
      Top             =   390
      Width           =   2490
      _ExtentX        =   4392
      _ExtentY        =   4180
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   15658734
      Appearance      =   1
      StartOfWeek     =   50921473
      TitleBackColor  =   192
      TitleForeColor  =   12648447
      CurrentDate     =   37439
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   1
      Left            =   4620
      TabIndex        =   3
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   1050
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmDebAutomatico.frx":0953
      PICN            =   "frmDebAutomatico.frx":096F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   2
      Left            =   4620
      TabIndex        =   4
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   2010
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmDebAutomatico.frx":1A2B
      PICN            =   "frmDebAutomatico.frx":1D45
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   3
      Left            =   4620
      TabIndex        =   15
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   3000
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmDebAutomatico.frx":2765
      PICN            =   "frmDebAutomatico.frx":2A7F
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdBanco 
      Height          =   975
      Index           =   4
      Left            =   4620
      TabIndex        =   18
      ToolTipText     =   "Ajuda desta Tela"
      Top             =   3990
      Width           =   2715
      _ExtentX        =   4789
      _ExtentY        =   1720
      BTYPE           =   9
      TX              =   ""
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
      COLTYPE         =   1
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   99
      MICON           =   "frmDebAutomatico.frx":3A03
      PICN            =   "frmDebAutomatico.frx":3D1D
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   -1  'True
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   4
      Left            =   4170
      TabIndex        =   20
      Top             =   2190
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Banespa..............:"
      Height          =   225
      Index           =   2
      Left            =   2760
      TabIndex        =   19
      Top             =   2205
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "HSBC.................:"
      Height          =   225
      Index           =   1
      Left            =   2760
      TabIndex        =   17
      Top             =   1785
      Width           =   1365
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   3
      Left            =   4170
      TabIndex        =   16
      Top             =   1770
      Width           =   315
   End
   Begin VB.Label lblAux 
      Height          =   225
      Left            =   3060
      TabIndex        =   14
      Top             =   6450
      Width           =   1755
   End
   Begin VB.Label lblData 
      Height          =   255
      Left            =   4830
      TabIndex        =   13
      Top             =   5160
      Width           =   1365
   End
   Begin VB.Label lblTit 
      BackStyle       =   0  'Transparent
      Caption         =   "Arquivos Disponíveis em"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Left            =   2790
      TabIndex        =   12
      Top             =   90
      Width           =   1785
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Bradesco.............:"
      Height          =   225
      Index           =   3
      Left            =   2760
      TabIndex        =   11
      Top             =   975
      Width           =   1365
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Itau......................:"
      Height          =   225
      Index           =   6
      Left            =   2760
      TabIndex        =   10
      Top             =   600
      Width           =   1365
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   1
      Left            =   4170
      TabIndex        =   9
      Top             =   960
      Width           =   315
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   0
      Left            =   4170
      TabIndex        =   8
      Top             =   615
      Width           =   315
   End
   Begin VB.Label lblMsg 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Selecione a Data de Geração dos Arquivos"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   255
      Left            =   90
      TabIndex        =   7
      Top             =   5130
      Width           =   4605
   End
   Begin VB.Label lblBanco 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      ForeColor       =   &H000000C0&
      Height          =   225
      Index           =   2
      Left            =   4170
      TabIndex        =   6
      Top             =   1350
      Width           =   315
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Unibanco............:"
      Height          =   225
      Index           =   0
      Left            =   2760
      TabIndex        =   5
      Top             =   1365
      Width           =   1365
   End
End
Attribute VB_Name = "frmDebAutomatico"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Type FebrabanA
    CodigoRegistro As String * 1
    CodigoRemessa As String * 1
    CodigoConvenio As String * 20
    NomeEmpresa As String * 20
    CodigoBanco As String * 3
    NomeBanco As String * 20
    DataGeracao As String * 8
    NumeroSeq As String * 6
    VersaoLayout As String * 2
    Filler As String * 69
End Type


Private Type FebrabanF 'RETORNO DO DEBITO AUTOMATICO
   CodigoRegistro As String * 1
   Distrito As String * 2
   Setor As String * 2
   Quadra As String * 4
   Lote As String * 5
   Seq As String * 2
   FillerID As String * 10
   CodAgencia As String * 4
   ContaCliente As String * 14
   DataVencto As String * 8
   ValorDebito As String * 15
   CodRetorno As String * 2
   NumDoc As String * 9
   Filler1 As String * 51
   Filler2 As String * 20
   CodMovimento As String * 1
End Type


Private Sub cmdBanco_Click(Index As Integer)

If Val(lblBanco(Index).Caption) = 0 Then
    MsgBox "Não existem arquivos para este banco na data especificada.", vbInformation, "Atenção"
    Exit Sub
End If
lblAux.Caption = Index
frmDebBanco.show vbModeless, frmMdi

Select Case Index
    Case 0
        frmDebBanco.lblBanco.Caption = "341 - ITAU"
    Case 1
        frmDebBanco.lblBanco.Caption = "237 - BRADESCO"
    Case 2
        frmDebBanco.lblBanco.Caption = "409 - UNIBANCO"
    Case 3
        frmDebBanco.lblBanco.Caption = "399 - HSBC"
    Case 4
        frmDebBanco.lblBanco.Caption = "033 - BANESPA"
End Select

frmDebBanco.grdArq.Rows = 1
Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(lblData.Caption, "mm/dd/yyyy") & "' AND CODBANCO=" & Val(Left$(frmDebBanco.lblBanco.Caption, 3)) & " AND DA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       frmDebBanco.grdArq.AddItem sPathArqBanco & "\" & Mv.Year & "\" & Format(Mv.Month, "00") & "\" & Format(Mv.Day, "00") & "\" & Chr(9) & !NOMEARQ
      .MoveNext
    Loop
End With

End Sub


Private Sub cmdSair_Click()
Unload Me
End Sub


Private Function LikroMila(sFullPath As Variant, dData As Date) As String

On Error GoTo Erro

Dim Header As FebrabanA
Dim Corpo As FebrabanF
Dim sData As String
'Dim sData As String
Dim f As String

'01/03/02
f = CStr(dData)
If Len(dData) < 10 Then
    f = Left$(f, 6) & "20" & Right$(f, 2)
    sData = Right$(f, 4) & Mid$(f, 4, 2) & Left$(f, 2)
Else
    sData = Right$(dData, 4) & Mid$(dData, 4, 2) & Left$(dData, 2)
End If

Posicao = 0
Open sFullPath For Binary Access Read Write As #1
    Get #1, 1, Header
    Posicao = Len(Header) + 3
    Get #1, Posicao, Corpo
    If Corpo.DataVencto = sData Then
         LikroMila = Header.CodigoBanco
    End If
 Close #1

 Exit Function
Erro:
 MsgBox Err.Description
 
End Function

Private Sub Form_Load()

Ocupado
Centraliza Me
Mv.Day = Day(Now)
Mv.Month = Month(Now)
Mv.Year = Year(Now)
Liberado
frmMdi.AddWindow Me.Name, Me.Caption
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
frmMdi.RemoveWindow Me.Name
End Sub

Private Sub Mv_DateClick(ByVal DateClicked As Date)
Dim RdoAux As rdoResultset
lblData.Caption = Mv.Value

Screen.MousePointer = vbHourglass
lblMsg.ForeColor = vbRed
lblMsg.Caption = "Aguarde... Lendo Arquivos."
lblMsg.Refresh
LimpaContador

Sql = "SELECT * FROM ARQUIVOBANCO WHERE DATACREDITO='" & Format(Mv.Value, "mm/dd/yyyy") & "' AND DA=1"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
       Select Case !CodBanco
            Case 237 'BRADESCO
                lblBanco(1).Caption = Val(lblBanco(1).Caption) + 1
            Case 341 'ITAU
                lblBanco(0).Caption = Val(lblBanco(0).Caption) + 1
            Case 409 'UNIBANCO
                lblBanco(2).Caption = Val(lblBanco(2).Caption) + 1
            Case 399 'HSBC
                lblBanco(3).Caption = Val(lblBanco(3).Caption) + 1
            Case 33 'BANESPA
                lblBanco(4).Caption = Val(lblBanco(4).Caption) + 1
       End Select
      .MoveNext
    Loop
End With

Screen.MousePointer = vbDefault
lblMsg.ForeColor = &HC00000
lblMsg.Caption = "Selecione a Data de Geração dos Arquivos"
lblMsg.Refresh
   
End Sub

Private Sub LimpaContador()
For x = 0 To 4
      lblBanco(x).Caption = 0
Next

End Sub

'Private Sub AtualizaContador()
'Dim x As Integer
'With grdMain
'      For x = 1 To .Rows - 1
'        If .TextMatrix(x, 2) = "341" Then
'             lblBanco(0).Caption = Val(lblBanco(0).Caption) + 1
'             lblBanco(0).Refresh
'        ElseIf .TextMatrix(x, 2) = "237" Then
'             lblBanco(1).Caption = Val(lblBanco(1).Caption) + 1
'             lblBanco(1).Refresh
'        ElseIf .TextMatrix(x, 2) = "409" Then
'             lblBanco(2).Caption = Val(lblBanco(2).Caption) + 1
'             lblBanco(2).Refresh
'        ElseIf .TextMatrix(x, 2) = "399" Then
'             lblBanco(3).Caption = Val(lblBanco(3).Caption) + 1
'             lblBanco(3).Refresh
'        ElseIf .TextMatrix(x, 2) = "033" Then
'             lblBanco(4).Caption = Val(lblBanco(4).Caption) + 1
'             lblBanco(4).Refresh
'        End If
'      Next
'End With

'End Sub

