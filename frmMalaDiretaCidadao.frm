VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmMalaDiretaCidadao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mala Direta - Cidadão"
   ClientHeight    =   1980
   ClientLeft      =   3330
   ClientTop       =   4710
   ClientWidth     =   4215
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1980
   ScaleWidth      =   4215
   Begin prjChameleon.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   2925
      TabIndex        =   10
      ToolTipText     =   "Imprimir as cartas"
      Top             =   1575
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   556
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
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmMalaDiretaCidadao.frx":0000
      PICN            =   "frmMalaDiretaCidadao.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame8 
      BackColor       =   &H00EEEEEE&
      Caption         =   "Importação de Arquivos"
      ForeColor       =   &H00800000&
      Height          =   1455
      Left            =   45
      TabIndex        =   1
      Top             =   0
      Width           =   4155
      Begin VB.TextBox txtArq 
         Appearance      =   0  'Flat
         BackColor       =   &H00EEEEEE&
         Height          =   285
         Left            =   90
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   315
         Width           =   3390
      End
      Begin VB.TextBox txtDelimiter 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Left            =   1035
         MaxLength       =   1
         TabIndex        =   2
         Text            =   ","
         Top             =   675
         Width           =   330
      End
      Begin prjChameleon.chameleonButton cmdOpen 
         Height          =   315
         Left            =   3555
         TabIndex        =   0
         ToolTipText     =   "Localizar arquivo texto"
         Top             =   315
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   5
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMalaDiretaCidadao.frx":0176
         PICN            =   "frmMalaDiretaCidadao.frx":0192
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdImportar 
         Height          =   315
         Left            =   1395
         TabIndex        =   4
         ToolTipText     =   "Importar o arquivo selecionado"
         Top             =   675
         Width           =   1170
         _ExtentX        =   2064
         _ExtentY        =   556
         BTYPE           =   5
         TX              =   "Importar"
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
         MICON           =   "frmMalaDiretaCidadao.frx":0219
         PICN            =   "frmMalaDiretaCidadao.frx":0235
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdLimpar 
         Height          =   315
         Left            =   3555
         TabIndex        =   5
         ToolTipText     =   "Limpar texto"
         Top             =   675
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   5
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMalaDiretaCidadao.frx":0448
         PICN            =   "frmMalaDiretaCidadao.frx":0464
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdPreview 
         Height          =   315
         Left            =   3105
         TabIndex        =   6
         ToolTipText     =   "Visualizar arquivo"
         Top             =   675
         Width           =   405
         _ExtentX        =   714
         _ExtentY        =   556
         BTYPE           =   5
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
         COLTYPE         =   2
         FOCUSR          =   0   'False
         BCOL            =   15658734
         BCOLO           =   15658734
         FCOL            =   0
         FCOLO           =   0
         MCOL            =   12632256
         MPTR            =   1
         MICON           =   "frmMalaDiretaCidadao.frx":0681
         PICN            =   "frmMalaDiretaCidadao.frx":069D
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Delimitador.:"
         Height          =   240
         Index           =   0
         Left            =   90
         TabIndex        =   9
         Top             =   735
         Width           =   870
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Registros localizadas no arquivo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   8
         Top             =   1125
         Width           =   2895
      End
      Begin VB.Label lblTotImp 
         BackStyle       =   0  'Transparent
         Caption         =   "0"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000080&
         Height          =   195
         Left            =   3285
         TabIndex        =   7
         Top             =   1125
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmMalaDiretaCidadao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim aCodigosImp() As Long, xImovel As clsImovel

Private Sub cmdImportar_Click()
Dim strLinha As String, z As Variant, x As Integer, nCodigo As Long, strCodigos As String
lblTotImp.Caption = 0
If txtDelimiter.text = "" Then
    MsgBox "Especifique um delimitador", vbCritical, "Erro"
    Exit Sub
End If
If txtArq.text = "" Then
    MsgBox "Selecione um arquivo", vbCritical, "Erro"
    Exit Sub
End If

ReDim aCodigosImp(0): strCodigos = ""
Open txtArq.text For Input As #1
   Do While Not EOF(1)
        Line Input #1, strLinha
        z = Split(strLinha, txtDelimiter.text)
        For x = 0 To UBound(z)
            If Not IsNumeric(z(x)) Then
               GoTo proximo
            End If
            nCodigo = CLng(z(x))
'            Sql = "insert protesto(codigo) values(" & nCodigo & ")"
'            cn.Execute Sql, rdExecDirect
            If nCodigo < 500000 Then
               GoTo Erro
            End If
            ReDim Preserve aCodigosImp(UBound(aCodigosImp) + 1)
            aCodigosImp(UBound(aCodigosImp)) = nCodigo
            strCodigos = strCodigos & nCodigo & ","
        Next
proximo:
   Loop
Close #1
strCodigos = Chomp(strCodigos, chomp_righT, 1)
lblTotImp.Caption = UBound(aCodigosImp)

Exit Sub
Erro:
MsgBox "Arquivo inválido !!!", vbCritical, "Erro de importação"
Close #1

End Sub

Private Sub cmdImprimir_Click()

If Val(lblTotImp.Caption) = 0 Then
    MsgBox "Nenhum cidadão selecionado.", vbCritical, "Erro"
    Exit Sub
End If

PrintLabel

End Sub

Private Sub cmdLimpar_Click()
txtArq.text = "": lblTotImp.Caption = 0
End Sub

Private Sub cmdOpen_Click()
Dim fName As String, cc As cCommonDlg

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , , , , "Documento de Texto|*.txt;*.csv|Todos os Arquivos|*.*", , App.Path & "\Bin", "Selecione um arquivo texto", , Me.hwnd, OFN_HIDEREADONLY, False
txtArq.text = fName

End Sub

Private Sub cmdPreview_Click()
If (txtArq.text) <> "" Then
    z = Shell(App.Path & "\NOTEPAD2" & " " & txtArq.text, vbNormalFocus)
End If

End Sub

Private Sub Form_Load()
Centraliza Me
Set xImovel = New clsImovel

End Sub

Private Sub PrintLabel()
Dim x As Integer, nCodReduz As Long, Sql As String, RdoAux As rdoResultset, sNome As String, sEndereco As String
Dim nNum As Integer, sCep As String, sComplemento As String, sBairro As String, sCidade As String, sUF As String
Ocupado

If cGetInputState() <> 0 Then DoEvents

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
    
For x = 0 To UBound(aCodigosImp)
    nCodReduz = aCodigosImp(x)
    If nCodReduz > 0 Then
        Sql = "SELECT CODCIDADAO,NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                sNome = !nomecidadao
                xImovel.RetornaEndereco nCodReduz, cidadao, cadastrocidadao
                sEndereco = xImovel.Endereco
                nNum = xImovel.Numero
                sCep = xImovel.Cep
                sComplemento = xImovel.Complemento
                sBairro = xImovel.Bairro
                sCidade = xImovel.Cidade
                sUF = xImovel.UF
                Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
                Sql = Sql & NomeDeLogin & "'," & x & ",'" & Format(nCodReduz, "000000") & "','" & Mask(sNome) & "','"
                Sql = Sql & sEndereco & " " & nNum & " " & sComplemento & "','" & sBairro & " - " & sCidade & "','" & sUF & " - " & sCep & "')"
                cn.Execute Sql, rdExecDirect
            Else
                MsgBox "Código: " & nCodReduz & " não cadastrado.", vbCritical, "Erro"
            End If
           .Close
        End With
    End If
Next

Liberado

If cGetInputState() <> 0 Then DoEvents
frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub
