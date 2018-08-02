VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmSalaEmpreendedor 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Sala do Empreendedor"
   ClientHeight    =   1635
   ClientLeft      =   4860
   ClientTop       =   5430
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   1635
   ScaleWidth      =   5025
   Begin esMaskEdit.esMaskedEdit mskDataIni 
      Height          =   285
      Left            =   1320
      TabIndex        =   0
      Top             =   570
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      MouseIcon       =   "frmSalaEmpreendedor.frx":0000
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
      Mask            =   "##/##/####"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskDataFim 
      Height          =   285
      Left            =   3720
      TabIndex        =   1
      Top             =   570
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   503
      MouseIcon       =   "frmSalaEmpreendedor.frx":001C
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
      Mask            =   "##/##/####"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Default         =   -1  'True
      Height          =   345
      Left            =   1890
      TabIndex        =   3
      ToolTipText     =   "Imprimir Relatório"
      Top             =   1125
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Relatório"
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
      MICON           =   "frmSalaEmpreendedor.frx":0038
      PICN            =   "frmSalaEmpreendedor.frx":0054
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
      Cancel          =   -1  'True
      Height          =   345
      Left            =   3210
      TabIndex        =   4
      ToolTipText     =   "Sair da Tela"
      Top             =   1125
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Sair"
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
      MICON           =   "frmSalaEmpreendedor.frx":01AE
      PICN            =   "frmSalaEmpreendedor.frx":01CA
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdEtiqueta 
      Height          =   375
      Left            =   570
      TabIndex        =   2
      ToolTipText     =   "Imprimir Etiqueta"
      Top             =   1110
      Width           =   1245
      _ExtentX        =   2196
      _ExtentY        =   661
      BTYPE           =   3
      TX              =   "Etiquetas"
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
      FOCUSR          =   -1  'True
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSalaEmpreendedor.frx":0238
      PICN            =   "frmSalaEmpreendedor.frx":0254
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
      Caption         =   "Digite a data de abertura inicial e final das empresas."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   180
      TabIndex        =   7
      Top             =   120
      Width           =   5565
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Final..:"
      Height          =   225
      Index           =   0
      Left            =   2730
      TabIndex        =   6
      Top             =   600
      Width           =   945
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Inicial..:"
      Height          =   225
      Index           =   9
      Left            =   270
      TabIndex        =   5
      Top             =   600
      Width           =   945
   End
End
Attribute VB_Name = "frmSalaEmpreendedor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdEtiqueta_Click()
If Not Valida Then Exit Sub

Ocupado
Sql = "DELETE FROM ETIQUETAGTI where usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT CODIGOMOB FROM vwFULLEMPRESA Where DATAABERTURA BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
        Sql = "SELECT * FROM vwFULLEMPRESA Where CODIGOMOB=" & RdoAux2!codigomob
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nCodLogr = !CodLogradouro
            sCodInscricao = Format(!codigomob, "000000")
            sContribuinte = !razaosocial
            sEnd = SubNull(!Logradouro) & " Nº " & CStr(!Numero)
            sCEP = RetornaCEP(!CodLogradouro, !Numero)
            sCompl = SubNull(Left(!Complemento, 20))
            sBairro = SubNull(!DescBairro)
    
            sEndEntrega = sEnd
            sBairroEntrega = sBairro
            sCidEntrega = !desccidade
            sCepEntrega = sCEP
            sComplEntrega = sCompl
            sUFEntrega = !DESCUF
            Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
            Sql = Sql & NomeDeLogin & "'," & 0 & ",'" & sCodInscricao & "','" & Mask(!razaosocial) & "','"
            Sql = Sql & sEndEntrega & " " & sComplEntrega & "','" & sCepEntrega & "   " & sBairroEntrega & "','" & sCidEntrega & "   " & sUFEntrega & "')"
            cn.Execute Sql, rdExecDirect
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport "ETIQUETACONSIST", frmMdi.hwnd, Me.hwnd
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

End Sub

Private Sub cmdPrint_Click()
Dim RdoAux As rdoResultset, RdoAux2 As rdoResultset, RdoAux3 As rdoResultset, Sql As String
Dim nCodLogr As Long, sCodInscricao As String, sContribuinte As String, sNomeEsc As String
Dim sEnd As String, nNum As Integer, sCEP As String, sCompl As String, sBairro As String

If Not Valida Then Exit Sub
Ocupado
Sql = "DELETE FROM SALAEMPREENDEDOR"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT CODIGOMOB FROM vwFULLEMPRESA Where DATAABERTURA BETWEEN '" & Format(mskDataIni.Text, "mm/dd/yyyy") & "' and '" & Format(mskDataFim.Text, "mm/dd/yyyy") & "'"
Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux2
    Do Until .EOF
        Sql = "SELECT * FROM vwFULLEMPRESA Where CODIGOMOB=" & RdoAux2!codigomob
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            nCodLogr = !CodLogradouro
            sEnd = SubNull(!Logradouro) & " Nº " & CStr(!Numero)
            If Val(SubNull(!RESPCONTABIL)) > 0 Then
                Sql = "SELECT NOMEESC FROM ESCRITORIOCONTABIL WHERE CODIGOESC=" & Val(SubNull(!RESPCONTABIL))
                Set RdoAux3 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux3.RowCount > 0 Then
                    sNomeEsc = RdoAux3!NOMEESC
                Else
                    sNomeEsc = ""
                End If
                RdoAux3.Close
            Else
                sNomeEsc = ""
            End If
            Sql = "INSERT SALAEMPREENDEDOR (CODIGOMOB,RAZAOSOCIAL,ENDEREÇO,CNPJ,IE,DATAAB,ENCERRADO,ATIVIDADE,TELEFONE,CONTADOR) VALUES("
            Sql = Sql & !codigomob & ",'" & Mask(!razaosocial) & "','"
            Sql = Sql & sEnd & "','" & SubNull(!Cnpj) & "','" & SubNull(!INSCESTADUAL) & "','" & Format(!DATAABERTURA, "mm/dd/yyyy") & "','" & !ENCERRADO & "','" & !Atividade & "','" & SubNull(!FONECONTATO) & "','" & sNomeEsc & "')"
            cn.Execute Sql, rdExecDirect
           .Close
        End With
       .MoveNext
    Loop
   .Close
End With

frmReport.ShowReport "SALAEMPREENDEDOR", frmMdi.hwnd, Me.hwnd

Sql = "DELETE FROM SALAEMPREENDEDOR"
cn.Execute Sql, rdExecDirect

Liberado
End Sub

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Private Function Valida() As Boolean

Valida = False
If Not IsDate(mskDataIni.Text) Or Not IsDate(mskDataFim.Text) Then
    MsgBox "Data inicial e/ou final inválida.", vbCritical, "ERRO"
    Exit Function
End If

If CDate(mskDataIni.Text) > CDate(mskDataFim.Text) Then
    MsgBox "Data inicial maior que data final.", vbCritical, "ERRO"
    Exit Function
End If
Valida = True

End Function
