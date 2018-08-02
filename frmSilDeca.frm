VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmSilDeca 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Resumo Cadastro no SIL"
   ClientHeight    =   2235
   ClientLeft      =   3030
   ClientTop       =   2205
   ClientWidth     =   5145
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2235
   ScaleWidth      =   5145
   Begin VB.ListBox lstEmpresa 
      Height          =   1035
      Left            =   45
      TabIndex        =   2
      Top             =   360
      Width           =   5010
   End
   Begin prjChameleon.chameleonButton cmdImportar 
      Height          =   645
      Left            =   3555
      TabIndex        =   0
      ToolTipText     =   "Exporta o arquivo de dados do GTI para o ISS Eletrônico"
      Top             =   1485
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1138
      BTYPE           =   14
      TX              =   "Importar XML"
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
      BCOL            =   15790320
      BCOLO           =   15790320
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmSilDeca.frx":0000
      PICN            =   "frmSilDeca.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   3405
      Left            =   270
      TabIndex        =   1
      Top             =   2610
      Width           =   4635
      _ExtentX        =   8176
      _ExtentY        =   6006
      _Version        =   393217
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin prjChameleon.chameleonButton cmdImprimir 
      Height          =   315
      Left            =   90
      TabIndex        =   4
      ToolTipText     =   "Impressão do Protocolo de Entrada e Requerimento"
      Top             =   1800
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Imprimir"
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
      MICON           =   "frmSilDeca.frx":02F4
      PICN            =   "frmSilDeca.frx":0310
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
      Caption         =   "Empresas contidas no arquivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   90
      TabIndex        =   3
      Top             =   135
      Width           =   4740
   End
End
Attribute VB_Name = "frmSilDeca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim sCnaePrincipal As String

Private Sub cmdImportar_Click()
Dim fName As String, cc As cCommonDlg, aName() As String, nFile As Integer, x As Integer, y As Integer
Dim oDoc As MYXMLDOM.XMLDOC, nPos As Integer, CodEmpresa As String

Set cc = New cCommonDlg
cc.VBGetOpenFileName fName, , , True, , , "Documento xml|*.xml", , App.Path & "\Bin", "Selecione um arquivo xml do SIL", , Me.hwnd, OFN_HIDEREADONLY, False
aName = Split(fName, " ")

If UBound(aName) = -1 Then Exit Sub

Set oDoc = New MYXMLDOM.XMLDOC
With oDoc
    .loadXML (load(aName(0)))
End With

LoadDocIntoTree oDoc

Set oDoc = Nothing
lstEmpresa.Clear
    
With tv
    For x = 1 To .Nodes.Count
        If .Nodes(x).Text = "Empresa" Then
            CodEmpresa = .Nodes(x).Tag
            For y = x To x + .Nodes(x).Children
                If .Nodes(y).Text = "NomeEmpresarial" Then
                    lstEmpresa.AddItem .Nodes(y).Tag
                    lstEmpresa.ItemData(lstEmpresa.NewIndex) = CodEmpresa
                End If
            Next
            
        End If
    Next
End With

If lstEmpresa.ListCount > 0 Then lstEmpresa.ListIndex = 0

End Sub

Private Sub Form_Load()
Centraliza Me
End Sub

Public Function load(ByVal FileName As String) As String
    Dim xmlStr As String, xmlLine As String
    Dim freeFileNo As Integer
    freeFileNo = FreeFile
    Open FileName For Input As #freeFileNo
        If Not EOF(1) Then
            Do While Not EOF(freeFileNo)
                Input #freeFileNo, xmlLine
                xmlStr = xmlStr & " " & xmlLine
            Loop
        End If
    Close #freeFileNo
    load = xmlStr
End Function

Private Sub LoadChildNodesIntoTree(oXmlNode As XMLNODE, ParentTvNode As MSComctlLib.Node)
    Dim x As Long
    Dim oChildXmlNode As XMLNODE
    Dim oChildTvNode As MSComctlLib.Node
    
    With oXmlNode
        If .HasChildNodes Then
            For x = 0 To .ChildNodes.Count - 1
                Set oChildXmlNode = .ChildNodes(x)
                
                If oChildXmlNode.NodeType <> NODE_COMMENT And oChildXmlNode.NodeType <> NODE_PROCESSING_INSTRUCTION Then
                    If oChildXmlNode.HasAttributes Then
                        Set oChildTvNode = tv.Nodes.Add(ParentTvNode.Key, tvwChild, NextKey(), oChildXmlNode.NodeName)
                        If oChildXmlNode.NodeName = "Municipio" Then
                            oChildTvNode.Tag = oChildXmlNode.Text
                        Else
                            If oChildXmlNode.NodeName = "CNAE" Then
                                If oChildXmlNode.Attributes(1).Value = "1" Then
                                    sCnaePrincipal = oChildXmlNode.Attributes(0).Value
                                End If
                            End If
                            If sCnaePrincipal <> "" Then
                                oChildTvNode.Tag = oChildXmlNode.Attributes(0).Value & " (Principal)"
                                sCnaePrincipal = ""
                            Else
                                oChildTvNode.Tag = oChildXmlNode.Attributes(0).Value
                            End If
                        End If
                    Else
                        Set oChildTvNode = tv.Nodes.Add(ParentTvNode.Key, tvwChild, NextKey(), oChildXmlNode.NodeName)
                        oChildTvNode.Tag = oChildXmlNode.Text
                    End If
                    oChildTvNode.EnsureVisible
                    
                    If oChildXmlNode.HasChildNodes Then
                        LoadChildNodesIntoTree oChildXmlNode, oChildTvNode
                    End If
                End If
                
                Set oChildXmlNode = Nothing
                Set oChildTvNode = Nothing
            Next
        End If
    End With
    
End Sub

Private Function NextKey() As String
    Static KeyCount As Integer
    
    KeyCount = KeyCount + 1
    NextKey = "x" & KeyCount
    
End Function

Private Sub LoadDocIntoTree(oDoc As MYXMLDOM.XMLDOC)
    Dim oTvNode As MSComctlLib.Node
    Dim oXmlNode As MYXMLDOM.XMLNODE
    Dim x As Long
    
    tv.Nodes.Clear

    With oDoc
        'grab the top node - there can only be one
        If .HasChildNodes Then
            For x = 0 To .ChildNodes.Count - 1
                Set oXmlNode = .ChildNodes(x)
                
                If oXmlNode.NodeType <> NODE_COMMENT And oXmlNode.NodeType <> NODE_PROCESSING_INSTRUCTION Then
                    Set oTvNode = tv.Nodes.Add(, , NextKey(), oXmlNode.NodeName)
                    oTvNode.EnsureVisible
                    
                    If oXmlNode.HasChildNodes Then
                        LoadChildNodesIntoTree oXmlNode, oTvNode
                    End If
                End If
            
            Next
        End If
    End With
    
    'make sure the top node is visible
    If tv.Nodes.Count > 0 Then tv.Nodes(1).EnsureVisible
    
    Set oTvNode = Nothing
    Set oXmlNode = Nothing
End Sub

Private Sub cmdImprimir_Click()
Dim FF1 As Integer, sNomeArq As String, ax As String, ret As Long, x As Integer, y As Integer, CodEmpresa As String, RazaoSocial As String
Dim z As Integer, bSocio As Boolean, bImovel As Boolean

bSocio = False: bImovel = False
If lstEmpresa.ListIndex = -1 Then
    MsgBox "Selecione uma empresa.", vbExclamation, "Atenção"
    Exit Sub
Else
    RazaoSocial = lstEmpresa.Text
    CodEmpresa = lstEmpresa.ItemData(lstEmpresa.ListIndex)
End If

sNomeArq = sPathBin & "\RELSIL.TXT"
FF1 = FreeFile()
Open sNomeArq For Output As FF1

ax = "RESUMO DO CADASTRO RECEBIDO PELO SIL" & vbCrLf
ax = ax & "************************************" & vbCrLf & vbCrLf
ax = ax & "ID............: " & CodEmpresa & vbCrLf
ax = ax & "RAZÃO SOCIAL..: " & RazaoSocial & vbCrLf

On Error Resume Next
With tv
    For x = 1 To .Nodes.Count
        If .Nodes(x).Text = "Empresas" Then GoTo PROXIMO
        If (.Nodes(x).Text = "Empresa" And .Nodes(x).Tag = CodEmpresa) Or (.Nodes(x).Parent.Text = "Empresa" And .Nodes(x).Parent.Tag = CodEmpresa) Or (.Nodes(x).Parent.Parent.Text = "Empresa" And .Nodes(x).Parent.Parent.Tag = CodEmpresa) Then
            If .Nodes(x).Text = "CNPJ" Then
                ax = ax & "CNPJ..........: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "DataAbertura" Then
                ax = ax & "DATA ABERTURA.: " & Format(.Nodes(x).Tag, "dd/mm/yyyy") & vbCrLf
            ElseIf .Nodes(x).Text = "NomeResponsavel" Then
                ax = ax & "NOME RESPONSAV: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "CPFResponsavel" Then
                ax = ax & "CPF RESPONSAV.: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "DDD1" Then
                ax = ax & "DDD...........: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "Telefone1" Then
                ax = ax & "Telefone......: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "Email" Then
                ax = ax & "Email.........: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "Endereco" Then
                For z = x + 1 To x + .Nodes(x).Children
                    If .Nodes(z).Text = "TipoLogradouro" Then
                        ax = ax & "Tipo Lograd...: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "Logradouro" Then
                        ax = ax & "Logradouro....: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "NumeroLogradouro" Then
                        ax = ax & "Nº Lograd.....: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "Bairro" Then
                        ax = ax & "Bairro........: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "CEP" Then
                        ax = ax & "CEP...........: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "UF" Then
                        ax = ax & "Sigla UF......: " & .Nodes(z).Tag & vbCrLf
                    ElseIf Left(.Nodes(z).Text, 9) = "Municipio" Then
                        ax = ax & "Município.....: " & .Nodes(z).Tag & vbCrLf
                    End If
                Next
            ElseIf .Nodes(x).Text = "Atividades" Then
                For z = x To x + .Nodes(x).Children
                    If .Nodes(z).Text = "CNAE" Then
                        ax = ax & "CNAE..........: " & .Nodes(z).Tag & vbCrLf
                    End If
                Next
            ElseIf .Nodes(x).Text = "NumeroCRCContadorPJ" Then
                ax = ax & "CRC.Contador.J: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "CNPJContador" Then
                ax = ax & "CNPJ Contador.: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "NumeroCRCContadorPF" Then
                ax = ax & "CRC.Contador.F: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "CPFContador" Then
                ax = ax & "CPF Contador..: " & .Nodes(x).Tag & vbCrLf
            ElseIf .Nodes(x).Text = "Imovel" Then
                If Not bImovel Then
                    ax = ax & vbCrLf
                    ax = ax & "IMóvel:" & vbCrLf
                    ax = ax & "======" & vbCrLf
                    ax = ax & vbCrLf
                    bSocio = True
                End If
                For z = x To x + .Nodes(x).Children
                    If .Nodes(z).Text = "AreaEstabelecimento" Then
                        ax = ax & "Area Estabel..: " & .Nodes(z).Tag & "m²" & vbCrLf
                    ElseIf .Nodes(z).Text = "NomeProprietario" Then
                        ax = ax & "Nome Propriet.: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "EmailProprietario" Then
                        ax = ax & "Email Propriet: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "TelefoneProprietario" Then
                        ax = ax & "Tel. Propriet.: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "NomeResponsavelUso" Then
                        ax = ax & "Nome Resp Uso.: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "TelefoneResponsavelUso" Then
                        ax = ax & "Tel. Resp Uso.: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "AreaTotal" Then
                        ax = ax & "Areas Total...: " & .Nodes(z).Tag & "m²" & vbCrLf
                    ElseIf .Nodes(z).Text = "Pavimentos" Then
                        ax = ax & "Pavimentos....: " & .Nodes(z).Tag & vbCrLf
                    End If
                Next
            ElseIf .Nodes(x).Text = "Socio" Then
                If Not bSocio Then
                    ax = ax & vbCrLf
                    ax = ax & "Sócios:" & vbCrLf
                    ax = ax & "======" & vbCrLf
                    ax = ax & vbCrLf
                    bSocio = True
                End If
                For z = x To x + .Nodes(x).Children
                    If .Nodes(z).Text = "Nome" Then
                        ax = ax & "Nome do Socio.: " & .Nodes(z).Tag & vbCrLf
                    ElseIf .Nodes(z).Text = "Numero" Then
                        ax = ax & "CPF do Socio..: " & .Nodes(z).Tag & vbCrLf
                    End If
                Next
            End If
        End If
PROXIMO:
    Next
End With

Print #FF1, ax

Close #FF1

ret = Shell("NOTEPAD" & " " & sNomeArq, vbNormalFocus)

End Sub

