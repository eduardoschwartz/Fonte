VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "Mscomct2.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmAlvaraNovo 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Alvará de Funcionamento"
   ClientHeight    =   3420
   ClientLeft      =   15945
   ClientTop       =   5475
   ClientWidth     =   6255
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3420
   ScaleWidth      =   6255
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Height          =   3375
      Left            =   30
      TabIndex        =   5
      Top             =   0
      Width           =   6195
      Begin VB.TextBox txtNumProc 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3510
         MaxLength       =   12
         TabIndex        =   1
         Top             =   180
         Width           =   1155
      End
      Begin MSComCtl2.DTPicker dtData 
         Height          =   315
         Left            =   3000
         TabIndex        =   3
         Top             =   2910
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   126812161
         CurrentDate     =   43493
      End
      Begin VB.CheckBox chkProvisorio 
         BackColor       =   &H00EEEEEE&
         Caption         =   "Alvará provisório"
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
         Left            =   150
         TabIndex        =   2
         Top             =   2970
         Width           =   1815
      End
      Begin VB.TextBox txtCodigo 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1110
         MaxLength       =   6
         TabIndex        =   0
         Top             =   180
         Width           =   915
      End
      Begin prjChameleon.chameleonButton btPrint 
         Cancel          =   -1  'True
         Height          =   345
         Left            =   4740
         TabIndex        =   4
         ToolTipText     =   "Imprimir o Alvará"
         Top             =   2850
         Width           =   1245
         _ExtentX        =   2196
         _ExtentY        =   609
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
         MICON           =   "frmAlvaraNovo.frx":0000
         PICN            =   "frmAlvaraNovo.frx":001C
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
         Caption         =   "Nº Processo..:"
         Height          =   225
         Index           =   1
         Left            =   2430
         TabIndex        =   28
         Top             =   240
         Width           =   1035
      End
      Begin VB.Label lblHorario 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   27
         Top             =   2550
         Width           =   4875
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Horário.......:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   26
         Top             =   2550
         Width           =   945
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Nome.........:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   25
         Top             =   540
         Width           =   885
      End
      Begin VB.Label lblNome 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   24
         Top             =   540
         Width           =   4905
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Bairro.........:"
         Height          =   225
         Left            =   180
         TabIndex        =   23
         Top             =   1680
         Width           =   945
      End
      Begin VB.Label lblBairro 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   22
         Top             =   1680
         Width           =   2625
      End
      Begin VB.Label Label9 
         BackStyle       =   0  'Transparent
         Caption         =   "CEP....:"
         Height          =   225
         Left            =   3930
         TabIndex        =   21
         Top             =   1725
         Width           =   585
      End
      Begin VB.Label lblCEP 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   4590
         TabIndex        =   20
         Top             =   1725
         Width           =   1335
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Atividade....:"
         Height          =   225
         Left            =   180
         TabIndex        =   19
         Top             =   1965
         Width           =   1005
      End
      Begin VB.Label lblAtividade 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   18
         Top             =   1965
         Width           =   4815
      End
      Begin VB.Label lblCidade 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   17
         Top             =   2250
         Width           =   4875
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Cidade.......:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   16
         Top             =   2250
         Width           =   1005
      End
      Begin VB.Label lblValidade 
         BackStyle       =   0  'Transparent
         Caption         =   "Validade..:"
         Height          =   195
         Left            =   2160
         TabIndex        =   15
         Top             =   2970
         Width           =   780
      End
      Begin VB.Label lblCPF 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   14
         Top             =   825
         Width           =   1995
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Endereço...:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   13
         Top             =   1110
         Width           =   945
      End
      Begin VB.Label lblEndereco 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   12
         Top             =   1110
         Width           =   4875
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº...:"
         Height          =   225
         Left            =   2625
         TabIndex        =   11
         Top             =   1440
         Width           =   495
      End
      Begin VB.Label lblNum 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   3105
         TabIndex        =   10
         Top             =   1440
         Width           =   765
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Completo....:"
         Height          =   225
         Left            =   180
         TabIndex        =   9
         Top             =   1395
         Width           =   870
      End
      Begin VB.Label lblCompl 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00800000&
         Height          =   225
         Left            =   1110
         TabIndex        =   8
         Top             =   1395
         Width           =   765
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "CPF/CNPJ.:"
         Height          =   225
         Index           =   1
         Left            =   180
         TabIndex        =   7
         Top             =   825
         Width           =   1005
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Código.......:"
         Height          =   225
         Index           =   0
         Left            =   180
         TabIndex        =   6
         Top             =   240
         Width           =   885
      End
   End
End
Attribute VB_Name = "frmAlvaraNovo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sControle As String, sValidade As String, sEndereco As String

Private Sub btPrint_Click()
Dim nNumAlvara As Long, Sql As String, RdoAux As rdoResultset

Sql = "select max(numero) as maximo from alvara_funcionamento where ano=" & Year(Now)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    If RdoAux!maximo = Null Then
        nNumAlvara = 1
    Else
        nNumAlvara = !maximo + 1
    End If
    .Close
End With
sEndereco = lblEndereco & ", " & lblNum.Caption & " " & lblCompl.Caption

If lblNome.Caption = "" Then
    MsgBox "Selecione uma empresa.", vbCritical, "Atenção"
Else
    If Trim(txtNumProc.Text) = "" Then
        MsgBox "Digite o número do processo.", vbCritical, "Atenção"
    Else
        If chkProvisorio.value = vbChecked Then
            sValidade = Format(dtData.value, "dd/mm/yyyy")
            sControle = Format(nNumAlvara, "00000") & Format(Year(Now), "0000") & "/" & Format(Val(txtCodigo.Text), "000000") & "-AP"
        Else
            sValidade = ""
            sControle = Format(nNumAlvara, "00000") & Format(Year(Now), "0000") & "/" & Format(Val(txtCodigo.Text), "000000") & "-AN"
        End If
        
        Sql = " insert alvara_funcionamento(ano,numero,controle,codigo,razao_social,documento,endereco,bairro,atividade,horario,validade,data_gravada) values("
        Sql = Sql & Year(Now) & "," & nNumAlvara & ",'" & sControle & "'," & Val(txtCodigo.Text) & ",'" & Mask(lblNome.Caption) & "','" & lblCPF.Caption & "','" & Mask(sEndereco) & "','" & Mask(lblBairro.Caption) & "','" & lblAtividade.Caption & "','" & lblHorario.Caption & "','" & Format(sValidade, "mm/dd/yyyy") & "','" & Format(Now, "mm/dd/yyyy hh:mm:ss") & "')"
        cn.Execute Sql, rdExecDirect
    
    
        If chkProvisorio.value = vbUnchecked Then
            frmReport.ShowReport3 "ALVARAFUNCIONAMENTO", frmMdi.HWND, Me.HWND
        Else
            frmReport.ShowReport3 "ALVARAFUNCIONAMENTOPROVISORIO", frmMdi.HWND, Me.HWND
        End If
    End If
End If

End Sub

Private Sub chkProvisorio_Click()
If chkProvisorio.value = vbChecked Then
    dtData.Visible = True
    lblValidade.Visible = True
Else
    dtData.Visible = False
    lblValidade.Visible = False
End If
End Sub

Private Sub Form_Load()
Centraliza Me
Limpa
End Sub

Private Sub Limpa()
lblNome.Caption = ""
lblCPF.Caption = ""
lblEndereco.Caption = ""
lblCompl.Caption = ""
lblNum.Caption = ""
lblBairro.Caption = ""
lblCEP.Caption = ""
lblAtividade.Caption = ""
lblCidade.Caption = ""
lblHorario.Caption = ""
chkProvisorio.value = vbUnchecked
dtData.value = DateAdd("d", 30, Now)
dtData.Visible = False
lblValidade.Visible = False

End Sub

Private Sub txtCodigo_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyReturn Then
    KeyAscii = 0
    txtCodigo_LostFocus
Else
    Tweak txtCodigo, KeyAscii, IntegerPositive
End If

End Sub

Private Sub txtCodigo_LostFocus()
Dim nCodigo As Long, RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset
Dim nDoc As Double, sCep As String, sHor As String


nCodigo = Val(txtCodigo.Text)
If nCodigo = 0 Then
    Exit Sub
Else
    If nCodigo < 100000 Or nCodigo >= 300000 Then
        MsgBox "Inscrição municipal inválida.", vbCritical, "Erro"
    Else
        Limpa
        Sql = "select * from vwfullempresa where codigomob=" & nCodigo
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux
            If .RowCount > 0 Then
                If Not IsNull(!dataencerramento) Then
                    MsgBox "Inscrição municipal encerrada.", vbCritical, "Erro"
                Else
                    Sql = "SELECT * From vwMOBILIARIOSUSPENSO Where codmobiliario=" & nCodigo & " and codtipoevento=2"
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If RdoAux2.RowCount > 0 Then
                        MsgBox "Inscrição municipal suspensa.", vbCritical, "Erro"
                    Else
                        lblNome.Caption = !razaosocial
                        nDoc = Val(RetornaNumero(SubNull(!Cnpj)))
                        If nDoc > 0 Then
                            lblCPF.Caption = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
                        Else
                            nDoc = Val(RetornaNumero(SubNull(!CPF)))
                            If nDoc > 0 Then
                                lblCPF.Caption = Format(Trim(!CPF), "000\.000\.000-00")
                            End If
                        End If
                            lblEndereco.Caption = SubNull(!Logradouro)
                            lblCompl.Caption = SubNull(!Complemento)
                            lblNum.Caption = SubNull(!Numero)
                            lblBairro.Caption = SubNull(!DescBairro)
                            lblCidade.Caption = SubNull(!descCidade) & " - " & SubNull(!SiglaUF)
                            sCep = RetornaCEP(Val(SubNull(!CodLogradouro)), Val(SubNull(!Numero)))
                            lblCEP.Caption = Format(sCep, "00000-000")
                            lblAtividade.Caption = SubNull(!ativextenso)
                        End If
                                                                                           
                        sHor = Mask(SubNull(!HORARIOEXT))
                        If sHor = "" Then
'                            ql = "SELECT DESCHORARIO FROM HORARIOFUNC WHERE CODHORARIO=" & !Horario
'                            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
'                            With RdoAux2
 '                               If .RowCount > 0 Then
                                    sHor = SubNull(!DESCHORARIO)
  '                              End If
   '                            .Close
    '                        End With
                        End If
                        lblHorario.Caption = sHor
                    End If
                    
            Else
                MsgBox "Inscrição municipal não cadastrada.", vbCritical, "Erro"
            End If
           .Close
        End With
    End If
End If

End Sub
