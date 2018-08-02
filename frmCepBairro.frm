VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmCepBairro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Correção de Bairro"
   ClientHeight    =   4995
   ClientLeft      =   13890
   ClientTop       =   5805
   ClientWidth     =   7680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4995
   ScaleWidth      =   7680
   Begin VB.Frame fr1 
      Height          =   4905
      Left            =   30
      TabIndex        =   0
      Top             =   60
      Width           =   7605
      Begin MSFlexGridLib.MSFlexGrid grdMain 
         Height          =   1455
         Left            =   120
         TabIndex        =   6
         Top             =   1980
         Width           =   7365
         _ExtentX        =   12991
         _ExtentY        =   2566
         _Version        =   393216
         Cols            =   4
         FixedCols       =   0
         FormatString    =   ">Código  |^Dist.Setor.Quadra  |Logradouro                                              |Bairro                                "
      End
      Begin VB.TextBox txtCod 
         Height          =   945
         Left            =   150
         MultiLine       =   -1  'True
         TabIndex        =   5
         Top             =   450
         Width           =   7335
      End
      Begin VB.ComboBox cmbBairro 
         Height          =   315
         Left            =   1530
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   3930
         Width           =   5865
      End
      Begin prjChameleon.chameleonButton cmdGravar 
         Height          =   375
         Left            =   6300
         TabIndex        =   3
         ToolTipText     =   "Gravar os Dados"
         Top             =   4350
         Width           =   1125
         _ExtentX        =   1984
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "&Gravar"
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
         MICON           =   "frmCepBairro.frx":0000
         PICN            =   "frmCepBairro.frx":001C
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin prjChameleon.chameleonButton cmdFiltrar 
         Default         =   -1  'True
         Height          =   375
         Left            =   5820
         TabIndex        =   7
         ToolTipText     =   "Consulta processos baseados no filtro selecionado"
         Top             =   1470
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         BTYPE           =   3
         TX              =   "Verificar"
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
         MICON           =   "frmCepBairro.frx":03C1
         PICN            =   "frmCepBairro.frx":03DD
         UMCOL           =   -1  'True
         SOFT            =   0   'False
         PICPOS          =   0
         NGREY           =   0   'False
         FX              =   0
         HAND            =   0   'False
         CHECK           =   0   'False
         VALUE           =   0   'False
      End
      Begin VB.Label Label3 
         Caption         =   "Selecione o novo bairro"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   150
         TabIndex        =   8
         Top             =   3600
         Width           =   7035
      End
      Begin VB.Label Label2 
         Caption         =   "Digite os códigos separados por vírgula e clique em verificar"
         ForeColor       =   &H000000C0&
         Height          =   225
         Left            =   180
         TabIndex        =   4
         Top             =   210
         Width           =   6075
      End
      Begin VB.Label Label1 
         Caption         =   "Novo bairro...:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   180
         TabIndex        =   1
         Top             =   3990
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmCepBairro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdFiltrar_Click()
Dim aCod() As String, x As Integer
aCod = Split(txtCod.Text, ",")
grdMain.Rows = 1

For x = 0 To UBound(aCod)
    grdMain.AddItem aCod(x)
Next


cmbLograduro.AddItem ("(Selecione um logradouro...)")
cmbLograduro.ItemData(cmbLograduro.NewIndex) = 0
Sql = "select codlogradouro,endereco from logradouro order by endereco"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbLograduro.AddItem (!Endereco)
        cmbLograduro.ItemData(cmbLograduro.NewIndex) = !CodLogradouro
       .MoveNext
    Loop
   .Close
End With
cmbLograduro.ListIndex = 0

End Sub

Private Sub cmdGravar_Click()
Dim Sql As String, RdoAux As rdoResultset, RdoAux2 As rdoResultset


End Sub

Private Sub Form_Load()
Dim Sql As String, RdoAux As rdoResultset
Centraliza Me

cmbBairro.AddItem ("(Selecione um bairro...)")
cmbBairro.ItemData(cmbBairro.NewIndex) = 0
Sql = "select * from bairro where siglauf='SP' and codcidade =413 and codbairro<>999 order by descbairro"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
        cmbBairro.AddItem (!DescBairro)
        cmbBairro.ItemData(cmbBairro.NewIndex) = !CodBairro
       .MoveNext
    Loop
   .Close
End With

cmbBairro.ListIndex = 0


End Sub

Private Sub mskCep_Change()

If Len(mskCEP.ClipText) = 8 Then
'    CarregaBairro
Else
    cmbBairro.ListIndex = 0
End If

End Sub

