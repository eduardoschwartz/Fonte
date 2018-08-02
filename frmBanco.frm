VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "Msflxgrd.ocx"
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmBanco 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Cadastro de Bancos / Conta da Prefeitura"
   ClientHeight    =   3450
   ClientLeft      =   2205
   ClientTop       =   1995
   ClientWidth     =   6765
   Icon            =   "frmBanco.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3450
   ScaleWidth      =   6765
   Begin prjChameleon.chameleonButton cmdSair 
      Height          =   315
      Left            =   5670
      TabIndex        =   1
      ToolTipText     =   "Sair da Tela"
      Top             =   3060
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
      COLTYPE         =   2
      FOCUSR          =   0   'False
      BCOL            =   12632256
      BCOLO           =   12632256
      FCOL            =   0
      FCOLO           =   0
      MCOL            =   12632256
      MPTR            =   1
      MICON           =   "frmBanco.frx":030A
      PICN            =   "frmBanco.frx":0326
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin MSFlexGridLib.MSFlexGrid grdBanco 
      Height          =   2955
      Left            =   30
      TabIndex        =   0
      Top             =   30
      Width           =   6705
      _ExtentX        =   11827
      _ExtentY        =   5212
      _Version        =   393216
      Rows            =   1
      Cols            =   5
      FixedCols       =   0
      BackColorFixed  =   15658734
      BackColorSel    =   128
      FocusRect       =   0
      SelectionMode   =   1
      Appearance      =   0
      FormatString    =   "Código    |Nome do Banco                               |Agência    |Nº da Conta                 |DV   "
   End
End
Attribute VB_Name = "frmBanco"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdSair_Click()
Unload Me
End Sub

Private Sub Form_Load()
Ocupado
Centraliza Me
CarregaLista
Liberado
End Sub

Private Sub CarregaLista()
Dim RdoAux As rdoResultset, Sql As String

Sql = "SELECT CODBANCO,NOMEBANCO,AGCONTAPREF,NUCONTAPREF,DVCONTAPREF FROM BANCO WHERE CODBANCO<>0"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
         grdBanco.AddItem !CodBanco & Chr(9) & !NomeBanco & Chr(9) & !AGCONTAPREF & Chr(9) & !NUCONTAPREF & Chr(9) & !DVCONTAPREF
        .MoveNext
    Loop
End With

End Sub
