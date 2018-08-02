VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Begin VB.Form frmEditImob 
   BackColor       =   &H00EEEEEE&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Alteração do Cadastro Imobiliário"
   ClientHeight    =   3195
   ClientLeft      =   4695
   ClientTop       =   2580
   ClientWidth     =   5295
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5295
   ShowInTaskbar   =   0   'False
   Begin prjChameleon.chameleonButton cmdRetorna 
      Height          =   315
      Left            =   3060
      TabIndex        =   4
      ToolTipText     =   "Editar Registro"
      Top             =   2700
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Alterar"
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
      MICON           =   "frmEditImob.frx":0000
      PICN            =   "frmEditImob.frx":001C
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin prjChameleon.chameleonButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   315
      Left            =   4200
      TabIndex        =   5
      ToolTipText     =   "Cancelar Edição"
      Top             =   2700
      Width           =   1035
      _ExtentX        =   1826
      _ExtentY        =   556
      BTYPE           =   3
      TX              =   "&Cancelar"
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
      MICON           =   "frmEditImob.frx":0176
      PICN            =   "frmEditImob.frx":0192
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00EEEEEE&
      Height          =   645
      Left            =   120
      TabIndex        =   2
      Top             =   4695
      Width           =   5295
      Begin VB.TextBox txtDoc 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   315
         Left            =   1770
         TabIndex        =   0
         Top             =   210
         Width           =   3075
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Nº do Documento..:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   270
         Width           =   1455
      End
   End
   Begin VB.ListBox lstItem 
      Appearance      =   0  'Flat
      Height          =   2565
      ItemData        =   "frmEditImob.frx":02EC
      Left            =   0
      List            =   "frmEditImob.frx":02EE
      TabIndex        =   1
      Top             =   15
      Width           =   5295
   End
End
Attribute VB_Name = "frmEditImob"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub cmdCancel_Click()
sItemEdit = ""
Unload Me
End Sub


Private Sub cmdRetorna_Click()

If lstItem.ListIndex = -1 Then
   MsgBox "Selecione o Item a ser alterado.", vbExclamation, "Atenção"
   Exit Sub
End If

Select Case lstItem.ListIndex
    Case 0
        sItemEdit = "PC" 'proprietario/compromissario
    Case 1
        If Val(Left$(frmCadImob.lblCond.Caption, 4)) > 0 Then
            MsgBox "Não é possivel alterar o Local do imóvel de um condominio.", vbExclamation, "Atenção"
            sItemEdit = ""
        Else
            sItemEdit = "LI" 'local do imovel
        End If
    Case 2
        sItemEdit = "EE" 'endereco entrega
    Case 3
        sItemEdit = "AT" 'area do terreno
    Case 4
        sItemEdit = "DT" 'dados do terreno
    Case 5
        If Val(Left$(frmCadImob.lblCond.Caption, 4)) > 0 Then
            MsgBox "Não é possivel alterar as testadas de um condominio.", vbExclamation, "Atenção"
            sItemEdit = ""
        Else
            sItemEdit = "TT" 'testadas
        End If
    Case 6
        sItemEdit = "DC" 'dados construcao
    Case 7
        sItemEdit = "HI" 'historico
End Select
CodImovel = ""

frmCadImob.AlteraCadastro
Unload frmEditImob

End Sub

Private Sub Form_Load()

Centraliza Me

With lstItem
    .AddItem "Proprietário/Proprietário Solidário"
    .AddItem "Local do Imóvel"
    .AddItem "Endereço de Entrega"
    .AddItem "Área do Terreno"
    .AddItem "Dados do Terreno"
    .AddItem "Testadas"
    .AddItem "Dados da Construção"
    .AddItem "Histórico"
End With

End Sub


