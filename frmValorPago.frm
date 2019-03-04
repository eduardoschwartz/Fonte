VERSION 5.00
Object = "{93019C16-6A9D-4E32-A995-8B9C1D41D5FE}#1.0#0"; "prjChameleon.ocx"
Object = "{F48120B2-B059-11D7-BF14-0010B5B69B54}#1.0#0"; "esMaskEdit.ocx"
Begin VB.Form frmValorPago 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Consultar total pago por período"
   ClientHeight    =   2115
   ClientLeft      =   9345
   ClientTop       =   5295
   ClientWidth     =   6135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2115
   ScaleWidth      =   6135
   ShowInTaskbar   =   0   'False
   Begin VB.ComboBox cmbLanc 
      Height          =   315
      Left            =   1305
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   240
      Width           =   4605
   End
   Begin prjChameleon.chameleonButton cmdPrint 
      Height          =   345
      Left            =   4770
      TabIndex        =   3
      ToolTipText     =   "Consultar valor pago"
      Top             =   750
      Width           =   1065
      _ExtentX        =   1879
      _ExtentY        =   609
      BTYPE           =   3
      TX              =   "Consultar"
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
      MICON           =   "frmValorPago.frx":0000
      UMCOL           =   -1  'True
      SOFT            =   0   'False
      PICPOS          =   0
      NGREY           =   0   'False
      FX              =   0
      HAND            =   0   'False
      CHECK           =   0   'False
      VALUE           =   0   'False
   End
   Begin esMaskEdit.esMaskedEdit mskData 
      Height          =   285
      Left            =   1305
      TabIndex        =   1
      Top             =   765
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   503
      MouseIcon       =   "frmValorPago.frx":001C
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
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin esMaskEdit.esMaskedEdit mskData2 
      Height          =   285
      Left            =   3495
      TabIndex        =   2
      Top             =   765
      Width           =   1050
      _ExtentX        =   1852
      _ExtentY        =   503
      MouseIcon       =   "frmValorPago.frx":0038
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
      Mask            =   "99/99/9999"
      SelText         =   ""
      Text            =   "__/__/____"
      HideSelection   =   -1  'True
   End
   Begin VB.Label lblTotal 
      BackStyle       =   0  'Transparent
      Caption         =   "R$ 0,00"
      ForeColor       =   &H00000080&
      Height          =   195
      Left            =   2070
      TabIndex        =   8
      Top             =   1410
      Width           =   1815
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Valor total arrecadado..:"
      Height          =   195
      Index           =   2
      Left            =   240
      TabIndex        =   7
      Top             =   1410
      Width           =   1815
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Lançamento..:"
      Height          =   195
      Index           =   3
      Left            =   150
      TabIndex        =   6
      Top             =   315
      Width           =   1080
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Data inicio:"
      Height          =   195
      Index           =   1
      Left            =   195
      TabIndex        =   5
      Top             =   825
      Width           =   795
   End
   Begin VB.Label lblVenc 
      BackStyle       =   0  'Transparent
      Caption         =   "Data final:"
      Height          =   195
      Index           =   0
      Left            =   2595
      TabIndex        =   4
      Top             =   825
      Width           =   795
   End
End
Attribute VB_Name = "frmValorPago"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmbLanc_Change()
lblTotal.Caption = "R$ 0,00"
End Sub

Private Sub cmdPrint_Click()
Dim RdoAux  As rdoResultset, Sql As String, nTotal As Double

lblTotal.Caption = "R$ 0,00"
nTotal = 0

If Not IsDate(mskData.Text) Then
    MsgBox "Data inicial inválida", vbExclamation, "Atenção"
    Exit Sub
End If

If Not IsDate(mskData2.Text) Then
    MsgBox "Data final inválida", vbExclamation, "Atenção"
    Exit Sub
End If

If CDate(mskData.Text) > CDate(mskData2.Text) Then
    MsgBox "Data inicial maior que data final", vbExclamation, "Atenção"
    Exit Sub
End If

Ocupado
Refresh
DoEvents
Sql = "SELECT SUM(VALORPAGOREAL) AS SOMA FROM DEBITOPAGO WHERE DATARECEBIMENTO BETWEEN '" & Format(mskData.Text, "mm/dd/yyyy") & "' AND '" & Format(mskData2.Text, "mm/dd/yyyy") & "' AND CODLANCAMENTO=" & cmbLanc.ItemData(cmbLanc.ListIndex)
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
nTotal = IIf(IsNull(RdoAux!soma), 0, RdoAux!soma)
RdoAux.Close
Liberado

lblTotal.Caption = "R$ " & FormatNumber(nTotal, 2)

End Sub

Private Sub Form_Load()

Sql = "SELECT CODLANCAMENTO,DESCFULL FROM LANCAMENTO WHERE CODLANCAMENTO<>20 AND CODLANCAMENTO<>11 ORDER BY DESCFULL"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
     Do Until .EOF
          cmbLanc.AddItem !DESCFULL
          cmbLanc.ItemData(cmbLanc.NewIndex) = !CodLancamento
         .MoveNext
     Loop
    .Close
End With

cmbLanc.ListIndex = 0
Centraliza Me
End Sub

Private Sub mskData_GotFocus()
mskData.SelStart = 0
mskData.SelLength = Len(mskData.Text)
End Sub

Private Sub mskData2_GotFocus()
mskData2.SelStart = 0
mskData2.SelLength = Len(mskData2.Text)

End Sub
