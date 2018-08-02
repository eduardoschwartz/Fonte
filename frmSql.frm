VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmSql 
   Caption         =   "Sql"
   ClientHeight    =   6030
   ClientLeft      =   4350
   ClientTop       =   690
   ClientWidth     =   7740
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   NegotiateMenus  =   0   'False
   ScaleHeight     =   6030
   ScaleWidth      =   7740
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox txtLog 
      Height          =   945
      Left            =   30
      TabIndex        =   3
      Top             =   1800
      Width           =   7665
   End
   Begin VB.CommandButton cmdRun 
      Caption         =   "RUN"
      Height          =   345
      Left            =   6750
      TabIndex        =   2
      Top             =   5610
      Width           =   975
   End
   Begin MSFlexGridLib.MSFlexGrid grdTemp 
      Height          =   2715
      Left            =   30
      TabIndex        =   1
      Top             =   2820
      Width           =   7665
      _ExtentX        =   13520
      _ExtentY        =   4789
      _Version        =   393216
      FixedCols       =   0
      AllowUserResizing=   3
      Appearance      =   0
   End
   Begin VB.TextBox txtSql 
      Height          =   1725
      Left            =   30
      MultiLine       =   -1  'True
      TabIndex        =   0
      Top             =   30
      Width           =   7665
   End
End
Attribute VB_Name = "frmSql"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Sql As String, RdoAux As rdoResultset

Private Sub cmdRun_Click()
Dim x As Integer, sTemp As String, bAction As Boolean
On Error GoTo Erro

txtLog.Text = ""
grdTemp.Cols = 1
grdTemp.Rows = 1


Sql = txtSql.Text

If InStr(1, Sql, "INSERT", vbBinaryCompare) > 0 Or _
   InStr(1, Sql, "UPDATE", vbBinaryCompare) > 0 Or _
   InStr(1, Sql, "DELETE", vbBinaryCompare) > 0 Or _
   InStr(1, Sql, "GRANT", vbBinaryCompare) > 0 Or _
   InStr(1, Sql, "REVOKE", vbBinaryCompare) > 0 Then
   bAction = True
Else
   bAction = False
End If

If Not bAction Then
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If .RowCount > 0 Then
            grdTemp.Cols = .rdoColumns.Count
            For x = 0 To .rdoColumns.Count - 1
                grdTemp.TextMatrix(0, x) = .rdoColumns(x).Name
            Next
            
            Do Until .EOF
                sTemp = ""
                For x = 0 To .rdoColumns.Count - 1
                    sTemp = sTemp & .rdoColumns(x).Value & Chr(9)
                Next
                sTemp = Left(sTemp, Len(sTemp) - 1)
                grdTemp.AddItem sTemp
               .MoveNext
            Loop
        End If
    End With
Else
    cn.Execute Sql, rdExecDirect
    txtLog.Text = "Linha afetadas: " & cn.RowsAffected
End If

Exit Sub
Erro:
For x = 0 To rdoErrors.Count - 1
    txtLog.Text = txtLog.Text & vbCrLf & rdoErrors(x).Description
Next

End Sub

Private Sub Form_Load()
Me.Width = 7875
Me.Height = 6510
Centraliza Me
End Sub

Private Sub Form_Resize()
txtSql.Width = Me.Width - 200
'txtSql.Height = Me.Height - grdTemp.Height - txtLog.Height - 200
'txtLog.Top = txtSql.Height + 100
txtLog.Width = Me.Width - 200
'txtLog.Height = Me.Height - grdTemp.Height - txtSql.Height - 200
'grdTemp.Top = txtSql.Height + txtLog.Height + 100
grdTemp.Width = Me.Width - 200
'grdTemp.Height = Me.Height - txtLog.Height - txtSql.Height - 200
cmdRun.Left = Me.Width - Me.Left + cmdRun.Width - 120
End Sub
