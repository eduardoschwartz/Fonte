VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "Richtx32.ocx"
Begin VB.Form frmBug 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Histórico de correções e melhorias no GTI"
   ClientHeight    =   3810
   ClientLeft      =   2505
   ClientTop       =   2400
   ClientWidth     =   5805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3810
   ScaleWidth      =   5805
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox Rtb 
      Height          =   3795
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5775
      _ExtentX        =   10186
      _ExtentY        =   6694
      _Version        =   393217
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmBug.frx":0000
   End
End
Attribute VB_Name = "frmBug"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RdoAux As rdoResultset, Sql As String

Public Sub CarregaBug()
Dim sVersao As String

Sql = "SELECT * FROM bugs  ORDER BY MAJOR,MINOR,REVISION DESC"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)

Do Until RdoAux.EOF
 '   If nMajor <> RdoAux!Major And nMinor <> RdoAux!Minor And nRevision <> RdoAux!Revision Then
        sVersao = RdoAux!Major & "." & RdoAux!Minor & "." & RdoAux!Revision
        With Rtb
            '.SelUnderline = True
             Negrito
            .SelText = "Versão nº " & sVersao & " Seq: " & RdoAux!Seq & " (" & Format(RdoAux!Data, "dd/mm/yyyy hh:mm;ss") & ") - " & SubNull(RdoAux!SOLICITADO) & vbCrLf
            Normal
            .SelText = RdoAux!Texto & vbCrLf & vbCrLf
            ' Negrito
            '.SelText = "Inscrição Cadastral: ":     Normal
        End With
        
'    End If
    RdoAux.MoveNext
Loop
Rtb.SelStart = 0
End Sub

Private Sub Negrito()
Rtb.SelBold = True
End Sub

Private Sub Normal()
Rtb.SelBold = False
End Sub

