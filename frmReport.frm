VERSION 5.00
Object = "{8767A745-088E-4CA6-8594-073D6D2DE57A}#9.2#0"; "crviewer9.dll"
Begin VB.Form frmReport 
   Caption         =   "Report"
   ClientHeight    =   5400
   ClientLeft      =   2085
   ClientTop       =   2640
   ClientWidth     =   11190
   LinkTopic       =   "Form1"
   ScaleHeight     =   5400
   ScaleWidth      =   11190
   Begin CRVIEWER9LibCtl.CRViewer9 CRViewer1 
      Height          =   4695
      Left            =   540
      TabIndex        =   0
      Top             =   180
      Width           =   6495
      lastProp        =   500
      _cx             =   11456
      _cy             =   8281
      DisplayGroupTree=   -1  'True
      DisplayToolbar  =   -1  'True
      EnableGroupTree =   -1  'True
      EnableNavigationControls=   -1  'True
      EnableStopButton=   -1  'True
      EnablePrintButton=   -1  'True
      EnableZoomControl=   -1  'True
      EnableCloseButton=   -1  'True
      EnableProgressControl=   -1  'True
      EnableSearchControl=   -1  'True
      EnableRefreshButton=   -1  'True
      EnableDrillDown =   -1  'True
      EnableAnimationControl=   -1  'True
      EnableSelectExpertButton=   0   'False
      EnableToolbar   =   -1  'True
      DisplayBorder   =   -1  'True
      DisplayTabs     =   -1  'True
      DisplayBackgroundEdge=   -1  'True
      SelectionFormula=   ""
      EnablePopupMenu =   -1  'True
      EnableExportButton=   -1  'True
      EnableSearchExpertButton=   0   'False
      EnableHelpButton=   0   'False
      LaunchHTTPHyperlinksInNewBrowser=   -1  'True
   End
End
Attribute VB_Name = "frmReport"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private hFrmMdi As Long
Private hFrmCall As Long
Dim crApp As New CRAXDRT.Application

Dim rpt As CRAXDRT.Report, bRefisAtivo As Boolean

Private Sub Form_Load()
Dim dDataIni As Date, dDataFim As Date


Sql = "select valparam from parametros where nomeparam='REFIS_INICIO'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataIni = CDate(RdoAux!valparam)

Sql = "select valparam from parametros where nomeparam='REFIS_FIM'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
dDataFim = CDate(RdoAux!valparam)

RdoAux.Close

If Now >= dDataIni And Now <= dDataFim Then
    bRefisAtivo = True
Else
    bRefisAtivo = False
End If

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)

Set crApp = Nothing
Set rpt = Nothing

End Sub

Private Sub Form_Resize()
CRViewer1.Top = 0
CRViewer1.Left = 0
CRViewer1.Height = ScaleHeight
CRViewer1.Width = ScaleWidth
CRViewer1.DisplayGroupTree = False
End Sub

Public Function ShowReport(sReport As String, hMDI As Long, hFormCalling As Long, Optional nNumDoc As Long, Optional nNumGuia As Long)
Dim RdoAux As rdoResultset, Sql As String, sTipo As String, nTotal As Double, sHor As String
Dim sTexto1 As String, sTexto2 As String, sTexto3 As String, bHeader As Boolean, bDam As Boolean
Dim z As Variant, RdoAux2 As rdoResultset, z2 As Variant, z3 As Variant, z4 As Variant, z5 As Variant, z6 As Variant, z7 As Variant
Dim sNumProc As String, nNumproc As Long, nAno As Integer, bAchou As Boolean, aTributo() As String, x As Integer
Dim sNome As String, sEnd As String, sCidade As String, sUF As String, sBairro As String, sRG As String
If bLocal Then
    Exit Function
End If


On Error GoTo Erro
bDam = False
bHeader = False
hFrmMdi = hMDI
hFrmCall = hFormCalling
Ocupado
DoEvents
If IsNull(nNumDoc) Then nNumDoc = 0

If sReport = "MALADIRETAPARC" Then
    MontaMalaDiretaParc
    sReport = "ETIQUETAPROTOCOLO2"
End If
If UCase(sReport) = "CARNE2" Then
    bHeader = True
    sReport = "CARNE"
End If

Set rpt = crApp.OpenReport(sPathReport & "\" & sReport & ".Rpt", 1)
If Left(sReport, 3) = "CDB" Then
    frmReport.Caption = "CERTIDÃO DE DÉBITO"
ASSINATURA:
    If NomeDeLogin = "RENATA" Or NomeDeLogin = "SOLANGE" Or NomeDeLogin = "SCHWARTZ" Then
        z = InputBox("Deseja ocultar a assinatura (S ou N)?", "Assinatura")
        If UCase(z) <> "S" And UCase(z) <> "N" Then GoTo ASSINATURA
    Else
        z = "N"
    End If
    
    Dim m_Report As CRAXDRT.Report
    Dim m_Application As New CRAXDRT.Application
    Set m_Report = Nothing
    Set m_Report = m_Application.OpenReport(sPathReport + "\" & sReport & ".rpt", 1)
    m_Report.EnableParameterPrompting = False
    m_Report.DiscardSavedData
    m_Report.ParameterFields.Item(1).AddCurrentValue CodCidadao
    m_Report.ParameterFields.Item(2).AddCurrentValue NumeroProcesso
    m_Report.ParameterFields.Item(3).AddCurrentValue NomeDeLogin
    m_Report.FormulaFields.GetItemByName("ASSINATURA").Text = "'" & IIf(UCase(z) = "S", "A", "B") & "'"
    m_Report.Database.Tables(1).SetLogOnInfo "192.168.15.160", "Tributacao", UL, UP
    'm_Report.Database.LogOnServer "PDSODBC.DLL", "odbcTributacao", "Tributacao", UL, UP
    m_Report.PaperSize = crPaperA4
   
    With CRViewer1
    
    .EnableExportButton = True
    .EnablePrintButton = True
    .EnableCloseButton = True
    .ReportSource = m_Report ''
    .ViewReport
    Liberado
    frmReport.show 1
    End With
    
    On Error Resume Next
    m_Report.ExportOptions.DestinationType = crEDTDiskFile
    m_Report.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\CERTIDOES\" & UCase(sReport) & "[" & Format(CodCidadao, "000000") & "].PDF"
    m_Report.ExportOptions.FormatType = crEFTPortableDocFormat
    m_Report.ExportOptions.PDFExportAllPages = True
    m_Report.Export (False)
    On Error GoTo 0

    Exit Function
    
End If

Select Case UCase(sReport)
    Case "ANALISE", "ANALISE1", "ANALISE2", "ANALISESD,ANALISE2_TMP"
            frmReport.Caption = "Analise da receita"
            rpt.RecordSelectionFormula = "{analise2.USUARIO}='" & NomeDeLogin & "'"
    Case "AJUIZAMENTO"
            frmReport.Caption = "Ajuizamento de Dívida"
    Case "MMG"
            frmReport.Caption = "Gerador de cartas de correspondência"
            rpt.RecordSelectionFormula = "{MMGREGISTRO.USUARIO}='" & NomeDoUsuario & "'"
    Case "MULTAINF"
            frmReport.Caption = "Multa de Infração"
            rpt.FormulaFields(1).Text = "'" & NumeroProcesso & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frm2ViaLaser.lblProp.Caption) & "'"
            rpt.FormulaFields(3).Text = "'" & Format(frm2ViaLaser.txtCod.Text, "000000") & "-" & RetornaDVCodReduzido(CLng(frm2ViaLaser.txtCod.Text)) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frm2ViaLaser.lblRua.Caption) & ", " & frm2ViaLaser.lblNumImovel.Caption & "'"
            Sql = "SELECT descbairro,dt_areaterreno FROM vwfullimovel2 WHERE CODREDUZIDO=" & Val(frm2ViaLaser.txtCod.Text)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                rpt.FormulaFields(5).Text = "'" & FormatNumber(!Dt_AreaTerreno, 2) & "'"
'                If !Dt_AreaTerreno <= 125 Then
'                    rpt.FormulaFields(6).Text = "'I'"
'                    rpt.FormulaFields(7).Text = "'65,75'"
'                ElseIf !Dt_AreaTerreno > 125 And !Dt_AreaTerreno <= 250 Then
'                    rpt.FormulaFields(6).Text = "'II'"
'                    rpt.FormulaFields(7).Text = "'164,37'"
'                ElseIf !Dt_AreaTerreno > 250 And !Dt_AreaTerreno <= 500 Then
'                    rpt.FormulaFields(6).Text = "'III'"
'                   rpt.FormulaFields(7).Text = "'262,99'"
'                ElseIf !Dt_AreaTerreno > 500 Then
'                    rpt.FormulaFields(6).Text = "'IV'"
'                    rpt.FormulaFields(7).Text = "'394,49'"
'                End If
'                rpt.FormulaFields(8).Text = "'" & FormatNumber(!Dt_AreaTerreno * 0.8113, 2) & "'"
                rpt.FormulaFields(6).Text = "''"
                rpt.FormulaFields(7).Text = "'500,00'"
                rpt.FormulaFields(8).Text = "'" & FormatNumber(!Dt_AreaTerreno * 0.9636, 2) & "'"
                rpt.FormulaFields(10).Text = "'" & SubNull(!DescBairro) & "'"
               .Close
            End With
    Case "CARNE", "CARNETMP"
            frmReport.Caption = "Impressão de Carnê"
            
            If hFrmCall <> frmConfissaoDivida.hwnd Then
                rpt.RecordSelectionFormula = "{CARNETMP.COMPUTER}='" & NomeDoUsuario & "'"
            Else
                rpt.RecordSelectionFormula = "{CARNETMP.COMPUTER}='" & NomeDoUsuario & "' AND {CARNETMP.NUMPARCELA}=1"
            End If
            frmConfissaoDivida.Hide
    Case "GUIAPRATICO1"
            frmReport.Caption = "Guias diversas"
            rpt.FormulaFields(1).Text = "'" & Mask(frmGuiaPratico1.txtNome.Text) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmGuiaPratico1.txtValor.Text) & " (" & Extenso(frmGuiaPratico1.txtValor.Text) & ")'"
            rpt.FormulaFields(3).Text = "'" & Mask(frmGuiaPratico1.txtCod.Text) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmGuiaPratico1.cmbTipo.Text) & "'"
            rpt.FormulaFields(5).Text = "'" & Mask(frmGuiaPratico1.cmbCateg.Text) & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmGuiaPratico1.txtArea.Text) & "'"
            rpt.FormulaFields(7).Text = "'" & Mask(frmGuiaPratico1.txtNot.Text) & "'"
            rpt.FormulaFields(8).Text = "'" & Mask(frmGuiaPratico1.txtProc.Text) & "'"
            rpt.FormulaFields(9).Text = "'" & Mask(frmGuiaPratico1.txtCod2.Text) & "'"
            rpt.FormulaFields(10).Text = "'" & Mask(RetornaUsuarioFullName) & "'"
            If frmGuiaPratico1.cmbPag.ListIndex = 0 Then
                z = "à vista."
            Else
                z = "parcelado em " & frmGuiaPratico1.txtParc.Text & " vezes."
            End If
            rpt.FormulaFields(11).Text = "'" & z & "'"
    Case "GUIAPRATICO2"
            frmReport.Caption = "Guias diversas"
            rpt.FormulaFields(1).Text = "'" & Mask(RetornaUsuarioFullName) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmGuiaPratico2.txtAno.Text) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(frmGuiaPratico2.txtNot.Text) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmGuiaPratico2.txtProc.Text) & "'"
            rpt.FormulaFields(5).Text = "'" & Mask(frmGuiaPratico2.txtNome.Text) & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmGuiaPratico2.txtCod.Text) & "'"
            If frmGuiaPratico2.cmbPag.ListIndex = 0 Then
                z = "à vista."
            Else
                z = "parcelado em " & frmGuiaPratico2.txtParc.Text & " vezes."
            End If
            rpt.FormulaFields(7).Text = "'" & z & "'"
    Case "CALCULOPARCELAMENTO", "CALCULOPARCELAMENTOTMP"
            frmReport.Caption = "Calculo de Parcelamento"
            rpt.RecordSelectionFormula = "{CALCULOPARCELAMENTO.COMPUTER}='" & NomeDeLogin & "'"
            For x = 0 To Forms.Count - 1
                If Forms(x).Name = "frmParcelamento2" Then
                   rpt.FormulaFields(1).Text = "'" & frmConfissaoDivida.lblAno.Caption & "'"
                   Exit For
                ElseIf Forms(x).Name = "frmDebitoImob" Then
                   rpt.FormulaFields(1).Text = "'" & frmDebitoImob.lblAno.Caption & "'"
                   Exit For
                End If
            Next
    Case "DEBITOAJPAGO"
            frmReport.Caption = "Débito ajuizado Pago"
Data1:
            z = InputBox("Digite a data inicial.", "Entre com a informação")
            If z = "" Then GoTo Data1
            If Not IsDate(z) Then GoTo Data1
Data2:
            z2 = InputBox("Digite a data final.", "Entre com a informação")
            If z2 = "" Then GoTo Data2
            If Not IsDate(z2) Then GoTo Data2

            rpt.RecordSelectionFormula = "{vwAnistia.DATARECEBIMENTO}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {vwAnistia.DATARECEBIMENTO}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
            rpt.FormulaFields(1).Text = "'" & Format(z, "dd/mm/yyyy") & " á " & Format(z2, "dd/mm/yyyy") & "'"
    Case "GARE"
            frmReport.Caption = "Impressão de GARE"
            rpt.FormulaFields(1).Text = "'" & frmGare.lblRequerente.Caption & "'"
            rpt.FormulaFields(2).Text = "'" & frmGare.lblEndereco.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & frmGare.lblCidade.Caption & "'"
            rpt.FormulaFields(4).Text = "'" & frmGare.lblUF.Caption & "'"
            rpt.FormulaFields(5).Text = "'" & frmGare.txtVencto.Text & "'"
            rpt.FormulaFields(6).Text = "'" & frmGare.txtValor.Text & "'"
            rpt.FormulaFields(7).Text = "'" & Mask(frmGare.txtNumExec.Text) & "'"
            rpt.FormulaFields(8).Text = "'" & frmGare.lblCPF.Caption & "'"
            rpt.FormulaFields(9).Text = "'" & frmGare.txtCod.Text & "'"
            rpt.FormulaFields(10).Text = "'" & Mask(frmGare.txtExecutado.Text) & "'"
    Case "FUNDOESPDESPESA"
            frmReport.Caption = "Impressão de F.E.D.T.J."
            rpt.FormulaFields(1).Text = "'" & frmFundoDespesa.lblRequerente.Caption & "'"
            rpt.FormulaFields(2).Text = "'" & frmFundoDespesa.lblRG.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & frmFundoDespesa.lblCPF.Caption & "'"
            rpt.FormulaFields(4).Text = "'" & frmFundoDespesa.txtNumExec.Text & "'"
            rpt.FormulaFields(5).Text = "'" & frmFundoDespesa.txtCod.Text & "'"
            rpt.FormulaFields(6).Text = "'" & frmFundoDespesa.txtValor.Text & "'"
            rpt.FormulaFields(7).Text = "'" & Mask(frmFundoDespesa.txtExecutado.Text) & "'"
    Case "DEPOSITOCRI"
            frmReport.Caption = "Depósito CRI"
            z = InputBox("Digite o valor do depósito.", "Valor do Depósito")
            z1 = InputBox("Depositado por:", "Digite o Nome")
            z2 = InputBox("Número da Execução Fiscal", "Digite o Número")
            rpt.FormulaFields(1).Text = "'" & z & "'"
            rpt.FormulaFields(2).Text = "'" & z1 & "'"
            rpt.FormulaFields(3).Text = "'" & z2 & "'"
    Case "DECLARAISENTOIPTU"
            frmReport.Caption = "Declaração de Isento de IPTU"
            rpt.FormulaFields(2).Text = "'" & Mask(frmDeclaraIsento.lblRequerente.Caption) & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmDeclaraIsento.lblRG.Caption) & "'"
            rpt.FormulaFields(7).Text = "'" & frmDeclaraIsento.lblCPF.Caption & "'"
            
            rpt.FormulaFields(1).Text = "'" & frmDeclaraIsento.lblCodImovel.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(frmDeclaraIsento.lblEndereco.Caption) & "'"
            rpt.FormulaFields(4).Text = "'" & frmDeclaraIsento.lblNum.Caption & "'"
            rpt.FormulaFields(5).Text = "'" & Mask(frmDeclaraIsento.lblBairro.Caption) & "'"
            
            If frmDeclaraIsento.chk(0).value = 1 Then
                rpt.FormulaFields(8).Text = "'X'"
            Else
                rpt.FormulaFields(8).Text = "' '"
            End If

            If frmDeclaraIsento.chk(1).value = 1 Then
                rpt.FormulaFields(9).Text = "'X'"
            Else
                rpt.FormulaFields(9).Text = "' '"
            End If
            
            If frmDeclaraIsento.chk(2).value = 1 Then
                rpt.FormulaFields(10).Text = "'X'"
            Else
                rpt.FormulaFields(10).Text = "' '"
            End If
                    
            If frmDeclaraIsento.chk(3).value = 1 Then
                rpt.FormulaFields(11).Text = "'X'"
            Else
                rpt.FormulaFields(11).Text = "' '"
            End If
            
            If Val(frmDeclaraIsento.txtValor(0).Text) > 0 Then
                nTotal = CDbl(frmDeclaraIsento.txtValor(0).Text)
                rpt.FormulaFields(12).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(0).Text, 2) & "'"
            Else
                rpt.FormulaFields(12).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(1).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(1).Text)
                rpt.FormulaFields(13).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(1).Text, 2) & "'"
            Else
                rpt.FormulaFields(13).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(2).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(2).Text)
                rpt.FormulaFields(14).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(2).Text, 2) & "'"
            Else
                rpt.FormulaFields(14).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(3).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(3).Text)
                rpt.FormulaFields(15).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(3).Text, 2) & "'"
            Else
                rpt.FormulaFields(15).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(4).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(4).Text)
                rpt.FormulaFields(16).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(4).Text, 2) & "'"
            Else
                rpt.FormulaFields(16).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(5).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(5).Text)
                rpt.FormulaFields(17).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(5).Text, 2) & "'"
            Else
                rpt.FormulaFields(17).Text = "'R$'"
            End If
            If Val(frmDeclaraIsento.txtValor(6).Text) > 0 Then
                nTotal = nTotal + CDbl(frmDeclaraIsento.txtValor(6).Text)
                rpt.FormulaFields(18).Text = "'R$ " & FormatNumber(frmDeclaraIsento.txtValor(6).Text, 2) & "'"
            Else
                rpt.FormulaFields(18).Text = "'R$'"
            End If
            rpt.FormulaFields(19).Text = "'R$ " & FormatNumber(nTotal, 2) & "'"
                
            
    Case "ETIQUETAPROTOCOLO", "ETIQUETAPROTOCOLO2", "ETIQUETAPROTOCOLO3", "ETIQUETACONSIST", "ETIQUETAGTI", "ETIQUETAIPTU", "ETIQUETACIP"
            frmReport.Caption = "Emissão de Etiquetas"
            rpt.RecordSelectionFormula = "{ETIQUETAGTI.USUARIO}='" & NomeDeLogin & "'"
    Case "REQUERIMENTOPROC"
            rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
            frmReport.Caption = "Requerimento de abertura de processo"
            If frmRequerimento.OptP(0).value = True Then
                sTexto1 = "Eu, " & frmRequerimento.lblRequerente.Caption
                sTexto1 = sTexto1 & " ,portador do RG Nº.: " & frmRequerimento.lblRG.Caption & " e CPF/CNPJ Nº.: " & frmRequerimento.lblCPF.Caption
                sTexto1 = sTexto1 & " , residente/estabelecido à " & frmRequerimento.txtEndereco.Text & " vem mui respeitosamente "
                sTexto1 = sTexto1 & "a presença de V.Exa. solicitar se digne providenciar através do setor competente, o que segue: "
            Else
                sTexto1 = frmRequerimento.lblRequerente.Caption
                sTexto1 = sTexto1 & " ,portador do CNPJ Nº.: " & frmRequerimento.lblCPF.Caption
                sTexto1 = sTexto1 & ", estabelecido à " & frmRequerimento.txtEndereco.Text & " vem mui respeitosamente "
                sTexto1 = sTexto1 & "a presença de V.Exa. solicitar se digne providenciar através do setor competente, o que segue: "
            End If
            rpt.FormulaFields(1).Text = "'" & Mask(sTexto1) & "'"
            rpt.FormulaFields(2).Text = "'as.) " & Mask(frmRequerimento.lblRequerente.Caption) & "'"
    
    Case "CADMOBILIARIO"
            Set m_Report = Nothing
            Set m_Report = m_Application.OpenReport(sPathReport + "\" & sReport & ".rpt", 1)
            m_Report.EnableParameterPrompting = False
            m_Report.ParameterFields.Item(1).AddCurrentValue NomeDeLogin
            m_Report.Database.Tables(1).SetLogOnInfo IPServer, "Tributacao", UL, UP
            m_Report.FormulaFields(8).Text = "'" & frmCadMob.mskCEP.Text & "'"
            m_Report.PaperSize = crPaperA4
            With CRViewer1
                .EnableExportButton = True
                .EnablePrintButton = True
                .EnableCloseButton = True
                .ReportSource = m_Report
                
                .ViewReport
                Liberado
                frmReport.show 1
            End With
        
            Exit Function
           
    Case "PARCELA"
            frmReport.Caption = "Detalhes de um lançamento tributário"
            rpt.RecordSelectionFormula = "{DAM.COMPUTER}='" & NomeDoUsuario & "'"
            rpt.FormulaFields(1).Text = "'" & Mask(frmCnsParcela.lblContrib.Caption) & "'"
            rpt.FormulaFields(2).Text = "'" & frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 1) & "'"
            rpt.FormulaFields(3).Text = "'" & frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 2) & "'"
            rpt.FormulaFields(4).Text = "'" & frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 3) & "'"
            rpt.FormulaFields(5).Text = "'" & frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 4) & "'"
            rpt.FormulaFields(6).Text = "'" & frmDebitoImob.grdExtrato.CellText(frmDebitoImob.grdExtrato.SelectedRow, 5) & "'"
            rpt.FormulaFields(7).Text = "'" & frmCnsParcela.lblStatus.Caption & "'"
            rpt.FormulaFields(8).Text = "'" & frmCnsParcela.lblIsentoMJ.Caption & "'"
            rpt.FormulaFields(9).Text = "'" & frmCnsParcela.lblDesconto.Caption & "'"
            rpt.FormulaFields(10).Text = "'" & frmCnsParcela.txtLivro.Text & "'"
            rpt.FormulaFields(11).Text = "'" & frmCnsParcela.mskIncricao.Text & "'"
            rpt.FormulaFields(12).Text = "'" & frmCnsParcela.txtPagina.Text & "'"
            rpt.FormulaFields(13).Text = "'" & frmCnsParcela.txtCertidao.Text & "'"
            rpt.FormulaFields(14).Text = "'" & frmCnsParcela.lblAjuizamento.Caption & "'"
            rpt.FormulaFields(15).Text = "'" & frmCnsParcela.lblDataBase.Caption & "'"
            rpt.FormulaFields(16).Text = "'" & frmCnsParcela.lblDataVencto.Caption & "'"
            rpt.FormulaFields(17).Text = "'" & frmCnsParcela.lblDataVenctoCalc.Caption & "'"
            rpt.FormulaFields(18).Text = "'" & frmCnsParcela.lblValorLancado.Caption & "'"
            rpt.FormulaFields(19).Text = "'" & frmCnsParcela.lblValorAtualizado.Caption & "'"
            rpt.FormulaFields(20).Text = "'" & frmCnsParcela.lblDataPagto.Caption & "'"
            rpt.FormulaFields(21).Text = "'" & frmCnsParcela.lblDataReceita.Caption & "'"
            rpt.FormulaFields(22).Text = "'" & frmCnsParcela.lblBanco.Caption & "'"
            rpt.FormulaFields(23).Text = "'" & frmCnsParcela.lblAgencia.Caption & "'"
            rpt.FormulaFields(24).Text = "'" & frmCnsParcela.lblValorTaxa.Caption & "'"
            rpt.FormulaFields(25).Text = "'" & frmCnsParcela.txtValorPago.Text & "'"
            rpt.FormulaFields(26).Text = "'" & frmCnsParcela.txtValorDiferenca.Text & "'"
    Case "SENHAISS"
            frmReport.Caption = "Requerimento de senha para ISS Eletrônico"
            Liberado
           z = InputBox("Digite o Código da Empresa.", "Código da Empresa")
           If z = "" Then Exit Function
           If Val(z) = 0 Then
              MsgBox "Código Inválido.", vbExclamation, "Atenção"
              Exit Function
           Else
              z2 = InputBox("Digite o nome do assinante.", "Informe")
              z3 = InputBox("Digite o RG do assinante.", "Informe")
              z4 = InputBox("Digite o Telefone do assinante.", "Informe")
              
              If MsgBox("Empresa: " & z & vbCrLf & "Nome: " & z2 & vbCrLf & "RG: " & z3 & vbCrLf & "Telefone: " & z4, vbQuestion + vbYesNo, "Confirme os Dados") = vbYes Then
                Sql = "SELECT * FROM vwFULLEMPRESA WHERE CODIGOMOB=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                   If .RowCount = 0 Then
                       MsgBox "Empresa não Cadastrada.", vbExclamation, "Atenção"
                       Exit Function
                   Else
                       If Not IsNull(!dataencerramento) Or !dataencerramento <> CDate("01/01/1900") Then
                          MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
                          Exit Function
                       End If
                   End If
                End With
             Else
                Exit Function
             End If
           End If
           rpt.FormulaFields(1).Text = "'" & Format(RdoAux!codigomob, "000000") & "'"
           rpt.FormulaFields(2).Text = "'" & RdoAux!RazaoSocial & "'"
           rpt.FormulaFields(3).Text = "'" & RdoAux!Logradouro & ", " & Val(SubNull(RdoAux!Numero)) & " Bairro: " & SubNull(RdoAux!DescBairro) & " - " & SubNull(RdoAux!descCidade) & "/ " & SubNull(RdoAux!SiglaUF) & " CEP: " & RetornaCEP(RdoAux!CodLogradouro, RdoAux!Numero) & "'"
           rpt.FormulaFields(4).Text = "'" & RdoAux!ativextenso & "'"
           rpt.FormulaFields(5).Text = "'" & z2 & "'"
           rpt.FormulaFields(6).Text = "'" & z3 & "'"
           rpt.FormulaFields(7).Text = "'" & z4 & "'"
    Case "2VIA"
            frmReport.Caption = "Emissão de 2ª Via de Carnê"
            rpt.RecordSelectionFormula = "{CARNETMP.COMPUTER}='" & NomeDoUsuario & "'"
    Case "ARRECADACAOSN"
            frmReport.Caption = "Arrecadação do Simples Nacional"
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {ARRECADACAOSN.COMPUTER}='" & NomeDoUsuario & "'"
    Case "PROCESSOENVIADO"
            frmReport.Caption = "Processos enviados por Centro de Custo"
            rpt.FormulaFields(1).Text = "'" & frmProcessosEnviados.mskData.Text & " e " & frmProcessosEnviados.mskData2.Text & "'"
            rpt.RecordSelectionFormula = "{PROCESSOENVIO.COMPUTER}='" & NomeDeLogin & "'"
    Case "DOCUMENTOASSUNTO"
            frmReport.Caption = "Documentos por Assunto"
            rpt.RecordSelectionFormula = "{COMMAND.CODIGO}=" & frmAssunto.lstAssunto.ItemData(frmAssunto.lstAssunto.ListIndex)
    Case "DAM", "DAMHONORARIO", "DAMTMP"
            frmReport.Caption = "Impressão de DAM"
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {DAM.SID}=" & nNumDoc
            'rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {DAM.COMPUTER}='" & NomeDoUsuario & "' AND {DAM.USUARIO}='" & NomeDeLogin & "'"
            If UCase(sReport) = "DAM" Then
                For x = 0 To Forms.Count - 1
                    If Forms(x).Name = "frmDAM" Then
                        If Forms(x).Visible Then
                            bDam = True
                            If frmDAM.chkCobranca.value = vbChecked Then
                                rpt.FormulaFields(2).Text = "'S'"
                            Else
                                rpt.FormulaFields(2).Text = "'N'"
                            End If
                            rpt.FormulaFields(1).Text = "'" & frmDAM.mskVencimento.Text & "'"
                        Else
                            rpt.FormulaFields(1).Text = "'" & frmITBI.mskVencto.Text & "'"
                        End If
                       Exit For
                    Else
                        If Forms(x).Name = "frmITBI" Then
                            rpt.FormulaFields(1).Text = "'" & frmITBI.mskVencto.Text & "'"
                            Exit For
                        End If
                    End If
                Next
            ElseIf UCase(sReport) = "DAMHONORARIO" Then
                rpt.FormulaFields(1).Text = "'" & frmDAM.mskVencimento.Text & "'"
            End If
    Case "DIVIDATIVACANCELADO"
            rpt.RecordSelectionFormula = "{DIVIDATIVA.USUARIO}='" & NomeDeLogin & "'"
    Case "DEVEDORES", "DEVEDORES2", "DEVEDORES3"
            frmReport.Caption = "Relação de Devedores"
            rpt.RecordSelectionFormula = "{DAM.COMPUTER}='" & NomeDeLogin & "'"
    Case "LISTAPAGTOSN"
            frmReport.Caption = "LISTA DE PAGAMENTO SIMPLES NACIONAL"
            rpt.FormulaFields(1).Text = "'" & frmListaSN.cmbAno.Text & "'"
    Case "ISSMENSAL"
            frmReport.Caption = "ISS MENSAL"
            rpt.RecordSelectionFormula = "{ISSMENSAL.COMPUTER}='" & NomeDoUsuario & "'"
            If frmISSMensal.optTipo(0).value = True Then
                sTipo = "ESTIMADO"
            ElseIf frmISSMensal.optTipo(1).value = True Then
                sTipo = "VARIÁVEL"
            ElseIf frmISSMensal.optTipo(2).value = True Then
                sTipo = "FIXO"
            End If
            rpt.FormulaFields(2).Text = "'" & sTipo & "'"
    Case "ISSMENSALNAOPAGO"
            frmReport.Caption = "ISS MENSAL NÃO PAGO"
            rpt.RecordSelectionFormula = "{vwISSMENSALNAOPAGO.CODLANCAMENTO}=" & IIf(frmISSMensal.optTipo(0).value = True, 3, 5)
            rpt.FormulaFields(2).Text = "'" & IIf(frmISSMensal.optTipo(0).value, "ESTIMADO", "VARIÁVEL") & "'"
    Case "ISSMENSALFORA"
            frmReport.Caption = "ISS MENSAL"
            rpt.RecordSelectionFormula = "{ISSMENSAL.COMPUTER}='" & NomeDoUsuario & "'"
            rpt.FormulaFields(2).Text = "'" & IIf(frmISSMensal.optTipo(0).value, "ESTIMADO", "VARIÁVEL") & "'"
    Case "LISTARURAL"
            frmReport.Caption = "Cadastro das Propriedades Rurais"
            If NomeDeLogin = "FABIO" Or NomeDeLogin = "SCHWARTZ" Then
                z = InputBox("Digite o valor inicial!", "Valor por Hectare")
                z2 = InputBox("Digite o valor Final!", "Valor por Hectare")
                If Val(z2) > 0 Then
                   rpt.RecordSelectionFormula = "{cadastrorural.valor1}/{cadastrorural.hectare}>=" & Val(z) & " and {cadastrorural.valor1}/{cadastrorural.hectare}<=" & Val(z2)
                End If
            End If
    Case "LISTARURAL3"
            frmReport.Caption = "Simulação de Cálculo"
            z = InputBox("Digite o valor para o simulado", "Valor de Simulação")
            If Val(z) > 0 Then
               rpt.FormulaFields(2).Text = z
            End If
    Case "EXTRATO", "EXTRATOFULL", "EXTRATOFORUM", "EXTRATO3"
            rpt.RecordSelectionFormula = "{EXTRATOTMP.COMPUTER}='" & NomeDeLogin & "'"
            frmReport.Caption = "Extrato de Lançamento(s)"
            rpt.FormulaFields(6).Text = "'Gerado pelo Sistema Tributário Municipal (GTI) - Os débitos foram atualizados até: " & Format(dDataAtualiza, "dd/mm/yyyy") & "'"
    Case "AVERBACAO"
            rpt.RecordSelectionFormula = "{AVERBACAO.COMPUTER}='" & NomeDoUsuario & "'"
            frmReport.Caption = "Certidão de Averbação"
    Case "BAIXATMP2"
            rpt.RecordSelectionFormula = "{BAIXATMP.COMPUTADOR}='" & NomeDoUsuario & "'"
            frmReport.Caption = "Relatório de Baixas"
    Case "EMPRESAPORCNPJ"
           z = InputBox("Deseja imprimir apenas as empresas do Simples Nacional?", "Tipo de Relatório", "N")
           If UCase$(z) <> "N" And UCase$(z) <> "S" Then
                MsgBox "Digite S ou N"
                Exit Function
            End If
           If UCase$(z) = "S" Then
                rpt.RecordSelectionFormula = "{COMMAND.SN}=1"
                rpt.FormulaFields(1).Text = "'S'"
           Else
                rpt.FormulaFields(1).Text = "'C'"
           End If
    Case "PROCESSOASSUNTO"
           frmReport.Caption = "PROCESSOS POR ASSUNTO"
           z = InputBox("Digite o ano.", "ano do relatório", Year(Now))
           If Val(z) < 1900 Or Val(z) > 2030 Then
                MsgBox "Ano inválido"
                Exit Function
            End If
            rpt.RecordSelectionFormula = "{COMMAND.ANO}=" & Val(z)
    Case "EMPRESAATIVIDADE"
        z = InputBox("Data inicial de abertura", "Datas", "01/01/1970")
        If Not IsDate(z) Then
            MsgBox "Data inválida !!!", vbCritical, "Atenção"
            Exit Function
        End If
        z1 = InputBox("Data final de abertura", "Datas", Format(Now, "dd/mm/yyyy"))
        If Not IsDate(z1) Then
            MsgBox "Data inválida !!!", vbCritical, "Atenção"
            Exit Function
        End If
        rpt.FormulaFields(1).Text = "'DATA DE ABERTURA ENTRE:" & z & " E " & z1 & "'"
        rpt.RecordSelectionFormula = "{COMMAND.DATAABERTURA}>=#" & CDate(z) & "# AND {COMMAND.DATAABERTURA}<=#" & CDate(z1) & "#"
    Case "ALVARAPROVISORIO", "ALVARAPROVISORIOVICE"
            If (frmAlvara.cmbAss.ListIndex > 0) Then
                 Sql = "SELECT USUARIO FROM ASSINATURA WHERE NOME='" & frmAlvara.cmbAss.Text & "'"
                 Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                 With RdoAux
                      rpt.RecordSelectionFormula = "{ASSINATURA.USUARIO}='" & !USUARIO & "'"
                     .Close
                 End With
            Else
                 rpt.RecordSelectionFormula = "{ASSINATURA.USUARIO}='NOBODY'"
            End If

            rpt.FormulaFields(1).Text = "'" & frmAlvara.txtAlvara.Text & frmAlvara.lblAnoAlvara.Caption & "'"
            rpt.FormulaFields(2).Text = "'" & frmAlvara.lblNome.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & IIf(Left(frmAlvara.mskCNPJ.Text, 1) <> "_", frmAlvara.mskCNPJ.Text, frmAlvara.mskCPF.Text) & "'"
            rpt.FormulaFields(4).Text = "'" & frmAlvara.txtCodigo.Text & "'"
            rpt.FormulaFields(5).Text = "'" & frmAlvara.lblEndereco.Caption & "'"
            rpt.FormulaFields(6).Text = "'" & frmAlvara.lblNum.Caption & "'"
            rpt.FormulaFields(7).Text = "'" & frmAlvara.lblBairro.Caption & "'"
            rpt.FormulaFields(8).Text = "'" & frmAlvara.lblCep.Caption & "'"
            rpt.FormulaFields(9).Text = "'" & frmAlvara.lblAtividade.Caption & "'"
            rpt.FormulaFields(10).Text = "'" & frmAlvara.txtProcesso.Text & "'"
            sTexto1 = "'"
            For z = 0 To frmAlvara.lstDoc.ListCount - 1
                If frmAlvara.lstDoc.Selected(z) = True Then
                    sTexto1 = sTexto1 & "- " & frmAlvara.lstDoc.List(z) & "' + Chr(13) + '"
                End If
            Next
            sTexto1 = Left(sTexto1, Len(sTexto1) - 13)
            rpt.FormulaFields(11).Text = sTexto1
    
           Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & Val(frmAlvara.txtCodigo.Text)
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
                sHor = Mask(SubNull(!HORARIOEXT))
                If sHor = "" Then
                     ql = "SELECT DESCHORARIO FROM HORARIOFUNC WHERE CODHORARIO=" & !HORARIO
                     Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                     With RdoAux2
                          If .RowCount > 0 Then
                             sHor = !DESCHORARIO
                          End If
                         .Close
                     End With
                End If
                sTexto1 = sHor
                rpt.FormulaFields(12).Text = "'" & sTexto1 & "'"
                rpt.FormulaFields(13).Text = "'" & Mask(frmAlvara.txtObs.Text) & "'"
           End With
           

    Case "NOTIFICACAO"
           z = InputBox("Digite o numero do processo", "Informação requerida")
           z2 = InputBox("Digite o numero da notificação", "Informação requerida")
           MsgBox z
           MsgBox z2
Exit Function
    
    Case "ITBI"
    On Error Resume Next
            rpt.FormulaFields(1).Text = "'Conforme disposto no Artigo " & frmITBI.txtArtigo.Text & " da Lei Complementar nº 07/92 não " & _
            "incide cobrança de ITBI (imposto sobre transimissão de bens intervivos) sobre a transação acima referida. Nada mais.'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmITBI.lblProp.Caption) & "'"
            rpt.FormulaFields(3).Text = "'" & frmITBI.lblNumInsc.Caption & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmITBI.txtTipo.Text) & "'"
            rpt.FormulaFields(5).Text = "'" & frmITBI.txtValor.Text & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmITBI.txtDesc.Text) & "'"
            rpt.FormulaFields(7).Text = "'" & frmITBI.txtFunc.Text & "'"
            frmReport.Caption = "imposto sobre transimissão de bens intervivos"
    Case "COBRANCAAMIGAVEL", "COBRANCAAMIGAVELDA", "COBRANCAAMIGAVELSUJEITO", "COBRANCAAMIGAVELREFIS"
            rpt.FormulaFields(1).Text = "'" & Mask(frmCobrancaAmigavel.txtResp.Text) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmCobrancaAmigavel.txtObs1.Text) & "'"
            frmReport.Caption = "Carta de cobrança amigável"
    Case "COBRANCAAMIGAVELISS"
            rpt.RecordSelectionFormula = "{vwCOBRANCAAMIGAVELISS.CODREDUZIDO} >= " & Val(frmCobrancaAmigavel.txtCod1.Text) & " AND {vwCOBRANCAAMIGAVELISS.CODREDUZIDO} <= " & Val(frmCobrancaAmigavel.txtCod2.Text)
            rpt.FormulaFields(7).Text = "'" & Mask(frmCobrancaAmigavel.txtResp.Text) & "'"
            rpt.FormulaFields(8).Text = "'" & Mask(frmCobrancaAmigavel.txtObs1.Text) & "'"
            rpt.FormulaFields(9).Text = "'" & Mask(frmCobrancaAmigavel.txtObs2.Text) & "'"
            frmReport.Caption = "Carta de cobrança amigável"
    Case "CADASTRORURAL"
            rpt.RecordSelectionFormula = "{vwCADASTRORURAL.CODREDUZIDO} = " & Val(frmCadastroRural.lblCodReduzido.Caption)
            frmReport.Caption = "Cadastro Rural"
    Case "DIVIDATIVATOTAL"
            frmReport.Caption = "Relatório Sintético da Divida Ativa"
    Case "DIVIDATIVA", "DIVIDATIVAPARC"
            frmReport.Caption = "Relatório Analítico da Divida Ativa"
            rpt.RecordSelectionFormula = "{DIVIDATIVA.USUARIO}='" & NomeDeLogin & "'"
    Case "CONFDIVIDA", "CONFDIVIDATMP"
            frmReport.Caption = "Termo de Confissão de Divida Fiscal"
            nNumproc = Val(Left$(frmConfissaoDivida.txtNumProc.Text, Len(frmConfissaoDivida.txtNumProc.Text) - 5))
            sNumProc = CStr(nNumproc) & RetornaDVProcesso(nNumproc)
            nAno = Val(Right$(frmConfissaoDivida.txtNumProc.Text, 4))
            sNumProc = sNumProc & "/" & CStr(nAno)

            rpt.FormulaFields(1).Text = "'" & Replace$(frmConfissaoDivida.lblProp.Caption, "'", "") & "'"
            rpt.FormulaFields(2).Text = "'" & frmConfissaoDivida.lblCod.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & Replace$(frmConfissaoDivida.lblEnd.Caption, "'", "") & "'"
            rpt.FormulaFields(4).Text = "'" & Replace$(frmConfissaoDivida.lblRequerente.Caption, "'", "") & "'"
            rpt.FormulaFields(5).Text = "'" & Replace$(frmConfissaoDivida.lblEndCor.Caption, "'", "") & "'"
            
            rpt.FormulaFields(6).Text = "'" & frmConfissaoDivida.lblCPF.Caption & "'"
            rpt.FormulaFields(7).Text = "'" & sNumProc & "'"
            rpt.FormulaFields(8).Text = "'" & frmConfissaoDivida.lblAno.Caption & "'"
            rpt.FormulaFields(9).Text = "'" & frmConfissaoDivida.lblValor.Caption & "'"
            rpt.FormulaFields(10).Text = "'" & frmConfissaoDivida.lblQtdeParc.Caption & "'"
            rpt.FormulaFields(11).Text = "'" & frmConfissaoDivida.lblVenc.Caption & "'"
                If frmConfissaoDivida.lblDI.Caption = "S" Then
                    rpt.FormulaFields(13).Text = "'" & "30 (Trinta)" & "'"
                Else
                    If bRefisAtivo Then
                        rpt.FormulaFields(13).Text = "'" & "30 (Trinta)" & "'"
                    Else
                        'rpt.FormulaFields(13).Text = "'" & "90 (Noventa)" & "'"
                        rpt.FormulaFields(13).Text = "'" & "30 (Trinta)" & "'"
                    End If
                End If
'            rpt.FormulaFields(13).Text = "'Jaboticabal, " & Format(CDate(frmConfissaoDivida.lblDataProc.Caption), "dd") & " de " & Format(CDate(frmConfissaoDivida.lblDataProc.Caption), "mmmm") & " de " & Format(CDate(frmConfissaoDivida.lblDataProc.Caption), "yyyy") & "'"
            
            Sql = "SELECT DISTINCT CODREDUZIDO From ORIGEMREPARC WHERE NUMPROCESSO = '" & frmConfissaoDivida.txtNumProc.Text & "'"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                 If .RowCount > 1 Then
                    z = "Os imóveis que fazem parte deste reparcelamento são os seguintes: "
                    Do Until .EOF
                        z = z & !CODREDUZIDO & ", "
                       .MoveNext
                    Loop
                    z = Left$(z, Len(z) - 2)
                    rpt.FormulaFields(12).Text = "'" & z & "'"
                 End If
                .Close
            End With
    Case "SIMULADO", "SIMULADOTMP"
            frmReport.Caption = "Simulado de Parcelamento"
            rpt.FormulaFields(3).Text = "'" & frmParcelamento2.lblAno.Caption & "'"
            rpt.FormulaFields(2).Text = "'" & Replace$(frmParcelamento2.lblNome.Caption, "'", "") & "'"
            rpt.FormulaFields(1).Text = "'" & frmParcelamento2.txtCod.Text & "'"
            If frmParcelamento2.chkAnistia.value = vbChecked Then
'                   rpt.FormulaFields(4).Text = "'1'"
            Else
'                  rpt.FormulaFields(4).Text = "'0'"
            End If
            rpt.RecordSelectionFormula = "{SIMULADOREPARC.COMPUTER}='" & NomeDoUsuario & "'"
     Case "CERTIDAODEMOLICAO", "CERTIDAOENDERECO", "CERTIDAOISENCAO"
            rpt.RecordSelectionFormula = "{CERTIDAO.COMPUTER}='" & NomeDoUsuario & "'"
     Case "PROTOCOLOENTRADA"
            Dim sDoc As String
            rpt.RecordSelectionFormula = "{PROCESSOGTI.ANO}=" & Val(frmProcesso.lblAno.Caption) & " AND {PROCESSOGTI.NUMERO}=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
            If frmProcesso.optEnd(0).value = True Then
                Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            Else
                Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            End If
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            On Error Resume Next
            With RdoAux2
                If .RowCount > 0 Then
                     sNome = !nomecidadao
                     If Val(SubNull(!FCodLogradouro)) > 0 Then
                         Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                         Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                         Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                         Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                         Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                         With RdoS
                             If .RowCount > 0 Then
                                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                             Else
                                sEnd = ""
                             End If
                            .Close
                         End With
                     Else
                        sEnd = SubNull(!FNomeLogradouro)
                     End If
                     sEnd = sEnd & " " & SubNull(RdoAux2!fNUMIMOVEL)
                     sDoc = ""
                     If SubNull(!CPF) <> "" Then
                         sDoc = !CPF
                     Else
                         If SubNull(!Cnpj) <> "" Then
                             sDoc = !Cnpj
                         Else
                             sDoc = SubNull(!frg)
                         End If
                     End If
                      
                     Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                     If RdoS.RowCount > 0 Then
                         sCidade = RdoS!descCidade
                     Else
                          sCidade = ""
                     End If
                     If Not IsNull(!CodBairro) Then
                         Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                         Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                         If .RowCount > 0 Then
                             sBairro = RdoS!DescBairro
                         Else
                             sBairro = ""
                         End If
                     Else
                         sBairro = ""
                     End If
                     sUF = SubNull(!fsiglauf)
                     sFone = SubNull(!telefone)
                     sCep = SubNull(!FCEP)
                Else
                    sEnd = ""
                    sBairro = ""
                    sCidade = ""
                    sFone = ""
                    sUF = ""
                    sCep = ""
                End If
               .Close
            End With
            rpt.FormulaFields(1).Text = "'" & Mask(sNome) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(sEnd) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(sCidade) & "'"
            rpt.FormulaFields(4).Text = "'" & sDoc & "'"
            rpt.FormulaFields(5).Text = "'" & sUF & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(sBairro) & "'"
            
            rpt.FormulaFields(7).Text = "'" & RetornaDVProcesso(Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))) & "'"
     Case "CPPROTOCOLOENTRADA"
            Dim sDoc2 As String
            rpt.RecordSelectionFormula = "{CPPROCESSOGTI.ANO}=" & Val(frmCPProcesso.lblAno.Caption) & " AND {CPPROCESSOGTI.NUMERO}=" & Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))
            rpt.FormulaFields(1).Text = "'" & frmCPProcesso.cmbReq.Text & "'"
            Sql = "SELECT * FROM VWCIDADAO WHERE CODCIDADAO=" & Val(frmCPProcesso.lblCodCid.Caption)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                 If .RowCount > 0 Then
                     rpt.FormulaFields(1).Text = "'" & Mask(!nomecidadao) & "'"
                     If Val(SubNull(!CodLogradouro)) = 0 Then
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL)) & "'"
                     Else
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & "'"
                     End If
                     
                     If Not IsNull(!CodCidade) Then
                        Sql = "SELECT DESCCIDADE FROM CIDADE WHERE CODCIDADE=" & IIf(Val(!CodCidade) = 0, 413, !CodCidade) & " AND SIGLAUF='" & IIf(Val(!SiglaUF) > 0, "SP", !SiglaUF) & "'"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If Not IsNull(RdoAux2!descCidade) Then
                            rpt.FormulaFields(3).Text = "'" & RdoAux2!descCidade & "'"
                        End If
                        RdoAux2.Close
                     Else
                        rpt.FormulaFields(3).Text = "''"
                     End If
                     sDoc2 = ""
                     If Not IsNull(!CPF) Then
                         sDoc2 = !CPF
                     Else
                         If Not IsNull(!Cnpj) Then
                             sDoc2 = !Cnpj
                         Else
                             If Not IsNull(!rg) Then
                                 sDoc2 = !rg
                             End If
                         End If
                     End If
                     rpt.FormulaFields(4).Text = "'" & sDoc & "'"
                     If Not IsNull(!SiglaUF) Then
                        rpt.FormulaFields(5).Text = "'" & IIf(Val(!SiglaUF) > 0, "SP", !SiglaUF) & "'"
                     Else
                        rpt.FormulaFields(5).Text = "''"
                     End If
                                             
                     If Not IsNull(!CodBairro) Then
                         Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE CODCIDADE=" & IIf(Val(!CodCidade) = 0, 413, !CodCidade) & " AND CODBAIRRO=" & !CodBairro
                         Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                         If RdoAux2.RowCount > 0 Then
                            If Not IsNull(RdoAux2!DescBairro) And RdoAux2!DescBairro <> "" Then
                                rpt.FormulaFields(6).Text = "'" & RdoAux2!DescBairro & "'"
                            End If
                         End If
                         RdoAux2.Close
                     End If
                 End If
                .Close
            End With
            rpt.FormulaFields(7).Text = "'" & RetornaDVProcesso(Val(Left$(frmCPProcesso.lblNumProc.Caption, Len(frmCPProcesso.lblNumProc.Caption) - 2))) & "'"
     Case "REQUERIMENTO", "REQUERIMENTOCANCEL"
            rpt.RecordSelectionFormula = "{PROCESSOGTI.ANO}=" & Val(frmProcesso.lblAno.Caption) & " AND {PROCESSOGTI.NUMERO}=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
            
            If frmProcesso.optEnd(0).value = True Then
                Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO AS fCODLOGRADOURO,NUMIMOVEL AS fNUMIMOVEL,"
                Sql = Sql & "COMPLEMENTO AS fCOMPLEMENTO,CODBAIRRO AS fCODBAIRRO,CODCIDADE AS fCODCIDADE,SIGLAUF AS fSIGLAUF,"
                Sql = Sql & "CEP AS fCEP,TELEFONE AS fTELEFONE,EMAIL AS fEMAIL,RG AS fRG,NOMELOGRADOURO AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            Else
                Sql = "SELECT CODCIDADAO,NOMECIDADAO,CPF,CNPJ,CODLOGRADOURO2 AS fCODLOGRADOURO,NUMIMOVEL2 AS fNUMIMOVEL,"
                Sql = Sql & "COMPLEMENTO2 AS fCOMPLEMENTO,CODBAIRRO2 AS fCODBAIRRO,CODCIDADE2 AS fCODCIDADE,SIGLAUF2 AS fSIGLAUF,"
                Sql = Sql & "CEP2 AS fCEP,TELEFONE2 AS fTELEFONE,EMAIL2 AS fEMAIL,RG AS fRG,NOMELOGRADOURO2 AS fNOMELOGRADOURO,ORGAO AS fORGAO"
                Sql = Sql & " FROM CIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            End If
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            On Error Resume Next
            With RdoAux2
                If .RowCount > 0 Then
                     sNome = !nomecidadao
                     If Val(SubNull(!FCodLogradouro)) > 0 Then
                         Sql = "SELECT CODLOGRADOURO,CODTIPOLOG,NOMETIPOLOG,"
                         Sql = Sql & "ABREVTIPOLOG,CODTITLOG,NOMETITLOG,"
                         Sql = Sql & "ABREVTITLOG,NOMELOGRADOURO "
                         Sql = Sql & "FROM vwLOGRADOURO WHERE CODLOGRADOURO=" & !FCodLogradouro
                         Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                         With RdoS
                             If .RowCount > 0 Then
                                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NomeLogradouro
                             Else
                                sEnd = ""
                             End If
                            .Close
                         End With
                     Else
                        sEnd = SubNull(!FNomeLogradouro)
                     End If
                     sEnd = sEnd & " " & SubNull(RdoAux2!fNUMIMOVEL)
                     sDoc = ""
                     If SubNull(!CPF) <> "" Then
                         rpt.FormulaFields(7).Text = "'" & "Pessoa Física" & "'"
                         sDoc = !CPF
                     Else
                         If SubNull(!Cnpj) <> "" Then
                             rpt.FormulaFields(7).Text = "'" & "Pessoa Jurídica" & "'"
                             sDoc = Format(Trim(!Cnpj), "00\.000\.000/0000-00")
'                         Else
'                             rpt.FormulaFields(7).Text = "'" & "Pessoa Física" & "'"
'                             sDoc = SubNull(!frg)
                         End If
                     End If
                     rpt.FormulaFields(8).Text = "'" & SubNull(!frg) & "'"
                     rpt.FormulaFields(10).Text = "'" & Mask(SubNull(!fORGAO)) & "'"
                     Sql = "SELECT DESCCIDADE FROM CIDADE WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade
                     Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                     If RdoS.RowCount > 0 Then
                         sCidade = RdoS!descCidade
                     Else
                          sCidade = ""
                     End If
                     If Not IsNull(!CodBairro) Then
                         Sql = "SELECT DESCBAIRRO FROM BAIRRO WHERE SIGLAUF='" & !fsiglauf & "' AND CODCIDADE=" & !fCodCidade & " AND CODBAIRRO=" & !fCodBairro
                         Set RdoS = cn.OpenResultset(Sql, rdOpenKeyset)
                         If .RowCount > 0 Then
                             sBairro = RdoS!DescBairro
                         Else
                             sBairro = ""
                         End If
                     Else
                         sBairro = ""
                     End If
                     sUF = SubNull(!fsiglauf)
                     sFone = SubNull(!telefone)
                     sCep = SubNull(!FCEP)
                Else
                    sEnd = ""
                    sBairro = ""
                    sCidade = ""
                    sFone = ""
                    sUF = ""
                    sCep = ""
                End If
               .Close
            End With
            rpt.FormulaFields(1).Text = "'" & Mask(sNome) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(sEnd) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(sCidade) & "'"
            rpt.FormulaFields(4).Text = "'" & sDoc & "'"
            rpt.FormulaFields(5).Text = "'" & sUF & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(sBairro) & "'"
            rpt.FormulaFields(11).Text = "'......." & frmProcesso.lblNumProc.Caption & "'"
            
                Sql = "SELECT PROCESSOEND.CODLOGR,PROCESSOEND.NUMERO, vwLOGRADOURO.ABREVTIPOLOG, vwLOGRADOURO.ABREVTITLOG,"
                Sql = Sql & "vwLOGRADOURO.NomeLogradouro FROM PROCESSOEND INNER JOIN "
                Sql = Sql & "vwLOGRADOURO ON PROCESSOEND.CODLOGR = vwLOGRADOURO.CODLOGRADOURO "
                Sql = Sql & "Where PROCESSOEND.ANO = " & Val(frmProcesso.lblAno.Caption) & " And "
                Sql = Sql & "PROCESSOEND.NUMPROCESSO = " & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
                Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux2.RowCount > 0 Then
                    sTexto1 = "'"
                    Do Until RdoAux2.EOF
                        sTexto1 = sTexto1 & CStr(RdoAux2.AbsolutePosition) & ") " & (Trim$(RdoAux2!AbrevTipoLog) & " " & Trim$(SubNull(RdoAux2!AbrevTitLog)) & " " & RdoAux2!NomeLogradouro & ", " & RdoAux2!Numero) & "  "
                       RdoAux2.MoveNext
                    Loop

                    rpt.FormulaFields(9).Text = sTexto1 & "'"  'endereco
                End If
                RdoAux2.Close
            
     Case "COMUNICADOJUDICIAL"
            frmReport.Caption = "Comunicado judicial"
            z = InputBox("Digite o código/inscrição", "Informação requerida")
            sTexto1 = ""
            Liberado
            If Val(z) > 0 And Val(z) < 100000 Then 'IMOVEL
                Sql = "SELECT NOMECIDADAO FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    sTexto1 = SubNull(RdoAux!nomecidadao)
                Else
                    MsgBox "Cadastro não existe", vbCritical, "Erro"
                    Exit Function
                End If
                RdoAux.Close
            ElseIf Val(z) >= 100000 And Val(z) < 300000 Then 'EMPRESA
                Sql = "SELECT RAZAOSOCIAL FROM vwFULLEMPRESA WHERE CODIGOMOB=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    sTexto1 = RdoAux!RazaoSocial
                Else
                    MsgBox "Cadastro não existe", vbCritical, "Erro"
                    Exit Function
                End If
                RdoAux.Close
            ElseIf Val(z) >= 500000 And Val(z) < 700000 Then 'CIDADAO
                Sql = "SELECT NOMECIDADAO FROM CIDADAO WHERE CODCIDADAO=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount > 0 Then
                    sTexto1 = RdoAux!nomecidadao
                Else
                    MsgBox "Cadastro não existe", vbCritical, "Erro"
                    Exit Function
                End If
                RdoAux.Close
            End If
            z1 = InputBox("Digite o número da execução fiscal", "Informação requerida")
            
            rpt.FormulaFields(1).Text = "'" & Mask(CStr(z)) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(CStr(z1)) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(CStr(sTexto1)) & "'"
     Case "COMUNICADODOC"
            rpt.RecordSelectionFormula = "{PROCESSOGTI.ANO}=" & Val(frmProcesso.lblAno.Caption) & " And {PROCESSOGTI.NUMERO}=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))
            Sql = "SELECT * FROM VWCIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                 If .RowCount > 0 Then
                     rpt.FormulaFields(1).Text = "'" & Mask(!nomecidadao) & "'"
                     If Val(SubNull(!CodLogradouro)) = 0 Then
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL)) & "'"
                     Else
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & "'"
                     End If
                     
                     If Not IsNull(!CodCidade) And Not IsNull(!SiglaUF) Then
                        Sql = "SELECT DESCCIDADE FROM CIDADE WHERE CODCIDADE=" & !CodCidade & " AND SIGLAUF='" & !SiglaUF & "'"
                        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                        If Not IsNull(RdoAux2!descCidade) Then
                            rpt.FormulaFields(3).Text = "'" & RdoAux2!descCidade & "'"
                        End If
                        RdoAux2.Close
                     Else
                        rpt.FormulaFields(3).Text = "''"
                     End If
                     sDoc = ""
                     If Not IsNull(!CPF) Then
                         sDoc = !CPF
                     Else
                         If Not IsNull(!Cnpj) Then
                             sDoc = !Cnpj
                         Else
                             If Not IsNull(!rg) Then
                                 sDoc = !rg
                             End If
                         End If
                     End If
                     rpt.FormulaFields(4).Text = "'" & sDoc & "'"
                     rpt.FormulaFields(5).Text = "'" & !SiglaUF & "'"
                     rpt.FormulaFields(7).Text = "'" & RetornaDVProcesso(Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))) & "'"
                 End If
                .Close
            End With
     Case "COMPROVANTEDOC"
            rpt.RecordSelectionFormula = "{PROCESSOGTI.ANO}=" & Val(frmProcesso.lblAno.Caption) & " And {PROCESSOGTI.NUMERO}=" & Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2)) & " And  NOT isNull({PROCESSODOC.DATA})"
            Sql = "SELECT * FROM VWCIDADAO WHERE CODCIDADAO=" & Val(frmProcesso.lblCodCid.Caption)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                 If .RowCount > 0 Then
                     rpt.FormulaFields(1).Text = "'" & Mask(!nomecidadao) & "'"
                     If Val(SubNull(!CodLogradouro)) = 0 Then
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!NomeLogradouro)) & ", " & Val(SubNull(!NUMIMOVEL)) & "'"
                     Else
                        rpt.FormulaFields(2).Text = "'" & Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & !NOMELOGRADOURO2 & ", " & !NUMIMOVEL & "'"
                     End If
                     
                     Sql = "SELECT DESCCIDADE FROM CIDADE WHERE CODCIDADE=" & !CodCidade & " AND SIGLAUF='" & !SiglaUF & "'"
                     Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                     If Not IsNull(RdoAux2!descCidade) Then
                         rpt.FormulaFields(3).Text = "'" & RdoAux2!descCidade & "'"
                     End If
                     RdoAux2.Close
                     
                     sDoc = ""
                     If Not IsNull(!CPF) Then
                         sDoc = !CPF
                     Else
                         If Not IsNull(!Cnpj) Then
                             sDoc = !Cnpj
                         Else
                             If Not IsNull(!rg) Then
                                 sDoc = !rg
                             End If
                         End If
                     End If
                     rpt.FormulaFields(4).Text = "'" & sDoc & "'"
                     rpt.FormulaFields(5).Text = "'" & !SiglaUF & "'"
                                             
                 End If
                .Close
            End With
            rpt.FormulaFields(7).Text = "'" & RetornaDVProcesso(Val(Left$(frmProcesso.lblNumProc.Caption, Len(frmProcesso.lblNumProc.Caption) - 2))) & "'"
      Case "RESUMOPROTOCOLO", "RESUMOPROTOCOLOREQ"
            rpt.RecordSelectionFormula = "{RESUMODIARIO.USUARIO}='" & NomeDoUsuario & "'"
      Case "TRAMITEABERTOLOCAL"
           If IsDate(frmResumoProtocolo.mskDataDe.Text) Then
                rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & "{COMMAND.DATAENTRADA}>=#" & Format(frmResumoProtocolo.mskDataDe.Text, "mm/dd/yyyy") & " # AND {COMMAND.DATAENTRADA}<=#" & Format(frmResumoProtocolo.mskDataAte.Text, "mm/dd/yyyy") & "# "
           Else
                rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & "{COMMAND.DATAENTRADA}>=#01/01/1970# AND {COMMAND.DATAENTRADA}<=#" & Format(Now, "mm/dd/yyyy") & "# "
           End If
           rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {COMMAND.CODASSUNTO} IN [" & frmResumoProtocolo.txtAssunto.Text & "]"
           If frmResumoProtocolo.cmbSetor.ListIndex > -1 Then
                rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {COMMAND.DESCRICAO2}='" & frmResumoProtocolo.cmbSetor.Text & "' AND ISNULL({COMMAND.CODCIDADAO})"
           End If
           If frmResumoProtocolo.lblReq.Caption <> "0" Then
                rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {COMMAND.CODCIDADAO}=" & Val(Left(frmResumoProtocolo.lblReq.Caption, 6))
           End If
           Select Case frmResumoProtocolo.cmbOrder.ListIndex
                Case 0 'datahora
                    rpt.FormulaFields(3).Text = "{COMMAND.DATAENTRADA}"
                Case 1 'requerente
                    rpt.FormulaFields(3).Text = "{COMMAND.NOMECIDADAO}"
                Case 2 'assunto
                    rpt.FormulaFields(3).Text = "{COMMAND.COMPLEMENTO}"
               Case 3 'centro custo
                    rpt.FormulaFields(3).Text = "{COMMAND.DESCRICAO2}"
           End Select
           If frmResumoProtocolo.chkExterno.value = 1 Then
                rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {COMMAND.INTERNO}=FALSE"
           End If
End Select
    
If sReport <> "DAM" And sReport <> "DAMHONORARIO" And sReport <> "DAMTMP" Then
    rpt.PaperSize = crPaperA4
    
End If

If sReport = "DAM" And bDam Then
    If frmDAM.chkCobranca.value = vbChecked Then
        rpt.LeftMargin = 2500
    End If
End If

If frmMdi.m_cMenuPrincipal.Checked(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True Then
    rpt.PaperSource = crPRBinLower
End If

rpt.DisplayProgressDialog = True
If UCase$(sReport) = "CARNE" Then
    rpt.Sections(1).Suppress = bHeader
End If

Select Case UCase$(sReport)
    Case "CARNETMP", "CONFDIVIDATMP", "SIMULADOTMP", "BOLETOGUIATMP", "CALCULOPARCELAMENTOTMP", "DAMTMP", "ANALISE2_TMP"
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "TributacaoTeste", UL, UP
    Case Else
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "Tributacao", UL, UP
End Select


'If UCase$(sReport) = "CARNETMP" Or UCase$(sReport) = "CONFDIVIDATMP" Or UCase$(sReport) = "SIMULADOTMP" Or UCase$(sReport) = "COBRANCAAMIGAVELTMP" Or UCase$(sReport) = "CALCULOPARCELAMENTOTMP" Or UCase$(sReport) = "DAMTMP" Or UCase$(sReport) = "BOLETOGUIATMP" Then
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTribTeste", "TributacaoTeste", UL, UP
'ElseIf UCase$(sReport) = "CARNELOCAL" Or UCase$(sReport) = "EXTRATOFULL" Then
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTribLocal", "Tributacao_Full", UL, UP
'ElseIf UCase$(sReport) = "EXTRATO" Then
'    rpt.Database.Tables(1).SetLogOnInfo "192.168.15.160", "Tributacao", UL, UP
'Else
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTributacao", "Tributacao", UL, UP
'End If
rpt.DiscardSavedData

CRViewer1.ReportSource = rpt

show:
CRViewer1.ViewReport
Liberado

If nNumDoc > 0 And sReport <> "DAMHONORARIO" And NomeDoComputador <> "GENESIS" Then
'If nNumDoc > 0 And NomeDoComputador <> "ADM-PC" Then
    rpt.ExportOptions.DestinationType = crEDTDiskFile
    If bLocal Then
        rpt.ExportOptions.DiskFileName = "C:\TMP\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
    Else
        rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\SEGUNDAVIA\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
    End If
    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rpt.ExportOptions.PDFExportAllPages = True
    rpt.Export (False)
End If

If UCase(sReport) = "ALVARAPROVISORIO" Or UCase(sReport) = "ALVARAPROVISORIOVICE" Then
    Sql = "select count(seq) as maximo from documentopic where codigo=" & Val(frmAlvara.txtCodigo.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 1
    Else
        nSeq = RdoAux!maximo + 1
    End If
    RdoAux.Close
    
    Sql = "select max(seq) as maximo from documentopic"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq2 = 1
    Else
        nSeq2 = RdoAux!maximo + 1
    End If
    RdoAux.Close
    sTexto1 = "08" & Year(Now) & Format(nSeq, "00") & Format(frmAlvara.txtCodigo.Text, "000000") & ".pdf"
    
    Sql = "insert documentopic(seq,codigo,documento) values(" & nSeq2 & "," & Val(frmAlvara.txtCodigo.Text) & ",'" & sTexto1 & "')"
    cn.Execute Sql, rdExecDirect
    
    rpt.ExportOptions.DestinationType = crEDTDiskFile
    rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\Documentos\" & Year(Now) & "\" & sTexto1
    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rpt.ExportOptions.PDFExportAllPages = True
    rpt.Export (False)
End If


frmReport.show 1

Exit Function
Erro:

Liberado
MsgBox Err.Description
Resume Next
End Function

Public Function ShowReport2(sReport As String, hMDI As Long, hFormCalling As Long, Optional nNumDoc As Long, Optional nNumGuia As Long)
Dim RdoAux As rdoResultset, Sql As String, sTipo As String, dData As Date, nAno As Integer, sDoc As String, nSeq2 As Integer
Dim sTexto1 As String, sTexto2 As String, sTexto3 As String, sHor As String, sSenha As String, nSeq As Integer, nCodReduz As Long
Dim z As Variant, RdoAux2 As rdoResultset, z2 As Variant, z3 As Variant, z4 As Variant, z5 As Variant
Dim sNumProc As String, nNumproc As Long, bAchou As Boolean, aTributo() As String, x As Integer, Y As Integer
Dim qd As New rdoQuery, bHeader As Boolean

If bLocal Then
    Exit Function
End If

bHeader = False
On Error GoTo Erro
hFrmMdi = hMDI
hFrmCall = hFormCalling
Ocupado
DoEvents
If IsNull(nNumDoc) Then nNumDoc = 0

If UCase(sReport) = "BOLETOGUIA2" Then
    bHeader = True
    sReport = "BOLETOGUIA"
End If
If UCase(sReport) = "BOLETOGUIA2TMP" Then
    bHeader = True
    sReport = "BOLETOGUIATMP"
End If

Set rpt = crApp.OpenReport(sPathReport & "\" & sReport & ".Rpt", 1)

Select Case UCase(sReport)
    Case "CALCULOIPTU"
        frmReport.Caption = "Amostra de Cálculo de IPTU"
        z = InputBox("Digite o código do imóvel.", "Entre com a informação")
        z2 = InputBox("Digite o ano de cálculo.", "Entre com a informação")
        If z = "" Or z2 = "" Then Exit Function
        If Val(z) <= 0 Or Val(z) > 50000 Then
            MsgBox "Código de imóvel inválido", vbCritical, "Erro"
            Exit Function
        End If
        If Val(z2) < 2006 Or Val(z2) > Year(Now) Then
            MsgBox "Ano de cálculo inválido!" & vbCrLf & "(Somente a partir de 2006)", vbCritical, "Erro"
            Exit Function
        End If
        
        Sql = "select * from vwfullimovel2 where codreduzido=" & Val(z)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount = 0 Then
            MsgBox "Código de imóvel não cadastrado.", vbCritical, "Erro"
            Exit Function
        End If
        
        rpt.FormulaFields(1).Text = "'" & z2 & "'"
        rpt.FormulaFields(2).Text = "'" & Format(z, "000000") & " - " & Mask(RdoAux!nomecidadao) & "'"
        If RdoAux!CodCondominio <> 999 Then
            rpt.FormulaFields(3).Text = "'" & Mask(RdoAux!Logradouro) & ", " & RdoAux!Li_Num & " " & RdoAux!cd_nomecond & " " & SubNull(RdoAux!Li_Compl) & " - " & RdoAux!DescBairro & "'"
        Else
            rpt.FormulaFields(3).Text = "'" & Mask(RdoAux!Logradouro) & ", " & RdoAux!Li_Num & " " & SubNull(RdoAux!Li_Compl) & " - " & RdoAux!DescBairro & "'"
        End If
        rpt.FormulaFields(4).Text = "'Distrito: " & RdoAux!Distrito & " Setor: " & RdoAux!Setor & " Quadra: " & Format(RdoAux!Quadra, "0000") & " Lote: " & Format(RdoAux!Lote, "00000") & " Face: " & RdoAux!Seq & " '"
        
        Set qd.ActiveConnection = cn
        qd.QueryTimeout = 0
        RdoAux.Close
        qd.Sql = "{ Call spCALCULO(?,?) }"
        qd(0) = Val(z)
        qd(1) = Val(z2)
        Set RdoAux = qd.OpenResultset(rdOpenKeyset)
        rpt.FormulaFields(5).Text = "'" & FormatNumber(RdoAux!AreaTerreno, 2) & " m²'"
        rpt.FormulaFields(6).Text = "'" & FormatNumber(RdoAux!TESTADAPRINC, 2) & " m'"
        
        RdoAux.Close
    
    Case "AVISODEBITO"
            frmReport.Caption = "Aviso de débito"
            With frmAvisoDebito
                ReDim aTributo(0)
                For x = 0 To 6
                    If .chkT(x).value = 1 Then
                        ReDim Preserve aTributo(UBound(aTributo) + 1)
                        aTributo(UBound(aTributo)) = .chkT(x).Caption & " R$ " & .lblT(x).Caption
                    End If
                Next
            End With
            For x = 1 To UBound(aTributo)
                rpt.FormulaFields(x).Text = "'" & aTributo(x) & "'"
            Next
            rpt.FormulaFields(15).Text = "'" & FormatNumber(frmAvisoDebito.txtTaxa.Text, 2) & frmAvisoDebito.txtExtenso.Text & "'"
            rpt.FormulaFields(16).Text = "'" & frmAvisoDebito.txtNumProc.Text & "'"
            rpt.FormulaFields(9).Text = "'" & Format(frmAvisoDebito.txtCod.Text, "000000") & "'"
            rpt.FormulaFields(10).Text = "'" & Mask(frmAvisoDebito.txtNome.Text) & "'"
            rpt.FormulaFields(11).Text = "'" & Mask(frmAvisoDebito.txtEndereco.Text) & "'"
            rpt.FormulaFields(12).Text = "'" & Mask(frmAvisoDebito.txtRequerente.Text) & "'"
            rpt.FormulaFields(13).Text = "'" & Mask(frmAvisoDebito.txtEnd2.Text) & " - " & frmAvisoDebito.txtBairroCidade.Text & " - CEP: " & frmAvisoDebito.txtCEP.Text & "'"
            rpt.FormulaFields(14).Text = "'" & frmAvisoDebito.txtNumProcessoE.Text & "'"
    Case "EmpresaCnae"
            frmReport.Caption = "Empresas por Cnae"
            CRViewer1.DisplayGroupTree = True
    Case "ASSUNTO_DOC"
            frmReport.Caption = "Assuntos por documento"
            CRViewer1.DisplayGroupTree = True
    Case "MAIORPAGADOR"
            frmReport.Caption = "Maiores Pagadores"
            rpt.RecordSelectionFormula = "{MAIORPAGADOR.USUARIO}='" & NomeDeLogin & "'"
            rpt.FormulaFields(1).Text = "'" & frmDevedor.txtAnoDe.Text & "'"
            rpt.FormulaFields(2).Text = "'" & frmDevedor.txtAnoAte.Text & "'"
    Case "CONFDIVIDADAM"
            frmReport.Caption = "Termo de Confissão de Divida Fiscal (DAM)"

            rpt.FormulaFields(1).Text = "'" & Replace$(frmConfissaoDivida.lblProp.Caption, "'", "") & "'"
            rpt.FormulaFields(2).Text = "'" & frmConfissaoDivida.lblCod.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & Replace$(frmConfissaoDivida.lblEnd.Caption, "'", "") & "'"
            rpt.FormulaFields(4).Text = "'" & Replace$(frmConfissaoDivida.txtRequerente.Text, "'", "") & "'"
            If Len(frmConfissaoDivida.txtCPF.Text) = 11 Then
                rpt.FormulaFields(5).Text = "'" & Format(frmConfissaoDivida.txtCPF.Text, "000\.000\.000-00") & "'"
            Else
                rpt.FormulaFields(5).Text = "'" & Format(frmConfissaoDivida.txtCPF.Text, "00\.000\.000/0000-00") & "'"
            End If
            rpt.FormulaFields(6).Text = "'" & frmConfissaoDivida.lblAno.Caption & "'"
            rpt.FormulaFields(7).Text = "'" & frmConfissaoDivida.lblValor.Caption & "'"
            rpt.FormulaFields(8).Text = "'" & frmConfissaoDivida.mskVenc.Text & "'"
            rpt.FormulaFields(9).Text = "'" & frmConfissaoDivida.txtNumDoc.Text & "-" & RetornaDVNumDoc(CLng(frmConfissaoDivida.txtNumDoc.Text)) & "'"
            
            Sql = "SELECT DISTINCT CODREDUZIDO From ORIGEMREPARC WHERE NUMPROCESSO = '" & frmConfissaoDivida.txtNumProc.Text & "'"
            Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux2
                 If .RowCount > 1 Then
                    z = "Os imóveis que fazem parte deste reparcelamento são os seguintes: "
                    Do Until .EOF
                        z = z & !CODREDUZIDO & ", "
                       .MoveNext
                    Loop
                    z = Left$(z, Len(z) - 2)
                    rpt.FormulaFields(12).Text = "'" & z & "'"
                 End If
                .Close
            End With
    Case "DECA"
            rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
            frmReport.Caption = "Impressão de DECA frente"
            If Val(Right(frmDeca.Caption, 6)) > 500000 Then
                rpt.FormulaFields(1).Text = "''"
            Else
                rpt.FormulaFields(1).Text = IIf(Val(Right(frmDeca.Caption, 6)) = 0, "", Val(Right(frmDeca.Caption, 6)))
            End If
            rpt.FormulaFields(2).Text = "'" & Mask(frmDeca.txtNome.Text) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(frmDeca.txtRamo1.Text) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmDeca.txtRamo2.Text) & "'"
            rpt.FormulaFields(5).Text = "'" & Mask(frmDeca.txtCodAtiv.Text) & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmDeca.txtEnd.Text) & "'"
            rpt.FormulaFields(7).Text = "'" & Mask(frmDeca.txtAndar.Text) & "'"
            rpt.FormulaFields(8).Text = "'" & Mask(frmDeca.txtSala.Text) & "'"
            rpt.FormulaFields(9).Text = "'" & Mask(frmDeca.txtBairro.Text) & "'"
            rpt.FormulaFields(10).Text = "'" & Mask(frmDeca.txtCEP.Text) & "'"
            rpt.FormulaFields(11).Text = "'" & Mask(frmDeca.txtCidade.Text) & "'"
            rpt.FormulaFields(12).Text = "'" & Mask(frmDeca.txtZona.Text) & "'"
            rpt.FormulaFields(13).Text = "'" & Mask(frmDeca.txtFone.Text) & "'"
            rpt.FormulaFields(14).Text = "'" & Mask(frmDeca.txtDataAbe.Text) & "'"
            rpt.FormulaFields(15).Text = "'" & Mask(frmDeca.txtArea.Text) & "'"
            rpt.FormulaFields(16).Text = "'" & Mask(frmDeca.txtNumemp.Text) & "'"
            rpt.FormulaFields(17).Text = "'" & Mask(frmDeca.txtMunicipio.Text) & "'"
            rpt.FormulaFields(18).Text = "'" & Mask(frmDeca.txtOrgao.Text) & "'"
            rpt.FormulaFields(19).Text = "'" & Mask(frmDeca.txtNumReg.Text) & "'"
            rpt.FormulaFields(20).Text = "'" & Mask(frmDeca.txtCapital.Text) & "'"
            rpt.FormulaFields(21).Text = "'" & Mask(frmDeca.txtRG.Text) & "'"
            rpt.FormulaFields(22).Text = "'" & Mask(frmDeca.txtCPF.Text) & "'"
            rpt.FormulaFields(23).Text = "'" & IIf(frmDeca.chkO(0).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(25).Text = "'" & IIf(frmDeca.chkO(2).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(26).Text = "'" & IIf(frmDeca.chkO(3).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(27).Text = "'" & IIf(frmDeca.chkO(4).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(28).Text = "'" & IIf(frmDeca.chkO(5).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(29).Text = "'" & IIf(frmDeca.chkO(6).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(30).Text = "'" & IIf(frmDeca.chkO(7).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(31).Text = "'" & IIf(frmDeca.chkO(8).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(32).Text = "'" & IIf(frmDeca.chkO(9).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(33).Text = "'" & frmDeca.mskO(0).Text & "'"
            rpt.FormulaFields(35).Text = "'" & frmDeca.mskO(2).Text & "'"
            rpt.FormulaFields(36).Text = "'" & frmDeca.mskO(3).Text & "'"
            rpt.FormulaFields(37).Text = "'" & frmDeca.mskO(4).Text & "'"
            rpt.FormulaFields(38).Text = "'" & frmDeca.mskO(5).Text & "'"
            rpt.FormulaFields(39).Text = "'" & frmDeca.mskO(6).Text & "'"
            rpt.FormulaFields(40).Text = "'" & frmDeca.mskO(7).Text & "'"
            rpt.FormulaFields(41).Text = "'" & frmDeca.mskO(8).Text & "'"
            rpt.FormulaFields(42).Text = "'" & frmDeca.mskO(9).Text & "'"
            rpt.FormulaFields(43).Text = "'" & IIf(frmDeca.chkT(0).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(44).Text = "'" & IIf(frmDeca.chkT(1).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(45).Text = "'" & IIf(frmDeca.chkE(0).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(46).Text = "'" & IIf(frmDeca.chkE(1).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(47).Text = "'" & IIf(frmDeca.chkE(2).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(48).Text = "'" & IIf(frmDeca.chkE(3).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(49).Text = "'" & IIf(frmDeca.chkE(4).value = vbUnchecked, " ", "X") & "'"
            rpt.FormulaFields(50).Text = "'" & Mask(frmDeca.txtHist.Text) & "'"
            rpt.FormulaFields(51).Text = "'" & Mask(frmDeca.txtAssinatura.Text) & "'"
            rpt.FormulaFields(52).Text = "'" & Mask(frmDeca.txtEndEntrega.Text) & "'"
            If frmDeca.chkAmbulante.value = vbChecked Then
                rpt.FormulaFields(53).Text = "'X'"
                rpt.FormulaFields(54).Text = "'" & frmDeca.cmbAmbulante.Text & "'"
                rpt.FormulaFields(55).Text = "'Trabalho como comércio ambulante de: " & Mask(frmDeca.txtDescAmbulante.Text) & "'"
            Else
                rpt.FormulaFields(53).Text = "''"
                rpt.FormulaFields(54).Text = "''"
                rpt.FormulaFields(55).Text = "''"
            End If
            rpt.FormulaFields(56).Text = "'" & Mask(frmDeca.txtDescAmb.Text) & "'"
            rpt.FormulaFields(57).Text = "'" & Mask(frmDeca.txtEmailEmpresa.Text) & "'"
    Case "DOCUMENTOSEMITIDOS"
            frmReport.Caption = "Documentos emitidos"
Data1:
            z = InputBox("Digite a data inicial.", "Entre com a informação")
            If z = "" Then GoTo Data1
            If Not IsDate(z) Then GoTo Data1
Data2:
            z2 = InputBox("Digite a data final.", "Entre com a informação")
            If z2 = "" Then GoTo Data2
            If Not IsDate(z2) Then GoTo Data2

            rpt.RecordSelectionFormula = "{vwdocumentosemitidos.DATADOCUMENTO}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {vwdocumentosemitidos.DATADOCUMENTO}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
            rpt.FormulaFields(1).Text = "'" & Format(z, "dd/mm/yyyy") & " e " & Format(z2, "dd/mm/yyyy") & "'"
    Case "DOCEMITIDO"
Data3:
            z = InputBox("Digite a data inicial.", "Entre com a informação")
            If z = "" Then GoTo Data3
            If Not IsDate(z) Then GoTo Data3
Data4:
            z2 = InputBox("Digite a data final.", "Entre com a informação")
            If z2 = "" Then GoTo Data4
            If Not IsDate(z2) Then GoTo Data4
            frmReport.Caption = "Documentos emitidos por usuário"
'            On Error Resume Next
            rpt.RecordSelectionFormula = "{Comando.DATADOCUMENTO}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {Comando.DATADOCUMENTO}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
 '           rpt.RecordSelectionFormula = "{Command.DATADOCUMENTO}>=#" & CDate(Format(Z, "mm/dd/yyyy")) & "# AND {Command.DATADOCUMENTO}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
'            On Error GoTo Erro
            rpt.FormulaFields(1).Text = "'" & Format(z, "dd/mm/yyyy") & "'"
            rpt.FormulaFields(2).Text = "'" & Format(z2, "dd/mm/yyyy") & "'"
    Case "GARE"
            frmReport.Caption = "Impressão de GARE"
            rpt.FormulaFields(1).Text = "'" & frmGare.lblRequerente.Caption & "'"
            rpt.FormulaFields(2).Text = "'" & frmGare.lblEndereco.Caption & "'"
            rpt.FormulaFields(3).Text = "'" & frmGare.lblCidade.Caption & "'"
            rpt.FormulaFields(4).Text = "'" & frmGare.lblUF.Caption & "'"
            rpt.FormulaFields(5).Text = "'" & frmGare.txtVencto.Text & "'"
            rpt.FormulaFields(6).Text = "'" & frmGare.txtValor.Text & "'"
            rpt.FormulaFields(7).Text = "'" & Mask(frmGare.txtNumExec.Text) & "'"
            rpt.FormulaFields(8).Text = "'" & frmGare.lblCPF.Caption & "'"
            rpt.FormulaFields(9).Text = "'" & frmGare.txtCod.Text & "'"
            rpt.FormulaFields(10).Text = "'" & Mask(frmGare.txtExecutado.Text) & "'"
    Case "DECA2"
            frmReport.Caption = "Impressão de DECA verso"
            rpt.FormulaFields(1).Text = "'" & frmDeca.txtNomeP(0).Text & "'"
            rpt.FormulaFields(2).Text = "'" & frmDeca.txtRuaP(0).Text & "'"
            rpt.FormulaFields(3).Text = "'" & frmDeca.txtBairroP(0).Text & "'"
            rpt.FormulaFields(4).Text = "'" & frmDeca.txtRGP(0).Text & "'"
            rpt.FormulaFields(5).Text = "'" & frmDeca.txtCPFP(0).Text & "'"
            rpt.FormulaFields(6).Text = "'" & frmDeca.txtNomeP(1).Text & "'"
            rpt.FormulaFields(7).Text = "'" & frmDeca.txtRuaP(1).Text & "'"
            rpt.FormulaFields(8).Text = "'" & frmDeca.txtBairroP(1).Text & "'"
            rpt.FormulaFields(9).Text = "'" & frmDeca.txtRGP(1).Text & "'"
            rpt.FormulaFields(10).Text = "'" & frmDeca.txtCPFP(1).Text & "'"
            rpt.FormulaFields(11).Text = "'" & frmDeca.txtNomeP(2).Text & "'"
            rpt.FormulaFields(12).Text = "'" & frmDeca.txtRuaP(2).Text & "'"
            rpt.FormulaFields(13).Text = "'" & frmDeca.txtBairroP(2).Text & "'"
            rpt.FormulaFields(14).Text = "'" & frmDeca.txtRGP(2).Text & "'"
            rpt.FormulaFields(15).Text = "'" & frmDeca.txtCPFP(2).Text & "'"
            rpt.FormulaFields(16).Text = "'" & frmDeca.txtNomeP(3).Text & "'"
            rpt.FormulaFields(17).Text = "'" & frmDeca.txtRuaP(3).Text & "'"
            rpt.FormulaFields(18).Text = "'" & frmDeca.txtBairroP(3).Text & "'"
            rpt.FormulaFields(19).Text = "'" & frmDeca.txtRGP(3).Text & "'"
            rpt.FormulaFields(20).Text = "'" & frmDeca.txtCPFP(3).Text & "'"
            rpt.FormulaFields(21).Text = "'" & frmDeca.txtNomeP(7).Text & "'"
            rpt.FormulaFields(22).Text = "'" & frmDeca.txtRuaP(7).Text & "'"
            rpt.FormulaFields(23).Text = "'" & frmDeca.txtBairroP(7).Text & "'"
            rpt.FormulaFields(24).Text = "'" & frmDeca.txtRGP(7).Text & "'"
            rpt.FormulaFields(25).Text = "'" & frmDeca.txtCPFP(7).Text & "'"
            rpt.FormulaFields(26).Text = "'" & frmDeca.txtNomeP(6).Text & "'"
            rpt.FormulaFields(27).Text = "'" & frmDeca.txtRuaP(6).Text & "'"
            rpt.FormulaFields(28).Text = "'" & frmDeca.txtBairroP(6).Text & "'"
            rpt.FormulaFields(29).Text = "'" & frmDeca.txtRGP(6).Text & "'"
            rpt.FormulaFields(30).Text = "'" & frmDeca.txtCPFP(6).Text & "'"
            rpt.FormulaFields(31).Text = "'" & frmDeca.txtNomeP(5).Text & "'"
            rpt.FormulaFields(32).Text = "'" & frmDeca.txtRuaP(5).Text & "'"
            rpt.FormulaFields(33).Text = "'" & frmDeca.txtBairroP(5).Text & "'"
            rpt.FormulaFields(34).Text = "'" & frmDeca.txtRGP(5).Text & "'"
            rpt.FormulaFields(35).Text = "'" & frmDeca.txtCPFP(5).Text & "'"
            rpt.FormulaFields(36).Text = "'" & frmDeca.txtNomeP(4).Text & "'"
            rpt.FormulaFields(37).Text = "'" & frmDeca.txtRuaP(4).Text & "'"
            rpt.FormulaFields(38).Text = "'" & frmDeca.txtBairroP(4).Text & "'"
            rpt.FormulaFields(39).Text = "'" & frmDeca.txtRGP(4).Text & "'"
            rpt.FormulaFields(40).Text = "'" & frmDeca.txtCPFP(4).Text & "'"
            rpt.FormulaFields(41).Text = "'" & frmDeca.txtNomeC.Text & "'"
            rpt.FormulaFields(42).Text = "'" & frmDeca.txtEndC.Text & "'"
            rpt.FormulaFields(43).Text = "'" & frmDeca.txtBairroC.Text & "'"
            rpt.FormulaFields(44).Text = "'" & frmDeca.txtFoneC.Text & "'"
            rpt.FormulaFields(45).Text = "'" & frmDeca.txtnumC.Text & "'"
            rpt.FormulaFields(46).Text = "'" & frmDeca.txtCEPC.Text & "'"
            rpt.FormulaFields(47).Text = "'" & frmDeca.txtRGC.Text & "'"
            rpt.FormulaFields(48).Text = "'" & frmDeca.txtOrgaoC.Text & "'"
            rpt.FormulaFields(49).Text = "'" & Mask(frmDeca.txtOBSC.Text) & "'"
            rpt.FormulaFields(58).Text = "'" & Mask(frmDeca.txtEmail.Text) & "'"
            rpt.FormulaFields(59).Text = "'" & Mask(frmDeca.txtTelefone(0).Text) & "'"
            rpt.FormulaFields(60).Text = "'" & Mask(frmDeca.txtTelefone(1).Text) & "'"
            rpt.FormulaFields(61).Text = "'" & Mask(frmDeca.txtTelefone(2).Text) & "'"
            rpt.FormulaFields(62).Text = "'" & Mask(frmDeca.txtTelefone(3).Text) & "'"
            rpt.FormulaFields(63).Text = "'" & Mask(frmDeca.txtTelefone(4).Text) & "'"
            rpt.FormulaFields(64).Text = "'" & Mask(frmDeca.txtTelefone(5).Text) & "'"
            rpt.FormulaFields(65).Text = "'" & Mask(frmDeca.txtTelefone(6).Text) & "'"
            rpt.FormulaFields(66).Text = "'" & Mask(frmDeca.txtTelefone(7).Text) & "'"
            rpt.FormulaFields(67).Text = "'" & Mask(frmDeca.txtCidadeC.Text) & "'"
            rpt.FormulaFields(68).Text = "'" & Mask(frmDeca.txtUFC.Text) & "'"
            
    Case "MULTAINF"
            frmReport.Caption = "Multa de Infração 2ª via"
            
            z = InputBox("Digite o número do processo")
            
            rpt.FormulaFields(1).Text = "'" & z & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmDebitoImob.lblProp.Caption) & "'"
            rpt.FormulaFields(3).Text = "'" & Format(frmDebitoImob.txtCod.Text, "000000") & "-" & RetornaDVCodReduzido(CLng(frmDebitoImob.txtCod.Text)) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmDebitoImob.lblRua.Caption) & "'"
            Sql = "SELECT DT_AREATERRENO FROM CADIMOB WHERE CODREDUZIDO=" & Val(frmDebitoImob.txtCod.Text)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                rpt.FormulaFields(5).Text = "'" & FormatNumber(!Dt_AreaTerreno, 2) & "'"
'                If !Dt_AreaTerreno <= 125 Then
'                    rpt.FormulaFields(6).Text = "'I'"
'                    rpt.FormulaFields(7).Text = 50
'                ElseIf !Dt_AreaTerreno > 125 And !Dt_AreaTerreno <= 250 Then
'                    rpt.FormulaFields(6).Text = "'II'"
'                    rpt.FormulaFields(7).Text = 100
'                ElseIf !Dt_AreaTerreno > 250 And !Dt_AreaTerreno <= 500 Then
'                    rpt.FormulaFields(6).Text = "'III'"
'                    rpt.FormulaFields(7).Text = 200
'                ElseIf !Dt_AreaTerreno > 500 Then
'                    rpt.FormulaFields(6).Text = "'IV'"
'                    rpt.FormulaFields(7).Text = 300
'                End If
'                rpt.FormulaFields(8).Text = "'" & FormatNumber(!Dt_AreaTerreno * 0.72, 2) & "'"
                rpt.FormulaFields(6).Text = "''"
                rpt.FormulaFields(7).Text = "'500,00'"
                rpt.FormulaFields(8).Text = "'" & FormatNumber(!Dt_AreaTerreno * 0.8883, 2) & "'"
               
               .Close
            End With
    Case "MULTAINF2"
            frmReport.Caption = "Multa de Infração 2ª via"
            Liberado
            z = InputBox("Digite o valor da taxa de serviço", "Informação requerida")
            If Val(z) > 0 Then
                rpt.FormulaFields(8).Text = "'" & z & "'"
            Else
                MsgBox "Valor inválido", vbCritical, "Erro"
                Exit Function
                rpt.FormulaFields(8).Text = "'0'"
            End If
            
            
            z = InputBox("Digite o número do processo")
            If z = "" Then Exit Function
            rpt.FormulaFields(1).Text = "'" & z & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmDebitoImob.lblProp.Caption) & "'"
            rpt.FormulaFields(3).Text = "'" & Format(frmDebitoImob.txtCod.Text, "000000") & "-" & RetornaDVCodReduzido(CLng(frmDebitoImob.txtCod.Text)) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmDebitoImob.lblRua.Caption) & "'"
            Sql = "SELECT DT_AREATERRENO FROM CADIMOB WHERE CODREDUZIDO=" & Val(frmDebitoImob.txtCod.Text)
            Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
            With RdoAux
                If RdoAux.RowCount = 0 Then
                    MsgBox "Imóvel não localizado", vbCritical, "Erro"
                    Exit Function
                End If
                rpt.FormulaFields(5).Text = "'" & FormatNumber(!Dt_AreaTerreno, 2) & "'"
 '               If !Dt_AreaTerreno <= 125 Then
 '                   rpt.FormulaFields(6).Text = "'I'"
 '                   rpt.FormulaFields(7).Text = "65.75"
 '               ElseIf !Dt_AreaTerreno > 125 And !Dt_AreaTerreno <= 250 Then
 '                   rpt.FormulaFields(6).Text = "'II'"
 '                   rpt.FormulaFields(7).Text = "164.37"
 '               ElseIf !Dt_AreaTerreno > 250 And !Dt_AreaTerreno <= 500 Then
 '                   rpt.FormulaFields(6).Text = "'III'"
 '                   rpt.FormulaFields(7).Text = "262.99"
 '               ElseIf !Dt_AreaTerreno > 500 Then
 '                   rpt.FormulaFields(6).Text = "'IV'"
 '                   rpt.FormulaFields(7).Text = "394.49"
 '               End If
                 rpt.FormulaFields(6).Text = "''"
                rpt.FormulaFields(7).Text = "'500,00'"
'                rpt.FormulaFields(8).Text = "'" & FormatNumber(!Dt_AreaTerreno * 0.8883, 2) & "'"

               .Close
            End With
    Case "ALVARA", "ALVARASEMDATA", "ALVARAVICE", "ALVARASEMDATAVICE"
            Liberado
            If frmAlvara.cmbDataAlvara.ListIndex = 0 Then
                z6 = Year(Now)
            Else
                z6 = Year(Now) + 1
            End If
           
           Select Case frmAlvara.cmbTipo.ListIndex
               Case 0
                  z3 = "N"
               Case 1
                  z3 = "B"
               Case 2
                  z3 = "V"
               Case 3
                  z3 = "BV"
               Case 4
                  z3 = "P"
           End Select
           
           frmReport.Caption = "Alvará de Funcionamento "
           Sql = "SELECT * FROM VWCNSMOBILIARIO WHERE CODIGOMOB=" & Val(frmAlvara.txtCodigo.Text)
           Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
           With RdoAux
               sHor = Mask(SubNull(!HORARIOEXT))
               If sHor = "" Then
                    ql = "SELECT DESCHORARIO FROM HORARIOFUNC WHERE CODHORARIO=" & !HORARIO
                    Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux2
                         If .RowCount > 0 Then
                            sHor = SubNull(!DESCHORARIO)
                         End If
                        .Close
                    End With
               End If
               sTexto1 = sHor
               sTexto3 = z2
               
               If (frmAlvara.cmbAss.ListIndex > 0) Then
                    Sql = "SELECT USUARIO FROM ASSINATURA WHERE NOME='" & frmAlvara.cmbAss.Text & "'"
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    With RdoAux
                         rpt.RecordSelectionFormula = "{ASSINATURA.USUARIO}='" & !USUARIO & "'"
                        .Close
                    End With
               Else
                    rpt.RecordSelectionFormula = "{ASSINATURA.USUARIO}='NOBODY'"
               End If
               rpt.FormulaFields(1).Text = "'" & Mask(sTexto1) & "'"
               rpt.FormulaFields(2).Text = "'" & Mask(sTexto2) & "'"
               rpt.FormulaFields(3).Text = "'" & frmAlvara.txtProcesso2.Text & "'"
               rpt.FormulaFields(4).Text = "'" & Mask(frmAlvara.lblNome.Caption) & "'"
               rpt.FormulaFields(10).Text = "'" & z3 & "'"
               rpt.FormulaFields(11).Text = "'" & frmAlvara.mskDataBomb.Text & "'"
               rpt.FormulaFields(12).Text = "'" & frmAlvara.mskDataVS.Text & "'"
               
               If z3 = "P" Then
                    rpt.FormulaFields(13).Text = "'" & z2 & "'"
               Else
                    rpt.FormulaFields(13).Text = "''"
               End If
               rpt.FormulaFields(5).Text = "'" & Mask(frmAlvara.lblEndereco.Caption & ", " & frmAlvara.lblNum.Caption & " " & frmAlvara.lblCompl.Caption) & "'"
               rpt.FormulaFields(7).Text = "'" & Mask(frmAlvara.lblAtividade.Caption) & "'"
               rpt.FormulaFields(8).Text = "'" & frmAlvara.txtCodigo.Text & "'"
               rpt.FormulaFields(9).Text = "'" & frmAlvara.lblCep.Caption & "'"
               rpt.FormulaFields(6).Text = "'" & Mask(frmAlvara.lblBairro.Caption) & "'"
               rpt.FormulaFields(14).Text = "'" & Mask(frmAlvara.lblCidade.Caption) & "'"
               If frmAlvara.mskCNPJ.ClipText <> "" Then
                   z2 = frmAlvara.mskCNPJ.Text
               Else
                   z2 = frmAlvara.mskCPF.Text
               End If
               rpt.FormulaFields(15).Text = "'" & z2 & "'"
               rpt.FormulaFields(16).Text = "'" & sTr(z6) & "'"
               
               If frmAlvara.cmbTipo.ListIndex = 4 Then
                   rpt.FormulaFields(17).Text = "'1'"
               Else
                   rpt.FormulaFields(17).Text = "''"
               End If
               rpt.FormulaFields(19).Text = "'" & frmAlvara.mskDataSaaej.Text & "'"
               rpt.FormulaFields(20).Text = "'" & frmAlvara.mskDataCETESB.Text & "'"
               rpt.FormulaFields(18).Text = "'" & IIf(frmAlvara.chkPrefeito.value = vbChecked, "A", "B") & "'"
               rpt.FormulaFields(21).Text = "'" & IIf(IsDate(frmAlvara.mskDataBomb.Text), "S", "N") & "'"
               rpt.FormulaFields(22).Text = "'" & IIf(IsDate(frmAlvara.mskDataVS.Text), "S", "N") & "'"
               rpt.FormulaFields(23).Text = "'" & IIf(IsDate(frmAlvara.mskDataSaaej.Text), "S", "N") & "'"
               rpt.FormulaFields(24).Text = "'" & IIf(IsDate(frmAlvara.mskDataCETESB.Text), "S", "N") & "'"
               rpt.FormulaFields(25).Text = "'" & IIf(frmAlvara.chk24Hrs.value = vbChecked, "S", "N") & "'"
               rpt.FormulaFields(26).Text = "'" & IIf(frmAlvara.chkBombon.value = vbChecked, "S", "N") & "'"
               rpt.FormulaFields(27).Text = "'" & Mask(frmAlvara.txtObs.Text) & "'"
               If sReport = "ALVARASEMDATA" Or sReport = "ALVARASEMDATAVICE" Then
                 rpt.FormulaFields(28).Text = "'" & Mask(frmAlvara.txtData.Text) & "'"
               End If
               Sql = "select placa from mobiliarioplaca where codigo=" & Val(frmAlvara.txtCodigo.Text)
               Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
               Dim sPlaca As String
'               If RdoAux2.RowCount > 0 Then
                   'sPlaca = RdoAux2!PLACA
                   Do Until RdoAux2.EOF
                     sPlaca = sPlaca & RdoAux2!PLACA & ", "
                    RdoAux2.MoveNext
                   Loop
 '              End If
               RdoAux2.Close
               If Len(sPlaca) > 0 Then sPlaca = Left(sPlaca, Len(sPlaca) - 2)
               rpt.FormulaFields(29).Text = "'" & sPlaca & "'"
               rpt.FormulaFields(30).Text = "'" & Mask(frmAlvara.lblPontoAgencia.Caption) & "'"
               rpt.FormulaFields(31).Text = "'" & frmAlvara.lblIE.Caption & "'"
              .Close
            End With
    Case "ISSPAGOPERIODO"
            frmReport.Caption = "Iss pago por atividade"
            CRViewer1.EnableGroupTree = True
            rpt.FormulaFields(1).Text = "'" & Mask(frmIssPagoAtividade.dtDataDe.value) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmIssPagoAtividade.dtDataAte.value) & "'"
            rpt.RecordSelectionFormula = "{vwDebitoPago4.datavencimento} >=#" & Format(frmIssPagoAtividade.dtDataDe.value, "mm/dd/yyyy") & "# and {vwDebitoPago4.datavencimento} <=#" & Format(frmIssPagoAtividade.dtDataAte.value, "mm/dd/yyyy") & "#"
            Liberado
    Case "SENHAISS3"
            frmReport.Caption = "Requerimento de senha para ISS Eletrônico"
            Liberado
            z = InputBox("Digite o Código da Empresa.", "Código da Empresa")
            If z = "" Then Exit Function
            If Val(z) = 0 Then
                MsgBox "Código Inválido.", vbExclamation, "Atenção"
                Exit Function
            Else
                Sql = "select * from mobiliarioatividadeiss where codmobiliario=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                If RdoAux.RowCount = 0 Then
                     MsgBox "Esta empresa não possui enquadramento de atividade de ISS.", vbExclamation, "Atenção"
                     Exit Function
                End If
                Sql = "SELECT * FROM vwFULLEMPRESA WHERE CODIGOMOB=" & Val(z)
                Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                With RdoAux
                    If .RowCount = 0 Then
                        MsgBox "Empresa não Cadastrada.", vbExclamation, "Atenção"
                        Exit Function
                    Else
                        If Not IsNull(!dataencerramento) And !dataencerramento <> CDate("01/01/1900") Then
                           MsgBox "Esta empresa foi encerrada em " & Format(!dataencerramento, "dd/mm/yyyy"), vbExclamation, "Atenção"
                           Exit Function
                        End If
                        sSenha = SubNull(!SENHAISS)
                        If sSenha = "" Then
                           MsgBox "Esta empresa não possui senha cadastrada.", vbExclamation, "Atenção"
                           Exit Function
                        End If
                    End If
                End With
            End If
           
           z2 = InputBox("Digite o nome do assinante.", "Informe")
           z3 = InputBox("Digite o RG do assinante.", "Informe")
           z4 = InputBox("Digite o número do processo.", "Informe")
            If MsgBox("Empresa: " & z & vbCrLf & "Nome: " & z2 & vbCrLf & "RG: " & z3 & vbCrLf & "Processo nº: " & z4, vbQuestion + vbYesNo, "Confirme os Dados") = vbYes Then
                rpt.FormulaFields(1).Text = "'" & RetornaUsuarioFullName & "'"
                rpt.FormulaFields(2).Text = "'" & z4 & "'"
                rpt.FormulaFields(3).Text = "'" & Format(RdoAux!codigomob, "000000") & "'"
                rpt.FormulaFields(4).Text = "'" & RdoAux!RazaoSocial & "'"
                rpt.FormulaFields(5).Text = "'" & RdoAux!Logradouro & ", " & Val(SubNull(RdoAux!Numero)) & " Bairro: " & SubNull(RdoAux!DescBairro) & " - " & SubNull(RdoAux!descCidade) & "/ " & SubNull(RdoAux!SiglaUF) & " CEP: " & RetornaCEP(RdoAux!CodLogradouro, RdoAux!Numero) & "'"
                rpt.FormulaFields(6).Text = "'" & RdoAux!ativextenso & "'"
                rpt.FormulaFields(7).Text = "'IM00" & Format(RdoAux!codigomob, "000000") & "'"
                rpt.FormulaFields(8).Text = "'" & sSenha & "'"
                rpt.FormulaFields(9).Text = "'" & z3 & "'"
                rpt.FormulaFields(10).Text = "'" & z2 & "'"
            Else
               Exit Function
            End If
    Case "GUIAPRATICO5"
            frmReport.Caption = "Guias diversas"
            rpt.FormulaFields(1).Text = "'" & Mask(RetornaUsuarioFullName) & "'"
            rpt.FormulaFields(2).Text = "'" & Mask(frmGuiaPratico5.txtAno.Text) & "'"
            rpt.FormulaFields(3).Text = "'" & Mask(frmGuiaPratico5.txtNot.Text) & "'"
            rpt.FormulaFields(4).Text = "'" & Mask(frmGuiaPratico5.txtProc.Text) & "'"
            rpt.FormulaFields(5).Text = "'" & Mask(frmGuiaPratico5.txtNome.Text) & "'"
            rpt.FormulaFields(6).Text = "'" & Mask(frmGuiaPratico5.txtCod.Text) & "'"
            If frmGuiaPratico5.cmbPag.ListIndex = 0 Then
                z = "à vista."
            Else
                z = "parcelado em " & frmGuiaPratico5.txtParc.Text & " vezes."
            End If
            rpt.FormulaFields(7).Text = "'" & z & "'"
            rpt.FormulaFields(8).Text = "'" & Mask(frmGuiaPratico5.txtPerc.Text) & "'"
    Case "PROCESSOENVIADO2"
            frmReport.Caption = "Processos tramitados por Centro de Custo"
            rpt.FormulaFields(1).Text = "'" & frmProcessosEnviados.mskData.Text & " e " & frmProcessosEnviados.mskData2.Text & "'"
            rpt.RecordSelectionFormula = "{PROCESSOENVIO.COMPUTER}='" & NomeDeLogin & "'"
    Case "SITUACAOTRIBUTO"
            frmReport.Caption = "Situação dos tributos lançados"
            rpt.FormulaFields(1).Text = "'" & frmSituacaoTributo.mskDataIni.Text & "'"
            rpt.FormulaFields(2).Text = "'" & frmSituacaoTributo.mskDataFim.Text & "'"
            rpt.RecordSelectionFormula = "{RELSITUACAOTRIBUTO.USUARIO}='" & NomeDeLogin & "'"
    Case "ALVARARENOVA", "ALVARARENOVAVICE"
            z = InputBox("Digite o Código da Empresa.", "Código da Empresa")
            If z = "" Then Exit Function
            If Val(z) = 0 Then
                MsgBox "Código Inválido.", vbExclamation, "Atenção"
                Exit Function
            End If
            
            On Error Resume Next
            RdoAux.Close
            On Error GoTo 0
            frmReport.Caption = "Renovação de Alvará"
            Set qd.ActiveConnection = cn
            qd.QueryTimeout = 0
            qd.Sql = "{ Call spALVARA(?) }"
            qd(0) = z
            Set RdoAux = qd.OpenResultset(rdOpenKeyset)
            If RdoAux!Tipo = 0 Then
                MsgBox "Empresa inválida ou que não possui renovação automática de alvará.", vbCritical, "Atenção"
                Liberado
                Exit Function
            Else
                If RdoAux!Tipo = 1 Then
                    rpt.FormulaFields(1).Text = "'" & RdoAux!RazaoSocial & "'"
                    sDoc = Format(SubNull(RdoAux!Cnpj), "00\.000\.000/0000-00")
                    If sDoc = "" Then
                        sDoc = Format(SubNull(RdoAux!CPF), "000\.000\.000-00")
                    End If
                    rpt.FormulaFields(2).Text = "'" & sDoc & "'"
                    rpt.FormulaFields(3).Text = "'" & RdoAux!Logradouro & ", " & RdoAux!Numero & "'"
                    rpt.FormulaFields(4).Text = "'" & SubNull(RdoAux!Bairro) & "'"
                    rpt.FormulaFields(5).Text = "'" & SubNull(RdoAux!Cep) & "'"
                    rpt.FormulaFields(6).Text = "'" & SubNull(RdoAux!Cidade) & "'"
                    rpt.FormulaFields(7).Text = "'" & SubNull(RdoAux!codigomob) & "'"
                    rpt.FormulaFields(8).Text = "'" & SubNull(RdoAux!Atividade) & "'"
                    rpt.FormulaFields(9).Text = "'" & SubNull(RdoAux!HORARIO) & "'"
                    rpt.FormulaFields(10).Text = "'" & SubNull(RdoAux!CHAVE) & "'"
                    rpt.FormulaFields(11).Text = "'" & SubNull(RdoAux!ANOALVARA) & "'"
                    nCodReduz = RdoAux!codigomob
                    RdoAux.Close
                    sTexto1 = "Emisão de Renovação de Alvará."
                    Sql = "SELECT MAX(SEQ) AS MAXIMO FROM MOBILIARIOHIST WHERE CODMOBILIARIO=" & nCodReduz
                    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
                    If IsNull(RdoAux!maximo) Then
                        nSeq = 0
                    Else
                        nSeq = RdoAux!maximo + 1
                    End If
                                
                    Sql = "INSERT MOBILIARIOHIST(CODMOBILIARIO,SEQ,DATAHIST,OBS,USERID) VALUES("
                    Sql = Sql & nCodReduz & "," & nSeq & ",'" & Format(Now, "mm/dd/yyyy") & "','" & Mask(sTexto1) & "'," & RetornaUsuarioID(NomeDeLogin) & ")"
                    cn.Execute Sql, rdExecDirect
                Else
                    MsgBox "Esta empresa não esta com todas as condições necessárias para renovação do alvará.", vbCritical, "Atenção"
                    Liberado
                    Exit Function
                End If
            End If
            
    Case "BOLETODAM", "BOLETODAMTESTE", "BOLETODAM_V4", "BOLETODAM_V4TMP", "BOLETODAM_V4TMP2", "BOLETODAM_V3", "BOLETODAM_V5"
        frmReport.Caption = "Impressão de DAM - Boleto"
        rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND {BOLETO.SID}=" & nNumDoc
        If UCase(sReport) = "BOLETODAM" Or UCase(sReport) = "BOLETODAMTESTE" Or UCase(sReport) = "BOLETODAM_V4" Or UCase(sReport) = "BOLETODAM_V3" Or UCase(sReport) = "BOLETODAM_V5" Then
            For x = 0 To Forms.Count - 1
                If Forms(x).Name = "frmDAM" Then
                    If Forms(x).Visible Then
                        bDam = True
                        If frmDAM.chkCobranca.value = vbChecked Then
                            rpt.FormulaFields(1).Text = "'S'"
                        Else
                            rpt.FormulaFields(1).Text = "'N'"
                        End If
                    End If
                   Exit For
                End If
            Next
            If frmDAM.Honorarios = True Then
                rpt.FormulaFields(2).Text = "'S'"
            Else
                rpt.FormulaFields(2).Text = "'N'"
            End If
        End If
    Case "BOLETOGUIA", "BOLETOGUIATMP", "BOLETOGUIA_V4", "BOLETOGUIA_V4TMP", "BOLETOGUIA_V4TMP2", "BOLETOGUIA_CIP"
        frmReport.Caption = "Impressão de guias - Boleto"
            If hFrmCall <> frmConfissaoDivida.hwnd Then
                rpt.RecordSelectionFormula = "{BOLETOGUIA.SID}=" & nNumDoc
            Else
                rpt.RecordSelectionFormula = "{BOLETOGUIA.SID}=" & nNumDoc & " AND {BOLETOGUIA.NUMPARCELA}=1"
            End If
            frmConfissaoDivida.Hide
    Case "BOLETOCOBRANCA", "BOLETOCOBRANCA_V4"
        frmReport.Caption = "Impressão de guias - Boleto de Cobrança"
        rpt.RecordSelectionFormula = "{BOLETO.SID}=" & nNumDoc
    Case "BOLETOCIP"
        frmReport.Caption = "Impressão de guias - Boleto CIP"
        rpt.RecordSelectionFormula = "{BOLETOGUIA.SID}=" & nNumDoc
        
    Case "CARNEINDIVIDUAL"
        frmReport.Caption = "Impressão de Carnê"
        rpt.RecordSelectionFormula = "{CARNETMP.COMPUTER}='" & NomeDoUsuario & "'"
            
    Case "MOBILIARIODEVEDOR"
            rpt.RecordSelectionFormula = "{MOBILIARIODEVEDOR.USUARIO}='" & NomeDeLogin & "'"
    
    Case "NOTIFICACAO3", "NOTIFICACAO4"
            If UCase(NomeDeLogin) = "GLEISE" Then
                rpt.RecordSelectionFormula = "{NOTIFICACAOISS.USUARIO}='EDUARDO'"
            Else
                rpt.RecordSelectionFormula = "{NOTIFICACAOISS.USUARIO}='" & NomeDeLogin & "'"
            End If
    Case "REQUERIPTU"
            rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
            frmReport.Caption = "Requerimento para Isenção de IPTU"
            If frmRequerIPTU.optTipo(0).value = True Then
                z = "isenção de IPTU"
                rpt.FormulaFields(4).Text = "'ISENÇÃO DE IPTU'"
            Else
                z = "renovação de isenção de IPTU"
                rpt.FormulaFields(4).Text = "'RENOVAÇÃO DE ISENÇÃO DE IPTU'"
            End If
            
            
            If frmRequerIPTU.Opt(0).value = True Then
                sTexto1 = frmRequerIPTU.lblRequerente.Caption & " RG " & frmRequerIPTU.lblRG.Caption & " CPF " & frmRequerIPTU.lblCPF.Caption & " abaixo "
                sTexto1 = sTexto1 & "assinado, residente " & frmRequerIPTU.lblEndereco.Caption & " vem respeitosamente requerer de V.Exa. que seja concedida "
                sTexto1 = sTexto1 & z & " do exercício de " & Year(Now) + 1 & ", para o imóvel código " & frmRequerIPTU.lblCodImovel.Caption & ", "
                sTexto1 = sTexto1 & "situado à " & frmRequerIPTU.lblEndImovel.Caption & " nº " & frmRequerIPTU.lblNumImovel.Caption & ", bairro " & frmRequerIPTU.lblBairroImovel.Caption & ", "
                sTexto1 = sTexto1 & "com fundamento na Lei Complementar 07/92, art. 50, inciso VII, regulamentada pelo Decreto 5569/2010, apresentando para essa finalidade a documentação exigida. Anexar Processo de Avaliação Social nº " & frmRequerIPTU.txtNumProc1.Text
                rpt.FormulaFields(1).Text = "'" & frmRequerIPTU.lblRequerente.Caption & "'"
            Else
                sTexto1 = frmRequerIPTU.lblRazao.Caption & " inscrito no CNPJ " & frmRequerIPTU.lblCNPJ.Caption & " estabelecido à " & frmRequerIPTU.txtEnd.Text & ", neste ato "
                sTexto1 = sTexto1 & "representado por " & Mask(frmRequerIPTU.txtRepresentante.Text) & ",RG " & frmRequerIPTU.txtRG.Text & ", CPF " & frmRequerIPTU.txtCPF.Text & ", "
                sTexto1 = sTexto1 & "vem respeitosamente requerer de V.Exa. que seja concedida " & z & " do exercício de " & Year(Now) + 1 & ", para o imóvel código " & frmRequerIPTU.lblCodImovel.Caption & ", "
                sTexto1 = sTexto1 & "situado à " & frmRequerIPTU.lblEndImovel.Caption & " nº " & frmRequerIPTU.lblNumImovel.Caption & ", bairro " & frmRequerIPTU.lblBairroImovel.Caption & ", "
                sTexto1 = sTexto1 & "com fundamento na Lei Complementar 07/92, art. 50, regulamentada pelo Decreto 5569/2010, apresentando para essa finalidade a documentação exigida. Processo Anterior nº " & frmRequerIPTU.txtNumProc2.Text
                rpt.FormulaFields(1).Text = "'" & frmRequerIPTU.lblRazao.Caption & "'"
            End If
            
            rpt.FormulaFields(2).Text = "'" & sTexto1 & "'"
            If frmRequerIPTU.chkObs.value = 0 Then
                rpt.FormulaFields(3).Text = "'N'"
            Else
                rpt.FormulaFields(3).Text = "'S'"
            End If
    
    Case "GUIAPRATICO3"
        frmReport.Caption = "Guias diversas"
        rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
        rpt.FormulaFields(1).Text = "'" & Mask(frmGuiaPratico3.lblRequerente.Caption) & "'"
        rpt.FormulaFields(2).Text = "'" & Mask(frmGuiaPratico3.lblRG.Caption) & "'"
        rpt.FormulaFields(3).Text = "'" & Mask(frmGuiaPratico3.lblCPF.Caption) & "'"
        rpt.FormulaFields(4).Text = "'" & Mask(frmGuiaPratico3.lblEndereco.Caption) & "'"
        rpt.FormulaFields(5).Text = "'" & Mask(frmGuiaPratico3.lblNum.Caption) & "'"
        rpt.FormulaFields(6).Text = "'" & Mask(frmGuiaPratico3.lblBairro.Caption) & "'"
        rpt.FormulaFields(7).Text = "'" & Mask(frmGuiaPratico3.lblFone.Caption) & "'"
        rpt.FormulaFields(8).Text = "'" & Mask(frmGuiaPratico3.txtMarca.Text) & "'"
        rpt.FormulaFields(9).Text = "'" & Mask(frmGuiaPratico3.txtModelo.Text) & "'"
        rpt.FormulaFields(10).Text = "'" & Mask(frmGuiaPratico3.txtAno.Text) & "'"
        rpt.FormulaFields(11).Text = "'" & Mask(frmGuiaPratico3.txtCor.Text) & "'"
        rpt.FormulaFields(12).Text = "'" & Mask(frmGuiaPratico3.txtRenavam.Text) & "'"
        rpt.FormulaFields(13).Text = "'" & Mask(frmGuiaPratico3.txtPlaca.Text) & "'"
        rpt.FormulaFields(14).Text = "'" & IIf(frmGuiaPratico3.chk(0).value = vbChecked, "X", " ") & "'"
        rpt.FormulaFields(15).Text = "'" & IIf(frmGuiaPratico3.chk(1).value = vbChecked, "X", " ") & "'"
        rpt.FormulaFields(16).Text = "'" & IIf(frmGuiaPratico3.chk(2).value = vbChecked, "X", " ") & "'"
    Case "GUIAPRATICO4"
        On Error Resume Next
        frmReport.Caption = "Guias diversas"
        rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
        rpt.FormulaFields(1).Text = "'" & Mask(frmGuiaPratico4.lblRequerente.Caption) & "'"
        rpt.FormulaFields(2).Text = "'" & Mask(frmGuiaPratico4.lblEndereco.Caption) & "'"
        rpt.FormulaFields(3).Text = "'" & Mask(frmGuiaPratico4.txtCPF.Text) & "'"
        
        sTexto1 = frmGuiaPratico4.txtExp.Text
        If Left(sTexto1, 10) = "Artigo 150" Then
            sTexto1 = "Aquisição de imóvel por templo religioso"
        Else
            x = InStr(1, sTexto1, "Inciso", vbBinaryCompare)
            Y = InStr(x, sTexto1, "-", vbBinaryCompare)
            sTexto1 = Mid(sTexto1, Y + 2, Len(sTexto1) - Y - 1)
        End If
        rpt.FormulaFields(4).Text = "'" & Mask(sTexto1) & "'"
        rpt.FormulaFields(11).Text = "'" & IIf(frmGuiaPratico4.cmbAss.ListIndex = 0, "A", "B") & "'"
        rpt.FormulaFields(5).Text = "'" & Mask(frmGuiaPratico4.txtValor.Text) & " (" & Extenso(frmGuiaPratico4.txtValor.Text) & ")'"
      rpt.FormulaFields(9).Text = "'" & RetornaUsuarioFullName & "'"
        sTexto1 = frmGuiaPratico4.txtExp.Text
        If sTexto1 = "" Then GoTo FIMPRATICO4
        If Left(sTexto1, 11) = "Artigo 150b" Then
            sTexto1 = "Artigo 150 - Inciso VI, letra ""b"" da Constituição Federal da República Federativa do Brasil, "
        ElseIf Left(sTexto1, 10) = "Artigo 150" Then
            sTexto1 = "Artigo 150 - Inciso VI, letra ""b"" da Constituição Federal da República Federativa do Brasil, combinado com o Artigo 111 - Inciso IV, "
        Else
            x = InStr(1, sTexto1, "Inciso", vbBinaryCompare)
            Y = InStr(x, sTexto1, "-", vbBinaryCompare)
            sTexto1 = Left(sTexto1, Y - 2)
        End If
FIMPRATICO4:
        rpt.FormulaFields(10).Text = "'" & sTexto1 & "'"
    Case "REGATENDIMENTO"
        frmReport.Caption = "Registro de Atendimento"
        rpt.RecordSelectionFormula = "{REGISTROATENDIMENTOTMP.USUARIO}='" & NomeDeLogin & "'"
    Case "REFIS"
        frmReport.Caption = "Relatório do Refis DAM"
DataR1:
            z = InputBox("Digite a data inicial.", "Entre com a informação")
            If z = "" Then GoTo DataR1
            If Not IsDate(z) Then GoTo DataR1
DataR2:
            z2 = InputBox("Digite a data final.", "Entre com a informação")
            If z2 = "" Then GoTo DataR2
            If Not IsDate(z2) Then GoTo DataR2
    '        rpt.RecordSelectionFormula = "{vwrefisnovo2.datapagamento}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {vwrefisnovo2.datapagamento}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
            GeraRefisDam CStr(z), CStr(z2)
            rpt.RecordSelectionFormula = "{relatorio_refis.usuario}='" & NomeDeLogin & "'"
            rpt.FormulaFields(1).Text = "'" & Format(z, "dd/mm/yyyy") & " e " & Format(z2, "dd/mm/yyyy") & "'"
    Case "REFISPARC"
        frmReport.Caption = "Relatório do Refis parcelado"
        rpt.RecordSelectionFormula = "{EXTRATOTMP.COMPUTER}='" & NomeDeLogin & "'"
        rpt.FormulaFields(1).Text = "'2017'"
        GeraRefis (nAno)
    Case "QTDEPROCESSOSANO"
        frmReport.Caption = "Qtde de processos tramitados no ano"
        z = InputBox("Digite o Ano", "Informação requerida")
        If Val(z) > 1980 And Val(z) < 2020 Then
            rpt.RecordSelectionFormula = "{COMMAND.ANO}=" & Val(z)
            rpt.FormulaFields(1).Text = "'" & z & "'"
        Else
            MsgBox "Ano inválido.", vbExclamation, "Atenção"
            Liberado
            Exit Function
        End If
    Case "PRODMOB1"
        frmReport.Caption = "Produtividade Mensal - Fiscal de Tributos"
        rpt.RecordSelectionFormula = "{PRODUTIVIDADEREL1.USUARIO}='" & NomeDeLogin & "'"
        rpt.FormulaFields(1).Text = "'" & frmProdutividadeMensal.cmbFiscal.Text & " (Matrícula nº " & frmProdutividadeMensal.lblMatricula.Caption & ")'"
        rpt.FormulaFields(2).Text = "'" & frmProdutividadeMensal.cmbMes.Text & "/" & frmProdutividadeMensal.cmbAno.Text & "'"
        rpt.FormulaFields(3).Text = "'" & frmProdutividadeMensal.txtTotal.Text & "'"
        rpt.FormulaFields(4).Text = "'" & frmProdutividadeMensal.txtSaldo.Text & "'"
        rpt.FormulaFields(5).Text = "'" & frmProdutividadeMensal.txtPontos.Text & "'"
        rpt.FormulaFields(6).Text = "'" & frmProdutividadeMensal.txtResultado.Text & "'"
        rpt.FormulaFields(7).Text = "'" & frmProdutividadeMensal.txtReceber.Text & "'"
        rpt.FormulaFields(8).Text = "'" & frmProdutividadeMensal.txtTransportar.Text & "'"
    Case "MAIORDEVEDOR"
        frmReport.Caption = "Maiores devedores"
        rpt.RecordSelectionFormula = "{EXTRATOTMP.COMPUTER}='" & NomeDeLogin & "'"
    Case "EXTRATONF"
        frmReport.Caption = "Extrato do ISS Eletrônico"
        rpt.RecordSelectionFormula = "{EXTRATONF.USUARIO}='" & NomeDeLogin & "'"
        rpt.FormulaFields(2).Text = "'" & frmCadMob.txtCodIss.Text & " - " & Mask(frmCadMob.txtNomeISS.Text) & "'"
    Case "NOTIFICACAO2"
        frmReport.Caption = "Notificação de lançamento"
        rpt.RecordSelectionFormula = "{NOTIFICACAO.USUARIO}='" & NomeDeLogin & "'"
    Case "REGATENDIMENTO3"
        frmReport.Caption = "Registro de Atendimento por Equipe"
        If frmRelatObra.cmbEquipe.ListIndex = 0 Then
            rpt.RecordSelectionFormula = "{vwFULLREGATENDIMENTO.DATA}>=#" & CDate(Format(frmRelatObra.mskDataIni.Text, "mm/dd/yyyy")) & "# AND {vwFULLREGATENDIMENTO.DATA}<=#" & CDate(Format(frmRelatObra.mskDataFim.Text, "mm/dd/yyyy")) & "#"
        Else
            rpt.RecordSelectionFormula = "{vwFULLREGATENDIMENTO.DATA}>=#" & CDate(Format(frmRelatObra.mskDataIni.Text, "mm/dd/yyyy")) & "# AND {vwFULLREGATENDIMENTO.DATA}<=#" & CDate(Format(frmRelatObra.mskDataFim.Text, "mm/dd/yyyy")) & "# AND {vwFULLREGATENDIMENTO.EQUIPE}=" & frmRelatObra.cmbEquipe.ItemData(frmRelatObra.cmbEquipe.ListIndex)
        End If
        rpt.FormulaFields(3).Text = "'" & frmRelatObra.mskDataIni.Text & "'"
        rpt.FormulaFields(4).Text = "'" & frmRelatObra.mskDataFim.Text & "'"
        If frmRelatObra.cmbSit.ListIndex = 1 Then '//concluidos
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 3 Then '//cancelados
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 2 Then '//aguardando
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) AND ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        End If
    
    
    Case "REGATENDIMENTO1", "REGATENDIMENTO4"
        frmReport.Caption = "Registro de Atendimento"
        rpt.RecordSelectionFormula = "{vwFULLREGATENDIMENTO.DATA}>=#" & CDate(Format(frmRelatObra.mskDataIni.Text, "mm/dd/yyyy")) & "# AND {vwFULLREGATENDIMENTO.DATA}<=#" & CDate(Format(frmRelatObra.mskDataFim.Text, "mm/dd/yyyy")) & "#"
        rpt.FormulaFields(3).Text = "'" & frmRelatObra.mskDataIni.Text & "'"
        rpt.FormulaFields(4).Text = "'" & frmRelatObra.mskDataFim.Text & "'"
        
        If frmRelatObra.cmbSit.ListIndex = 1 Then '//concluidos
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 3 Then '//cancelados
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 2 Then '//aguardando
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) AND ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        End If
     Case "REGATENDIMENTO5"
        frmReport.Caption = "Registro de Atendimento"
        rpt.RecordSelectionFormula = "{relatorio_obra.dtentrada}>=#" & CDate(Format(frmRelatObra.mskDataIni.Text, "mm/dd/yyyy")) & "# AND {relatorio_obra.dtentrada}<=#" & CDate(Format(frmRelatObra.mskDataFim.Text, "mm/dd/yyyy")) & "#"
        rpt.FormulaFields(1).Text = "'" & frmRelatObra.mskDataIni.Text & "'"
        rpt.FormulaFields(2).Text = "'" & frmRelatObra.mskDataFim.Text & "'"
    Case "REGATENDIMENTO2"
        frmReport.Caption = "Resumo dos Atendimento"
        rpt.RecordSelectionFormula = "{vwFULLREGATENDIMENTO.DATA}>=#" & CDate(Format(frmRelatObra.mskDataIni.Text, "mm/dd/yyyy")) & "# AND {vwFULLREGATENDIMENTO.DATA}<=#" & CDate(Format(frmRelatObra.mskDataFim.Text, "mm/dd/yyyy")) & "#"
        rpt.FormulaFields(3).Text = "'" & frmRelatObra.mskDataIni.Text & "'"
        rpt.FormulaFields(4).Text = "'" & frmRelatObra.mskDataFim.Text & "'"
        If frmRelatObra.cmbSit.ListIndex = 1 Then '//concluidos
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 3 Then '//cancelados
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND NOT ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) "
        ElseIf frmRelatObra.cmbSit.ListIndex = 2 Then '//aguardando
            rpt.RecordSelectionFormula = rpt.RecordSelectionFormula & " AND ISNULL( {vwFULLREGATENDIMENTO.DATACANCEL}) AND ISNULL( {vwFULLREGATENDIMENTO.DATAEND}) "
        End If
        
    Case "SSPAC"
        rpt.FormulaFields(1).Text = "'" & Format(Now, "dd/mm/yyyy") & "'"
        rpt.FormulaFields(2).Text = "'" & Format(Now, "hh:mm") & "'"
        rpt.FormulaFields(3).Text = "'" & frmSenhaPre.lblBanda.Caption & "'"
        rpt.FormulaFields(4).Text = "'" & frmSenhaPre.lblSenha.Caption & "'"
    Case "PAGAMENTOROCADA"
        frmReport.Caption = "Pagamento de Roçada"
        z = InputBox("Digite o ano do relatório", "Informação", Year(Now))
        If Val(z) < 2010 Or Val(z) > Year(Now) Then
            MsgBox "Ano inválido !!!", vbCritical, "Atenção"
            Exit Function
        End If
        rpt.FormulaFields(1).Text = "'" & z & "'"
        rpt.RecordSelectionFormula = "YEAR({vwPAGAMENTOROCADA.DATAVENCIMENTO})=" & Val(z)
    Case "ATIVIDADETL", "ATIVIDADETLA"
        frmReport.Caption = "ATIVIDADES"
        z = InputBox("Data inicial de abertura", "Datas", "01/01/" & Year(Now))
        If Not IsDate(z) Then
            MsgBox "Data inválida !!!", vbCritical, "Atenção"
            Exit Function
        End If
        z1 = InputBox("Data final de abertura", "Datas", Format(Now, "dd/mm/yyyy"))
        If Not IsDate(z1) Then
            MsgBox "Data inválida !!!", vbCritical, "Atenção"
            Exit Function
        End If
        rpt.RecordSelectionFormula = "{vwRELATORIOATIVIDADETL.DATAABERTURA}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {vwRELATORIOATIVIDADETL.DATAABERTURA}<=#" & CDate(Format(z1, "mm/dd/yyyy")) & "#"
    Case "CVVIMOVEL"
        frmReport.Caption = "CERTIDÃO DE VALOR VENAL"
        rpt.FormulaFields(1).Text = "'" & Format(frmCertidao.lblCertidao.Caption, "0000") & "/" & Year(Now) & "'"
        rpt.FormulaFields(2).Text = "'" & frmCertidao.txtNumProc.Text & "'"
        If IsDate(frmCertidao.lblDataProc.Caption) Then
            dData = CDate(frmCertidao.lblDataProc.Caption)
            rpt.FormulaFields(3).Text = "'" & Format(dData, "dd", vbLongDate) & " de " & Format(dData, "mmmm", vbLongDate) & " de " & Format(dData, "yyyy", vbLongDate) & "'"
            'rpt.FormulaFields(3).Text = "'" & Format(Day(dData), "00") & " de " & Format(Month(dData), "mmmm") & " de " & Year(dData) & "'"
        End If
        rpt.FormulaFields(4).Text = "'" & Mask(frmCertidao.lblEnd.Caption) & "'"
        rpt.FormulaFields(5).Text = "'" & Mask(frmCertidao.lblNum.Caption) & "'"
        rpt.FormulaFields(6).Text = "'" & IIf(frmCertidao.lblComplemento.Caption = "", "", Mask(frmCertidao.lblComplemento.Caption) & ", ") & " " & Mask(frmCertidao.lblBairro.Caption) & "'"
        rpt.FormulaFields(7).Text = "'" & Mask(frmCertidao.lblQuadra.Caption) & "'"
        rpt.FormulaFields(8).Text = "'" & Mask(frmCertidao.lblLote.Caption) & "'"
        rpt.FormulaFields(9).Text = "'" & Format(frmCertidao.txtCod.Text) & "'"
        rpt.FormulaFields(10).Text = "'" & frmCertidao.lblInscricao.Caption & "'"
        rpt.FormulaFields(11).Text = "'" & Mask(frmCertidao.lblProp.Caption) & "'"
        rpt.FormulaFields(12).Text = "'" & frmCertidao.lblVVT.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & frmCertidao.lblVVC.Caption & "'"
        rpt.FormulaFields(14).Text = "'" & frmCertidao.lblVVI.Caption & "'"
        rpt.FormulaFields(15).Text = "'" & RetornaUsuarioFullName() & "'"
        rpt.FormulaFields(16).Text = "'" & IIf(frmCertidao.chkAss.value = vbChecked, "A", "B") & "'"
    Case "CENDERECO"
        frmReport.Caption = "CERTIDÃO DE ENDEREÇO"
        rpt.FormulaFields(1).Text = "'" & Format(frmCertidao.lblCertidao.Caption, "0000") & "/" & Year(Now) & "'"
        rpt.FormulaFields(2).Text = "'" & frmCertidao.txtNumProc.Text & "'"
        If IsDate(frmCertidao.lblDataProc.Caption) Then
            dData = CDate(frmCertidao.lblDataProc.Caption)
            rpt.FormulaFields(3).Text = "'" & Format(dData, "dd", vbLongDate) & " de " & Format(dData, "mmmm", vbLongDate) & " de " & Format(dData, "yyyy", vbLongDate) & "'"
'            rpt.FormulaFields(3).Text = "'" & Format(Day(dData), "00") & " de " & Format(Month(dData), "mmmm") & " de " & Year(dData) & "'"
        End If
        rpt.FormulaFields(4).Text = "'" & frmCertidao.lblEnd.Caption & "'"
        rpt.FormulaFields(5).Text = "'" & frmCertidao.lblNum.Caption & "'"
        If frmCertidao.lblComplemento.Caption = "" Then
            rpt.FormulaFields(6).Text = "''"
        Else
            rpt.FormulaFields(6).Text = "'" & Virg2Ponto(Mask(frmCertidao.lblComplemento.Caption)) & " " & "'"
        End If
      '  rpt.FormulaFields(6).Text = IIf(frmCertidao.lblComplemento.Caption = "", "", Mask(frmCertidao.lblComplemento.Caption) & ", '")
        rpt.FormulaFields(7).Text = "'" & frmCertidao.lblQuadra.Caption & "'"
        rpt.FormulaFields(8).Text = "'" & frmCertidao.lblLote.Caption & "'"
        rpt.FormulaFields(9).Text = "'" & Format(frmCertidao.txtCod.Text) & "'"
        rpt.FormulaFields(10).Text = "'" & frmCertidao.lblInscricao.Caption & "'"
        rpt.FormulaFields(11).Text = "'" & Mask(frmCertidao.lblProp.Caption) & "'"
        rpt.FormulaFields(12).Text = "'" & frmCertidao.lblVVT.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & frmCertidao.lblVVC.Caption & "'"
        rpt.FormulaFields(14).Text = "'" & frmCertidao.lblVVI.Caption & "'"
        rpt.FormulaFields(15).Text = "'" & RetornaUsuarioFullName() & "'"
        rpt.FormulaFields(16).Text = "'" & frmCertidao.lblBairro.Caption & "'"
        rpt.FormulaFields(17).Text = "'" & IIf(frmCertidao.chkAss.value = vbChecked, "A", "B") & "'"
    Case "CISENCAO", "CISENCAOAREA"
        frmReport.Caption = "CERTIDÃO DE ISENÇÃO"
        rpt.FormulaFields(1).Text = "'" & Format(frmCertidao.lblCertidao.Caption, "0000") & "/" & Year(Now) & "'"
        rpt.FormulaFields(2).Text = "'" & frmCertidao.txtNumProc.Text & "'"
        If IsDate(frmCertidao.lblDataProc.Caption) Then
            dData = CDate(frmCertidao.lblDataProc.Caption)
            rpt.FormulaFields(3).Text = "'" & Format(dData, "dd", vbLongDate) & " de " & Format(dData, "mmmm", vbLongDate) & " de " & Format(dData, "yyyy", vbLongDate) & "'"
        End If
        rpt.FormulaFields(4).Text = "'" & frmCertidao.lblEnd.Caption & "'"
        rpt.FormulaFields(5).Text = "'" & frmCertidao.lblNum.Caption & "'"
        rpt.FormulaFields(6).Text = "'" & IIf(frmCertidao.lblComplemento.Caption = "", "", frmCertidao.lblComplemento.Caption & ", ") & "'"
        rpt.FormulaFields(7).Text = "'" & frmCertidao.lblQuadra.Caption & "'"
        rpt.FormulaFields(8).Text = "'" & frmCertidao.lblLote.Caption & "'"
        rpt.FormulaFields(9).Text = "'" & Format(frmCertidao.txtCod.Text) & "'"
        rpt.FormulaFields(10).Text = "'" & frmCertidao.lblInscricao.Caption & "'"
        rpt.FormulaFields(11).Text = "'" & frmCertidao.lblProp.Caption & "'"
        rpt.FormulaFields(12).Text = "'" & frmCertidao.lblVVT.Caption & "'"
        rpt.FormulaFields(13).Text = "'" & frmCertidao.lblVVC.Caption & "'"
        rpt.FormulaFields(14).Text = "'" & frmCertidao.lblVVI.Caption & "'"
        rpt.FormulaFields(15).Text = "'" & RetornaUsuarioFullName() & "'"
       ' rpt.FormulaFields(16).Text = "'" & RetornaUsuarioFullName() & "'"
        rpt.FormulaFields(17).Text = "'" & frmCertidao.lblProcIsencao.Caption & "'"
        If IsDate(frmCertidao.lblDataProcIsencao.Caption) Then
            dData = CDate(frmCertidao.lblDataProcIsencao.Caption)
            rpt.FormulaFields(18).Text = "'" & Format(dData, "dd", vbLongDate) & " de " & Format(dData, "mmmm", vbLongDate) & " de " & Format(dData, "yyyy", vbLongDate) & "'"
        End If
        rpt.FormulaFields(19).Text = "'" & frmCertidao.lblPercIsencao.Caption & "'"
End Select

If sReport <> "DAM" And sReport <> "DAMHONORARIO" And sReport <> "DAMTMP" Then
    rpt.PaperSize = crPaperA4
End If

If Left(sReport, 3) = "BOL" Then
    Sql = "select * from machines2 where computer='" & NomeDoComputador & "'"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    With RdoAux
        If !margin_top = 0 And !Margin_left = 0 And !margin_bottom = 0 And !margin_right = 0 Then
        Else
            rpt.TopMargin = !margin_top
            rpt.LeftMargin = !Margin_left
            rpt.BottomMargin = !margin_bottom
            rpt.RightMargin = !margin_right
        End If
    End With
End If

If frmMdi.m_cMenuPrincipal.Checked(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True Then
    rpt.PaperSource = crPRBinLower
End If

rpt.DisplayProgressDialog = True

If UCase$(sReport) = "BOLETOGUIA" Or UCase$(sReport) = "BOLETOGUIATMP" Then
    rpt.Sections(1).Suppress = bHeader
End If



Select Case UCase$(sReport)
    Case "REGATENDIMENTO5"
        rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTributacao", "Tributacao", UL, UP
    Case "CARNETMP", "CONFDIVIDATMP", "SIMULADOTMP", "COBRANCAAMIGAVELTMP", "CALCULOPARCELAMENTOTMP", "DAMTMP", "BOLETOGUIATMP", "BOLETODAMTESTE", "BOLETODAM_V4TMP", "BOLETODAM_V4TMP2"
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "TributacaoTeste", UL, UP
    Case Else
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "Tributacao", UL, UP
End Select



'If UCase$(sReport) = "CARNETMP" Or UCase$(sReport) = "CONFDIVIDATMP" Or UCase$(sReport) = "SIMULADOTMP" Or UCase$(sReport) = "COBRANCAAMIGAVELTMP" Or UCase$(sReport) = "CALCULOPARCELAMENTOTMP" Or UCase$(sReport) = "DAMTMP" Or UCase$(sReport) = "BOLETOGUIATMP" Then
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTribTeste", "TributacaoTeste", LastUser, UserPwd
'ElseIf UCase$(sReport) = "CARNELOCAL" Or UCase$(sReport) = "EXTRATOFULL" Then
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTribLocal", "Tributacao_Full", LastUser, UserPwd
'Else
    'rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTributacao", "Tributacao", LastUser, UserPwd
'    rpt.Database.LogOnServer "PDSODBC.DLL", "odbcTributacao", "Tributacao", UL, UP
'    rpt.Database.Tables(1).SetLogOnInfo "192.168.15.160", "Tributacao", UL, UP
'End If
    
rpt.DiscardSavedData

CRViewer1.ReportSource = rpt

show:
CRViewer1.ViewReport
Liberado

'If nNumDoc > 0 Then
'    rpt.ExportOptions.DestinationType = crEDTDiskFile
'    If bLocal Then
'        rpt.ExportOptions.DiskFileName = "C:\TMP\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
'    Else
'        rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\SEGUNDAVIA\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
'    End If
'    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
'    rpt.ExportOptions.PDFExportAllPages = True
'    rpt.Export (False)
'End If

If UCase(sReport) = "NOTIFICACAO3" Or UCase(sReport) = "NOTIFICACAO4" Then
    Sql = "select count(seq) as maximo from documentopic where codigo=" & Val(z)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 1
    Else
        nSeq = RdoAux!maximo + 1
    End If
    RdoAux.Close
    
    Sql = "select max(seq) as maximo from documentopic"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq2 = 1
    Else
        nSeq2 = RdoAux!maximo + 1
    End If
    RdoAux.Close
    sTexto1 = "05" & frmNotificacao2.cmbAno.Text & Format(nSeq, "00") & Format(frmNotificacao2.txtCodImovel.Text, "000000") & ".pdf"
    
    Sql = "insert documentopic(seq,codigo,documento) values(" & nSeq2 & "," & Val(z) & ",'" & sTexto1 & "')"
    cn.Execute Sql, rdExecDirect
    
    rpt.ExportOptions.DestinationType = crEDTDiskFile
    rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\Documentos\" & Year(Now) & "\" & sTexto1
    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rpt.ExportOptions.PDFExportAllPages = True
    rpt.Export (False)
End If

If UCase(sReport) = "ALVARA" Or UCase(sReport) = "ALVARASEMDATA" Or UCase(sReport) = "ALVARAVICE" Or UCase(sReport) = "ALVARASEMDATAVICE" Then
    Sql = "select count(seq) as maximo from documentopic where codigo=" & Val(frmAlvara.txtCodigo.Text)
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq = 1
    Else
        nSeq = RdoAux!maximo + 1
    End If
    RdoAux.Close
    
    Sql = "select max(seq) as maximo from documentopic"
    Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
    If IsNull(RdoAux!maximo) Then
        nSeq2 = 1
    Else
        nSeq2 = RdoAux!maximo + 1
    End If
    RdoAux.Close
    sTexto1 = "08" & Year(Now) & Format(nSeq, "00") & Format(frmAlvara.txtCodigo.Text, "000000") & ".pdf"
    
    Sql = "insert documentopic(seq,codigo,documento) values(" & nSeq2 & "," & Val(frmAlvara.txtCodigo.Text) & ",'" & sTexto1 & "')"
    cn.Execute Sql, rdExecDirect
    
    rpt.ExportOptions.DestinationType = crEDTDiskFile
    rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\Documentos\" & Year(Now) & "\" & sTexto1
    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rpt.ExportOptions.PDFExportAllPages = True
    rpt.Export (False)
End If


frmReport.show 1

If UCase(sReport) = "REFISPARC" Then
    Sql = "delete from extratotmp where computer='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
ElseIf UCase(sReport) = "REFIS" Then
    Sql = "DELETE FROM relatorio_refis WHERE usuario='" & NomeDeLogin & "'"
    cn.Execute Sql, rdExecDirect
End If

Exit Function
Erro:

Liberado
MsgBox Err.Description
Resume Next

End Function

Public Function ShowReport3(sReport As String, hMDI As Long, hFormCalling As Long, Optional nNumDoc As Long, Optional nNumGuia As Long)
Dim RdoAux As rdoResultset, Sql As String, sTipo As String, dData As Date, nAno As Integer, sDoc As String, nSeq2 As Integer
Dim sTexto1 As String, sTexto2 As String, sTexto3 As String, sHor As String, sSenha As String, nSeq As Integer, nCodReduz As Long
Dim z As Variant, RdoAux2 As rdoResultset, z2 As Variant, z3 As Variant, z4 As Variant, z5 As Variant
Dim sNumProc As String, nNumproc As Long, bAchou As Boolean, aTributo() As String, x As Integer, Y As Integer
Dim qd As New rdoQuery, bHeader As Boolean

If UCase(sReport) = "BOLETOGUIA2" Then
    bHeader = True
    sReport = "BOLETOGUIA"
End If
If UCase(sReport) = "BOLETOGUIA2TMP" Then
    bHeader = True
    sReport = "BOLETOGUIATMP"
End If

Set rpt = crApp.OpenReport(sPathReport & "\" & sReport & ".Rpt", 1)
If sReport <> "DAM" And sReport <> "DAMHONORARIO" And sReport <> "DAMTMP" Then
    rpt.PaperSize = crPaperA4
End If

If frmMdi.m_cMenuPrincipal.Checked(frmMdi.m_cMenuPrincipal.IndexForKey("mnuPrintBottom")) = True Then
    rpt.PaperSource = crPRBinLower
End If

rpt.DisplayProgressDialog = True

If UCase$(sReport) = "BOLETOGUIA" Or UCase$(sReport) = "BOLETOGUIATMP" Or UCase$(sReport) = "BOLETOGUIA_V3" Then
    rpt.Sections(1).Suppress = bHeader
End If


Select Case UCase(sReport)
    Case "EMPRESA_QTDEATIVIDADE"
        rpt.FormulaFields(1).Text = "'DATA DE ABERTURA ENTRE:" & frmCnsAvancadaMob.mskDataAbeIni.Text & " E " & frmCnsAvancadaMob.mskDataAbeFim.Text & "'"
    Case "NF_EMITIDA"
Data1:
            z = InputBox("Digite a data inicial.", "Entre com a informação")
            If z = "" Then GoTo Data1
            If Not IsDate(z) Then GoTo Data1
Data2:
            z2 = InputBox("Digite a data final.", "Entre com a informação")
            If z2 = "" Then GoTo Data2
            If Not IsDate(z2) Then GoTo Data2
        rpt.RecordSelectionFormula = "{Comando.dataemissao}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {Comando.dataemissao}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
        frmReport.Caption = "Notas fiscais emitidas por período"
        rpt.FormulaFields(1).Text = "'PERÍODO DE " & z & " À " & z2 & "'"
    Case "CALCULO_PARCELAMENTO2", "CALCULO_PARCELAMENTO2_TMP"
            frmReport.Caption = "Calculo de Parcelamento"
    Case "REGATENDIMENTO_ENDERECO"
            frmReport.Caption = "Registro Atendimento por endereço"
    Case "DECA"
        rpt.RecordSelectionFormula = "{REPORTTMP.USUARIO}='" & NomeDeLogin & "'"
        frmReport.Caption = "Impressão de DECA frente"
        rpt.FormulaFields(1).Text = frmCadMob.txtCodEmpresa.Text
        rpt.FormulaFields(2).Text = "'" & Mask(frmCadMob.txtRazao.Text) & "'"
        rpt.FormulaFields(3).Text = "'" & Mask(frmCadMob.txtAtivExt.Text) & "'"
        rpt.FormulaFields(4).Text = "'" & "" & "'" 'ativext2
        rpt.FormulaFields(5).Text = "'" & "" & "'" 'codatividade
        rpt.FormulaFields(6).Text = "'" & Mask(frmCadMob.txtNomeLogr.Text) & ", " & frmCadMob.txtNumero.Text & " " & Mask(frmCadMob.txtCompl.Text) & "'"
        rpt.FormulaFields(7).Text = "'" & "" & "'" 'andar
        rpt.FormulaFields(8).Text = "'" & "" & "'" 'sala
        rpt.FormulaFields(9).Text = "'" & Mask(frmCadMob.cmbBairro.Text) & "'"
        rpt.FormulaFields(10).Text = "'" & Mask(frmCadMob.mskCEP.Text) & "'"
        rpt.FormulaFields(11).Text = "'" & Mask(frmCadMob.cmbCidade.Text) & "'"
        rpt.FormulaFields(12).Text = "'" & "" & "'" 'zona
        rpt.FormulaFields(13).Text = "'" & Mask(frmCadMob.txtFone.Text) & "'"
        rpt.FormulaFields(14).Text = "'" & "" & "'"
        rpt.FormulaFields(15).Text = "'" & frmCadMob.txtArea.Text & "'"
        rpt.FormulaFields(16).Text = "'" & frmCadMob.txtNumFunc.Text & "'" 'numemp
        rpt.FormulaFields(17).Text = "'" & "" & "'" 'txtmunicipio
        rpt.FormulaFields(18).Text = "'" & Mask(frmCadMob.txtTipoConselho.Text) & "'" 'txtorgao
        rpt.FormulaFields(21).Text = "'" & Mask(frmCadMob.txtInscEst.Text) & "'" 'txtnumreg
        rpt.FormulaFields(20).Text = "'" & frmCadMob.txtCapital.Text & "'" 'txtcapital
        rpt.FormulaFields(19).Text = "'" & Mask(frmCadMob.txtNumRegistro.Text) & "'" 'txtrg
        rpt.FormulaFields(22).Text = "'" & Mask(frmCadMob.mskCPF.Text) & "'" 'txtcpf
        rpt.FormulaFields(24).Text = "'" & " " & "'"
        rpt.FormulaFields(23).Text = "'" & " " & "'"
        rpt.FormulaFields(25).Text = "'" & " " & "'"
        rpt.FormulaFields(26).Text = "'" & " " & "'"
        rpt.FormulaFields(27).Text = "'" & " " & "'"
        rpt.FormulaFields(28).Text = "'" & " " & "'"
        rpt.FormulaFields(29).Text = "'" & " " & "'"
        rpt.FormulaFields(30).Text = "'" & " " & "'"
        rpt.FormulaFields(31).Text = "'" & " " & "'"
        rpt.FormulaFields(32).Text = "'" & " " & "'"
        rpt.FormulaFields(33).Text = "'" & frmCadMob.mskDataAb.Text & "'"
        rpt.FormulaFields(35).Text = "'" & "" & "'"
        rpt.FormulaFields(36).Text = "'" & "" & "'"
        rpt.FormulaFields(37).Text = "'" & "" & "'"
        rpt.FormulaFields(38).Text = "'" & "" & "'"
        rpt.FormulaFields(39).Text = "'" & "" & "'"
        rpt.FormulaFields(40).Text = "'" & "" & "'"
        rpt.FormulaFields(41).Text = "'" & "" & "'"
        rpt.FormulaFields(42).Text = "'" & "" & "'"
        If frmCadMob.mskCNPJ.ClipText <> "" Then
            rpt.FormulaFields(44).Text = "'" & "X" & "'"
            rpt.FormulaFields(43).Text = "'" & " " & "'"
        Else
            rpt.FormulaFields(44).Text = "'" & " " & "'"
            rpt.FormulaFields(43).Text = "'" & "X" & "'"
        End If
        If Left(frmCadMob.txtAtiv.Text, 1) = "2" Then
            rpt.FormulaFields(45).Text = "'" & "X" & "'"
            rpt.FormulaFields(46).Text = "'" & " " & "'"
            rpt.FormulaFields(47).Text = "'" & " " & "'"
        ElseIf Left(frmCadMob.txtAtiv.Text, 1) = "1" Then
            rpt.FormulaFields(45).Text = "'" & " " & "'"
            rpt.FormulaFields(46).Text = "'" & "X" & "'"
            rpt.FormulaFields(47).Text = "'" & " " & "'"
        ElseIf Left(frmCadMob.txtAtiv.Text, 1) = "3" Then
            rpt.FormulaFields(45).Text = "'" & " " & "'"
            rpt.FormulaFields(46).Text = "'" & " " & "'"
            rpt.FormulaFields(47).Text = "'" & "X " & "'"
        Else
            rpt.FormulaFields(45).Text = "'" & " " & "'"
            rpt.FormulaFields(46).Text = "'" & " " & "'"
            rpt.FormulaFields(47).Text = "'" & " " & "'"
        End If
        
        rpt.FormulaFields(48).Text = "'" & " " & "'"
        rpt.FormulaFields(49).Text = "'" & " " & "'"
        rpt.FormulaFields(50).Text = "'" & "" & "'" 'txtHist
        rpt.FormulaFields(51).Text = "'" & "" & "'" 'txtassinatura
        rpt.FormulaFields(52).Text = "'" & "" & "'" 'end entrega
'        If frmDeca.chkAmbulante.Value = vbChecked Then
'            rpt.FormulaFields(53).Text = "'X'"
'            rpt.FormulaFields(54).Text = "'" & frmDeca.cmbAmbulante.Text & "'"
'            rpt.FormulaFields(55).Text = "'Trabalho como comércio ambulante de: " & Mask(frmDeca.txtDescAmbulante.Text) & "'"
'        Else
            rpt.FormulaFields(53).Text = "''" 'ambulante
            rpt.FormulaFields(54).Text = "''" 'tipo
            rpt.FormulaFields(55).Text = "''" 'especificacao ativ.ambulante
       ' End If
        rpt.FormulaFields(56).Text = "'" & "" & "'" 'txtdescambulante
        rpt.FormulaFields(57).Text = "'" & Mask(frmCadMob.txtEmail.Text) & "'"
    Case "PARCELAMENTO_SIMULADO", "PARCELAMENTO_SIMULADO_TMP"
        frmReport.Caption = "Simulação de Parcelamento"
        rpt.RecordSelectionFormula = "{PARCELAMENTO_SIMULADO.USUARIO}='" & NomeDeLogin & "'"
    Case "ISSCCIVIL"
            frmReport.Caption = "Resumo Iss costrução civil"
DataCC1:
            z = InputBox("Digite a data inicial.", "Entre com o período de lançamento")
            If z = "" Then GoTo DataCC1
            If Not IsDate(z) Then GoTo DataCC1
DataCC2:
            z2 = InputBox("Digite a data final.", "Entre com o período de lançamento")
            If z2 = "" Then GoTo DataCC2
            If Not IsDate(z2) Then GoTo DataCC2
            On Error Resume Next
            rpt.RecordSelectionFormula = "{vwISSCCivil.DATADEBASE}>=#" & CDate(Format(z, "mm/dd/yyyy")) & "# AND {vwISSCCivil.DATADEBASE}<=#" & CDate(Format(z2, "mm/dd/yyyy")) & "#"
            On Error GoTo 0
            rpt.FormulaFields(1).Text = "'" & Format(z, "dd/mm/yyyy") & "'"
            rpt.FormulaFields(2).Text = "'" & Format(z2, "dd/mm/yyyy") & "'"
    Case "DECA2"
        frmReport.Caption = "Impressão de DECA verso"
        rpt.FormulaFields(1).Text = "'" & "" & "'"
        rpt.FormulaFields(2).Text = "'" & "" & "'"
        rpt.FormulaFields(3).Text = "'" & "" & "'"
        rpt.FormulaFields(4).Text = "'" & "" & "'"
        rpt.FormulaFields(5).Text = "'" & "" & "'"
        rpt.FormulaFields(6).Text = "'" & "" & "'"
        rpt.FormulaFields(7).Text = "'" & "" & "'"
        rpt.FormulaFields(8).Text = "'" & "" & "'"
        rpt.FormulaFields(9).Text = "'" & "" & "'"
        rpt.FormulaFields(10).Text = "'" & "" & "'"
        rpt.FormulaFields(11).Text = "'" & "" & "'"
        rpt.FormulaFields(12).Text = "'" & "" & "'"
        rpt.FormulaFields(13).Text = "'" & "" & "'"
        rpt.FormulaFields(14).Text = "'" & "" & "'"
        rpt.FormulaFields(15).Text = "'" & "" & "'"
        rpt.FormulaFields(16).Text = "'" & "" & "'"
        rpt.FormulaFields(17).Text = "'" & "" & "'"
        rpt.FormulaFields(18).Text = "'" & "" & "'"
        rpt.FormulaFields(19).Text = "'" & "" & "'"
        rpt.FormulaFields(20).Text = "'" & "" & "'"
        rpt.FormulaFields(21).Text = "'" & "" & "'"
        rpt.FormulaFields(22).Text = "'" & "" & "'"
        rpt.FormulaFields(23).Text = "'" & "" & "'"
        rpt.FormulaFields(24).Text = "'" & "" & "'"
        rpt.FormulaFields(25).Text = "'" & "" & "'"
        rpt.FormulaFields(26).Text = "'" & "" & "'"
        rpt.FormulaFields(27).Text = "'" & "" & "'"
        rpt.FormulaFields(28).Text = "'" & "" & "'"
        rpt.FormulaFields(29).Text = "'" & "" & "'"
        rpt.FormulaFields(30).Text = "'" & "" & "'"
        rpt.FormulaFields(31).Text = "'" & "" & "'"
        rpt.FormulaFields(32).Text = "'" & "" & "'"
        rpt.FormulaFields(33).Text = "'" & "" & "'"
        rpt.FormulaFields(34).Text = "'" & "" & "'"
        rpt.FormulaFields(35).Text = "'" & "" & "'"
        rpt.FormulaFields(36).Text = "'" & "" & "'"
        rpt.FormulaFields(37).Text = "'" & "" & "'"
        rpt.FormulaFields(38).Text = "'" & "" & "'"
        rpt.FormulaFields(39).Text = "'" & "" & "'"
        rpt.FormulaFields(40).Text = "'" & "" & "'"
        Sql = "select * from escritoriocontabil where codigoesc=" & Val(frmCadMob.txtCodEsc.Text)
        Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If RdoAux.RowCount > 0 Then
            rpt.FormulaFields(41).Text = "'" & RdoAux!NOMEESC & "'"
            rpt.FormulaFields(42).Text = "'" & SubNull(RdoAux!NomeLogradouro) & "'"
            rpt.FormulaFields(43).Text = "'" & SubNull(RdoAux!NOMEBairro) & "'"
            rpt.FormulaFields(44).Text = "'" & SubNull(RdoAux!telefone) & "'"
            rpt.FormulaFields(45).Text = "'" & SubNull(RdoAux!Numero) & "'"
            rpt.FormulaFields(46).Text = "'" & SubNull(RdoAux!Cep) & "'"
            rpt.FormulaFields(58).Text = "'" & SubNull(RdoAux!email) & "'"
        Else
            rpt.FormulaFields(41).Text = "'" & "" & "'"
            rpt.FormulaFields(42).Text = "'" & "" & "'"
            rpt.FormulaFields(43).Text = "'" & "" & "'"
            rpt.FormulaFields(44).Text = "'" & "" & "'"
            rpt.FormulaFields(45).Text = "'" & "" & "'"
            rpt.FormulaFields(46).Text = "'" & "" & "'"
            rpt.FormulaFields(58).Text = "'" & "" & "'"
        End If
        RdoAux.Close
        
        rpt.FormulaFields(47).Text = "'" & "" & "'" 'txtrgc
        rpt.FormulaFields(48).Text = "'" & "" & "'" 'txtorgaoc
        rpt.FormulaFields(49).Text = "'" & "" & "'" 'mskobsc
        rpt.FormulaFields(59).Text = "'" & "" & "'" 'fone0
        rpt.FormulaFields(60).Text = "'" & "" & "'"
        rpt.FormulaFields(61).Text = "'" & "" & "'"
        rpt.FormulaFields(62).Text = "'" & "" & "'"
        rpt.FormulaFields(63).Text = "'" & "" & "'"
        rpt.FormulaFields(64).Text = "'" & "" & "'"
        rpt.FormulaFields(65).Text = "'" & "" & "'"
        rpt.FormulaFields(66).Text = "'" & "" & "'" 'fone7
        rpt.FormulaFields(67).Text = "'" & "" & "'" 'cidade C
        rpt.FormulaFields(68).Text = "'" & "" & "'" 'uf c
End Select

Select Case UCase$(sReport)
    Case "CARNETMP", "PARCELAMENTO_SIMULADO_TMP", "CALCULO_PARCELAMENTO2_TMP"
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "TributacaoTeste", UL, UP
    Case "DECA", "DECA2"
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "Tributacao", UL, UP
    Case Else
        rpt.Database.Tables(1).SetLogOnInfo IPServer, "Tributacao", UL, UP
End Select

rpt.DiscardSavedData

CRViewer1.ReportSource = rpt

show:
CRViewer1.ViewReport
Liberado

If nNumDoc > 0 And NomeDoComputador <> "MATHWORLD" And NomeDoComputador <> "GTI" Then
    rpt.ExportOptions.DestinationType = crEDTDiskFile
    If bLocal Then
        rpt.ExportOptions.DiskFileName = "C:\TMP\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
    Else
        rpt.ExportOptions.DiskFileName = "\\192.168.200.130\ATUALIZAGTI\SEGUNDAVIA\" & Format(nNumGuia, "000000000") & "[" & NomeDeLogin & "].PDF"
    End If
    rpt.ExportOptions.FormatType = crEFTPortableDocFormat
    rpt.ExportOptions.PDFExportAllPages = True
    rpt.Export (False)
End If



frmReport.show 1

Exit Function
Erro:

Liberado
MsgBox Err.Description


End Function


Private Sub MontaMalaDiretaParc()
Dim RdoAux As rdoResultset, Sql As String, RdoAux2 As rdoResultset, nSeq As Integer
Dim sNome As String, sEnd As String, sCompl As String, sBairro As String, sCid As String, sCep As String, sUF As String
Dim sNumProc As String, nAnoproc As Integer, nNumproc As Long
On Error GoTo Erro
nSeq = 1
Sql = "DELETE FROM ETIQUETAGTI WHERE USUARIO='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "SELECT DISTINCT processogti.CODCIDADAO, cidadao.nomecidadao FROM processogti INNER JOIN Cidadao ON processogti.CODCIDADAO = cidadao.codcidadao "
Sql = Sql & "WHERE (processogti.CODASSUNTO = 759 OR processogti.CODASSUNTO = 828 OR processogti.CODASSUNTO = 817) AND (processogti.ANO = 2009) AND (processogti.CODCIDADAO > 0) ORDER BY NOMECIDADAO"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
'        If IsNull(!NUMPROC) Then GoTo PROXIMO
'        sNumProc = !numprocesso
'        nNumProc = !NUMPROC
'        nAno = !anoproc
        Sql = "SELECT * FROM vwFULLCIDADAO WHERE CODCIDADAO=" & !CodCidadao
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        With RdoAux2
            sNome = !nomecidadao
'            If Val(SubNull(!CODLOGRADOURO)) > 0 Then
'                sEnd = Trim$(SubNull(!AbrevTipoLog)) & " " & Trim$(SubNull(!AbrevTitLog)) & " " & SubNull(!NomeLogradouro) & ", " & SubNull(!NUMIMOVEL)
'                sCEP = RetornaCEP(!CODLOGRADOURO, !NUMIMOVEL)
'            Else
'                sEnd = SubNull(!NomeLogradouro2) & ", " & SubNull(!NUMIMOVEL)
'                sCEP = SubNull(!cep2)
'            End If
            sEnd = SubNull(!Endereco) & "," & SubNull(!NUMIMOVEL)
            sCompl = SubNull(!Complemento)
            'sBairro = SubNull(!NOMEBairro)
            'If sBairro = "" Then
                 sBairro = SubNull(!DescBairro)
            'End If
            'sCid = SubNull(!NomeCidade)
            'If sCidade = "" Then
                sCid = SubNull(!descCidade)
            'End If
            sUF = SubNull(!SiglaUF)
            If !CodLogradouro > 0 Then
                sCep = RetornaCEP(!CodLogradouro, !NUMIMOVEL)
            Else
                sCep = SubNull(!Cep)
            End If
            
           .Close
        End With
        
        Sql = "INSERT ETIQUETAGTI (USUARIO,SEQ,CAMPO1,CAMPO2,CAMPO3,CAMPO4,CAMPO5) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & nSeq & ",'" & Format(!CodCidadao, "000000") & "','" & Mask(sNome) & "','" & sEnd & "','" & sBairro & " - " & sCid & "','" & sUF & " - " & sCep & "')"
        cn.Execute Sql, rdExecDirect
proximo:
        nSeq = nSeq + 1
       .MoveNext
    Loop
   .Close
End With

Exit Sub
Erro:
MsgBox Err.Description
Resume Next
End Sub

Private Sub GeraRefis(nAno As Integer)
Dim RdoAux As rdoResultset, sPlano As String
On Error GoTo Erro
Ocupado
Sql = "DELETE FROM EXTRATOTMP WHERE COMPUTER='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect
'nAno = 2015
Sql = "SELECT DISTINCT debitoparcela.codreduzido, debitoparcela.numprocesso, SUM(debitotributo.valortributo) AS Soma, debitoparcela.codlancamento, parceladocumento.plano,plano.Nome "
Sql = Sql & "FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND debitoparcela.numparcela = debitotributo.numparcela AND "
Sql = Sql & "debitoparcela.codcomplemento = debitotributo.codcomplemento INNER JOIN parceladocumento ON debitoparcela.codreduzido = parceladocumento.codreduzido AND debitoparcela.anoexercicio = parceladocumento.anoexercicio AND "
Sql = Sql & "debitoparcela.codlancamento = parceladocumento.codlancamento AND debitoparcela.seqlancamento = parceladocumento.seqlancamento AND debitoparcela.numparcela = parceladocumento.numparcela AND "
Sql = Sql & "debitoparcela.codcomplemento = parceladocumento.codcomplemento INNER JOIN plano ON parceladocumento.plano = plano.codigo GROUP BY debitoparcela.codreduzido, debitoparcela.numprocesso, debitoparcela.codlancamento, "
Sql = Sql & "parceladocumento.plano, plano.nome HAVING parceladocumento.plano IN (20, 21, 22,24,25) ORDER BY debitoparcela.codreduzido"

Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
    
        nCodReduz = !CODREDUZIDO
        sNumProc = !numprocesso
        sPlano = !Nome
        If nCodReduz < 100000 Then
            Sql = "SELECT NOMECIDADAO AS NOME FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & nCodReduz
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            Sql = "SELECT RAZAOSOCIAL AS NOME FROM vwFULLEMPRESA WHERE CODIGOMOB=" & nCodReduz
        Else
            Sql = "SELECT NOMECIDADAO AS NOME FROM vwFULLCIDADAO WHERE CODCIDADAO=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = RdoAux2!Nome
        RdoAux2.Close
        Sql = "SELECT * FROM PROCESSOREPARC WHERE NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        'If RdoAux2!plano = 0 Then GoTo proximo
        nQtdeParc = RdoAux2!qtdeparcela
        bCancel = RdoAux2!Cancelado
        RdoAux2.Close
        
        Sql = "SELECT COUNT(*) AS CONTADOR FROM DEBITOPARCELA WHERE CODREDUZIDO=" & nCodReduz & " AND STATUSLANC=2 AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nQtdePago = RdoAux2!contador
        RdoAux2.Close
        
        Sql = "SELECT SUM(debitotributo.valortributo) AS soma FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
        Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO And debitoparcela.NumParcela = debitotributo.NumParcela "
        Sql = Sql & "WHERE debitoparcela.CODREDUZIDO=" & nCodReduz & " AND NUMPROCESSO='" & sNumProc & "'"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        nValorTotal = RdoAux2!soma
        RdoAux2.Close
                
        Sql = "SELECT SUM(debitotributo.valortributo) AS soma FROM debitoparcela INNER JOIN debitotributo ON debitoparcela.codreduzido = debitotributo.codreduzido AND debitoparcela.anoexercicio = debitotributo.anoexercicio AND "
        Sql = Sql & "debitoparcela.codlancamento = debitotributo.codlancamento AND debitoparcela.seqlancamento = debitotributo.seqlancamento AND "
        Sql = Sql & "debitoparcela.CODCOMPLEMENTO = debitotributo.CODCOMPLEMENTO And debitoparcela.NumParcela = debitotributo.NumParcela "
        Sql = Sql & "WHERE debitoparcela.CODREDUZIDO=" & nCodReduz & " AND STATUSLANC=2 AND NUMPROCESSO='" & sNumProc & "' AND STATUSLANC=2"
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        If IsNull(RdoAux2!soma) Then
            nValorPago = 0
            GoTo proximo
        Else
            nValorPago = RdoAux2!soma
        End If
        RdoAux2.Close
                
        If IsNull(nQtdeParc) Then nQtdeParc = 0
        Sql = "INSERT EXTRATOTMP(COMPUTER,SEQ,CODREDUZIDO,DESCLANCAMENTO,NOMEPROP,ANOEXERCICIO,NUMSEQUENCIA,NUMPARCELA,VALORLANCADO,VALORCORRECAO) VALUES('"
        Sql = Sql & NomeDeLogin & "'," & .AbsolutePosition & "," & nCodReduz & ",'" & sPlano & "','" & Left(sNome, 30) & "'," & nQtdeParc & "," & nQtdePago & "," & IIf(bCancel, 1, 0) & ","
        Sql = Sql & Virg2Ponto(CStr(nValorTotal)) & "," & Virg2Ponto(CStr(nValorPago)) & ")"
        cn.Execute Sql, rdExecDirect
proximo:
       .MoveNext
    Loop
   .Close
End With
Liberado
Exit Sub
Erro:
'MsgBox Err.Description
Resume Next
End Sub

Private Sub GeraRefisDam(sDataIni As String, sDataFim As String)
Dim RdoAux As rdoResultset, sNome As String, x As Integer
On Error GoTo Erro
Ocupado
x = 1
Sql = "DELETE FROM relatorio_refis WHERE usuario='" & NomeDeLogin & "'"
cn.Execute Sql, rdExecDirect

Sql = "select * from vwrefisnovo2 where  datapagamento between '" & Format(sDataIni, "mm/dd/yyyy") & "' and '" & Format(sDataFim, "mm/dd/yyyy") & "'"
Set RdoAux = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
With RdoAux
    Do Until .EOF
   
        nCodReduz = !CODREDUZIDO
        If nCodReduz < 100000 Then
            Sql = "SELECT NOMECIDADAO AS NOME FROM vwFULLIMOVEL WHERE CODREDUZIDO=" & nCodReduz
        ElseIf nCodReduz >= 100000 And nCodReduz < 500000 Then
            Sql = "SELECT RAZAOSOCIAL AS NOME FROM MOBILIARIO WHERE CODIGOMOB=" & nCodReduz
        Else
            Sql = "SELECT NOMECIDADAO AS NOME FROM CIDADAO WHERE CODCIDADAO=" & nCodReduz
        End If
        Set RdoAux2 = cn.OpenResultset(Sql, rdOpenKeyset, rdConcurValues)
        sNome = RdoAux2!Nome
        RdoAux2.Close
        Sql = "insert relatorio_refis(seq,usuario,numdocumento,codreduzido,datapagamento,valorpago,valordoc,plano,nome,nomecontribuinte) values("
        Sql = Sql & x & ",'" & NomeDeLogin & "'," & !NumDocumento & "," & !CODREDUZIDO & ",'" & Format(!DataPagamento, "mm/dd/yyyy") & "',"
        Sql = Sql & Virg2Ponto(CStr(!ValorPago)) & "," & Virg2Ponto(CStr(!valordoc)) & "," & !plano & ",'" & !Nome & "','" & Mask(sNome) & "')"
        cn.Execute Sql, rdExecDirect
        x = x + 1
       .MoveNext
    Loop
   .Close
End With
Liberado

Exit Sub

Erro:
Resume Next

End Sub

