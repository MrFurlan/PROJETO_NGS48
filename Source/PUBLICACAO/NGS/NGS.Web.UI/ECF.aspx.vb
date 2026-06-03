Imports System.IO
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ECF
    Inherits BasePage

    Dim objEmpresa As ClienteXEmpresa
    Dim objParametros As ClienteXEmpresaXECF

#Region "Session"
    Private Sub SessaoSalvaEmpresa()
        Session("objParametros" + HID.Value) = objEmpresa
    End Sub

    Private Sub SessaoRecuperaEmpresa()
        objEmpresa = CType(Session("objParametros" + HID.Value), [Lib].Negocio.ClienteXEmpresa)
        objParametros = objEmpresa.ParametrosECF
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ECF", "ACESSAR") Then
                txtDT_INI.Text = String.Format("01/01/{0}", DateTime.Now.Year - 1)
                txtDT_FIN.Text = String.Format("31/12/{0}", DateTime.Now.Year - 1)
                ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, " CE.Matriz = 'S'")
                Funcoes.VerificaEmpresa(DdlEmpresa)

                txtNUM_REC.Parent.Visible = False
                DivIsentaImune.Visible = False
                HID.Value = Guid.NewGuid().ToString
                CarregarIND_ALIQ_CSLL()
                IND_ALIQ_CSLL_Mostrar_Ocultar()
                LiberaEmpresa()
            Else

                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub CarregarIND_ALIQ_CSLL()
        ddlIND_ALIQ_CSLL_I.Items.Clear()
        ddlIND_ALIQ_CSLL_I.Items.Insert(0, "")
        ddlIND_ALIQ_CSLL_I.Items.Insert(1, "9%")
        ddlIND_ALIQ_CSLL_I.Items.Insert(2, "17%")
        ddlIND_ALIQ_CSLL_I.Items.Insert(3, "20%")
        ddlIND_ALIQ_CSLL_I.SelectedIndex = 0
    End Sub

    Protected Sub ddlFORMA_TRIB_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFORMA_TRIB.SelectedIndexChanged
        DivIsentaImune.Visible = (ddlFORMA_TRIB.SelectedValue = 5 Or ddlFORMA_TRIB.SelectedValue = 8 Or ddlFORMA_TRIB.SelectedValue = 9)
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresa.SelectedIndexChanged
        If DdlEmpresa.SelectedIndex > 0 Then
            Dim Emp() As String = DdlEmpresa.SelectedValue.Split("-")
            objEmpresa = New ClienteXEmpresa(Emp(0), Emp(1))
            objParametros = objEmpresa.ParametrosECF
            SessaoSalvaEmpresa()

            chkOPT_REFIS.Checked = objEmpresa.ParametrosECF.OPT_REFIS
            chkOPT_PAES.Checked = objEmpresa.ParametrosECF.OPT_PAES
            ddlFORMA_TRIB.SelectedValue = objEmpresa.ParametrosECF.FORMA_TRIB
            DivIsentaImune.Visible = (ddlFORMA_TRIB.SelectedValue = 5 Or ddlFORMA_TRIB.SelectedValue = 8 Or ddlFORMA_TRIB.SelectedValue = 9)
            ddlFORMA_APUR.SelectedValue = objEmpresa.ParametrosECF.FORMA_APUR
            ddlCOD_QUALIF_PJ.SelectedValue = objEmpresa.ParametrosECF.COD_QUALIF_PJ
            ddlFORMA_TRIB_PER.SelectedValue = objEmpresa.ParametrosECF.FORMA_TRIB_PER
            ddlMES_BAL_RED.SelectedValue = objEmpresa.ParametrosECF.MES_BAL_RED
            ddlTIP_ESC_PRE.SelectedValue = objEmpresa.ParametrosECF.TIP_ESC_PRE

            If Trim(objEmpresa.ParametrosECF.TIP_ENT).Length > 0 AndAlso objEmpresa.ParametrosECF.TIP_ENT > 0 Then
                ddlTIP_ENT.SelectedValue = objEmpresa.ParametrosECF.TIP_ENT
            End If

            If Trim(objEmpresa.ParametrosECF.FORMA_APUR_I).Length > 0 Then
                ddlFORMA_APUR_I.SelectedValue = objEmpresa.ParametrosECF.FORMA_APUR_I
            End If

            If Trim(objEmpresa.ParametrosECF.APUR_CSLL).Length > 0 Then
                ddlAPUR_CSLL.SelectedValue = objEmpresa.ParametrosECF.APUR_CSLL
            End If

            chkOPT_EXT_RTT.Checked = objEmpresa.ParametrosECF.OPT_EXT_RTT
            chkDIF_FCONT.Checked = objEmpresa.ParametrosECF.DIF_FCONT

            chkIND_ALIQ_CSLL.Checked = objEmpresa.ParametrosECF.IND_ALIQ_CSLL
            ddlIND_ALIQ_CSLL_I.SelectedIndex = objEmpresa.ParametrosECF.IND_ALIQ_CSLL_I
            txtIND_QTE_SCP.Text = objEmpresa.ParametrosECF.IND_QTE_SCP
            chkIND_ADM_FUN_CLU.Checked = objEmpresa.ParametrosECF.IND_ADM_FUN_CLU
            chkIND_PART_CONS.Checked = objEmpresa.ParametrosECF.IND_PART_CONS
            chkIND_OP_EXT.Checked = objEmpresa.ParametrosECF.IND_OP_EXT
            chkIND_OP_VINC.Checked = objEmpresa.ParametrosECF.IND_OP_VINC
            chkIND_PJ_ENQUAD.Checked = objEmpresa.ParametrosECF.IND_PJ_ENQUAD
            chkIND_PART_EXT.Checked = objEmpresa.ParametrosECF.IND_PART_EXT
            chkIND_ATIV_RURAL.Checked = objEmpresa.ParametrosECF.IND_ATIV_RURAL
            chkIND_LUC_EXP.Checked = objEmpresa.ParametrosECF.IND_LUC_EXP
            chkIND_RED_ISEN.Checked = objEmpresa.ParametrosECF.IND_RED_ISEN
            chkIND_FIN.Checked = objEmpresa.ParametrosECF.IND_FIN
            chkIND_DOA_ELEIT.Checked = objEmpresa.ParametrosECF.IND_DOA_ELEIT
            chkIND_PART_COLIG.Checked = objEmpresa.ParametrosECF.IND_PART_COLIG
            chkIND_VEND_EXP.Checked = objEmpresa.ParametrosECF.IND_VEND_EXP
            chkIND_REC_EXT.Checked = objEmpresa.ParametrosECF.IND_REC_EXT
            chkIND_ATIV_EXT.Checked = objEmpresa.ParametrosECF.IND_ATIV_EXT
            chkIND_COM_EXP.Checked = objEmpresa.ParametrosECF.IND_COM_EXP
            chkIND_PGTO_EXT.Checked = objEmpresa.ParametrosECF.IND_PGTO_EXT
            chkIND_E_COM_TI.Checked = objEmpresa.ParametrosECF.IND_E_COM_TI
            chkIND_ROY_REC.Checked = objEmpresa.ParametrosECF.IND_ROY_REC
            chkIND_ROY_PAG.Checked = objEmpresa.ParametrosECF.IND_ROY_PAG
            chkIND_REND_SERV.Checked = objEmpresa.ParametrosECF.IND_REND_SERV
            chkIND_PGTO_REM.Checked = objEmpresa.ParametrosECF.IND_PGTO_REM
            chkIND_INOV_TEC.Checked = objEmpresa.ParametrosECF.IND_INOV_TEC
            chkIND_CAP_INF.Checked = objEmpresa.ParametrosECF.IND_CAP_INF
            chkIND_POLO_AM.Checked = objEmpresa.ParametrosECF.IND_POLO_AM
            chkIND_ZON_EXP.Checked = objEmpresa.ParametrosECF.IND_ZON_EXP
            chkIND_AREA_COM.Checked = objEmpresa.ParametrosECF.IND_AREA_COM
            chkIND_PJ_HAB.Checked = objEmpresa.ParametrosECF.IND_PJ_HAB

            txtArquivoDeSaida.Text = "SpedECF-" & objEmpresa.Empresa_id & "-" & DateValue(txtDT_INI.Text).ToString("yyyyMMdd") & "-a-" & DateValue(txtDT_FIN.Text).ToString("yyyyMMdd") & ".txt"
            txtArquivoAuxiliar.Text = "SpedECFAuxiliar-" & objEmpresa.Empresa_id & "-" & DateValue(txtDT_INI.Text).ToString("yyyyMMdd") & "-a-" & DateValue(txtDT_FIN.Text).ToString("yyyyMMdd") & ".txt"
        End If
    End Sub

    Private Function Validar() As Boolean
        'Verificacao Registro 0000
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDT_INI.Text) Then
            MsgBox(Me.Page, "Informe o periodo inicial.")
            Return False
        ElseIf Not IsDate(txtDT_INI.Text) Then
            MsgBox(Me.Page, "Valor informado no período inicial não é uma data válida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDT_FIN.Text) Then
            MsgBox(Me.Page, "Informe o período final.")
            Return False
        ElseIf Not IsDate(txtDT_FIN.Text) Then
            MsgBox(Me.Page, "Valor informado no período final não é uma data válida.")
            Return False

            '***********************
            '**** PAT_REMAN_CIS ****
            '***********************
        ElseIf DdlSIT_ESPECIAL.SelectedValue = 6 AndAlso (Not IsNumeric(txtPAT_REMAN_CIS.Text) OrElse (CDec(txtPAT_REMAN_CIS.Text) <= 0)) Then
            MsgBox(Me.Page, "REGRA_PAT_REMAN_CIS_OBRIGATORIO: Verifica se o campo foi preenchido quando 0000.SIT_ESPECIAL for igual a '6' (Cisão Parcial).")
            Return False
        ElseIf DdlSIT_ESPECIAL.SelectedValue <> 6 AndAlso Not (txtPAT_REMAN_CIS.Text.Length = 0) Then
            MsgBox(Me.Page, "REGRA_NAO_PREENCHER_SIT_ESP_CISAO_PARCIAL: Verifica se o campo está em branco quando 0000.SIT_ESPECIAL for diferente de '6' (Cisão Parcial).")
            Return False
            '**********************
            '****  DT_SIT_ESP  ****
            '**********************
        ElseIf DdlSIT_ESPECIAL.SelectedValue <> 0 AndAlso Not IsDate(txtDT_SIT_ESP.Text) Then
            MsgBox(Me.Page, "REGRA_DT_SIT_OBRIGATORIO: Verifica se o campo foi preenchido quando 0000.SIT_ESPECIAL estiver preenchido.")
            Return False
        ElseIf DdlSIT_ESPECIAL.SelectedValue = 0 AndAlso Not (txtDT_SIT_ESP.Text.Length = 0) Then
            MsgBox(Me.Page, "REGRA_NAO_PREENCHER_SIT_ESP_NORMAL: Verifica se campo está me branco quando 0000.SIT_ESPECIAL for igual a '0' (Normal = Sem situação especial no período).")
            Return False
            '**********************
            '******  DT_INI  ******
            '**********************
        ElseIf CDate(txtDT_INI.Text) < CDate("01-01-2014") Then
            MsgBox(Me.Page, "REGRA_DATA_MINIMA: Verifica se 0000.DT_INI é maior que 01/01/2014.")
            Return False
        ElseIf ddlIND_SIT_INI_PER.SelectedValue = 0 AndAlso Not (CDate(txtDT_INI.Text).Day = 1 OrElse CDate(txtDT_INI.Text).Month = 1) Then
            MsgBox(Me.Page, "REGRA_DT_INICIO_ESCRITURACAO: Verifica, quando o 0000.IND_SIT_INI_PER é igual a '0' (Normal), se 0000.DT_INI é igual a 01/01/XXXX.")
            Return False
        ElseIf CDate(txtDT_INI.Text).Year < 2014 Then
            MsgBox(Me.Page, "REGRA_SEM_LEIAUTE: Verifica se, para o ano informado, há leiaute disponível.")
            Return False
        ElseIf CDate(txtDT_INI.Text) > CDate(txtDT_FIN.Text) Then
            MsgBox(Me.Page, "REGRA_DATA_INI_MAIOR: Verifica se 0000.DT_FIN foi preenchido com a data maior que a data informada em 0000.DT_INI.")
            Return False
        ElseIf CDate(txtDT_INI.Text).Year <> CDate(txtDT_FIN.Text).Year Then
            MsgBox(Me.Page, "REGRA_ANO_DIFERENTE: Verifica se o ano informado em 0000.DT_FIN é igual ao ano informado em 0000.DT_INI.")
            Return False
            '**********************
            '******  DT_FIN  ******
            '**********************
        ElseIf DdlSIT_ESPECIAL.SelectedValue = 0 AndAlso Not (CDate(txtDT_FIN.Text).Day = 31 OrElse CDate(txtDT_FIN.Text).Month = 12) Then
            MsgBox(Me.Page, "REGRA_DT_FINAL_ESCRITURACAO: Verifica, quando 0000.SIT_ESPECIAL for igual a '0' (Normal), se o dia e o mês em 0000.DT_FIN é igual a 31/12.")
            Return False
        ElseIf (DdlSIT_ESPECIAL.SelectedValue = 1 OrElse DdlSIT_ESPECIAL.SelectedValue = 2 OrElse DdlSIT_ESPECIAL.SelectedValue = 3 OrElse DdlSIT_ESPECIAL.SelectedValue = 4 OrElse DdlSIT_ESPECIAL.SelectedValue = 5 OrElse DdlSIT_ESPECIAL.SelectedValue = 6) AndAlso Not (CDate(txtDT_FIN.Text) = CDate(txtDT_SIT_ESP.Text)) Then
            MsgBox(Me.Page, "REGRA_EVENTO_ACONTECIMENTO: Verifica: - Quando 0000.SIT_ESPECIAL for igual a '1' (Extinção), '2' (Fusão), '3' (Incorporação/Incorporada), '4' (Incorporação/Incorporadora), '5' (Cisão Total) ou '6' (Cisão Parcial), se 0000.DT_FIN é igual a 0000.DT_SIT_ESP.")
            Return False
        ElseIf (DdlSIT_ESPECIAL.SelectedValue = 7 OrElse DdlSIT_ESPECIAL.SelectedValue = 8 OrElse DdlSIT_ESPECIAL.SelectedValue = 9) AndAlso Not (CDate(txtDT_FIN.Text) = CDate(txtDT_SIT_ESP.Text).AddDays(-1)) Then
            MsgBox(Me.Page, "REGRA_EVENTO_ACONTECIMENTO: Verifica: - Quando 0000.SIT_ESPECIAL for '7' (Transformação), '8' (Desenquadramento de Imune e Isenta ou '9' (Inclusão no Simples Nacional), se 0000.DT_FIN é igual a 0000.DT_SIT_ESP – 1.")
            Return False
            '***********************
            '******  NUM_REC  ******
            '***********************
        ElseIf (ddlRETIFICADORA.SelectedValue = "S" OrElse ddlRETIFICADORA.SelectedValue = "F") AndAlso (txtNUM_REC.Text.Length = 0) Then
            MsgBox(Me.Page, "REGRA_REC_ANTERIOR_OBRIGATORIO: Verifica, quando o campo 0000.RETIFICADORA é igual a 'S' (ECF Retificadora) ou 'F' (ECF original com mudança de forma de tributação), se 0000.NUM_REC está preenchido.")
            Return False
        ElseIf (ddlRETIFICADORA.SelectedValue = "N") AndAlso (txtNUM_REC.Text.Length > 0) Then
            MsgBox(Me.Page, "REGRA_NRO_REC_ANTERIOR_NAO_SE_APLICA: Verifica, quando 0000.RETIFICADORA é igual a 'N' (ECF Original), se 0000.NUM_REC não está preenchido.")
            Return False
        End If

        'Verificacao Registro 0010
        '********************
        '**** FORMA_APUR ****
        '********************
        If ddlFORMA_APUR.SelectedValue = "A" AndAlso Not (((ddlFORMA_TRIB.SelectedValue = 1 OrElse ddlFORMA_TRIB.SelectedValue = 2) OrElse ((ddlFORMA_TRIB.SelectedValue = 3 OrElse ddlFORMA_TRIB.SelectedValue = 4) AndAlso chkOPT_REFIS.Checked))) Then
            MsgBox(Me.Page, "REGRA_FORMA_APUR_VALIDA: Verifica: - Quando 0010.FORMA_APUR igual a 'A' (Anual), se 0010.FORMA_TRIB igual a '1' (Lucro Real) ou '2' (Lucro Real/Arbitrado) ou ['3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado) e 0010.OPT_REFIS igual a 'S'].")
            Return False
        ElseIf ddlFORMA_APUR.SelectedValue = "T" AndAlso Not ((ddlFORMA_TRIB.SelectedValue = 1 OrElse ddlFORMA_TRIB.SelectedValue = 2 OrElse ddlFORMA_TRIB.SelectedValue = 3 OrElse ddlFORMA_TRIB.SelectedValue = 4 OrElse ddlFORMA_TRIB.SelectedValue = 5 OrElse ddlFORMA_TRIB.SelectedValue = 6 OrElse ddlFORMA_TRIB.SelectedValue = 7)) Then
            MsgBox(Me.Page, "REGRA_FORMA_APUR_VALIDA: Verifica: - Quando 0010.FORMA_APUR igual a 'T' (Trimestral), se 0010.FORMA_TRIB igual a '1' (Lucro Real) ou '2' (Lucro Real/Arbitrado) ou '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado) ou '5' (Lucro Presumido) ou '6' (Lucro Arbitrado) ou '7' (Lucro Presumido/Arbitrado).")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (ddlFORMA_APUR.SelectedValue = " ") Then
            MsgBox(Me.Page, "REGRA_ NAO_PREENCHER_IMUNE: Verifica, quando 0010.FORMA_TRIB igual a '8' (Imune do IRPJ) ou '9' (Isento do IPRJ), se 0010.FORMA_APUR não foi preenchido.")
            Return False
            '**********************
            '*** COD_QUALIF_PJ ****
            '**********************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 3 OrElse ddlFORMA_TRIB.SelectedValue = 4 OrElse ddlFORMA_TRIB.SelectedValue = 5 OrElse ddlFORMA_TRIB.SelectedValue = 7) AndAlso Not (ddlCOD_QUALIF_PJ.SelectedValue = 1) Then
            MsgBox(Me.Page, "REGRA_COD_QUALIF_PJ: Verifica, quando 0010.FORMA_TRIB igual a '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado) ou '5' (Lucro Presumido) ou '7' (Lucro Presumido/Arbitrado), se 0010.COD_QUALIF_PJ é igual a '01' (PJ em geral).")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (ddlCOD_QUALIF_PJ.SelectedValue = 0) Then
            MsgBox(Me.Page, "REGRA_ NAO_PREENCHER_IMUNE: Verifica, quando 0010.FORMA_TRIB igual a '8' (Imune do IRPJ) ou '9' (Isento do IPRJ), se 0010.COD_QUALIF_PJ não foi preenchido.")
            Return False
        ElseIf Not ((ddlFORMA_TRIB.SelectedValue = 1 OrElse ddlFORMA_TRIB.SelectedValue = 2 OrElse ddlFORMA_TRIB.SelectedValue = 3 OrElse ddlFORMA_TRIB.SelectedValue = 4 OrElse ddlFORMA_TRIB.SelectedValue = 5 OrElse ddlFORMA_TRIB.SelectedValue = 6 OrElse ddlFORMA_TRIB.SelectedValue = 7) AndAlso ddlCOD_QUALIF_PJ.SelectedValue <> "0") Then
            MsgBox(Me.Page, "REGRA_COD_QUALIF_PJ_OBRIGATORIO: Verifica, quando se 0010.FORMA_TRIB igual a '1' (Lucro Real) ou '2' (Lucro Real/Arbitrado) ou '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado) ou '5' (Lucro Presumido) ou '6' (Lucro Arbitrado) ou '7' (Lucro Presumido/Arbitrado), se 0010.COD_QUALIF_PJ foi preenchido.")
            Return False
            '***********************
            '*** FORMA_TRIB_PER ****
            '***********************
            'ElseIf Not ((ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso ddlFORMA_TRIB_PER.SelectedValue = " ") Then
            '    MsgBox(Me.Page, "REGRA_NAO_PREENCHER_IMUNE: Verifica, quando 0010.FORMA_TRIB igual a '8' (Imune do IRPJ) ou '9' (Isento do IPRJ), se 0010.FORMA_TRIB_PER não foi preenchido.")
            '    Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 1) AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '1' (Lucro Real), então 0010.FORMA_TRIB_PER deve ser igual a '0' ou 'R'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 2) AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '2' (Lucro Real/Arbitrado), então 0010.FORMA_TRIB_PER deve ser igual a '0', 'R', ou 'A'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 3 AndAlso chkOPT_REFIS.Checked = False) AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '3' (Lucro Presumido/Real) e 0010.OPT_REFIS é igual a 'N', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'P' ou 'R'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 3 AndAlso chkOPT_REFIS.Checked AndAlso ddlFORMA_APUR.SelectedValue = "A") AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "E" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '3' (Lucro Presumido/Real) e 0010.OPT_REFIS é igual a 'S' e 0010.FORMA_APUR é igual a 'A', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'E' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 3 AndAlso chkOPT_REFIS.Checked AndAlso ddlFORMA_APUR.SelectedValue = "T") AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '3' (Lucro Presumido/Real) e 0010.OPT_REFIS é igual a 'S' e 0010.FORMA_APUR é igual a 'T', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'R' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 4 AndAlso chkOPT_REFIS.Checked = False) AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '4' (Lucro Presumido/Real/Arbitrado) e 0010.OPT_REFIS é igual a 'N', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'A', 'P' ou 'R'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 4 AndAlso chkOPT_REFIS.Checked AndAlso ddlFORMA_APUR.SelectedValue = "A") AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A" OrElse ddlFORMA_TRIB_PER.SelectedValue = "E" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '4' (Lucro Presumido/Real/Arbitrado) e 0010.OPT_REFIS é igual a 'S' e 0010.FORMA_APUR é igual a 'A', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'A', 'E' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 4 AndAlso chkOPT_REFIS.Checked AndAlso ddlFORMA_APUR.SelectedValue = "T") AndAlso Not ((ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A" OrElse ddlFORMA_TRIB_PER.SelectedValue = "R" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P")) Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '4' (Lucro Presumido/Real/Arbitrado) e 0010.OPT_REFIS é igual a 'S' e 0010.FORMA_APUR é igual a 'T', então 0010.FORMA_TRIB_PER deve ser igual a '0', 'A', 'R' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 5) AndAlso Not (ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P") Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '5' (Lucro Presumido), então 0010.FORMA_TRIB_PER deve ser igual a '0' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 6) AndAlso Not (ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A") Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '6' (Lucro Arbitrado), então 0010.FORMA_TRIB_PER deve ser igual a '0' ou 'A'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 7) AndAlso Not (ddlFORMA_TRIB_PER.SelectedValue = "0" OrElse ddlFORMA_TRIB_PER.SelectedValue = "A" OrElse ddlFORMA_TRIB_PER.SelectedValue = "P") Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '7' (Lucro Presumido/Arbitrado), então 0010.FORMA_TRIB_PER deve ser igual a '0', 'A' ou 'P'.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8) AndAlso Not (ddlFORMA_TRIB_PER.SelectedValue = " ") Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '8' (Imune do IRPJ), então 0010.FORMA_TRIB_PER não deve ser preenchido.")
            Return False
        ElseIf (ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (ddlFORMA_TRIB_PER.SelectedValue = " ") Then
            MsgBox(Me.Page, "REGRA_TRIBUT_INVALIDA: Verificar, Se 0010.FORMA_TRIB é igual a '9' (Isento do IRPJ), então 0010.FORMA_TRIB_PER não deve ser preenchido.")
            Return False
        ElseIf (ddlAPUR_CSLL.SelectedValue.Length = 0) Then
            MsgBox(Me.Page, " Apuração CSLL p/ Imunes/Isentas não foi selecionado.")
            Return False
        End If

        'Verificacao Registro 0020

        '****************************
        '**** 03 - IND_QTE_SCP ******
        '****************************
        'If (ddlTIP_ECF.SelectedValue = 0 OrElse ddlTIP_ECF.SelectedValue = 2) AndAlso Not (txtIND_QTE_SCP.Text.Length = 0) Then
        '    MsgBox(Me.Page, "REGRA_SCP_NAO_PREENCHER_QTD: Verifica, quando 0000.TIP_ECF é igual a '0' (ECF de empresa não participante de SCP como sócio ostensivo)ou '2' (ECF da SCP), se 0020.IND_QTE_SCP não está preenchido.")
        '    Return False
        'Else

        If CDate(txtDT_FIN.Text).Year >= 2015 AndAlso ddlIND_ALIQ_CSLL_I.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione a Aliquota")
            Return False
        ElseIf (ddlTIP_ECF.SelectedValue = 1) AndAlso (txtIND_QTE_SCP.Text.Length = 0) Then
            MsgBox(Me.Page, "REGRA_SCP_OBRIGATORIO_QTD: Verifica, quando 0000.TIP_ECF é igual a '1' (ECF de empresa participante de SCP como sócio ostensivo), se 0020.IND_QTE_SCP está preenchido.")
            Return False
            '*********************************
            '****** 05 - IND_PART_CONS  ******
            '*********************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso ddlTIP_ENT.SelectedValue = 13 AndAlso Not (chkIND_PART_CONS.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0010.IND_PART_CONS é igual a 'N'.")
            Return False
            '*********************************
            '****** 06 - IND_OP_EXT  *********
            '*********************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_OP_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune do IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_OP_EXT é igual a 'N'.")
            Return False
            '*********************************
            '****** 07 - IND_OP_VINC  ********
            '*********************************
        ElseIf chkIND_OP_EXT.Checked = False And Not (chkIND_OP_VINC.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_OP_VINC: Verifica, quando 0020.IND_OP_EXT é igual a 'N', se 0020.IND_OP_VINC também é igual a 'N'.")
            Return False
            '**********************************
            '****** 09 - IND_PART_EXT  ********
            '**********************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_PART_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_PART_EXT: Verifica, Se 0010.FORMA_TRIB igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), então 0020.IND_PART_EXT deve ser igual a 'N'.")
            Return False
        ElseIf ((ddlFORMA_TRIB.SelectedValue = 5 OrElse ddlFORMA_TRIB.SelectedValue = 7) AndAlso Not chkOPT_REFIS.Checked) AndAlso Not (chkIND_PART_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_PART_EXT: Verifica, Se [0010.FORMA_TRIB é igual a '5' (Lucro Presumido) ou '7' (Lucro Presumido/Arbitrado)) e 0010.OPT_REFIS igual a 'N'], então IND_PART_EXT deveser igual a 'N'.")
            Return False
            '************************************
            '****** 10 - IND_ATIV_RURAL  ********
            '************************************
            'ElseIf (ddlCOD_QUALIF_PJ.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 2 OrElse ddlFORMA_TRIB.SelectedValue <> 3 OrElse ddlFORMA_TRIB.SelectedValue <> 4) AndAlso Not (chkIND_ATIV_RURAL.Checked = False) Then
            '    MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_ATIV_RURAL: Verifica, quando 0010.COD_QUALIF_PJ é diferente de '01' (PJ em Geral) ou 0010.FORMA_TRIB é diferente de '1' (Lucro Real), '2' (Lucro Real/Arbitrado), '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado), se 0020.IND_ATIV_RURAL é igual a 'N'.")
            '    Return False
            '************************************
            '****** 11 - IND_LUC_EXP  ***********
            '************************************
        ElseIf (ddlCOD_QUALIF_PJ.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 2 OrElse ddlFORMA_TRIB.SelectedValue <> 3 OrElse ddlFORMA_TRIB.SelectedValue <> 4) AndAlso Not (chkIND_LUC_EXP.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_LUC_EXP: Verifica, quando 0010.COD_QUALIF_PJ é diferente de '01' (PJ em Geral) ou 0010.FORMA_TRIB é diferente de '1' (Lucro Real), '2' (Lucro Real/Arbitrado), '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado), se 0020.IND_LUC_EXP é igual a 'N'.")
            Return False
            '************************************
            '****** 12 - IND_RED_ISEN  **********
            '************************************
        ElseIf (chkOPT_REFIS.Checked OrElse ddlCOD_QUALIF_PJ.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 5 OrElse ddlFORMA_TRIB.SelectedValue <> 7) AndAlso Not (chkIND_RED_ISEN.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_RED_ISEN: Verifica, quando 0010.OPT_REFIS é diferente de 'S' ou 0010.COD_QUALIF_PJ é diferente de '01' (PJ em Geral) ou 0010.FORMA_TRIB é diferente de '5' (Lucro Presumido) ou '7' (Lucro Presumido/Arbitrado), se 0020.IND_RED_ISEN é igual a 'N'.")
            Return False
            '************************************
            '****** 13 - IND_FIN  ***************
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue <> 1 OrElse ddlFORMA_TRIB.SelectedValue <> 2 OrElse ddlFORMA_TRIB.SelectedValue <> 3 OrElse ddlFORMA_TRIB.SelectedValue <> 4) AndAlso Not (chkIND_FIN.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IND_FIN: Verifica, quando 0010.FORMA_TRIB é diferente de '1' (Lucro Real), '2' (Lucro Real/Arbitrado), '3' (Lucro Presumido/Real) ou '4' (Lucro Presumido/Real/Arbitrado), se 0020.IND_FIN é igual a 'N'.")
            Return False
            '************************************
            '****** 15 - IND_PART_COLIG  ********
            '************************************
        ElseIf (ddlTIP_ECF.SelectedValue = 13) AndAlso Not (chkIND_PART_COLIG.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_PART_COLIG é igual a 'N'.")
            Return False
            '************************************
            '****** 16 - IND_VEND_EXP  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_VEND_EXP.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune do IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_VEND_EXP é igual a 'N'.")
            Return False
            'ElseIf (ddlCOD_QUALIF_PJ.SelectedValue <> 0) AndAlso Not (chkIND_VEND_EXP.Checked = False) Then
            '    MsgBox(Me.Page, "REGRA_PREENCHIMENTO_PJ: Verifica, quando 0010.COD_QUALIF_PJ é diferente de ''00' (PJ em Geral), se 0020.IND_VEND_EXP é igual a 'N'.")
            '    Return False
            '************************************
            '****** 17 - IND_ REC_EXT  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_REC_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_REC_EXT é igual a 'N'.")
            Return False
            '************************************
            '****** 18 - IND_ATIV_EXT  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_ATIV_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_ATIV_EXT é igual a 'N'.")
            Return False
        ElseIf (ddlCOD_QUALIF_PJ.SelectedValue <> 1) AndAlso Not (chkIND_ATIV_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_PJ: Verifica, quando 0010.COD_QUALIF_PJ diferente de '01' (PJ em Geral), se 0020.IND_ATIV_EXT é igual a 'N'.")
            Return False
            '************************************
            '****** 19 - IND_COM_EXP  ***********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_COM_EXP.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_COM_EXP é igual a 'N'.")
            Return False
        ElseIf (ddlCOD_QUALIF_PJ.SelectedValue <> 1) AndAlso Not (chkIND_COM_EXP.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_PJ_00: Verifica, quando 0010.COD_QUALIF_PJ é diferente de '01' (PJ em Geral), se 0020.IND_COM_EXP é igual a 'N'.")
            Return False
            '************************************
            '****** 20 - IND_PGTO_EXT  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_PGTO_EXT.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_PGTO_EXT é igual a 'N'.")
            Return False
            '************************************
            '****** 21 - IND_E_COM_TI  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_E_COM_TI.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_E-COM_TI é igual a 'N'.")
            Return False
            '************************************
            '****** 22 - IND_ROY_PAG   **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_ROY_PAG.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_ROY_REC é igual a 'N'.")
            Return False
            '************************************
            '****** 23 - IND_ROY_PAG  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_ROY_PAG.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_ROY_PAG é igual a 'N'.")
            Return False
            '************************************
            '****** 24 - IND_REND_SERV  *********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_REND_SERV.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_REND_SERV é igual a 'N'.")
            Return False
            '************************************
            '****** 25 - IND_PGTO_REM  **********
            '************************************
        ElseIf ddlTIP_ENT.SelectedValue > 0 AndAlso (ddlTIP_ENT.SelectedValue = 13) AndAlso Not (chkIND_PGTO_REM.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_ATIV_13: Verifica, quando 0010_TIP_ENT é igual a '13' (Fifa e Entidades Relacionadas), se 0020.IND_PGTO_REM é igual a 'N'.")
            Return False
            '************************************
            '****** 26 - IND_INOV_TEC  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_INOV_TEC.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_INOV_TEC é igual a 'N'.")
            Return False
            '************************************
            '****** 27 - IND_CAP_INF  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_CAP_INF.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_CAP_INF é igual a 'N'.")
            Return False
            '************************************
            '****** 28 - IND_PJ_HAB  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_PJ_HAB.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_PJ_HAB é igual a 'N'.")
            Return False
            '************************************
            '****** 29 - IND_POLO_AM  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_POLO_AM.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_POLO_AM é igual a 'N'.")
            Return False
            '************************************
            '****** 30 - IND_ZON_EXP  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_ZON_EXP.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_ZON_EXP é igual a 'N'.")
            Return False
            '************************************
            '****** 31 - IND_AREA_COM  **********
            '************************************
        ElseIf (ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9) AndAlso Not (chkIND_AREA_COM.Checked = False) Then
            MsgBox(Me.Page, "REGRA_PREENCHIMENTO_IMUNE_ISENTA: Verifica, quando 0010.FORMA_TRIB é igual a '8' (Imune de IRPJ) ou '9' (Isento do IPRJ), se 0020.IND_AREA_COM é igual a 'N'.")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Dim Emp() As String = DdlEmpresa.SelectedValue.Split("-")
        txtArquivoDeSaida.Text = "SpedECF-" & Emp(0) & "-" & DateValue(txtDT_INI.Text).ToString("yyyyMMdd") & "-a-" & DateValue(txtDT_FIN.Text).ToString("yyyyMMdd") & ".txt"
        txtArquivoAuxiliar.Text = "SpedECFAuxiliar-" & Emp(0) & "-" & DateValue(txtDT_INI.Text).ToString("yyyyMMdd") & "-a-" & DateValue(txtDT_FIN.Text).ToString("yyyyMMdd") & ".txt"

        If Not Validar() Then Exit Sub

        If String.IsNullOrWhiteSpace(txtArquivoDeSaida.Text) Then
            MsgBox(Me.Page, "É necessário selecionar o caminho do arquivo de saída!")
            Exit Sub
        End If

        Dim NomeArquivo As String
        Dim NomeArquivo2 As String = "ECF/" & txtArquivoDeSaida.Text.ToUpper
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Using Strm As New StreamWriter(Server.MapPath("~/ECF/" & txtArquivoDeSaida.Text.Trim().ToUpper()), True, Encoding.Default)
            Try
                Dim sql As String = ""
                Dim Empresa As New Cliente(Emp(0), Emp(1))

                Dim PeriodoInicial As Date = DateValue(txtDT_INI.Text).ToString("yyyy/MM/dd")
                Dim PeriodoFinal As Date = DateValue(txtDT_FIN.Text).ToString("yyyy/MM/dd")
                Dim Trimestre As String = String.Empty

                cmdArquivoDeSaida.Visible = False
                cmdArquivoAuxDeSaida.Visible = False

                Dim linha As String = String.Empty
                Dim RegistroGeral As Integer = 0    'Numero de linhas total do Arquivo
                '******************************************************************************************************
                '******** REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA   *************
                '******************************************************************************************************
                Dim PeriodoEspecial As String = ""
                If DdlSIT_ESPECIAL.SelectedValue <> 0 Then PeriodoEspecial = CDate(txtDT_SIT_ESP.Text).ToString("ddMMyyyy")
                CompoeRegistro0000(Strm, Empresa, PeriodoInicial, PeriodoFinal, PeriodoEspecial, linha, RegistroGeral)

                '******************************************************************************************************
                '**************************** REGISTRO 0001: ABERTURA DO BLOCO 0   ************************************
                '******************************************************************************************************

                CompoeRegistro0001(Strm, linha, RegistroGeral)

                '******************************************************************************************************
                '*************************** REGISTRO 0010: PARÂMETROS DE TRIBUTAÇÃO   ********************************
                '******************************************************************************************************

                CompoeRegistro0010(Strm, linha, RegistroGeral, PeriodoFinal)

                '******************************************************************************************************
                '************************** REGISTRO 0020: PARÂMETROS COMPLEMENTARES   ********************************
                '******************************************************************************************************

                CompoeRegistro0020(Strm, PeriodoFinal, linha, RegistroGeral)

                '******************************************************************************************************
                '**************************  REGISTRO 0030: DADOS CADASTRAIS   ********************************
                '******************************************************************************************************

                CompoeRegistro0030(Strm, Empresa, linha, RegistroGeral)

                '******************************************************************************************************
                '********************* REGISTRO 0930: IDENTIFICAÇÃO DOS SIGNATÁRIOS DA ECF   **************************
                '******************************************************************************************************

                CompoeRegistro0930(Strm, Empresa, linha, RegistroGeral)

                '******************************************************************************************************
                '**************************  REGISTRO 0990: ENCERRAMENTO DO BLOCO 0   ********************************
                '******************************************************************************************************

                CompoeRegistro0990(Strm, linha, RegistroGeral)

                '******************************************************************************************************
                '**************************** REGISTRO J001: ABERTURA DO BLOCO J   ************************************
                '******************************************************************************************************

                Dim RegistroJ As Integer

                CompoeRegistroJ001(Strm, linha, RegistroJ, RegistroGeral)

                '******************************************************************************************************
                '*******************    REGISTRO J050: PLANO DE CONTAS DO CONTRIBUINTE     ****************************
                '*******************    REGISTRO J051: PLANO DE CONTAS REFERENCIAL         ****************************
                '******************************************************************************************************

                CompoeRegistroJ050_J051(Strm, Empresa, linha, RegistroJ, RegistroGeral, DateValue(txtDT_FIN.Text))

                '******************************************************************************************************
                '**************************  REGISTRO J990: ENCERRAMENTO DO BLOCO J   ********************************
                '******************************************************************************************************

                CompoeRegistroJ990(Strm, linha, RegistroJ, RegistroGeral)

                '******************************************************************************************************
                '**************************** REGISTRO K001: ABERTURA DO BLOCO K   ************************************
                '******************************************************************************************************

                Dim RegistroK As Integer

                CompoeRegistroK001(Strm, linha, RegistroK, RegistroGeral)

                ' ''***************************************************************************************************************
                ' ''*****  REGISTRO K030: IDENTIFICAÇÃO DO PERÍODO E FORMAS DE APURAÇÃO DO IRPJ E DA CSLL NO ANO-CALENDÁRIO  ******
                ' ''***************************************************************************************************************

                If ddlFORMA_APUR.SelectedValue = "T" Then 'Trimestral
                    'Quatro Trimestres - T01, T02, T3 e T04
                    Trimestre = String.Empty
                    Dim DataInicial As Date = DateValue(txtDT_INI.Text)
                    'Dim DataFinal As Date = DateValue(txtDT_FIN.Text).AddMonths(-9) 'Neri 15;10;2025 Fica derrado quando o mes final for menos que 12
                    Dim DataFinal As Date = DataInicial.AddMonths(3)
                    DataFinal = DataFinal.AddDays(-1)

                    For index = 1 To 4
                        Select Case index
                            Case 1
                                Trimestre = "T01"

                                'If Left(Empresa.Codigo, 8) = "38388314" Then
                                '    DataInicial = CDate("01/01/2021")
                                '    DataFinal = CDate("31/03/2021")
                                'End If
                            Case 2
                                Trimestre = "T02"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)

                                If Left(Empresa.Codigo, 8) = "40938762" AndAlso PeriodoFinal.Year.ToString() = "2021" Then
                                    DataInicial = CDate("01/04/2021")
                                End If
                            Case 3
                                Trimestre = "T03"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)

                                If Left(Empresa.Codigo, 8) = "40938762" AndAlso PeriodoFinal.Year.ToString() = "2021" Then
                                    DataInicial = CDate("01/07/2021")
                                End If
                            Case Else
                                Trimestre = "T04"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)
                                DataFinal = DataFinal.AddDays(1)

                                'If Left(Empresa.Codigo, 8) = "38388314" Then
                                '    Continue For
                                'End If
                                If Left(Empresa.Codigo, 8) = "40938762" AndAlso PeriodoFinal.Year.ToString() = "2021" Then
                                    DataInicial = CDate("01/10/2021")
                                End If
                        End Select

                        CompoeRegistroK030(Strm, Trimestre, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)

                        '***************************************************************************************************************
                        '*****  REGISTRO K155: Detalhes dos Saldos Contábeis (Depois do Encerramento do Resultado do Período)  ******
                        '*****  REGISTRO K156: Mapeamento Referencial do Saldo Final  **************************************************
                        '***************************************************************************************************************

                        CompoeRegistroK155_K156(Strm, Empresa, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)

                        '*************************************************************************************************************************
                        '*****  REGISTRO K355: Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento  ****************************
                        '*****  REGISTRO K356: Mapeamento Referencial dos Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento **
                        '*************************************************************************************************************************
                        CompoeRegistroK355_K356(Strm, Empresa, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)

                    Next
                End If


                If ddlFORMA_APUR.SelectedValue = "A" Then 'Anual
                    'Anual = - A00, A01, A02, A03, A04, A05, A06, A07, A08, A09, A10, A11, A12
                    Trimestre = String.Empty
                    Dim DataInicial As Date = DateValue(txtDT_INI.Text)
                    Dim DataFinal As Date = DateValue(txtDT_FIN.Text)

                    For index = 1 To 13
                        Select Case index
                            Case 1
                                Trimestre = "A00"
                            Case 2
                                Trimestre = "A01"
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-11)
                            Case 3
                                Trimestre = "A02"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-10)
                            Case 4
                                Trimestre = "A03"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-9)
                            Case 5
                                Trimestre = "A04"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-8)
                            Case 6
                                Trimestre = "A05"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-7)
                            Case 7
                                Trimestre = "A06"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-6)
                            Case 8
                                Trimestre = "A07"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-5)
                            Case 9
                                Trimestre = "A08"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-4)
                            Case 10
                                Trimestre = "A09"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-3)
                            Case 11
                                Trimestre = "A10"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-2)
                            Case 12
                                Trimestre = "A11"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text).AddMonths(-1)
                            Case 13
                                Trimestre = "A12"
                                'DataInicial = DataInicial.AddMonths(1)
                                DataFinal = DateValue(txtDT_FIN.Text)
                        End Select


                        CompoeRegistroK030(Strm, Trimestre, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)

                        '***************************************************************************************************************
                        '*****  REGISTRO K155: Detalhes dos Saldos Contábeis (Depois do Encerramento do Resultado do Período)  ******
                        '*****  REGISTRO K156: Mapeamento Referencial do Saldo Final  **************************************************
                        '***************************************************************************************************************

                        CompoeRegistroK155_K156(Strm, Empresa, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)

                        '*************************************************************************************************************************
                        '*****  REGISTRO K355: Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento  ****************************
                        '*****  REGISTRO K356: Mapeamento Referencial dos Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento **
                        '*************************************************************************************************************************
                        CompoeRegistroK355_K356(Strm, Empresa, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)
                    Next
                End If


                '******************************************************************************************************
                '************************************** FECHAMENTO DO BLOCO K *****************************************
                '******************************************************************************************************

                '******************************************************************************************************
                '***************************************    ABERTURA DO BLOCO L   *************************************
                '******************************************************************************************************


                Dim RegistroL As Integer
                CompoeRegistroL001(Strm, linha, RegistroL, RegistroGeral)


                ''******************************************************************************************************
                ''**** Registro L030:  ****************************
                ''******************************************************************************************************


                If Not ddlFORMA_TRIB.SelectedValue = 5 AndAlso ddlFORMA_APUR.SelectedValue = "T" Then 'Trimestral
                    'Quatro Trimestres - T01, T02, T3 e T04
                    Trimestre = String.Empty
                    Dim DataInicial As Date = DateValue(txtDT_INI.Text)

                    'Dim DataFinal As Date = DateValue(txtDT_FIN.Text).AddMonths(-9)
                    'Dim DataFinal As Date = DateValue(txtDT_FIN.Text).AddMonths(-9) 'Neri 15;10;2025 Fica derrado quando o mes final for menos que 12
                    Dim DataFinal As Date = DataInicial.AddMonths(3)
                    DataFinal = DataFinal.AddDays(-1)

                    For index = 1 To 4
                        Select Case index
                            Case 1
                                Trimestre = "T01"
                            Case 2
                                Trimestre = "T02"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)
                            Case 3
                                Trimestre = "T03"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)
                            Case Else
                                Trimestre = "T04"
                                DataInicial = DataInicial.AddMonths(3)
                                DataFinal = DataFinal.AddMonths(3)
                                DataFinal = DataFinal.AddDays(1)
                        End Select

                        CompoeRegistroL030(Strm, Trimestre, DataInicial, DataFinal, linha, RegistroL, RegistroGeral)

                        '***************************************************************************************************************
                        '*****  REGISTRO K155: Detalhes dos Saldos Contábeis (Depois do Encerramento do Resultado do Período)  ******
                        '*****  REGISTRO K156: Mapeamento Referencial do Saldo Final  **************************************************
                        '***************************************************************************************************************

                        CompoeRegistroL200(Strm, linha, RegistroL, RegistroGeral)

                        '*************************************************************************************************************************
                        '*****  REGISTRO K355: Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento  ****************************
                        '*****  REGISTRO K356: Mapeamento Referencial dos Saldos Finais das Contas Contábeis de Resultado Antes do Encerramento **
                        '*************************************************************************************************************************
                        'CompoeRegistroK355_K356(Strm, Empresa, DataInicial, DataFinal, linha, RegistroK, RegistroGeral)
                    Next
                End If


                '******************************************************************************************************
                '**************************** REGISTRO L: FECHAMENTO DO BLOCO L   **********************************
                '******************************************************************************************************



                '******************************************************************************************************
                '***************************************    ABERTURA DO BLOCO M   *************************************
                '******************************************************************************************************


                'Dim RegistroM As Integer
                'CompoeRegistroM001(Strm, linha, RegistroM, RegistroGeral)


                '******************************************************************************************************
                '**** Registro M030:  ****************************
                '******************************************************************************************************


                '******************************************************************************************************
                '**************************** REGISTRO M: FECHAMENTO DO BLOCO M   **********************************
                '******************************************************************************************************


                '******************************************************************************************************
                '**************************** REGISTRO Y001: ABERTURA DO BLOCO Y   ************************************
                '******************************************************************************************************

                Dim RegistroY As Integer
                CompoeRegistroY001(Strm, linha, RegistroY, RegistroGeral)

                '******************************************************************************************************
                '**** Registro Y520: Pagamentos/Recebimentos do Exterior ou de Não Residentes *************************
                '******************************************************************************************************

                CompoeRegistroY520(Strm, Empresa, linha, RegistroY, RegistroGeral)


                '******************************************************************************************************
                '**** Registro Y540: Discriminação da Receita de Vendas dos Estabelecimentos por Atividade Econômica **
                '******************************************************************************************************
                Dim Leiaute As String = GetLeiaute(PeriodoFinal.Year, PeriodoFinal.Month)
                'Ajuste para Layout à partir de 2020 Y540 deixou de existir - Furlan - 24/08/2020
                If CInt(Leiaute) < 7 Then
                    If chkIND_VEND_EXP.Checked Then
                        CompoeRegistroY540(Strm, Empresa, linha, RegistroY, RegistroGeral)
                    End If
                End If


                '******************************************************************************************************
                '**** Registro Y550: Vendas a Comercial Exportadora com Fim Específico de Exportação ******************
                '******************************************************************************************************
                If chkIND_REC_EXT.Checked Then
                    CompoeRegistroY550(Strm, linha, RegistroY, RegistroGeral)
                End If

                '******************************************************************************************************
                '**** Registro Y560: Detalhamento das Exportações da Comercial Exportadora ****************************
                '******************************************************************************************************
                If chkIND_COM_EXP.Checked Then
                    CompoeRegistroY560(Strm, linha, RegistroY, RegistroGeral)
                End If

                '******************************************************************************************************
                '************************************** FECHAMENTO DO BLOCO Y *****************************************
                '******************************************************************************************************

                cmdArquivoDeSaida.Visible = True
                cmdArquivoAuxDeSaida.Visible = True
                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "downloadArquivo();", True)
            Catch ex As Exception
                Throw New Exception(ex.Message)
            Finally
                Strm.Close()
                Strm.Dispose()
            End Try
        End Using

    End Sub

#Region "Registro 0"
    Private Function GetLeiaute(ByRef Ano As Integer, ByRef Mes As Integer) As String
        Dim Leiaute As String = String.Empty

        Select Case Ano
            Case Is < 2015
                Return "0001"
            Case Is = 2015
                Return "0002"
            Case Is = 2016
                Return "0003"
            Case Is = 2017
                Return "0004"
            Case Is = 2018
                Return "0005"
            Case Is = 2019
                Return "0006"
            Case Is = 2020
                Return "0007"
            Case Is = 2021
                Return "0008"
            Case Is = 2022
                Return "0009"
            Case Is = 2023
                Return "0010"
            Case Is = 2024
                Return "0011"
            Case Else
                Return "0011"
        End Select
    End Function

    Private Sub CompoeRegistro0000(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef PeriodoInicial As Date, ByRef PeriodoFinal As Date, ByRef PeriodoEspecial As String, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '******** REGISTRO 0000: ABERTURA DO ARQUIVO DIGITAL E IDENTIFICAÇÃO DA PESSOA JURÍDICA   *************
        '******************************************************************************************************

        'Leiaute Mudar conforme as datas.

        Dim Leiaute As String = GetLeiaute(PeriodoFinal.Year, PeriodoFinal.Month)

        'Dim PeriodoEspecial As String = ""
        'If DdlSIT_ESPECIAL.SelectedValue <> 0 Then PeriodoEspecial = CDate(txtDT_SIT_ESP.Text).ToString("yyyyMMdd")

        linha = "|0000"                                            '01 - REG      
        linha &= "|LECF"                                           '02 - NOME_ESC 
        linha &= "|" & Leiaute                                     '03 - COD_VER 
        linha &= "|" & Empresa.Codigo                              '04 - CNPJ
        linha &= "|" & Empresa.Nome                                '05 - NOME
        linha &= "|" & ddlIND_SIT_INI_PER.SelectedValue            '06 - IND_SIT_INI_PER  
        linha &= "|" & DdlSIT_ESPECIAL.SelectedValue               '07 - SIT_ESPECIAL  
        linha &= "|" & txtPAT_REMAN_CIS.Text                       '08 - PAT_REMAN_CIS 
        linha &= "|" & PeriodoEspecial                             '09 - DT_SIT_ESP 
        linha &= "|" & PeriodoInicial.ToString("ddMMyyyy")  '10 - DT_INI
        linha &= "|" & PeriodoFinal.ToString("ddMMyyyy")  '11 - DT_FIN
        linha &= "|" & ddlRETIFICADORA.SelectedValue               '12 - RETIFICADORA 
        linha &= "|" & txtNUM_REC.Text                             '13 - NUM_REC
        linha &= "|" & ddlTIP_ECF.SelectedValue                    '14 - TIP_ECF 
        linha &= "|"                                               '15 - COD_SCP 
        linha &= "|"

        Strm.WriteLine(linha)
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistro0001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO 0001: ABERTURA DO BLOCO 0   ************************************
        '******************************************************************************************************

        linha = "|0001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistro0010(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroGeral As Integer, ByRef PeriodoFinal As Date)
        '******************************************************************************************************
        '*************************** REGISTRO 0010: PARÂMETROS DE TRIBUTAÇÃO   ********************************
        '******************************************************************************************************
        Dim Leiaute As String = GetLeiaute(PeriodoFinal.Year, PeriodoFinal.Month)

        linha = "|0010"                                           '01 - REG
        linha &= "|"                                              '02 - HASH_ECF_ANTERIOR
        linha &= "|" & IIf(chkOPT_REFIS.Checked, "S", "N")        '03 - OPT_REFIS

        'Ajuste para Layout à partir de 2020 - Furlan - 19/08/2020
        If CInt(Leiaute) < 7 Then
            linha &= "|" & IIf(chkOPT_PAES.Checked, "S", "N")         '04 - OPT_PAES
        End If

        linha &= "|" & ddlFORMA_TRIB.SelectedValue                '05 - FORMA_TRIB

        linha &= "|" & ddlFORMA_APUR.SelectedValue                '06 - FORMA_APUR

        linha &= "|" & CInt(ddlCOD_QUALIF_PJ.SelectedValue).ToString("00")  '07 - COD_QUALIF_PJ 

        'If cbTrimestral.Checked Then
        '    linha &= "|" & Replace(Space(4), " ", ddlFORMA_TRIB_PER.SelectedValue)   '08 - FORMA_TRIB_PER
        'Else
        '    linha &= "|" & ddlFORMA_TRIB_PER.SelectedValue            '08 - FORMA_TRIB_PER
        'End If


        If ddlFORMA_APUR.SelectedValue = "T" Then
            linha &= "|" & Replace(Space(4), " ", ddlFORMA_TRIB_PER.SelectedValue)   '08 - FORMA_TRIB_PER
            linha &= "|" '& ddlMES_BAL_RED.SelectedValue               '09 - MES_BAL_RED 
        Else
            linha &= "|" & Replace(Space(4), " ", ddlFORMA_TRIB_PER.SelectedValue)   '08 - FORMA_TRIB_PER
            linha &= "|" & Replace(Space(12), " ", ddlMES_BAL_RED.SelectedValue)               '09 - MES_BAL_RED 

            'linha &= "|" & ddlFORMA_TRIB_PER.SelectedValue            '08 - FORMA_TRIB_PER
        End If

        'linha &= "|" & ddlMES_BAL_RED.SelectedValue               '09 - MES_BAL_RED 

        '3,4,5,7,8,9
        If ddlFORMA_TRIB.SelectedValue = 3 OrElse ddlFORMA_TRIB.SelectedValue = 4 OrElse ddlFORMA_TRIB.SelectedValue = 5 _
            OrElse ddlFORMA_TRIB.SelectedValue = 7 OrElse ddlFORMA_TRIB.SelectedValue = 8 OrElse ddlFORMA_TRIB.SelectedValue = 9 Then
            linha &= "|" & ddlTIP_ESC_PRE.SelectedValue               '10 - TIP_ESC_PRE
        Else
            linha &= "|"        '10 - TIP_ESC_PRE
        End If


        If ddlFORMA_TRIB.SelectedValue = 5 Or ddlFORMA_TRIB.SelectedValue = 8 Or ddlFORMA_TRIB.SelectedValue = 9 Then
            If ddlTIP_ENT.SelectedValue > 0 Then
                linha &= "|" & ddlTIP_ENT.SelectedValue                  '11 - TIP_ENT
            Else
                linha &= "|"
            End If

            If ddlFORMA_TRIB.SelectedValue = 5 Then
                linha &= "||"                                            ' 12 - 13
            Else
                linha &= "|" & ddlFORMA_APUR_I.SelectedValue              '12 - FORMA_APUR_I  
                linha &= "|" & ddlAPUR_CSLL.SelectedValue                 '13 - APUR_CSLL
            End If

        Else
            linha &= "|||"                                            '11 - 12 - 13
        End If

        If PeriodoFinal.Year > 2016 Then
            linha &= "|" & ddlIND_REC_RECEITA.SelectedValue           '14 - TIP_ENT
        Else
            linha &= "|" '& IIf(chkOPT_EXT_RTT.Checked, "S", "N")     '14 - PT_EXT_RTT
        End If

        'linha &= "|" & IIf(chkDIF_FCONT.Checked, "S", "N")        '15 - DIF_FCONT
        linha &= "|"

        Strm.WriteLine(linha)
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistro0020(ByRef Strm As StreamWriter, ByRef PeriodoFinal As Date, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '************************** REGISTRO 0020: PARÂMETROS COMPLEMENTARES   ********************************
        '******************************************************************************************************
        Dim Leiaute As String = GetLeiaute(PeriodoFinal.Year, PeriodoFinal.Month)

        linha = "|0020"                                           '01 - REG

        If PeriodoFinal.Year >= 2015 Then
            linha &= "|" & ddlIND_ALIQ_CSLL_I.SelectedIndex             '02 - IND_ALIQ_CSLL
        Else
            linha &= "|" & IIf(chkIND_ALIQ_CSLL.Checked, "S", "N")   '02 - IND_ALIQ_CSLL
        End If

        linha &= "|" & txtIND_QTE_SCP.Text '03 - IND_QTE_SCP
        linha &= "|" & IIf(chkIND_ADM_FUN_CLU.Checked, "S", "N")  '04 - IND_ADM_FUN_CLU
        linha &= "|" & IIf(chkIND_PART_CONS.Checked, "S", "N")    '05 - IND_PART_CONS
        linha &= "|" & IIf(chkIND_OP_EXT.Checked, "S", "N")       '06 - IND_OP_EXT
        linha &= "|" & IIf(chkIND_OP_VINC.Checked, "S", "N")      '07 - IND_OP_VINC
        linha &= "|" & IIf(chkIND_PJ_ENQUAD.Checked, "S", "N")    '08 - IND_PJ_ENQUAD
        linha &= "|" & IIf(chkIND_PART_EXT.Checked, "S", "N")     '09 - IND_PART_EXT
        linha &= "|" & IIf(chkIND_ATIV_RURAL.Checked, "S", "N")   '10 - IND_ATIV_RURAL
        linha &= "|" & IIf(chkIND_LUC_EXP.Checked, "S", "N")      '11 - IND_LUC_EX
        linha &= "|" & IIf(chkIND_RED_ISEN.Checked, "S", "N")     '12 - IND_RED_ISEN
        linha &= "|" & IIf(chkIND_FIN.Checked, "S", "N")          '13 - IND_FIN

        'Ajuste para Layout à partir de 2020 - Furlan - 19/08/2020
        If CInt(Leiaute) < 7 Then
            linha &= "|" & IIf(chkIND_DOA_ELEIT.Checked, "S", "N")    '14 - IND_DOA_ELEIT
        End If

        linha &= "|" & IIf(chkIND_PART_COLIG.Checked, "S", "N")   '14 - IND_PART_COLIG

        If CInt(Leiaute) < 7 Then
            linha &= "|" & IIf(chkIND_VEND_EXP.Checked, "S", "N")     '16 - IND_VEND_EXP
        End If

        linha &= "|" & IIf(chkIND_REC_EXT.Checked, "S", "N")      '15 - IND_REC_EXT
        linha &= "|" & IIf(chkIND_ATIV_EXT.Checked, "S", "N")     '16 - IND_ATIV_EXT

        If CInt(Leiaute) < 7 Then
            linha &= "|" & IIf(chkIND_COM_EXP.Checked, "S", "N")      '19 - IND_COM_EXP
        End If

        linha &= "|" & IIf(chkIND_PGTO_EXT.Checked, "S", "N")     '17 - IND_PGTO_EXT
        linha &= "|" & IIf(chkIND_E_COM_TI.Checked, "S", "N")     '18 - IND_E_COM_TI
        linha &= "|" & IIf(chkIND_ROY_REC.Checked, "S", "N")      '19 - IND_ROY_REC
        linha &= "|" & IIf(chkIND_ROY_PAG.Checked, "S", "N")      '20 - IND_ROY_PAG
        linha &= "|" & IIf(chkIND_REND_SERV.Checked, "S", "N")    '21 - IND_REND_SERV
        linha &= "|" & IIf(chkIND_PGTO_REM.Checked, "S", "N")     '22 - IND_PGTO_REM
        linha &= "|" & IIf(chkIND_INOV_TEC.Checked, "S", "N")     '23 - IND_INOV_TEC
        linha &= "|" & IIf(chkIND_CAP_INF.Checked, "S", "N")      '24 - IND_CAP_INF
        linha &= "|" & IIf(chkIND_PJ_HAB.Checked, "S", "N")       '25 - IND_PJ_HAB
        linha &= "|" & IIf(chkIND_POLO_AM.Checked, "S", "N")      '26 - IND_POLO_AM
        linha &= "|" & IIf(chkIND_ZON_EXP.Checked, "S", "N")      '27 - IND_ZON_EXP
        linha &= "|" & IIf(chkIND_AREA_COM.Checked, "S", "N")     '28 - IND_AREA_COM

        linha &= "|N" '29 - IND_PAIS_A_PAIS
        linha &= "|N" '30 IND_DEREX - Furlan 09/07/2018
        linha &= "|N" '31 IND_PR_TRANSF - Felipe 15/07/2024
        linha &= "|"

        Strm.WriteLine(linha)
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistro0030(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************  REGISTRO 0030: DADOS CADASTRAIS   ********************************
        '******************************************************************************************************
        linha = "|0030"                                           '01 - REG
        linha &= "|" & Empresa.Empresa.CodigoNaturezaJuridica     '02 - COD_NAT
        linha &= "|" & Empresa.Empresa.CodigoCnae                 '03 - CNAE_FISCAL
        linha &= "|" & Empresa.Endereco                           '04 - ENDERECO
        linha &= "|" & Empresa.Numero                             '05 - NUM
        linha &= "|" & Empresa.Complemento                        '06 - COMPL
        linha &= "|" & Empresa.Bairro                             '07 - BAIRRO
        linha &= "|" & Empresa.CodigoEstado                       '08 - UF
        linha &= "|" & Empresa.Municipio.CodigoIbgeCompleto       '09 - COD_MUN
        linha &= "|" & Empresa.CEP                                '10 - CEP
        linha &= "|" & Empresa.Telefone                           '11 - NUM_TEL
        linha &= "|" & Empresa.Email                              '12 - EMAIL
        linha &= "|"

        Strm.WriteLine(linha)
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistro0930(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '********************* REGISTRO 0930: IDENTIFICAÇÃO DOS SIGNATÁRIOS DA ECF   **************************
        '******************************************************************************************************
        linha = "|0930"                                          '01 - REG
        linha &= "|" & Empresa.Empresa.NomeDoTitular             '02 - IDENT_NOM
        linha &= "|" & Empresa.Empresa.CPFTitular                '03 - IDENT_CPF_CNPJ
        linha &= "|" & Empresa.Empresa.CodigoQualificacaoTitular '04 - IDENT_QUALIF
        linha &= "|"                                             '05 - IND_CRC
        linha &= "|" & Empresa.Empresa.EmailContador             '06 - EMAIL
        linha &= "|" & Empresa.Empresa.TelefoneContador          '07 - FONE
        linha &= "|"

        Strm.WriteLine(linha)

        linha = "|0930"                                           '01 - REG
        linha &= "|" & Empresa.Empresa.NomeDoContador             '02 - IDENT_NOM
        linha &= "|" & Empresa.Empresa.CPFContador                '03 - IDENT_CPF_CNPJ
        linha &= "|" & Empresa.Empresa.CodigoQualificacaoContador '04 - IDENT_QUALIF
        linha &= "|" & Empresa.Empresa.CRCContador                '05 - IND_CRC
        linha &= "|" & Empresa.Empresa.EmailContador              '06 - EMAIL
        linha &= "|" & Empresa.Empresa.TelefoneContador           '07 - FONE
        linha &= "|"

        Strm.WriteLine(linha)
        RegistroGeral += 2
    End Sub

    Private Sub CompoeRegistro0990(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************  REGISTRO 0990: ENCERRAMENTO DO BLOCO 0   ********************************
        '******************************************************************************************************

        RegistroGeral += 1 'Deve contar a si mesmo

        linha = "|0990"                                           '01 - REG
        linha &= "|" & RegistroGeral & "|"                        '02 - QTD_LIN

        Strm.WriteLine(linha)
    End Sub
#End Region

#Region "Registro J"
    Private Sub CompoeRegistroJ001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroJ As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO J001: ABERTURA DO BLOCO J   ************************************
        '******************************************************************************************************

        linha = "|J001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroJ += 1
        RegistroGeral += 1
    End Sub

    Private Function ConsultaRegistroJ050_J051(ByRef Empresa As Cliente) As DataSet
        Dim Sql As String

        Sql = "SELECT DISTINCT Consulta1.Conta_Id," & vbCrLf &
                      "       Consulta1.Linha as J050," & vbCrLf &
                      "	      Consulta1.Referencial as J051" & vbCrLf &
                      "  FROM (" & vbCrLf &
                      "        SELECT Consulta1.Conta_Id ,Consulta1.Linha, Consulta1.Referencial" & vbCrLf &
                      "          FROM (SELECT PC.conta_Id," & vbCrLf &
                      "					      --CASE" & vbCrLf &
                      "		                  --     WHEN PCB.Conta_Id IS NOT NULL AND ISNULL(R.Cliente_ID,'') <> '' " & vbCrLf &
                      " 			          --         THEN R.Conta_Id + R.Cliente_Id +  CONVERT(NVARCHAR,R.EndCliente_Id)" & vbCrLf &
                      "			              --         ELSE PC.Conta_ID" & vbCrLf &
                      "                       --END AS Conta_Id," & vbCrLf &
                      "		                  CASE" & vbCrLf &
                      "					          WHEN PCB.Conta_Id IS NOT NULL" & vbCrLf &
                      "					              THEN 'A'" & vbCrLf &
                      "					              ELSE 'S'" & vbCrLf &
                      "					      END AS Tipo," & vbCrLf &
                      "                       '|J050|01012014|'+ " & vbCrLf &
                      "                       /* COD_NAT Código da Natureza da Conta Analítica ou Sintética:*/ " & vbCrLf &
                      "                       CASE " & vbCrLf &
                      "                             WHEN LEFT(PC.conta_Id,1)  = '1' and not LEFT(PC.conta_Id,3) in ('105')        THEN '01'" & vbCrLf &
                      "	                            WHEN LEFT(PC.conta_Id,1)  = '2' and not LEFT(PC.conta_Id,3) in ('204','205')  THEN '02'" & vbCrLf &
                      "                             WHEN LEFT(PC.conta_Id,3)              = '105' THEN '05'" & vbCrLf &
                      "                             WHEN LEFT(PC.conta_Id,3)              = '205' THEN '05'" & vbCrLf &
                      "                             WHEN LEFT(PC.conta_Id,3)              = '204' THEN '03'" & vbCrLf &
                      "                         ELSE '04'" & vbCrLf &
                      "			              END " & vbCrLf &
                      "                       +'|'+ " & vbCrLf &
                      "                       /* IND_CTA Indicador do Tipo de Conta: */" & vbCrLf &
                      "		                  CASE" & vbCrLf &
                      "					          WHEN PCB.Conta_Id IS NOT NULL" & vbCrLf &
                      "					              THEN 'A'" & vbCrLf &
                      "					              ELSE 'S'" & vbCrLf &
                      "					      END " & vbCrLf &
                      "                       +'|'+ " & vbCrLf &
                      "                       /* NIVEL Nível da Conta Analítica/Sintética: Número crescente a partir da conta de menor detalhamento. Deve ser acrescido de 1 a cada mudança de nível. */ " & vbCrLf &
                      "						  CASE WHEN LEN(PC.Conta_Id) = 1 THEN '1'" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) = 3 THEN '2'" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) = 5 THEN '3'" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) = 7 THEN '4'" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) > 7 THEN '5'" & vbCrLf &
                      "                       END " & vbCrLf &
                      "                       +'|'+" & vbCrLf &
                      "                       PC.Conta_Id " & vbCrLf &
                      "					      --CASE" & vbCrLf &
                      "		                  --     WHEN PCB.Conta_Id IS NOT NULL AND ISNULL(R.Cliente_ID,'') <> '' " & vbCrLf &
                      " 			          --         THEN R.Conta_Id + R.Cliente_Id +  CONVERT(NVARCHAR,R.EndCliente_Id)" & vbCrLf &
                      "			              --         ELSE PC.Conta_ID" & vbCrLf &
                      "                       --END" & vbCrLf &
                      "                       +'|'+" & vbCrLf &
                      "                      /* COD_CTA_SUP Código da Conta Sintética de Nível Imediatamente Superior. */" & vbCrLf &
                      "						  CASE WHEN LEN(PC.Conta_Id) = 1 THEN ''" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) = 3 THEN SUBSTRING(PC.Conta_Id,1,1)" & vbCrLf &
                      "							   WHEN LEN(PC.Conta_Id) = 5 THEN SUBSTRING(PC.Conta_Id,1,3)" & vbCrLf &
                      "							   WHEN LEN(PC.Conta_Id) = 7 THEN SUBSTRING(PC.Conta_Id,1,5)" & vbCrLf &
                      "                            WHEN LEN(PC.Conta_Id) > 7 THEN SUBSTRING(PC.Conta_Id,1,7)" & vbCrLf &
                      "                       END  +'|'+" & vbCrLf &
                      "                         RTrim(LTrim(PC.Titulo))" & vbCrLf &
                      "                       --CASE" & vbCrLf &
                      "                       --    WHEN PCB.Conta_Id IS NOT NULL AND ISNULL(R.Cliente_ID,'') <> ''" & vbCrLf &
                      "                       --         THEN RTRIM(LTRIM(ISNULL(Cli.Nome, PC.Titulo)))" & vbCrLf &
                      "                       --         ELSE RTRIM(LTRIM(PC.Titulo)) " & vbCrLf &
                      "                       --END " & vbCrLf &
                      "                           +'|'" & vbCrLf &
                      "                       AS Linha," & vbCrLf &
                      "                       CASE " & vbCrLf &
                      "					          WHEN PCB.Conta_Id IS NULL" & vbCrLf &
                      "						          THEN ''" & vbCrLf &
                      "						          ELSE '|J051||'+ RTRIM(LTRIM(PCB.Conta_Id))+'|'" & vbCrLf &
                      "					      END AS Referencial" & vbCrLf &
                      "                  FROM PlanoDeContas PC" & vbCrLf &
                      "                  LEFT JOIN Razao R" & vbCrLf &
                      "                    ON PC.Conta_id = R.Conta_id " & vbCrLf &
                      "                  LEFT JOIN Clientes Cli" & vbCrLf &
                      "                    ON Cli.Cliente_ID = R.Cliente_ID" & vbCrLf &
                      "                   AND Cli.Endereco_Id = R.EndCliente_Id" & vbCrLf &
                      "                  LEFT JOIN PlanoDeContasXReferencialBacen PlxB " & vbCrLf &
                      "                    ON PC.Empresa_Id = PlxB.Empresa " & vbCrLf &
                      "                   AND PC.EndEmpresa_Id = Plxb.EndEmpresa" & vbCrLf &
                      "                   AND PC.Conta_Id = Plxb.Conta" & vbCrLf &
                      "                  LEFT JOIN PlanoDeContasReferencialBacen PCB " & vbCrLf &
                      "                    ON Plxb.Referencial  = PCB.Conta_Id " & vbCrLf &
                      "		            WHERE len(PC.Conta_Id)  <= 7" & vbCrLf &
                      "			           OR (len(PC.Conta_Id) > 7" & vbCrLf &
                      "					          AND EXISTS(SELECT 1" & vbCrLf &
                      "		                                   FROM Razao RZ" & vbCrLf &
                      "                                       WHERE RZ.Conta_Id = PC.Conta_Id" & vbCrLf &
                      "                                     )" & vbCrLf &
                      "				           )" & vbCrLf &
                      "                ) AS Consulta1" & vbCrLf &
                      "       ) AS Consulta1" & vbCrLf &
                      " WHERE Conta_ID IS NOT NULL " & vbCrLf &
                      "    AND (LEN(Consulta1.Referencial) > 0" & vbCrLf &
                      "          OR LEN(Conta_Id)  <= 9 )" & vbCrLf &
                      " ORDER BY Consulta1.Conta_Id ASC"

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroJ050_J051(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef linha As String, ByRef RegistroJ As Integer, ByRef RegistroGeral As Integer, ByRef PeriodoFinal As Date)
        '******************************************************************************************************
        '*******************    REGISTRO J050: PLANO DE CONTAS DO CONTRIBUINTE     ****************************
        '*******************    REGISTRO J051: PLANO DE CONTAS REFERENCIAL         ****************************
        '******************************************************************************************************


        'Dim ds As DataSet = ConsultaRegistroJ050_J051(Empresa)

        'Dim COD_CTA As String = String.Empty

        'For Each row In ds.Tables(0).Rows
        '    If COD_CTA <> row("Conta_Id") Then
        '        Strm.WriteLine(row("J050"))
        '        RegistroJ += 1
        '        RegistroGeral += 1

        '        If row("J051").ToString.Length > 0 Then
        '            Strm.WriteLine(row("J051"))
        '            RegistroJ += 1
        '            RegistroGeral += 1
        '        End If
        '        COD_CTA = row("Conta_Id")
        '    Else
        '        If row("J051").ToString.Length > 0 Then
        '            Strm.WriteLine(row("J051"))
        '            RegistroJ += 1
        '            RegistroGeral += 1
        '        End If
        '    End If
        'Next


        Dim Sql As String

        'Sql = " SELECT  Razao.Conta_Id + Razao.Cliente_Id +  convert(nvarchar,Razao.EndCliente_Id) AS Conta_Id, isnull(Clientes.Nome, 'Cliente Nao Encontrado') AS Titulo, Razao.Produto" & vbCrLf &
        Sql = " SELECT distinct Razao.Conta_Id + Razao.Cliente_Id +  convert(nvarchar,Razao.EndCliente_Id) AS Conta_Id, isnull(Clientes.Nome, 'Cliente Nao Encontrado') AS Titulo, PlanoDeContas.Produto" & vbCrLf &
                  "     FROM       Razao " & vbCrLf &
                  "            LEFT OUTER JOIN Clientes " & vbCrLf &
                  "                         ON Razao.Cliente_Id = Clientes.Cliente_Id" & vbCrLf &
                  "                         And Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
                  "            LEFT OUTER JOIN PlanoDeContas " & vbCrLf &
                  "                         ON PlanoDeContas.Conta_Id = Razao.Conta_Id " & vbCrLf &
                  "     WHERE      (Razao.Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%') And Razao.Cliente_ID <> '')" & vbCrLf &
                  "            And (Razao.Movimento_Id <= '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')" & vbCrLf &
                  " Group by Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, Clientes.Nome, PlanoDeContas.Produto" & vbCrLf &
                  "                                                            " & vbCrLf &
                  " Union" & vbCrLf &
                  "                                                             " & vbCrLf &
                  " SELECT Conta_Id, Titulo, Produto " & vbCrLf &
                  "     FROM   PlanoDeContas" & vbCrLf &
                  "     WHERE     Len(Conta_Id) < 9" & vbCrLf &
                  "       AND left(Conta_Id,1) not in('5')" & vbCrLf &
                  "                                                     " & vbCrLf &
                  " Union" & vbCrLf &
                  "                                                     " & vbCrLf &
                  " SELECT  PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PlanoDeContas.Produto  " & vbCrLf &
                  "     From Razao " & vbCrLf &
                  "         LEFT OUTER JOIN PlanoDeContas" & vbCrLf &
                  "             ON PlanoDeContas.Conta_Id = Razao.Conta_Id " & vbCrLf &
                  "     WHERE    (LEN(PlanoDeContas.Conta_Id) = 9) And Razao.Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%') AND left(Razao.Conta_Id,1) not in('5')" & vbCrLf &
                  " Group by PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PlanoDeContas.Produto " & vbCrLf &
                  " Order by Conta_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "PlanoDeContas")

        Dim Nivel1 As String = ""
        Dim Nivel2 As String = ""
        Dim Conta As String = ""

        If ds.Tables(0).Rows.Count > 0 Then
            For Each dr In ds.Tables(0).Rows
                If Conta = dr("Conta_Id") Then
                Else
                    With dr
                        linha = "|J050"
                        linha &= "|0101" & PeriodoFinal.Year.ToString()

                        '****Plano Novo
                        Nivel1 = Left(.Item("Conta_Id"), 1)
                        Nivel2 = Left(.Item("Conta_Id"), 3)
                        If Nivel2 = "204" Then
                            linha &= "|03"
                        ElseIf Nivel2 = "105" Then
                            linha &= "|05"
                        ElseIf Nivel2 = "205" Then
                            linha &= "|05"
                        Else
                            If Nivel1 = "1" Then
                                linha &= "|01"
                            ElseIf Nivel1 = "2" Then
                                linha &= "|02"
                            ElseIf Nivel1 = "3" Then
                                linha &= "|04"
                            ElseIf Nivel1 = "4" Then
                                linha &= "|04"
                            ElseIf Nivel1 = "5" Then
                                linha &= "|04"
                            End If
                        End If

                        If Len(.Item("Conta_Id")) > 7 Then
                            linha &= "|A"
                        Else
                            linha &= "|S"
                        End If

                        If Len(.Item("Conta_Id")) > 7 Then
                            linha &= "|5"
                        ElseIf Len(.Item("Conta_Id")) = 7 Then
                            linha &= "|4"
                        ElseIf Len(.Item("Conta_Id")) = 5 Then
                            linha &= "|3"
                        ElseIf Len(.Item("Conta_Id")) = 3 Then
                            linha &= "|2"
                        ElseIf Len(.Item("Conta_Id")) = 1 Then
                            linha &= "|1"
                        End If

                        linha &= "|" & .Item("Conta_Id")
                        Conta = .Item("Conta_Id")

                        If Len(.Item("Conta_Id")) > 7 Then
                            linha &= "|" & Left(.Item("Conta_Id"), 7)
                        ElseIf Len(.Item("Conta_Id")) = 7 Then
                            linha &= "|" & Left(.Item("Conta_Id"), 5)
                        ElseIf Len(.Item("Conta_Id")) = 5 Then
                            linha &= "|" & Left(.Item("Conta_Id"), 3)
                        ElseIf Len(.Item("Conta_Id")) = 3 Then
                            linha &= "|" & Left(.Item("Conta_Id"), 1)
                        ElseIf Len(.Item("Conta_Id")) = 1 Then
                            linha &= "|"
                        End If


                        linha &= "|" & RTrim(.Item("Titulo"))
                        linha &= "|"

                        Strm.WriteLine(linha)

                        RegistroJ += 1
                        RegistroGeral += 1

                        '-------Plano de Contas Referencial-------

                        If Len(.Item("Conta_Id")) > 7 Then
                            If Len(.Item("Conta_Id")) > 7 Then
                                Conta = Left(.Item("Conta_Id"), 9)
                            End If

                            If Not IsDBNull(.Item("Produto")) AndAlso .Item("Produto") = "S" AndAlso (Left(.Item("Conta_Id"), 1) = "1" Or Left(.Item("Conta_Id"), 1) = "2") Then
                                Sql = "SELECT  RB.Referencial " & vbCrLf &
                                        "From Razao R " & vbCrLf &
                                        "	INNER JOIN PlanoDeContasXReferencialBacen RB" & vbCrLf &
                                        "		ON RB.Conta = '" & Left(Conta, 7) & "'" & vbCrLf &
                                        "		AND RB.Grupo = Left(R.Produto, 5) " & vbCrLf &
                                        "WHERE R.Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf &
                                        "AND R.conta_ID like '" & .Item("Conta_Id") & "'" & vbCrLf &
                                        "And R.Movimento_Id <= '" & PeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                                        "GROUP BY RB.Referencial "
                            Else
                                Sql = "  SELECT   Referencial "
                                Sql &= " FROM     PlanoDeContasXReferencialBacen"
                                Sql &= " WHERE    (Conta = '" & Conta & "' OR Conta = '" & Left(Conta, 7) & "') Group by Referencial"
                            End If

                            Dim dsr As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                            If dsr.Tables(0).Rows.Count > 0 Then
                                For Each drr In dsr.Tables(0).Rows
                                    With drr
                                        linha = "|J051"
                                        linha &= "|"
                                        linha &= "|" & .Item("Referencial")
                                        linha &= "|"
                                    End With

                                    Strm.WriteLine(linha)

                                    RegistroJ += 1
                                    RegistroGeral += 1
                                Next
                            End If

                        End If
                    End With
                End If
            Next
        End If
    End Sub

    Private Sub CompoeRegistroJ990(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroJ As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************  REGISTRO J990: ENCERRAMENTO DO BLOCO J   ********************************
        '******************************************************************************************************

        RegistroGeral += 1 'Deve contar a si mesmo
        RegistroJ += 1 'Deve contar a si mesmo

        linha = "|J990"                                           '01 - REG
        linha &= "|" & RegistroJ                              '02 - QTD_LIN
        linha &= "|"


        Strm.WriteLine(linha)
    End Sub
#End Region

#Region "Registro K"
    Private Sub CompoeRegistroK001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroK As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO K001: ABERTURA DO BLOCO K   ************************************
        '******************************************************************************************************

        linha = "|K001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroK += 1
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistroK030(ByRef Strm As StreamWriter, ByRef Trimestre As String, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date, _
                                   ByRef linha As String, ByRef RegistroK As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  REGISTRO K030: IDENTIFICAÇÃO DO PERÍODO E FORMAS DE APURAÇÃO DO IRPJ E DA CSLL NO ANO-CALENDÁRIO  ******
        '***************************************************************************************************************

        linha = "|K030"                                            '01 - REG      
        linha &= "|" & TrimestrePeriodoInicial.ToString("ddMMyyyy")  '02 - DT_INI
        linha &= "|" & TrimestrePeriodoFinal.ToString("ddMMyyyy")  '03 - DT_FIN
        linha &= "|" & Trimestre & "|"                                            '04 - PER_APUR   


        Strm.WriteLine(linha)
        RegistroK += 1
        RegistroGeral += 1
    End Sub

    Private Function ConsultaRegistroK155(ByRef Empresa As Cliente, ByRef PeriodoInicial As Date, ByRef PeriodoFinal As Date) As DataSet
        Dim Sql As String

        'Sql = "SELECT Linha , Conta_id" & vbCrLf &
        '      "  FROM(" & vbCrLf &
        '      "		  SELECT  Consulta.Empresa_Id,LEFT(Consulta.Conta_Id,9) AS Conta_id" & vbCrLf &
        '      "				  ,'|K155|'+LEFT(Conta_Id,9)" & vbCrLf &
        '      "				 +'||'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Inicial)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ CASE WHEN SUM(Inicial) > 0 then 'D' else 'C' end" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Debitos)) ,'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,(Sum(Inicial) + Sum(Debitos)) - Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ CASE WHEN ((Sum(Inicial)+Sum(Debitos)) - Sum(Creditos)) > 0 then 'D' else 'C' end" & vbCrLf &
        '      "				 +'|' as Linha" & vbCrLf &
        '      "				 ,(SELECT TOP 1 [Referencial]" & vbCrLf &
        '      "					 FROM [PlanoDeContasXReferencialBacen] RB" & vbCrLf &
        '      "					WHERE (RB.Empresa = LEFT(Consulta.Empresa_Id,8) OR RB.Empresa = '99999999999999')" & vbCrLf &
        '      "					  AND RB.Conta = LEFT(Consulta.Conta_Id,9)" & vbCrLf &
        '      "                   And (Sum(Inicial)   <> 0 or Sum(Debitos) <> 0 or Sum(Creditos) <> 0)" & vbCrLf &
        '      "                     ) AS Referencial" & vbCrLf &
        '      "		  FROM" & vbCrLf &
        '      "		  (" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				  FROM" & vbCrLf &
        '      "					 (SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							,EndEmpresa_Id" & vbCrLf &
        '      "							,Left(Conta_Id, 7) AS Conta_id" & vbCrLf &
        '      "							,SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
        '      "							,0 as Debitos" & vbCrLf &
        '      "							,0 AS Creditos" & vbCrLf &
        '      "						FROM Razao" & vbCrLf &
        '      "					   WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "						 AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "						 And Movimento_Id < '" & PeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '      "						 AND Lote_Id < 9000" & vbCrLf &
        '      "					GROUP BY Left(Empresa_Id, 8),EndEmpresa_Id, Left(Conta_Id, 7)" & vbCrLf &
        '      "					 ) AS Consulta" & vbCrLf &
        '      "				UNION" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				 FROM (" & vbCrLf &
        '      "					   SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							 ,EndEmpresa_Id" & vbCrLf &
        '      "							 ,LEFT(Conta_id,9) AS Conta_Id" & vbCrLf &
        '      "							 ,SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
        '      "							 ,0 as Debitos" & vbCrLf &
        '      "							 ,0 AS Creditos" & vbCrLf &
        '      "						 FROM Razao" & vbCrLf &
        '      "						WHERE  LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "						  AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "						 And Movimento_Id < '" & PeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '      "						  AND Lote_Id < 9000" & vbCrLf &
        '      "					 GROUP BY Left(Empresa_Id, 8), EndEmpresa_Id , LEFT(Conta_id,9)" & vbCrLf &
        '      "					  ) AS Consulta" & vbCrLf &
        '      "				 UNION" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				  FROM (" & vbCrLf &
        '      "						 SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							   ,EndEmpresa_Id" & vbCrLf &
        '      "							   ,Left(Conta_Id, 7) AS Conta_id" & vbCrLf &
        '      "							   ,0 as Inicial" & vbCrLf &
        '      "							   ,SUM(DebitoOficial) as Debitos" & vbCrLf &
        '      "							   ,SUM(CreditoOficial) AS Creditos" & vbCrLf &
        '      "						   FROM Razao" & vbCrLf &
        '      "						  WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "						    AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "							AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')                             " & vbCrLf &
        '      "							AND Lote_Id < 9000" & vbCrLf &
        '      "						  GROUP BY Left(Empresa_Id, 8), EndEmpresa_Id, Left(Conta_Id, 7)" & vbCrLf &
        '      "						) AS Consulta" & vbCrLf &
        '      " 				 UNION" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id,  LEFT(Consulta.Conta_id,9), Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				  FROM (" & vbCrLf &
        '      "						 SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							   ,EndEmpresa_Id" & vbCrLf &
        '      "							   ,LEFT(Conta_id,9) as Conta_id" & vbCrLf &
        '      "							   ,0 as Inicial" & vbCrLf &
        '      "							   ,SUM(DebitoOficial) as Debitos" & vbCrLf &
        '      "							   ,SUM(CreditoOficial) AS Creditos" & vbCrLf &
        '      "						   FROM Razao" & vbCrLf &
        '      "						  WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "							  AND Empresa_Id  LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "							  AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')                           " & vbCrLf &
        '      "							  AND Lote_Id < 9000" & vbCrLf &
        '      "					   GROUP BY Left(Empresa_Id, 8), EndEmpresa_Id, LEFT(Conta_id,9)" & vbCrLf &
        '      "					   ) AS Consulta" & vbCrLf &
        '      "			 ) AS Consulta" & vbCrLf &
        '      " 			   GROUP BY LEFT(Conta_id,9), Empresa_Id" & vbCrLf &
        '      "     )AS Consulta" & vbCrLf &
        '      "     WHERE LEN(Consulta.Referencial) > 1" & vbCrLf &
        '      "      ORDER BY Conta_id " & vbCrLf


        Sql = "SELECT Linha , Conta_id" & vbCrLf &
              "  FROM(" & vbCrLf &
              "		  SELECT  Consulta.Empresa_Id, Consulta.Conta_Id" & vbCrLf &
              "				  ,'|K155|'+replace(Conta_Id,';','')" & vbCrLf &
              "				 +'||'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Inicial)),'.',','),'-','')" & vbCrLf &
              "				 +'|'+ CASE WHEN SUM(Inicial) > 0 then 'D' else 'C' end" & vbCrLf &
              "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Debitos)) ,'.',','),'-','')" & vbCrLf &
              "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Creditos)),'.',','),'-','')" & vbCrLf &
              "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,(Sum(Inicial) + Sum(Debitos)) - Sum(Creditos)),'.',','),'-','')" & vbCrLf &
              "				 +'|'+ CASE WHEN ((Sum(Inicial)+Sum(Debitos)) - Sum(Creditos)) > 0 then 'D' else 'C' end" & vbCrLf &
              "				 +'|' as Linha" & vbCrLf &
              "" & vbCrLf &
              "" & vbCrLf &
              "		  FROM" & vbCrLf &
              "		  (" & vbCrLf &
              "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
              "				  FROM" & vbCrLf &
              "					 (SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
              "							,EndEmpresa_Id" & vbCrLf &
              "							,Case" & vbCrLf &
              "								When len(isnull(Cliente_id,0)) > 0" & vbCrLf &
              "									then Conta_id + ';' + Cliente_id + ';' + CONVERT(VARCHAR,EndCliente_Id)" & vbCrLf &
              "									else Conta_id" & vbCrLf &
              "							end as Conta_id" & vbCrLf &
              "							,SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
              "							,0 as Debitos" & vbCrLf &
              "							,0 AS Creditos" & vbCrLf &
              "						FROM Razao" & vbCrLf &
              "					   WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
              "						 AND (left(Conta_Id,3) <> '105' and left(Conta_Id,3) <> '205')" & vbCrLf &
              "						 AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
              "						 And Movimento_Id < '" & PeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
              "						 AND Lote_Id < 9000" & vbCrLf &
              "					   GROUP BY Left(Empresa_Id, 8),EndEmpresa_Id, Conta_Id, Cliente_id, EndCliente_Id" & vbCrLf &
              "					   HAVING SUM(DebitoOficial - CreditoOficial) <> 0" & vbCrLf &
              "					 ) AS Consulta" & vbCrLf &
              "				UNION" & vbCrLf &
              "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
              "				  FROM (" & vbCrLf &
              "						 SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
              "							   ,EndEmpresa_Id" & vbCrLf &
              "					 		    ,Case" & vbCrLf &
              "								    When len(isnull(Cliente_id,0)) > 0" & vbCrLf &
              "									    then Conta_id + ';' + Cliente_id + ';' + CONVERT(VARCHAR,EndCliente_Id)" & vbCrLf &
              "									    else Conta_id" & vbCrLf &
              "							    end as Conta_id" & vbCrLf &
              "							   ,0 as Inicial" & vbCrLf &
              "							   ,SUM(DebitoOficial) as Debitos" & vbCrLf &
              "							   ,SUM(CreditoOficial) AS Creditos" & vbCrLf &
              "						   FROM Razao" & vbCrLf &
              "						  WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
              "					 	    AND (left(Conta_Id,3) <> '105' and left(Conta_Id,3) <> '205')" & vbCrLf &
              "						    AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
              "							AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')" & vbCrLf &
              "							AND Lote_Id < 9000" & vbCrLf &
              "						  GROUP BY Left(Empresa_Id, 8), EndEmpresa_Id, Conta_id, Cliente_id, EndCliente_Id" & vbCrLf &
              "						) AS Consulta" & vbCrLf &
              "			 ) AS Consulta" & vbCrLf &
              " 			   GROUP BY Conta_id, Empresa_Id" & vbCrLf &
              "     )AS Consulta" & vbCrLf &
              "      ORDER BY Conta_id " & vbCrLf


        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function ConsultaRegistroK156(ByRef Empresa As Cliente, ByRef Conta As String, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date) As DataSet
        Dim Sql As String
        'Sql = "SELECT Linha FROM ( SELECT '|K156|'+LTRIM(RTRIM([Referencial]))" & vbCrLf & _
        '      "      +'|'+REPLACE(CONVERT(VARCHAR,SUM(CASE WHEN SaldoFinal < 0 THEN  SaldoFinal*-1 ELSE SaldoFinal END)) ,'.',',')" & vbCrLf & _
        '      "      +'|'+CASE WHEN SUM(SaldoFinal) > 0 THEN 'D' ELSE 'C' END +'|' AS Linha" & vbCrLf & _
        '      "FROM (" & vbCrLf & _
        '      "         SELECT Referencial, sum(SaldoFinal) as SaldoFinal  From (" & vbCrLf & _
        '      "               SELECT RB.[Empresa]" & vbCrLf & _
        '      " 					  ,RB.[Conta]" & vbCrLf & _
        '      " 					  ,RB.[Referencial]" & vbCrLf & _
        '      " 					  ,RB.[Grupo]" & vbCrLf & _
        '      " 					  ,RB.[Produto]" & vbCrLf & _
        '      " 					  ,CASE WHEN LEN(RB.[Grupo]) > 0" & vbCrLf & _
        '      " 					        THEN" & vbCrLf & _
        '      " 							  ( SELECT SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS SaldoFinal" & vbCrLf & _
        '      "									  FROM Razao" & vbCrLf & _
        '      "									 WHERE Left(Razao.Empresa_Id, 8)  = " & Empresa.Codigo.Substring(0, 8) & vbCrLf & _
        '      "									   AND (Razao.Movimento_Id <= '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "')" & vbCrLf & _
        '      "									   AND left(Razao.Conta_Id,LEN(RB.Conta)) = RB.Conta" & vbCrLf & _
        '      "									   AND Left(Razao.Produto, 5) =  RB.Grupo" & vbCrLf & _
        '      "							  )" & vbCrLf & _
        '      "							ELSE" & vbCrLf & _
        '      "								  ( SELECT SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS SaldoFinal" & vbCrLf & _
        '      "									  FROM Razao" & vbCrLf & _
        '      "									 WHERE Left(Razao.Empresa_Id, 8)  = " & Empresa.Codigo.Substring(0, 8) & vbCrLf & _
        '      "									   AND (Razao.Movimento_Id <= '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "')                                     " & vbCrLf & _
        '      "									   AND left(Razao.Conta_Id,LEN(RB.Conta)) = RB.Conta" & vbCrLf & _
        '      "							  )" & vbCrLf & _
        '      "							END AS SaldoFinal" & vbCrLf & _
        '      "				  FROM [PlanoDeContasXReferencialBacen] RB" & vbCrLf & _
        '      "				 WHERE (RB.Empresa LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%') OR RB.Empresa = '99999999999999')" & vbCrLf & _
        '      "				   AND RB.Conta = '" & Conta & "') AS Resumo " & vbCrLf & _
        '      "         group by Referencial" & vbCrLf & _
        '      ") AS C1" & vbCrLf & _
        '      "GROUP BY [Referencial]   ) As C1 WHERE LEN(Linha) > 0" & vbCrLf

        'Dim PeriodoAnterior = TrimestrePeriodoInicial.AddDays(-1)

        'Sql = "SELECT Linha , Conta_id" & vbCrLf &
        '      "  FROM(" & vbCrLf &
        '      "		  SELECT  Consulta.Empresa_Id, Consulta.Conta_Id" & vbCrLf &
        '      "				  ,'|K155|'+Conta_Id" & vbCrLf &
        '      "				 +'||'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Inicial)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ CASE WHEN SUM(Inicial) > 0 then 'D' else 'C' end" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Debitos)) ,'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,(Sum(Inicial) + Sum(Debitos)) - Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '      "				 +'|'+ CASE WHEN ((Sum(Inicial)+Sum(Debitos)) - Sum(Creditos)) > 0 then 'D' else 'C' end" & vbCrLf &
        '      "				 +'|' as Linha" & vbCrLf &
        '      "				 ,(SELECT TOP 1 [Referencial]" & vbCrLf &
        '      "					 FROM [PlanoDeContasXReferencialBacen] RB" & vbCrLf &
        '      "					WHERE (RB.Empresa = LEFT(Consulta.Empresa_Id,8) OR RB.Empresa = '99999999999999')" & vbCrLf &
        '      "					  AND (RB.Conta = LEFT(Consulta.Conta_Id,9) OR RB.Conta = LEFT(Consulta.Conta_Id,7))" & vbCrLf &
        '      "" & vbCrLf &
        '      "" & vbCrLf &
        '      "		  FROM" & vbCrLf &
        '      "		  (" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				  FROM" & vbCrLf &
        '      "					 (SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							,EndEmpresa_Id" & vbCrLf &
        '      "							,Case" & vbCrLf &
        '      "								When len(isnull(Cliente_id,0)) > 0" & vbCrLf &
        '      "									then Conta_id + Cliente_id + CONVERT(VARCHAR,EndCliente_Id)" & vbCrLf &
        '      "									else Conta_id" & vbCrLf &
        '      "							end as Conta_id" & vbCrLf &
        '      "							,SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
        '      "							,0 as Debitos" & vbCrLf &
        '      "							,0 AS Creditos" & vbCrLf &
        '      "						FROM Razao" & vbCrLf &
        '      "					   WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "						 AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "						 And Movimento_Id < '" & PeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '      "						 AND Lote_Id < 9000" & vbCrLf &
        '      "					   GROUP BY Left(Empresa_Id, 8),EndEmpresa_Id, Conta_Id, Cliente_id, EndCliente_Id" & vbCrLf &
        '      "					   HAVING SUM(DebitoOficial - CreditoOficial) <> 0" & vbCrLf &
        '      "					 ) AS Consulta" & vbCrLf &
        '      "				UNION" & vbCrLf &
        '      "				SELECT Consulta.Empresa_Id, Consulta.Conta_id, Consulta.Inicial, Consulta.Debitos, Consulta.Creditos" & vbCrLf &
        '      "				  FROM (" & vbCrLf &
        '      "						 SELECT Left(Empresa_Id, 8) as Empresa_Id" & vbCrLf &
        '      "							   ,EndEmpresa_Id" & vbCrLf &
        '      "					 		    ,Case" & vbCrLf &
        '      "								    When len(isnull(Cliente_id,0)) > 0" & vbCrLf &
        '      "									    then Conta_id + Cliente_id + CONVERT(VARCHAR,EndCliente_Id)" & vbCrLf &
        '      "									    else Conta_id" & vbCrLf &
        '      "							    end as Conta_id" & vbCrLf &
        '      "							   ,0 as Inicial" & vbCrLf &
        '      "							   ,SUM(DebitoOficial) as Debitos" & vbCrLf &
        '      "							   ,SUM(CreditoOficial) AS Creditos" & vbCrLf &
        '      "						   FROM Razao" & vbCrLf &
        '      "						  WHERE LEFT(Conta_Id, 1) IN ('1', '2')" & vbCrLf &
        '      "						    AND Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
        '      "							AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')" & vbCrLf &
        '      "							AND Lote_Id < 9000" & vbCrLf &
        '      "						  GROUP BY Left(Empresa_Id, 8), EndEmpresa_Id, Conta_id, Cliente_id, EndCliente_Id" & vbCrLf &
        '      "						) AS Consulta" & vbCrLf &
        '      "			 ) AS Consulta" & vbCrLf &
        '      " 			   GROUP BY Conta_id, Empresa_Id" & vbCrLf &
        '      "     )AS Consulta" & vbCrLf &
        '      "      ORDER BY Conta_id " & vbCrLf


        'Sql = "SELECT " & vbCrLf &
        '        "'|K156|'+Referencial " & vbCrLf &
        '        "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Inicial)),'.',','),'-','')" & vbCrLf &
        '        "+'|'+ CASE WHEN SUM(Inicial) > 0 then 'D' else 'C' end" & vbCrLf &
        '        "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Debitos)) ,'.',','),'-','')" & vbCrLf &
        '        "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '        "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,(Sum(Inicial) + Sum(Debitos)) - Sum(Creditos)),'.',','),'-','')" & vbCrLf &
        '        "+'|'+ CASE WHEN ((Sum(Inicial)+Sum(Debitos)) - Sum(Creditos)) > 0 then 'D' else 'C' end" & vbCrLf &
        '        "+'|' as Linha," & vbCrLf &
        '        "Referencial, Sum(Inicial) AS Inicial, Sum(Debitos) AS Debitos, Sum(Creditos) AS Creditos" & vbCrLf &
        '        "FROM (" & vbCrLf &
        '        "		SELECT RB.Referencial," & vbCrLf &
        '        "				SUM(R.DebitoOficial - R.CreditoOficial) AS Inicial," & vbCrLf &
        '        "				0 as Debitos," & vbCrLf &
        '        "				0 AS Creditos" & vbCrLf &
        '        "		FROM Razao R" & vbCrLf &
        '        "			LEFT JOIN PlanoDeContasXReferencialBacen RB" & vbCrLf &
        '        "					ON RB.Empresa = '99999999999999'" & vbCrLf &
        '        "					AND RB.Conta = LEFT(R.Conta_Id,9) OR RB.Conta = LEFT(R.Conta_Id,7)" & vbCrLf &
        '        "					AND RB.Grupo =" & vbCrLf &
        '        "						(CASE WHEN len(isnull(R.Produto,'')) > 0" & vbCrLf &
        '        "						THEN Left(R.Produto, 5)" & vbCrLf &
        '        "						ELSE NULL" & vbCrLf &
        '        "						END)" & vbCrLf &
        '        "		WHERE R.Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf

        'If Len(Conta) > 9 Then
        '    Sql &= "		AND R.Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
        '           "		AND R.Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
        '           "		AND R.EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        'Else
        '    Sql &= "		AND R.Conta_Id = '" & Conta & "'" & vbCrLf
        'End If

        'Sql &= "		AND R.Movimento_Id < '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '        "		AND R.Lote_Id < 9000" & vbCrLf &
        '        "		GROUP BY RB.Referencial" & vbCrLf &
        '        "		HAVING SUM(R.DebitoOficial - R.CreditoOficial) <> 0" & vbCrLf &
        '        "		UNION" & vbCrLf &
        '        "		SELECT RB.Referencial," & vbCrLf &
        '        "				0 AS Inicial," & vbCrLf &
        '        "				SUM(R.DebitoOficial) AS Debitos," & vbCrLf &
        '        "				SUM(R.CreditoOficial) AS Creditos" & vbCrLf &
        '        "		FROM Razao R" & vbCrLf &
        '        "			LEFT JOIN PlanoDeContasXReferencialBacen RB" & vbCrLf &
        '        "					ON RB.Empresa = '99999999999999'" & vbCrLf &
        '        "					AND RB.Conta = LEFT(R.Conta_Id,9) OR RB.Conta = LEFT(R.Conta_Id,7)" & vbCrLf &
        '        "					AND RB.Grupo =" & vbCrLf &
        '        "						(CASE WHEN len(isnull(R.Produto,'')) > 0" & vbCrLf &
        '        "						THEN Left(R.Produto, 5)" & vbCrLf &
        '        "						ELSE NULL" & vbCrLf &
        '        "						END)" & vbCrLf &
        '        "		WHERE R.Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf

        'If Len(Conta) > 9 Then
        '    Sql &= "		AND R.Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
        '           "		AND R.Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
        '           "		AND R.EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        'Else
        '    Sql &= "		AND R.Conta_Id = '" & Conta & "'" & vbCrLf
        'End If

        'Sql &= "		AND R.Movimento_Id BETWEEN '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '        "		AND R.Lote_Id < 9000" & vbCrLf &
        '        "		GROUP BY RB.Referencial" & vbCrLf &
        '        ") AS Consulta" & vbCrLf &
        '        "GROUP BY Referencial"


        'If Len(Conta) > 9 Then
        '    Dim teste = ""
        'End If

        'If Conta = "101060151" Or Conta = "1010601" Then
        '    Dim teste2 = ""
        'End If

        Sql = "SELECT " & vbCrLf &
                "'|K156|'+Referencial " & vbCrLf &
                "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(isnull(Inicial,0))),'.',','),'-','')" & vbCrLf &
                "+'|'+ CASE WHEN SUM(isnull(Inicial,0)) > 0 then 'D' else 'C' end" & vbCrLf &
                "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(isnull(Debitos,0))) ,'.',','),'-','')" & vbCrLf &
                "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(isnull(Creditos,0))),'.',','),'-','')" & vbCrLf &
                "+'|'+ REPLACE(REPLACE(CONVERT(VARCHAR,(Sum(isnull(Inicial,0)) + Sum(isnull(Debitos,0))) - Sum(isnull(Creditos,0))),'.',','),'-','')" & vbCrLf &
                "+'|'+ CASE WHEN ((Sum(isnull(Inicial,0))+Sum(isnull(Debitos,0))) - Sum(isnull(Creditos,0))) > 0 then 'D' else 'C' end" & vbCrLf &
                "+'|' as Linha" & vbCrLf &
                "--Referencial, Sum(Inicial) AS Inicial, Sum(Debitos) AS Debitos, Sum(Creditos) AS Creditos" & vbCrLf &
                "FROM (Select RB.Referencial," & vbCrLf &
                "		CASE " & vbCrLf &
                "			WHEN LEN(RB.Grupo) > 0" & vbCrLf &
                "			THEN (SELECT SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Left(Produto, 5) = RB.Grupo" & vbCrLf &
                "					AND Movimento_Id < '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "			ELSE (SELECT SUM(DebitoOficial - CreditoOficial) AS Inicial" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Movimento_Id < '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "		END AS Inicial," & vbCrLf &
                "		CASE " & vbCrLf &
                "			WHEN LEN(RB.Grupo) > 0" & vbCrLf &
                "			THEN (SELECT SUM(DebitoOficial) AS Debitos" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Left(Produto, 5) = RB.Grupo" & vbCrLf &
                "					AND Movimento_Id BETWEEN '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "			ELSE (SELECT SUM(DebitoOficial) AS Debitos" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Movimento_Id BETWEEN '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "		END AS Debitos," & vbCrLf &
                "		CASE " & vbCrLf &
                "			WHEN LEN(RB.Grupo) > 0" & vbCrLf &
                "			THEN (SELECT SUM(CreditoOficial) AS Creditos" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Left(Produto, 5) = RB.Grupo" & vbCrLf &
                "					AND Movimento_Id BETWEEN '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "			ELSE (SELECT SUM(CreditoOficial) AS Creditos" & vbCrLf &
                "					FROM Razao " & vbCrLf &
                "					WHERE Empresa_Id LIKE '" & Empresa.Codigo.Substring(0, 8) & "%'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "					AND Conta_Id = '" & Conta.Split(";")(0) & "'" & vbCrLf &
                   "					AND Cliente_Id = '" & Conta.Split(";")(1) & "'" & vbCrLf &
                   "					AND EndCliente_Id = '" & Conta.Split(";")(2) & "'" & vbCrLf
        Else
            Sql &= "					AND Conta_Id = '" & Conta & "'" & vbCrLf
        End If
        Sql &= "					AND Movimento_Id BETWEEN '" & TrimestrePeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & TrimestrePeriodoFinal.ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "					AND Lote_Id < 9000) " & vbCrLf &
                "		END AS Creditos" & vbCrLf &
                "		FROM PlanoDeContasXReferencialBacen RB" & vbCrLf &
                "		WHERE RB.Empresa = '99999999999999'" & vbCrLf
        If Len(Conta) > 9 Then
            Sql &= "		AND RB.Conta = LEFT('" & Conta.Split(";")(0) & "',9) OR RB.Conta = LEFT('" & Conta.Split(";")(0) & "',7)" & vbCrLf
        Else
            Sql &= "		AND RB.Conta = LEFT('" & Conta & "',9) OR RB.Conta = LEFT('" & Conta & "',7)" & vbCrLf
        End If

        Sql &= "		group by RB.Referencial, RB.Grupo" & vbCrLf

        Sql &= ") AS Consulta" & vbCrLf &
                "GROUP BY Referencial" & vbCrLf &
                "HAVING (Sum(isnull(Inicial,0)) <> 0) OR Sum(isnull(Debitos,0)) <> 0 OR Sum(isnull(Creditos,0)) <> 0"

        If Conta = "101060102" Then
            Dim teste As String = "conta"
        End If

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroK155_K156(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date, _
                                        ByRef linha As String, ByRef RegistroK As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  REGISTRO K155 e K156 ******
        '***************************************************************************************************************
        Dim dsK155 As DataSet = ConsultaRegistroK155(Empresa, TrimestrePeriodoInicial, TrimestrePeriodoFinal)
        Dim dsK156 As DataSet

        For Each row In dsK155.Tables(0).Select("")
            linha = row("linha")
            Strm.WriteLine(linha)

            RegistroK += 1
            RegistroGeral += 1

            dsK156 = ConsultaRegistroK156(Empresa, row("Conta_ID"), TrimestrePeriodoInicial, TrimestrePeriodoFinal)
            For Each rowK156 In dsK156.Tables(0).Rows
                If Not IsDBNull(rowK156("linha")) Then
                    linha = rowK156("linha")
                    Strm.WriteLine(linha)

                    RegistroK += 1
                    RegistroGeral += 1
                End If
            Next
        Next
        'RegistroK += 1
        'RegistroGeral += 1
    End Sub

    Private Function ConsultaRegistroK355(ByRef Empresa As Cliente, ByRef PeriodoInicial As Date, ByRef PeriodoFinal As Date) As DataSet

        Dim Sql As String = String.Empty

        Sql = "SELECT Empresa_Id, Linha, LEFT(Conta_Id,9) as Conta_Id , Referencial" & vbCrLf & _
              "  FROM(" & vbCrLf & _
              "		SELECT C2.Empresa_Id, C2.Linha, LEFT(C2.Conta_Id,9) as Conta_Id" & vbCrLf & _
              "			  ,(SELECT TOP 1 [Referencial]" & vbCrLf & _
              "				  FROM [PlanoDeContasXReferencialBacen] RB" & vbCrLf & _
              "				 WHERE (RB.Empresa = LEFT(C2.Empresa_id,8) OR RB.Empresa = '99999999999999')" & vbCrLf & _
              "				   AND RB.Conta = LEFT(C2.Conta_id,9)) AS Referencial" & vbCrLf & _
              "		  FROM(" & vbCrLf & _
              "				 SELECT C1.Empresa_Id, LEFT(C1.Conta_Id,9) AS Conta_Id, '|K355|'+LEFT(C1.Conta_Id,9)" & vbCrLf & _
              "						+'||'+ REPLACE(REPLACE(CONVERT(VARCHAR,Sum(Saldo)),'.',','),'-','')" & vbCrLf & _
              "						+'|'+ Case when Sum(Saldo) > 0 then 'D' else 'C' end +'|' AS Linha" & vbCrLf & _
              "				  FROM" & vbCrLf & _
              "					  (" & vbCrLf & _
              "						SELECT Consulta.Empresa_Id, Consulta.Conta_Id, Consulta.Saldo" & vbCrLf & _
              "						  FROM" & vbCrLf & _
              "							  (SELECT LEFT(Empresa_Id, 8) AS Empresa_Id, EndEmpresa_Id" & vbCrLf & _
              "									 ,LEFT(Conta_Id, 7)  AS Conta_Id" & vbCrLf & _
              "									 ,SUM(DebitoOficial - CreditoOficial) AS Saldo" & vbCrLf & _
              "								 FROM Razao" & vbCrLf & _
              "								WHERE (LEFT(Conta_Id, 1) IN ('3', '4','6'))" & vbCrLf & _
              "								  AND Empresa_Id Like ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf & _
              "								  AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "')" & vbCrLf & _
              "								  AND (Lote_Id < 7500)" & vbCrLf & _
              "								  GROUP BY LEFT(Empresa_Id, 8), EndEmpresa_Id , Left(Conta_Id, 7)" & vbCrLf & _
              "							   ) AS Consulta" & vbCrLf & _
              "			      UNION" & vbCrLf & _
              "						SELECT Consulta.Empresa_Id, Consulta.Conta_Id, Consulta.Saldo" & vbCrLf & _
              "						   FROM" & vbCrLf & _
              "							   (SELECT LEFT(Empresa_Id, 8) AS Empresa_Id, EndEmpresa_Id" & vbCrLf & _
              "									  ,LEFT(Conta_id,9) AS Conta_Id" & vbCrLf & _
              "									  ,SUM(DebitoOficial - CreditoOficial) AS Saldo" & vbCrLf & _
              "								  FROM Razao" & vbCrLf & _
              "								 WHERE (LEFT(Conta_Id, 1) IN ('3', '4', '6'))" & vbCrLf & _
              "								   AND Empresa_Id Like ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf & _
              "								   AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "') " & vbCrLf & _
              "								   AND (Lote_Id < 7500)" & vbCrLf & _
              "								   GROUP BY LEFT(Empresa_Id, 8), EndEmpresa_Id , LEFT(Conta_id,9)" & vbCrLf & _
              "							   ) AS Consulta" & vbCrLf & _
              "				 ) AS C1" & vbCrLf & _
              "				   GROUP BY LEFT(C1.Conta_Id,9), C1.Saldo, C1.Empresa_Id" & vbCrLf & _
              "			 ) AS C2" & vbCrLf & _
              "    ) AS C3" & vbCrLf & _
              "      WHERE LEN(C3.Referencial) > 0" & vbCrLf & _
              "      GROUP BY C3.Conta_Id, C3.Empresa_Id, C3.Linha, C3.Referencial" & vbCrLf & _
              "      ORDER BY C3.Conta_Id" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")

    End Function

    Private Function ConsultaRegistroK356(ByRef Empresa As Cliente, ByRef Conta As String, ByRef PeriodoInicial As Date, ByRef PeriodoFinal As Date)
        Dim Sql As String = String.Empty

        Sql = " SELECT Linha" & vbCrLf & _
              "   FROM (" & vbCrLf & _
              " 		SELECT '|K356|'+LTRIM(RTRIM([Referencial]))" & vbCrLf & _
              " 			  +'|'+REPLACE(CONVERT(VARCHAR,CASE WHEN SaldoFinal < 0 THEN  SaldoFinal*-1 ELSE SaldoFinal END) ,'.',',')" & vbCrLf & _
              " 			  +'|'+CASE WHEN SUM(SaldoFinal) > 0 THEN 'D' ELSE 'C' END +'|' AS Linha" & vbCrLf & _
              " 		FROM (" & vbCrLf & _
              " 				SELECT  SUM(SaldoFinal)  AS  SaldoFinal  ,C1.[Referencial]" & vbCrLf & _
              " 				FROM (" & vbCrLf & _
              " 							   SELECT RB.[Empresa]" & vbCrLf & _
              "  									  ,RB.[Conta]" & vbCrLf & _
              "  									  ,RB.[Referencial]" & vbCrLf & _
              "  									  ,RB.[Grupo]" & vbCrLf & _
              "  									  ,RB.Produto" & vbCrLf & _
              "  									  , CASE WHEN LEN(RB.[Produto]) > 0" & vbCrLf & _
              "  											THEN" & vbCrLf & _
              "  											  ( SELECT ISNULL(SUM(Razao.DebitoOficial - Razao.CreditoOficial),0) AS SaldoFinal" & vbCrLf & _
              " 													  FROM Razao" & vbCrLf & _
              " 												WHERE Left(Razao.Empresa_Id, 8)  = '" & Empresa.Codigo.Substring(0, 8) & "'" & vbCrLf & _
              " 													   AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "') " & vbCrLf & _
              " 													   AND Razao.Lote_Id < 7500" & vbCrLf & _
              " 													   AND LEFT(Razao.Conta_Id,9) = LEFT(RB.Conta,9)" & vbCrLf & _
              " 													   AND  CASE WHEN LEN(RB.Produto ) > 0" & vbCrLf & _
              " 									  							  THEN RB.Produto ELSE 1 END =" & vbCrLf & _
              " 									  						 CASE WHEN LEN(RB.Produto) > 0 THEN RAZAO.Produto ELSE 1 END" & vbCrLf & _
              " 											  )" & vbCrLf & _
              " 											ELSE" & vbCrLf & _
              " 												  ( SELECT ISNULL(SUM(Razao.DebitoOficial - Razao.CreditoOficial),0) AS SaldoFinal  " & vbCrLf & _
              " 													  FROM Razao" & vbCrLf & _
              " 													 WHERE Left(Razao.Empresa_Id, 8)  = '" & Empresa.Codigo.Substring(0, 8) & "'" & vbCrLf & _
              "                                                        AND (Movimento_Id BETWEEN '" & PeriodoInicial.ToString("yyyy/MM/dd") & "' AND '" & PeriodoFinal.ToString("yyyy/MM/dd") & "') " & vbCrLf & _
              " 													   AND Razao.Lote_Id < 7500" & vbCrLf & _
              " 													   AND LEFT(Razao.Conta_Id,9) = LEFT(RB.Conta,9)" & vbCrLf & _
              " 													   AND  CASE WHEN LEN(RB.Grupo) > 0" & vbCrLf & _
              " 									   							 THEN  RB.Grupo  ELSE 1 END =" & vbCrLf & _
              " 									   						CASE WHEN LEN(RB.Grupo) > 0 THEN LEFT(RAZAO.Produto,3) ELSE 1 END" & vbCrLf & _
              " 											  )" & vbCrLf & _
              " 											END  AS SaldoFinal" & vbCrLf & _
              " 								  FROM [PlanoDeContasXReferencialBacen] RB" & vbCrLf & _
              " 								 WHERE (RB.Empresa LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%') OR RB.Empresa = '99999999999999')" & vbCrLf & _
              " 								   AND LEFT(RB.Conta,9) LIKE ('" & Conta & "%')" & vbCrLf & _
              " 				) AS C1" & vbCrLf & _
              " 				GROUP BY [Referencial]" & vbCrLf & _
              " 		) As C1" & vbCrLf & _
              " 		GROUP BY [Referencial], [SaldoFinal]" & vbCrLf & _
              " )AS C1" & vbCrLf & _
              " WHERE LEN(C1.Linha) > 0" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroK355_K356(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date,
                                       ByRef linha As String, ByRef RegistroK As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  REGISTRO K155 e K156 ******
        '***************************************************************************************************************
        Dim dsK355 As DataSet = ConsultaRegistroK355(Empresa, TrimestrePeriodoInicial, TrimestrePeriodoFinal)
        Dim dsK356 As DataSet

        For Each row In dsK355.Tables(0).Rows
            linha = row("linha")
            Strm.WriteLine(linha)

            RegistroK += 1
            RegistroGeral += 1

            dsK356 = ConsultaRegistroK356(Empresa, row("Conta_ID"), TrimestrePeriodoInicial, TrimestrePeriodoFinal)
            For Each rowK156 In dsK356.Tables(0).Rows
                linha = rowK156("linha")
                Strm.WriteLine(linha)

                RegistroK += 1
                RegistroGeral += 1
            Next
        Next

        'Strm.WriteLine(linha)
    End Sub

#End Region

#Region "Registro L"
    Private Sub CompoeRegistroL001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroL As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO M001: ABERTURA DO BLOCO M   ************************************
        '******************************************************************************************************

        linha = "|L001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroL += 1
        RegistroGeral += 1
    End Sub

    Private Sub CompoeRegistroL030(ByRef Strm As StreamWriter, ByRef Trimestre As String, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date, _
                                  ByRef linha As String, ByRef RegistroL As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  Registro L030: Identificação dos Períodos e Formas de Apuração do IRPJ e da CSLL no Ano-Calendário ******
        '***************************************************************************************************************
        linha = "|L030"                                            '01 - REG      
        linha &= "|" & TrimestrePeriodoInicial.ToString("ddMMyyyy")  '02 - DT_INI
        linha &= "|" & TrimestrePeriodoFinal.ToString("ddMMyyyy")  '03 - DT_FIN
        linha &= "|" & Trimestre & "|"                                            '04 - PER_APUR   

        Strm.WriteLine(linha)
        RegistroL += 1
        RegistroGeral += 1
    End Sub


    Private Sub CompoeRegistroL200(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroL As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  Registro L200: Método de Avaliação do Estoque Final ****************************************************
        '***************************************************************************************************************
        linha = "|L200"                                            '01 - REG      
        linha &= "|1|"

        Strm.WriteLine(linha)
        RegistroL += 1
        RegistroGeral += 1
    End Sub


    Private Function ConsultaRegistroL210(ByRef Empresa As Cliente, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date) As DataSet

        Dim Sql As String = String.Empty

        Sql = " 	 SELECT '|L210|'" & vbCrLf & _
              " 		     +CONVERT(VARCHAR,CodigoDeCustoECF)" & vbCrLf & _
              " 		     +'|'+(SELECT Descricao FROM PlanoDeContasCustosECF WHERE Codigo_Id = CodigoDeCustoECF)" & vbCrLf & _
              " 		     +'|'+REPLACE(CONVERT(VARCHAR,(SUM(Valor))),'.',',')+'|'  AS Linha" & vbCrLf & _
              " 	   FROM (" & vbCrLf & _
              " 			 SELECT  PlanoDeContas.CodigoDeCustoECF,(Consulta.Debitos-Consulta.Creditos) AS Valor" & vbCrLf & _
              "		       FROM (" & vbCrLf & _
              "					 SELECT Pais_Id, LEFT(Empresa_Id, 8) AS Empresa" & vbCrLf & _
              "						   ,EndEmpresa_Id" & vbCrLf & _
              "						   ,Conta_Id  AS Conta" & vbCrLf & _
              "						   ,0 as Inicial" & vbCrLf & _
              "						   ,SUM(DebitoOficial) AS Debitos" & vbCrLf & _
              "						   ,SUM(CreditoOficial) AS Creditos	,LEFT(Produto,3) as GrupoProduto" & vbCrLf & _
              "					   FROM Razao" & vbCrLf & _
              "			  WHERE Empresa_Id  Like ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf & _
              "				AND (Movimento_Id <= '" & TrimestrePeriodoFinal & "')" & vbCrLf & _
              "				AND Lote_Id < 7500" & vbCrLf & _
              "		   GROUP BY Pais_Id, LEFT(Empresa_Id, 8), EndEmpresa_Id, Conta_Id,Produto" & vbCrLf & _
              "		   ) AS Consulta" & vbCrLf & _
              "		    LEFT OUTER JOIN PlanoDeContasXCustosECF AS PlanoDeContas" & vbCrLf & _
              "		         ON CASE WHEN LEN(Consulta.Conta) > 5 THEN LEFT(Consulta.Conta,9) ELSE Consulta.Conta END = PlanoDeContas.Conta_Id " & vbCrLf & _
              "				AND Consulta.Pais_Id = PlanoDeContas.Pais_Id" & vbCrLf & _
              "   			AND Consulta.Empresa = Left(PlanoDeContas.Empresa_Id, 8)" & vbCrLf & _
              "				AND Consulta.EndEmpresa_Id = PlanoDeContas.EndEmpresa_Id" & vbCrLf & _
              "				AND  CASE WHEN LEN(PlanoDeContas.Grupo_Id) > 0" & vbCrLf & _
              "				          THEN Consulta.GrupoProduto ELSE 1 END =" & vbCrLf & _
              "				     CASE WHEN LEN(PlanoDeContas.Grupo_Id) > 0 THEN PlanoDeContas.Grupo_Id ELSE 1 END" & vbCrLf & _
              "	          WHERE LEN(CodigoDeCustoECF) > 0 OR CodigoDeCustoECF <> ''" & vbCrLf & _
              "   ) AS Consulta	GROUP BY CodigoDeCustoECF" & vbCrLf & _
              "    UNION" & vbCrLf & _
              "  	 SELECT '|L210|'" & vbCrLf & _
              "  		     +CONVERT(VARCHAR,CodigoDeCustoECF)" & vbCrLf & _
              "  		     +'|'+(SELECT Descricao FROM PlanoDeContasCustosECF WHERE Codigo_Id = CodigoDeCustoECF)" & vbCrLf & _
              "  		     +'|'+REPLACE(CONVERT(VARCHAR,(SUM(Valor))),'.',',')+'|'  AS Linha" & vbCrLf & _
              "	 	   FROM (" & vbCrLf & _
              "	 			 SELECT  PlanoDeContas.CodigoDeCustoECF,(Consulta.Debitos-Consulta.Creditos) AS Valor" & vbCrLf & _
              "	 		       FROM (" & vbCrLf & _
              "	 					 SELECT Pais_Id, LEFT(Empresa_Id, 8) AS Empresa" & vbCrLf & _
              "	 						   ,EndEmpresa_Id" & vbCrLf & _
              "	 						   ,Conta_Id  AS Conta" & vbCrLf & _
              "	 						   ,0 as Inicial" & vbCrLf & _
              "	 						   ,SUM(DebitoOficial) AS Debitos" & vbCrLf & _
              "	 						   ,SUM(CreditoOficial) AS Creditos	,LEFT(Produto,3) as GrupoProduto" & vbCrLf & _
              "	 					   FROM Razao" & vbCrLf & _
              "	 			  WHERE Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf & _
              "	 				AND (Movimento_Id <= '" & TrimestrePeriodoFinal & "')" & vbCrLf & _
              "	 				AND Lote_Id < 7500" & vbCrLf & _
              "	 		   GROUP BY Pais_Id, LEFT(Empresa_Id, 8), EndEmpresa_Id, Conta_Id ,Produto" & vbCrLf & _
              "	 		   ) AS Consulta" & vbCrLf & _
              "	 		    LEFT OUTER JOIN PlanoDeContasXCustosECF AS PlanoDeContas" & vbCrLf & _
              "	 		         ON LEFT(Consulta.Conta,5) = PlanoDeContas.Conta_Id AND LEN(PlanoDeContas.Conta_Id) = 5" & vbCrLf & _
              "  				AND Consulta.Pais_Id = PlanoDeContas.Pais_Id" & vbCrLf & _
              "	  			AND Consulta.Empresa = Left(PlanoDeContas.Empresa_Id, 8)" & vbCrLf & _
              "	 				AND Consulta.EndEmpresa_Id = PlanoDeContas.EndEmpresa_Id" & vbCrLf & _
              "	 				AND  CASE WHEN LEN(PlanoDeContas.Grupo_Id) > 0" & vbCrLf & _
              "	 				          THEN Consulta.GrupoProduto ELSE 1 END =" & vbCrLf & _
              "	 				     CASE WHEN LEN(PlanoDeContas.Grupo_Id) > 0 THEN PlanoDeContas.Grupo_Id ELSE 1 END" & vbCrLf & _
              "	 	          WHERE LEN(CodigoDeCustoECF) > 0" & vbCrLf & _
              "        ) AS Consulta	GROUP BY CodigoDeCustoECF" & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroL210(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef TrimestrePeriodoInicial As Date, ByRef TrimestrePeriodoFinal As Date, _
                                        ByRef linha As String, ByRef RegistroL As Integer, ByRef RegistroGeral As Integer)
        '***************************************************************************************************************
        '*****  REGISTRO K155 e K156 ******
        '***************************************************************************************************************
        Dim dsL210 As DataSet = ConsultaRegistroL210(Empresa, TrimestrePeriodoInicial, TrimestrePeriodoFinal)

        For Each row In dsL210.Tables(0).Select("")
            linha = row("linha")
            'dsK156 = ConsultaRegistroK156(Empresa, row("Conta_ID"), TrimestrePeriodoInicial, TrimestrePeriodoFinal)
            'For Each rowK156 In dsK156.Tables(0).Rows
            '    linha = row("linha")
            'Next
        Next

        Strm.WriteLine(linha)
        RegistroL += 1
        RegistroGeral += 1
    End Sub


#End Region

#Region "Registro M"
    Private Sub CompoeRegistroM001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroM As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO M001: ABERTURA DO BLOCO M   ************************************
        '******************************************************************************************************

        linha = "|M001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroM += 1
        RegistroGeral += 1
    End Sub
#End Region

#Region "Registro Y"
    Private Sub CompoeRegistroY001(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroY As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '**************************** REGISTRO J001: ABERTURA DO BLOCO J   ************************************
        '******************************************************************************************************

        linha = "|Y001"                                         '01 - REG      
        linha &= "|0|"                                          '02 - IND_DAD

        Strm.WriteLine(linha)
        RegistroY += 1
        RegistroGeral += 1
    End Sub

    Private Function ConsultaRegistroY520(ByRef Empresa As Cliente) As DataSet
        Dim Sql As String

        'Sql = "SELECT 'Y520' REG, T.Tip_Ext, LEFT(Cli.Pais, LEN(Cli.Pais) -1) Pais, ECF_FormaDePagamento Forma, ECF_NaturezaOpereracao Nat_Oper,SUM(T.ValorLiquido) Vl_Periodo" & vbCrLf &
        '      "FROM (" & vbCrLf &
        '      "		SELECT 'P' Tip_Ext, EmpresaPagadora, 1 as ECF_FormaDePagamento, 12012 as ECF_NaturezaOpereracao, Situacao, " & vbCrLf &
        '      "		       Provisao, Cliente, EndCliente, Baixa, ValorLiquido " & vbCrLf &
        '      "		  FROM ContasAPagar CP" & vbCrLf &
        '      "		 UNION ALL" & vbCrLf &
        '      "		SELECT 'R' Tip_Ext, EmpresaPagadora,  1 as ECF_FormaDePagamento, 12005 as ECF_NaturezaOpereracao, Situacao, " & vbCrLf &
        '      "		        Provisao, Cliente, EndCliente, Baixa, ValorLiquido " & vbCrLf &
        '      "		  FROM ContasAReceber Cr" & vbCrLf &
        '      " ) T" & vbCrLf &
        '      " JOIN Clientes Cli" & vbCrLf &
        '      "    ON T.Cliente = Cli.Cliente_Id" & vbCrLf &
        '      "   AND T.EndCliente = Cli.Endereco_Id" & vbCrLf &
        '      " WHERE Cli.Estado = 'EX'" & vbCrLf &
        '      "   AND T.EmpresaPagadora = '" & Empresa.Codigo & "'" & vbCrLf &
        '      "   AND T.Baixa BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
        '      "   AND T.Provisao = 1 " & vbCrLf &
        '      "   AND T.Situacao = 1" & vbCrLf &
        '      " GROUP BY T.Tip_Ext, Cli.Pais, ECF_FormaDePagamento, ECF_NaturezaOpereracao" & vbCrLf

        Sql = "SELECT 'Y520' REG, T.Tip_Ext, LEFT(Cli.Pais, LEN(Cli.Pais) -1) Pais, ECF_FormaDePagamento Forma, ECF_NaturezaOpereracao Nat_Oper,SUM(T.ValorLiquido) Vl_Periodo" & vbCrLf &
                "FROM (" & vbCrLf &
                "	SELECT 'P' Tip_Ext, CP.EmpresaPagadora, 1 as ECF_FormaDePagamento, 12012 as ECF_NaturezaOpereracao, CP.Situacao, " & vbCrLf &
                "			CP.Provisao, CP.Cliente, CP.EndCliente, CP.Baixa, CP.ValorLiquido " & vbCrLf &
                "	FROM ContasAPagar CP " & vbCrLf &
                "		INNER JOIN Clientes Cli" & vbCrLf &
                "				ON Cli.Cliente_Id   = CP.Cliente" & vbCrLf &
                "				AND Cli.Endereco_Id = CP.EndCliente" & vbCrLf &
                "	WHERE CP.EmpresaPagadora LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
                "	AND CP.Baixa BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "	AND CP.Provisao = 1 " & vbCrLf &
                "	AND CP.Situacao = 1" & vbCrLf &
                "	AND Cli.Estado = 'EX'" & vbCrLf &
                "	UNION ALL" & vbCrLf &
                "	SELECT 'R' Tip_Ext, CR.EmpresaPagadora,  1 as ECF_FormaDePagamento, 12005 as ECF_NaturezaOpereracao, CR.Situacao, " & vbCrLf &
                "			CR.Provisao, CR.Cliente, CR.EndCliente, CR.Baixa, CR.ValorLiquido " & vbCrLf &
                "	FROM ContasAReceber CR" & vbCrLf &
                "		INNER JOIN Clientes Cli" & vbCrLf &
                "				ON Cli.Cliente_Id   = CR.Cliente" & vbCrLf &
                "				AND Cli.Endereco_Id = CR.EndCliente" & vbCrLf &
                "	WHERE CR.EmpresaPagadora LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
                "	AND CR.Baixa BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "	AND CR.Provisao = 1 " & vbCrLf &
                "	AND CR.Situacao = 1" & vbCrLf &
                "	AND Cli.Estado = 'EX'" & vbCrLf &
                "	UNION ALL" & vbCrLf &
                "	SELECT 'R' Tip_Ext, r.Empresa_Id AS EmpresaPagadora, 1 as ECF_FormaDePagamento, 12005 as ECF_NaturezaOpereracao, 1 AS Situacao, " & vbCrLf &
                "			1 AS Provisao, r.Cliente_Id AS Cliente, r.EndCliente_Id AS EndCliente, r.Movimento_Id AS Baixa, SUM(CreditoOficial) AS ValorLiquido" & vbCrLf &
                "	FROM Razao r" & vbCrLf &
                "			INNER JOIN Clientes Cli" & vbCrLf &
                "					ON Cli.Cliente_Id = r.Cliente_Id" & vbCrLf &
                "					AND Cli.Endereco_Id = r.EndCliente_Id" & vbCrLf &
                "	WHERE r.Empresa_Id LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
                "	AND r.Movimento_id between '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "	AND Lote_id in(1,2,3,4,5)" & vbCrLf &
                "	AND Cli.Estado = 'EX'" & vbCrLf &
                "	AND Conta_Id IN('1010202')" & vbCrLf &
                "	AND Historico not LIKE ('%VARIACAO%')" & vbCrLf &
                "	AND Historico not LIKE ('%CAMBIAL%')" & vbCrLf &
                "	AND Historico not LIKE ('%PGTO COMISSAO%')" & vbCrLf &
                "	AND Historico not LIKE ('%DEBITO COMISSAO%')" & vbCrLf &
                "	GROUP BY r.Empresa_Id, r.Cliente_Id, r.EndCliente_Id, r.Movimento_Id" & vbCrLf &
                ") T" & vbCrLf &
                "JOIN Clientes Cli" & vbCrLf &
                "    ON T.Cliente = Cli.Cliente_Id" & vbCrLf &
                "    AND T.EndCliente = Cli.Endereco_Id" & vbCrLf &
                "WHERE Cli.Estado = 'EX'" & vbCrLf &
                "    AND T.EmpresaPagadora LIKE ('" & Empresa.Codigo.Substring(0, 8) & "%')" & vbCrLf &
                "    AND T.Baixa BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
                "    AND T.Provisao = 1 " & vbCrLf &
                "    AND T.Situacao = 1" & vbCrLf &
                "GROUP BY T.Tip_Ext, Cli.Pais, ECF_FormaDePagamento, ECF_NaturezaOpereracao" & vbCrLf &
                "HAVING SUM(T.ValorLiquido) > 0" & vbCrLf &
                "ORDER BY T.Tip_Ext"

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroY520(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef linha As String, ByRef RegistroY As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '*******************    REGISTRO Y520:      ****************************
        '******************************************************************************************************

        Dim ds As DataSet = ConsultaRegistroY520(Empresa)

        For Each row In ds.Tables(0).Rows
            linha = "|" & row("REG")
            linha &= "|" & row("TIP_EXT")
            linha &= "|" & row("PAIS")
            linha &= "|" & row("FORMA")
            linha &= "|" & row("NAT_OPER")
            linha &= "|" & row("VL_PERIODO") & "|"

            Strm.WriteLine(linha)
            RegistroY += 1
            RegistroGeral += 1
        Next

    End Sub

    Private Function ConsultaRegistroY540(ByRef Empresa As Cliente) As DataSet
        Dim Sql As String

        Sql = "SELECT 'Y540' as REG, NF.Empresa_Id AS CNPJ_ESTAB," & vbCrLf &
              "       SUM(CASE" & vbCrLf &
              "               WHEN So.Devolucao = 'S'" & vbCrLf &
              "                   THEN NFxE.valor * -1" & vbCrLf &
              "                   ELSE NFxE.valor" & vbCrLf &
              "           END) as VL_REC_ESTAB," & vbCrLf &
              "        Pr.CNAE" & vbCrLf &
              "  FROM NotasFiscais AS NF " & vbCrLf &
              "  JOIN NotasFiscaisXItens AS NFxI " & vbCrLf &
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
              "   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
              "   AND NF.Nota_Id         = NFxI.Nota_Id  " & vbCrLf &
              "  JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf &
              "    ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf &
              "   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf &
              "   AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf &
              "   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf &
              "   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "   AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf &
              "   AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf &
              "   AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf &
              "   AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf &
              "   AND NFxI.Sequencia_Id    = NFxE.Sequencia_id" & vbCrLf &
              "   AND NFxE.Encargo_Id      = 'LIQUIDO'" & vbCrLf &
              "  JOIN (" & vbCrLf &
              "         SELECT Empresa_Id, EndEmpresa_Id, Cliente_NF, EndCliente_Nf, Numero_NF, Serie_NF, EntradaSaida_Nf, Produto_Nf" & vbCrLf &
              "           FROM Razao " & vbCrLf &
              "          WHERE Conta_id IN ('301010101', '301010103', '301020301', '301010102')" & vbCrLf &
              "            AND Movimento_Id BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
              "        ) AS R" & vbCrLf &
              "    ON R.Empresa_Id = Nf.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf &
              "   AND R.Cliente_Nf = NF.Cliente_Id" & vbCrLf &
              "   AND R.EndCliente_Nf = NF.EndCliente_Id" & vbCrLf &
              "   AND R.Numero_Nf = NF.Nota_Id" & vbCrLf &
              "   AND R.EntradaSaida_Nf = NF.EntradaSaida_Id" & vbCrLf &
              "   AND R.Serie_Nf = NF.Serie_Id" & vbCrLf &
              "   AND R.Produto_Nf = NFxi.Produto_Id" & vbCrLf &
              "  JOIN Produtos Pr" & vbCrLf &
              "    ON NFxI.Produto_id = Pr.Produto_id" & vbCrLf &
              "  JOIN Suboperacoes SO" & vbCrLf &
              "    ON SO.Operacao_Id        = NFxI.Operacao" & vbCrLf &
              "   AND SO.Suboperacoes_Id    = NFxI.Suboperacao" & vbCrLf &
              " WHERE NF.Movimento BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf &
              "   AND NF.Situacao = 1" & vbCrLf &
              " GROUP BY NF.Empresa_Id, PR.Cnae" & vbCrLf &
              " ORDER BY NF.Empresa_Id, PR.Cnae " & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroY540(ByRef Strm As StreamWriter, ByRef Empresa As Cliente, ByRef linha As String, ByRef RegistroY As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '*******************    REGISTRO Y540:      ****************************
        '******************************************************************************************************

        Dim ds As DataSet = ConsultaRegistroY540(Empresa)

        For Each row In ds.Tables(0).Rows
            linha = "|" & row("REG")
            linha &= "|" & row("CNPJ_ESTAB")
            linha &= "|" & row("VL_REC_ESTAB")
            linha &= "|" & row("CNAE") & "|"

            Strm.WriteLine(linha)
            RegistroY += 1
            RegistroGeral += 1
        Next

    End Sub

    Private Function ConsultaRegistroY550_Y560() As DataSet
        Dim Sql As String

        Sql = "SELECT NF.Cliente_Id AS CNPJ," & vbCrLf & _
              "       Pr.NCM AS COD_NCM, " & vbCrLf & _
              "       SUM(NFR.Valor) VL_COMPRA, " & vbCrLf & _
              "       SUM(CASE" & vbCrLf & _
              "               WHEN So.Devolucao = 'S'" & vbCrLf & _
              "                   THEN NFxE.valor * -1" & vbCrLf & _
              "                   ELSE NFxE.valor" & vbCrLf & _
              "           END) as VL_VENDA," & vbCrLf & _
              "       NFxI.CFOP_Id AS CFOP " & vbCrLf & _
              "  FROM NotasFiscais AS NF " & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens AS NFxI " & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id  " & vbCrLf & _
              " INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf & _
              "    ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf & _
              "   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf & _
              "   AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf & _
              "   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf & _
              "   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf & _
              "   AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf & _
              "   AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf & _
              "   AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf & _
              "   AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf & _
              "   AND NFxI.Sequencia_Id    = NFxE.Sequencia_id" & vbCrLf & _
              "   AND NFxE.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
              "  LEFT JOIN NotaFiscalReferencial NFR" & vbCrLf & _
              "    ON NFR.Empresa_Id = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NFR.EndEmpresa_Id  = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NFR.Cliente_Id = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NFR.EndCliente_Id = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NFR.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NFR.Nota_Id = NFxI.Nota_Id" & vbCrLf & _
              "   AND NFR.Serie_Id = NFxI.Serie_Id" & vbCrLf & _
              "   AND NFR.Produto_Id = NFxI.Produto_Id" & vbCrLf & _
              "   AND NFR.CFOP_Id = NFxI.CFOP_Id" & vbCrLf & _
              "   AND NFR.Sequencia_Id = NFxI.Sequencia_Id" & vbCrLf & _
              " INNER JOIN produtos PR" & vbCrLf & _
              "    ON NFxI.Produto_id = Pr.Produto_id" & vbCrLf & _
              " INNER JOIN Suboperacoes SO" & vbCrLf & _
              "    ON SO.Operacao_Id        = NFxI.Operacao" & vbCrLf & _
              "   AND SO.Suboperacoes_Id    = NFxI.Suboperacao" & vbCrLf & _
              " WHERE NF.Movimento BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
              "   AND NF.Situacao = 1" & vbCrLf & _
              "   AND NFxi.CFOP_Id IN (2501, 1501, 5502, 6502)  " & vbCrLf & _
              " GROUP BY NFxI.CFOP_ID, NF.Cliente_Id, Pr.NCM " & vbCrLf & _
              " ORDER BY NFxI.CFOP_ID, NF.Cliente_Id, PR.NCM " & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Function ConsultaRegistroY550() As DataSet
        Dim Sql As String

        Sql = "SELECT 'Y550' as REG, NF.Cliente_Id AS CNPJ_EXP," & vbCrLf & _
              "       Pr.NCM AS COD_NCM," & vbCrLf & _
              "       SUM(CASE" & vbCrLf & _
              "               WHEN So.Devolucao = 'S'" & vbCrLf & _
              "                   THEN NFxE.valor * -1" & vbCrLf & _
              "                   ELSE NFxE.valor" & vbCrLf & _
              "           END) as VL_VENDA" & vbCrLf & _
              "  FROM NotasFiscais AS NF " & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens AS NFxI " & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id  " & vbCrLf & _
              " INNER JOIN NotasFiscaisXEncargos AS NFxE " & vbCrLf & _
              "    ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf & _
              "   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf & _
              "   AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf & _
              "   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf & _
              "   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf & _
              "   AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf & _
              "   AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf & _
              "   AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf & _
              "   AND NFxI.CFOP_Id         = NFxE.CFOP_Id " & vbCrLf & _
              "   AND NFxI.Sequencia_Id    = NFxE.Sequencia_id" & vbCrLf & _
              "   AND NFxE.Encargo_Id      = 'LIQUIDO'" & vbCrLf & _
              " INNER JOIN produtos Pr" & vbCrLf & _
              "    ON NFxI.Produto_id = Pr.Produto_id" & vbCrLf & _
              " INNER JOIN Suboperacoes SO" & vbCrLf & _
              "    ON SO.Operacao_Id        = NFxI.Operacao" & vbCrLf & _
              "   AND SO.Suboperacoes_Id    = NFxI.Suboperacao" & vbCrLf & _
              " WHERE NF.Movimento BETWEEN '" & CDate(txtDT_INI.Text).ToString("yyyy/MM/dd") & "' AND '" & CDate(txtDT_FIN.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf & _
              "   AND NF.Situacao = 1" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = 'S'" & vbCrLf & _
              "   AND NFxi.CFOP_Id IN (5502, 6502) " & vbCrLf & _
              " GROUP BY NF.Cliente_Id, Pr.NCM " & vbCrLf & _
              " ORDER BY NF.Cliente_Id, PR.NCM " & vbCrLf

        Return Banco.ConsultaDataSet(Sql, "Consulta")
    End Function

    Private Sub CompoeRegistroY550(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroY As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '*******************    REGISTRO Y540:      ****************************
        '******************************************************************************************************

        Dim ds As DataSet = ConsultaRegistroY550_Y560()

        For Each row As DataRow In ds.Tables(0).Select("CFOP IN (5502, 6502)")
            linha = "|" & "Y550" & "|"
            linha &= row("CNPJ") & "|"
            linha &= row("COD_NCM") & "|"
            linha &= row("VL_VENDA") & "|"

            Strm.WriteLine(linha)
            RegistroY += 1
            RegistroGeral += 1
        Next

    End Sub

    Private Sub CompoeRegistroY560(ByRef Strm As StreamWriter, ByRef linha As String, ByRef RegistroY As Integer, ByRef RegistroGeral As Integer)
        '******************************************************************************************************
        '*******************    REGISTRO Y540:      ****************************
        '******************************************************************************************************

        Dim ds As DataSet = ConsultaRegistroY550_Y560()

        For Each row As DataRow In ds.Tables(0).Select("CFOP IN (1501, 2501)")
            linha = "|" & "Y560" & "|"
            linha &= row("CNPJ") & "|"
            linha &= row("COD_NCM") & "|"
            linha &= row("VL_COMPRA") & "|"
            linha &= row("VL_VENDA") & "|"

            Strm.WriteLine(linha)
            RegistroY += 1
            RegistroGeral += 1
        Next

    End Sub

#End Region

    Protected Sub DdlSIT_ESPECIAL_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlSIT_ESPECIAL.SelectedIndexChanged
        txtPAT_REMAN_CIS.Parent.Visible = DdlSIT_ESPECIAL.SelectedValue = 6
        txtPAT_REMAN_CIS.Text = String.Empty

        txtDT_SIT_ESP.Parent.Visible = DdlSIT_ESPECIAL.SelectedValue <> 0
        txtDT_SIT_ESP.Text = String.Empty
    End Sub

    Protected Sub ddlRETIFICADORA_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlRETIFICADORA.SelectedIndexChanged
        txtNUM_REC.Parent.Visible = ddlRETIFICADORA.SelectedValue <> "N"
        txtNUM_REC.Text = String.Empty
    End Sub

    Protected Sub cmdArquivoDeSaida_Click(sender As Object, e As EventArgs) Handles cmdArquivoDeSaida.Click
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & txtArquivoDeSaida.Text)
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(Server.MapPath("~/ECF/" & txtArquivoDeSaida.Text), FileMode.Open, FileAccess.Read)
                Do
                    If Response.IsClientConnected Then
                        length = download.Read(buffer, 0, bufferLength)
                        Response.OutputStream.Write(buffer, 0, length)
                        buffer = New Byte(bufferLength - 1) {}
                    Else
                        length = -1
                    End If
                Loop While length > 0
                Response.Flush()
                Response.End()
            Finally
                If download IsNot Nothing Then
                    download.Close()
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub IND_ALIQ_CSLL_Mostrar_Ocultar()
        If CDate(txtDT_FIN.Text).Year = 2014 Then
            Div_Ind_Aliq_Csll_I.Style.Add("Display", "none")
            Div_Ind_Aliq_Csll.Style.Remove("Display")
        Else
            Div_Ind_Aliq_Csll_I.Style.Remove("Display")
            Div_Ind_Aliq_Csll.Style.Add("Display", "none")
        End If
    End Sub

    Protected Sub txtDT_FIN_TextChanged(sender As Object, e As EventArgs) Handles txtDT_FIN.TextChanged
        IND_ALIQ_CSLL_Mostrar_Ocultar()
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Response.Redirect(Me.Request.Url.ToString())
    End Sub

End Class