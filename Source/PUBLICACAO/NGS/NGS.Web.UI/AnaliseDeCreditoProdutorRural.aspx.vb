Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Partial Class AnaliseDeCreditoProdutorRural
    Inherits BasePage

#Region "Page_Load"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AnaliseDeCreditoProdutorRural", "ACESSAR") Then
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                tcAnaliseCredito.ActiveTabIndex = 1
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub
#End Region

#Region "Variaveis Locais"
    Dim ObjAnalise As AnaliseDeCredito
#End Region

#Region "Sessao"
    Public Sub CarregarCliente()

    End Sub

    Public Sub CarregarClienteCC()
        If Session("objClienteAnCredCC" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteAnCredCC" & HID.Value), [Lib].Negocio.Cliente))
            txtClienteConsulta.Text = itemCliente.Text
            HCliente.Value = itemCliente.Value
            Session.Remove("objClienteAXP" & HID.Value)
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteAnCredRural" & HID.Value) Is Nothing Then
            SessaoRecuperaAnaliseDeCredito()
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteAnCredRural" & HID.Value)
            If ObjAnalise.Clientes.AdicionarCliente(cli) Then
                VerificaIUD()
                ObjAnalise.AtualizarAnalise()
                SessaoSalvaAnaliseDeCredito()
                AtualizaFormComAClasse()
                If ObjAnalise.JaExisteCreditoLiberadoNesteAno Then
                    lnkAprovarSimulacao.Enabled = False
                    lnkLiberarCredito.Enabled = False
                    MsgBox(Me.Page, "Já existe uma simulação Liberada para este cliente no Ano.")
                End If
            End If
            Session.Remove("objClienteAnCredRural" & HID.Value)
        End If
    End Sub

    Private Sub SessaoSalvaAnaliseDeCredito()
        Session("ObjAnalise") = ObjAnalise
    End Sub

    Private Sub SessaoRecuperaAnaliseDeCredito()
        ObjAnalise = CType(Session("ObjAnalise"), AnaliseDeCredito)
    End Sub
#End Region

    Public Sub VerificaIUD()
        If ObjAnalise.IUD = "U" Then
            MsgBox(Me.Page, "Nao se pode alterar uma analise de credito ja existente, uma nova analise sera criada a partir dos novos parametros")
            Dim num As New [Lib].Negocio.Numerador(3)
            ObjAnalise.IUD = "I"
            ObjAnalise.CodigoAnalise = num.Sequencia + 1
            ObjAnalise.ParametrosAnaliseDeCredito = New [Lib].Negocio.ParametrosAnaliseDeCredito(ObjAnalise.ParametrosAnaliseDeCredito.Ano)
            ObjAnalise.DataAnalise = Date.Now
            ObjAnalise.UsuarioAprovacao = ""
            ObjAnalise.UsuarioLiberacao = ""
            ObjAnalise.UsuarioCancelamento = ""
            ObjAnalise.Situacao = "NORMAL"
            lblAnalise.Text = ObjAnalise.CodigoAnalise
            lblDefinicao.Text = "Definicao Num. " & ObjAnalise.ParametrosAnaliseDeCredito.DefinicaoAno & " de " & ObjAnalise.ParametrosAnaliseDeCredito.DataDefinicao.ToString("dd/MM/yyyy")
        End If
    End Sub

    Protected Sub ddlAno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlAno.SelectedIndex = 0 Then
            lblDefinicao.Text = ""
            Exit Sub
        End If

        Dim ObjParametros As [Lib].Negocio.ParametrosAnaliseDeCredito = New [Lib].Negocio.ParametrosAnaliseDeCredito(ddlAno.SelectedValue)
        'SessaoSalvaParametros()
        Dim num As New [Lib].Negocio.Numerador(3)

        lblDefinicao.Text = "Definicao Num. " & ObjParametros.DefinicaoAno & " de " & ObjParametros.DataDefinicao.ToString("dd/MM/yyyy")

        ObjAnalise = New AnaliseDeCredito(ObjParametros)
        ObjAnalise.IUD = "I"
        ObjAnalise.CodigoAnalise = num.Sequencia + 1
        ObjAnalise.DataAnalise = Date.Now

        lblAnalise.Text = ObjAnalise.CodigoAnalise
        ddlSituacao.SelectedValue = "NORMAL"
        SessaoSalvaAnaliseDeCredito()
    End Sub

    Protected Sub btnIncluirCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnIncluirCliente.Click
        If ddlAno.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Selecione um ano para Parametrizacao")
            Exit Sub
        End If
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteAnCredRural" & HID.Value, "txtNome")
    End Sub

    Public Sub AtualizaFormComAClasse()
        SessaoRecuperaAnaliseDeCredito()
        If ObjAnalise.ParametrosAnaliseDeCredito.Ano > 0 Then
            ddlAno.SelectedValue = ObjAnalise.ParametrosAnaliseDeCredito.Ano
            lblDefinicao.Text = "Definicao Num. " & ObjAnalise.ParametrosAnaliseDeCredito.DefinicaoAno & " de " & ObjAnalise.ParametrosAnaliseDeCredito.DataDefinicao.ToString("dd/MM/yyyy")
        End If

        lblAnalise.Text = ObjAnalise.CodigoAnalise
        ddlSituacao.SelectedValue = ObjAnalise.Situacao

        If ObjAnalise.UsuarioAprovacao <> "" Then
            lblUsuarioApr.Text = ObjAnalise.DataAprovacao.ToString("dd-MM-yyyy") & " | " & ObjAnalise.UsuarioAprovacao
        Else
            lblUsuarioApr.Text = ""
        End If

        If ObjAnalise.UsuarioLiberacao <> "" Then
            lblUsuarioLib.Text = ObjAnalise.DataLiberacao.ToString("dd-MM-yyyy") & " | " & ObjAnalise.UsuarioLiberacao
        Else
            lblUsuarioLib.Text = ""
        End If

        If ObjAnalise.UsuarioCancelamento <> "" Then
            lblUsuarioCan.Text = ObjAnalise.DataCancelamento.ToString("dd-MM-yyyy") & " | " & ObjAnalise.UsuarioCancelamento
        Else
            lblUsuarioCan.Text = ""
        End If

        gridClientes.DataSource = ObjAnalise.Clientes.ToArray
        gridClientes.DataBind()

        gridPerguntas.DataSource = ObjAnalise.Perguntas.ToArray
        gridPerguntas.DataBind()

        gridCulturas.DataSource = ObjAnalise.Culturas.ToArray
        gridCulturas.DataBind()

        FormatRatingDeCredito()
    End Sub

    Public Sub AtualizaResumo()
        txtCoefRedutor.Text = ObjAnalise.Culturas.CoefRedutorRiscoCultura.ToString("N2")
        txtPotencialCompraPortifolio.Text = ObjAnalise.Culturas.TotalCustoPortifolio.ToString("N2")

        txtReceitaCulturas.Text = ObjAnalise.Culturas.TotalReceitaCultura.ToString("N2")
        txtOutrasReceitas.Text = ObjAnalise.OutrasReceitas.ToString("N2")
        txtTotalReceitas.Text = ObjAnalise.TotalReceitas.ToString("N2")

        txtCustoCultura.Text = ObjAnalise.Culturas.TotalCustoCultura.ToString("N2")
        txtCustoArrendamento.Text = ObjAnalise.CustoArrendamento.ToString("N2")
        txtOutrosCustos.Text = ObjAnalise.OutrasDespesas.ToString("N2")
        txtTotalCustos.Text = ObjAnalise.TotalCustos.ToString("N2")

        txtSaldoAnalise.Text = ObjAnalise.SaldoCliente.ToString("N2")
        txtCreditoConcedido.Text = ObjAnalise.CreditoConcedido.ToString("N2")
    End Sub

    Private Sub FormatRatingDeCredito()
        Select Case ObjAnalise.Perguntas.RatingCreditoCliente
            Case "A"
                lblRatingCredito.Text = "| A |"
                txtPercRatingCredito.Text = ObjAnalise.ParametrosAnaliseDeCredito.PercLimiteCreditoA
                lblRatingCredito.BackColor = Drawing.Color.Green
                lblRatingCredito.ForeColor = Drawing.Color.White
            Case "B"
                lblRatingCredito.Text = "| B |"
                txtPercRatingCredito.Text = ObjAnalise.ParametrosAnaliseDeCredito.PercLimiteCreditoB
                lblRatingCredito.BackColor = Drawing.Color.Yellow
                lblRatingCredito.ForeColor = Drawing.Color.Black
            Case "C"
                lblRatingCredito.Text = "| C |"
                txtPercRatingCredito.Text = ObjAnalise.ParametrosAnaliseDeCredito.PercLimiteCreditoC
                lblRatingCredito.BackColor = Drawing.Color.Red
                lblRatingCredito.ForeColor = Drawing.Color.White
        End Select
        AtualizaResumo()
    End Sub

    Protected Sub ddlResposta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddlResposta As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddlResposta.NamingContainer, GridViewRow)

        SessaoRecuperaAnaliseDeCredito()
        VerificaIUD()
        ObjAnalise.Perguntas(row.RowIndex).Resposta = ddlResposta.SelectedValue
        FormatRatingDeCredito()
        SessaoSalvaAnaliseDeCredito()
    End Sub

    Protected Sub ddlRiscoCultura_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim ddlRiscoCultura As DropDownList = CType(sender, DropDownList)
        Dim row As GridViewRow = CType(ddlRiscoCultura.NamingContainer, GridViewRow)

        SessaoRecuperaAnaliseDeCredito()
        VerificaIUD()
        ObjAnalise.Culturas(row.RowIndex).RiscoCultura = ddlRiscoCultura.SelectedValue
        txtCoefRedutor.Text = ObjAnalise.Culturas.CoefRedutorRiscoCultura
        AtualizaResumo()
        SessaoSalvaAnaliseDeCredito()
    End Sub

    Protected Sub ImgExcluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim btn As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(btn.NamingContainer, GridViewRow)

        SessaoRecuperaAnaliseDeCredito()
        VerificaIUD()
        ObjAnalise.Clientes.Remove(ObjAnalise.Clientes(row.RowIndex))
        ObjAnalise.AtualizarAnalise()
        SessaoSalvaAnaliseDeCredito()
        AtualizaFormComAClasse()
    End Sub

    Public Sub Limpar()
        ddlAno.SelectedIndex = 0
        lblDefinicao.Text = ""
        ObjAnalise = New AnaliseDeCredito(New [Lib].Negocio.ParametrosAnaliseDeCredito)
        ObjAnalise.IUD = "I"
        SessaoSalvaAnaliseDeCredito()
        AtualizaFormComAClasse()
        Dim num As New [Lib].Negocio.Numerador(3)
        lblAnalise.Text = num.Sequencia + 1
        'VerificaPermissoes()
    End Sub

    Private Function getSql() As String
        Dim sql As String = ""
        sql &= "SELECT An.Analise_Id," & vbCrLf & _
               "       An.Ano_Id," & vbCrLf & _
               "       An.DefinicaoAno_Id," & vbCrLf & _
               "       An.Situacao," & vbCrLf & _
               "       dbo.ClientesAnalise(An.Analise_Id) as nome," & vbCrLf & _
               "       AnXCu.Plantios," & vbCrLf & _
               "       AnXCu.AreaPlantio," & vbCrLf & _
               "       An.AreaArrendada," & vbCrLf & _
               "       An.CustoArrendamento," & vbCrLf & _
               "       AnXCu.CustoPortifolio," & vbCrLf & _
               "       AnXCu.CustoCultura," & vbCrLf & _
               "       An.OutrasDespesas," & vbCrLf & _
               "       An.OutrasReceitas," & vbCrLf & _
               "       AnXCu.ReceitaCultura," & vbCrLf & _
               "       An.LimiteCredito," & vbCrLf & _
               "       case An.LimiteCredito" & vbCrLf & _
               "          when 'A' then PercLimiteCreditoA" & vbCrLf & _
               "          when 'B' then PercLimiteCreditoB" & vbCrLf & _
               "          when 'C' then PercLimiteCreditoC" & vbCrLf & _
               "       end PercLimCredito," & vbCrLf & _
               "       An.PercRedutorRiscoCultura," & vbCrLf & _
               "       (AnXCu.CustoPortifolio *(" & vbCrLf & _
               "								case An.LimiteCredito" & vbCrLf & _
               "								  when 'A' then PercLimiteCreditoA" & vbCrLf & _
               "								  when 'B' then PercLimiteCreditoB" & vbCrLf & _
               "								  when 'C' then PercLimiteCreditoC" & vbCrLf & _
               "							    end - An.PercRedutorRiscoCultura" & vbCrLf & _
               "                                )) / 100 as CreditoAnalise" & vbCrLf & _
               "  FROM AnaliseDeCredito An" & vbCrLf & _
               " Inner Join AnaliseDeCreditoParametro AnP" & vbCrLf & _
               "    on AnP.Ano_Id      = An.Ano_Id" & vbCrLf & _
               "   and AnP.DefinicaoAno_Id = An.DefinicaoAno_Id" & vbCrLf & _
               " Inner Join (" & vbCrLf & _
               " 			 Select Analise_Id," & vbCrLf & _
               "			 	    Ano_Id," & vbCrLf & _
               "				    DefinicaoAno_id," & vbCrLf & _
               "				    Count(*) as Plantios," & vbCrLf & _
               "				    Sum(AreaPlantio) as AreaPlantio," & vbCrLf & _
               "				    Sum(CustoPortifolio) as CustoPortifolio," & vbCrLf & _
               "				    Sum(CustoCultura) as CustoCultura," & vbCrLf & _
               "				    sum(ReceitaCultura) as ReceitaCultura" & vbCrLf & _
               "			   From AnaliseDeCreditoCultura" & vbCrLf & _
               "			  Group by Analise_Id, Ano_Id, DefinicaoAno_id" & vbCrLf & _
               "             ) AnXCu" & vbCrLf & _
               "      on An.Analise_Id      = AnXCu.Analise_Id" & vbCrLf & _
               "	 and An.Ano_Id          = AnXCu.Ano_Id" & vbCrLf & _
               "     and An.DefinicaoAno_id = AnXCu.DefinicaoAno_id" & vbCrLf & _
               "   Where 1 = 1" & vbCrLf

        If txtAnaliseConsulta.Text.Trim <> "" Then
            sql &= "     and An.Analise_Id = " & txtAnaliseConsulta.Text & vbCrLf
        End If

        If ddlAnoConsulta.SelectedIndex > 0 Then
            sql &= "     and An.Ano_Id = " & ddlAnoConsulta.SelectedValue & vbCrLf
        End If


        Dim Situacao As String = ""
        If chkNormal.Checked Then Situacao &= IIf(Situacao = "", "", ",") & "'NORMAL'"
        If chkCancelada.Checked Then Situacao &= IIf(Situacao = "", "", ",") & "'CANCELADA'"
        If chkAprovada.Checked Then Situacao &= IIf(Situacao = "", "", ",") & "'APROVADA'"
        If chkLiberada.Checked Then Situacao &= IIf(Situacao = "", "", ",") & "'LIBERADA'"
        If Situacao <> "" Then sql &= "     and An.Situacao in (" & Situacao & ")" & vbCrLf


        If chkUsarPeriodoAprovacao.Checked And chkUsarPeriodoLiberacao.Checked Then
            sql &= "     and (" & vbCrLf & _
                   "          An.DataAprovacao between '" & CDate(txtDataAprDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataAprAte.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "         or" & vbCrLf & _
                   "          An.DataLiberacao between '" & CDate(txtDataLibDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataLibAte.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   "          )" & vbCrLf
        ElseIf chkUsarPeriodoAprovacao.Checked Then
            sql &= "     and An.DataAprovacao between '" & CDate(txtDataAprDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataAprAte.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        ElseIf chkUsarPeriodoLiberacao.Checked Then
            sql &= "     and An.DataLiberacao between '" & CDate(txtDataLibDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataLibAte.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        If chkUsuarioAprovacao.Checked And chkUsuarioLiberacao.Checked Then
            sql &= "     and (" & vbCrLf & _
                   "          An.UsuarioAprovacao = '" & txtUsuario.Text & "'" & vbCrLf & _
                   "         or" & vbCrLf & _
                   "          An.UsuarioLiberacao = '" & txtUsuario.Text & "'" & vbCrLf & _
                   "          )" & vbCrLf
        ElseIf chkUsuarioAprovacao.Checked Then
            sql &= "     and An.UsuarioAprovacao = '" & txtUsuario.Text & "'" & vbCrLf
        ElseIf chkUsuarioLiberacao.Checked Then
            sql &= "     and An.UsuarioLiberacao = '" & txtUsuario.Text & "'" & vbCrLf
        End If

        If HCliente.Value.Length > 0 Then
            Dim cli As String()
            cli = HCliente.Value.Split("-")
            sql &= "     and exists (select 1 " & vbCrLf & _
                   "                   from AnaliseDeCreditoCliente AnC" & vbCrLf & _
                   "                  where AnC.Analise_Id = An.Analise_Id" & vbCrLf & _
                   "                    and AnC.Cliente_id = '" & cli(0) & "')" & vbCrLf
        End If
        Return sql
    End Function

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteAnCredCC", "txtNome")
    End Sub

    Protected Sub gridConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridConsulta.SelectedIndexChanged
        ObjAnalise = New AnaliseDeCredito(gridConsulta.SelectedRow.Cells(2).Text)
        ObjAnalise.IUD = "U"
        'ObjAnalise.AtualizarAnalise()
        SessaoSalvaAnaliseDeCredito()
        AtualizaFormComAClasse()
        tcAnaliseCredito.ActiveTabIndex = 0
    End Sub

    Protected Sub imgImpressaoPend_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim btnImprimir As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(btnImprimir.NamingContainer, GridViewRow)
        ExibirRelatorio(gridConsulta.Rows(Gridrow.RowIndex).Cells(2).Text, gridConsulta.Rows(Gridrow.RowIndex).Cells(3).Text, gridConsulta.Rows(Gridrow.RowIndex).Cells(4).Text)
    End Sub

    Public Sub ExibirRelatorio(ByVal Analise As Integer, ByVal Ano As Integer, ByVal Definicao As Integer)
        Dim sql As String
        '--********** Analise De Credito **************
        sql = "SELECT AC.Analise_Id," & vbCrLf & _
              "       AC.Ano_Id," & vbCrLf & _
              "       AC.DefinicaoAno_Id," & vbCrLf & _
              "       AC.DataAnalise," & vbCrLf & _
              "       AC.AreaArrendada," & vbCrLf & _
              "       AC.CustoArrendamento," & vbCrLf & _
              "       AC.OutrasReceitas," & vbCrLf & _
              "       AC.OutrasDespesas," & vbCrLf & _
              "       AC.LimiteCredito," & vbCrLf & _
              "       Case AC.LimiteCredito" & vbCrLf & _
              "         when 'A' then PercLimiteCreditoA" & vbCrLf & _
              "         when 'B' then PercLimiteCreditoB" & vbCrLf & _
              "         when 'C' then PercLimiteCreditoC" & vbCrLf & _
              "       end PercCredito," & vbCrLf & _
              "       AC.PercRedutorRiscoCultura," & vbCrLf & _
              "       AC.Situacao," & vbCrLf & _
              "       AC.UsuarioAprovacao,    AC.DataAprovacao," & vbCrLf & _
              "       AC.UsuarioLiberacao,    AC.DataLiberacao," & vbCrLf & _
              "       AC.UsuarioCancelamento, AC.DataCancelamento," & vbCrLf & _
              "       SB_Cultura.TotalReceitaCultura," & vbCrLf & _
              "       SB_Cultura.TotalCustoCultura," & vbCrLf & _
              "       SB_Cultura.TotalCustoPortifolio," & vbCrLf & _
              "       P.DataDefinicao" & vbCrLf & _
              "  FROM AnaliseDeCredito AC" & vbCrLf & _
              " INNER JOIN AnaliseDeCreditoParametro P" & vbCrLf & _
              "    on AC.Ano_Id          = P.Ano_Id" & vbCrLf & _
              "   and AC.DefinicaoAno_Id = P.DefinicaoAno_Id" & vbCrLf & _
              " INNER Join(" & vbCrLf & _
              "			SELECT Analise_Id," & vbCrLf & _
              "			       Ano_Id," & vbCrLf & _
              "			       DefinicaoAno_Id," & vbCrLf & _
              "			       Sum(ReceitaCultura) as TotalReceitaCultura," & vbCrLf & _
              "			       Sum(CustoCultura) as TotalCustoCultura," & vbCrLf & _
              "			       Sum(CustoPortifolio) as TotalCustoPortifolio" & vbCrLf & _
              "			  FROM AnaliseDeCreditoCultura" & vbCrLf & _
              "			 Group by Analise_Id," & vbCrLf & _
              "			          Ano_Id," & vbCrLf & _
              "			          DefinicaoAno_Id" & vbCrLf & _
              "            ) SB_Cultura" & vbCrLf & _
              "    on AC.Analise_id      = SB_Cultura.Analise_Id" & vbCrLf & _
              "   and AC.Ano_Id          = SB_Cultura.Ano_Id" & vbCrLf & _
              "   and AC.DefinicaoAno_Id = SB_Cultura.DefinicaoAno_Id" & vbCrLf & _
              " Where AC.Analise_Id      = " & Analise & vbCrLf & _
              "   and AC.Ano_Id          = " & Ano & vbCrLf & _
              "   And AC.DefinicaoAno_Id = " & Definicao & vbCrLf

        '--****************** CLIENTE *****************
        sql &= "SELECT ACC.Analise_Id, " & vbCrLf & _
               "       ACC.Ano_Id, " & vbCrLf & _
               "       ACC.DefinicaoAno_Id, " & vbCrLf & _
               "       ACC.Cliente_Id," & vbCrLf & _
               "       (Select top 1 C.Nome from Clientes C where C.Cliente_id = ACC.Cliente_Id) as NomeCliente" & vbCrLf & _
               "  FROM AnaliseDeCreditoCliente ACC  " & vbCrLf & _
               " Where ACC.Analise_Id      = " & Analise & vbCrLf & _
               "   and ACC.Ano_Id          = " & Ano & vbCrLf & _
               "   And ACC.DefinicaoAno_Id = " & Definicao & vbCrLf


        '--**** Analise De Credito Cultura ************
        sql &= "SELECT ACC.Analise_Id, " & vbCrLf & _
               "       ACC.Ano_Id, " & vbCrLf & _
               "       ACC.DefinicaoAno_Id, " & vbCrLf & _
               "       ACC.Safra_Id, " & vbCrLf & _
               "       ACC.Cultura_Id, " & vbCrLf & _
               "       C.Descricao as DescCultura," & vbCrLf & _
               "       ACC.AreaPlantio, " & vbCrLf & _
               "       case ACC.RiscoCultura" & vbCrLf & _
               "         when 3 then 'ALTO'" & vbCrLf & _
               "         when 2 then 'MEDIO'" & vbCrLf & _
               "         when 1 then 'BAIXO'" & vbCrLf & _
               "       end as RiscoCultura, " & vbCrLf & _
               "       ACC.CustoPortifolio, " & vbCrLf & _
               "       ACC.Producao, " & vbCrLf & _
               "       ACC.ReceitaCultura, " & vbCrLf & _
               "       ACC.CustoCultura" & vbCrLf & _
               "  FROM AnaliseDeCreditoCultura ACC" & vbCrLf & _
               " INNER JOIN Cultura C" & vbCrLf & _
               "    ON ACC.Cultura_Id = C.Cultura_Id" & vbCrLf & _
               " Where ACC.Analise_Id      = " & Analise & vbCrLf & _
               "   and ACC.Ano_Id          = " & Ano & vbCrLf & _
               "   And ACC.DefinicaoAno_Id = " & Definicao & vbCrLf

        '--**** Aalise De Credito Pergunta ***********
        sql &= "SELECT AC.Analise_Id, " & vbCrLf & _
               "       AC.Ano_Id, " & vbCrLf & _
               "       AC.DefinicaoAno_Id, " & vbCrLf & _
               "       AC.Pergunta_Id, " & vbCrLf & _
               "       P.Descricao AS DescPergunta," & vbCrLf & _
               "       case AC.Resposta" & vbCrLf & _
               "         when 3 then 'BOM'" & vbCrLf & _
               "         when 2 then 'REGULAR'" & vbCrLf & _
               "         when 1 then 'RUIM'" & vbCrLf & _
               "       end as Resposta " & vbCrLf & _
               "  FROM AnaliseDeCreditoPergunta AC" & vbCrLf & _
               " INNER JOIN AnaliseDeCreditoParametroPergunta P " & vbCrLf & _
               "    ON AC.Ano_Id          = P.Ano_Id " & vbCrLf & _
               "   AND AC.DefinicaoAno_Id = P.DefinicaoAno_Id " & vbCrLf & _
               "   AND AC.Pergunta_Id     = P.Pergunta_Id" & vbCrLf & _
               " Where AC.Analise_Id      = " & Analise & vbCrLf & _
               "   and AC.Ano_Id          = " & Ano & vbCrLf & _
               "   And AC.DefinicaoAno_Id = " & Definicao & vbCrLf

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Consulta")
        ds.Tables(0).TableName = "AnaliseDeCredito"
        ds.Tables(1).TableName = "Cliente"
        ds.Tables(2).TableName = "Cultura"
        ds.Tables(3).TableName = "Perguntas"

        'LOGOTIPO
        Dim dt As DataTable = ds.Tables.Add("Logotipo")
        dt.Columns.Add("path", GetType(String))
        dt.Columns.Add("Imagem", GetType(System.Byte()))
        dt.Columns.Add("Nome", GetType(String))
        dt.Columns.Add("Cidade", GetType(String))
        dt.Columns.Add("Estado_Id", GetType(String))
        Dim drImagem As DataRow = dt.NewRow()
        Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
        drImagem("path") = strCaminhoImagem
        drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        drImagem("Nome") = HttpContext.Current.Session("ssNomeEmpresa")
        drImagem("Cidade") = HttpContext.Current.Session("ssCidadeEmpresa")
        drImagem("Estado_Id") = HttpContext.Current.Session("ssEstadoEmpresa")
        dt.Rows.Add(drImagem)

        Dim rpt As New ReportDocument()
        rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_AnaliseDeCreditoProdutorRural.rpt")
        rpt.Load(rpt.FileName, OpenReportMethod.OpenReportByDefault)

        Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

        rpt.SetDataSource(ds)

        Dim parameters = New Dictionary(Of String, Object)()

        If Dir(NomeArquivo).Length > 0 Then
            Kill(NomeArquivo)
        End If

        Try
            rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)
            If System.IO.File.Exists(NomeArquivo) Then
                Funcoes.AbrirArquivo(Page, UrlArquivo)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(getSql() & " order by dbo.ClientesAnalise(An.Analise_Id)", "Consulta")
            ds.Tables(0).TableName = "AnaliseDeCredito"

            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            ds.Tables(0).Columns.Add("Imagem", GetType(System.Byte()))
            ds.Tables(0).Columns.Add("NomeEmp", GetType(String))
            ds.Tables(0).Columns.Add("CidadeEstado", GetType(String))

            For Each row As DataRow In ds.Tables(0).Rows
                row("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                row("NomeEmp") = HttpContext.Current.Session("ssNomeEmpresa") & " (" & Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("ssEmpresa")) & ")"
                row("CidadeEstado") = HttpContext.Current.Session("ssCidadeEmpresa") & "/" & HttpContext.Current.Session("ssEstadoEmpresa")
            Next

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", "Relatório Geral de Análise de Créditos Produtor Rural.")

            Funcoes.BindReport(Me.Page, ds, "Cr_AnaliseDeCreditos", eExportType.PDF, parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        txtAnaliseConsulta.Text = String.Empty
        txtClienteConsulta.Text = String.Empty

        chkAprovada.Checked = False
        chkCancelada.Checked = False
        chkLiberada.Checked = False
        chkNormal.Checked = False

        txtDataAprAte.Text = String.Empty
        txtDataAprDe.Text = String.Empty
        txtDataLibAte.Text = String.Empty
        txtDataLibDe.Text = String.Empty

        txtUsuario.Text = String.Empty
        chkUsuarioAprovacao.Checked = False
        chkUsuarioLiberacao.Checked = False
        chkUsarPeriodoAprovacao.Checked = False
        chkUsarPeriodoLiberacao.Checked = False


    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        SessaoRecuperaAnaliseDeCredito()
        'ObjAnalise.IUD = "I"
        ObjAnalise.DataAnalise = Now
        If ObjAnalise.Salvar Then
            MsgBox(Me.Page, "Análise Salva com Sucesso.", eTitulo.Sucess)
            Limpar()
        Else
            MsgBox(Me.Page, "Erro ao Salvar...")
        End If
    End Sub

    Protected Sub lnkCancelar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkCancelar.Click
        SessaoRecuperaAnaliseDeCredito()
        ObjAnalise.IUD = "C"
        ObjAnalise.Situacao = "CANCELADA"
        ObjAnalise.DataCancelamento = Now
        ObjAnalise.UsuarioCancelamento = Session("ssNomeUsuario")
        If ObjAnalise.Salvar Then
            MsgBox(Me.Page, "Análise Cancelada com Sucesso.", eTitulo.Sucess)
            Limpar()
        Else
            MsgBox(Me.Page, "Erro ao Cancelar...")
        End If
    End Sub

    Protected Sub lnkAprovarSimulacao_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAprovarSimulacao.Click
        If Funcoes.VerificaPermissao("AnaliseDeCreditoAprovar", "ACESSAR") Then

            SessaoRecuperaAnaliseDeCredito()
            If ObjAnalise.IUD = "I" Then
                ObjAnalise.DataAnalise = Now
                If Not ObjAnalise.Salvar Then
                    MsgBox(Me.Page, "Erro ao Salvar a Análise, não foi possível fazer a aprovação.")
                    Exit Sub
                End If
            End If

            ObjAnalise.IUD = "A"
            ObjAnalise.Situacao = "APROVADA"
            ObjAnalise.UsuarioAprovacao = Session("ssNomeUsuario")
            ObjAnalise.DataAprovacao = Date.Now

            If ObjAnalise.Salvar Then
                MsgBox(Me.Page, "Análise Aprovada com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, "Erro ao Aprovar a Simulação...")
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para Aprovar Crédito!!!")
        End If
    End Sub

    Protected Sub lnkLiberarCredito_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLiberarCredito.Click
        If Funcoes.VerificaPermissao("AnaliseDeCreditoLiberar", "ACESSAR") Then

            SessaoRecuperaAnaliseDeCredito()
            If ObjAnalise.IUD = "I" Then
                ObjAnalise.DataAnalise = Now
                ObjAnalise.Situacao = "NORMAL"
                If Not ObjAnalise.Salvar Then
                    MsgBox(Me.Page, "Erro ao Salvar a Análise, não foi possivel fazer a aprovação.")
                    Exit Sub
                End If
            End If

            '******************************************************
            '**************** Autorizar ***************************
            '******************************************************
            If ObjAnalise.IUD = "I" Or ObjAnalise.UsuarioAprovacao = "" Then
                ObjAnalise.IUD = "A"
                ObjAnalise.Situacao = "APROVADA"
                ObjAnalise.UsuarioAprovacao = Session("ssNomeUsuario")
                ObjAnalise.DataAprovacao = Date.Now

                If ObjAnalise.Salvar Then
                    MsgBox(Me.Page, "Análise Aprovada com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "Erro ao Aprovar a Simulacao...")
                    Exit Sub
                End If
            End If

            '******************************************************
            '**************** LIBERAR *****************************
            '******************************************************
            ObjAnalise.IUD = "L"
            ObjAnalise.Situacao = "LIBERADA"
            ObjAnalise.UsuarioLiberacao = Session("ssNomeUsuario")
            ObjAnalise.DataLiberacao = Date.Now

            If ObjAnalise.Salvar Then
                MsgBox(Me.Page, "Análise Liberada com Sucesso. Todas as outras análises desses clientes feitas para este ano safra foram canceladas automaticamente.")
            Else
                MsgBox(Me.Page, "Erro ao Liberar a Simulação...")
                Exit Sub
            End If

        Else
            MsgBox(Me.Page, "Usuário sem permissão para Liberar Crédito!!!")
        End If
    End Sub

    Protected Sub lnkLimparAnalise_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparAnalise.Click
        Limpar()
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(getSql, "Consulta")
        gridConsulta.DataSource = ds
        gridConsulta.DataBind()
    End Sub

    Protected Sub lnkAjudaB_Click(sender As Object, e As EventArgs) Handles lnkAjudaB.Click
        Try
            Funcoes.Ajuda(Me.Page, "AnaliseDeCreditoProdutorRural")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AnaliseDeCreditoProdutorRural")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class