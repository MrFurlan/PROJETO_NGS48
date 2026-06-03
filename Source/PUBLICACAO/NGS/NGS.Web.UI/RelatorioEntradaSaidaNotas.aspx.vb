Imports System.IO
Imports System.Data
Imports System.Data.SqlClient
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelatorioEntradaSaidaNotas
    Inherits BasePage

#Region "Variáveis"
    Dim ListGrupo As [Lib].Negocio.ListGrupoProduto
    Dim Parametros As String = String.Empty
#End Region

#Region "SESSÃO"
    Private Sub SessaoSalvarGrupo()
        Session("Grupo" & HID.Value) = ListGrupo
    End Sub

    Private Sub SessaoRecuperaGrupo()
        ListGrupo = Session("Grupo" & HID.Value)
    End Sub
#End Region

#Region "Procedimentos"

    Private Sub CarregarSituacao()
        ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao)
        ddlSituacao.SelectedValue = "1" 'NORMAL
    End Sub

    Private Sub CarregarTipoDeDocumento(ByVal todos As Boolean)
        chkTipoDeDocumento.Items.Clear()
        chkTipoDeDocumento.DataValueField = "Codigo"
        chkTipoDeDocumento.DataTextField = "Descricao"
        Dim lst As New [Lib].Negocio.ListTipoDeDocumento()
        chkTipoDeDocumento.DataSource = lst.ToArray()
        chkTipoDeDocumento.DataBind()
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If todos Then
                chkTipoDeDocumento.Items(i).Selected = True
            ElseIf chkTipoDeDocumento.Items(i).Value = 1 Then
                chkTipoDeDocumento.Items(i).Selected = True
            End If
        Next
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub CarregarPlanoDeCusto()
        Dim ListPC As New [Lib].Negocio.ListPlanoDeCusto()

        Dim j As Integer = 0
        While j < ListPC.Count
            LstPlanoDeCusto.Items.Add(New ListItem(Format(ListPC(j).Codigo, "000") & "-" & ListPC(j).Descricao, ListPC(j).Codigo))
            j += 1
        End While
    End Sub

    Public Function ValidarCampos() As Boolean
        If txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida.", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        If txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida.", eTitulo.Info)
            txtDataFinal.Focus()
            Return False
        End If

        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        If lstEncargos.GetSelectedValues().Count > 0 AndAlso lstEncargos.GetSelectedValues().Count <> 5 Then
            MsgBox(Me.Page, "Informe exatamente 5 Encargos, para consulta.")
            Return False
        End If

        Return True
    End Function

    Protected Sub chkAllTipos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If chkAllTipos.Checked Then
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
                    chkTipoDeDocumento.Items(i).Selected = True
                Next
            Else
                For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1

                    If chkTipoDeDocumento.Items(i).Value = 1 Then
                        chkTipoDeDocumento.Items(i).Selected = True
                    Else
                        chkTipoDeDocumento.Items(i).Selected = False
                    End If
                Next
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function BuscaSQLNotasFisico(ByVal PC As String) As String
        Dim sql As String
        sql = "SELECT NF.Movimento," & vbCrLf &
              "       NF.Pedido," & vbCrLf &
              "       CONVERT(varchar, R.Operacao) + '-' + CONVERT(varchar, R.SubOperacao) AS Operacao," & vbCrLf &
              "       NF.Empresa_Id," & vbCrLf &
              "       NF.EndEmpresa_Id," & vbCrLf &
              "       R.PesoLiquido AS Liquido," & vbCrLf &
              "       R.PesoBruto AS BrutoBalanca," & vbCrLf &
              "       P.Produto_Id," & vbCrLf &
              "       P.Nome," & vbCrLf &
              "       C.Reduzido," & vbCrLf &
              "       C.Cliente_Id," & vbCrLf &
              "       C.Endereco_Id," & vbCrLf &
              "       C.Nome AS NomeCliente," & vbCrLf &
              "       ISNULL(Analises.PercentualUmidade, 0)     AS PercUmidade," & vbCrLf &
              "       ISNULL(Analises.DescontoUmidade, 0)       AS DescUmidade," & vbCrLf &
              "       ISNULL(Analises.PercentualImpureza, 0)    AS PercImpureza," & vbCrLf &
              "       ISNULL(Analises.DescontoImpureza, 0)      AS DescImpureza," & vbCrLf &
              "       ISNULL(Analises.PercentualAvariados, 0)   AS PercAveriados," & vbCrLf &
              "       ISNULL(Analises.DescontoAvariados, 0)     AS DescAveriados," & vbCrLf &
              "       ISNULL(Analises.PercentualPH, 0)          AS PercPH," & vbCrLf &
              "       ISNULL(Analises.DescontoPH, 0)            AS DescPH," & vbCrLf &
              "       ISNULL(Analises.PercentualVerdoso, 0)     AS PercVerdoso," & vbCrLf &
              "       ISNULL(Analises.DescontoVerdoso, 0)       AS DescVerdoso," & vbCrLf &
              "       ISNULL(Analises.PercentualTransgenico, 0) AS PercTrans," & vbCrLf &
              "       ISNULL(Analises.DescontoTransgenico, 0)   AS DescTrans," & vbCrLf &
              "       ISNULL(Analises.PercentualOutros, 0)      AS PercOutros," & vbCrLf &
              "       ISNULL(Analises.DescontoOutros, 0)        AS DescOutros," & vbCrLf &
              "       D.Cliente_Id AS CnpjDeposito," & vbCrLf &
              "       D.Endereco_Id AS EnderecoDeposito," & vbCrLf &
              "       D.Fantasia AS FantasiaDeposito," & vbCrLf &
              "       D.Reduzido AS ReduzidoDeposito," & vbCrLf &
              "       D.Cidade AS CidadeDeposito," & vbCrLf &
              "       D.Estado AS EstadoDeposito," & vbCrLf &
              "       E.Nome AS EmpresaNome," & vbCrLf &
              "       E.Reduzido AS EmpresaReduzido," & vbCrLf &
              "       NF.Nota_Id," & vbCrLf &
              "       C.Cidade AS CidadeCliente," & vbCrLf &
              "       C.Estado AS EstadoCliente," & vbCrLf &
              "       E.Cidade AS CidadeEmpresa," & vbCrLf &
              "       E.Estado AS EstadoEmpresa," & vbCrLf &
              "       NFI.PesoFiscal," & vbCrLf &
              "       RomaneiosXPesagens.Pesagem_Id AS Laudo," & vbCrLf &
              "       ISNULL(LP.Placa, '') AS Placa, " & vbCrLf &
              "       NFI.Valor, " & vbCrLf &
              "       Pe.Pedido_Id, " & vbCrLf &
              "       NFER.ChaveNfe, " & vbCrLf &
              "       NF.Representante " & vbCrLf &              
              "  FROM NotasFiscais AS NF WITH (NOLOCK)" & vbCrLf &
              " INNER JOIN NotasFiscaisXItens AS NFI WITH (NOLOCK)" & vbCrLf &
              "    ON NF.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFI.Serie_Id " & vbCrLf &
              "   AND NF.Nota_Id         = NFI.Nota_Id" & vbCrLf &
              "  LEFT JOIN Operacoes AS OP WITH (NOLOCK)" & vbCrLf &
              "    ON OP.Operacao_Id     = NF.Operacao " & vbCrLf &
              "  LEFT JOIN SubOperacoes AS SO WITH (NOLOCK)" & vbCrLf &
              "    ON SO.Operacao_Id     = NF.Operacao " & vbCrLf &
              "   AND SO.SubOperacoes_Id = NF.SubOperacao " & vbCrLf &
              " Inner Join OperacaoXEstado OE" & vbCrLf &
              "    on OE.Codigo_id = NFI.OperacaoXEstado" & vbCrLf &
              " Inner Join NotasFiscaisxEncargos NFxE WITH (NOLOCK)" & vbCrLf &
              "    ON NFI.Empresa_Id      = NFxE.Empresa_Id" & vbCrLf &
              "   AND NFI.EndEmpresa_Id   = NFxE.EndEmpresa_Id" & vbCrLf &
              "   AND NFI.Cliente_Id      = NFxE.Cliente_Id" & vbCrLf &
              "   AND NFI.EndCliente_Id   = NFxE.EndCliente_Id" & vbCrLf &
              "   AND NFI.EntradaSaida_Id = NFxE.EntradaSaida_Id" & vbCrLf &
              "   AND NFI.Serie_Id        = NFxE.Serie_Id" & vbCrLf &
              "   AND NFI.Nota_Id         = NFxE.Nota_Id  " & vbCrLf &
              "   AND NFI.Produto_Id      = NFxE.Produto_Id" & vbCrLf &
              "   AND NFI.Sequencia_Id    = NFxE.Sequencia_Id" & vbCrLf &
              "   AND NFxE.Encargo_Id     = 'PRODUTO'" & vbCrLf &
              "  LEFT Join NFERealizadas NFER WITH (NOLOCK)" & vbCrLf &
              "    ON NF.Empresa_Id      = NFER.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFER.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFER.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFER.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = NFER.Nota_Id" & vbCrLf &
              "  LEFT Join Pedidos Pe WITH (NOLOCK)" & vbCrLf &
              "    on NF.Empresa_Id    = Pe.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id = Pe.EndEmpresa_Id" & vbCrLf &
              "   and NF.Pedido        = Pe.Pedido_Id" & vbCrLf &
              "  LEFT Join PedidoxTabelaDeComissao PTC WITH (NOLOCK)" & vbCrLf &
              "    ON Pe.Empresa_Id    = PTC.Empresa_Id " & vbCrLf &
              "   And Pe.EndEmpresa_Id = PTC.EndEmpresa_Id " & vbCrLf &
              "   And Pe.Pedido_Id     = PTC.Pedido_Id " & vbCrLf &
              "   AND NFI.Produto_id    = PTC.Produto_Id " & vbCrLf &
              " INNER JOIN NotasFiscaisXRomaneios AS NFR WITH (NOLOCK)" & vbCrLf &
              "    ON NFR.Empresa_Id      = NFI.Empresa_Id" & vbCrLf &
              "   AND NFR.EndEmpresa_Id   = NFI.EndEmpresa_Id" & vbCrLf &
              "   AND NFR.Cliente_Id      = NFI.Cliente_Id" & vbCrLf &
              "   AND NFR.EndCliente_Id   = NFI.EndCliente_Id" & vbCrLf &
              "   AND NFR.EntradaSaida_Id = NFI.EntradaSaida_Id" & vbCrLf &
              "   AND NFR.Serie_Id        = NFI.Serie_Id" & vbCrLf &
              "   AND NFR.Nota_Id         = NFI.Nota_Id" & vbCrLf &
              " INNER JOIN Romaneios AS R WITH (NOLOCK)" & vbCrLf &
              "    ON R.Empresa_Id    = NFR.Empresa_Id" & vbCrLf &
              "   AND R.EndEmpresa_Id = NFR.EndEmpresa_Id" & vbCrLf &
              "   AND R.Romaneio_Id   = NFR.Romaneio_Id" & vbCrLf &
              " INNER JOIN Produtos AS P WITH (NOLOCK)" & vbCrLf &
              "    ON NFI.Produto_Id = P.Produto_Id" & vbCrLf &
              " INNER JOIN Clientes AS C WITH (NOLOCK)" & vbCrLf &
              "    ON C.Cliente_Id  = NF.Cliente_Id" & vbCrLf &
              "   AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf &
              " INNER JOIN Clientes AS D WITH (NOLOCK)" & vbCrLf &
              "    ON D.Cliente_Id  = NFI.Deposito" & vbCrLf &
              "   AND D.Endereco_Id = NFI.EndDeposito" & vbCrLf &
              " INNER JOIN Clientes AS E WITH (NOLOCK)" & vbCrLf &
              "    ON E.Cliente_Id  = NFI.Empresa_Id" & vbCrLf &
              "   AND E.Endereco_Id = NFI.EndEmpresa_Id" & vbCrLf &
              " LEFT JOIN RomaneiosXPesagens WITH (NOLOCK)" & vbCrLf &
              "    ON R.Empresa_Id    = RomaneiosXPesagens.Empresa_Id " & vbCrLf &
              "   AND R.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id " & vbCrLf &
              "   AND R.Romaneio_Id   = RomaneiosXPesagens.Romaneio_Id" & vbCrLf &
              " LEFT JOIN Pesagem LP WITH (NOLOCK)" & vbCrLf &
              "    ON RomaneiosXPesagens.Empresa_Id         = LP.Empresa_Id " & vbCrLf &
              "   AND RomaneiosXPesagens.EndEmpresa_Id      = LP.EndEmpresa_Id " & vbCrLf &
              "   AND RomaneiosXPesagens.Pesagem_Id         = LP.Pesagem_Id" & vbCrLf &
              "  LEFT JOIN (SELECT Romaneio_ID, Empresa_Id, EndEmpresa_Id," & vbCrLf &
              "                    sum(case When Analise_Id = 1 then Percentual else 0 end) PercentualUmidade," & vbCrLf &
              "                    sum(case When Analise_Id = 1 then Desconto   else 0 end) DescontoUmidade," & vbCrLf &
              "                    sum(case When Analise_Id = 2 then Percentual else 0 end) PercentualImpureza," & vbCrLf &
              "                    sum(case When Analise_Id = 2 then Desconto   else 0 end) DescontoImpureza," & vbCrLf &
              "                    sum(case When Analise_Id = 3 then Percentual else 0 end) PercentualAvariados," & vbCrLf &
              "                    sum(case When Analise_Id = 3 then Desconto   else 0 end) DescontoAvariados," & vbCrLf &
              "                    sum(case When Analise_Id = 4 then Percentual else 0 end) PercentualPH," & vbCrLf &
              "                    sum(case When Analise_Id = 4 then Desconto   else 0 end) DescontoPH," & vbCrLf &
              "                    sum(case When Analise_Id = 5 then Percentual else 0 end) PercentualVerdoso," & vbCrLf &
              "                    sum(case When Analise_Id = 5 then Desconto   else 0 end) DescontoVerdoso," & vbCrLf &
              "                    sum(case When Analise_Id = 6 then Percentual else 0 end) PercentualTransgenico," & vbCrLf &
              "                    sum(case When Analise_Id = 6 then Desconto   else 0 end) DescontoTransgenico," & vbCrLf &
              "                    avg(case When Analise_Id > 6 then Percentual else 0 end) PercentualOutros," & vbCrLf &
              "                    sum(case When Analise_Id > 6 then Desconto   else 0 end) DescontoOutros" & vbCrLf &
              "               FROM RomaneiosXDescontos WITH (NOLOCK)" & vbCrLf &
              "              Group by Romaneio_ID, Empresa_Id, EndEmpresa_Id" & vbCrLf &
              "             ) AS Analises" & vbCrLf &
              "    ON Analises.Romaneio_ID   = R.Romaneio_Id" & vbCrLf &
              "   AND Analises.Empresa_Id    = R.Empresa_Id" & vbCrLf &
              "   AND Analises.EndEmpresa_Id = R.EndEmpresa_Id" & vbCrLf

        If txtCodigoTransportador.Value.Length > 0 Then
            Dim TRA As String() = txtCodigoTransportador.Value.Split(New Char() {"-"})
            sql &= " INNER JOIN NotasFiscaisXTransportadores AS nfxT WITH (NOLOCK)" & vbCrLf & _
                   "    ON nf.Empresa_Id        = nfxT.Empresa_Id" & vbCrLf & _
                   "   AND nf.EndEmpresa_Id     = nfxT.EndEmpresa_Id" & vbCrLf & _
                   "   AND nf.Cliente_Id        = nfxT.Cliente_Id" & vbCrLf & _
                   "   AND nf.EndCliente_Id     = nfxT.EndCliente_Id" & vbCrLf & _
                   "   AND nf.EntradaSaida_Id   = nfxT.EntradaSaida_Id" & vbCrLf & _
                   "   AND nf.Serie_Id          = nfxT.Serie_Id" & vbCrLf & _
                   "   AND nf.Nota_Id           = nfxT.Nota_Id" & vbCrLf & _
                   "   AND nfxT.Proprietario    ='" & TRA(0) & "'" & vbCrLf & _
                   "   AND nfxT.EndProprietario = " & TRA(1)
        End If

        sql &= " WHERE NF.Situacao = " & ddlSituacao.SelectedValue & " " & vbCrLf

        If ddlTroca.SelectedIndex = 1 Then
            sql &= "   and isnull(Pe.Troca,0) = 1"
        ElseIf ddlTroca.SelectedIndex = 2 Then
            sql &= "   and isnull(Pe.Troca,0) <> 1"
        End If

        If ddlNfTroca.SelectedIndex = 1 Then
            sql &= "   and isnull(NF.Troca,0) = 1"
        ElseIf ddlNfTroca.SelectedIndex = 2 Then
            sql &= "   and isnull(NF.Troca,0) <> 1"
        End If

        If ddlTipoFrete.SelectedIndex <> 0 Then
            sql &= "   AND isnull(NF.CIFFOB,'NEN') = '" & ddlTipoFrete.SelectedValue & "'" & vbCrLf
        End If

        If chkPessoaJuridica.Checked Then
            sql &= " and len(C.Cliente_ID) <> 14"
        End If

        If chkPessoaFisica.Checked Then
            sql &= " and len(C.Cliente_ID) <> 11 "
        End If

        If chkPessoaExterior.Checked Then
            sql &= " and C.Estado <> 'EX'"
        End If


        If txtSerie.Text.Length > 0 Then sql &= " and NF.Serie_Id ='" & txtSerie.Text & "'"

        If ddlStICMS.SelectedIndex > 0 Then sql &= " and isnull(OE.STICMS,0) = " & ddlStICMS.SelectedValue
        If ddlStIPI.SelectedIndex > 0 Then sql &= " and isnull(OE.STIPI,0) = " & ddlStIPI.SelectedValue
        If ddlStPISCOFINS.SelectedIndex > 0 Then sql &= " and isnull(OE.STPISCOFINS,0) = " & ddlStPISCOFINS.SelectedValue


        If PC.Length > 0 Then sql &= " And SO.ApuracaoDeCustos in (" & PC & ")"
        If txtPedido.Text.Trim.Length > 0 Then sql &= "AND PE.Pedido_Id        = " & txtPedido.Text & " " & vbCrLf
        If txtCN.Text.Trim.Length > 0 Then sql &= "AND PE.PedidoEfetivo ='" & txtCN.Text & "' " & vbCrLf
        If txtContrato.Text.Trim.Length > 0 Then sql &= "AND PE.Contrato      ='" & txtContrato.Text & "' " & vbCrLf

        If chkPeriodo.Checked And txtPedido.Text.Trim.Length = 0 And txtCN.Text.Trim.Length = 0 And txtContrato.Text.Trim.Length = 0 Then
            sql &= " AND NF.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
        End If

        If chkEntrada.Checked And Not chkSaida.Checked Then
            sql &= " AND NF.EntradaSaida_Id = 'E' " & vbCrLf
        ElseIf chkSaida.Checked And Not chkEntrada.Checked Then
            sql &= " AND NF.EntradaSaida_Id = 'S' " & vbCrLf
        End If

        If cmbEmpresa.SelectedIndex > 0 And (txtPedido.Text.Trim.Length = 0 Or txtCN.Text.Trim.Length = 0 Or txtContrato.Text.Trim.Length = 0) Then
            Dim strEmpresa As String()
            strEmpresa = cmbEmpresa.SelectedValue.Split(New Char() {"-"})
            If chkConsolidarEmpresa.Checked Then
                sql &= " AND left(NF.Empresa_Id,8) ='" & strEmpresa(0).Substring(0, 8) & "'"
            Else
                sql &= " AND NF.Empresa_Id    ='" & strEmpresa(0) & "' " & vbCrLf & _
                       " AND NF.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
            End If

        End If

        If txtCodigoCliente.Value.Length > 0 And (txtPedido.Text.Trim.Length = 0 Or txtCN.Text.Trim.Length = 0 Or txtContrato.Text.Trim.Length = 0) Then
            Dim strCliente As String()
            strCliente = txtCodigoCliente.Value.Split(New Char() {"-"})
            If chkConsolidarCliente.Checked Then
                sql &= " AND Left(NF.Cliente_Id,8) = '" & strCliente(0).Substring(0, 8) & "' " & vbCrLf
            Else
                sql &= " AND NF.Cliente_Id = '" & strCliente(0) & "'" & vbCrLf & _
                       " AND NF.EndCliente_Id = " & strCliente(1) & vbCrLf
            End If
        End If

        If txtCodigoDeposito.Value.Length > 0 Then
            Dim strDeposito As String()
            strDeposito = txtCodigoDeposito.Value.Split(New Char() {"-"})
            sql &= " AND NFI.Deposito    ='" & strDeposito(0) & "'" & vbCrLf & _
                   " AND NFI.EndDeposito = " & strDeposito(1) & vbCrLf
        End If

        If txtCodigoRepresentante.Value.Length > 0 Then
            Dim strRepresentante As String()
            strRepresentante = txtCodigoRepresentante.Value.Split(New Char() {"-"})
            sql &= " AND NF.Representante    ='" & strRepresentante(0) & "'" & vbCrLf &
                   " AND NF.EndRepresentante = " & strRepresentante(1) & vbCrLf
        End If

        If HOrigemDestino.Value.Length > 0 Then
            Dim strOrigemDestino As String()
            strOrigemDestino = HOrigemDestino.Value.Split("-")
            sql &= " AND NF.Destino    ='" & strOrigemDestino(0) & "'" & vbCrLf & _
                   " AND NF.EndDestino = " & strOrigemDestino(1) & vbCrLf
        End If

        VerificarGrupoProduto(sql)

        If cmbOperacao.SelectedIndex > 0 Then sql &= "AND OP.Operacao_id = " & cmbOperacao.SelectedValue & " " & vbCrLf
        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            sql &= "AND SO.SubOperacoes_ID = " & strSubOpe(1) & " " & vbCrLf
        End If

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            sql &= "   AND op.classe " & IIf(rdComClasse.Checked, "", "not") & " in ('" & op & "')"
        End If

        op = String.Join("','", lstFinalidade.GetSelecteds)
        If op.Length > 0 Then
            sql &= "   AND NF.Finalidade " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
        End If

        If RadSerieNao.Checked = True Then sql &= "AND NF.Serie_Id not in('D','F') " & vbCrLf
        If cmbSafra.SelectedIndex > 0 Then sql &= "AND PE.Safra = '" & cmbSafra.SelectedValue & "' " & vbCrLf

        If ddlMarca.SelectedIndex > 0 Then sql &= " And P.Marca =" & ddlMarca.SelectedValue

        If chkCusto.Checked Then
            sql &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) "
        End If

        If txtPlaca.Text.Length > 0 Then
            sql &= " AND LP.Placa = '" & txtPlaca.Text & "' " & vbCrLf
        End If

        If lstCfopSelecionados.Items.Count > 0 Then
            sql &= "AND ("
            Dim strOr As String = ""

            Dim k As Integer = 0
            While k < lstCfopSelecionados.Items.Count
                sql &= strOr & "OE.CodigoFiscal = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                strOr = " OR "
                k += 1
            End While

            sql &= ") " & vbCrLf
        End If

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If tp <> "" Then
            sql &= " AND NF.TipoDeDocumento in(" & tp & ")" & vbCrLf
        End If



        Dim strClienteRepresentante() As String = txtCodigoCliRepre.Value.Split("-")
        If strClienteRepresentante(0).Length > 0 Then
            Parametros &= "Representante: " & txtClienteRepresentante.Text & vbCrLf
            sql &= " AND PTC.Representante_Id = '" & strClienteRepresentante(0) & "'" & vbCrLf
            sql &= " AND PTC.EndRepresentante_Id = " & strClienteRepresentante(1) & " " & vbCrLf
        End If

        If txtSerie.Text.Length > 0 Then
            sql &= "AND NF.Serie_Id = '" & txtSerie.Text & "'" & vbCrLf
        End If

        If txtNumeroNota.Text.Length > 0 Then
            sql &= "AND NF.Nota_Id = '" & txtNumeroNota.Text & "'" & vbCrLf
        End If

        If RdNota.Checked Then
            sql &= "Order by NF.Movimento, NF.Empresa_Id, NF.EndEmpresa_Id, NF.Nota_Id, C.Nome"
        Else
            sql &= "Order by NF.Movimento, NF.Empresa_Id, NF.EndEmpresa_Id, C.Nome, NF.Nota_Id"
        End If

        Return sql

    End Function

    Private Function BuscaSQLResumoProduto(ByVal CFOP As Boolean)
        Dim sql As String

        sql = "/* RESUMO */ " & vbCrLf

        If CFOP Then
            sql &= " SELECT NF.EntradaSaida_Id AS EntradaSaida, NF.CFOP, " & vbCrLf & _
              "        NF.Produto, NF.NomeProduto,  " & vbCrLf & _
              "        SUM(NF.PesoBruto) AS PesoBruto, " & vbCrLf & _
              "        SUM(NF.Desconto) AS Desconto, " & vbCrLf & _
              "        SUM(NF.PesoLiquido) AS Liquido," & vbCrLf & _
              "        SUM(NF.Variacao) AS Variacao," & vbCrLf & _
              "        SUM(NF.Valor) AS Valor," & vbCrLf & _
              "        SUM(NF.ValorEncargo1) AS ValorEncargo1, " & vbCrLf & _
              "        SUM(NF.ValorEncargo2) AS ValorEncargo2," & vbCrLf & _
              "        SUM(NF.ValorEncargo3) AS ValorEncargo3," & vbCrLf & _
              "        SUM (NF.ValorEncargo4) AS ValorEncargo4," & vbCrLf & _
              "        SUM (NF.ValorEncargo5) AS ValorEncargo5," & vbCrLf & _
              "        SUM (NF.ValorLiquido) AS ValorLiquido " & vbCrLf & _
              "   FROM #Notas NF" & vbCrLf & _
              "  GROUP BY NF.Produto, NF.NomeProduto, NF.EntradaSaida_Id, NF.CFOP" & vbCrLf & _
              "  ORDER BY EntradaSaida, NF.CFOP, NF.NomeProduto" & vbCrLf
        Else
            sql &= " SELECT NF.EntradaSaida_Id AS EntradaSaida, " & vbCrLf & _
              "        NF.Produto, NF.NomeProduto,  " & vbCrLf & _
              "        SUM(NF.PesoBruto) AS PesoBruto, " & vbCrLf & _
              "        SUM(NF.Desconto) AS Desconto, " & vbCrLf & _
              "        SUM(NF.PesoLiquido) AS Liquido," & vbCrLf & _
              "        SUM(NF.Variacao) AS Variacao," & vbCrLf & _
              "        SUM(NF.Valor) AS Valor," & vbCrLf & _
              "        SUM(NF.ValorEncargo1) AS ValorEncargo1, " & vbCrLf & _
              "        SUM(NF.ValorEncargo2) AS ValorEncargo2," & vbCrLf & _
              "        SUM(NF.ValorEncargo3) AS ValorEncargo3," & vbCrLf & _
              "        SUM (NF.ValorEncargo4) AS ValorEncargo4," & vbCrLf & _
              "        SUM (NF.ValorEncargo5) AS ValorEncargo5," & vbCrLf & _
              "        SUM (NF.ValorLiquido) AS ValorLiquido " & vbCrLf & _
              "   FROM #Notas NF" & vbCrLf & _
              "  GROUP BY NF.Produto, NF.NomeProduto, NF.EntradaSaida_Id " & vbCrLf & _
              "  ORDER BY EntradaSaida, NF.NomeProduto" & vbCrLf
        End If
        Return sql
    End Function

    Function BuscaSQLNotasFiscal(ByVal pc As String, ByVal Encargos As List(Of String)) As String
        Dim dsRelatorio As New DataSet
        Dim Sql As String

        Sql = "SELECT NFE.Empresa_Id, NFE.EndEmpresa_Id, NFE.Cliente_Id, NFE.EndCliente_Id, NFE.EntradaSaida_Id, NFE.Serie_Id, NFE.Nota_Id, NFE.Produto_Id, NFE.Sequencia_Id," & vbCrLf &
               "                    sum(case when (NFE.Encargo_Id ='" & Encargos(0) & "' OR NFE.Encargo_Id = 'ICMS A REC.') THEN isnull(NFE.Valor, 0) else 0 end) AS Encargo1," & vbCrLf &
               "                    sum(case when NFE.Encargo_Id ='" & Encargos(1) & "' THEN isnull(NFE.Valor, 0) else 0 end) AS Encargo2," & vbCrLf &
               "                    sum(case when NFE.Encargo_Id ='" & Encargos(2) & "' THEN isnull(NFE.Valor, 0) else 0 end) AS Encargo3," & vbCrLf &
               "                    sum(case when NFE.Encargo_Id ='" & Encargos(3) & "' THEN isnull(NFE.Valor, 0) else 0 end) AS Encargo4," & vbCrLf &
               "                    sum(case when NFE.Encargo_Id ='" & Encargos(4) & "' THEN isnull(NFE.Valor, 0) else 0 end) AS Encargo5," & vbCrLf &
               "                    sum(case when NFE.Encargo_Id ='LIQUIDO'             THEN isnull(NFE.Valor, 0) else 0 end) AS LIQUIDO" & vbCrLf &
               "               INTO #Encargos " & vbCrLf &
               "               FROM NotasFiscaisXEncargos NFE WITH (NOLOCK)" & vbCrLf &
               "              INNER JOIN NotasFiscais AS NF " & vbCrLf &
               "                 ON NFE.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
               "                AND NFE.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
               "                AND NFE.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
               "                AND NFE.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
               "                AND NFE.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
               "                AND NFE.Serie_Id        = NF.Serie_Id " & vbCrLf &
               "                AND NFE.Nota_Id         = NF.Nota_Id" & vbCrLf &
               "              WHERE NFE.Encargo_Id in ('LIQUIDO', 'PRODUTO','" & Encargos(0) & "','" & Encargos(1) & "','" & Encargos(2) & "','" & Encargos(3) & "','" & Encargos(4) & "','ICMS A REC.')" & vbCrLf
       
        'Parâmetros somente para as Notas Fiscais
        GetSqlWhereNF(Sql)
        Sql &= "              GROUP BY NFE.Empresa_Id, NFE.EndEmpresa_Id, NFE.Cliente_Id, NFE.EndCliente_Id, NFE.EntradaSaida_Id, NFE.Serie_Id, NFE.Nota_Id, NFE.Produto_Id, NFE.Sequencia_Id" & vbCrLf


        Sql &= "SELECT E.Reduzido AS EmpresaReduzido, " & vbCrLf &
              "       NF.Empresa_Id, " & vbCrLf &
              "       NF.EndEmpresa_Id," & vbCrLf &
              "       E.Nome AS EmpresaNome, " & vbCrLf &
              "       E.Cidade AS EmpresaCidade, " & vbCrLf &
              "       E.Estado AS EmpresaEstado, " & vbCrLf &
              "       NF.Movimento, " & vbCrLf &
              "       NF.DataDaNota AS DataNota," & vbCrLf &
              "       NF.Nota_Id, " & vbCrLf &
              "       NF.Serie_Id, " & vbCrLf &
              "       NF.EntradaSaida_Id, " & vbCrLf &
              "       NFI.Produto_Id AS Produto," & vbCrLf &
              "       P.NCM," & vbCrLf &
              "       P.Nome AS NomeProduto, " & vbCrLf &
              "       Pg.Descricao as CondicaoPagamento, " & vbCrLf &
              "       OE.CodigoFiscal AS CFOP, " & vbCrLf &
              "       NF.Operacao, " & vbCrLf &
              "       NF.SubOperacao," & vbCrLf &
              "       ISNULL(NF.Finalidade, '') AS Finalidade, " & vbCrLf &
              "       NF.Cliente_Id, " & vbCrLf &
              "       NF.EndCliente_Id, " & vbCrLf &
              "       ISNULL(CLI.Nome, '') AS NomeCliente," & vbCrLf &
              "       CLI.Cidade AS CidadeCliente, " & vbCrLf &
              "       CLI.Estado AS EstadoCliente, " & vbCrLf &
              "       COALESCE (DO.Reduzido, N'') AS Deposito, " & vbCrLf &
              "       COALESCE (DD.Reduzido, N'') AS DepositoDestino, " & vbCrLf &
              "       case " & vbCrLf &
              "		    when isnull(P.ControlarRomaneio,0) = 1 " & vbCrLf &
              "			    then ISNULL(R.PesoBruto, 0) " & vbCrLf &
              "			    else NFI.QuantidadeFisica " & vbCrLf &
              "		    end PesoBruto, " & vbCrLf &
              "       case " & vbCrLf &
              "		    when isnull(P.ControlarRomaneio,0) = 1 " & vbCrLf &
              "			    then ISNULL(R.Desconto,  0) " & vbCrLf &
              "			    else 0 " & vbCrLf &
              "		    end Desconto, " & vbCrLf &
              "       case " & vbCrLf &
              "		    when isnull(P.ControlarRomaneio,0) = 1 " & vbCrLf &
              "			    then ISNULL(R.PesoLiquido,0) " & vbCrLf &
              "			    else NFI.QuantidadeFisica " & vbCrLf &
              "		    end PesoLiquido, " & vbCrLf &
              "       NFI.QuantidadeFiscal AS Variacao," & vbCrLf &
              "       ISNULL(NXD.PesoLiquido,0) as PesoChegada, " & vbCrLf &
              "       NFI.Valor, " & vbCrLf &
              "       COALESCE (Encargos.Encargo1, 0) AS ValorEncargo1, " & vbCrLf &
              "       COALESCE (Encargos.Encargo2, 0) AS ValorEncargo2," & vbCrLf &
              "       COALESCE (Encargos.Encargo3, 0) AS ValorEncargo3," & vbCrLf &
              "       COALESCE (Encargos.Encargo4, 0) AS ValorEncargo4," & vbCrLf &
              "       COALESCE (Encargos.Encargo5, 0) AS ValorEncargo5," & vbCrLf &
              "       COALESCE (Encargos.LIQUIDO, 0) AS ValorLiquido," & vbCrLf &
              "       P.Grupo, " & vbCrLf &
              "       ISNULL(RomaneiosXPesagens.Pesagem_Id, '') AS Laudo," & vbCrLf &
              "       ISNULL(LP.Placa, '') AS Placa," & vbCrLf &
              "       case " & vbCrLf &
              "          when isnull(NFI.capacidadeEmbalagem,0) = 0 " & vbCrLf &
              "            then '' " & vbCrLf &
              "            else isnull(EM.EmbalagemIndea,'') + '-' + isnull(NFI.TipoDeEmbalagem,'') +'-' + convert(nvarchar(10),NFI.capacidadeEmbalagem) " & vbCrLf &
              "       end Embalagem, " & vbCrLf &
              "       Pe.Pedido_Id, " & vbCrLf &
              "       isnull(REP.Nome,'') as NomeRepresentante, " & vbCrLf &
              "       isnull(M.Descricao,'') as MarcaProduto," & vbCrLf &
              "       NFER.ChaveNfe, " & vbCrLf &
              "       NF.Representante" & vbCrLf

        If chkResumo.Checked Then
            Sql &= "  INTO #Notas" & vbCrLf
        End If

        Sql &= "  FROM NotasFiscais AS NF WITH (NOLOCK) " & vbCrLf &
               " INNER JOIN NotasFiscaisXItens AS NFI WITH (NOLOCK)" & vbCrLf &
               "    ON NF.Empresa_Id      = NFI.Empresa_Id " & vbCrLf &
               "   AND NF.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf &
               "   AND NF.Cliente_Id      = NFI.Cliente_Id " & vbCrLf &
               "   AND NF.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf &
               "   AND NF.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
               "   AND NF.Serie_Id        = NFI.Serie_Id " & vbCrLf &
               "   AND NF.Nota_Id         = NFI.Nota_Id " & vbCrLf &
               "  Left Join Operacoes as OP WITH (NOLOCK)" & vbCrLf &
               "    ON OP.Operacao_Id     = NFI.Operacao " & vbCrLf &
               "  Left JOIN SubOperacoes AS SO WITH (NOLOCK)" & vbCrLf &
               "    ON SO.Operacao_Id     = NFI.Operacao " & vbCrLf &
               "   AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf &
               " Inner Join OperacaoXEstado OE WITH (NOLOCK)" & vbCrLf &
               "    on OE.Codigo_id = NFI.OperacaoXEstado" & vbCrLf &
               "  LEFT Join NFERealizadas NFER WITH (NOLOCK)" & vbCrLf &
               "    ON NF.Empresa_Id      = NFER.Empresa_Id" & vbCrLf &
               "   AND NF.EndEmpresa_Id   = NFER.EndEmpresa_Id" & vbCrLf &
               "   AND NF.Cliente_Id      = NFER.Cliente_Id" & vbCrLf &
               "   AND NF.EndCliente_Id   = NFER.EndCliente_Id" & vbCrLf &
               "   AND NF.EntradaSaida_Id = NFER.EntradaSaida_Id" & vbCrLf &
               "   AND NF.Serie_Id        = NFER.Serie_Id" & vbCrLf &
               "   AND NF.Nota_Id         = NFER.Nota_Id" & vbCrLf &
               " INNER JOIN Produtos AS P WITH (NOLOCK)" & vbCrLf &
               "    ON P.Produto_Id = NFI.Produto_Id " & vbCrLf &
               "  LEFT JOIN Marca M WITH (NOLOCK)" & vbCrLf &
               "    ON M.Marca_id = P.Marca" & vbCrLf &
               " LEFT JOIN NotasXDestinos NXD WITH (NOLOCK)" & vbCrLf &
               "    ON NXD.Empresa_Id      = NFI.Empresa_Id      " & vbCrLf &
               "   AND NXD.EndEmpresa_Id   = NFI.EndEmpresa_Id   " & vbCrLf &
               "   AND NXD.Cliente_Id      = NFI.Cliente_Id      " & vbCrLf &
               "   AND NXD.EndCliente_Id   = NFI.EndCliente_Id   " & vbCrLf &
               "   AND NXD.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
               "   AND NXD.Serie_Id        = NFI.Serie_Id        " & vbCrLf &
               "   AND NXD.Nota_Id         = NFI.Nota_Id         " & vbCrLf &
               "  LEFT Join Pedidos Pe WITH (NOLOCK)" & vbCrLf &
               "    on NF.Empresa_Id    = Pe.Empresa_Id " & vbCrLf &
               "   and NF.EndEmpresa_Id = Pe.EndEmpresa_Id " & vbCrLf &
               "   and NF.Pedido        = Pe.Pedido_Id " & vbCrLf &
               "   Left Join Pagamentos Pg " & vbCrLf &
               "   on Pg.Pagamento_Id = Pe.CondicaoPagamento " & vbCrLf &
               " INNER JOIN (SELECT NFE.Empresa_Id, NFE.EndEmpresa_Id, NFE.Cliente_Id, NFE.EndCliente_Id, NFE.EntradaSaida_Id, NFE.Serie_Id, NFE.Nota_Id, NFE.Produto_Id, NFE.Sequencia_Id," & vbCrLf &
               "                    Encargo1," & vbCrLf &
               "                    Encargo2," & vbCrLf &
               "                    Encargo3," & vbCrLf &
               "                    Encargo4," & vbCrLf &
               "                    Encargo5," & vbCrLf &
               "                    LIQUIDO" & vbCrLf &
               "               FROM #Encargos Nfe " & vbCrLf
        'Parâmetros somente para as Notas Fiscais
        'GetSqlWhereNF(Sql)


        Sql &= "             ) AS Encargos" & vbCrLf &
               "    ON NFI.Empresa_Id      = Encargos.Empresa_Id" & vbCrLf &
               "   AND NFI.EndEmpresa_Id   = Encargos.EndEmpresa_Id" & vbCrLf &
               "   AND NFI.Cliente_Id      = Encargos.Cliente_Id" & vbCrLf &
               "   AND NFI.EndCliente_Id   = Encargos.EndCliente_Id" & vbCrLf &
               "   AND NFI.EntradaSaida_Id = Encargos.EntradaSaida_Id" & vbCrLf &
               "   AND NFI.Serie_Id        = Encargos.Serie_Id" & vbCrLf &
               "   AND NFI.Nota_Id         = Encargos.Nota_Id  " & vbCrLf &
               "   AND NFI.Produto_Id      = Encargos.Produto_Id" & vbCrLf &
               "   AND NFI.Sequencia_Id    = Encargos.Sequencia_Id" & vbCrLf &
               "  Left Join Embalagens EM WITH (NOLOCK)" & vbCrLf &
               "    on EM.Embalagem_ID = NFI.Embalagem " & vbCrLf &
               "  LEFT JOIN NotasFiscaisXRomaneios AS NFR WITH (NOLOCK)" & vbCrLf &
               "    ON NFR.Empresa_Id      = NFI.Empresa_Id " & vbCrLf &
               "   AND NFR.EndEmpresa_Id   = NFI.EndEmpresa_Id " & vbCrLf &
               "   AND NFR.Cliente_Id      = NFI.Cliente_Id " & vbCrLf &
               "   AND NFR.EndCliente_Id   = NFI.EndCliente_Id " & vbCrLf &
               "   AND NFR.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
               "   AND NFR.Serie_Id        = NFI.Serie_Id " & vbCrLf &
               "   AND NFR.Nota_Id         = NFI.Nota_Id " & vbCrLf &
               "  LEFT JOIN Romaneios AS R WITH (NOLOCK)" & vbCrLf &
               "    ON R.Empresa_Id    = NFR.Empresa_Id " & vbCrLf &
               "   AND R.EndEmpresa_Id = NFR.EndEmpresa_Id " & vbCrLf &
               "   AND R.Romaneio_Id   = NFR.Romaneio_Id " & vbCrLf &
               "  LEFT JOIN Clientes AS CLI WITH (NOLOCK)" & vbCrLf &
               "    ON nf.Cliente_Id    = CLI.Cliente_Id " & vbCrLf &
               "   AND nf.EndCliente_Id = CLI.Endereco_Id " & vbCrLf &
               "  LEFT JOIN Clientes AS E WITH (NOLOCK)" & vbCrLf &
               "    ON nf.Empresa_Id    = E.Cliente_Id " & vbCrLf &
               "   AND nf.EndEmpresa_Id = E.Endereco_Id" & vbCrLf &
               "  LEFT JOIN Clientes AS DO WITH (NOLOCK)" & vbCrLf &
               "    ON NFI.Deposito    = DO.Cliente_Id " & vbCrLf &
               "   AND NFI.EndDeposito = DO.Endereco_Id " & vbCrLf &
               "  LEFT JOIN Clientes AS DD WITH (NOLOCK)" & vbCrLf &
               "    ON NF.Destino    = DD.Cliente_Id " & vbCrLf &
               "   AND NF.EndDestino = DD.Endereco_Id " & vbCrLf &
               "  LEFT JOIN RomaneiosXPesagens WITH (NOLOCK)" & vbCrLf &
               "    ON R.Empresa_Id    = RomaneiosXPesagens.Empresa_Id " & vbCrLf &
               "   AND R.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id" & vbCrLf &
               "   AND R.Romaneio_Id   = RomaneiosXPesagens.Romaneio_Id" & vbCrLf &
               "  LEFT JOIN Pesagem LP WITH (NOLOCK) " & vbCrLf &
               "    ON RomaneiosXPesagens.Empresa_Id        = LP.Empresa_Id " & vbCrLf &
               "   AND RomaneiosXPesagens.EndEmpresa_Id     = LP.EndEmpresa_Id " & vbCrLf &
               "   AND RomaneiosXPesagens.Pesagem_Id        = LP.Pesagem_Id" & vbCrLf &
               "  LEFT Join Comissoes Com WITH (NOLOCK)" & vbCrLf &
               "    ON Pe.Empresa_Id           = Com.Empresa_Id" & vbCrLf &
               "   And Pe.EndEmpresa_Id        = Com.EndEmpresa_Id" & vbCrLf &
               "   And Pe.Pedido_Id            = Com.Pedido_Id" & vbCrLf &
               "   AND Com.Principal           = 'S'" & vbCrLf

        Dim strClienteRepresentante() As String = txtCodigoCliRepre.Value.Split("-")
        If strClienteRepresentante(0).Length > 0 Then
            Sql &= " INNER Join PedidoxTabelaDeComissao PTC WITH (NOLOCK)" & vbCrLf & _
                   "    ON Com.Empresa_Id           = PTC.Empresa_Id" & vbCrLf & _
                   "   And Com.EndEmpresa_Id        = PTC.EndEmpresa_Id" & vbCrLf & _
                   "   And Com.Pedido_Id            = PTC.Pedido_Id" & vbCrLf & _
                   "   AND NFI.Produto_id           = PTC.Produto_Id" & vbCrLf & _
                   "   AND PTC.Representante_Id    ='" & strClienteRepresentante(0) & "'" & vbCrLf & _
                   "   AND PTC.EndRepresentante_Id = " & strClienteRepresentante(1) & vbCrLf
            Parametros &= "Representante: " & txtClienteRepresentante.Text & vbCrLf
        Else
            Sql &= "  LEFT Join PedidoxTabelaDeComissao PTC WITH (NOLOCK)" & vbCrLf & _
                   "    ON Com.Empresa_Id           = PTC.Empresa_Id" & vbCrLf & _
                   "   And Com.EndEmpresa_Id        = PTC.EndEmpresa_Id" & vbCrLf & _
                   "   And Com.Pedido_Id            = PTC.Pedido_Id" & vbCrLf & _
                   "   AND NFI.Produto_id           = PTC.Produto_Id" & vbCrLf
        End If

        Sql &= "  LEFT JOIN Clientes REP WITH (NOLOCK)" & vbCrLf & _
               "    ON REP.Cliente_id  = PTC.Representante_Id" & vbCrLf & _
               "   AND REP.Endereco_Id = PTC.EndRepresentante_Id" & vbCrLf

        If txtCodigoTransportador.Value.Length > 0 Then
            Dim TRA As String() = txtCodigoTransportador.Value.Split(New Char() {"-"})
            Sql &= " INNER JOIN NotasFiscaisXTransportadores AS nfxT WITH (NOLOCK)" & vbCrLf & _
                   "    ON nf.Empresa_Id        = nfxT.Empresa_Id" & vbCrLf & _
                   "   AND nf.EndEmpresa_Id     = nfxT.EndEmpresa_Id" & vbCrLf & _
                   "   AND nf.Cliente_Id        = nfxT.Cliente_Id" & vbCrLf & _
                   "   AND nf.EndCliente_Id     = nfxT.EndCliente_Id" & vbCrLf & _
                   "   AND nf.EntradaSaida_Id   = nfxT.EntradaSaida_Id" & vbCrLf & _
                   "   AND nf.Serie_Id          = nfxT.Serie_Id" & vbCrLf & _
                   "   AND nf.Nota_Id           = nfxT.Nota_Id" & vbCrLf & _
                   "   AND nfxT.Proprietario    ='" & TRA(0) & "'" & vbCrLf & _
                   "   AND nfxT.EndProprietario = " & TRA(1) & vbCrLf
        End If

        Sql &= " WHERE 1=1 " & vbCrLf

        'Parâmetros somente para as Notas Fiscais
        GetSqlWhereNF(Sql)

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            Sql &= "   AND op.classe " & IIf(rdComClasse.Checked, "", "not") & " in ('" & op & "')"
        End If

        If ddlTroca.SelectedIndex = 1 Then
            Sql &= "   and isnull(Pe.Troca,0) = 1"
        ElseIf ddlTroca.SelectedIndex = 2 Then
            Sql &= "   and isnull(Pe.Troca,0) <> 1"
        End If

        If chkPessoaJuridica.Checked Then
            Sql &= " and len(CLI.Cliente_ID) <> 14"
        End If

        If chkPessoaFisica.Checked Then
            Sql &= " and len(CLI.Cliente_ID) <> 11 "
        End If

        If chkPessoaExterior.Checked Then
            Sql &= " and CLI.Estado <> 'EX'"
        End If

        If ddlStICMS.SelectedIndex > 0 Then Sql &= "  and OE.STICMS = " & ddlStICMS.SelectedValue
        If ddlStIPI.SelectedIndex > 0 Then Sql &= "  and OE.STIPI = " & ddlStIPI.SelectedValue
        If ddlStPISCOFINS.SelectedIndex > 0 Then Sql &= "  and OE.STPISCOFINS = " & ddlStPISCOFINS.SelectedValue

        If pc.Length > 0 Then Sql &= " And SO.ApuracaoDeCustos in (" & pc & ")"
        If txtPedido.Text.Trim.Length > 0 Then Sql &= "AND PE.Pedido_id        = " & txtPedido.Text & " " & vbCrLf
        If txtCN.Text.Trim.Length > 0 Then Sql &= "AND PE.PedidoEfetivo ='" & txtCN.Text & "'" & vbCrLf
        If txtContrato.Text.Trim.Length > 0 Then Sql &= "AND PE.Contrato      ='" & txtContrato.Text & "'" & vbCrLf

        If cmbUnidadeNegocio.SelectedIndex > 0 And (txtPedido.Text.Trim.Length = 0 Or txtCN.Text.Trim.Length = 0 Or txtContrato.Text.Trim.Length = 0) Then
            Sql &= " AND isnull(Pe.UnidadeDeNegocio,'" & cmbUnidadeNegocio.SelectedValue & "') ='" & cmbUnidadeNegocio.SelectedValue & "'" & vbCrLf
            Parametros &= "Unidade De Negocio:" & cmbUnidadeNegocio.SelectedItem.Text & vbCrLf
        End If

        If txtCodigoDeposito.Value.Length > 0 Then
            Dim strDeposito As String()
            strDeposito = txtCodigoDeposito.Value.Split(New Char() {"-"})
            Sql &= " AND NFI.Deposito    ='" & strDeposito(0) & "'" & vbCrLf & _
                   " AND NFI.EndDeposito = " & strDeposito(1) & vbCrLf
        End If

        If txtCodigoRepresentante.Value.Length > 0 Then
            Dim strRepresentante As String()
            strRepresentante = txtCodigoRepresentante.Value.Split(New Char() {"-"})
            Sql &= " AND NF.Representante    ='" & strRepresentante(0) & "'" & vbCrLf &
                   " AND NF.EndRepresentante = " & strRepresentante(1) & vbCrLf
        End If

        VerificarGrupoProduto(Sql)
        If cmbOperacao.SelectedIndex > 0 Then Sql &= "AND OP.Operacao_Id = " & cmbOperacao.SelectedValue & " " & vbCrLf
        If cmbSubOperacao.SelectedIndex > 0 Then
            Dim strSubOpe() As String = cmbSubOperacao.SelectedValue.ToString.Split("-")
            Sql &= "AND SO.SubOperacoes_ID = " & strSubOpe(1) & " " & vbCrLf
        End If

        If chkRemessa.Checked Then Sql &= "AND SO.Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "'"
        If chkGlobal.Checked Then Sql &= "AND SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'"

        If cmbSafra.SelectedIndex > 0 Then Sql &= "AND PE.Safra = '" & cmbSafra.SelectedValue & "' " & vbCrLf
        If ddlMarca.SelectedIndex > 0 Then Sql &= " And P.Marca =" & ddlMarca.SelectedValue

        If chkCusto.Checked Then
            Sql &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) "
        End If

        If txtPlaca.Text.Length > 0 Then
            Sql &= " AND LP.Placa = '" & txtPlaca.Text & "' " & vbCrLf
        End If

        If lstCfopSelecionados.Items.Count > 0 Then
            Sql &= "AND ("
            Dim strOr As String = ""

            Dim k As Integer = 0
            While k < lstCfopSelecionados.Items.Count
                Sql &= strOr & "OE.CodigoFiscal = " & lstCfopSelecionados.Items(k).Text.Substring(0, 4)
                strOr = " OR "
                k += 1
            End While

            Sql &= ") " & vbCrLf
        End If

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If txtSerie.Text.Length > 0 Then
            Sql &= "AND NF.Serie_Id = '" & txtSerie.Text & "'" & vbCrLf
        End If

        If txtNumeroNota.Text.Length > 0 Then
            Sql &= "AND NF.Nota_Id = '" & txtNumeroNota.Text & "'" & vbCrLf
        End If

        If RdNome.Checked OrElse chkQuebraPorCliente.Checked Then
            Sql &= "ORDER BY NF.Empresa_Id, NF.Entradasaida_Id, Cli.Nome, NF.Nota_Id" & vbCrLf
        ElseIf RdMovimento.Checked Then
            Sql &= "ORDER BY NF.Empresa_Id, NF.Entradasaida_Id, NF.Movimento" & vbCrLf
        Else
            Sql &= "ORDER BY NF.Empresa_Id, NF.Entradasaida_Id, NF.Nota_Id, Cli.Nome" & vbCrLf
        End If

        If chkResumo.Checked Then
            Sql &= " SELECT * FROM #Notas " & vbCrLf
        End If

        Return Sql
        '#NFD 05
    End Function

    Private Function GetSqlWhereNF(ByRef Sql) As String
        Sql &= "   AND NF.Situacao = " & ddlSituacao.SelectedValue & " " & vbCrLf

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        op = String.Join("','", lstFinalidade.GetSelecteds)
        If op.Length > 0 Then
            Sql &= "   AND NF.Finalidade " & IIf(rdComFinalidade.Checked, "", "not") & " in ('" & op & "')"
        End If

        If ddlNfTroca.SelectedIndex = 1 Then
            Sql &= "   and isnull(NF.Troca,0) = 1"
        ElseIf ddlNfTroca.SelectedIndex = 2 Then
            Sql &= "   and isnull(NF.Troca,0) <> 1"
        End If

        If ddlTipoFrete.SelectedIndex <> 0 Then
            Sql &= "   AND isnull(NF.CIFFOB,'NEN') = '" & ddlTipoFrete.SelectedValue & "'" & vbCrLf
        End If

        If txtSerie.Text.Length > 0 Then Sql &= " and NF.Serie_Id ='" & txtSerie.Text & "'"

        If chkPeriodo.Checked And txtPedido.Text.Trim.Length = 0 And txtCN.Text.Trim.Length = 0 And txtContrato.Text.Trim.Length = 0 Then
            Sql &= " AND NF.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "' " & vbCrLf
        End If

        If chkEntrada.Checked And Not chkSaida.Checked Then
            Sql &= " AND NF.EntradaSaida_Id = 'E' " & vbCrLf
        ElseIf chkSaida.Checked And Not chkEntrada.Checked Then
            Sql &= " AND NF.EntradaSaida_Id = 'S' " & vbCrLf
        End If

        If cmbEmpresa.SelectedIndex > 0 And (txtPedido.Text.Trim.Length = 0 Or txtCN.Text.Trim.Length = 0 Or txtContrato.Text.Trim.Length = 0) Then
            Dim strEmpresa As String()
            strEmpresa = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

            If chkConsolidarEmpresa.Checked Then
                Sql &= " AND left(NF.Empresa_Id,8) ='" & strEmpresa(0).Substring(0, 8) & "'" & vbCrLf
            Else
                Sql &= " AND NF.Empresa_Id    ='" & strEmpresa(0) & "'" & vbCrLf & _
                       " AND NF.EndEmpresa_Id = " & strEmpresa(1) & vbCrLf
            End If
        End If

        If txtCodigoCliente.Value.Length > 0 And (txtPedido.Text.Trim.Length = 0 Or txtCN.Text.Trim.Length = 0 Or txtContrato.Text.Trim.Length = 0) Then
            Dim strCliente As String()
            strCliente = txtCodigoCliente.Value.Split(New Char() {"-"})

            If chkConsolidarCliente.Checked Then
                Sql &= " AND Left(NF.Cliente_Id,8) = '" & strCliente(0).Substring(0, 8) & "' " & vbCrLf
            Else
                Sql &= " AND NF.Cliente_Id = '" & strCliente(0) & "'" & vbCrLf & _
                       " AND NF.EndCliente_Id = " & strCliente(1) & vbCrLf
            End If
        End If

        If HOrigemDestino.Value.Length > 0 Then
            Dim strOrigemDestino As String()
            strOrigemDestino = HOrigemDestino.Value.Split("-")
            Sql &= " AND NF.Destino    ='" & strOrigemDestino(0) & "'" & vbCrLf & _
                   " AND NF.EndDestino = " & strOrigemDestino(1) & vbCrLf
        End If

        If RadSerieNao.Checked = True Then Sql &= "AND NF.Serie_Id not in('D','F') " & vbCrLf

        Dim tp As String = ""
        For i As Integer = 0 To chkTipoDeDocumento.Items.Count - 1
            If chkTipoDeDocumento.Items(i).Selected Then
                tp &= IIf(tp.Length = 0, "", ",") & chkTipoDeDocumento.Items(i).Value
            End If
        Next

        If tp <> "" Then
            Sql &= " AND NF.TipoDeDocumento in(" & tp & ")" & vbCrLf
        End If

        Return Sql
    End Function

    Private Function VerificarGrupoProduto(ByRef Sql As String) As String
        If ucSelecaoProduto.TemSelecionado() Then
            Dim retorno As ArrayList
            retorno = ucSelecaoProduto.GetSqlEParametrosRelatorio("P.Grupo", "NFI.Produto_id")
            Sql &= " AND " & retorno(0)
            Return retorno(1)
        Else
            Return ""
        End If
    End Function

    Private Sub LimparCampos()
        ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, "", True)
        Funcoes.VerificaUnidadeDeNegocio(cmbUnidadeNegocio, cmbEmpresa)
        CarregarSituacao()

        txtClientes.Text = ""
        txtCodigoCliente.Value = ""

        txtDeposito.Text = ""
        txtCodigoDeposito.Value = ""

        txtOrigemDestino.Text = ""
        HOrigemDestino.Value = ""

        txtClienteRepresentante.Text = ""
        txtCodigoCliRepre.Value = ""

        txtClienteTransportador.Text = ""
        txtCodigoTransportador.Value = ""

        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()

        'cmbFinalidade.SelectedIndex = 0
        lstFinalidade.SelectedIndex = -1
        lstClasseOp.SelectedIndex = -1

        cmbGrupoCFOP.SelectedIndex = 0

        lstCfop.Items.Clear()
        lstCfopSelecionados.Items.Clear()

        cmbSafra.SelectedIndex = 0

        txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")

        chkEntrada.Checked = True
        chkSaida.Checked = True
        optPonderada.Checked = True
        optFisico.Checked = True
        chkQuebraPorCliente.Visible = False
        chkQuebraPorCliente.Checked = False
        chkCusto.Checked = False

        txtPedido.Text = ""
        txtPlaca.Text = ""

        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        Panel3.Parent.Visible = False
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidadeNegocio.Enabled = False
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteESNotasCliente" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteESNotasCliente" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteESNotasCliente" & HID.Value.ToString())
        ElseIf Not Session("objClienteESNotasDeposito" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteESNotasDeposito" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtDeposito.Text = itemCliente.Text
            txtCodigoDeposito.Value = itemCliente.Value
            Session.Remove("objClienteESNotasDeposito" & HID.Value.ToString())
        ElseIf Not Session("objClienteESNotasOrigemDestino" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteESNotasOrigemDestino" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtOrigemDestino.Text = itemCliente.Text
            HOrigemDestino.Value = itemCliente.Value
            Session.Remove("objClienteESNotasOrigemDestino" & HID.Value.ToString())
        ElseIf Not Session("objClienteESNotasCarCliRepre" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteESNotasCarCliRepre" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClienteRepresentante.Text = itemCliente.Text
            txtCodigoCliRepre.Value = itemCliente.Value
            Session.Remove("objClienteESNotasCarCliRepre" & HID.Value.ToString())
        ElseIf Not Session("objClienteESNotasCarCliTransportador" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objClienteESNotasCarCliTransportador" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClienteTransportador.Text = itemCliente.Text
            txtCodigoTransportador.Value = itemCliente.Value
            Session.Remove("objClienteESNotasCarCliTransportador" & HID.Value.ToString())
        ElseIf Not Session("objRepresentante" & HID.Value.ToString) Is Nothing Then
            Dim cli As [Lib].Negocio.Cliente = Session("objRepresentante" & HID.Value.ToString())
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtRepresentante.Text = itemCliente.Text
            txtCodigoRepresentante.Value = itemCliente.Value
            Session.Remove("objRepresentante" & HID.Value.ToString())
        End If
    End Sub

    Private Function getListEncargos() As List(Of String)
        Try
            Dim lst As New List(Of String)

            For Each item As String In lstEncargos.GetSelectedValues()
                lst.Add(item)
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioEntradaSaidaNotas", "ACESSAR") Then
                Panel3.Parent.Visible = False
                chkQuebraPorCliente.Visible = False

                ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
                ddl.Carregar(lstFinalidade, CarregarDDL.Tabela.Finalidade, "", False)
                ddl.Carregar(cmbSafra, CarregarDDL.Tabela.Safra, "", True)
                ddl.Carregar(cmbGrupoCFOP, CarregarDDL.Tabela.CFOPGrupo, "", True)
                CarregarPlanoDeCusto()
                CarregarTipoDeDocumento(False)
                CarregarSituacao()
                ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "")
                ddl.Carregar(ddlStICMS, CarregarDDL.Tabela.SituacaoTributariaICMS, "")
                ddl.Carregar(ddlStIPI, CarregarDDL.Tabela.SituacaoTributariaIPI, "")
                ddl.Carregar(ddlStPISCOFINS, CarregarDDL.Tabela.SituacaoTributariaPISCOFINS, "")
                ddl.Carregar(lstEncargos, CarregarDDL.Tabela.Encargos, "", False)
                ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)
                LimparCampos()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If

        If (Not Session("txtCnpjESNotas") Is Nothing) And txtClientes.Text.Length = 0 Then
            txtClientes.Text = Session("txtNomeESNotas") & " - " & Session("txtCidadeESNotas") & " - " & Session("txtEstadoESNotas") & " - " & Funcoes.FormatarCpfCnpj(Session("txtCnpjESNotas"))
            txtCodigoCliente.Value = Session("txtCnpjESNotas") & ";" & Session("txtEndESNotas")
            HttpContext.Current.Session.Remove("txtCnpjESNotas")
            HttpContext.Current.Session.Remove("txtEndESNotas")
            HttpContext.Current.Session.Remove("txtNomeESNotas")
            HttpContext.Current.Session.Remove("txtEnderecoESNotas")
            HttpContext.Current.Session.Remove("txtCidadeESNotas")
            HttpContext.Current.Session.Remove("txtEstadoESNotas")
        End If

        If (Not Session("txtCnpjESNotasDeposito") Is Nothing) And txtDeposito.Text.Length = 0 Then
            txtDeposito.Text = Session("txtNomeESNotasDeposito") & " - " & Session("txtCidadeESNotasDeposito") & " - " & Session("txtEstadoESNotasDeposito") & " - " & Funcoes.FormatarCpfCnpj(Session("txtCnpjESNotasDeposito"))
            txtCodigoDeposito.Value = Session("txtCnpjESNotasDeposito") & ";" & Session("txtEndESNotasDeposito")
            HttpContext.Current.Session.Remove("txtCnpjESNotasDeposito")
            HttpContext.Current.Session.Remove("txtEndESNotasDeposito")
            HttpContext.Current.Session.Remove("txtNomeESNotasDeposito")
            HttpContext.Current.Session.Remove("txtEnderecoESNotasDeposito")
            HttpContext.Current.Session.Remove("txtCidadeESNotasDeposito")
            HttpContext.Current.Session.Remove("txtEstadoESNotasDeposito")
        End If
    End Sub

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmbUnidadeNegocio.SelectedIndexChanged
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue, True)
    End Sub

    Protected Sub EntradaSaida_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If Not (chkEntrada.Checked Or chkSaida.Checked) Then
            chkEntrada.Checked = True
            chkSaida.Checked = True
        End If
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        BuscarSubOperacoes(cmbOperacao.SelectedValue)
    End Sub

    Protected Sub cmbGrupoCFOP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If cmbGrupoCFOP.SelectedIndex > 0 Then
            Panel3.Parent.Visible = True
            Dim ListaCFOP As New [Lib].Negocio.ListCFOP(True, cmbGrupoCFOP.SelectedValue)

            lstCfop.Items.Clear()

            Dim j As Integer = 0
            While j < ListaCFOP.Count
                lstCfop.Items.Add(New ListItem(Format(ListaCFOP(j).Codigo, "0000") & "-" & ListaCFOP(j).Descricao, ListaCFOP(j).Codigo))
                j += 1
            End While
        Else
            Panel3.Parent.Visible = False
        End If
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaCliente.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteESNotasCliente" & HID.Value, "txtNome")
    End Sub

    Protected Sub cmdConsultaDeposito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaDeposito.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteESNotasDeposito" & HID.Value, "txtNome")
    End Sub

    Protected Sub imgAdicionar_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If lstCfop.SelectedIndex > -1 Then

            Dim temCFOP As Boolean = False

            Dim i As Integer = 0
            While i < lstCfopSelecionados.Items.Count
                If lstCfopSelecionados.Items(i).Text = lstCfop.SelectedItem.Text Then
                    temCFOP = True
                End If
                i += 1
            End While

            If temCFOP Then
                MsgBox(Me.Page, "CFOP já foi selecionado")
            Else
                lstCfopSelecionados.Items.Add(New ListItem(lstCfop.SelectedItem.Text))
            End If
        End If
    End Sub

    Protected Sub imgRemover_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        If lstCfopSelecionados.SelectedIndex > -1 Then
            lstCfopSelecionados.Items.RemoveAt(lstCfopSelecionados.SelectedIndex)
        End If
    End Sub

    Protected Sub cmbSafra_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If cmbSafra.SelectedIndex > 0 Then
            chkPeriodo.Checked = False
        Else
            chkPeriodo.Checked = True
        End If
    End Sub

    Protected Sub btnOrigemDestino_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnOrigemDestino.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteESNotasOrigemDestino" & HID.Value, "txtNome")
    End Sub

    Protected Sub optFisico_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optFisico.CheckedChanged
        chkQuebraPorCliente.Visible = False
        chkQuebraPorCliente.Checked = False
        pnlMedia.Visible = True
        pnlOpcaoFiscal.Visible = Not optFisico.Checked
        divSeries.Visible = False
        CarregarTipoDeDocumento(False)
    End Sub

    Protected Sub optFiscal_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles optFiscal.CheckedChanged
        chkQuebraPorCliente.Visible = True
        pnlMedia.Visible = False
        pnlOpcaoFiscal.Visible = optFiscal.Checked
        divSeries.Visible = True
        CarregarTipoDeDocumento(True)
    End Sub

    Protected Sub cmdBuscaCliRepre_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaCliRepre.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteESNotasCarCliRepre" & HID.Value, "txtNome")

    End Sub

    Protected Sub cmdBuscaCliTransportador_Click(sender As Object, e As EventArgs) Handles cmdBuscaCliTransportador.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteESNotasCarCliTransportador" & HID.Value, "txtNome")
    End Sub

    Protected Sub chkPeriodo_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPeriodo.CheckedChanged
        pnlDataMovimento.Visible = chkPeriodo.Checked
    End Sub

    Protected Sub ddlMarca_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMarca.SelectedIndexChanged
        If ddlMarca.SelectedIndex = 0 Then
            ucSelecaoProduto.WhereProduto = ""
        Else
            ucSelecaoProduto.WhereProduto = "marca = " & ddlMarca.SelectedValue
        End If
        ucSelecaoProduto.CarregarNivel(1)
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirExcel(ByRef dsExcel As DataSet)
        Try
            Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            'Remove coluna do DataTable
            'Dim dt As DataTable = dsExcel.Tables(0)
            'dt.Columns.Remove("EmpresaNome")

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha títulos
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Movimentações Fiscais.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o nome da empresa e o cnpj
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa a cidade e o estado da empresa
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Movimentações Fiscais.")
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "DATA : " & String.Format(Now(), "dd/MM/yyyy"))
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:E{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dsExcel.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dsExcel.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In dsExcel.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("G{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'formatando células numéricas
                        'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    'worksheet.Cells(String.Format("AE{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                    ''criando colunas de totalizadores
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Font.Bold = True
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Formula = String.Format("=SUM(AF6:AF{0})", rowIndex - 1)
                    'worksheet.Cells(String.Format("AF{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelDados.Click
        Try
            EmitirRelatorio(False, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean, Optional ByVal ExcelDados As Boolean = False)
        Try
            If Funcoes.VerificaPermissao("RelatorioEntradaSaidaNotas", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim Encargos As List(Of String) = getListEncargos()

                    If Encargos.Count = 0 AndAlso optFiscal.Checked Then
                        Encargos.Add("ICMS")
                        Encargos.Add("FUNRURAL")
                        Encargos.Add("FUNRURAL TERCEI")
                        Encargos.Add("FETHAB")
                        Encargos.Add("FACS")
                    End If

                    Dim Pc As String = ""
                    If LstPlanoDeCusto.GetSelectedValues().Count > 0 Then
                        Dim i As Integer() = LstPlanoDeCusto.GetSelectedIndices()

                        For x As Integer = 0 To i.GetLength(0) - 1
                            Pc &= IIf(Pc.Length > 0, ",", "") & LstPlanoDeCusto.Items(i(x)).Value
                        Next
                    End If

                    Dim dsRelatorio As New DataSet
                    Dim strSQL As String

                    If optFisico.Checked Then
                        strSQL = BuscaSQLNotasFisico(Pc)
                    Else
                        strSQL = BuscaSQLNotasFiscal(Pc, Encargos)
                    End If

                    If chkResumo.Checked Then
                        strSQL &= BuscaSQLResumoProduto(ddlTipoDeResumo.SelectedIndex = 1)
                    End If

                    dsRelatorio = Banco.ConsultaDataSet(strSQL, "NotasFiscais")

                    If dsRelatorio.Tables.Count > 1 Then
                        dsRelatorio.Tables(1).TableName = "Resumo"
                    End If

                    Dim crpt As New ReportDocument()

                    Try
                        Dim strCaminho As String = String.Empty
                        If optFisico.Checked Then
                            strCaminho &= IIf(RdNota.Checked, HttpContext.Current.Server.MapPath("~/Reports/Cr_EntradaSaidaNotaOrdenadoPorNota.rpt"), HttpContext.Current.Server.MapPath("~/Reports/Cr_EntradaSaidaNotas.rpt"))
                        Else
                            strCaminho &= HttpContext.Current.Server.MapPath("~/Reports/Cr_EntradasSaidasNotasFiscalNome.rpt")
                        End If

                        crpt.Load(strCaminho)

                        Dim strNomeArquivo As String
                        If Pdf Then
                            strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                        Else
                            strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
                        End If

                        Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)

                        crpt.SetDataSource(dsRelatorio)

                        Dim parameters = New Dictionary(Of String, Object)()
                        Dim strHistorico As String = "Período de " & txtDataInicial.Text & " à " & txtDataFinal.Text
                        parameters.Add("NomeEmpresa", HttpContext.Current.Session("ssNomeEmpresa"))

                        Dim i As Integer = 1

                        If optFiscal.Checked Then
                            If Encargos.Count > 0 Then
                                For Each item As String In Encargos
                                    parameters.Add("Encargo" & i, item)
                                    parameters.Add("ResEncargo" & i, item)
                                    i += 1
                                Next
                            Else
                                parameters.Add("Encargo1", "Icms")
                                parameters.Add("Encargo2", "Funrural")
                                parameters.Add("Encargo3", "Funrural Tercei")
                                parameters.Add("Encargo4", "Fethab")
                                parameters.Add("Encargo5", "Facs")
                                parameters.Add("ResEncargo1", "Icms")
                                parameters.Add("ResEncargo2", "Funrural")
                                parameters.Add("ResEncargo3", "Funrural Tercei")
                                parameters.Add("ResEncargo4", "Fethab")
                                parameters.Add("ResEncargo5", "Facs")
                            End If
                        End If

                        If chkEntrada.Checked And chkSaida.Checked Then
                            strHistorico &= " - ENTRADA(S)/SAÍDA(S)"
                        ElseIf chkEntrada.Checked Then
                            strHistorico &= " - ENTRADA(S)"
                        ElseIf chkSaida.Checked Then
                            strHistorico &= " - SAÍDA(S)"
                        End If

                        If optFisico.Checked Then
                            strHistorico &= " - FÍSICO"
                        ElseIf optFiscal.Checked Then
                            strHistorico &= " - FISCAL"
                        End If

                        If RdNota.Checked Then
                            strHistorico &= " - POR NOTA"
                        ElseIf RdNome.Checked Then
                            strHistorico &= " - POR NOME"
                        End If

                        If ddlTipoFrete.SelectedIndex <> 0 Then
                        End If

                        If Parametros.Length > 0 Then strHistorico &= " - " & Parametros
                        parameters.Add("Periodo", strHistorico)
                        If optFisico.Checked Then
                            parameters.Add("Media", IIf(optAritmetica.Checked, "Aritmetica", "Ponderada"))
                        Else
                            parameters.Add("@QuebraCliente", chkQuebraPorCliente.Checked)
                        End If

                        If optFiscal.Checked Then
                            parameters.Add("MostrarResumo", chkResumo.Checked)
                            If ddlTipoDeResumo.SelectedIndex = 0 Then
                                parameters.Add("TituloDoResumo", "Resumo Por Produto")
                            Else
                                parameters.Add("TituloDoResumo", "Resumo Por CFOP")
                            End If
                        End If

                        If optFiscal.Checked Then
                            If Not Pdf Then
                                parameters.Add("Excel", True)
                            Else
                                parameters.Add("Excel", False)
                            End If
                        End If

                        Funcoes.BindParameters(crpt, parameters)

                        If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

                        If Pdf Then
                            crpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strArquivo)
                        ElseIf ExcelDados Then
                            EmitirExcel(dsRelatorio)
                        Else
                            crpt.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, strArquivo)
                        End If

                        If IO.File.Exists(strArquivo) Then
                            If Pdf = True Then
                                ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & strNomeArquivo & "');", True)
                            Else
                                ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.location = '" & strNomeArquivo & "';", True)
                            End If
                        End If

                        If chkZIP.Checked OrElse chkImpressao.Checked OrElse chkImpressao.Checked Then
                            PrintOpenPDF(dsRelatorio)
                        End If

                    Catch ex As Exception
                        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                    Finally
                        crpt.Close()
                        crpt.Dispose()
                    End Try
                End If
            Else
                MsgBox(Me.Page, "Usuario sem permissao para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub PrintOpenPDF(dsRelatorio As DataSet)

        Dim dt As DataTable = dsRelatorio.Tables(0) ' A tabela do DataSet onde está a coluna "nota"
        Dim pageNumber As Integer = 1 ' Inicia a contagem de páginas

        If txtPaginaInicial.Text.Length > 0 Then
            pageNumber = txtPaginaInicial.Text
        End If

        Dim rowCount As Integer = 0 ' Contador de linhas para controle da troca de página
        Dim caminhoLivro As String = Path.Combine(Server.MapPath("~"), "MovimentoFiscais")
        Dim empresa() As String = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

        Dim pastaLivro As String = DateTime.Now.ToString("yyyyMM")
        Dim pathFiles As String = String.Format("{0}\{1}\{2}", caminhoLivro, empresa(0), pastaLivro)
        Dim urlPath As String = String.Format("MovimentoFiscais/{0}/{1}", empresa(0), pastaLivro)

        If chkZIP.Checked Then

            'Se a opção de zipar estiver marcado, precisamos limpar os arquivos se houver seleção de paginas
            If Directory.Exists(String.Format("{0}\{1}\{2}", caminhoLivro, empresa(0), pastaLivro)) Then

                Directory.Delete(String.Format("{0}\{1}\{2}", caminhoLivro, empresa(0), pastaLivro), True)

            End If

        End If

        Funcoes.CriarPastaLivroFiscal(String.Format("{0}\{1}\{2}", caminhoLivro, empresa(0), pastaLivro))

        Dim NotasPorPagina As Integer = 34

        For Each rowPrint As DataRow In dt.Rows

            rowCount += 1

            ' A cada 24 linhas, mudar de página
            If rowCount > NotasPorPagina Then

                pageNumber += 1
                Console.WriteLine("Página " & pageNumber)
                rowCount = 1 ' Reseta o contador de linhas para a nova página
                'As demais paginas tem mais notas por pagina, mudamos para 40
                NotasPorPagina = 40
            End If

            If txtPaginaInicial.Text.Length > 0 Then
                If pageNumber < CType(txtPaginaInicial.Text, Integer) Then
                    Continue For
                End If
            End If

            If txtPaginaFinal.Text.Length > 0 Then
                If pageNumber > CType(txtPaginaFinal.Text, Integer) Then
                    Exit For
                End If
            End If

            Dim objNFe As New [Lib].Negocio.NotaFiscal()

            objNFe.CodigoEmpresa = rowPrint("Empresa_Id")
            objNFe.EnderecoEmpresa = rowPrint("EndEmpresa_Id")
            objNFe.CodigoCliente = rowPrint("Cliente_Id")
            objNFe.EnderecoCliente = rowPrint("EndCliente_Id")
            objNFe.EntradaSaida = IIf(rowPrint("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objNFe.Serie = rowPrint("Serie_Id")
            objNFe.Codigo = rowPrint("Nota_Id")
            objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

            Dim msgNFE As String = ""

            Dim caminhoArquivo As String

            If objNFe.ChaveNFE.Length > 0 Then
                caminhoArquivo = String.Format("{0}\{1}.pdf", pathFiles, objNFe.ChaveNFE)
            Else
                caminhoArquivo = String.Format("{0}\{1}_{2}_{3}.pdf", pathFiles, objNFe.CodigoEmpresa, objNFe.CodigoCliente, objNFe.Codigo.ToString())
            End If

            If Not File.Exists(caminhoArquivo) Then

                Dim bytes As Byte() = Nothing

                If objNFe.Arquivos.Where(Function(x) x.Arquivo.Length > 0 And x.Descricao.ToUpper().Contains(".PDF")).Count > 0 Then

                    Dim arquivo As New [Lib].Negocio.Arquivo
                    arquivo = objNFe.Arquivos.Where(Function(x) x.Arquivo.Length > 0 And x.Descricao.ToUpper().Contains(".PDF")).FirstOrDefault()
                    bytes = arquivo.Arquivo
                End If

                If bytes Is Nothing OrElse bytes.Length = 0 Then

                    If DocumentoEletronico.ImprimirNFeDanfe(objNFe, 1, msgNFE) Then
                        bytes = New FilesManager().getFile(String.Format("{0}.pdf", objNFe.ChaveNFE), eTipoDeDocumento.Nota)
                    Else
                        If objNFe.ChaveNFE.Length = 0 Then
                            msgNFE = String.Format("{0} - Empresa: {1} - Cliente: {2} - Nota Fiscal: {3}", msgNFE, objNFe.CodigoEmpresa, objNFe.CodigoCliente, objNFe.Codigo.ToString())
                        End If
                        MsgBox(Me.Page, msgNFE)
                        Continue For
                    End If

                End If

                If bytes IsNot Nothing Then

                    System.IO.File.WriteAllBytes(caminhoArquivo, bytes)

                End If

            End If

            If chkImpressao.Checked Then
                Funcoes.ImprimirArquivo(Me.Page, caminhoArquivo, objNFe)
            End If

            If chkAbrir.Checked Then

                Dim url As String = HttpContext.Current.Request.Url.AbsoluteUri
                Dim domainServer As String

                If HttpContext.Current.Request.Url.Query.Length > 0 Then
                    domainServer = url.Replace(HttpContext.Current.Request.Url.Query, "").Replace(HttpContext.Current.Request.Url.LocalPath, "")
                Else
                    domainServer = url.Replace(HttpContext.Current.Request.Url.LocalPath, "")
                End If

                Dim fileInfo As FileInfo = New FileInfo(caminhoArquivo)
                Dim bServidor As Boolean

                If url.ToUpper().Contains("/NGS/") Or url.ToUpper().Contains("/NGSTESTE/") Then
                    bServidor = True
                End If

                If bServidor Then
                    If url.ToUpper().Contains("/NGS/") Then
                        Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/ngs/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                    Else
                        Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/ngsTeste/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                    End If
                Else
                    Funcoes.AbrirArquivo(Me.Page, String.Format("{0}/{1}/{2}", domainServer, urlPath, fileInfo.Name))
                End If

            End If

        Next

        If chkZIP.Checked Then

            Funcoes.DownloadZIP(Me.Page, pathFiles)

        End If

    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioEntradaSaidaNotas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        LimparCampos()
    End Sub

    Protected Sub chkResumo_CheckedChanged(sender As Object, e As EventArgs) Handles chkResumo.CheckedChanged
        ddlTipoDeResumo.Visible = chkResumo.Checked
    End Sub

    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaClientes.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeClientes(Me, "objRepresentante" & HID.Value, txtNome.ClientID)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

End Class