Imports NGS.Lib.Negocio

Public Class RelatorioDeFretesExcel
    Inherits BasePage

#Region "Methods"

    Private Sub BuscaUnidade()
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
    End Sub

    Private Sub BuscarGruposProdutos()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub BuscarProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupo.SelectedValue & "'", True)
    End Sub

    Private Sub BuscarClientes()
        ddl.Carregar(ddlCliente, CarregarDDL.Tabela.Clientes, "", True)
    End Sub

    Private Function getTituloAba() As String
        If rdoFretesSemNota.Checked Then
            Return "Fretes sem notas."
        ElseIf rdoNotasSemFrete.Checked Then
            Return "Notas sem fretes."
        ElseIf rdoResumoCustos.Checked Then
            Return "Resumo dos custos."
        Else
            Return "Todos os fretes."
        End If
    End Function

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataSet() As DataSet
        Dim sql As String = ""

        If rdoTodosFretes.Checked Then
            sql = " SELECT CTRC.Empresa_Id AS CTRC_Empresa, CTRC.EndEmpresa_Id AS CTRC_EndEmpresa, CTRC.Cliente_Id AS CTRC_Transportador," & vbCrLf &
                  "       CTRC.EndCliente_Id AS CTRC_EndTransportador, CTRC.EntradaSaida_Id AS CTRC_ES, CTRC.Serie_Id AS CTRC_Serie, CTRC.Nota_Id AS CTRC_Numero," & vbCrLf &
                  "       CTRC.Movimento AS CTRC_Movimento, isnull(NotasFiscais.Movimento, '') as NF_Movimento, CTRC.DataDaNota AS CTRC_DataDeEmissao, CTRCXItens.CFOP_Id AS CTRC_Cfop," & vbCrLf &
                  "       Transportador.Nome AS CTRC_NomeTransportador, Transportador.Cidade AS CTRC_Cidade, Transportador.Estado AS CTRC_Estado," & vbCrLf &
                  "       CTRCXItens.QuantidadeFiscal AS CTRC_Peso, CTRCXItens.Unitario AS CTRC_Unitario, CTRCXItens.Valor AS CTRC_Valor, ISNULL(CTRCxEncargos.Percentual, 0)" & vbCrLf &
                  "       AS CTRC_Aliquota_Icms, ISNULL(CTRCxEncargos.Valor, 0) AS CTRC_Valor_ICMS, NotasFiscais.Empresa_Id AS Nota_Empresa," & vbCrLf &
                  "       NotasFiscais.EndEmpresa_Id AS Nota_EndEmpresa, NotasFiscais.Cliente_Id AS Nota_Cliente, NotasFiscais.EndCliente_Id AS Nota_EndCliente," & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id AS Nota_ES, NotasFiscais.Serie_Id AS Nota_Serie, NotasFiscais.Nota_Id AS Nota_Numero," & vbCrLf &
                  "       NotasFiscaisXItens.Produto_Id AS Nota_Produto, Produtos.Nome AS Nota_NomeProduto, NotasFiscaisXItens.Operacao AS Nota_Operacao," & vbCrLf &
                  "       NotasFiscaisXItens.SubOperacao AS Nota_SubOperacao, SubOperacoes.Descricao AS Nota_DescricaoOperacao, NotasFiscaisXItens.CFOP_Id AS Nota_CFOP," & vbCrLf &
                  "       Cfop.Descricao AS Nota_DescricaoCFOP, NotasFiscais.Pedido, Pedidos.FreteCIFFOB AS CIF_FOB, NotasFiscais.Deposito AS Nota_Origem_Deposito," & vbCrLf &
                  "       Origem.Nome AS Nota_Origem_Nome, Origem.Cidade AS Nota_Origem_Cidade, Origem.Estado AS Nota_Origem_Estado, Destino.Nome AS Nota_Destino_Nome," & vbCrLf &
                  "       Destino.Cidade AS Nota_Destino_Cidade, Destino.Estado AS Nota_Destino_Estado, NotasFiscaisXItens.PesoFiscal AS Nota_Peso," & vbCrLf &
                  "       NotasFiscaisXItens.Valor AS Nota_Valor, ISNULL(NotasFiscaisXEncargos.Percentual, 0) AS Nota_Aliquota_Icms, ISNULL(NotasFiscaisXEncargos.Valor, 0)" & vbCrLf &
                  "       AS Nota_Valor_ICMS, ISNULL(CTRCXItens.QuantidadeFiscal - NotasFiscaisXItens.PesoFiscal, 0) AS Diferenca_Peso" & vbCrLf &
                  " FROM  Clientes AS Destino RIGHT OUTER JOIN" & vbCrLf &
                  "       NotasFiscaisXEncargos RIGHT OUTER JOIN" & vbCrLf &
                  "       Produtos INNER JOIN" & vbCrLf &
                  "       NotasFiscais INNER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                  "       NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                  "       SubOperacoes ON NotasFiscaisXItens.Operacao = SubOperacoes.Operacao_Id AND" & vbCrLf &
                  "       NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
                  "       Cfop ON NotasFiscaisXItens.CFOP_Id = Cfop.Cfop_Id ON Produtos.Produto_Id = NotasFiscaisXItens.Produto_Id ON" & vbCrLf &
                  "       NotasFiscaisXEncargos.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscaisXEncargos.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscaisXEncargos.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscaisXEncargos.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                  "       NotasFiscaisXEncargos.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscaisXEncargos.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                  "       NotasFiscaisXEncargos.Nota_Id = NotasFiscaisXItens.Nota_Id AND NotasFiscaisXEncargos.Produto_Id = NotasFiscaisXItens.Produto_Id AND" & vbCrLf &
                  "       NotasFiscaisXEncargos.CFOP_Id = NotasFiscaisXItens.CFOP_Id AND NotasFiscaisXEncargos.Encargo_Id = 'ICMS' ON" & vbCrLf &
                  "       Destino.Cliente_Id = NotasFiscais.Destino AND Destino.Endereco_Id = NotasFiscais.EndDestino LEFT OUTER JOIN" & vbCrLf &
                  "       Clientes AS Origem ON NotasFiscais.Deposito = Origem.Cliente_Id AND NotasFiscais.EndDeposito = Origem.Endereco_Id LEFT OUTER JOIN" & vbCrLf &
                  "       Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Pedido = Pedidos.Pedido_Id FULL OUTER JOIN" & vbCrLf &
                  "       NotasXNotas ON NotasFiscais.Empresa_Id = NotasXNotas.OrigemEmpresa_Id AND NotasFiscais.EndEmpresa_Id = NotasXNotas.OrigemEndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Cliente_Id = NotasXNotas.OrigemCliente_Id AND NotasFiscais.EndCliente_Id = NotasXNotas.OrigemEndCliente_Id AND" & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id AND NotasFiscais.Serie_Id = NotasXNotas.OrigemSerie_Id AND" & vbCrLf &
                  "       NotasFiscais.Nota_Id = NotasXNotas.OrigemNota_Id FULL OUTER JOIN" & vbCrLf &
                  "       Clientes AS Transportador INNER JOIN" & vbCrLf &
                  "       NotasFiscais AS CTRC INNER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens AS CTRCXItens ON CTRC.Empresa_Id = CTRCXItens.Empresa_Id AND CTRC.EndEmpresa_Id = CTRCXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       CTRC.Cliente_Id = CTRCXItens.Cliente_Id AND CTRC.EndCliente_Id = CTRCXItens.EndCliente_Id AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id AND" & vbCrLf &
                  "       CTRC.Serie_Id = CTRCXItens.Serie_Id AND CTRC.Nota_Id = CTRCXItens.Nota_Id ON Transportador.Cliente_Id = CTRC.Cliente_Id AND" & vbCrLf &
                  "       Transportador.Endereco_Id = CTRC.EndCliente_Id LEFT OUTER JOIN" & vbCrLf &
                  "       NotasFiscaisXEncargos AS CTRCxEncargos ON CTRCXItens.Empresa_Id = CTRCxEncargos.Empresa_Id AND" & vbCrLf &
                  "       CTRCXItens.EndEmpresa_Id = CTRCxEncargos.EndEmpresa_Id AND CTRCXItens.Cliente_Id = CTRCxEncargos.Cliente_Id AND" & vbCrLf &
                  "       CTRCXItens.EndCliente_Id = CTRCxEncargos.EndCliente_Id AND CTRCXItens.EntradaSaida_Id = CTRCxEncargos.EntradaSaida_Id AND" & vbCrLf &
                  "       CTRCXItens.Serie_Id = CTRCxEncargos.Serie_Id AND CTRCXItens.Nota_Id = CTRCxEncargos.Nota_Id AND" & vbCrLf &
                  "       CTRCXItens.Produto_Id = CTRCxEncargos.Produto_Id AND CTRCXItens.CFOP_Id = CTRCxEncargos.CFOP_Id AND CTRCxEncargos.Encargo_Id = 'ICMS' ON" & vbCrLf &
                  "       NotasXNotas.Empresa_Id = CTRC.Empresa_Id AND NotasXNotas.EndEmpresa_Id = CTRC.EndEmpresa_Id AND NotasXNotas.Cliente_Id = CTRC.Cliente_Id AND" & vbCrLf &
                  "       NotasXNotas.EndCliente_Id = CTRC.EndCliente_Id AND NotasXNotas.EntradaSaida_Id = CTRC.EntradaSaida_Id AND NotasXNotas.Serie_Id = CTRC.Serie_Id AND" & vbCrLf &
                  "       NotasXNotas.Nota_Id = CTRC.Nota_Id" & vbCrLf &
                  " WHERE  CTRC.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' AND (CTRC.Situacao = 1) AND CTRCXItens.QuantidadeFiscal > 1" & vbCrLf &
                  "  AND (CTRCXItens.CFOP_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 6350 AND 6360)" & vbCrLf &
                  " And  (CTRC.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue) Then
                sql &= " And CTRC.Cliente_Id = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
                sql &= " And Left(NotasFiscaisXItens.Produto_Id, 5) = '" & ddlGrupo.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                sql &= " And NotasFiscaisXItens.Produto_Id = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            sql &= " GROUP BY CTRC.Empresa_Id, CTRC.EndEmpresa_Id, CTRC.Cliente_Id, CTRC.EndCliente_Id, CTRC.EntradaSaida_Id, CTRC.Serie_Id, CTRC.Nota_Id, CTRC.Movimento, NotasFiscais.Movimento, " & vbCrLf &
                   "        CTRC.DataDaNota, CTRCXItens.CFOP_Id, Transportador.Nome, Transportador.Cidade, Transportador.Estado, NotasFiscais.Empresa_Id," & vbCrLf &
                   "        NotasFiscais.EndEmpresa_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id," & vbCrLf &
                   "        NotasFiscaisXItens.Produto_Id, NotasFiscais.Pedido, NotasFiscais.Deposito, Origem.Nome, Origem.Cidade, Origem.Estado, Destino.Nome, Destino.Cidade," & vbCrLf &
                   "        Destino.Estado, Pedidos.FreteCIFFOB, CTRCXItens.QuantidadeFiscal, NotasFiscaisXItens.PesoFiscal, CTRCXItens.Unitario, CTRCXItens.Valor," & vbCrLf &
                   "        CTRCxEncargos.Encargo_Id, CTRCxEncargos.Percentual, CTRCxEncargos.Valor, NotasFiscaisXEncargos.Encargo_Id, NotasFiscaisXEncargos.Percentual," & vbCrLf &
                   "        NotasFiscaisXEncargos.Valor, NotasFiscaisXItens.Valor, NotasFiscaisXItens.Operacao, NotasFiscaisXItens.SubOperacao, SubOperacoes.Descricao," & vbCrLf &
                   "        Produtos.Nome , NotasFiscaisXItens.CFOP_Id, CFOP.Descricao" & vbCrLf &
                   " ORDER BY CTRC_Movimento" & vbCrLf

        ElseIf rdoFretesSemNota.Checked Then

            sql = " SELECT CTRC.Empresa_Id AS CTRC_Empresa, CTRC.EndEmpresa_Id AS CTRC_EndEmpresa, CTRC.Cliente_Id AS CTRC_Transportador," & vbCrLf &
                  "       CTRC.EndCliente_Id AS CTRC_EndTransportador, CTRC.EntradaSaida_Id AS CTRC_ES, CTRC.Serie_Id AS CTRC_Serie, CTRC.Nota_Id AS CTRC_Numero," & vbCrLf &
                  "       CTRC.Movimento AS CTRC_Movimento, CTRC.DataDaNota AS CTRC_DataDeEmissao, CTRCXItens.CFOP_Id AS CTRC_Cfop," & vbCrLf &
                  "       NotasFiscais.Empresa_Id AS Nota_Empresa, NotasFiscais.EndEmpresa_Id AS Nota_EndEmpresa, NotasFiscais.Cliente_Id AS Nota_Cliente," & vbCrLf &
                  "       NotasFiscais.EndCliente_Id AS Nota_EndCliente, NotasFiscais.EntradaSaida_Id AS Nota_ES, NotasFiscais.Serie_Id AS Nota_Serie," & vbCrLf &
                  "       NotasFiscais.Nota_Id AS Nota_Numero, NotasFiscaisXItens.Produto_Id AS Nota_Produto, NotasFiscais.Pedido, Pedidos.FreteCIFFOB AS CIF_FOB," & vbCrLf &
                  "       CTRCXItens.QuantidadeFiscal AS CTRC_Peso, NotasFiscaisXItens.PesoFiscal AS Nota_Peso, isnull((CTRCXItens.QuantidadeFiscal - NotasFiscaisXItens.PesoFiscal), 0) as Diferenca, CTRCXItens.Unitario AS CTRC_Unitario," & vbCrLf &
                  "       CTRCXItens.Valor AS CTRC_Valor" & vbCrLf &
                  " FROM  NotasFiscais AS CTRC INNER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens AS CTRCXItens ON CTRC.Empresa_Id = CTRCXItens.Empresa_Id AND CTRC.EndEmpresa_Id = CTRCXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       CTRC.Cliente_Id = CTRCXItens.Cliente_Id AND CTRC.EndCliente_Id = CTRCXItens.EndCliente_Id AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id AND" & vbCrLf &
                  "       CTRC.Serie_Id = CTRCXItens.Serie_Id AND CTRC.Nota_Id = CTRCXItens.Nota_Id FULL OUTER JOIN" & vbCrLf &
                  "       NotasXNotas FULL OUTER JOIN" & vbCrLf &
                  "       NotasFiscais INNER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                  "       NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id LEFT OUTER JOIN" & vbCrLf &
                  "       Pedidos ON NotasFiscais.Empresa_Id = Pedidos.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Pedido = Pedidos.Pedido_Id ON NotasXNotas.OrigemEmpresa_Id = NotasFiscais.Empresa_Id AND" & vbCrLf &
                  "       NotasXNotas.OrigemEndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND NotasXNotas.OrigemCliente_Id = NotasFiscais.Cliente_Id AND" & vbCrLf &
                  "       NotasXNotas.OrigemEndCliente_Id = NotasFiscais.EndCliente_Id AND NotasXNotas.OrigemEntradaSaida_Id = NotasFiscais.EntradaSaida_Id AND" & vbCrLf &
                  "       NotasXNotas.OrigemSerie_Id = NotasFiscais.Serie_Id AND NotasXNotas.OrigemNota_Id = NotasFiscais.Nota_Id ON" & vbCrLf &
                  "       CTRC.Empresa_Id = NotasXNotas.Empresa_Id AND CTRC.EndEmpresa_Id = NotasXNotas.EndEmpresa_Id AND CTRC.Cliente_Id = NotasXNotas.Cliente_Id AND" & vbCrLf &
                  "       CTRC.EndCliente_Id = NotasXNotas.EndCliente_Id AND CTRC.EntradaSaida_Id = NotasXNotas.EntradaSaida_Id AND CTRC.Serie_Id = NotasXNotas.Serie_Id AND" & vbCrLf &
                  "       CTRC.Nota_Id = NotasXNotas.Nota_Id" & vbCrLf &
                  " WHERE  CTRC.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' AND (CTRC.Situacao = 1) AND CTRCXItens.QuantidadeFiscal > 1 And NotasFiscaisXItens.Produto_Id is null" & vbCrLf &
                  "  AND (CTRCXItens.CFOP_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
                  "       CTRCXItens.CFOP_Id BETWEEN 6350 AND 6360)" & vbCrLf &
                  " And  (CTRC.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue) Then
                sql &= " And CTRC.Cliente_Id = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
                sql &= " And Left(NotasFiscaisXItens.Produto_Id, 5) = '" & ddlGrupo.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                sql &= " And NotasFiscaisXItens.Produto_Id = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            sql &= "Order by  CTRC.Movimento"

        ElseIf rdoNotasSemFrete.Checked Then

            sql = "SELECT CTRC.Empresa_Id AS CTRC_Empresa, CTRC.EndEmpresa_Id AS CTRC_EndEmpresa, CTRC.Cliente_Id AS CTRC_Transportador," & vbCrLf &
                  "       CTRC.EndCliente_Id AS CTRC_EndTransportador, CTRC.EntradaSaida_Id AS CTRC_ES, CTRC.Serie_Id AS CTRC_Serie, CTRC.Nota_Id AS CTRC_Numero," & vbCrLf &
                  "       isnull(CTRC.Movimento, '') AS CTRC_Movimento, isnull(CTRC.DataDaNota, '') AS CTRC_DataDeEmissao, CTRCXItens.CFOP_Id AS CTRC_Cfop," & vbCrLf &
                  "       NotasFiscais.Empresa_Id AS Nota_Empresa, NotasFiscais.EndEmpresa_Id AS Nota_EndEmpresa, NotasFiscais.Cliente_Id AS Nota_Cliente," & vbCrLf &
                  "       NotasFiscais.EndCliente_Id AS Nota_EndCliente, NotasFiscais.EntradaSaida_Id AS Nota_ES, NotasFiscais.Serie_Id AS Nota_Serie," & vbCrLf &
                  "       NotasFiscais.Nota_Id AS Nota_Numero, NotasFiscaisXItens.Produto_Id AS Nota_Produto, NotasFiscais.Pedido, Pedidos.FreteCIFFOB AS CIF_FOB," & vbCrLf &
                  "       CTRCXItens.QuantidadeFiscal AS CTRC_Peso, NotasFiscaisXItens.PesoFiscal AS Nota_Peso," & vbCrLf &
                  "       ISNULL(CTRCXItens.QuantidadeFiscal - NotasFiscaisXItens.PesoFiscal, 0) AS Diferenca, CTRCXItens.Unitario AS CTRC_Unitario," & vbCrLf &
                  "       CTRCXItens.Valor AS CTRC_Valor" & vbCrLf &
                  " FROM  NotasFiscais INNER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                  "       NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                  "       Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf &
                  "       SubOperacoes INNER JOIN" & vbCrLf &
                  "       Pedidos ON SubOperacoes.Operacao_Id = Pedidos.Operacao AND SubOperacoes.SubOperacoes_Id = Pedidos.SubOperacao ON" & vbCrLf &
                  "       NotasFiscais.Empresa_Id = Pedidos.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Pedidos.EndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Pedido = Pedidos.Pedido_Id FULL OUTER JOIN" & vbCrLf &
                  "       NotasXNotas ON NotasFiscais.Empresa_Id = NotasXNotas.OrigemEmpresa_Id AND NotasFiscais.EndEmpresa_Id = NotasXNotas.OrigemEndEmpresa_Id AND" & vbCrLf &
                  "       NotasFiscais.Cliente_Id = NotasXNotas.OrigemCliente_Id AND NotasFiscais.EndCliente_Id = NotasXNotas.OrigemEndCliente_Id AND" & vbCrLf &
                  "       NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id AND NotasFiscais.Serie_Id = NotasXNotas.OrigemSerie_Id AND" & vbCrLf &
                  "       NotasFiscais.Nota_Id = NotasXNotas.OrigemNota_Id FULL OUTER JOIN" & vbCrLf &
                  "       NotasFiscais AS CTRC LEFT OUTER JOIN" & vbCrLf &
                  "       NotasFiscaisXItens AS CTRCXItens ON CTRC.Empresa_Id = CTRCXItens.Empresa_Id AND CTRC.EndEmpresa_Id = CTRCXItens.EndEmpresa_Id AND" & vbCrLf &
                  "       CTRC.Cliente_Id = CTRCXItens.Cliente_Id AND CTRC.EndCliente_Id = CTRCXItens.EndCliente_Id AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id AND" & vbCrLf &
                  "       CTRC.Serie_Id = CTRCXItens.Serie_Id AND CTRC.Nota_Id = CTRCXItens.Nota_Id ON NotasXNotas.Empresa_Id = CTRC.Empresa_Id AND" & vbCrLf &
                  "       NotasXNotas.EndEmpresa_Id = CTRC.EndEmpresa_Id AND NotasXNotas.Cliente_Id = CTRC.Cliente_Id AND NotasXNotas.EndCliente_Id = CTRC.EndCliente_Id AND" & vbCrLf &
                  "       NotasXNotas.EntradaSaida_Id = CTRC.EntradaSaida_Id And NotasXNotas.Serie_Id = CTRC.Serie_Id And NotasXNotas.Nota_Id = CTRC.Nota_Id" & vbCrLf &
                  "  WHERE (NotasFiscais.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "')  AND (NotasFiscais.Situacao = 1)  And Produtos.Agrupar = 'N' And isnull(CTRC.Nota_Id, 0) < 1" & vbCrLf &
                  "    AND (NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') AND" & vbCrLf &
                  "       ( (Pedidos.FreteCIFFOB = 'FOB' AND SubOperacoes.EntradaSaida = 'E') OR  (Pedidos.FreteCIFFOB = 'CIF' AND SubOperacoes.EntradaSaida = 'S') )" & vbCrLf

        ElseIf rdoResumoCustos.Checked Then

            sql = " SELECT     CTRC.Empresa_Id As Empresa, Clientes.Nome AS NomeDaEmpresa, Clientes.Cidade AS CidadeDaEmpresa,  NotasFiscais.Deposito,  NotasFiscais.EndDeposito," & vbCrLf &
                  "      Deposito.Nome AS NomeDoDeposito, Deposito.Cidade as CidadeDoDeposito," & vbCrLf

            If chkDataNota.Checked Then
                sql &= "       MONTH(isnull(NotasFiscais.Movimento, '')) AS Mes," & vbCrLf
            Else
                sql &= "       MONTH(CTRC.Movimento) AS Mes," & vbCrLf
            End If

            sql &= "                 NotasFiscais.EntradaSaida_Id AS ES, NotasFiscaisXItens.Produto_Id AS Produto, Produtos.Nome, NotasFiscaisXItens.Operacao, NotasFiscaisXItens.SubOperacao," & vbCrLf &
                   "                 SubOperacoes.Descricao, SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Produto' THEN NotasFiscaisXItens.PesoFiscal ELSE 0 END) AS PesoTransportado," & vbCrLf &
                   "                 SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Produto' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS ValordoFrete," & vbCrLf &
                   "                 SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Icms' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Icms," & vbCrLf &
                   "                 SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Pis' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Pis," & vbCrLf &
                   "                 SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Cofins' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Cofins" & vbCrLf &
                   " FROM         SubOperacoes INNER JOIN" & vbCrLf &
                   "                  NotasFiscais INNER JOIN" & vbCrLf &
                   "                  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                   "                  NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                   "                  NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                   "                  NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                   "                  Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao AND" & vbCrLf &
                   "                  SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao INNER JOIN" & vbCrLf &
                   "                  Clientes ON NotasFiscais.Empresa_Id = Clientes.Cliente_Id AND NotasFiscais.EndEmpresa_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf &
                   "                  Clientes AS Deposito ON NotasFiscais.Deposito = Deposito.Cliente_Id AND NotasFiscais.EndDeposito = Deposito.Endereco_Id RIGHT OUTER JOIN" & vbCrLf &
                   "                  NotasXNotas ON NotasFiscais.Empresa_Id = NotasXNotas.OrigemEmpresa_Id AND NotasFiscais.EndEmpresa_Id = NotasXNotas.OrigemEndEmpresa_Id AND" & vbCrLf &
                   "                  NotasFiscais.Cliente_Id = NotasXNotas.OrigemCliente_Id AND NotasFiscais.EndCliente_Id = NotasXNotas.OrigemEndCliente_Id AND" & vbCrLf &
                   "                  NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id AND NotasFiscais.Serie_Id = NotasXNotas.OrigemSerie_Id AND" & vbCrLf &
                   "                  NotasFiscais.Nota_Id = NotasXNotas.OrigemNota_Id RIGHT OUTER JOIN" & vbCrLf &
                   "                  NotasFiscais AS CTRC INNER JOIN" & vbCrLf &
                   "                   NotasFiscaisXItens AS CTRCXItens ON CTRC.Empresa_Id = CTRCXItens.Empresa_Id AND CTRC.EndEmpresa_Id = CTRCXItens.EndEmpresa_Id AND" & vbCrLf &
                   "                  CTRC.Cliente_Id = CTRCXItens.Cliente_Id AND CTRC.EndCliente_Id = CTRCXItens.EndCliente_Id AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id AND" & vbCrLf &
                   "                  CTRC.Serie_Id = CTRCXItens.Serie_Id AND CTRC.Nota_Id = CTRCXItens.Nota_Id INNER JOIN" & vbCrLf &
                   "                  NotasFiscaisXEncargos AS CTRCXItensXEncargos ON CTRCXItens.Empresa_Id = CTRCXItensXEncargos.Empresa_Id AND" & vbCrLf &
                   "                  CTRCXItens.EndEmpresa_Id = CTRCXItensXEncargos.EndEmpresa_Id AND CTRCXItens.Cliente_Id = CTRCXItensXEncargos.Cliente_Id AND" & vbCrLf &
                   "                  CTRCXItens.EndCliente_Id = CTRCXItensXEncargos.EndCliente_Id AND CTRCXItens.EntradaSaida_Id = CTRCXItensXEncargos.EntradaSaida_Id AND" & vbCrLf &
                   "                  CTRCXItens.Serie_Id = CTRCXItensXEncargos.Serie_Id AND CTRCXItens.Nota_Id = CTRCXItensXEncargos.Nota_Id AND" & vbCrLf &
                   "                  CTRCXItens.Produto_Id = CTRCXItensXEncargos.Produto_Id AND CTRCXItens.CFOP_Id = CTRCXItensXEncargos.CFOP_Id ON" & vbCrLf &
                   "                  NotasXNotas.Empresa_Id = CTRC.Empresa_Id AND NotasXNotas.EndEmpresa_Id = CTRC.EndEmpresa_Id AND NotasXNotas.Cliente_Id = CTRC.Cliente_Id AND" & vbCrLf &
                   "                  NotasXNotas.EndCliente_Id = CTRC.EndCliente_Id AND NotasXNotas.EntradaSaida_Id = CTRC.EntradaSaida_Id AND NotasXNotas.Serie_Id = CTRC.Serie_Id AND" & vbCrLf &
                   "                  NotasXNotas.Nota_Id = CTRC.Nota_Id" & vbCrLf &
                   " WHERE     CTRC.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'  AND (CTRC.Situacao = 1) AND CTRCXItens.QuantidadeFiscal > 1 And (CTRCXItens.CFOP_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 6350 AND 6360)" & vbCrLf

            If chkDataNota.Checked Then
                sql &= " And  (NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
            Else
                sql &= " And  (CTRC.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
            End If

            sql &= "                  And (CTRCXItensXEncargos.Encargo_Id NOT IN ('Liquido')) AND (Produtos.Agrupar = 'N')  And NotasFiscais.EntradaSaida_Id = 'E'" & vbCrLf &
                   " GROUP BY CTRC.Empresa_Id, " & vbCrLf

            If chkDataNota.Checked Then
                sql &= " MONTH(NotasFiscais.Movimento), NotasFiscais.Movimento, " & vbCrLf
            Else
                sql &= " MONTH(CTRC.Movimento), CTRC.Movimento, " & vbCrLf
            End If

            sql &= " NotasFiscais.EntradaSaida_Id , NotasFiscaisXItens.Produto_Id, NotasFiscaisXItens.Operacao, " & vbCrLf &
                   "                  NotasFiscaisXItens.SubOperacao, SubOperacoes.Descricao, Produtos.Nome, " & vbCrLf &
                   "                  Clientes.Nome, Clientes.Cidade, NotasFiscais.Deposito, NotasFiscais.EndDeposito," & vbCrLf &
                   "                  Deposito.Nome , Deposito.Cidade" & vbCrLf &
                   "" & vbCrLf &
                   " Union" & vbCrLf &
                   "" & vbCrLf &
                   " SELECT     CTRC.Empresa_Id as Empresa, Clientes.Nome AS NomeDaEmpresa, Clientes.Cidade AS CidadeDaEmpresa,  " & vbCrLf &
                   "            NotasFiscais.Destino as Deposito, NotasFiscais.EndDestino as EndDeposito," & vbCrLf &
                   "                  Deposito.Nome AS NomeDoDeposito," & vbCrLf &
                   "                  Deposito.Cidade AS CidadeDoDeposito, " & vbCrLf

            If chkDataNota.Checked Then
                sql &= " MONTH(isnull(NotasFiscais.Movimento, '')) AS Mes," & vbCrLf
            Else
                sql &= " MONTH(CTRC.Movimento) AS Mes," & vbCrLf
            End If

            sql &= "                  NotasFiscais.EntradaSaida_Id AS ES, NotasFiscaisXItens.Produto_Id AS Produto," & vbCrLf &
                   "                  Produtos.Nome, NotasFiscaisXItens.Operacao, NotasFiscaisXItens.SubOperacao, SubOperacoes.Descricao," & vbCrLf &
                   "                  SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Produto' THEN NotasFiscaisXItens.PesoFiscal ELSE 0 END) AS PesoTransportado," & vbCrLf &
                   "                  SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Produto' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS ValordoFrete," & vbCrLf &
                   "                  SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Icms' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Icms," & vbCrLf &
                   "                  SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Pis' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Pis," & vbCrLf &
                   "                  SUM(CASE WHEN CTRCXItensXEncargos.Encargo_Id = 'Cofins' THEN CTRCXItensXEncargos.Valor ELSE 0 END) AS Cofins" & vbCrLf &
                   " FROM         SubOperacoes INNER JOIN" & vbCrLf &
                   "                  NotasFiscais INNER JOIN" & vbCrLf &
                   "                  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf &
                   "                  NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf &
                   "                  NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf &
                   "                  NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf &
                   "                  Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id ON SubOperacoes.Operacao_Id = NotasFiscaisXItens.Operacao AND" & vbCrLf &
                   "                  SubOperacoes.SubOperacoes_Id = NotasFiscaisXItens.SubOperacao INNER JOIN" & vbCrLf &
                   "                  Clientes ON NotasFiscais.Empresa_Id = Clientes.Cliente_Id AND NotasFiscais.EndEmpresa_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf &
                   "                  Clientes AS Deposito ON NotasFiscais.Destino = Deposito.Cliente_Id AND NotasFiscais.EndDestino = Deposito.Endereco_Id RIGHT OUTER JOIN" & vbCrLf &
                   "                  NotasXNotas ON NotasFiscais.Empresa_Id = NotasXNotas.OrigemEmpresa_Id AND NotasFiscais.EndEmpresa_Id = NotasXNotas.OrigemEndEmpresa_Id AND" & vbCrLf &
                   "                  NotasFiscais.Cliente_Id = NotasXNotas.OrigemCliente_Id AND NotasFiscais.EndCliente_Id = NotasXNotas.OrigemEndCliente_Id AND" & vbCrLf &
                   "                  NotasFiscais.EntradaSaida_Id = NotasXNotas.OrigemEntradaSaida_Id AND NotasFiscais.Serie_Id = NotasXNotas.OrigemSerie_Id AND" & vbCrLf &
                   "                  NotasFiscais.Nota_Id = NotasXNotas.OrigemNota_Id RIGHT OUTER JOIN" & vbCrLf &
                   "                  NotasFiscais AS CTRC INNER JOIN" & vbCrLf &
                   "                  NotasFiscaisXItens AS CTRCXItens ON CTRC.Empresa_Id = CTRCXItens.Empresa_Id AND CTRC.EndEmpresa_Id = CTRCXItens.EndEmpresa_Id AND" & vbCrLf &
                   "                  CTRC.Cliente_Id = CTRCXItens.Cliente_Id AND CTRC.EndCliente_Id = CTRCXItens.EndCliente_Id AND CTRC.EntradaSaida_Id = CTRCXItens.EntradaSaida_Id AND" & vbCrLf &
                   "                  CTRC.Serie_Id = CTRCXItens.Serie_Id AND CTRC.Nota_Id = CTRCXItens.Nota_Id INNER JOIN" & vbCrLf &
                   "                  NotasFiscaisXEncargos AS CTRCXItensXEncargos ON CTRCXItens.Empresa_Id = CTRCXItensXEncargos.Empresa_Id AND" & vbCrLf &
                   "                  CTRCXItens.EndEmpresa_Id = CTRCXItensXEncargos.EndEmpresa_Id AND CTRCXItens.Cliente_Id = CTRCXItensXEncargos.Cliente_Id AND" & vbCrLf &
                   "                  CTRCXItens.EndCliente_Id = CTRCXItensXEncargos.EndCliente_Id AND CTRCXItens.EntradaSaida_Id = CTRCXItensXEncargos.EntradaSaida_Id AND" & vbCrLf &
                   "                  CTRCXItens.Serie_Id = CTRCXItensXEncargos.Serie_Id AND CTRCXItens.Nota_Id = CTRCXItensXEncargos.Nota_Id AND" & vbCrLf &
                   "                  CTRCXItens.Produto_Id = CTRCXItensXEncargos.Produto_Id AND CTRCXItens.CFOP_Id = CTRCXItensXEncargos.CFOP_Id ON" & vbCrLf &
                   "                  NotasXNotas.Empresa_Id = CTRC.Empresa_Id AND NotasXNotas.EndEmpresa_Id = CTRC.EndEmpresa_Id AND NotasXNotas.Cliente_Id = CTRC.Cliente_Id AND" & vbCrLf &
                   "                  NotasXNotas.EndCliente_Id = CTRC.EndCliente_Id AND NotasXNotas.EntradaSaida_Id = CTRC.EntradaSaida_Id AND NotasXNotas.Serie_Id = CTRC.Serie_Id AND" & vbCrLf &
                   "                  NotasXNotas.Nota_Id = CTRC.Nota_Id" & vbCrLf &
                   " WHERE     CTRC.Empresa_ID = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'  AND (CTRC.Situacao = 1) AND CTRCXItens.QuantidadeFiscal > 1 AND (CTRCXItens.CFOP_Id BETWEEN 1350 AND 1360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 2350 AND 2360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 5350 AND 5360 OR" & vbCrLf &
                   "                  CTRCXItens.CFOP_Id BETWEEN 6350 AND 6360) " & vbCrLf

            If chkDataNota.Checked Then
                sql &= " And  (NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
            Else
                sql &= " And  (CTRC.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
            End If

            sql &= "                  And (CTRCXItensXEncargos.Encargo_Id NOT IN ('Liquido'))" & vbCrLf &
                   "                  AND (Produtos.Agrupar = 'N') AND (NotasFiscais.EntradaSaida_Id = 'S')" & vbCrLf &
                   " GROUP BY CTRC.Empresa_Id, " & vbCrLf

            If chkDataNota.Checked Then
                sql &= " MONTH(NotasFiscais.Movimento), NotasFiscais.Movimento, " & vbCrLf
            Else
                sql &= " MONTH(CTRC.Movimento), CTRC.Movimento, " & vbCrLf
            End If

            sql &= "  NotasFiscais.EntradaSaida_Id , NotasFiscaisXItens.Produto_Id, NotasFiscaisXItens.Operacao, " & vbCrLf &
                   "                  NotasFiscaisXItens.SubOperacao, SubOperacoes.Descricao, Produtos.Nome, Clientes.Nome, Clientes.Cidade, Deposito.Nome, Deposito.Cidade," & vbCrLf &
                   "                  NotasFiscais.Destino , NotasFiscais.EndDestino" & vbCrLf

            sql &= " ORDER BY ES, Mes"
        End If

        Return Banco.ConsultaDataSet(sql, "Fretes")
    End Function

    Private Sub Limpar()
        BuscaUnidade()
        ddlCliente.SelectedValue = ""
        ddlGrupo.SelectedValue = ""
        ddlProduto.Items.Clear()
        rdoTodosFretes.Checked = True
        chkDataNota.Checked = False
        txtDataInicial.Text = New DateTime(Now.Year, Now.Month, 1)
        txtDataFinal.Text = New DateTime(Now.Year, Now.Month, Now.Day)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeFretesExcel", "ACESSAR") Then
                    BuscaUnidade()
                    BuscarClientes()
                    BuscarGruposProdutos()
                    txtDataInicial.Text = New DateTime(Now.Year, Now.Month, 1)
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupo.SelectedIndexChanged
        Try
            BuscarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged1(sender As Object, e As EventArgs) Handles DdlUnidade.SelectedIndexChanged
        Try
            BuscaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim ds As DataSet = getDataSet()

            Funcoes.BindExcelOffice(Me.Page, ds, "Relatorio de Fretes")

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeFretesExcel")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class