Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelatorioDeNotasFiscais
    Inherits BasePage

#Region "Methods"

    Private Sub CarregaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Private Sub CarregaClientes()
        ddl.Carregar(ddlCliente, CarregarDDL.Tabela.Clientes, "", True)
    End Sub

    Private Sub CarregaDeposito()
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.Depositos, "", True)
        ddl.Carregar(ddlDepositoDestino, CarregarDDL.Tabela.Depositos, "", True)
    End Sub

    Private Sub CarregaGrupoDeProduto()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub CarregaProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.ProdutoPorGrupo, ddlGrupo.SelectedValue, True)
    End Sub

    Private Sub CarregaOperacoes()
        ddl.Carregar(ddlOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub BuscarSubOperacoes()
        If ddlOperacao.SelectedIndex > 0 Then
            ddlSubOperacao.Items.Clear()
            Dim obj As New [Lib].Negocio.ListSubOperacao()
            If obj.Selecionar(ddlOperacao.SelectedValue) Then
                For Each objSubOperacao As [Lib].Negocio.SubOperacao In obj
                    ddlSubOperacao.Items.Add(New ListItem(objSubOperacao.Codigo.ToString("00") & " - " & objSubOperacao.Descricao, _
                                                          objSubOperacao.Codigo))
                Next
                If obj IsNot Nothing Then
                    For Each li As ListItem In ddlSubOperacao.Items
                        li.Attributes("title") = li.Text
                    Next
                End If
                Funcoes.InserirLinhaEmBranco(ddlSubOperacao)
            Else
                MsgBox(Me.Page, obj.Erro.Message)
            End If
        Else
            ddlSubOperacao.Items.Clear()
        End If
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = ""

        If Not rdoLaudoSemNota.Checked Then
            sql = " SELECT     Empresa.Reduzido AS RedEmpresa, NotasFiscais.Empresa_Id AS CNPJEmpresa, NotasFiscais.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf & _
            "                Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa, NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.EndCliente_Id AS EndCliente," & vbCrLf & _
            "                Clientes.Nome AS NomeCliente, Clientes.Cidade AS CidadeCliente, Clientes.Estado AS UFCliente, NotasFiscais.Operacao, NotasFiscais.SubOperacao," & vbCrLf & _
            "                convert(varchar, NotasFiscais.Operacao) + '_' + convert(varchar, NotasFiscais.SubOperacao) as GrupoOperacao," & vbCrLf & _
            "                SubOperacoes.Descricao AS NomeOperacao, NotasFiscais.Deposito, NotasFiscais.EndDeposito, Deposito.Nome AS NomeDeposito," & vbCrLf & _
            "                Deposito.Cidade AS CidadeDeposito, Deposito.Estado AS UFDeposito, NotasFiscais.Destino AS DepositoDestino, NotasFiscais.EndDestino AS EndDepositoDestino," & vbCrLf & _
            "                DepositoDestino.Nome AS NomeDepositoDestino, DepositoDestino.Cidade AS CidadeDepositoDestino, DepositoDestino.Estado AS UFDepositoDestino," & vbCrLf & _
            "                GruposDeEstoques.Grupo_Id as Grupo, GruposDeEstoques.Descricao as NomeDoGrupo, NotasFiscaisXItens.Produto_Id AS Produto, Produtos.Nome AS NomeProduto, " & vbCrLf & _
            "                NotasFiscais.EntradaSaida_Id AS E_S, NotasFiscais.Serie_Id AS SerieDaNota," & vbCrLf & _
            "                NotasFiscais.Nota_Id AS NumeroDaNota, Romaneios.Romaneio_Id as NumeroDoRomaneio, RomaneiosXPesagens.Pesagem_Id as NumeroDoTicket, NotasFiscais.Pedido, " & vbCrLf & _
            "                NotasFiscais.Movimento AS MovimentoDaNota, ISNULL(Romaneios.Movimento, NotasFiscais.Movimento)" & vbCrLf & _
            "                AS MovimentoDoRomaneio, NotasFiscais.DataDaNota, NotasFiscais.NossaEmissao, NotasFiscais.CIFFOB, ISNULL(NotasFiscais.LocalEmbarque, '')" & vbCrLf & _
            "                AS LocalEmbarque, ISNULL(NotasFiscais.EndLocalEmbarque, 0) AS EndLocalEmbarque, ISNULL(LocalDeEmbarque.Nome, '') AS NomeLocalEmbarque," & vbCrLf & _
            "                ISNULL(LocalDeEmbarque.Cidade, '') AS CidadeLocalEmbarque, ISNULL(LocalDeEmbarque.Estado, '') AS UFLocalEmbarque, NotasFiscais.UsuarioInclusao," & vbCrLf & _
            "                 NotasFiscais.Autorizacao, NotasFiscaisXItens.CFOP_Id, Cfop.Descricao AS NomeCFOP, Romaneios.PesoBruto," & vbCrLf & _
            "                NotasFiscaisXItens.QuantidadeFiscal - Romaneios.PesoBruto AS DifBalanca, Romaneios.Desconto, NotasFiscaisXItens.QuantidadeFisica," & vbCrLf & _
            "                ISNULL(Romaneios.PesoLiquido, 0) AS Romaneio, NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor AS ValorDoProduto," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'ICMS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS ICMS," & vbCrLf & _
            "                SUM(CASE WHEN (NotasFiscaisXEncargos.Encargo_Id = 'FUNRURAL' or NotasFiscaisXEncargos.Encargo_Id = 'FUNRURAL JUDICIAL' or NotasFiscaisXEncargos.Encargo_Id = 'SENAR') THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS FUNRURAL," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'FETHAB' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS FETHAB," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS PIS," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'PIS CRED PRE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS PISCRE," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS COFINS," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'COFINS CRED PRE' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS COFINSCRE," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'IRRF PJ' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS IRRF," & vbCrLf & _
            "                SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_Id = 'LIQUIDO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) AS VALORLIQUIDO," & vbCrLf & _
            "                Comissoes.Representante_Id as CNPJ_Representante, Comissoes.EndRepresentante_Id as EndRenpresentante, Representante.Nome as NomeRepresentante," & vbCrLf & _
            "                isnull(Comissoes.Percentual, 0) as PercentualComissao," & vbCrLf & _
            "                case when NotasFiscais.Operacao <> 36 then NotasFiscaisXItens.Valor * ISNULL(Comissoes.Percentual, 0) / 100" & vbCrLf & _
            "                else ((NotasFiscaisXItens.QuantidadeFiscal / 1000) * ISNULL(Comissoes.Percentual, 0) ) * Cotacoes.Indice end AS ValordaComissao," & vbCrLf & _
            "                Cotacoes.Indice" & vbCrLf

            If chkObservacoesFiscais.Checked = True Then
                sql &= "              , convert(varchar(max), NotasFiscais.Observacoes) as Observacoes"
            End If

            sql &= " FROM     NotasFiscaisXRomaneios nfxr LEFT OUTER JOIN" & vbCrLf & _
                   "          Romaneios LEFT OUTER JOIN" & vbCrLf & _
                   "          RomaneiosXPesagens ON Romaneios.Empresa_Id = RomaneiosXPesagens.Empresa_Id AND" & vbCrLf & _
                   "          Romaneios.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id AND Romaneios.Romaneio_Id = RomaneiosXPesagens.Romaneio_Id ON" & vbCrLf & _
                   "          nfxr.Empresa_Id = Romaneios.Empresa_Id AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id AND" & vbCrLf & _
                   "          nfxr.Romaneio_Id = Romaneios.Romaneio_Id RIGHT OUTER JOIN" & vbCrLf & _
                   "          GruposDeEstoques INNER JOIN" & vbCrLf & _
                   "          NotasFiscais INNER JOIN" & vbCrLf & _
                   "          Clientes AS Empresa ON NotasFiscais.Empresa_Id = Empresa.Cliente_Id AND NotasFiscais.EndEmpresa_Id = Empresa.Endereco_Id INNER JOIN" & vbCrLf & _
                   "          Clientes AS Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
                   "          SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf & _
                   "          NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf & _
                   "          NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
                   "          NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
                   "          NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                   "          Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id INNER JOIN" & vbCrLf & _
                   "          NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf & _
                   "          NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf & _
                   "          NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND" & vbCrLf & _
                   "          NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND" & vbCrLf & _
                   "          NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf & _
                   "          NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id AND NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id ON" & vbCrLf & _
                   "          GruposDeEstoques.Grupo_Id = Produtos.Grupo LEFT OUTER JOIN" & vbCrLf & _
                   "          Cotacoes ON NotasFiscais.Movimento = Cotacoes.Data_Id AND Cotacoes.Indexador_Id = 3 LEFT OUTER JOIN" & vbCrLf & _
                   "          Comissoes INNER JOIN" & vbCrLf & _
                   "          Clientes AS Representante ON Comissoes.Representante_Id = Representante.Cliente_Id AND Comissoes.EndRepresentante_Id = Representante.Endereco_Id ON" & vbCrLf & _
                   "          NotasFiscais.Empresa_Id = Comissoes.Empresa_Id AND NotasFiscais.EndEmpresa_Id = Comissoes.EndEmpresa_Id AND" & vbCrLf & _
                   "          NotasFiscais.Pedido = Comissoes.Pedido_Id ON nfxr.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND" & vbCrLf & _
                   "          nfxr.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND nfxr.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND" & vbCrLf & _
                   "          nfxr.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
                   "          nfxr.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND nfxr.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
                   "          nfxr.Nota_Id = NotasFiscaisXItens.Nota_Id  LEFT OUTER JOIN" & vbCrLf & _
                   "          Cfop ON NotasFiscaisXItens.CFOP_Id = Cfop.Cfop_Id LEFT OUTER JOIN" & vbCrLf & _
                   "          Clientes AS LocalDeEmbarque ON NotasFiscais.LocalEmbarque = LocalDeEmbarque.Cliente_Id AND" & vbCrLf & _
                   "          NotasFiscais.EndLocalEmbarque = LocalDeEmbarque.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                   "          Clientes AS DepositoDestino ON NotasFiscais.Destino = DepositoDestino.Cliente_Id AND" & vbCrLf & _
                   "          NotasFiscais.EndDestino = DepositoDestino.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                   "          Clientes AS Deposito ON NotasFiscais.Deposito = Deposito.Cliente_Id AND NotasFiscais.EndDeposito = Deposito.Endereco_Id" & vbCrLf & _
                   " WHERE    NotasFiscais.Empresa_ID = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue.Split("-")(0)) Then
                sql &= " And NotasFiscais.Cliente_ID = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlDeposito.SelectedValue.Split("-")(0)) Then
                sql &= " And NotasFiscais.Deposito = '" & ddlDeposito.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlDepositoDestino.SelectedValue.Split("-")(0)) Then
                sql &= " And NotasFiscais.Destino = '" & ddlDepositoDestino.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
                sql &= " And Left(NotasFiscaisXItens.Produto_Id, 5) = '" & Left(ddlGrupo.SelectedValue, 5) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                sql &= " And NotasFiscaisXItens.Produto_Id = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue) Then
                sql &= " And NotasFiscais.Operacao = " & ddlOperacao.SelectedValue & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlSubOperacao.SelectedValue) Then
                sql &= " And NotasFiscais.SubOperacao = " & ddlSubOperacao.SelectedValue & vbCrLf
            End If

            If Not chkIncluiNotasProdutosDiversos.Checked Then
                sql &= " And (Produtos.Agrupar = 'N') " & vbCrLf
            End If

            sql &= " And (NotasFiscais.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') AND (NotasFiscais.Situacao = 1)" & vbCrLf

            If rdoNotasSemRomaneio.Checked Then
                sql &= " AND (isnull(Romaneios.PesoLiquido, 0) = 0 OR MONTH(NotasFiscais.Movimento) <> Month(Romaneios.Movimento))  And SubOperacoes.EstoqueFisico = 'S'" & vbCrLf
            End If

            sql &= " GROUP BY Empresa.Reduzido, NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, Empresa.Nome, Empresa.Cidade, Empresa.Estado, NotasFiscais.Cliente_Id," & vbCrLf & _
                   "                  NotasFiscais.EndCliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, NotasFiscais.Operacao, NotasFiscais.SubOperacao, SubOperacoes.Descricao," & vbCrLf & _
                   "                  NotasFiscais.Deposito, NotasFiscais.EndDeposito, Deposito.Nome, Deposito.Cidade, Deposito.Estado, NotasFiscais.Destino, NotasFiscais.EndDestino," & vbCrLf & _
                   "                  DepositoDestino.Nome, DepositoDestino.Cidade, DepositoDestino.Estado, NotasFiscaisXItens.Produto_Id, Produtos.Nome, NotasFiscais.EntradaSaida_Id," & vbCrLf & _
                   "                  NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, NotasFiscais.Pedido, NotasFiscais.Movimento, Romaneios.Movimento, NotasFiscais.DataDaNota," & vbCrLf & _
                   "                  NotasFiscais.NossaEmissao, NotasFiscais.CIFFOB, NotasFiscais.LocalEmbarque, NotasFiscais.EndLocalEmbarque, NotasFiscais.UsuarioInclusao," & vbCrLf & _
                   "                  NotasFiscais.Autorizacao, NotasFiscaisXItens.CFOP_Id, NotasFiscaisXItens.QuantidadeFisica, NotasFiscaisXItens.QuantidadeFiscal, NotasFiscaisXItens.Unitario," & vbCrLf & _
                   "                  NotasFiscaisXItens.Valor, LocalDeEmbarque.Nome, LocalDeEmbarque.Cidade, LocalDeEmbarque.Estado, Cfop.Descricao, Romaneios.PesoLiquido," & vbCrLf & _
                   "                  Romaneios.Desconto , Romaneios.PesoBruto, Comissoes.Representante_Id, Comissoes.EndRepresentante_Id, Representante.Nome, Comissoes.Percentual, Cotacoes.Indice, " & vbCrLf & _
                   "                  GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao, Romaneios.Romaneio_Id, RomaneiosXPesagens.Pesagem_Id" & vbCrLf

            If chkObservacoesFiscais.Checked Then
                sql &= "              , convert(varchar(max), NotasFiscais.Observacoes)" & vbCrLf
            End If

        Else
            sql = " SELECT Empresa.Reduzido AS RedEmpresa, Pesagem.Empresa_Id AS CNPJEmpresa, Pesagem.EndEmpresa_Id AS EndEmpresa, Empresa.Nome AS NomeEmpresa," & vbCrLf & _
                  "        Empresa.Cidade AS CidadeEmpresa, Empresa.Estado AS UFEmpresa, Pesagem.Cliente, Pesagem.EndCliente, Clientes.Nome AS NomeCliente," & vbCrLf & _
                  "        Clientes.Cidade AS CidadeCliente, Clientes.Estado AS UFCliente, Pesagem.Operacao, Pesagem.SubOperacao, SubOperacoes.Descricao AS NomeOperacao," & vbCrLf & _
                  "        Produtos.Nome AS NomeProduto, Pesagem.Pesagem_Id AS Laudo, ISNULL(NotasFiscais.Nota_Id, 0) AS Nota, Pesagem.Liquido," & vbCrLf & _
                  "        Pesagem.Movimento AS DataLaudo, Romaneios.Movimento AS DataRomaneio, ISNULL(NotasFiscais.Movimento, '') AS DataDaNota, Pesagem.Placa," & vbCrLf & _
                  "        Pesagem.UsuarioInclusao" & vbCrLf & _
                  "" & vbCrLf & _
                  " FROM   NotasFiscaisXRomaneios nfxr RIGHT OUTER JOIN" & vbCrLf & _
                  "        Clientes AS Clientes RIGHT OUTER JOIN" & vbCrLf & _
                  "        RomaneiosXPesagens LEFT OUTER JOIN" & vbCrLf & _
                  "        Romaneios ON RomaneiosXPesagens.Romaneio_Id = Romaneios.Romaneio_Id AND RomaneiosXPesagens.Empresa_Id = Romaneios.Empresa_Id AND" & vbCrLf & _
                  "        RomaneiosXPesagens.EndEmpresa_Id = Romaneios.EndEmpresa_Id RIGHT OUTER JOIN" & vbCrLf & _
                  "        SubOperacoes RIGHT OUTER JOIN" & vbCrLf & _
                  "        Clientes AS Empresa RIGHT OUTER JOIN" & vbCrLf & _
                  "        Pesagem ON Empresa.Cliente_Id = Pesagem.Empresa_Id AND Empresa.Endereco_Id = Pesagem.EndEmpresa_Id LEFT OUTER JOIN" & vbCrLf & _
                  "        Produtos ON Pesagem.Produto = Produtos.Produto_Id ON SubOperacoes.Operacao_Id = Pesagem.Operacao AND" & vbCrLf & _
                  "        SubOperacoes.SubOperacoes_Id = Pesagem.SubOperacao ON RomaneiosXPesagens.Empresa_Id = Pesagem.Empresa_Id AND" & vbCrLf & _
                  "        RomaneiosXPesagens.EndEmpresa_Id = Pesagem.EndEmpresa_Id AND RomaneiosXPesagens.Pesagem_Id = Pesagem.Pesagem_Id AND" & vbCrLf & _
                  "        RomaneiosXPesagens.Sequencia_Id = Pesagem.Sequencia_Id ON Clientes.Cliente_Id = Pesagem.Cliente AND Clientes.Endereco_Id = Pesagem.EndCliente ON" & vbCrLf & _
                  "        nfxr.Empresa_Id = Romaneios.Empresa_Id AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id AND" & vbCrLf & _
                  "        nfxr.Romaneio_Id = Romaneios.Romaneio_Id LEFT OUTER JOIN" & vbCrLf & _
                  "        NotasFiscais RIGHT OUTER JOIN" & vbCrLf & _
                  "        NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf & _
                  "        NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
                  "        NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
                  "        NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id ON nfxr.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND" & vbCrLf & _
                  "        nfxr.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND nfxr.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND" & vbCrLf & _
                  "        nfxr.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
                  "        nfxr.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND nfxr.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
                  "        nfxr.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf

            sql = sql & " WHERE  Pesagem.Empresa_ID = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "' "

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue.Split("-")(0)) Then
                sql &= " And Pesagem.Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlDeposito.SelectedValue.Split("-")(0)) Then
                sql &= " And NotasFiscais.Deposito = '" & ddlDeposito.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
                sql &= " And Left(Pesagem.Produto, 5) = '" & Left(ddlGrupo.SelectedValue, 5) & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
                sql &= " And Pesagem.Produto = '" & ddlProduto.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlOperacao.SelectedValue) Then
                sql &= " And NotasFiscais.Operacao = " & ddlOperacao.SelectedValue & vbCrLf
            End If

            sql &= " And (Produtos.Agrupar = 'N') And (Pesagem.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') AND (Pesagem.Situacao = 1)" & vbCrLf & _
                   " AND (Pesagem.Sequencia_Id = 0) AND (Pesagem.SegundaPesagem <> 0) And  ((ISNULL(NotasFiscaisXItens.QuantidadeFisica, 0) = 0) OR (MONTH(Pesagem.Movimento) <> MONTH(NotasFiscais.Movimento)))" & vbCrLf & _
                   " GROUP BY Empresa.Reduzido, Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Empresa.Nome, Empresa.Cidade, Empresa.Estado, Pesagem.Cliente," & vbCrLf & _
                   "        Pesagem.EndCliente, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Pesagem.Operacao, Pesagem.SubOperacao, SubOperacoes.Descricao, Produtos.Nome," & vbCrLf & _
                   "        NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, Pesagem.Movimento, Pesagem.Movimento, Pesagem.Pesagem_Id, Pesagem.Liquido," & vbCrLf & _
                   "        Pesagem.UsuarioInclusao , Pesagem.Placa, NotasFiscais.Movimento, Romaneios.Movimento " & vbCrLf & _
                   " ORDER BY NomeCliente" & vbCrLf
        End If

        Return Banco.ConsultaDataSet(sql, "NotasFiscais")
    End Function

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe o período.")
            Return False
        End If

        Return True
    End Function

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

#End Region

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeNotasFiscais", "ACESSAR") Then
                    CarregaUnidade()
                    CarregaClientes()
                    CarregaDeposito()
                    CarregaGrupoDeProduto()
                    CarregaOperacoes()

                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gerencial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupo.SelectedIndexChanged
        Try
            CarregaProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlOperacao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlOperacao.SelectedIndexChanged
        Try
            BuscarSubOperacoes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidaCampos() Then
                Dim ds As DataSet = getDataSet()

                If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum resultado encontrado.")
                Else
                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("MovimentoDaNota", eTipoCampo.Data)
                    colunas.Add("MovimentoDoRomaneio", eTipoCampo.Data)
                    colunas.Add("DataDaNota", eTipoCampo.Data)
                    colunas.Add("DifBalanca", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("Desconto", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("QuantidadeFisica", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("Unitario", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("ValorDoProduto", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("ICMS", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("Funrural", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("FETHAB", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("PIS", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("PISCRE", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("COFINS", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("COFINSCRE", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("IRRF", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("VALORLIQUIDO", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("PercentualComissao", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("ValorDaComissao", eTipoCampo.ValorSemTotalizador)
                    colunas.Add("Indice", eTipoCampo.ValorSemTotalizador)

                    BindExcelOffice(Me.Page, ds, "Notas Fiscais", colunas)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudar_Click(sender As Object, e As EventArgs) Handles lnkAjudar.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeNotasFiscais")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BindExcelOffice(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing)
        If ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Throw New Exception("Nenhum resultado encontrado!")
        Else
            Try
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then File.Delete(fileName)

                'emitir excel.xsls do office / relatório padrão em lista
                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando aba da planilha
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)

                        'inserindo o cabeçalho
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        Dim complemento As String() = {"Chegada", "BrutoDeChegada", "Retencao", "LiquidoDeChegada"}
                        For i = 0 To 3
                            worksheet.Cells(rowIndex, columnIndex).Value = complemento(i)
                            columnIndex += 1
                        Next

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count + 4)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using

                        rowIndex += 1

                        'exportando conteúdo da planilha com os dados da tabela
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                Else
                                    worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                End If

                                'formatações de valores
                                If colunas IsNot Nothing Then
                                    For Each coluna In colunas
                                        If coluna.Key = col.ColumnName Then
                                            If coluna.Value = eTipoCampo.Numerico Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                Exit For
                                            ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                Exit For
                                            ElseIf coluna.Value = eTipoCampo.Data Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                Exit For
                                            End If
                                        End If
                                    Next
                                End If
                                columnIndex += 1
                            Next

                            'formatações de celulas
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count + 4)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        If Not rdoLaudoSemNota.Checked Then
                            rowIndex = 1

                            For Each row As DataRow In ds.Tables(0).Rows

                                If Not String.IsNullOrWhiteSpace(row("RedEmpresa")) AndAlso row("E_S") = "S" AndAlso row("Operacao") = 66 Then

                                    Dim sql As String = "SELECT NxD.Movimento," & vbCrLf & _
                                                        " NxD.PesoBruto," & vbCrLf & _
                                                        " NxD.Desconto," & vbCrLf & _
                                                        " NxD.PesoLiquido" & vbCrLf & _
                                                        " FROM NotasXDestinos NxD" & vbCrLf & _
                                                        "" & vbCrLf & _
                                                        " Where  NxD.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'  And " & vbCrLf & _
                                                        " NxD.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & " And" & vbCrLf & _
                                                        " NxD.Cliente_Id = '" & row("Cliente") & " ' And " & vbCrLf & _
                                                        " NxD.EndCliente_Id = " & row("EndCliente") & " And" & vbCrLf & _
                                                        " NxD.EntradaSaida_Id = 'S' And" & vbCrLf & _
                                                        " NxD.Serie_Id = '" & row("SerieDaNota") & "' And" & vbCrLf & _
                                                        " NxD.Nota_Id = " & row("NumeroDaNota") & "" & vbCrLf

                                    Dim dsNF As DataSet = Banco.ConsultaDataSet(sql, "Nota")

                                    For Each dr As DataRow In dsNF.Tables(0).Rows
                                        columnIndex = ds.Tables(0).Columns.Count + 1
                                        For Each coldr As DataColumn In dsNF.Tables(0).Columns
                                            worksheet.Cells(rowIndex, columnIndex).Value = dr(coldr.ColumnName)
                                            If IsDate(dr(coldr.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                            ElseIf IsNumeric(dr(coldr.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                            End If
                                            columnIndex += 1
                                        Next
                                    Next

                                    sql = " SELECT top 1 Romaneios.Movimento, Romaneios.PesoBruto, Romaneios.Desconto, Romaneios.PesoLiquido" & vbCrLf & _
                                          " FROM  NotasFiscais INNER JOIN" & vbCrLf & _
                                          "    NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf & _
                                          "    NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
                                          "    NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
                                          "    NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                                          "    NotasFiscaisXRomaneios nfxr ON NotasFiscaisXItens.Empresa_Id = nfxr.Empresa_Id AND" & vbCrLf & _
                                          "    NotasFiscaisXItens.EndEmpresa_Id = nfxr.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = nfxr.Cliente_Id AND" & vbCrLf & _
                                          "    NotasFiscaisXItens.EndCliente_Id = nfxr.EndCliente_Id AND" & vbCrLf & _
                                          "    NotasFiscaisXItens.EntradaSaida_Id = nfxr.EntradaSaida_Id AND NotasFiscaisXItens.Serie_Id = nfxr.Serie_Id AND" & vbCrLf & _
                                          "    NotasFiscaisXItens.Nota_Id = nfxr.Nota_Id INNER JOIN" & vbCrLf & _
                                          "    Romaneios ON nfxr.Empresa_Id = Romaneios.Empresa_Id AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id AND" & vbCrLf & _
                                          "    nfxr.Romaneio_Id = Romaneios.Romaneio_Id" & vbCrLf & _
                                          " where NotasFiscais.Empresa_Id = '" & row("DepositoDestino") & "' And" & vbCrLf & _
                                          " NotasFiscais.EndEmpresa_Id = " & row("EndDepositoDestino") & " And" & vbCrLf & _
                                          " NotasFiscais.Cliente_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'  And " & vbCrLf & _
                                          " NotasFiscais.EndCliente_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & " And" & vbCrLf & _
                                          " NotasFiscais.EntradaSaida_Id = 'E' AND" & vbCrLf & _
                                          " NotasFiscais.Nota_Id = " & row("NumeroDaNota") & vbCrLf

                                    dsNF = New DataSet()
                                    dsNF = Banco.ConsultaDataSet(sql, "Nota")

                                    For Each dr As DataRow In dsNF.Tables(0).Rows
                                        columnIndex = ds.Tables(0).Columns.Count + 1
                                        For Each coldr As DataColumn In dsNF.Tables(0).Columns
                                            worksheet.Cells(rowIndex, columnIndex).Value = dr(coldr.ColumnName)
                                            If IsDate(dr(coldr.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                            ElseIf IsNumeric(dr(coldr.ColumnName)) Then
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                            End If
                                            columnIndex += 1
                                        Next
                                    Next
                                End If

                                rowIndex += 1
                            Next
                        End If

                        'soma dos campos de valores
                        If colunas IsNot Nothing Then
                            'aplicando formatação nas células do rodapé
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            columnIndex = 1

                            For Each col In colunas
                                If col.Value = eTipoCampo.ValorComTotalizador Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                    Exit For
                                End If
                            Next

                            For Each col As DataColumn In ds.Tables(0).Columns
                                For Each coluna In colunas
                                    If coluna.Key = col.ColumnName Then
                                        If coluna.Value = eTipoCampo.ValorComTotalizador Then
                                            worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                            Exit For
                                        End If
                                    End If
                                Next
                                columnIndex += 1
                            Next
                        End If

                        rowIndex += 1

                        'criando auto filtro na planilha
                        worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count).AutoFilter = True

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linham
                        worksheet.View.FreezePanes(2, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Catch ex As Exception
                Throw New Exception(ex.Message)
            End Try
        End If
    End Sub

#End Region


End Class