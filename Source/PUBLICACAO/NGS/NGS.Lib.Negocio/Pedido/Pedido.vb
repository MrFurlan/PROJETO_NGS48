Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data.SqlClient
Imports System.Configuration
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine
Imports System.Web.UI
Imports System.Linq.Expressions

<Serializable()>
Public Class Pedidos
    Inherits List(Of Pedido)

    Public Sub New()

    End Sub

    Public Sub New(ByVal Limite As Integer, Optional ByVal CodigoPedido As Integer = 0,
                               Optional ByVal PedidoEfetivo As String = "", Optional ByVal UnidadeNegocio As String = "",
                               Optional ByVal Empresa As String = "", Optional ByVal EnderecoEmpresa As Integer = 0,
                               Optional ByVal Cliente As String = "", Optional ByVal EnderecoCliente As Integer = 0,
                               Optional ByVal Situacao As eSituacao = Nothing, Optional ByVal Safra As String = "",
                               Optional ByVal Operacao As Integer = 0, Optional ByVal SubOperacao As Integer = 0,
                               Optional ByVal EntradaSaida As eEntradaSaidaNenhum = eEntradaSaidaNenhum.Nenhum,
                               Optional ByVal Agrupar As eTriBool = eTriBool.Nenhum,
                               Optional ByVal DataPedidoInicial As DateTime = Nothing,
                               Optional ByVal DataPedidoFinal As DateTime = Nothing,
                               Optional ByVal OperadorDataPedido As eOperadoresSelecao = Nothing,
                               Optional ByVal CodigoProduto As String = "",
                               Optional ByVal OrdenarPor As String = "",
                               Optional ByVal Classe As String = "",
                               Optional ByVal Pedidos As String = "")

        Selecionar(Limite, CodigoPedido, PedidoEfetivo, UnidadeNegocio, Empresa, EnderecoEmpresa, Cliente)
    End Sub

    Private _PedidoSelecionado As Pedido

    Public Property PedidoSelecionado() As Pedido
        Get
            Return _PedidoSelecionado
        End Get
        Set(ByVal value As Pedido)
            _PedidoSelecionado = value
        End Set
    End Property

    Public Function Selecionar(Optional ByVal Limite As Integer = 0, Optional ByVal CodigoPedido As Integer = 0,
                               Optional ByVal PedidoEfetivo As String = "", Optional ByVal UnidadeNegocio As String = "",
                               Optional ByVal Empresa As String = "", Optional ByVal EnderecoEmpresa As Integer = 0,
                               Optional ByVal Cliente As String = "", Optional ByVal EnderecoCliente As Integer = 0,
                               Optional ByVal Situacao As eSituacao = Nothing, Optional ByVal Safra As String = "",
                               Optional ByVal Operacao As Integer = 0, Optional ByVal SubOperacao As Integer = 0,
                               Optional ByVal EntradaSaida As eEntradaSaidaNenhum = eEntradaSaidaNenhum.Nenhum,
                               Optional ByVal Agrupar As eTriBool = eTriBool.Nenhum,
                               Optional ByVal DataPedidoInicial As DateTime = Nothing,
                               Optional ByVal DataPedidoFinal As DateTime = Nothing,
                               Optional ByVal OperadorDataPedido As eOperadoresSelecao = Nothing,
                               Optional ByVal CodigoProduto As String = "",
                               Optional ByVal OrdenarPor As String = "",
                               Optional ByVal Classe As String = "",
                               Optional ByVal Pedidos As String = "") As Boolean


        Dim objBanco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT" & IIf(Limite > 0, " TOP " & Limite.ToString(), "") & " P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id, P.UnidadeDeNegocio, P.Cliente, P.EndCliente, c.Nome as NomeCliente," & vbCrLf &
                  "       P.Praca, P.EndPraca, ISNULL(P.PedidoEfetivo,'') PedidoEfetivo, P.Safra, P.Moeda, isnull(TemVariacao,0) as TemVariacao, P.Indexador, P.Operacao, P.SubOperacao, P.Situacao, P.DataPedido, " & vbCrLf &
                  "       P.DataEntrega, isnull(P.DataInicioEntrega,P.DataPedido) as DataInicioEntrega, P.PedidoOrigem, P.FreteCIFFOB, " & vbCrLf &
                  "       case" & vbCrLf &
                  "          when isnull(p.origemdestino,'') = ''                       then 0" & vbCrLf &
                  "          When p.origemDestino in ('0','1')                          then P.OrigemDestino" & vbCrLf &
                  "          when so.entradasaida = 'E' and P.origemDestino = 'DESTINO' then 0" & vbCrLf &
                  "          when so.entradasaida = 'E' and P.origemDestino = 'ORIGEM'  then 1" & vbCrLf &
                  "          when so.entradasaida = 'S' and P.origemDestino = 'ORIGEM'  then 0" & vbCrLf &
                  "          when so.entradasaida = 'S' and P.origemDestino = 'DESTINO' then 1" & vbCrLf &
                  "       end as OrigemDestino," & vbCrLf &
                  "       P.Solicitacao, P.UsuarioInclusao, P.UsuarioInclusaoData, isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                  "       P.UsuarioAlteracaoData, isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento, P.UsuarioCancelamentoData, isnull(P.UsuarioLiberacao,'') AS UsuarioLiberacao, " & vbCrLf &
                  "       P.UsuarioLiberacaoData, P.Observacoes, isnull(P.CondicaoPagamento,0) as CondicaoPagamento,  P.BancoCliente, P.AgenciaCliente, " & vbCrLf &
                  "       P.DigitoAgenciaCliente, P.ContaCliente, P.DigitoContaCliente," & vbCrLf &
                  "       isnull(case" & vbCrLf &
                  "                when p.comercializacao = 'Industria' then 0" & vbCrLf &
                  "                when p.comercializacao = 'Comercio'  then 1" & vbCrLf &
                  "                else p.comercializacao " & vbCrLf &
                  "              end,1) as Comercializacao," & vbCrLf &
                  "       isnull(P.Finalidade,1) AS Finalidade, isnull(P.Contrato,'') as Contrato, isnull(p.Carteira,'') as Carteira, " & vbCrLf &
                  "       isnull(P.Taxa,0) as Taxa, isnull(P.VencimentoPedido,P.DataEntrega) as VencimentoPedido, " & vbCrLf &
                  "       isnull(p.LocalEmbarque,'') as LocalEmbarque, isnull(p.EndLocalEmbarque,0) as EndLocalEmbarque," & vbCrLf &
                  "       isnull(p.MomentoFinanceiro,0) as MomentoFinanceiro, isnull(p.AgruparFinanceiro,0) as AgruparFinanceiro, " & vbCrLf &
                  "       isnull(p.IndiceFixado, 0) AS IndiceFixado, isnull(p.EstadoEntrega,'') AS EstadoEntrega, isnull(p.CidadeEntrega,'') AS CidadeEntrega, " & vbCrLf &
                  "       isnull(p.FiscalAberto,1) as FiscalAberto, isnull(p.FinanceiroAberto,1) as FinanceiroAberto, " & vbCrLf &
                  "       isnull(p.EmpresaTroca, '') AS EmpresaTroca, isnull(p.EndEmpresaTroca, 0) AS EndEmpresaTroca, " & vbCrLf &
                  "       isnull(p.PedidoTroca, 0) AS PedidoTroca, isnull(p.ContaAdiantamentoTroca, '') AS ContaAdiantamentoTroca, " & vbCrLf &
                  "       isnull(p.CondicaoPagamentoEntrega, 0) AS CondicaoPagamentoEntrega, isnull(p.QuotaEntrega, 0) AS QuotaEntrega, " & vbCrLf &
                  "       isnull(p.PeriodicidadeEntrega, 0) AS PeriodicidadeEntrega, isnull(P.PedidoBloqueado,0) AS PedidoBloqueado, " & vbCrLf &
                  "       isnull(P.LiberaCarregamento,0) AS LiberaCarregamento, isnull(p.VersaoPedido,0) as VersaoPedido," & vbCrLf &
                  "       isnull(p.VersaoUsuario,isnull(p.UsuarioAlteracao,p.UsuarioInclusao)) as VersaoUsuario," & vbCrLf &
                  "       isnull(p.VersaoHorarioBloqueio,CAST(GETDATE() AS DATETIME)) as VersaoHorarioBloqueio, isnull(p.IndexadorFixo,1) AS IndexadorFixo, " & vbCrLf &
                  "       cXe.DataFinanceiro, isnull(p.troca,0) as Troca, isnull(p.Antecipada,0) as Antecipada, isnull(p.Recompra,0) as Recompra, " & vbCrLf &
                  "       isnull(p.XPedNFe,'') as XPedNFe, isnull(p.ItemXPedNFe,'') as ItemXPedNFe, isnull(P.InvoiceNavio,0) as InvoiceNavio, " & vbCrLf &
                  "       isnull(ObservacoesControleInterno,'') as ObservacoesControleInterno" & vbCrLf &
                  "  FROM Pedidos P " & vbCrLf &
                  " INNER JOIN ClientesXEmpresas cXe " & vbCrLf &
                  "    ON P.Empresa_Id    = cXe.Empresa_Id " & vbCrLf &
                  "   AND P.EndEmpresa_Id = cXe.EndEmpresa_Id " & vbCrLf &
                  " Inner Join Clientes c" & vbCrLf &
                  "    ON c.Cliente_Id = p.Cliente" & vbCrLf &
                  "   And c.Endereco_Id = p.EndCliente" & vbCrLf &
                  " Inner Join Suboperacoes so" & vbCrLf &
                  "    on p.operacao    = so.Operacao_id" & vbCrLf &
                  "   and p.suboperacao = so.suboperacoes_id" & vbCrLf &
                  " WHERE P.UnidadeDeNegocio IS NOT NULL                                                                              " & vbCrLf &
                  "   AND P.PedidoOrigem = 0                                                                                          " & vbCrLf

            If Not String.IsNullOrWhiteSpace(Pedidos) Then
                Sql &= "AND P.Pedido_Id in (" & Pedidos & ")"
            End If

            'Sql &= "AND P.Pedido_Id in (45,19220980,19228301,19229318,19231409,19231521,19231522,19231523,19231524,19231525,19231526,19231527,19231528,19231529,19231530,19231531,19231532,19231533,19231534,19231535,19231536,19231537,19231538,19231539,19231540,19231541,19231542,19231543,19231544,19231545)"

            If Classe.Length > 0 Then
                Sql &= " AND SO.Classe ='" & Classe & "'" & vbCrLf
            End If

            If UnidadeNegocio <> "" Then Sql &= "AND P.UnidadeDeNegocio = '" & UnidadeNegocio & "' " & vbCrLf
            If CodigoPedido > 0 Then Sql &= "AND P.Pedido_Id = " & CodigoPedido.ToString() & vbCrLf
            If PedidoEfetivo <> "" Then Sql &= "AND P.PedidoEfetivo = '" & PedidoEfetivo & "' " & vbCrLf

            If Empresa <> "" Then
                Sql &= "AND P.Empresa_Id = '" & Empresa & "' " & vbCrLf &
                          "AND P.EndEmpresa_Id = " & EnderecoEmpresa.ToString() & vbCrLf
            End If

            If Cliente <> "" Then
                Sql &= "AND P.Cliente = '" & Cliente & "' " & vbCrLf &
                          "AND P.EndCliente = " & EnderecoCliente.ToString() & vbCrLf
            End If

            If Situacao <> Nothing Then Sql &= "AND P.Situacao = " & Convert.ToInt32(Situacao).ToString() & vbCrLf
            If Safra <> "" Then Sql &= "AND P.Safra = '" & Safra & "' " & vbCrLf

            If Operacao > 0 Then
                Sql &= "AND P.Operacao = " & Operacao.ToString() & vbCrLf
                If SubOperacao > 0 Then Sql &= "AND P.SubOperacao = " & SubOperacao.ToString() & vbCrLf
            End If

            If OperadorDataPedido <> Nothing Then
                Sql &= "AND P.DataPedido " & Conversoes.ConverterOperadoresBancoParaString(OperadorDataPedido) & " " &
                          "'" & DataPedidoInicial.ToString("yyyy-MM-dd") & "' " & vbCrLf

                If OperadorDataPedido = eOperadoresSelecao.Entre Then Sql &= "AND '" & DataPedidoFinal.ToString("yyyy-MM-dd") & "' " & vbCrLf
            End If

            If EntradaSaida <> eEntradaSaidaNenhum.Nenhum Then
                Sql &= "AND SO.EntradaSaida = '" & EntradaSaida.ToString().Substring(0, 1) & "' " & vbCrLf
            End If

            Dim strSQLProduto As String = "AND EXISTS (SELECT NULL " & vbCrLf &
                                                      "FROM PedidoXItem PIT " & vbCrLf &
                                                      "[INNER] " & vbCrLf &
                                                      "WHERE P.Empresa_Id = PIT.Empresa_Id " & vbCrLf &
                                                      "AND P.EndEmpresa_Id = PIT.EndEmpresa_Id " & vbCrLf &
                                                      "AND P.Pedido_Id = PIT.Pedido_Id " & vbCrLf &
                                                      "[WHERE]) " & vbCrLf

            Dim strSQLProdFinal As String = ""

            If Agrupar <> eTriBool.Nenhum Then
                strSQLProdFinal = strSQLProduto.Replace("[INNER]", "INNER JOIN Produtos PR " &
                                                                   "ON PR.Produto_Id = PIT.Produto_Id")

                strSQLProduto = strSQLProduto.Replace("[WHERE]", "AND PR.Agrupar = '" & IIf(Agrupar = eTriBool.Verdadeiro, "S", "N") & "'" &
                                                                 IIf(CodigoProduto <> "", "AND PIT.Produto_Id = '" & CodigoProduto & "'", ""))
            ElseIf CodigoProduto <> "" Then
                strSQLProdFinal = strSQLProduto.Replace("[INNER]", "").Replace("[WHERE]", "AND PIT.Produto_Id = '" & CodigoProduto & "'")
            End If

            Sql &= strSQLProdFinal

            If OrdenarPor <> "" Then Sql &= "ORDER BY " & OrdenarPor

            Dim dsPedidos As DataSet = objBanco.ConsultaDataSet(Sql, "Pedidos")

            For Each row As DataRow In dsPedidos.Tables(0).Rows
                Dim objPedido As New Pedido()

                objPedido.CodigoEmpresa = row("Empresa_Id")
                objPedido.EnderecoEmpresa = row("EndEmpresa_Id")
                objPedido.Codigo = row("Pedido_Id")
                objPedido.CodigoUnidadeNegocio = row("UnidadeDeNegocio")
                objPedido.CodigoCliente = row("Cliente")
                objPedido.NomeCliente = Funcoes.FormatarCpfCnpj(row("Cliente")) & " - " & row("NomeCliente")
                objPedido.EnderecoCliente = row("EndCliente")
                objPedido.CodigoPraca = row("Praca")
                objPedido.EnderecoPraca = row("EndPraca")
                objPedido.PedidoEfetivo = row("PedidoEfetivo")
                objPedido.CodigoSafra = row("Safra")
                objPedido.CodigoMoeda = row("Moeda")
                objPedido.TemVariacao = row("TemVariacao")
                objPedido.CodigoIndexador = row("Indexador")
                objPedido.CodigoOperacao = row("Operacao")
                objPedido.CodigoSubOperacao = row("SubOperacao")
                objPedido.CodigoSituacao = row("Situacao")
                objPedido.DataPedido = row("DataPedido")
                objPedido.DataEntregaInicial = row("DataInicioEntrega")
                objPedido.DataEntregaFinal = row("DataEntrega")
                objPedido.FreteCIFFOB = [Enum].Parse(GetType(eTiposFrete), row("FreteCIFFOB").ToString())
                objPedido.OrigemDestino = row("OrigemDestino")
                objPedido.Solicitacao = row("Solicitacao")
                objPedido.UsuarioInclusao = row("UsuarioInclusao")
                objPedido.DataInclusao = row("UsuarioInclusaoData")
                objPedido.UsuarioAlteracao = row("UsuarioAlteracao")
                If Not IsDBNull(row("UsuarioAlteracaoData")) Then
                    objPedido.DataAlteracao = Convert.ToDateTime(row("UsuarioAlteracaoData"))
                End If

                objPedido.UsuarioCancelamento = row("UsuarioCancelamento").ToString()
                If Not IsDBNull(row("UsuarioCancelamentoData")) Then
                    objPedido.DataCancelamento = Convert.ToDateTime(row("UsuarioCancelamentoData"))
                End If

                objPedido.UsuarioLiberacao = row("UsuarioLiberacao").ToString()
                If Not IsDBNull(row("UsuarioLiberacaoData")) Then
                    objPedido.DataLiberacao = Convert.ToDateTime(row("UsuarioLiberacaoData"))
                End If

                objPedido.Observacoes = IIf(IsDBNull(row("Observacoes")), String.Empty, row("Observacoes").ToString)

                objPedido.CodigoCondicaoPagamento = row("CondicaoPagamento")

                objPedido.ContaBancariaSelecionada = Nothing

                If Not IsDBNull(row("BancoCliente")) Then
                    If row("BancoCliente") > 0 Then
                        Dim bancoCliente = row("BancoCliente")
                        Dim agenciaCliente = row("AgenciaCliente")
                        Dim digitoAgencia = row("DigitoAgenciaCliente")
                        Dim contaCliente = row("ContaCliente")
                        Dim digitoContaCliente = row("DigitoContaCliente")
                        Dim intIndice As Integer = objPedido.Cliente.ContasBancarias.FindIndex(Function(s) s.CodigoBanco = bancoCliente And s.CodigoAgencia = agenciaCliente And s.DigitoAgencia = digitoAgencia And s.ContaCorrente = contaCliente And s.DigitoConta = digitoContaCliente)
                        If intIndice > -1 Then objPedido.ContaBancariaSelecionada = objPedido.Cliente.ContasBancarias(intIndice)
                    End If
                End If

                objPedido.Comercializacao = row("Comercializacao")

                objPedido.CodigoFinalidade = row("Finalidade")
                objPedido.Contrato = row("Contrato")
                objPedido.CodigoCarteira = row("Carteira")
                objPedido.Taxa = row("Taxa")
                objPedido.DataVencimentoPedido = row("VencimentoPedido")
                objPedido.CodigoLocalEmbarque = row("LocalEmbarque")
                objPedido.EndLocalEmbarque = row("EndLocalEmbarque")
                objPedido.MomentoFinanceiro = row("MomentoFinanceiro")
                objPedido.AgruparFinanceiro = row("AgruparFinanceiro")
                objPedido.IndiceFixado = row("IndiceFixado")
                objPedido.EstadoEntrega = row("EstadoEntrega")
                objPedido.CidadeEntrega = row("CidadeEntrega")
                objPedido.FiscalAberto = row("FiscalAberto")
                objPedido.FinanceiroAberto = row("FinanceiroAberto")

                objPedido.CodigoEmpresaTroca = row("EmpresaTroca")
                objPedido.EnderecoEmpresaTroca = row("EndEmpresaTroca")
                objPedido.CodigoPedidoTroca = row("PedidoTroca")
                objPedido.ContaAdiantamentoTroca = row("ContaAdiantamentoTroca")
                objPedido.CodigoCondicaoPagamentoDaEntrega = row("CondicaoPagamentoEntrega")

                objPedido.QuotaEntrega = row("QuotaEntrega")
                objPedido.PeriodicidadeEntrega = row("PeriodicidadeEntrega")
                objPedido.PedidoBloqueado = row("PedidoBloqueado")
                objPedido.LiberaCarregamento = row("LiberaCarregamento")

                objPedido.VersaoPedido = row("VersaoPedido")
                objPedido.VersaoUsuario = row("VersaoUsuario")
                objPedido.VersaoHorarioBloqueio = row("VersaoHorarioBloqueio")

                objPedido.IndexadorFixo = row("IndexadorFixo")

                objPedido.Troca = row("Troca")
                objPedido.Antecipada = row("Antecipada")
                objPedido.Recompra = row("Recompra")

                objPedido.XPedNFe = row("XPedNFe")
                objPedido.ItemXPedNFe = row("ItemXPedNFe")

                objPedido.InvoiceNavio = row("InvoiceNavio")

                objPedido.ObservacoesControleInterno = row("ObservacoesControleInterno")

                'Verifica se o pedido foi aberto com financeiro novo
                If Not IsDBNull(row("DataFinanceiro")) Then
                    objPedido.FinanceiroNovo = objPedido.DataPedido >= Convert.ToDateTime(row("DataFinanceiro"))
                    If objPedido.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                        objPedido.Peridiocidade = [Enum].Parse(GetType(ePeriodicidade), row("PeriodicidadeEntrega"))
                    End If
                End If

                Me.Add(objPedido)
            Next

            Return True
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function
End Class


<Serializable()>
Public Class Pedido
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal Endereco As Integer, ByVal Codigo As Integer, ByVal notaDeTerceiro As Boolean)
        Dim db As New AcessaBanco()

        Try
            Dim strSQL

            strSQL = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id, P.UnidadeDeNegocio, P.Cliente, P.EndCliente, " & vbCrLf &
                     "       P.Praca, P.EndPraca, P.PedidoEfetivo, P.Safra, P.Moeda, isnull(TemVariacao,0) as TemVariacao, P.Indexador, P.Operacao, P.SubOperacao, P.Situacao, P.DataPedido, " & vbCrLf &
                     "       P.DataEntrega, isnull(P.DataInicioEntrega,P.DataPedido) as DataInicioEntrega, P.PedidoOrigem, isnull(P.FreteCIFFOB,'NEN') as FreteCIFFOB, " & vbCrLf &
                     "       case" & vbCrLf &
                     "          when isnull(P.origemdestino,'') = ''                       then 0" & vbCrLf &
                     "          When P.origemDestino in ('0','1')                          then P.OrigemDestino" & vbCrLf &
                     "          when so.entradasaida = 'E' and P.origemDestino = 'DESTINO' then 0" & vbCrLf &
                     "          when so.entradasaida = 'E' and P.origemDestino = 'ORIGEM'  then 1" & vbCrLf &
                     "          when so.entradasaida = 'S' and P.origemDestino = 'ORIGEM'  then 0" & vbCrLf &
                     "          when so.entradasaida = 'S' and P.origemDestino = 'DESTINO' then 1" & vbCrLf &
                     "       end as OrigemDestino," & vbCrLf &
                     "       isnull(P.Finalidade,1) AS Finalidade, P.Solicitacao, P.UsuarioInclusao, P.UsuarioInclusaoData, " & vbCrLf &
                     "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                     "       isnull(P.UsuarioAlteracaoData,CAST(GETDATE() AS DATETIME)) AS UsuarioAlteracaoData, isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento, " & vbCrLf &
                     "       isnull(P.UsuarioCancelamentoData,CAST(GETDATE() AS DATETIME)) AS UsuarioCancelamentoData, isnull(P.UsuarioLiberacao,'') AS UsuarioLiberacao, " & vbCrLf &
                     "       isnull(P.UsuarioLiberacaoData,CAST(GETDATE() AS DATETIME)) AS UsuarioLiberacaoData, P.Observacoes, isnull(P.CondicaoPagamento,0) as CondicaoPagamento,  " & vbCrLf &
                     "       P.BancoCliente, P.AgenciaCliente, P.DigitoAgenciaCliente, " & vbCrLf &
                     "       P.ContaCliente, P.DigitoContaCliente," & vbCrLf &
                     "       isnull(case" & vbCrLf &
                     "                when P.comercializacao = 'Industria' then 0" & vbCrLf &
                     "                when P.comercializacao = 'Comercio'  then 1" & vbCrLf &
                     "                else P.comercializacao " & vbCrLf &
                     "              end,1) as Comercializacao," & vbCrLf &
                     "       isnull(P.Contrato,'') as Contrato, isnull(P.carteira,'') as Carteira," & vbCrLf &
                     "       isnull(P.Taxa,0) as Taxa, isnull(P.VencimentoPedido,P.DataEntrega) as VencimentoPedido, " & vbCrLf &
                     "       isnull(P.LocalEmbarque,'') as LocalEmbarque, isnull(P.EndLocalEmbarque,0) as EndLocalEmbarque," & vbCrLf &
                     "       isnull(P.MomentoFinanceiro,0) as MomentoFinanceiro, isnull(P.AgruparFinanceiro,0) as AgruparFinanceiro, " & vbCrLf &
                     "       isnull(P.IndiceFixado,0) AS IndiceFixado, isnull(P.EstadoEntrega,'') AS EstadoEntrega, isnull(P.CidadeEntrega,'') AS CidadeEntrega, " & vbCrLf &
                     "       isnull(P.FiscalAberto,1) as FiscalAberto, isnull(P.FinanceiroAberto,1) as FinanceiroAberto, " & vbCrLf &
                     "       isnull(P.EmpresaTroca, '') AS EmpresaTroca, isnull(P.EndEmpresaTroca, 0) AS EndEmpresaTroca, " & vbCrLf &
                     "       isnull(P.PedidoTroca, 0) AS PedidoTroca, isnull(P.ContaAdiantamentoTroca, '') AS ContaAdiantamentoTroca, " & vbCrLf &
                     "       isnull(P.CondicaoPagamento, 0) AS CondicaoPagamento, isnull(P.CondicaoPagamentoEntrega, 0) AS CondicaoPagamentoEntrega, isnull(P.QuotaEntrega, 0) AS QuotaEntrega, " & vbCrLf &
                     "       isnull(P.PeriodicidadeEntrega, 0) AS PeriodicidadeEntrega, isnull(p.PedidoBloqueado,0) AS PedidoBloqueado, " & vbCrLf &
                     "       isnull(p.LiberaCarregamento,0) AS LiberaCarregamento, isnull(P.VersaoPedido,0) as VersaoPedido," & vbCrLf &
                     "       isnull(P.VersaoUsuario,isnull(P.UsuarioAlteracao,P.UsuarioInclusao)) as VersaoUsuario," & vbCrLf &
                     "       isnull(P.VersaoHorarioBloqueio,CAST(GETDATE() AS DATETIME)) as VersaoHorarioBloqueio, isnull(P.IndexadorFixo,1) AS IndexadorFixo, " & vbCrLf &
                     "       cXe.DataFinanceiro, isnull(P.Troca,0) as Troca, isnull(P.Antecipada,0) as Antecipada, isnull(P.Recompra,0) as Recompra, " & vbCrLf &
                     "       isnull(p.XPedNFe,'') as XPedNFe, isnull(p.ItemXPedNFe,'') as ItemXPedNFe, isnull(P.InvoiceNavio,0) as InvoiceNavio, ISNULL(P.Embalagem, 0) AS Embalagem, " & vbCrLf &
                     "       ISNULL(EM.Descricao, '') AS DescricaoEmbalagem, ISNULL(P.TipoCondicaoEntrega, '') AS TipoCondicaoEntrega, ISNULL(P.TipoPagamentoPtax, '') AS TipoPagamentoPtax, " & vbCrLf &
                     "       ISNULL(P.LocalDeEmbarque, '') AS LocalDeEmbarque, isnull(ObservacoesControleInterno,'') as ObservacoesControleInterno" & vbCrLf &
                     "  FROM Pedidos P "

            If notaDeTerceiro Then

                strSQL &= " LEFT JOIN ClientesXEmpresas cXe " & vbCrLf &
                          "     ON P.Empresa_Id    = cXe.Empresa_Id" & vbCrLf &
                          "     AND P.EndEmpresa_Id = cXe.EndEmpresa_Id "

            Else

                strSQL &= " INNER JOIN ClientesXEmpresas cXe " & vbCrLf &
                          "     ON P.Empresa_Id    = cXe.Empresa_Id" & vbCrLf &
                          "     AND P.EndEmpresa_Id = cXe.EndEmpresa_Id "

            End If

            strSQL &= "     LEFT JOIN Embalagens EM " & vbCrLf &
                        "       ON P.Embalagem    = EM.Embalagem_Id" & vbCrLf &
                        "   Inner Join Suboperacoes so" & vbCrLf &
                        "       on p.operacao    = so.Operacao_id" & vbCrLf &
                        "       and p.suboperacao = so.suboperacoes_id" & vbCrLf &
                        "   WHERE P.Empresa_Id    ='" & Empresa & "'" & vbCrLf &
                        "       AND P.EndEmpresa_Id = " & Endereco & vbCrLf &
                        "       AND P.Pedido_Id     = " & Codigo

            Dim dsPedido As DataSet = db.ConsultaDataSet(strSQL, "Pedidos")

            If dsPedido.Tables(0).Rows.Count > 0 Then
                Dim drPedido As DataRow = dsPedido.Tables(0).Rows(0)

                Me.CodigoEmpresa = drPedido("Empresa_Id")
                Me.EnderecoEmpresa = drPedido("EndEmpresa_Id")
                Me.Codigo = drPedido("Pedido_Id")
                Me.CodigoUnidadeNegocio = drPedido("UnidadeDeNegocio")
                Me.CodigoCliente = drPedido("Cliente")
                Me.EnderecoCliente = drPedido("EndCliente")
                Me.CodigoPraca = drPedido("Praca")
                Me.EnderecoPraca = drPedido("EndPraca")
                Me.PedidoEfetivo = drPedido("PedidoEfetivo").ToString
                Me.CodigoSafra = drPedido("Safra")
                Me.CodigoMoeda = drPedido("Moeda")
                Me.TemVariacao = drPedido("TemVariacao")

                Me.CodigoIndexador = drPedido("Indexador")
                Me.CodigoOperacao = drPedido("Operacao")
                Me.CodigoSubOperacao = drPedido("SubOperacao")
                Me.CodigoSituacao = drPedido("Situacao")
                Me.DataPedido = drPedido("DataPedido")

                Me.DataEntregaInicial = drPedido("DataInicioEntrega")
                Me.DataEntregaFinal = drPedido("DataEntrega")

                Me.FreteCIFFOB = [Enum].Parse(GetType(eTiposFrete), drPedido("FreteCIFFOB").ToString())
                Me.OrigemDestino = drPedido("OrigemDestino")
                Me.CodigoFinalidade = drPedido("Finalidade")
                Me.Solicitacao = drPedido("Solicitacao")
                Me.UsuarioInclusao = drPedido("UsuarioInclusao").ToString()
                Me.DataInclusao = Convert.ToDateTime(drPedido("UsuarioInclusaoData"))
                Me.UsuarioAlteracao = drPedido("UsuarioAlteracao").ToString()
                Me.DataAlteracao = Convert.ToDateTime(drPedido("UsuarioAlteracaoData"))
                Me.UsuarioCancelamento = drPedido("UsuarioCancelamento").ToString()
                Me.DataCancelamento = Convert.ToDateTime(drPedido("UsuarioCancelamentoData"))
                Me.UsuarioLiberacao = drPedido("UsuarioLiberacao").ToString()
                Me.DataLiberacao = Convert.ToDateTime(drPedido("UsuarioLiberacaoData"))
                Me.Observacoes = drPedido("Observacoes")
                Me.CodigoCondicaoPagamento = drPedido("CondicaoPagamento")
                Me.Contrato = drPedido("Contrato").ToString()
                Me.CodigoCarteira = drPedido("Carteira")
                Me.Taxa = drPedido("Taxa")
                Me.DataVencimentoPedido = drPedido("VencimentoPedido")
                Me.CodigoLocalEmbarque = drPedido("LocalEmbarque")
                Me.EndLocalEmbarque = drPedido("EndLocalEmbarque")

                If Not IsDBNull(drPedido("BancoCliente")) Then
                    If drPedido("BancoCliente") > 0 Then


                        If Not Me.CodigoCliente = Me.CodigoPraca Then
                            Dim intIndice As Integer = Me.Praca.ContasBancarias.FindIndex(Function(s) s.CodigoBanco = drPedido("BancoCliente") And s.CodigoAgencia = drPedido("AgenciaCliente") And s.DigitoAgencia = drPedido("DigitoAgenciaCliente") And s.ContaCorrente = drPedido("ContaCliente") And s.DigitoConta = drPedido("DigitoContaCliente"))
                            If intIndice > -1 Then Me.ContaBancariaSelecionada = Me.Praca.ContasBancarias(intIndice)
                        Else
                            Dim intIndice As Integer = Me.Cliente.ContasBancarias.FindIndex(Function(s) s.CodigoBanco = drPedido("BancoCliente") And s.CodigoAgencia = drPedido("AgenciaCliente") And s.DigitoAgencia = drPedido("DigitoAgenciaCliente") And s.ContaCorrente = drPedido("ContaCliente") And s.DigitoConta = drPedido("DigitoContaCliente"))
                            If intIndice > -1 Then Me.ContaBancariaSelecionada = Me.Cliente.ContasBancarias(intIndice)
                        End If
                    End If
                End If

                Me.Itens = New Negocio.ListPedidoXItem(Me)


                Me.Comercializacao = drPedido("Comercializacao")

                Me.CodigoCondicaoPagamento = drPedido("CondicaoPagamento")
                Me.MomentoFinanceiro = drPedido("MomentoFinanceiro")
                Me.AgruparFinanceiro = drPedido("AgruparFinanceiro")
                Me.IndiceFixado = drPedido("IndiceFixado")
                Me.EstadoEntrega = drPedido("EstadoEntrega")
                Me.CidadeEntrega = drPedido("CidadeEntrega")
                Me.FiscalAberto = drPedido("FiscalAberto")
                Me.FinanceiroAberto = drPedido("FinanceiroAberto")

                Me.CodigoEmpresaTroca = drPedido("EmpresaTroca")
                Me.EnderecoEmpresaTroca = drPedido("EndEmpresaTroca")
                Me.CodigoPedidoTroca = drPedido("PedidoTroca")
                Me.ContaAdiantamentoTroca = drPedido("ContaAdiantamentoTroca")

                Me.CodigoCondicaoPagamentoDaEntrega = drPedido("CondicaoPagamentoEntrega")
                Me.QuotaEntrega = drPedido("QuotaEntrega")
                Me.PeriodicidadeEntrega = drPedido("PeriodicidadeEntrega")
                Me.PedidoBloqueado = drPedido("PedidoBloqueado")
                Me.LiberaCarregamento = drPedido("LiberaCarregamento")

                Me.VersaoPedido = drPedido("VersaoPedido")
                Me.VersaoUsuario = drPedido("VersaoUsuario")
                Me.VersaoHorarioBloqueio = drPedido("VersaoHorarioBloqueio")

                Me.IndexadorFixo = drPedido("IndexadorFixo")

                Me.Troca = drPedido("Troca")
                Me.Antecipada = drPedido("Antecipada")
                Me.Recompra = drPedido("Recompra")

                Me.XPedNFe = drPedido("XPedNFe")
                Me.ItemXPedNFe = drPedido("ItemXPedNFe")

                Me.InvoiceNavio = drPedido("InvoiceNavio")

                Me.Embalagem = drPedido("Embalagem")
                Me.DescricaoEmbalagem = drPedido("DescricaoEmbalagem")
                Me.TipoCondicaoEntrega = drPedido("TipoCondicaoEntrega")
                Me.TipoPagamentoPtax = drPedido("TipoPagamentoPtax")
                Me.LocalDeEmbarque = drPedido("LocalDeEmbarque")

                Me.ObservacoesControleInterno = drPedido("ObservacoesControleInterno")

                'Verifica se o pedido foi aberto com financeiro novo
                If Not IsDBNull(drPedido("DataFinanceiro")) Then
                    Me.FinanceiroNovo = Me.DataPedido >= Convert.ToDateTime(drPedido("DataFinanceiro"))
                    If Me.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                        Me.Peridiocidade = [Enum].Parse(GetType(ePeriodicidade), drPedido("PeriodicidadeEntrega"))
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub New(ByVal Empresa As String, ByVal Endereco As Integer, ByVal Codigo As Integer)
        Dim db As New AcessaBanco()

        Try
            Dim strSQL
            strSQL = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_Id, P.UnidadeDeNegocio, P.Cliente, P.EndCliente, " & vbCrLf &
                     "       P.Praca, P.EndPraca, P.PedidoEfetivo, P.Safra, P.Moeda, isnull(TemVariacao,0) as TemVariacao, P.Indexador, P.Operacao, P.SubOperacao, P.Situacao, P.DataPedido, " & vbCrLf &
                     "       P.DataEntrega, isnull(P.DataInicioEntrega,P.DataPedido) as DataInicioEntrega, P.PedidoOrigem, isnull(P.FreteCIFFOB,'NEN') as FreteCIFFOB, " & vbCrLf &
                     "       case" & vbCrLf &
                     "          when isnull(P.origemdestino,'') = ''                       then 0" & vbCrLf &
                     "          When P.origemDestino in ('0','1')                          then P.OrigemDestino" & vbCrLf &
                     "          when so.entradasaida = 'E' and P.origemDestino = 'DESTINO' then 0" & vbCrLf &
                     "          when so.entradasaida = 'E' and P.origemDestino = 'ORIGEM'  then 1" & vbCrLf &
                     "          when so.entradasaida = 'S' and P.origemDestino = 'ORIGEM'  then 0" & vbCrLf &
                     "          when so.entradasaida = 'S' and P.origemDestino = 'DESTINO' then 1" & vbCrLf &
                     "       end as OrigemDestino," & vbCrLf &
                     "       isnull(P.Finalidade,1) AS Finalidade, P.Solicitacao, P.UsuarioInclusao, P.UsuarioInclusaoData, " & vbCrLf &
                     "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao, " & vbCrLf &
                     "       isnull(P.UsuarioAlteracaoData,CAST(GETDATE() AS DATETIME)) AS UsuarioAlteracaoData, isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento, " & vbCrLf &
                     "       isnull(P.UsuarioCancelamentoData,CAST(GETDATE() AS DATETIME)) AS UsuarioCancelamentoData, isnull(P.UsuarioLiberacao,'') AS UsuarioLiberacao, " & vbCrLf &
                     "       isnull(P.UsuarioLiberacaoData,CAST(GETDATE() AS DATETIME)) AS UsuarioLiberacaoData, P.Observacoes, isnull(P.CondicaoPagamento,0) as CondicaoPagamento,  " & vbCrLf &
                     "       P.BancoCliente, P.AgenciaCliente, P.DigitoAgenciaCliente, " & vbCrLf &
                     "       P.ContaCliente, P.DigitoContaCliente," & vbCrLf &
                     "       isnull(case" & vbCrLf &
                     "                when P.comercializacao = 'Industria' then 0" & vbCrLf &
                     "                when P.comercializacao = 'Comercio'  then 1" & vbCrLf &
                     "                else P.comercializacao " & vbCrLf &
                     "              end,1) as Comercializacao," & vbCrLf &
                     "       isnull(P.Contrato,'') as Contrato, isnull(P.carteira,'') as Carteira," & vbCrLf &
                     "       isnull(P.Taxa,0) as Taxa, isnull(P.VencimentoPedido,P.DataEntrega) as VencimentoPedido, " & vbCrLf &
                     "       isnull(P.LocalEmbarque,'') as LocalEmbarque, isnull(P.EndLocalEmbarque,0) as EndLocalEmbarque," & vbCrLf &
                     "       isnull(P.MomentoFinanceiro,0) as MomentoFinanceiro, isnull(P.AgruparFinanceiro,0) as AgruparFinanceiro, " & vbCrLf &
                     "       isnull(P.IndiceFixado,0) AS IndiceFixado, isnull(P.EstadoEntrega,'') AS EstadoEntrega, isnull(P.CidadeEntrega,'') AS CidadeEntrega, " & vbCrLf &
                     "       isnull(P.FiscalAberto,1) as FiscalAberto, isnull(P.FinanceiroAberto,1) as FinanceiroAberto, " & vbCrLf &
                     "       isnull(P.EmpresaTroca, '') AS EmpresaTroca, isnull(P.EndEmpresaTroca, 0) AS EndEmpresaTroca, " & vbCrLf &
                     "       isnull(P.PedidoTroca, 0) AS PedidoTroca, isnull(P.ContaAdiantamentoTroca, '') AS ContaAdiantamentoTroca, " & vbCrLf &
                     "       isnull(P.CondicaoPagamento, 0) AS CondicaoPagamento, isnull(P.CondicaoPagamentoEntrega, 0) AS CondicaoPagamentoEntrega, isnull(P.QuotaEntrega, 0) AS QuotaEntrega, " & vbCrLf &
                     "       isnull(P.PeriodicidadeEntrega, 0) AS PeriodicidadeEntrega, isnull(p.PedidoBloqueado,0) AS PedidoBloqueado, " & vbCrLf &
                     "       isnull(p.LiberaCarregamento,0) AS LiberaCarregamento, isnull(P.VersaoPedido,0) as VersaoPedido," & vbCrLf &
                     "       isnull(P.VersaoUsuario,isnull(P.UsuarioAlteracao,P.UsuarioInclusao)) as VersaoUsuario," & vbCrLf &
                     "       isnull(P.VersaoHorarioBloqueio,CAST(GETDATE() AS DATETIME)) as VersaoHorarioBloqueio, isnull(P.IndexadorFixo,1) AS IndexadorFixo, " & vbCrLf &
                     "       cXe.DataFinanceiro, isnull(P.Troca,0) as Troca, isnull(P.Antecipada,0) as Antecipada, isnull(P.Recompra,0) as Recompra, " & vbCrLf &
                     "       isnull(p.XPedNFe,'') as XPedNFe, isnull(p.ItemXPedNFe,'') as ItemXPedNFe, isnull(P.InvoiceNavio,0) as InvoiceNavio, ISNULL(P.Embalagem, 0) AS Embalagem, " & vbCrLf &
                     "       ISNULL(EM.Descricao, '') AS DescricaoEmbalagem, ISNULL(P.TipoCondicaoEntrega, '') AS TipoCondicaoEntrega, ISNULL(P.TipoPagamentoPtax, '') AS TipoPagamentoPtax, " & vbCrLf &
                     "       ISNULL(P.LocalDeEmbarque, '') AS LocalDeEmbarque, isnull(ObservacoesControleInterno,'') as ObservacoesControleInterno" & vbCrLf &
                     "  FROM Pedidos P " & vbCrLf &
                     " INNER JOIN ClientesXEmpresas cXe " & vbCrLf &
                     "    ON P.Empresa_Id    = cXe.Empresa_Id" & vbCrLf &
                     "   AND P.EndEmpresa_Id = cXe.EndEmpresa_Id" & vbCrLf &
                     " LEFT JOIN Embalagens EM " & vbCrLf &
                     "    ON P.Embalagem    = EM.Embalagem_Id" & vbCrLf &
                     " Inner Join Suboperacoes so" & vbCrLf &
                     "    on p.operacao    = so.Operacao_id" & vbCrLf &
                     "   and p.suboperacao = so.suboperacoes_id" & vbCrLf &
                     " WHERE P.Empresa_Id    ='" & Empresa & "'" & vbCrLf &
                     "   AND P.EndEmpresa_Id = " & Endereco & vbCrLf &
                     "   AND P.Pedido_Id     = " & Codigo

            Dim dsPedido As DataSet = db.ConsultaDataSet(strSQL, "Pedidos")

            If dsPedido.Tables(0).Rows.Count > 0 Then
                Dim drPedido As DataRow = dsPedido.Tables(0).Rows(0)

                Me.CodigoEmpresa = drPedido("Empresa_Id")
                Me.EnderecoEmpresa = drPedido("EndEmpresa_Id")
                Me.Codigo = drPedido("Pedido_Id")
                Me.CodigoUnidadeNegocio = drPedido("UnidadeDeNegocio")
                Me.CodigoCliente = drPedido("Cliente")
                Me.EnderecoCliente = drPedido("EndCliente")
                Me.CodigoPraca = drPedido("Praca")
                Me.EnderecoPraca = drPedido("EndPraca")
                Me.PedidoEfetivo = drPedido("PedidoEfetivo").ToString
                Me.CodigoSafra = drPedido("Safra")
                Me.CodigoMoeda = drPedido("Moeda")
                Me.TemVariacao = drPedido("TemVariacao")

                Me.CodigoIndexador = drPedido("Indexador")
                Me.CodigoOperacao = drPedido("Operacao")
                Me.CodigoSubOperacao = drPedido("SubOperacao")
                Me.CodigoSituacao = drPedido("Situacao")
                Me.DataPedido = drPedido("DataPedido")

                Me.DataEntregaInicial = drPedido("DataInicioEntrega")
                Me.DataEntregaFinal = drPedido("DataEntrega")

                Me.FreteCIFFOB = [Enum].Parse(GetType(eTiposFrete), drPedido("FreteCIFFOB").ToString())
                Me.OrigemDestino = drPedido("OrigemDestino")
                Me.CodigoFinalidade = drPedido("Finalidade")
                Me.Solicitacao = drPedido("Solicitacao")
                Me.UsuarioInclusao = drPedido("UsuarioInclusao").ToString()
                Me.DataInclusao = Convert.ToDateTime(drPedido("UsuarioInclusaoData"))
                Me.UsuarioAlteracao = drPedido("UsuarioAlteracao").ToString()
                Me.DataAlteracao = Convert.ToDateTime(drPedido("UsuarioAlteracaoData"))
                Me.UsuarioCancelamento = drPedido("UsuarioCancelamento").ToString()
                Me.DataCancelamento = Convert.ToDateTime(drPedido("UsuarioCancelamentoData"))
                Me.UsuarioLiberacao = drPedido("UsuarioLiberacao").ToString()
                Me.DataLiberacao = Convert.ToDateTime(drPedido("UsuarioLiberacaoData"))
                Me.Observacoes = drPedido("Observacoes")
                Me.CodigoCondicaoPagamento = drPedido("CondicaoPagamento")
                Me.Contrato = drPedido("Contrato").ToString()
                Me.CodigoCarteira = drPedido("Carteira")
                Me.Taxa = drPedido("Taxa")
                Me.DataVencimentoPedido = drPedido("VencimentoPedido")
                Me.CodigoLocalEmbarque = drPedido("LocalEmbarque")
                Me.EndLocalEmbarque = drPedido("EndLocalEmbarque")

                If Not IsDBNull(drPedido("BancoCliente")) Then
                    If drPedido("BancoCliente") > 0 Then


                        If Not Me.CodigoCliente = Me.CodigoPraca Then
                            Dim intIndice As Integer = Me.Praca.ContasBancarias.FindIndex(Function(s) s.CodigoBanco = drPedido("BancoCliente") And s.CodigoAgencia = drPedido("AgenciaCliente") And s.DigitoAgencia = drPedido("DigitoAgenciaCliente") And s.ContaCorrente = drPedido("ContaCliente") And s.DigitoConta = drPedido("DigitoContaCliente"))
                            If intIndice > -1 Then Me.ContaBancariaSelecionada = Me.Praca.ContasBancarias(intIndice)
                        Else
                            Dim intIndice As Integer = Me.Cliente.ContasBancarias.FindIndex(Function(s) s.CodigoBanco = drPedido("BancoCliente") And s.CodigoAgencia = drPedido("AgenciaCliente") And s.DigitoAgencia = drPedido("DigitoAgenciaCliente") And s.ContaCorrente = drPedido("ContaCliente") And s.DigitoConta = drPedido("DigitoContaCliente"))
                            If intIndice > -1 Then Me.ContaBancariaSelecionada = Me.Cliente.ContasBancarias(intIndice)
                        End If
                    End If
                End If

                Me.Itens = New Negocio.ListPedidoXItem(Me)


                Me.Comercializacao = drPedido("Comercializacao")

                Me.CodigoCondicaoPagamento = drPedido("CondicaoPagamento")
                Me.MomentoFinanceiro = drPedido("MomentoFinanceiro")
                Me.AgruparFinanceiro = drPedido("AgruparFinanceiro")
                Me.IndiceFixado = drPedido("IndiceFixado")
                Me.EstadoEntrega = drPedido("EstadoEntrega")
                Me.CidadeEntrega = drPedido("CidadeEntrega")
                Me.FiscalAberto = drPedido("FiscalAberto")
                Me.FinanceiroAberto = drPedido("FinanceiroAberto")

                Me.CodigoEmpresaTroca = drPedido("EmpresaTroca")
                Me.EnderecoEmpresaTroca = drPedido("EndEmpresaTroca")
                Me.CodigoPedidoTroca = drPedido("PedidoTroca")
                Me.ContaAdiantamentoTroca = drPedido("ContaAdiantamentoTroca")

                Me.CodigoCondicaoPagamentoDaEntrega = drPedido("CondicaoPagamentoEntrega")
                Me.QuotaEntrega = drPedido("QuotaEntrega")
                Me.PeriodicidadeEntrega = drPedido("PeriodicidadeEntrega")
                Me.PedidoBloqueado = drPedido("PedidoBloqueado")
                Me.LiberaCarregamento = drPedido("LiberaCarregamento")

                Me.VersaoPedido = drPedido("VersaoPedido")
                Me.VersaoUsuario = drPedido("VersaoUsuario")
                Me.VersaoHorarioBloqueio = drPedido("VersaoHorarioBloqueio")

                Me.IndexadorFixo = drPedido("IndexadorFixo")

                Me.Troca = drPedido("Troca")
                Me.Antecipada = drPedido("Antecipada")
                Me.Recompra = drPedido("Recompra")

                Me.XPedNFe = drPedido("XPedNFe")
                Me.ItemXPedNFe = drPedido("ItemXPedNFe")

                Me.InvoiceNavio = drPedido("InvoiceNavio")

                Me.Embalagem = drPedido("Embalagem")
                Me.DescricaoEmbalagem = drPedido("DescricaoEmbalagem")
                Me.TipoCondicaoEntrega = drPedido("TipoCondicaoEntrega")
                Me.TipoPagamentoPtax = drPedido("TipoPagamentoPtax")
                Me.LocalDeEmbarque = drPedido("LocalDeEmbarque")

                Me.ObservacoesControleInterno = drPedido("ObservacoesControleInterno")

                'Verifica se o pedido foi aberto com financeiro novo
                If Not IsDBNull(drPedido("DataFinanceiro")) Then
                    Me.FinanceiroNovo = Me.DataPedido >= Convert.ToDateTime(drPedido("DataFinanceiro"))
                    If Me.MomentoFinanceiro = eTipoFaturamento.Peridiocidade Then
                        Me.Peridiocidade = [Enum].Parse(GetType(ePeriodicidade), drPedido("PeriodicidadeEntrega"))
                    End If
                End If
            End If
        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub New(objNota As NGS.Lib.Negocio.NotaFiscal, pCodigoCondicaoDePagamento As Integer)
        'PREENCHE O PEDIDO COM A NOTA
        Me.IUD = "I"
        Me.CodigoUnidadeNegocio = objNota.CodigoUnidadeDeNegocio
        Me.CodigoEmpresa = objNota.CodigoEmpresa
        Me.EnderecoEmpresa = objNota.EnderecoEmpresa
        Me.CodigoCliente = objNota.CodigoCliente
        Me.EnderecoCliente = objNota.EnderecoCliente
        Me.CodigoPraca = objNota.CodigoCliente
        Me.EnderecoPraca = objNota.EnderecoCliente
        Me.CodigoSafra = objNota.CodigoSafra
        Me.CodigoMoeda = 1
        Me.TemVariacao = False
        Me.CodigoIndexador = 99

        If Not objNota.Pedido Is Nothing Then
            Me.IndiceFixado = objNota.Pedido.IndiceFixado
        End If

        Me.CodigoOperacao = objNota.CodigoOperacao
        Me.CodigoSubOperacao = objNota.CodigoSubOperacao
        Me.CodigoSituacao = 1
        Me.DataPedido = objNota.DataNota

        Me.DataEntregaInicial = objNota.DataNota
        Me.DataEntregaFinal = objNota.DataNota

        If objNota.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            Me.FreteCIFFOB = eTiposFrete.FOB
        Else
            Me.FreteCIFFOB = eTiposFrete.CIF
        End If
        Me.OrigemDestino = "0"
        Me.CodigoFinalidade = 1
        Me.DataVencimentoPedido = objNota.DataNota
        Me.UsuarioInclusao = objNota.UsuarioInclusao
        Me.DataInclusao = objNota.DataInclusao
        If Me.SubOperacao.Financeiro Then
            Me.CodigoCondicaoPagamento = pCodigoCondicaoDePagamento
        End If

        Dim Produtos As New List(Of Integer)
        Me.InvoiceNavio = objNota.InvoiceNavio

        'PREENCHE OS ITENS DO PEDIDO
        For Each ItemNota In objNota.Itens.Where(Function(I) I.IUD <> "D")

            If Not Produtos.Contains(ItemNota.CodigoProduto) Then

                Dim Pitem As New PedidoXItem(Me)
                Pitem.CodigoProduto = ItemNota.CodigoProduto
                Pitem.CodigoClassificacao = 1
                Pitem.Descricao = ItemNota.Produto.Nome
                Pitem.Retencao = ItemNota.Retencao
                Pitem.CodigoOperacaoXEstado = ItemNota.CodigoOperacaoEstado
                Pitem.CodigoUnidadeComercializacao = ItemNota.Produto.Unidade

                For Each ItemNotaSequencia In objNota.Itens.Where(Function(I) I.IUD <> "D" And I.CodigoProduto = ItemNota.CodigoProduto)

                    'PREENCHE O LANCAMENTO NORMAL
                    Dim Lan As New LancamentoItemPedido(Pitem)
                    Lan.TipoLancamento = eTiposLancamentosPedidos.Normal
                    Lan.CodigoPedidoItem = ItemNotaSequencia.Sequencia
                    Lan.Movimento = objNota.DataNota
                    Lan.DataEntrega = objNota.DataNota
                    Lan.QuantidadeFaturamento = ItemNotaSequencia.QuantidadeFiscal

                    Lan.UnitarioOficial = ItemNotaSequencia.Unitario
                    Lan.TotalOficial = ItemNotaSequencia.ValorTotal
                    Lan.QuantidadeComercializacao = ItemNotaSequencia.QuantidadeFiscal
                    'If objNota.IndiceNota = 0 Then objNota.IndiceNota = New Cotacao(Me.CodigoIndexador, Me.DataPedido).Indice

                    If objNota.Pedido Is Nothing OrElse objNota.Pedido.IndiceFixado = 0 Then
                        Lan.UnitarioMoeda = 0
                        Lan.TotalMoeda = 0
                    Else
                        Lan.UnitarioMoeda = ItemNotaSequencia.UnitarioMoeda

                        If Not objNota.Pedido Is Nothing AndAlso objNota.Pedido.IndiceFixado > 0 Then
                            Lan.TotalMoeda = Math.Round(ItemNotaSequencia.ValorTotal / objNota.Pedido.IndiceFixado, 2)
                        End If

                    End If

                    Pitem.Lancamentos.Add(Lan)

                Next

                'PREENCHE ENCARGOS

                Pitem.Encargos.Clear()

                Dim ProdutosEncargos As New Dictionary(Of String, Integer)

                For Each ItemNotaEncargo In objNota.Itens.Where(Function(I) I.IUD <> "D" And I.CodigoProduto = ItemNota.CodigoProduto)

                    For i As Integer = 0 To ItemNotaEncargo.Encargos.Count - 1

                        Dim Penc As New PedidoXEncargo(Pitem)
                        Penc.CodigoEncargo = ItemNotaEncargo.Encargos(i).Codigo
                        Penc.Percentual = ItemNotaEncargo.Encargos(i).Percentual

                        Penc.BaseOficial = ItemNotaEncargo.Encargos(i).Base
                        Penc.ValorOficial = ItemNotaEncargo.Encargos(i).Valor

                        If objNota.Pedido Is Nothing OrElse objNota.Pedido.IndiceFixado = 0 Then

                            If Not ProdutosEncargos.ContainsKey(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo) Then

                                Penc.BaseMoeda = 0
                                Penc.ValorMoeda = 0

                                Pitem.Encargos.Add(Penc)
                                ProdutosEncargos.Add(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo, Pitem.Encargos.Count - 1)
                            Else

                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).BaseOficial += Penc.BaseOficial
                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).ValorOficial += Penc.ValorOficial

                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).BaseMoeda += 0
                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).ValorMoeda += 0

                            End If

                        Else

                            If Not ProdutosEncargos.ContainsKey(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo) Then

                                If Not objNota.Pedido Is Nothing AndAlso objNota.Pedido.IndiceFixado > 0 Then
                                    Penc.BaseMoeda = Math.Round(ItemNotaEncargo.Encargos(i).Base / objNota.Pedido.IndiceFixado, 2)
                                    Penc.ValorMoeda = Math.Round(ItemNotaEncargo.Encargos(i).Valor / objNota.Pedido.IndiceFixado, 2)
                                End If

                                Pitem.Encargos.Add(Penc)
                                ProdutosEncargos.Add(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo, Pitem.Encargos.Count - 1)
                            Else

                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).BaseOficial += Penc.BaseOficial
                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).ValorOficial += Penc.ValorOficial

                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).BaseMoeda += Math.Round(ItemNotaEncargo.Encargos(i).Base / objNota.IndiceNota, 2)
                                Pitem.Encargos(ProdutosEncargos(ItemNotaEncargo.CodigoProduto & " " & ItemNotaEncargo.Encargos(i).Codigo)).ValorMoeda += Math.Round(ItemNotaEncargo.Encargos(i).Valor / objNota.IndiceNota, 2)

                            End If

                        End If

                    Next

                Next

                Me.Itens.Add(Pitem)

                Produtos.Add(ItemNota.CodigoProduto)

            End If


        Next

        'PREENCHE DEPOSITO
        Dim DepositoES As New PedidoXDeposito(Me)
        DepositoES.Codigo = Me.CodigoEmpresa
        DepositoES.CodigoEndereco = Me.EnderecoEmpresa
        DepositoES.Principal = True
        DepositoES.Quantidade = 0
        DepositoES.Tipo = "DE"
        Me.Depositos.Add(DepositoES)

        'PREENCHE ORIGEM DESTINO
        Dim DepositoOD As New PedidoXDeposito(Me)
        DepositoOD.Codigo = Me.CodigoCliente
        DepositoOD.CodigoEndereco = Me.EnderecoCliente
        DepositoOD.Principal = True
        DepositoOD.Quantidade = 0
        DepositoOD.Tipo = "OD"
        Me.Depositos.Add(DepositoOD)

        'Preenche Parcelas
        Me.Parcelas.Clear()
        Dim x As Integer = 1
        For Each fin In objNota.VencimentosNota.Where(Function(s) s.IUD <> "D")
            Dim Parc As New PedidoXParcela(Me)
            Parc.IUD = "I"
            Parc.CodigoParcela = x
            x += 1
            Parc.DataVencimento = fin.Prorrogacao
            Parc.Valor = fin.ValorDoDocumento
            Me.Parcelas.Add(Parc)
        Next

    End Sub
#End Region

#Region "Fields"
    Private _financeiroNovo As Boolean = False
    Private _ResumoFinanceiro As ResumoFinanceiro
    Private _IUD As String = ""

    '****** Empresa
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente

    '****** Pedido
    Private _Codigo As Integer

    '****** Unidade de Negocio
    Private _CodigoUnidadeNegocio As String = ""
    Private _UnidadeNegocio As Cliente

    '****** Cliente
    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Cliente As Cliente
    Private _NomeCliente As String
    Public Property NomeCliente() As String
        Get
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property


    '****** Praca de Pagamento "Endereco"
    Private _CodigoPraca As String = ""
    Private _EnderecoPraca As Integer
    Private _Praca As Cliente

    '****** CN - Confirmacao de Negocio
    Private _PedidoEfetivo As String = ""

    '****** Codigo Safra
    Private _CodigoSafra As String = ""

    '****** Moeda
    Private _CodigoMoeda As Integer
    Private _Moeda As Moeda
    Private _TemVariacao As Boolean

    '****** Indexador
    Private _CodigoIndexador As Integer
    Private _Indexador As Indexador

    '****** Operacao
    Private _CodigoOperacao As Integer
    Private _Operacao As Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao

    '****** Situacao
    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    '****** Datas
    Private _DataPedido As DateTime
    Private _DataEntregaInicial As DateTime
    Private _DataEntregaFinal As DateTime
    Private _DataFinanceiroNovo As Boolean

    '****** Frete
    Private _FreteCIFFOB As eTiposFrete

    '****** Qual Balanca Vale
    Private _OrigemDestino As String = ""

    '****** Finalidade
    Private _CodigoFinalidade As Integer
    Private _Finalidade As Finalidade


    Private _Solicitacao As Integer
    Private _Observacoes As String = ""
    Private _Comercializacao As eComercializacao

    '****** Listas
    Private _Itens As ListPedidoXItem
    Private _Depositos As ListPedidoxDeposito
    Private _Roteiros As ListPedidoXRoteiro
    Private _Transportadores As ListPedidoXTransportador
    Private _Representantes As ListPedidoXRepresentante
    Private _Vencimentos As MovimentacoesFinanceiras
    Private _Titulos As Novo.ListTituloNovo
    Private _AdiantamentosAbertos As Novo.ListAdiantamentoNovo
    Private _Parcelas As ListPedidoXParcela
    Private _Financeiro As ListPedidoxFinanceiro
    Private _Contratos As ListPedidoXContrato

    Private _CodigoCondicaoPagamento As Integer
    Private _CondicaoPagamento As CondicaoPagamento
    Private _Periodicidade As ePeriodicidade

    Private _ContaBancariaSelecionada As ClienteXContaBancaria
    Private _Contrato As String = ""

    Private _CodigoCarteira As String = ""
    Private _CarteiraFinanceira As CarteiraFinanceira

    Private _DataVencimentoPedido As Date
    Private _Taxa As Double

    Private _CodigoLocalEmbarque As String = ""
    Private _EndLocalEmbarque As Integer
    Private _LocalEmbarque As Cliente

    Private _MomentoFinanceiro As Integer
    Private _AgruparFinanceiro As Boolean
    Private _IndiceFixado As Decimal

    Private _EstadoEntrega As String = ""
    Private _CidadeEntrega As String = ""

    Private _FiscalAberto As Boolean = True
    Private _FinanceiroAberto As Boolean = True

    Private _CodigoEmpresaTroca As String = ""
    Private _EnderecoEmpresaTroca As Integer
    Private _EmpresaTroca As Cliente
    Private _CodigoPedidoTroca As Integer
    Private _PedidoTroca As Pedido
    Private _ContaAdiantamentoTroca As String = ""

    '***************************************
    Private _CodigoCondicaoPagamentoDaEntrega As Integer
    Private _CondicaoPagamentoDaEntrega As CondicaoPagamento
    Private _QuotaEntrega As Decimal
    Private _PeriodicidadeEntrega As Integer

    Private _PedidoBloqueado As Boolean
    Private _LiberaCarregamento As Boolean

    '****** Controle de Alteracao
    Private _UsuarioInclusao As String = ""
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _UsuarioAlteracaoData As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _UsuarioCancelamentoData As DateTime
    Private _UsuarioLiberacao As String = ""
    Private _UsuarioLiberacaoData As DateTime

    '******* VERSIONAMENTO ********
    Private _VersaoPedido As Integer
    Private _VersaoUsuario As String = ""
    Private _VersaoHorarioBloqueio As DateTime

    Private _SaldoItensPedido As ListSaldoPedido2015

    Private _IndexadorFixo As Boolean = True
    Private _Antecipada As Boolean = False
    Private _Troca As Boolean = False
    Private _Recompra As Boolean = False

    '******* INFORMA��O PARA NFE ********
    Private _XPedNFe As String = ""
    Private _ItemXPedNFe As String = ""

    Private _Embalagem As Integer
    Private _DescricaoEmbalagem As String = ""
    Private _TipoCondicaoEntrega As String = ""
    Private _TipoPagamentoPtax As String = ""

    '******* INFORMA��O PARA CUSTO DO NAVIO ********
    Private _InvoiceNavio As Integer
    Private _LocalDeEmbarque As String

    Private _ObservacoesControleInterno As String = ""

#End Region

#Region "Propriedades"

    Public ReadOnly Property ProjetoFinanceiroNovo As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
        End Get
    End Property

    Public Property FinanceiroNovo As Boolean
        Get
            Return _financeiroNovo
        End Get
        Set(value As Boolean)
            _financeiroNovo = value
        End Set
    End Property

    Public Property ResumoFinanceiro As ResumoFinanceiro
        Get
            If _ResumoFinanceiro Is Nothing And Me.Codigo > 0 Then _ResumoFinanceiro = New ResumoFinanceiro(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.Codigo)
            Return _ResumoFinanceiro
        End Get
        Set(value As ResumoFinanceiro)
            _ResumoFinanceiro = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And Not Me.CodigoEmpresa Is Nothing Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
            _ResumoFinanceiro = Nothing
        End Set
    End Property

    Public Property CodigoUnidadeNegocio() As String
        Get
            Return _CodigoUnidadeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeNegocio = value
        End Set
    End Property

    Public Property UnidadeNegocio() As Cliente
        Get
            If _UnidadeNegocio Is Nothing And Not Me.CodigoUnidadeNegocio Is Nothing Then _UnidadeNegocio = New Cliente(Me.CodigoUnidadeNegocio, 0)
            Return _UnidadeNegocio
        End Get
        Set(ByVal value As Cliente)
            _UnidadeNegocio = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public ReadOnly Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And Not Me.CodigoCliente Is Nothing Then _Cliente = New Cliente(Me.CodigoCliente, Me.EnderecoCliente)
            Return _Cliente
        End Get
    End Property

    Public Property CodigoPraca() As String
        Get
            Return _CodigoPraca
        End Get
        Set(ByVal value As String)
            _CodigoPraca = value
            _Praca = Nothing
        End Set
    End Property

    Public Property EnderecoPraca() As Integer
        Get
            Return _EnderecoPraca
        End Get
        Set(ByVal value As Integer)
            _EnderecoPraca = value
            _Praca = Nothing
        End Set
    End Property

    Public ReadOnly Property Praca() As Cliente
        Get
            If _Praca Is Nothing And Not Me.CodigoPraca Is Nothing Then _Praca = New Cliente(Me.CodigoPraca, Me.EnderecoPraca)
            Return _Praca
        End Get
    End Property

    Public Property PedidoEfetivo() As String
        Get
            Return _PedidoEfetivo
        End Get
        Set(ByVal value As String)
            _PedidoEfetivo = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoMoeda() As Integer
        Get
            Return _CodigoMoeda
        End Get
        Set(ByVal value As Integer)
            _CodigoMoeda = value
            _Moeda = Nothing
        End Set
    End Property

    Public Property TemVariacao As Boolean
        Get
            Return _TemVariacao
        End Get
        Set(value As Boolean)
            _TemVariacao = value
        End Set
    End Property

    Public Property Moeda() As Moeda
        Get
            If _Moeda Is Nothing And _CodigoMoeda > 0 Then _Moeda = New Moeda(_CodigoMoeda)
            Return _Moeda
        End Get
        Set(ByVal value As Moeda)
            _Moeda = value
        End Set
    End Property

    Public Property CodigoIndexador() As Integer
        Get
            Return _CodigoIndexador
        End Get
        Set(ByVal value As Integer)
            _CodigoIndexador = value
            _Indexador = Nothing
        End Set
    End Property

    Public Property Indexador() As Indexador
        Get
            If _Indexador Is Nothing And _CodigoIndexador > 0 Then _Indexador = New Indexador(_CodigoIndexador)
            Return _Indexador
        End Get
        Set(ByVal value As Indexador)
            _Indexador = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property Operacao() As Operacao
        Get
            If _Operacao Is Nothing And CodigoOperacao > 0 Then _Operacao = New Operacao(CodigoOperacao)
            Return _Operacao
        End Get
        Set(ByVal value As Operacao)
            _Operacao = value
            _SubOperacao = Nothing
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property

    Public Property SubOperacao() As SubOperacao
        Get
            If _SubOperacao Is Nothing And CodigoOperacao > 0 And CodigoSubOperacao > 0 Then _SubOperacao = New SubOperacao(CodigoOperacao, CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As SubOperacao)
            _SubOperacao = value
        End Set
    End Property

    Public Property CodigoSituacao As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(value As Integer)
            _CodigoSituacao = value
        End Set
    End Property

    Public ReadOnly Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And Me.CodigoSituacao > 0 Then _Situacao = New Situacao(Me.CodigoSituacao)
            Return _Situacao
        End Get
    End Property

    Public Property DataPedido() As DateTime
        Get
            Return _DataPedido
        End Get
        Set(ByVal value As DateTime)
            _DataPedido = value
        End Set
    End Property

    Public Property DataEntregaInicial() As DateTime
        Get
            Return _DataEntregaInicial
        End Get
        Set(ByVal value As DateTime)
            _DataEntregaInicial = value
        End Set
    End Property

    Public Property DataEntregaFinal() As DateTime
        Get
            Return _DataEntregaFinal
        End Get
        Set(ByVal value As DateTime)
            _DataEntregaFinal = value
        End Set
    End Property

    Public Property FreteCIFFOB() As eTiposFrete
        Get
            Return _FreteCIFFOB
        End Get
        Set(ByVal value As eTiposFrete)
            _FreteCIFFOB = value
        End Set
    End Property

    Public Property OrigemDestino() As String
        Get
            Return _OrigemDestino
        End Get
        Set(ByVal value As String)
            _OrigemDestino = value
        End Set
    End Property

    Public Property CodigoFinalidade As Integer
        Get
            Return _CodigoFinalidade
        End Get
        Set(value As Integer)
            _CodigoFinalidade = value
            _Finalidade = Nothing
        End Set
    End Property

    Public ReadOnly Property Finalidade() As Finalidade
        Get
            If _Finalidade Is Nothing And Me.CodigoFinalidade > 0 Then _Finalidade = New Finalidade(_Codigo)
            Return _Finalidade
        End Get
    End Property

    Public Property Solicitacao() As Integer
        Get
            Return _Solicitacao
        End Get
        Set(ByVal value As Integer)
            _Solicitacao = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property DataLiberacao() As DateTime
        Get
            Return _UsuarioLiberacaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioLiberacaoData = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property Comercializacao() As eComercializacao
        Get
            Return _Comercializacao
        End Get
        Set(ByVal value As eComercializacao)
            _Comercializacao = value
        End Set
    End Property

    Public Property Itens() As ListPedidoXItem
        Get
            If _Itens Is Nothing Then _Itens = New ListPedidoXItem(Me)
            Return _Itens
        End Get
        Set(ByVal value As ListPedidoXItem)
            _Itens = value
        End Set
    End Property

    Public Property Depositos() As ListPedidoxDeposito
        Get
            If _Depositos Is Nothing OrElse _Depositos.Count() = 0 Then _Depositos = New ListPedidoxDeposito(Me)
            Return _Depositos
        End Get
        Set(ByVal value As ListPedidoxDeposito)
            _Depositos = value
        End Set
    End Property

    Public Property Roteiros() As ListPedidoXRoteiro
        Get
            If _Roteiros Is Nothing Then _Roteiros = New ListPedidoXRoteiro(Me)
            Return _Roteiros
        End Get
        Set(ByVal value As ListPedidoXRoteiro)
            _Roteiros = value
        End Set
    End Property

    Public Property Transportadores() As ListPedidoXTransportador
        Get
            If _Transportadores Is Nothing Then _Transportadores = New ListPedidoXTransportador(Me)
            Return _Transportadores
        End Get
        Set(ByVal value As ListPedidoXTransportador)
            _Transportadores = value
        End Set
    End Property

    Public Property Representantes() As ListPedidoXRepresentante
        Get
            If _Representantes Is Nothing Then _Representantes = New ListPedidoXRepresentante(Me)
            Return _Representantes
        End Get
        Set(ByVal value As ListPedidoXRepresentante)
            _Representantes = value
        End Set
    End Property

    Public ReadOnly Property TemFixacoes() As Boolean
        Get
            For Each objItem As PedidoXItem In Me.Itens
                If objItem.Fixacoes.Count > 0 AndAlso objItem.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then Return True
            Next

            Return False
        End Get
    End Property

    Public Property Contrato() As String
        Get
            Return _Contrato
        End Get
        Set(ByVal value As String)
            _Contrato = value
        End Set
    End Property

    Public Property CodigoCarteira() As String
        Get
            Return _CodigoCarteira
        End Get
        Set(ByVal value As String)
            _CodigoCarteira = value
            _CarteiraFinanceira = Nothing
        End Set
    End Property

    Public Property CarteiraFinanceira() As CarteiraFinanceira
        Get
            If _CodigoCarteira.Length > 0 Then _CarteiraFinanceira = New CarteiraFinanceira(_CodigoCarteira)
            Return _CarteiraFinanceira
        End Get
        Set(ByVal value As CarteiraFinanceira)
            _CarteiraFinanceira = value
        End Set
    End Property

    Public Property DataVencimentoPedido() As Date
        Get
            Return _DataVencimentoPedido
        End Get
        Set(ByVal value As Date)
            _DataVencimentoPedido = value
        End Set
    End Property

    Public Property Taxa() As Double
        Get
            Return _Taxa
        End Get
        Set(ByVal value As Double)
            _Taxa = value
        End Set
    End Property

    Public Property CodigoLocalEmbarque() As String
        Get
            Return _CodigoLocalEmbarque
        End Get
        Set(ByVal value As String)
            _CodigoLocalEmbarque = value
            _LocalEmbarque = Nothing
        End Set
    End Property

    Public Property EndLocalEmbarque() As Integer
        Get
            Return _EndLocalEmbarque
        End Get
        Set(ByVal value As Integer)
            _EndLocalEmbarque = value
            _LocalEmbarque = Nothing
        End Set
    End Property

    Public ReadOnly Property LocalEmbarque() As Cliente
        Get
            If _LocalEmbarque Is Nothing And _CodigoLocalEmbarque.Length > 0 Then _LocalEmbarque = New Cliente(_CodigoLocalEmbarque, _EndLocalEmbarque)
            Return _LocalEmbarque
        End Get
    End Property

    Public Property IndiceFixado() As Decimal
        Get
            Return _IndiceFixado
        End Get
        Set(ByVal value As Decimal)
            _IndiceFixado = value
        End Set
    End Property

    Public Property EstadoEntrega() As String
        Get
            Return _EstadoEntrega
        End Get
        Set(ByVal value As String)
            _EstadoEntrega = value
        End Set
    End Property

    Public Property CidadeEntrega() As String
        Get
            Return _CidadeEntrega
        End Get
        Set(ByVal value As String)
            _CidadeEntrega = value
        End Set
    End Property

    Public Property FiscalAberto() As Boolean
        Get
            Return _FiscalAberto
        End Get
        Set(ByVal value As Boolean)
            _FiscalAberto = value
        End Set
    End Property

    Public Property FinanceiroAberto() As Boolean
        Get
            Return _FinanceiroAberto
        End Get
        Set(ByVal value As Boolean)
            _FinanceiroAberto = value
        End Set
    End Property

    Public Property CodigoEmpresaTroca() As String
        Get
            Return _CodigoEmpresaTroca
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaTroca = value
        End Set
    End Property

    Public Property EnderecoEmpresaTroca() As Integer
        Get
            Return _EnderecoEmpresaTroca
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresaTroca = value
        End Set
    End Property

    Public Property EmpresaTroca() As Cliente
        Get
            If _EmpresaTroca Is Nothing And Not Me.CodigoEmpresaTroca Is Nothing Then _EmpresaTroca = New Cliente(Me.CodigoEmpresaTroca, Me.EnderecoEmpresaTroca)
            Return _EmpresaTroca
        End Get
        Set(ByVal value As Cliente)
            _EmpresaTroca = value
        End Set
    End Property

    Public Property CodigoPedidoTroca() As Integer
        Get
            Return _CodigoPedidoTroca
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoTroca = value
        End Set
    End Property

    Public Property PedidoTroca() As Pedido
        Get
            If _PedidoTroca Is Nothing And Me.CodigoPedidoTroca > 0 Then _PedidoTroca = New Pedido(Me.CodigoEmpresaTroca, Me.EnderecoEmpresaTroca, Me.CodigoPedidoTroca)
            Return _PedidoTroca
        End Get
        Set(value As Pedido)
            _PedidoTroca = value
        End Set
    End Property

    Public Property ContaAdiantamentoTroca() As String
        Get
            Return _ContaAdiantamentoTroca
        End Get
        Set(value As String)
            _ContaAdiantamentoTroca = value
        End Set
    End Property

    Public Property PedidoBloqueado As Boolean
        Get
            Return _PedidoBloqueado
        End Get
        Set(value As Boolean)
            _PedidoBloqueado = value
        End Set
    End Property

    Public Property LiberaCarregamento As Boolean
        Get
            Return _LiberaCarregamento
        End Get
        Set(value As Boolean)
            _LiberaCarregamento = value
        End Set
    End Property

    Public Property CodigoCondicaoPagamentoDaEntrega As Integer
        Get
            Return _CodigoCondicaoPagamentoDaEntrega
        End Get
        Set(value As Integer)
            _CodigoCondicaoPagamentoDaEntrega = value
            _CondicaoPagamentoDaEntrega = Nothing
        End Set
    End Property

    Public ReadOnly Property CondicaoPagamentoDaEntrega() As CondicaoPagamento
        Get
            If _CondicaoPagamentoDaEntrega Is Nothing And Me.CodigoCondicaoPagamentoDaEntrega > 0 Then _CondicaoPagamentoDaEntrega = New CondicaoPagamento(Me.CodigoCondicaoPagamentoDaEntrega)
            Return _CondicaoPagamentoDaEntrega
        End Get
    End Property

    Public Property QuotaEntrega() As Decimal
        Get
            Return _QuotaEntrega
        End Get
        Set(value As Decimal)
            _QuotaEntrega = value
        End Set
    End Property

    Public Property PeriodicidadeEntrega() As Integer
        Get
            Return _PeriodicidadeEntrega
        End Get
        Set(ByVal value As Integer)
            _PeriodicidadeEntrega = value
        End Set
    End Property

    '******* FINANCEIRO ***********
    Public Property Vencimentos() As PedidoXParcelas
        Get
            If _Vencimentos Is Nothing Then _Vencimentos = New PedidoXParcelas(Me)
            Return _Vencimentos
        End Get
        Set(ByVal value As PedidoXParcelas)
            _Vencimentos = value
        End Set
    End Property

    Public Property Titulos() As Novo.ListTituloNovo
        Get
            If _Titulos Is Nothing OrElse _Titulos.Count = 0 Then
                If Me.MomentoFinanceiro = eTipoFaturamento.Lote Then
                    _Titulos = New Novo.ListTituloNovo(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.Codigo, 0, Me.MomentoFinanceiro)
                Else
                    _Titulos = New Novo.ListTituloNovo(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.Codigo)
                End If

            End If
            Return _Titulos
        End Get
        Set(ByVal value As Novo.ListTituloNovo)
            _Titulos = value
        End Set
    End Property

    Public Property CodigoCondicaoPagamento As Integer
        Get
            Return _CodigoCondicaoPagamento
        End Get
        Set(value As Integer)
            _CodigoCondicaoPagamento = value
            _CondicaoPagamento = Nothing
        End Set
    End Property

    Public Property Peridiocidade As ePeriodicidade
        Get
            Return _Periodicidade
        End Get
        Set(value As ePeriodicidade)
            _Periodicidade = value
        End Set
    End Property


    Public ReadOnly Property CondicaoPagamento() As CondicaoPagamento
        Get
            If _CondicaoPagamento Is Nothing And Me.CodigoCondicaoPagamento > 0 Then _CondicaoPagamento = New CondicaoPagamento(Me.CodigoCondicaoPagamento)
            Return _CondicaoPagamento
        End Get
    End Property

    Public Property ContaBancariaSelecionada() As ClienteXContaBancaria
        Get
            Return _ContaBancariaSelecionada
        End Get
        Set(ByVal value As ClienteXContaBancaria)
            _ContaBancariaSelecionada = value
        End Set
    End Property

    Public Property MomentoFinanceiro() As Integer
        Get
            Return _MomentoFinanceiro
        End Get
        Set(ByVal value As Integer)
            _MomentoFinanceiro = value
        End Set
    End Property

    Public Property AgruparFinanceiro() As Boolean
        Get
            Return _AgruparFinanceiro
        End Get
        Set(ByVal value As Boolean)
            _AgruparFinanceiro = value
        End Set
    End Property

    Public Property AdiantamentosAbertos() As Novo.ListAdiantamentoNovo
        Get
            If _AdiantamentosAbertos Is Nothing Then
                Dim Par As New Hashtable
                Par.Add("CodigoCliente", Me.CodigoCliente)
                Par.Add("EndCliente", Me.EnderecoCliente)
                Par.Add("ConsolidarCliente", False)
                Par.Add("SomenteComSaldo", True)

                Par.Add("CodigoEmpresa", Me.CodigoEmpresa)
                Par.Add("EndEmpresa", Me.EnderecoEmpresa)
                Par.Add("CodigoPedido", Me.Codigo)

                If Me.Troca Then
                    Par("isTroca") = True

                    If Me.ResumoFinanceiro.SaldoAdiantamento = 0 Then
                        Par("ContaAdiantamento") = Me.ResumoFinanceiro.ResumoTroca.ContaContabilAdiantamento
                    End If

                    Par.Add("CodigoEmpresaTroca", Me.CodigoEmpresaTroca)
                    Par.Add("EndEmpresaTroca", Me.EnderecoEmpresaTroca)
                    Par.Add("CodigoPedidoTroca", Me.CodigoPedidoTroca)
                End If

                _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)
            End If
            Return _AdiantamentosAbertos
        End Get
        Set(ByVal value As Novo.ListAdiantamentoNovo)
            _AdiantamentosAbertos = value
        End Set
    End Property

    Public Property AdiantamentosTodos() As Novo.ListAdiantamentoNovo
        Get
            If _AdiantamentosAbertos Is Nothing Then
                Dim Par As New Hashtable
                Par.Add("CodigoCliente", Me.CodigoCliente)
                Par.Add("EndCliente", Me.EnderecoCliente)
                Par.Add("ConsolidarCliente", False)
                Par.Add("SomenteComSaldo", False)

                Par.Add("CodigoEmpresa", Me.CodigoEmpresa)
                Par.Add("EndEmpresa", Me.EnderecoEmpresa)
                Par.Add("CodigoPedido", Me.Codigo)

                If Me.Troca Then
                    Par("isTroca") = True

                    If Me.ResumoFinanceiro.SaldoAdiantamento = 0 Then
                        Par("ContaAdiantamento") = Me.ResumoFinanceiro.ResumoTroca.ContaContabilAdiantamento
                    End If

                    Par.Add("CodigoEmpresaTroca", Me.CodigoEmpresaTroca)
                    Par.Add("EndEmpresaTroca", Me.EnderecoEmpresaTroca)
                    Par.Add("CodigoPedidoTroca", Me.CodigoPedidoTroca)
                End If

                _AdiantamentosAbertos = New Novo.ListAdiantamentoNovo(Par)
            End If
            Return _AdiantamentosAbertos
        End Get
        Set(ByVal value As Novo.ListAdiantamentoNovo)
            _AdiantamentosAbertos = value
        End Set
    End Property

    '*************************************************************************************************************
    '****************************  Financeiro Virtual e Resumos **************************************************
    '*************************************************************************************************************

    Public ReadOnly Property Financeiro As ListPedidoxFinanceiro
        Get
            If _Financeiro Is Nothing Then _Financeiro = New ListPedidoxFinanceiro(Me)
            Return _Financeiro
        End Get
    End Property

    Public Property Parcelas As ListPedidoXParcela
        Get
            If _Parcelas Is Nothing Then _Parcelas = New ListPedidoXParcela(Me)
            Return _Parcelas
        End Get
        Set(value As ListPedidoXParcela)
            _Parcelas = value
        End Set
    End Property

    Public ReadOnly Property ParcelasTitulosVituais As ListPedidoXTituloVirtual
        Get
            Return New ListPedidoXTituloVirtual(Me)
        End Get
    End Property

    '******* VERSIONAMENTO ********
    Public Property VersaoPedido As Integer
        Get
            Return _VersaoPedido
        End Get
        Set(value As Integer)
            _VersaoPedido = value
        End Set
    End Property

    Public Property VersaoUsuario As String
        Get
            Return _VersaoUsuario
        End Get
        Set(value As String)
            _VersaoUsuario = value
        End Set
    End Property

    Public Property VersaoHorarioBloqueio As DateTime
        Get
            Return _VersaoHorarioBloqueio
        End Get
        Set(value As DateTime)
            _VersaoHorarioBloqueio = value
        End Set
    End Property

    '******** Lista dos saldos dos itens do Pedido *******
    Public Property SaldoItensPedido As ListSaldoPedido2015
        Get
            If _SaldoItensPedido Is Nothing Then
                If Me.Codigo > 0 Then
                    Dim Parametros As New Hashtable
                    Parametros.Add("Empresa", Me.CodigoEmpresa)
                    Parametros.Add("EndEmpresa", Me.EnderecoEmpresa)
                    Parametros.Add("Pedido", Me.Codigo)
                    Parametros.Add("TipoApuracao", 1)
                    _SaldoItensPedido = New ListSaldoPedido2015(Parametros)
                Else
                    _SaldoItensPedido = New ListSaldoPedido2015()
                End If
            End If
            Return _SaldoItensPedido
        End Get
        Set(value As ListSaldoPedido2015)
            _SaldoItensPedido = value
        End Set
    End Property

    Public Property IndexadorFixo As Boolean
        Get
            Return _IndexadorFixo
        End Get
        Set(value As Boolean)
            _IndexadorFixo = value
        End Set
    End Property

    Public Property Antecipada As Boolean
        Get
            Return _Antecipada
        End Get
        Set(value As Boolean)
            _Antecipada = value
        End Set
    End Property

    Public Property Troca As Boolean
        Get
            Return _Troca
        End Get
        Set(value As Boolean)
            _Troca = value
        End Set
    End Property

    Public Property Recompra As Boolean
        Get
            Return _Recompra
        End Get
        Set(value As Boolean)
            _Recompra = value
        End Set
    End Property

    Public Property Contratos() As ListPedidoXContrato
        Get
            If _Contratos Is Nothing Then _Contratos = New ListPedidoXContrato(Me)
            Return _Contratos
        End Get
        Set(ByVal value As ListPedidoXContrato)
            _Contratos = value
        End Set
    End Property

    Public Property XPedNFe() As String
        Get
            Return _XPedNFe
        End Get
        Set(ByVal value As String)
            _XPedNFe = value
        End Set
    End Property

    Public Property ItemXPedNFe() As String
        Get
            Return _ItemXPedNFe
        End Get
        Set(ByVal value As String)
            _ItemXPedNFe = value
        End Set
    End Property

    Public Property InvoiceNavio() As Integer
        Get
            Return _InvoiceNavio
        End Get
        Set(ByVal value As Integer)
            _InvoiceNavio = value
        End Set
    End Property

    Public Property Embalagem As Integer
        Get
            Return _Embalagem
        End Get
        Set(value As Integer)
            _Embalagem = value
        End Set
    End Property

    Public Property TipoCondicaoEntrega As String
        Get
            Return _TipoCondicaoEntrega
        End Get
        Set(value As String)
            _TipoCondicaoEntrega = value
        End Set
    End Property

    Public Property TipoPagamentoPtax As String
        Get
            Return _TipoPagamentoPtax
        End Get
        Set(value As String)
            _TipoPagamentoPtax = value
        End Set
    End Property

    Public Property DescricaoEmbalagem As String
        Get
            Return _DescricaoEmbalagem
        End Get
        Set(value As String)
            _DescricaoEmbalagem = value
        End Set
    End Property

    Public Property LocalDeEmbarque As String
        Get
            Return _LocalDeEmbarque
        End Get
        Set(value As String)
            _LocalDeEmbarque = value
        End Set
    End Property

    Public Property ObservacoesControleInterno() As String
        Get
            Return _ObservacoesControleInterno
        End Get
        Set(ByVal value As String)
            _ObservacoesControleInterno = value
        End Set
    End Property

#End Region

#Region "Funções"

    Public Function BloquearPedido() As Boolean

        Dim sql As String = "update Pedidos set PedidoBloqueado = 1 " & vbCrLf &
                            "where Empresa_Id    =   '" & _CodigoEmpresa & "'" & vbCrLf &
                            "  and EndEmpresa_id =   " & _EnderecoEmpresa & vbCrLf &
                            "  and Pedido_Id     =   " & _Codigo

        Return New AcessaBanco().GravaBanco(sql)

    End Function

    Public Function AbrirFecharPedido(ByVal ajustaFinanceiro As Boolean) As Boolean
        Dim arrSQL As New ArrayList()

        Dim strSQL = "UPDATE Pedidos " & vbCrLf &
                   "   SET PedidoBloqueado          =  " & IIf(Me.PedidoBloqueado, 1, 0) & vbCrLf &
                   "      ,DataInicioEntrega        = '" & Me.DataEntregaInicial.ToString("yyyy-MM-dd") & "'" &
                   "      ,DataEntrega              = '" & Me.DataEntregaFinal.ToString("yyyy-MM-dd") & "'" &
                   "      ,CondicaoPagamento        =  " & Me.CodigoCondicaoPagamento.ToString() &
                   "      ,MomentoFinanceiro        =  " & Me.MomentoFinanceiro &
                   "      ,CondicaoPagamentoEntrega =  " & Me.CodigoCondicaoPagamentoDaEntrega.ToString() &
                   "      ,QuotaEntrega             =  " & Str(Me.QuotaEntrega) &
                   "      ,PeriodicidadeEntrega     =  " & Me.PeriodicidadeEntrega &
                   "      ,UsuarioLiberacao         = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" &
                   "      ,UsuarioLiberacaoData     = CURRENT_TIMESTAMP " &
                   " WHERE Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                   "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa.ToString() & vbCrLf &
                   "   AND Pedido_Id     = " & Me.Codigo.ToString()

        arrSQL.Add(strSQL)

        If ajustaFinanceiro Then
            arrSQL.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
            arrSQL.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), _MomentoFinanceiro))
        End If

        Return New AcessaBanco().GravaBanco(arrSQL)
    End Function

    Public Function AtualizaEmbarqueAtivo(ByVal Situacao As Integer) As Boolean
        Dim sql = " UPDATE Pedidos " & vbCrLf &
                  "    SET EmbarqueAtivo =  " & Situacao & vbCrLf &
                  "  WHERE Empresa_Id    = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "    AND EndEmpresa_Id =  " & Me.EnderecoEmpresa.ToString() & vbCrLf &
                  "    AND Pedido_Id     =  " & Me.Codigo.ToString()

        Return New AcessaBanco().GravaBanco(sql)
    End Function

    Public Function VerificarExclusao() As eValidacaoExclusaoPedido
        'Verificar se tem alguma baixa no financeiro deste pedido
        If PedidoXParcelas.Existe(Me.Codigo, Me.SubOperacao.EntradaSaida, eProvisao.Baixa) Then Return eValidacaoExclusaoPedido.TemBaixaFinanceira

        'Verificar se tem alguma nota fiscal deste pedido
        If ListNotaFiscalXItem.Existe(Me.Codigo, Me.CodigoEmpresa, Me.EnderecoEmpresa) Then Return eValidacaoExclusaoPedido.UsadoEmNotaFiscal

        'Verificar se tem alguma Cess�o de Cr�dito com este pedido
        If ListProcuracao.Existe(Me.Codigo) Then Return eValidacaoExclusaoPedido.UsadoEmProcuracao

        'Verificar se tem laudo lan�ado para este pedido
        If LaudosPesagem.Existe(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.Codigo, False) Then Return eValidacaoExclusaoPedido.UsadoEmLaudo

        Return eValidacaoExclusaoPedido.Nenhum
    End Function

    'N�o Usar
    'Public Function RemoverTroca() As Boolean
    '    Dim objBanco As New AcessaBanco()
    '    Dim arrSQL As New ArrayList()

    '    Dim strSQL As String = "UPDATE Pedidos SET" & _
    '           "   EmpresaTroca = ''" & _
    '           "  ,EndEmpresaTroca = 0 " & _
    '           "  ,PedidoTroca = 0 " & _
    '           " WHERE Empresa_Id = '" & Me.CodigoEmpresaTroca & "' " & _
    '           "   AND EndEmpresa_Id = " & Me.EnderecoEmpresaTroca.ToString() & _
    '           "   AND Pedido_Id = " & Me.CodigoPedidoTroca

    '    arrSQL.Add(strSQL)

    '    strSQL = "UPDATE Pedidos SET" & _
    '           "   EmpresaTroca = ''" & _
    '           "  ,EndEmpresaTroca = 0 " & _
    '           "  ,PedidoTroca = 0 " & _
    '           " WHERE Empresa_Id = '" & Me.CodigoEmpresa & "' " & _
    '           "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa.ToString() & _
    '           "   AND Pedido_Id = " & Me.Codigo

    '    arrSQL.Add(strSQL)

    '    If objBanco.GravaBanco(arrSQL) Then
    '        Return True
    '    Else
    '        Return False
    '    End If
    'End Function

    Public Function TemPedidoEfetivo() As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "select PedidoEfetivo from Pedidos " & vbCrLf &
              "where Situacao = 1  " & vbCrLf &
              "  and Empresa_Id    =   '" & _CodigoEmpresa & "'" & vbCrLf &
              "  and EndEmpresa_id =   " & _EnderecoEmpresa & vbCrLf &
              "  and PedidoEfetivo =   '" & _PedidoEfetivo & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "PedidoEfetivo")
        If ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function TemPedidoContrato() As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "select PedidoEfetivo from Pedidos " & vbCrLf &
              "where Situacao = 1  " & vbCrLf &
              "  and Empresa_Id    =   '" & _CodigoEmpresa & "'" & vbCrLf &
              "  and EndEmpresa_id =   " & _EnderecoEmpresa & vbCrLf &
              "  and Contrato      =   '" & _Contrato & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "PedidoContrato")
        If ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function TemPesagem() As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "select top 1 1 from Pesagem " & vbCrLf &
              "where situacao = 1  " & vbCrLf &
              "  and Sequencia_Id = 0 " & vbCrLf &
              "  and empresa_Id    =   '" & _CodigoEmpresa & "'" & vbCrLf &
              "  and endempresa_id =   " & _EnderecoEmpresa & vbCrLf &
              "  and pedido        =   " & _Codigo & vbCrLf


        ds = Banco.ConsultaDataSet(sql, "Mov")
        If ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function TemNFG() As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT 1" & vbCrLf &
              "  FROM Pedidos P " & vbCrLf &
              " INNER JOIN NotasFiscais NF " & vbCrLf &
              "    ON P.Empresa_id    = NF.Empresa_id" & vbCrLf &
              "   And P.EndEmpresa_id = NF.EndEmpresa_Id" & vbCrLf &
              "   AND P.Pedido_Id     = NF.Pedido" & vbCrLf &
              " WHERE isnull(NF.NFG,0)   = 1 " & vbCrLf &
              "   AND P.empresa_Id       ='" & _CodigoEmpresa & "'" & vbCrLf &
              "   AND P.endempresa_id    = " & _EnderecoEmpresa & vbCrLf &
              "   AND P.Pedido_Id        = " & _Codigo

        ds = Banco.ConsultaDataSet(sql, "NFG")

        Return ds.Tables(0).Rows.Count > 0
    End Function

    Public Function TemFaturamento() As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "select top 1 1 from Notasfiscais " & vbCrLf &
              "where situacao in (1,4,7) " & vbCrLf &
              "  and empresa_Id    = '" & _CodigoEmpresa & "'" & vbCrLf &
              "  and endempresa_id = " & _EnderecoEmpresa & vbCrLf &
              "  and pedido        = " & _Codigo

        ds = Banco.ConsultaDataSet(sql, "Mov")
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function TemFaturamentoComDeposito(indice As Integer) As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = " select TOP 1 1 " & vbCrLf &
              "   FROM Notasfiscais " & vbCrLf &
              "  WHERE Situacao in (1,4,7) " & vbCrLf &
              "    AND empresa_Id    = '" & _CodigoEmpresa & "'" & vbCrLf &
              "    AND endempresa_id = " & _EnderecoEmpresa & vbCrLf &
              "    AND pedido        = " & _Codigo & vbCrLf

        If Me.Depositos(indice).Tipo = "DE" Then
            sql &= "  AND Deposito = '" & Me.Depositos(indice).Codigo & "'" & vbCrLf &
                   "  AND EndDeposito = " & Me.Depositos(indice).CodigoEndereco & vbCrLf
        ElseIf Me.Depositos(indice).Tipo = "LE" Then
            sql &= "  AND LocalEmbarque = '" & Me.Depositos(indice).Codigo & "'" & vbCrLf &
                   "  AND EndLocalEmbarque = " & Me.Depositos(indice).CodigoEndereco & vbCrLf
        ElseIf Me.Depositos(indice).Tipo = "TR" Then
            sql &= "  AND Transbordo = '" & Me.Depositos(indice).Codigo & "'" & vbCrLf &
                   "  AND EndTransbordo = " & Me.Depositos(indice).CodigoEndereco & vbCrLf
        ElseIf Me.Depositos(indice).Tipo = "OD" Then
            sql &= "  AND Destino = '" & Me.Depositos(indice).Codigo & "'" & vbCrLf &
                   "  AND EndDestino = " & Me.Depositos(indice).CodigoEndereco & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "Mov")
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function TemFinanceiro(Optional ByVal CodigoProvisao As Integer = 0) As Boolean
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = " SELECT top 1 1 " & vbCrLf &
              "   FROM ContasAReceber " & vbCrLf &
              "  WHERE  Pedido = " & _Codigo & IIf(CodigoProvisao > 0, " And Provisao = " & CodigoProvisao, "") & vbCrLf &
              "  AND EmpresaPedido    = '" & _CodigoEmpresa & "'" & vbCrLf &
              "   AND EndempresaPedido = " & _EnderecoEmpresa & vbCrLf &
              "   AND Situacao         = 1 " & vbCrLf &
              " UNION ALL" & vbCrLf &
              "SELECT top 1 1 " & vbCrLf &
              "  FROM ContasAPagar   " & vbCrLf &
              " WHERE Pedido = " & _Codigo & IIf(CodigoProvisao > 0, " And Provisao = " & CodigoProvisao, "") & vbCrLf &
              "  AND EmpresaPedido    = '" & _CodigoEmpresa & "'" & vbCrLf &
              "   AND EndempresaPedido = " & _EnderecoEmpresa & vbCrLf &
              "   AND Situacao         = 1 " & vbCrLf
        If FinanceiroNovo Then
            sql &= " UNION ALL" & vbCrLf &
                   " Select Top 1 1" & vbCrLf &
                   "   from Titulos" & vbCrLf &
                   "  WHERE Pedido     = " & _Codigo & IIf(CodigoProvisao > 0, " And Provisao = " & CodigoProvisao, "") & vbCrLf &
                   "    AND Empresa    ='" & _CodigoEmpresa & "'" & vbCrLf &
                   "    AND Endempresa = " & _EnderecoEmpresa & vbCrLf &
                   "    AND Situacao   = 1 " & vbCrLf

        End If
        ds = Banco.ConsultaDataSet(sql, "Mov")
        If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    'Public Function AlteracaoIndiceFixado() As Boolean
    '    Try
    '        'Lancamentos
    '        For Each lan In Me.Itens.SelectMany(Function(s) s.Lancamentos)
    '            If lan.IUD <> "I" Then
    '                lan.ItemPedido.IUD = "U"
    '                lan.IUD = "U"
    '            End If

    '            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
    '                lan.UnitarioMoeda = Funcoes.ConverteMoeda(lan.UnitarioOficial, Me.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
    '                lan.TotalMoeda = Funcoes.ConverteMoeda(lan.TotalOficial, Me.IndiceFixado, eTiposMoeda.MoedaEstrangeira, True, False, 2)
    '            Else
    '                lan.UnitarioOficial = Funcoes.ConverteMoeda(lan.UnitarioMoeda, Me.IndiceFixado, eTiposMoeda.Oficial, True, False, 2)
    '                lan.TotalOficial = Funcoes.ConverteMoeda(lan.TotalMoeda, Me.IndiceFixado, eTiposMoeda.Oficial, True, False, 2)
    '            End If
    '        Next

    '        'Itens e Encargos
    '        For Each Item In Me.Itens
    '            If Me.Moeda.Classificacao = eTiposMoeda.Oficial Then
    '                Item.SaldoItem.SaldoValor = Item.PedidoValor
    '            Else
    '                Item.SaldoItem.SaldoValor = Item.PedidoValor * Me.IndiceFixado
    '            End If

    '            Item.Encargos = Nothing
    '            Item.Encargos.CriaListar()
    '        Next

    '        Me.Representantes.RecalcularComissoesFixas()

    '    Catch ex As Exception
    '        Return False
    '    End Try

    '    Return True
    'End Function

    Public Function TemReferenciaEmOutroPedidoDeTroca() As Integer
        If Me.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
            Dim ds As DataSet
            Dim Banco As New AcessaBanco
            Dim sql As String

            sql = " SELECT Pedido_Id " & vbCrLf &
                  "   FROM Pedidos " & vbCrLf &
                  "  WHERE EmpresaTroca    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "    AND EndEmpresaTroca = " & Me.EnderecoEmpresa & vbCrLf &
                  "    AND PedidoTroca     = " & Me.Codigo & vbCrLf &
                  "    And Situacao        = 1"
            ds = Banco.ConsultaDataSet(sql, "Referencia")
            If ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                Return ds.Tables(0).Rows(0)("Pedido_Id")
            Else
                Return 0
            End If
        End If
        Return 0
    End Function


#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        If FinanceiroNovo AndAlso Not String.IsNullOrWhiteSpace(Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> Empresa.Empresa.Servidor Then
            Banco = New AcessaBanco(2, Empresa.Empresa.Servidor)
        End If
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        If FinanceiroNovo Then getSqlException(Sqls)
        Dim aux As Boolean = Banco.GravaBanco(Sqls)
        If FinanceiroNovo AndAlso Not aux Then
            Liberar()
        End If
        Return aux
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList, Optional PreencherNumerador As Boolean = False)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I", "U"
                If PreencherNumerador And Me.IUD = "I" Then
                    Dim num As New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 10)
                    Me.Codigo = num.Sequencia
                    Sqls.Add(num.IncrementarNumeradorSql)
                End If

                Sql = "; Merge Pedidos as Dest" & vbCrLf &
                      " USING (Select '" & Me.CodigoEmpresa & "' as Empresa_Id," & Me.EnderecoEmpresa & " as EndEmpresa_Id," & Me.Codigo & " as Pedido_Id) AS Ori" & vbCrLf &
                      "    ON Dest.Empresa_Id    = Ori.Empresa_Id" & vbCrLf &
                      "   and Dest.EndEmpresa_Id = Ori.EndEmpresa_Id" & vbCrLf &
                      "   and Dest.Pedido_Id     = Ori.Pedido_Id" & vbCrLf &
                      "  WHEN NOT MATCHED" & vbCrLf &
                      "    THEN Insert (Empresa_Id, EndEmpresa_Id," & vbCrLf &
                      "        Pedido_Id, UnidadeDeNegocio, Cliente, EndCliente, " & vbCrLf &
                      "        Praca, EndPraca, PedidoEfetivo, Safra, Moeda, TemVariacao, Indexador, Operacao, SubOperacao, Situacao, " & vbCrLf &
                      "        DataPedido, DataInicioEntrega, DataEntrega, PedidoOrigem, Solicitacao, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf &
                      "        UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf &
                      "        UsuarioLiberacao, UsuarioLiberacaoData, Observacoes, CondicaoPagamento, BancoCliente, " & vbCrLf &
                      "        AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, FreteCIFFOB, OrigemDestino, Comercializacao, " & vbCrLf &
                      "        Finalidade, Contrato, Carteira, Taxa, VencimentoPedido,LocalEmbarque,EndLocalEmbarque, MomentoFinanceiro, " & vbCrLf &
                      "        AgruparFinanceiro, IndiceFixado, EstadoEntrega, CidadeEntrega, FiscalAberto, FinanceiroAberto, " & vbCrLf &
                      "        EmpresaTroca, EndEmpresaTroca, PedidoTroca, ContaAdiantamentoTroca, CondicaoPagamentoEntrega, QuotaEntrega, PeriodicidadeEntrega, PedidoBloqueado," & vbCrLf &
                      "        LiberaCarregamento,VersaoPedido, VersaoUsuario, VersaoHorarioBloqueio, IndexadorFixo, Troca, Antecipada, Recompra, XPedNFe, ItemXPedNFe, InvoiceNavio, Embalagem, TipoCondicaoEntrega, TipoPagamentoPtax, LocalDeEmbarque)" & vbCrLf &
                      " VALUES ('" & Me.CodigoEmpresa & "', " & Me.EnderecoEmpresa.ToString() & "," & vbCrLf &
                      Me.Codigo & ",'" & Me.UnidadeNegocio.Codigo & "'," & vbCrLf &
                      "'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & "," & vbCrLf &
                      "'" & Me.CodigoPraca & "'," & Me.EnderecoPraca & "," & vbCrLf &
                      "'" & Me.PedidoEfetivo & "','" & Me.CodigoSafra & "'," & vbCrLf &
                      Me.CodigoMoeda & "," & IIf(Me.TemVariacao, 1, 0) & "," & Me.CodigoIndexador & "," & vbCrLf &
                      Me.CodigoOperacao & "," & Me.CodigoSubOperacao & "," & vbCrLf &
                      Me.CodigoSituacao & ",'" & Me.DataPedido.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "'" & Me.DataEntregaInicial.ToString("yyyy-MM-dd") & "','" & Me.DataEntregaFinal.ToString("yyyy-MM-dd") & "', '0', 0," & vbCrLf &
                      "'" & Me.UsuarioInclusao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf &
                      "'" & Me.UsuarioAlteracao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf &
                      "'" & Me.UsuarioCancelamento & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf &
                      "'" & Me.UsuarioLiberacao & "', CONVERT(varchar, CURRENT_TIMESTAMP, 112)," & vbCrLf &
                      "'" & Me.Observacoes & "'," & vbCrLf

                If Not Me.CondicaoPagamento Is Nothing Then Sql &= Me.CondicaoPagamento.Codigo & ", " Else Sql &= "0, "

                'BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente
                If Not Me.ContaBancariaSelecionada Is Nothing Then
                    Sql &= Me.ContaBancariaSelecionada.CodigoBanco & "," & vbCrLf &
                           "'" & Me.ContaBancariaSelecionada.CodigoAgencia & "','" & Me.ContaBancariaSelecionada.DigitoAgencia & "'," & vbCrLf &
                           "'" & Me.ContaBancariaSelecionada.ContaCorrente & "','" & Me.ContaBancariaSelecionada.DigitoConta & "'," & vbCrLf
                Else
                    Sql &= "0, '', '', '', '', "
                End If

                'FreteCIFFOB, OrigemDestino, Comercializacao
                Sql &= "'" & Me.FreteCIFFOB.ToString() & "','" & Me.OrigemDestino & "', '" & CInt(Me.Comercializacao) & "', "
                'Finalidade
                Sql &= Me.CodigoFinalidade

                Sql &= ",'" & Me.Contrato.ToString() & "','" & _CodigoCarteira & "'" & vbCrLf &
                       "," & Str(Me.Taxa) & ",'" & Me.DataVencimentoPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "," & IIf(Me.CodigoLocalEmbarque.Length > 0, "'" & Me.CodigoLocalEmbarque & "'", "NULL") & vbCrLf &
                       "," & IIf(Me.CodigoLocalEmbarque.Length > 0, Me.EndLocalEmbarque, "NULL") & "," & Me.MomentoFinanceiro & "," & vbCrLf &
                       IIf(Me.AgruparFinanceiro, 1, 0) & "," & Str(Me.IndiceFixado) & ",'" & Me.EstadoEntrega & "','" & Me.CidadeEntrega & "'," & vbCrLf &
                       IIf(Me.FiscalAberto, 1, 0) & "," & IIf(Me.FinanceiroAberto, 1, 0) & "," & vbCrLf &
                       "'" & Me.CodigoEmpresaTroca & "'," & vbCrLf &
                       Me.EnderecoEmpresaTroca & "," & vbCrLf &
                       Me.CodigoPedidoTroca & "," & vbCrLf &
                       "'" & Me.ContaAdiantamentoTroca & "'," & vbCrLf

                Sql &= Me.CodigoCondicaoPagamentoDaEntrega & ", "
                Sql &= Str(Me.QuotaEntrega) & "," & Me.PeriodicidadeEntrega & "," & IIf(_PedidoBloqueado, 1, 0)
                Sql &= "," & IIf(_LiberaCarregamento, 1, 0) & "," & IIf(Me.IUD = "I", 1, Me.VersaoPedido + 1) & ",'" & UsuarioServidor.NomeUsuario & "', NULL,"
                Sql &= IIf(Me.IndexadorFixo, 1, 0) & "," & IIf(Me.Troca, 1, 0) & "," & IIf(Me.Antecipada, 1, 0) & "," & IIf(Me.Recompra, 1, 0) & ",'" & Me.XPedNFe & "','" & Me.ItemXPedNFe & "'," & Me.InvoiceNavio & ", " & Me.Embalagem & ", '" & Me.TipoCondicaoEntrega & "', '" & Me.TipoPagamentoPtax & "', '" & Me.LocalDeEmbarque & "')"
                Sql &= " WHEN MATCHED " & vbCrLf &
                       "   THEN Update set " & vbCrLf &
                       "    Cliente                ='" & Me.CodigoCliente & "'" & vbCrLf &
                       "   ,EndCliente             = " & Me.EnderecoCliente & vbCrLf &
                       "   ,Praca                  ='" & Me.CodigoPraca & "'" & vbCrLf &
                       "   ,EndPraca               = " & Me.EnderecoPraca & vbCrLf &
                       "   ,PedidoEfetivo          ='" & Me.PedidoEfetivo & "'" & vbCrLf &
                       "   ,Safra                  ='" & Me.CodigoSafra & "'" & vbCrLf &
                       "   ,Moeda                  = " & Me.CodigoMoeda & vbCrLf &
                       "   ,Indexador              = " & Me.CodigoIndexador & vbCrLf &
                       "   ,Operacao               = " & Me.CodigoOperacao & vbCrLf &
                       "   ,SubOperacao            = " & Me.CodigoSubOperacao & vbCrLf &
                       "   ,Situacao               = " & Me.CodigoSituacao & vbCrLf &
                       "   ,DataPedido             ='" & Me.DataPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "   ,DataInicioEntrega      ='" & Me.DataEntregaInicial.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "   ,DataEntrega            ='" & Me.DataEntregaFinal.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "   ,CondicaoPagamento      ='" & Me.CodigoCondicaoPagamento & "'" & vbCrLf &
                       "   ,Contrato               ='" & Me.Contrato & "'" & vbCrLf &
                       "   ,Carteira               ='" & Me.CodigoCarteira & "'" & vbCrLf &
                       "   ,Taxa                   = " & Str(Me.Taxa) & vbCrLf &
                       "   ,VencimentoPedido       ='" & Me.DataVencimentoPedido.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                       "   ,LocalEmbarque          = " & IIf(Me.CodigoLocalEmbarque.Length > 0, "'" & Me.CodigoLocalEmbarque & "'", "NULL") & vbCrLf &
                       "   ,EndLocalEmbarque       = " & IIf(Me.CodigoLocalEmbarque.Length > 0, "'" & Me.EndLocalEmbarque & "'", "NULL") & vbCrLf &
                       "   ,MomentoFinanceiro      = " & Me.MomentoFinanceiro & vbCrLf &
                       "   ,AgruparFinanceiro      = " & IIf(Me.AgruparFinanceiro, 1, 0) & vbCrLf &
                       "   ,IndiceFixado           = " & Str(Me.IndiceFixado) & vbCrLf

                If Not Me.ContaBancariaSelecionada Is Nothing Then
                    Sql &= "   ,BancoCliente         = " & Me.ContaBancariaSelecionada.CodigoBanco & vbCrLf &
                           "   ,AgenciaCliente       ='" & Me.ContaBancariaSelecionada.CodigoAgencia & "'" & vbCrLf &
                           "   ,DigitoAgenciaCliente ='" & Me.ContaBancariaSelecionada.DigitoAgencia & "'" & vbCrLf &
                           "   ,ContaCliente         ='" & Me.ContaBancariaSelecionada.ContaCorrente & "'" & vbCrLf &
                           "   ,DigitoContaCliente   ='" & Me.ContaBancariaSelecionada.DigitoConta & "'" & vbCrLf
                Else
                    Sql &= "   ,BancoCliente         = 0" & vbCrLf &
                           "   ,AgenciaCliente       =''" & vbCrLf &
                           "   ,DigitoAgenciaCliente =''" & vbCrLf &
                           "   ,ContaCliente         =''" & vbCrLf &
                           "   ,DigitoContaCliente   =''" & vbCrLf
                End If

                Sql &= "   ,Observacoes          ='" & Me.Observacoes & "'" & vbCrLf &
                       "   ,FreteCIFFOB          ='" & Me.FreteCIFFOB.ToString() & "'" & vbCrLf &
                       "   ,OrigemDestino        ='" & Me.OrigemDestino & "'" & vbCrLf &
                       "   ,Comercializacao      ='" & CInt(Me.Comercializacao) & "'" & vbCrLf &
                       "   ,UsuarioAlteracao     ='" & Me.UsuarioAlteracao & "'" & vbCrLf &
                       "   ,UsuarioAlteracaoData = CURRENT_TIMESTAMP" & vbCrLf

                If Me.CodigoSituacao = eSituacao.Cancelado Then
                    Sql &= "   ,UsuarioCancelamento     = '" & UsuarioCancelamento & "'" & vbCrLf &
                           "   ,UsuarioCancelamentoData = CURRENT_TIMESTAMP" & vbCrLf
                End If

                Sql &= "   ,Finalidade = " & Me.CodigoFinalidade & vbCrLf &
                       "   ,EstadoEntrega            ='" & Me.EstadoEntrega & "'" & vbCrLf &
                       "   ,CidadeEntrega            ='" & Me.CidadeEntrega & "'" & vbCrLf &
                       "   ,FiscalAberto             = " & IIf(Me.FiscalAberto, 1, 0) & vbCrLf &
                       "   ,FinanceiroAberto         = " & IIf(Me.FinanceiroAberto, 1, 0) & vbCrLf &
                       "   ,EmpresaTroca             ='" & Me.CodigoEmpresaTroca & "'" & vbCrLf &
                       "   ,EndEmpresaTroca          = " & Me.EnderecoEmpresaTroca & vbCrLf &
                       "   ,PedidoTroca              = " & Me.CodigoPedidoTroca & vbCrLf &
                       "   ,CondicaoPagamentoEntrega = " & Me.CodigoCondicaoPagamentoDaEntrega & vbCrLf &
                       "   ,QuotaEntrega             = " & Str(Me.QuotaEntrega) & vbCrLf &
                       "   ,PeriodicidadeEntrega     = " & Me.PeriodicidadeEntrega & vbCrLf &
                       "   ,PedidoBloqueado          = " & IIf(Me.PedidoBloqueado, 1, 0) & vbCrLf &
                       "   ,LiberaCarregamento       = " & IIf(Me.LiberaCarregamento, 1, 0) & vbCrLf &
                       "   ,VersaoPedido             = " & (Me.VersaoPedido + 1) & vbCrLf &
                       "   ,VersaoUsuario            = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                       "   ,IndexadorFixo            = " & IIf(Me.IndexadorFixo, 1, 0) & vbCrLf &
                       "   ,Troca                    = " & IIf(Me.Troca, 1, 0) & vbCrLf &
                       "   ,Antecipada               = " & IIf(Me.Antecipada, 1, 0) & vbCrLf &
                       "   ,Recompra                 = " & IIf(Me.Recompra, 1, 0) & vbCrLf &
                       "   ,XPedNFe                  = '" & Me.XPedNFe & "'" & vbCrLf &
                       "   ,ItemXPedNFe              = '" & Me.ItemXPedNFe & "'" & vbCrLf &
                       "   ,InvoiceNavio             = " & Me.InvoiceNavio & "" & vbCrLf &
                       "   ,Embalagem                = " & Me.Embalagem & "" & vbCrLf &
                       "   ,TipoCondicaoEntrega      = '" & Me.TipoCondicaoEntrega & "'" & vbCrLf &
                       "   ,TipoPagamentoPtax        = '" & Me.TipoPagamentoPtax & "'" & vbCrLf &
                       "   ,LocalDeEmbarque          = '" & Me.LocalDeEmbarque & "';" & vbCrLf

                Sqls.Add(Sql)
                SalvarTabelasRelacionadasSql(Sqls)

            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                Sql = "DELETE Pedidos " &
                      " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" &
                      "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa &
                      "   AND Pedido_Id     = " & Me.Codigo
                Sqls.Add(Sql)
            Case "C"
                SalvarTabelasRelacionadasSql(Sqls)
                Sql = "UPDATE Pedidos SET" &
                      "   Situacao                = " & [Enum].Parse(GetType(eSituacao), "Cancelado") & vbCrLf &
                      "  ,UsuarioCancelamento     ='" & Me.UsuarioCancelamento & "'" & vbCrLf &
                      "  ,UsuarioCancelamentoData = CURRENT_TIMESTAMP " & vbCrLf &
                      " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                      "   AND Pedido_Id     = " & Me.Codigo & vbCrLf
                Sqls.Add(Sql)
            Case Else
                SalvarTabelasRelacionadasSql(Sqls)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Me.Itens.SalvarSql(Sqls)
        Me.Depositos.SalvarSql(Sqls)
        Me.Roteiros.SalvarSql(Sqls)
        Me.Transportadores.SalvarSql(Sqls)
        Me.Representantes.SalvarSql(Sqls)
        Me.Parcelas.SalvarSql(Sqls)
        Me.Contratos.SalvarSql(Sqls)

        Dim sql As String
        If Not Me.PedidoTroca Is Nothing AndAlso Me.CodigoPedidoTroca > 0 AndAlso Me.Troca AndAlso Me.Operacao.CodigoClasse = eClassesOperacoes.COMPRAS.ToString Then
            sql = "UPDATE Pedidos SET" &
                  "   EmpresaTroca    ='" & Me.CodigoEmpresa & "'" &
                  "  ,EndEmpresaTroca = " & Me.EnderecoEmpresa &
                  "  ,PedidoTroca     = " & Me.Codigo &
                  " WHERE Empresa_Id    ='" & Me.CodigoEmpresaTroca & "'" &
                  "   AND EndEmpresa_Id = " & Me.EnderecoEmpresaTroca &
                  "   AND Pedido_Id     = " & Me.CodigoPedidoTroca

            Sqls.Add(sql)
        End If

        If Not Me.TemNFG Then
            If ProjetoFinanceiroNovo Then
                Me.Titulos.SalvarSql(Sqls, False)
            Else
                'Grava condicoes pagamento
                If Me.IUD = "I" Or Me.IUD = "U" Then
                    'LIBERAR FEX PARA GERAR FINANCEIRO NA TROCA - FURLAN - 08-12-2015
                    If Me.Troca AndAlso Me.Vencimentos.Count > 0 AndAlso Not Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "15204808" Then
                        Sqls.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
                    ElseIf (Me.Moeda.Classificacao = eTiposMoeda.Oficial AndAlso Me.Itens.TotalOficial = 0) Or (Me.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso Me.Itens.TotalMoeda = 0) Then
                        Sqls.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
                    Else
                        Me.Vencimentos.ModificarHistorico(eTabelas.Pedido, New String() {Codigo.ToString()})
                        'Sqls.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
                        Sqls.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Inclusao, UsuarioServidor.NomeServidor.ToUpper(), _MomentoFinanceiro))
                    End If

                ElseIf Me.IUD = "D" Or Me.IUD = "C" Then
                    Sqls.AddRange(Me.Vencimentos.GetSQL(eModoAlteracao.Exclusao))
                End If
            End If
        End If


        If Me.IUD = "D" Or Me.IUD = "C" Then
            sql = "Delete AutorizacaoDeRetirada " &
                  " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "' " &
                  "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa.ToString() &
                  "   AND Pedido_Id     = " & Me.Codigo.ToString()

            Sqls.Add(sql)
        End If
    End Sub

    Public Function Bloquear() As Boolean
        If Me.IUD <> "I" AndAlso Not String.IsNullOrWhiteSpace(Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> Empresa.Empresa.Servidor Then
            Dim Banco As New AcessaBanco(2, Empresa.Empresa.Servidor)
            Dim sql As String = "UPDATE Pedidos SET VersaoPedido = ISNULL(VersaoPedido,0) + 1, VersaoUsuario = '" & UsuarioServidor.NomeUsuario & "', VersaoHorarioBloqueio = GETDATE() " & vbCrLf &
                                "WHERE Pedido_Id = " & Me.Codigo & " AND VersaoPedido = " & Me.VersaoPedido & " AND ISNULL(VersaoHorarioBloqueio, DATEADD(MINUTE, -3, GETDATE())) <= DATEADD(MINUTE, -3, GETDATE());"
            'SER RETORNO FOR MAIOR QUE ZERO LINHAS, PEDIDO FOI ALTERADO COM SUCESSO (BLOQUEADO)
            'Return Banco.ExecuteNonQuery(sql)
            Return True
        ElseIf Me.IUD <> "I" Then
            Dim Banco As New AcessaBanco()
            Dim sql As String = "UPDATE Pedidos SET VersaoPedido = ISNULL(VersaoPedido,0) + 1, VersaoUsuario = '" & UsuarioServidor.NomeUsuario & "', VersaoHorarioBloqueio = GETDATE() " & vbCrLf &
                                "WHERE Pedido_Id = " & Me.Codigo & " AND VersaoPedido = " & Me.VersaoPedido & " AND ISNULL(VersaoHorarioBloqueio, DATEADD(MINUTE, -3, GETDATE())) <= DATEADD(MINUTE, -3, GETDATE());"
            'SER RETORNO FOR MAIOR QUE ZERO LINHAS, PEDIDO FOI ALTERADO COM SUCESSO (BLOQUEADO)
            'Return Banco.ExecuteNonQuery(sql)
            Return True
        End If
        Return True
    End Function

    Public Function Liberar() As Boolean
        If Me.IUD <> "I" AndAlso Not String.IsNullOrWhiteSpace(Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> Empresa.Empresa.Servidor Then
            Dim Banco As New AcessaBanco(2, Empresa.Empresa.Servidor)
            Dim sql As String = "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL, VERSAOUSUARIO = NULL WHERE PEDIDO_ID = '" & Me.Codigo & "' AND EMPRESA_ID = '" & Me.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.EnderecoEmpresa & "' AND VERSAOUSUARIO = '" & UsuarioServidor.NomeUsuario & "'; "
            Return Banco.ExecuteNonQuery(sql)
        End If
        Return True
    End Function

    Public Sub getSqlException(ByRef Sqls As ArrayList)
        Dim sql As String = "BEGIN TRY " & vbCrLf &
               "DECLARE @HORA_BLOQUEIO AS DATETIME = DATEADD(MINUTE, 3, (SELECT VERSAOHORARIOBLOQUEIO FROM PEDIDOS WHERE PEDIDO_ID = '" & Me.Codigo & "' AND EMPRESA_ID = '" & Me.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.EnderecoEmpresa & "')) " & vbCrLf &
               "PRINT 'HORA_ATUAL: ' + CAST(GETDATE() AS VARCHAR); " & vbCrLf &
               "PRINT 'HORA_BLOQUEIO: ' + CAST(@HORA_BLOQUEIO AS VARCHAR); " & vbCrLf &
               "IF (GETDATE() > @HORA_BLOQUEIO) " & vbCrLf &
               "BEGIN " & vbCrLf &
               "RAISERROR ('POR FAVOR, ATUALIZE O SALDO FINANCEIRO PARA REALIZAR ESTA A��O!', " & vbCrLf &
               "16, " & vbCrLf &
               "1); " & vbCrLf &
               "END " & vbCrLf &
               "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL, VERSAOUSUARIO = NULL WHERE PEDIDO_ID = '" & Me.Codigo & "' AND EMPRESA_ID = '" & Me.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.EnderecoEmpresa & "'; " & vbCrLf &
               "END TRY " & vbCrLf &
               "BEGIN CATCH " & vbCrLf &
               "DECLARE @ErrorMessage NVARCHAR(4000); " & vbCrLf &
               "DECLARE @ErrorSeverity INT; " & vbCrLf &
               "DECLARE @ErrorState INT; " & vbCrLf &
               "SELECT " & vbCrLf &
               "@ErrorMessage = ERROR_MESSAGE(), " & vbCrLf &
               "@ErrorSeverity = ERROR_SEVERITY(), " & vbCrLf &
               "@ErrorState = ERROR_STATE(); " & vbCrLf &
               "RAISERROR (@ErrorMessage, " & vbCrLf &
               "@ErrorSeverity, " & vbCrLf &
               "@ErrorState); " & vbCrLf &
               "END CATCH;"
        Sqls.Add(sql)
    End Sub

    Public Sub ImprimirPedido(ByVal page As Page, ByVal pdf As Boolean, Optional ByRef nameFile As String = "", Optional ByVal EmailNFePedido As Boolean = False)

        Dim Banco As New AcessaBanco()
        Dim ds As New DataSet
        Dim Sql As String

        Sql = "SELECT Nome, Cliente_Id as Codigo, Inscricao, Endereco, Numero, complemento, Bairro, Telefone, CEP, Cidade, Estado, Email, '' AS Site " & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              " Where Cliente_id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   and Endereco_id = " & Me.EnderecoEmpresa
        ds.Merge(Banco.ConsultaDataSet(Sql, "Empresa"))

        Sql = "SELECT Nome, Fantasia, Cliente_Id as Codigo, Inscricao, Endereco, complemento, Bairro, Telefone, CEP, Cidade, Estado, Email " & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              " Where Cliente_id    ='" & Me.CodigoCliente & "'" & vbCrLf &
              "   and Endereco_id = " & Me.EnderecoCliente
        ds.Merge(Banco.ConsultaDataSet(Sql, "Cliente"))

        Sql = "SELECT CLI.Nome, CLI.Fantasia, CLI.Cliente_Id AS Codigo, CLI.Inscricao, CLI.Endereco, CLI.complemento, CLI.Bairro, CLI.Telefone, CLI.CEP, CLI.Cidade, CLI.Estado, CLI.Email " & vbCrLf &
              "  FROM Clientes CLI" & vbCrLf &
              "  INNER JOIN PedidosxDepositos PXD " & vbCrLf &
              "   ON PXD.Deposito_Id            = CLI.Cliente_Id " & vbCrLf &
              "   AND PXD.EndDeposito_Id        = CLI.Endereco_id  " & vbCrLf &
              " WHERE PXD.Empresa_Id            = '" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   AND PXD.EndEmpresa_Id         = " & Me.EnderecoEmpresa & "" & vbCrLf &
              "   AND PXD.Pedido_Id             = " & Me.Codigo & " " & vbCrLf &
              "   AND PXD.Tipo                  = 'OD'; " & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(Sql, "Entrega"))

        Sql = "SELECT Nome, Cliente_Id as Codigo, Inscricao, Endereco,complemento, Telefone, CEP, Cidade, Estado " & vbCrLf &
              "  FROM Clientes" & vbCrLf &
              " Where Cliente_id    ='" & Me.CodigoPraca & "'" & vbCrLf &
              "   and Endereco_id = " & Me.EnderecoPraca
        ds.Merge(Banco.ConsultaDataSet(Sql, "Praca"))

        Sql = "SELECT P.Pedido_Id as Codigo," & vbCrLf &
              "       P.PedidoEfetivo," & vbCrLf &
              "       P.DataPedido," & vbCrLf &
              "       P.Vendedor," & vbCrLf &
              "       case When M.Classificacao = 'O' then PxI.TotalBrutoOficial   else PxI.TotalBrutoMoeda   end as TotalBruto," & vbCrLf &
              "       case When M.Classificacao = 'O' then PxI.TotalLiquidoOficial else PxI.TotalLiquidoMoeda end as TotalLiquido," & vbCrLf &
              "       FORMAT(P.DataInicioEntrega, 'dd/MM/yyyy') + ' - ' + FORMAT(P.DataEntrega, 'dd/MM/yyyy') as PeriodoEntrega," & vbCrLf &
              "       P.Observacoes," & vbCrLf &
              "       P.BancoCliente," & vbCrLf &
              "       isnull(B.Descricao,'') as NomeBanco," & vbCrLf &
              "       P.AgenciaCliente," & vbCrLf &
              "       P.DigitoAgenciaCliente," & vbCrLf &
              "       P.ContaCliente," & vbCrLf &
              "       P.DigitoContaCliente," & vbCrLf &
              "       SO.Classe," & vbCrLf &
              "       cast(so.operacao_id as varchar) + '-' + cast(so.SubOperacoes_Id as varchar) + ' - ' + so.Descricao + ' - (' + so.Classe + ')' as Operacao," & vbCrLf &
              "       P.Moeda, " & vbCrLf &
              "       P.FreteCIFFOB, ISNULL(P.Embalagem, 0) AS Embalagem, ISNULL(EM.Descricao, '') AS DescricaoEmbalagem, " & vbCrLf &
              "       ISNULL(P.TipoCondicaoEntrega, '') AS TipoCondicaoEntrega, ISNULL(P.TipoPagamentoPtax, '') AS TipoPagamentoPtax, ISNULL(UPPER(P.LocalDeEmbarque), '') AS LocalDeEmbarque " & vbCrLf &
              "  FROM Pedidos P" & vbCrLf &
              " Inner Join (Select Empresa_id," & vbCrLf &
              "                    EndEmpresa_Id," & vbCrLf &
              "                    Pedido_Id," & vbCrLf &
              "                    sum(case When Encargo_Id = 'PRODUTO' Then valorOficial else 0 end) as TotalBrutoOficial," & vbCrLf &
              "                    sum(case When Encargo_Id = 'PRODUTO' Then valorMoeda   else 0 end) as TotalBrutoMoeda," & vbCrLf &
              "                    sum(case When Encargo_Id = 'LIQUIDO' Then valorOficial else 0 end) as TotalLiquidoOficial," & vbCrLf &
              "                    sum(case When Encargo_Id = 'LIQUIDO' Then valorMoeda   else 0 end) as TotalLiquidoMoeda" & vbCrLf &
              "               From pedidosxencargos" & vbCrLf &
              "              GRoup by Empresa_id, EndEmpresa_Id, Pedido_Id " & vbCrLf &
              "             ) PxI" & vbCrLf &
              "    on P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf &
              "   and P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
              "   and P.Pedido_id     = PxI.Pedido_Id" & vbCrLf &
              " LEFT JOIN Embalagens EM " & vbCrLf &
              "    ON P.Embalagem    = EM.Embalagem_Id" & vbCrLf &
              "  Left Join Bancos B" & vbCrLf &
              "    on P.BancoCliente = B.Banco_Id" & vbCrLf &
              " Inner Join SubOperacoes SO" & vbCrLf &
              "    on SO.Operacao_Id     = P.Operacao" & vbCrLf &
              "   and SO.SubOperacoes_Id = P.SubOperacao" & vbCrLf &
              " Inner join Moedas M" & vbCrLf &
              "    on M.Moeda_Id = P.Moeda" & vbCrLf &
              " Where P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   and P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   and P.Pedido_Id     = " & Me.Codigo
        ds.Merge(Banco.ConsultaDataSet(Sql, "Pedido"))

        Sql = "   SELECT		co.Representante_Id + '-' + CAST(co.EndRepresentante_Id AS varchar) + ' - ' + c.Nome AS Representante,                         " & vbCrLf &
            "   			CONVERT(DECIMAL(18, 10), co.Percentual) AS Percentual, co.ValorComissao, co.Principal,                                         " & vbCrLf &
            "   			ISNULL(co.PercentualFixo, CASE WHEN co.valorcomissao > 0 THEN 1 ELSE 0 END) AS PercentualFixo, co.Pedido_Id  " & vbCrLf &
            "   FROM         Comissoes AS co                                                                                                               " & vbCrLf &
            "   		INNER JOIN Clientes AS c                                                                                                           " & vbCrLf &
            "   			ON c.Cliente_Id = co.Representante_Id                                                                                                " & vbCrLf &
            "   			AND c.Endereco_Id = co.EndRepresentante_Id                                                                                           " & vbCrLf &
            "   WHERE   (co.Empresa_Id = '" & Me.CodigoEmpresa & "') " & vbCrLf &
            "           AND (co.EndEmpresa_Id = " & Me.EnderecoEmpresa & ")" & vbCrLf &
            "           AND (co.Pedido_Id = " & Me.Codigo & ")                                      " & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(Sql, "Representantes"))

        Sql = "SELECT co.Transportador_Id + '-' + cast(co.EndTransportador_Id as varchar) + ' - ' + c.Nome as Transportador, co.Quantidade, co.QuotaDiaria, co.Redespacho, co.DataFrete_Id, isnull(co.UnitarioFrete,0) AS UnitarioFrete " & vbCrLf &
                  "  FROM PedidosXTransportadores co" & vbCrLf &
                  "   	INNER JOIN Clientes AS c                                                                                                           " & vbCrLf &
                  "   		ON c.Cliente_Id = co.Transportador_Id                                                                                                " & vbCrLf &
                  "		    AND c.Endereco_Id = co.EndTransportador_Id                                                                                           " & vbCrLf &
                  "     WHERE   co.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "         AND co.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                  "         AND co.Pedido_Id     = " & Me.Codigo & vbCrLf
        ds.Merge(Banco.ConsultaDataSet(Sql, "Transportadores"))

        If Left(Me.CodigoEmpresa, 8) = "24450490" Then

            Sql = " WITH PE AS (" & vbCrLf &
                  "    SELECT PxI.Pedido_Id AS Pedido," & vbCrLf &
                  "           PxI.Produto_Id AS CodigoProduto," & vbCrLf &
                  "           Prd.Unidade," & vbCrLf &
                  "           PEI.UnidadeComercializacao," & vbCrLf &
                  "           PEI.FatorConversao," & vbCrLf

            If Me.Empresa.Empresa.UsarRegistroMinAgr Then
                Sql &= "           Prd.Nome + '-' + Prd.Descricao + '(' + Prd.RegMinAgr + ')' AS DescricaoProduto," & vbCrLf
            ElseIf Me.Empresa.Empresa.UsarDescricaoProduto Then
                Sql &= "           Prd.Nome + '-' + Prd.Descricao AS DescricaoProduto," & vbCrLf
            Else
                Sql &= "           Prd.Nome AS DescricaoProduto," & vbCrLf
            End If

            Sql &= "           ROW_NUMBER() OVER (PARTITION BY PxI.Pedido_Id ORDER BY PxI.Produto_Id) AS rn" & vbCrLf &
                  "    FROM Pedidos P" & vbCrLf &
                  "    INNER JOIN PedidoXItemXLancamento PxI ON P.Empresa_Id = PxI.Empresa_Id" & vbCrLf &
                  "        AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
                  "        AND P.Pedido_Id = PxI.Pedido_Id" & vbCrLf &
                  "    INNER JOIN PedidoXItem PEI ON P.Empresa_Id = PEI.Empresa_Id" & vbCrLf &
                  "        AND P.EndEmpresa_Id = PEI.EndEmpresa_Id" & vbCrLf &
                  "        AND P.Pedido_Id = PEI.Pedido_Id" & vbCrLf &
                  "        AND PEI.Produto_id = PxI.Produto_id" & vbCrLf &
                  "    INNER JOIN Produtos Prd ON PxI.Produto_id = Prd.Produto_id" & vbCrLf &
                  "    INNER JOIN Moedas M ON M.Moeda_Id = P.Moeda" & vbCrLf &
                  "    WHERE P.Empresa_Id = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "        AND P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                  "        AND P.Pedido_Id = " & Me.Codigo & vbCrLf &
                  ")" & vbCrLf &
                  "SELECT " & vbCrLf &
                  "    (SELECT CodigoProduto FROM PE WHERE rn = 1) AS CodigoProduto," & vbCrLf &
                  "    (SELECT DescricaoProduto FROM PE WHERE rn = 1) AS DescricaoProduto," & vbCrLf &
                  "    (SELECT Unidade FROM PE WHERE rn = 1) AS Unidade," & vbCrLf &
                  "    (SELECT UnidadeComercializacao FROM PE WHERE rn = 1) AS UnidadeComercializacao," & vbCrLf &
                  "    (SELECT FatorConversao FROM PE WHERE rn = 1) AS FatorConversao," & vbCrLf &
                  "    PxI.Pedido_Id AS Pedido," & vbCrLf &
                  "    SUM(CASE WHEN PxI.TipodeLancamento = 'E' THEN PxI.Quantidade * -1 ELSE PxI.Quantidade END) AS Quantidade," & vbCrLf &
                  "    SUM(CASE WHEN PxI.TipodeLancamento = 'E' THEN PxI.QuantidadeComercializacao * -1 ELSE PxI.QuantidadeComercializacao END) AS QuantidadeComercializacao," & vbCrLf &
                  "    SUM(CASE" & vbCrLf &
                  "        WHEN PxI.TipodeLancamento = 'E' THEN" & vbCrLf &
                  "            CASE WHEN M.classificacao = 'O' THEN TotalOficial * -1 ELSE TotalMoeda * -1 END" & vbCrLf &
                  "        ELSE" & vbCrLf &
                  "            CASE WHEN M.classificacao = 'O' THEN TotalOficial ELSE TotalMoeda END" & vbCrLf &
                  "    END) AS Total" & vbCrLf &
                  "INTO #Temp" & vbCrLf &
                  "FROM Pedidos P" & vbCrLf &
                  "INNER JOIN PedidoXItemXLancamento PxI ON P.Empresa_Id = PxI.Empresa_Id" & vbCrLf &
                  "    AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
                  "    AND P.Pedido_Id = PxI.Pedido_Id" & vbCrLf &
                  "INNER JOIN PedidoXItem PEI ON P.Empresa_Id = PEI.Empresa_Id" & vbCrLf &
                  "    AND P.EndEmpresa_Id = PEI.EndEmpresa_Id" & vbCrLf &
                  "    AND P.Pedido_Id = PEI.Pedido_Id" & vbCrLf &
                  "    AND PEI.Produto_id = PxI.Produto_id" & vbCrLf &
                  "INNER JOIN Produtos Prd ON PxI.Produto_id = Prd.Produto_id" & vbCrLf &
                  "INNER JOIN Moedas M ON M.Moeda_Id = P.Moeda" & vbCrLf &
                  "WHERE P.Empresa_Id = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "    AND P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                  "    AND P.Pedido_Id = " & Me.Codigo & vbCrLf &
                  "GROUP BY PxI.Pedido_Id;" & vbCrLf &
                  "SELECT Pedido, CodigoProduto, DescricaoProduto, Unidade, UnidadeComercializacao, FatorConversao, Quantidade, QuantidadeComercializacao," & vbCrLf &
                  "       CASE WHEN QuantidadeComercializacao > 0 THEN Total / QuantidadeComercializacao ELSE 0 END AS Unitario," & vbCrLf &
                  "       Total" & vbCrLf &
                  "FROM #Temp" & vbCrLf &
                  "WHERE QuantidadeComercializacao > 0;" & vbCrLf

        Else

            Sql = "SELECT PxI.Pedido_Id as Pedido," & vbCrLf &
             "       sum(case" & vbCrLf &
             "             When PxI.TipodeLancamento = 'E'" & vbCrLf &
             "               then PxI.Quantidade * -1" & vbCrLf &
             "               else PxI.Quantidade" & vbCrLf &
             "           end) as Quantidade," & vbCrLf &
             "       sum(case" & vbCrLf &
             "             When PxI.TipodeLancamento = 'E'" & vbCrLf &
             "               then PxI.QuantidadeComercializacao * -1" & vbCrLf &
             "               else PxI.QuantidadeComercializacao" & vbCrLf &
             "           end) as QuantidadeComercializacao," & vbCrLf &
             "       Prd.Unidade," & vbCrLf &
             "       PEI.UnidadeComercializacao," & vbCrLf &
             "       PEI.FatorConversao," & vbCrLf

            If Me.Empresa.Empresa.UsarRegistroMinAgr Then
                Sql &= "       Prd.Nome + '-' + Prd.Descricao + '(' + Prd.RegMinAgr + ')' as DescricaoProduto," & vbCrLf
            ElseIf Me.Empresa.Empresa.UsarDescricaoProduto Then
                Sql &= "       Prd.Nome + '-' + Prd.Descricao as DescricaoProduto," & vbCrLf
            Else
                Sql &= "       Prd.Nome as DescricaoProduto," & vbCrLf
            End If

            Sql &= "       sum(case" & vbCrLf &
                  "            when PxI.TipodeLancamento = 'E'" & vbCrLf &
                  "             then" & vbCrLf &
                  "               case" & vbCrLf &
                  "                 When M.classificacao = 'O'" & vbCrLf &
                  "                   then TotalOficial * -1" & vbCrLf &
                  "                   else TotalMoeda * - 1" & vbCrLf &
                  "               end" & vbCrLf &
                  "             else" & vbCrLf &
                  "               case" & vbCrLf &
                  "                 When M.classificacao = 'O'" & vbCrLf &
                  "                   then TotalOficial " & vbCrLf &
                  "                   else TotalMoeda " & vbCrLf &
                  "               end           end" & vbCrLf &
                  "          ) as Total," & vbCrLf &
                  "       PxI.Produto_Id as CodigoProduto" & vbCrLf &
                  "  INTO #Temp " & vbCrLf &
                  "  FROM Pedidos P" & vbCrLf &
                  " INNER JOIN PedidoXItemXLancamento PxI" & vbCrLf &
                  "    ON P.Empresa_Id    = PxI.Empresa_Id" & vbCrLf &
                  "   AND P.EndEmpresa_Id = PxI.EndEmpresa_Id" & vbCrLf &
                  "   AND P.Pedido_Id     = PxI.Pedido_Id" & vbCrLf &
                  " INNER JOIN PedidoXItem PEI " & vbCrLf &
                  "    ON P.Empresa_Id    = PEI.Empresa_Id" & vbCrLf &
                  "   AND P.EndEmpresa_Id = PEI.EndEmpresa_Id" & vbCrLf &
                  "   AND P.Pedido_Id     = PEI.Pedido_Id" & vbCrLf &
                  "   AND PEI.Produto_id = PxI.Produto_id" & vbCrLf &
                  " Inner Join Produtos Prd" & vbCrLf &
                  "    on PxI.Produto_id = Prd.Produto_id" & vbCrLf &
                  " Inner Join Moedas M" & vbCrLf &
                  "    on M.Moeda_Id = P.Moeda" & vbCrLf &
                  " Where P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "   and P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
                  "   and P.Pedido_Id     = " & Me.Codigo & vbCrLf &
                  " group by PxI.Pedido_Id," & vbCrLf &
                  "          Prd.Unidade," & vbCrLf

            If Me.Empresa.Empresa.UsarRegistroMinAgr Then
                Sql &= "       Prd.Nome, Prd.Descricao, Prd.RegMinAgr," & vbCrLf
            ElseIf Me.Empresa.Empresa.UsarDescricaoProduto Then
                Sql &= "       Prd.Nome, Prd.Descricao," & vbCrLf
            Else
                Sql &= "       Prd.Nome," & vbCrLf
            End If

            Sql &= "		  PxI.Produto_Id, PEI.UnidadeComercializacao, PEI.FatorConversao; " & vbCrLf &
                  " --Composicao do Unitario " & vbCrLf &
                  "SELECT Pedido, CodigoProduto,  DescricaoProduto, Unidade, UnidadeComercializacao, FatorConversao, Quantidade, QuantidadeComercializacao, " & vbCrLf &
                  "       CASE " & vbCrLf &
                  "           WHEN QuantidadeComercializacao>0 " & vbCrLf &
                  "			       THEN Total / QuantidadeComercializacao " & vbCrLf &
                  "				   ELSE 0 " & vbCrLf &
                  "		  END Unitario,  " & vbCrLf &
                  "       Total" & vbCrLf &
                  "  FROM #TEMP" & vbCrLf

        End If

        ds.Merge(Banco.ConsultaDataSet(Sql, "PedidosXItens"))

        Sql = "SELECT P.Pedido_id as Pedido," & vbCrLf &
              "       PxE.Encargo_Id as Encargo," & vbCrLf &
              "       sum(case" & vbCrLf &
              "             when M.Classificacao = 'O'" & vbCrLf &
              "               Then PxE.ValorOficial" & vbCrLf &
              "               Else PxE.ValorMoeda" & vbCrLf &
              "           end" & vbCrLf &
              "           ) as Total" & vbCrLf &
              "  FROM Pedidos P" & vbCrLf &
              " INNER JOIN PedidosXEncargos PxE" & vbCrLf &
              "    ON P.Empresa_Id    = PxE.Empresa_Id" & vbCrLf &
              "   AND P.EndEmpresa_Id = PxE.EndEmpresa_Id" & vbCrLf &
              "   AND P.Pedido_Id     = PxE.Pedido_Id" & vbCrLf &
              " INNER JOIN Moedas M" & vbCrLf &
              "    ON P.Moeda = M.Moeda_Id" & vbCrLf &
              " Where P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   and P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   and P.Pedido_Id     = " & Me.Codigo & vbCrLf &
              " Group By P.Pedido_Id, PxE.Encargo_Id" & vbCrLf &
              " having sum(case when M.Classificacao = 'O' Then PxE.ValorOficial Else PxE.ValorMoeda end) > 0" & vbCrLf &
              " order by Case" & vbCrLf &
              "            when PxE.Encargo_Id = 'PRODUTO' THEN 1" & vbCrLf &
              "            When PxE.Encargo_Id = 'LIQUIDO' THEN 3" & vbCrLf &
              "            else 2" & vbCrLf &
              "          end " & vbCrLf

        ds.Merge(Banco.ConsultaDataSet(Sql, "PedidosXEncargos"))

        If Left(Me.CodigoEmpresa, 8) = "24450490" Then

            Sql = "SELECT * FROM ( Select P.Pedido_id as Pedido, '' AS Parcela, " & vbCrLf &
              "       FORMAT(Sb.Prorrogacao, 'dd/MM/yyyy')  AS Vencimento," & vbCrLf &
              "       case" & vbCrLf &
              "          When M.Classificacao = 'O'" & vbCrLf &
              "            then sb.ValorLiquido" & vbCrLf &
              "            else sb.MoedaValorLiquido" & vbCrLf &
              "       end Valor" & vbCrLf &
              "  From Pedidos P" & vbCrLf &
              " Inner Join" & vbCrLf &
              "      (Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf &
              "	   from ContasaReceber cr" & vbCrLf &
              "			inner join ComprasXProdutos cXp" & vbCrLf &
              "					on cXp.Produto_Id = cr.Carteira " & vbCrLf &
              "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf &
              "          union all" & vbCrLf &
              "       Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf &
              "	   from ContasaPagar cr" & vbCrLf &
              "			inner join ComprasXProdutos cXp" & vbCrLf &
              "					on cXp.Produto_Id = cr.Carteira " & vbCrLf &
              "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf &
              "       ) sb" & vbCrLf &
              "    on sb.Empresa    = P.Empresa_id" & vbCrLf &
              "   and Sb.EndEmpresa = P.EndEmpresa_id" & vbCrLf &
              "   and sb.Pedido     = P.Pedido_id" & vbCrLf &
              " Inner join Moedas M" & vbCrLf &
              "    on M.Moeda_Id = P.Moeda" & vbCrLf &
              " INNER JOIN Pagamentos Pag" & vbCrLf &
              "     On P.CondicaoPagamento = Pag.Pagamento_Id" & vbCrLf &
              " Where P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   and P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   and P.Pedido_Id     = " & Me.Codigo & vbCrLf &
              "   and NOT Pag.Descricao LIKE '%DDF%'" & vbCrLf &
              " UNION " & vbCrLf &
              " SELECT P.Pedido_id as Pedido, " & vbCrLf &
              "       '' AS Parcela, " & vbCrLf &
              "       FORMAT(P.DataPedido, 'dd/MM/yyyy') AS Vencimento, " & vbCrLf &
              "       0 Valor " & vbCrLf &
              " FROM Pedidos P " & vbCrLf &
              " INNER JOIN Pagamentos Pag " & vbCrLf &
              "    ON P.CondicaoPagamento = Pag.Pagamento_Id " & vbCrLf &
              " WHERE P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   AND P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   AND P.Pedido_Id     = " & Me.Codigo & vbCrLf &
              "    AND Pag.Descricao LIKE '%DDF%' " & vbCrLf &
              ") AS PE " & vbCrLf &
              "ORDER BY CONVERT(DATE, PE.Vencimento, 103) "

        Else

            Sql = "Select P.Pedido_id as Pedido, '' AS Parcela, " & vbCrLf &
              "       FORMAT(Sb.Prorrogacao, 'dd/MM/yyyy')  AS Vencimento," & vbCrLf &
              "       case" & vbCrLf &
              "          When M.Classificacao = 'O'" & vbCrLf &
              "            then sb.ValorLiquido" & vbCrLf &
              "            else sb.MoedaValorLiquido" & vbCrLf &
              "       end Valor" & vbCrLf &
              "  From Pedidos P" & vbCrLf &
              " Inner Join" & vbCrLf &
              "      (Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf &
              "	   from ContasaReceber cr" & vbCrLf &
              "			inner join ComprasXProdutos cXp" & vbCrLf &
              "					on cXp.Produto_Id = cr.Carteira " & vbCrLf &
              "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf &
              "          union all" & vbCrLf &
              "       Select cr.Empresa, cr.EndEmpresa, cr.Registro_id, cr.Provisao, cr.Pedido, cr.Prorrogacao, cr.ValorLiquido, cr.MoedaValorLiquido" & vbCrLf &
              "	   from ContasaPagar cr" & vbCrLf &
              "			inner join ComprasXProdutos cXp" & vbCrLf &
              "					on cXp.Produto_Id = cr.Carteira " & vbCrLf &
              "	   WHERE cr.Situacao  = 1 and not (cXp.Adiantamento = 'S' and cXp.BaixaAdiantamento = 0)  " & vbCrLf &
              "       ) sb" & vbCrLf &
              "    on sb.Empresa    = P.Empresa_id" & vbCrLf &
              "   and Sb.EndEmpresa = P.EndEmpresa_id" & vbCrLf &
              "   and sb.Pedido     = P.Pedido_id" & vbCrLf &
              " Inner join Moedas M" & vbCrLf &
              "    on M.Moeda_Id = P.Moeda" & vbCrLf &
              " Where P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   and P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   and P.Pedido_Id     = " & Me.Codigo & vbCrLf &
              " ORDER BY Sb.Provisao, Vencimento;" & vbCrLf

        End If

        ds.Merge(Banco.ConsultaDataSet(Sql, "PedidoXParcelas"))

        If Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then

            Sql = " SELECT PE.BancoCliente Banco_id,                                " & vbCrLf &
                  "         BA.Descricao,                                           " & vbCrLf &
                  "         PE.AGenciaCliente AS Agencia_id,                        " & vbCrLf &
                  "         PE.DigitoAGenciaCliente AS DigitoAgencia_id,            " & vbCrLf &
                  "         PE.ContaCliente AS ContaCorrente_id,                    " & vbCrLf &
                  "         PE.DigitoContaCliente AS DigitoConta_id,                " & vbCrLf &
                  "         CASE                                                    " & vbCrLf &
                  "             WHEN LEN(PE.Cliente) = 14 THEN                      " & vbCrLf &
                  "                  'J'                                            " & vbCrLf &
                  "         ELSE                                                    " & vbCrLf &
                  "                       'F'                                       " & vbCrLf &
                  "         END AS TipoConta                                        " & vbCrLf &
                  "  FROM Pedidos  PE                                               " & vbCrLf &
                  "  INNER JOIN Bancos BA                                           " & vbCrLf &
                  "     ON PE.BancoCliente = BA.Banco_id                            " & vbCrLf &
                  "  WHERE PE.Empresa_Id                                            = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "   AND PE.EndEmpresa_Id                                          = " & Me.EnderecoEmpresa & "" & vbCrLf &
                  "   AND PE.Pedido_Id                                              = " & Me.Codigo

        Else

            Sql = " SELECT CXC.Banco_id,                        " & vbCrLf &
                  "         BA.Descricao,                       " & vbCrLf &
                  "         CXC.Agencia_id,                     " & vbCrLf &
                  "         CXC.DigitoAgencia_id,               " & vbCrLf &
                  "         CXC.ContaCorrente_id,               " & vbCrLf &
                  "         CXC.DigitoConta_id,                 " & vbCrLf &
                  "         CXC.TipoConta                       " & vbCrLf &
                  "  FROM ClientesXContasBancarias  CXC         " & vbCrLf &
                  "  INNER JOIN Bancos BA                       " & vbCrLf &
                  "     ON CXC.Banco_id = BA.Banco_id           " & vbCrLf &
                  "  WHERE CXC.Cliente_id           = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "   AND CXC.Endereco_id           = " & Me.EnderecoEmpresa

        End If

        ds.Merge(Banco.ConsultaDataSet(Sql, "CondicoesDeEntrega"))


        'Nova versao - reload branch Gilberto
        'Imagem
        Dim dtImagem As DataTable = ds.Tables.Add("Images")
        dtImagem.Columns.Add("path", GetType(String))
        dtImagem.Columns.Add("image", GetType(System.Byte()))

        Dim drImagem As DataRow = dtImagem.NewRow()
        Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & Me.Empresa.Imagem)

        drImagem("path") = strCaminhoImagem
        drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        dtImagem.Rows.Add(drImagem)

        Dim param As New Dictionary(Of String, Object)
        param.Add("Rep", IIf(ds.Tables("Representantes").Rows.Count > 0, True, False))
        param.Add("Transp", IIf(ds.Tables("Transportadores").Rows.Count > 0, True, False))

        param.Add("UsuarioInclusao", "Usu�rio Inclusao: " & IIf(Me.UsuarioAlteracao.Length > 0, Me.UsuarioAlteracao, Me.UsuarioInclusao))

        param.Add("CliouFor", IIf(Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "Fornecedor:", "Cliente:"))


        If Left(Me.Empresa.Codigo, 8) = "24450490" Then
            AddParametros(ds)
        Else

            If Not Me.CondicaoPagamento Is Nothing AndAlso Me.CondicaoPagamento.Codigo > 0 Then
                param.Add("CPagamento", Me.CondicaoPagamento.Descricao & " " & Me.CondicaoPagamento.Parcelas & " PARCELA(S)")
            Else
                param.Add("CPagamento", "")
            End If

        End If

        Dim totalKGS As Decimal = 0

        For Each row As DataRow In ds.Tables("PedidosXItens").Rows

            totalKGS += row("Quantidade") * row("FatorConversao")

        Next

        If totalKGS > 0 Then
            param.Add("TotalKGS", "TOTAL: " & totalKGS.ToString("N0") & " KGS")
        Else
            param.Add("TotalKGS", "")
        End If

        For Each row As DataRow In ds.Tables("Cliente").Rows
            param.Add("Email", row("Email"))
        Next

        For Each row As DataRow In ds.Tables("Pedido").Rows
            param.Add("PedidoEfeito", row("PedidoEfetivo"))
        Next

        Dim crptRelatorio As New ReportDocument()

        Try
            Dim strCaminho As String = "~/Reports/Cr_Pedidos.rpt"
            If Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "05272759" Then
                strCaminho = "~/Reports/Cr_PedidosQuimica.rpt"
            ElseIf Left(HttpContext.Current.Session("ssEmpresa").ToString, 8) = "24450490" Then
                strCaminho = "~/Reports/Cr_PedidosCN.rpt"
                For Each row As DataRow In ds.Tables("Empresa").Rows
                    row("Site") = "www.rtgraos.com"
                    Exit For
                Next
            End If

            strCaminho = HttpContext.Current.Server.MapPath(strCaminho)
            crptRelatorio.Load(strCaminho)

            Dim strNomeArquivo As String = String.Empty
            If pdf Then
                strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Else
                strNomeArquivo = "Files/" & Funcoes.GeraNomeArquivo & ".XLS"
            End If

            Dim strArquivo As String = HttpContext.Current.Server.MapPath(strNomeArquivo)

            crptRelatorio.SetDataSource(ds)
            Funcoes.BindParameters(crptRelatorio, param)

            If Dir(strArquivo).Length > 0 Then Kill(strArquivo)

            If pdf Then
                crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.PortableDocFormat, strArquivo)
            Else
                crptRelatorio.ExportToDisk(CrystalDecisions.Shared.ExportFormatType.Excel, strArquivo)
            End If

            If IO.File.Exists(strArquivo) Then
                If EmailNFePedido = True Then
                    nameFile = strArquivo
                Else
                    If pdf Then
                        Funcoes.AbrirArquivo(page, strNomeArquivo)
                    Else
                        Funcoes.AbrirExcel(page, strNomeArquivo)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(page, eTitulo.Erro)
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Private Sub AddParametros(ds As DataSet)

        Dim dParcelas As Decimal
        Dim dValorFinanceiro As Decimal
        Dim dsParametro As New DataSet("Parametros")
        Dim sDataVencimento As String = ""
        Dim bDDF As Boolean = False

        ' Criação do DataTable
        Dim dt As New DataTable("Parametros")

        ' Adição de colunas ao DataTable
        dt.Columns.Add("Financeiro", GetType(String))
        dt.Columns.Add("UsuarioInclusao", GetType(String))
        dt.Columns.Add("EmpresaRecebedora", GetType(String))

        ' Adição do DataTable ao DataSet
        dsParametro.Tables.Add(dt)

        If Not Me.CondicaoPagamento Is Nothing AndAlso Me.CondicaoPagamento.Codigo > 0 Then

            Dim iParcelasTotal As Integer = Me.CondicaoPagamento.Parcelas

            If Not ds.Tables("PedidoXParcelas") Is Nothing Then
                For Each row As DataRow In ds.Tables("PedidoXParcelas").Rows
                    dParcelas += 1

                    If dParcelas > iParcelasTotal Then
                        Exit For
                    End If

                    If Me.CondicaoPagamento.Descricao.Contains("DDF") Then

                        For Each itens As DataRow In ds.Tables("PedidosXItens").Rows
                            dValorFinanceiro += itens("Total")
                        Next

                        row("Parcela") = String.Format("{0} - {1} {2}", Me.CondicaoPagamento.Descricao, IIf(Me.Moeda.Codigo = 1, "R$", "US$"), FormatNumber(dValorFinanceiro, 2), dValorFinanceiro)
                        bDDF = True

                    Else
                        dValorFinanceiro += row("Valor")
                        row("Parcela") = String.Format("{0} - PARCELA {1} DE {2} - {3} {4}", row("Vencimento"), dParcelas, iParcelasTotal, IIf(Me.Moeda.Codigo = 1, "R$", "US$"), FormatNumber(row("Valor"), 2), row("Valor"))
                    End If


                    If sDataVencimento.Length = 0 And Not Me.CondicaoPagamento.Descricao.Contains("DDF") Then
                        sDataVencimento = row("Vencimento")
                    End If

                Next
            End If

            ' Criação de uma nova linha
            Dim newRow As DataRow = dt.NewRow()

            If sDataVencimento.Length > 0 Then
                ' Atribuição de valores às colunas
                newRow("Financeiro") = String.Format("{0} - {1}{2}", Me.CondicaoPagamento.Descricao, IIf(Me.Moeda.Codigo = 1, "R$", "US$"), FormatNumber(dValorFinanceiro, 2))
            Else
                ' Atribuição de valores às colunas
                newRow("Financeiro") = String.Format("{0}", Me.CondicaoPagamento.Descricao)
            End If

            If bDDF Then
                newRow("Financeiro") = ""
            End If

            newRow("UsuarioInclusao") = "Usuário Inclusão: " & IIf(Me.UsuarioAlteracao.Length > 0, Me.UsuarioAlteracao, Me.UsuarioInclusao)

            Dim objBXC As New [Lib].Negocio.BancosXContas(Me.Empresa.ContasBancarias(0).CodigoBanco, Me.Empresa.ContasBancarias(0).CodigoAgencia, Me.Empresa.ContasBancarias(0).DigitoAgencia, Me.Empresa.ContasBancarias(0).ContaCorrente, Me.Empresa.ContasBancarias(0).DigitoConta)
            'Adiciona Empresa Recebedora
            newRow("EmpresaRecebedora") = Funcoes.FormatarCpfCnpj(objBXC.CodigoEmpresa) & " (PIX)"

            ' Adição da linha ao DataTable
            dt.Rows.Add(newRow)

        Else

            ' Criação de uma nova linha
            Dim newRow As DataRow = dt.NewRow()
            newRow("Financeiro") = ""
            newRow("UsuarioInclusao") = ""
            newRow("EmpresaRecebedora") = ""

            ' Adição da linha ao DataTable
            dt.Rows.Add(newRow)

        End If

        ds.Merge(dsParametro)

    End Sub

    Public Function DataFixa(ByRef DataProrrogacao As DateTime) As Boolean

        Dim dsDataFixa As New DataSet
        Dim SqlPag As String = ""

        Dim objBanco As New AcessaBanco()

        SqlPag = "SELECT Descricao" & vbCrLf &
                 "FROM  Pagamentos" & vbCrLf &
                 "WHERE Pagamento_Id = " & Me.CondicaoPagamento.Codigo

        dsDataFixa = objBanco.ConsultaDataSet(SqlPag, "DataFixa")

        If Not Me.Vencimentos Is Nothing AndAlso Me.Vencimentos.Count > 0 AndAlso Not dsDataFixa Is Nothing AndAlso dsDataFixa.Tables(0).Rows.Count > 0 Then

            For Each row As DataRow In dsDataFixa.Tables(0).Rows
                If row("Descricao").ToString().Contains("DATA FIXA") OrElse Me.CondicaoPagamento.Codigo = 23 OrElse Me.CondicaoPagamento.VencimentoPedido Then
                    DataProrrogacao = Me.Vencimentos(0).DataProrrogacao
                    Return True
                End If
            Next

        End If

        Return False
    End Function

#End Region

End Class