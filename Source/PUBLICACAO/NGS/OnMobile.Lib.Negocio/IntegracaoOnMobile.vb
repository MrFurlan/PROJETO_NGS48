Imports System.Configuration
Imports System.Globalization
Imports System.IO
Imports System.Web
Imports NGS.Lib.Negocio

Public Class IntegracaoOnMobile
    Public Async Function IntegracaoWsVendedor() As Task
        Try
            Dim _http As New HttpClientOnMobile

            Dim _vendedores As New ListWsVendedor()

            Await _http.DeleteAsync("WsVendedor")

            Await _http.PostAsync("WsVendedor", _vendedores)
        Catch ex As Exception
            Dim erro = ex.Message
        End Try

    End Function

    'WsCliente
    Public Async Function IntegracaoWsCliente() As Task
        Try
            Dim _http As New HttpClientOnMobile

            Dim _clientes As New ListWsCliente()

            Await _http.DeleteAsync("WsCliente")

            Await _http.PostAsync("WsCliente", _clientes)
        Catch ex As Exception
            Dim erro = ex.Message
        End Try
    End Function
    'WsCondicao
    Public Async Function IntegracaoWsCondicao() As Task
        Dim _http As New HttpClientOnMobile

        Dim _condicoes As New ListWsCondicao()

        Await _http.DeleteAsync("WsCondicao")

        Await _http.PostAsync("WsCondicao", _condicoes)
    End Function
    'WsProduto
    Public Async Function IntegracaoWsProduto() As Task
        Dim _http As New HttpClientOnMobile

        Dim _produtos As New ListWsProduto()

        Await _http.DeleteAsync("WsProduto")

        Await _http.PostAsync("WsProduto", _produtos)
    End Function
    'WsTabcondicao
    Public Async Function IntegracaoWsTabCondicao() As Task
        Dim _http As New HttpClientOnMobile

        Dim _tabcondicoes As New ListWsTabCondicao()

        Await _http.DeleteAsync("WsTabcondicao")

        Await _http.PostAsync("WsTabcondicao", _tabcondicoes)
    End Function
    'WsTabpreco
    Public Async Function IntegracaoWsTabPreco() As Task
        Dim _http As New HttpClientOnMobile

        Dim _precos As New ListWsTabPreco()

        Await _http.DeleteAsync("WsTabpreco")

        Await _http.PostAsync("WsTabpreco", _precos)
    End Function
    'WsTabprecoProduto
    Public Async Function IntegracaoWsTabPrecoProduto() As Task
        Dim _http As New HttpClientOnMobile

        Dim _tabPrecoProdutos As New ListWsTabPrecoProduto()

        Await _http.DeleteAsync("WsTabprecoProduto")

        Await _http.PostAsync("WsTabprecoProduto", _tabPrecoProdutos)
    End Function
    'WsTipooperacao
    Public Async Function IntegracaoWsTipoOperacao() As Task
        Dim _http As New HttpClientOnMobile

        Dim _tipoOperacoes As New ListWsTipoOperacao()

        Await _http.DeleteAsync("WsTipoOperacao")

        Await _http.PostAsync("WsTipoOperacao", _tipoOperacoes)
    End Function

    Public Async Function IntegracaoWsLoad() As Task
        Dim _http As New HttpClientOnMobile

        Await _http.GetWsLoadAsync()
    End Function

    Public Async Function IntegracaoPedidoOnSolft() As Task
        Dim _http As New HttpClientOnMobile
        Dim pedidosParaIntegracao As New List(Of WsPedidoIntegracao)
        Dim pedidosOnSoft = Await _http.GetPedidoOnSolftAsync()
        Dim clientesEmPreCadastro = Await _http.GetClientesPreCadastroOnSolftAsync()

        'Dim strm As New StreamWriter("c:\ngs\Files\log.txt")

        Dim BancoMobile As New AcessaBancoOnMobile

        Dim sqls As New ArrayList
        Dim Sql As String = String.Empty

        Dim _pedidosSeparadosPorOperacoes As New List(Of WsPedido)

        sqls.Clear()

        For Each p In pedidosOnSoft.Items

            Sql = "Merge PedidoNAOIntegracaoOnSoft as Dest" & vbCrLf &
                  "USING (SELECT '" & p.VendedorCod & "' as VendedorCod, '" & p.PedidoNum & "' as PedidoNum, '" & p.ClienteCod & "' as ClienteCod) AS Ori" & vbCrLf &
                  "	 ON Dest.VendedorCod  = Ori.VendedorCod" & vbCrLf &
                  "	and Dest.PedidoNum    = Ori.PedidoNum" & vbCrLf &
                  "	and Dest.ClienteCod   = Ori.ClienteCod" & vbCrLf &
                  "WHEN NOT MATCHED" & vbCrLf &
                  "	THEN Insert (VendedorCod, PedidoNum, ClienteCod, TabPrecoCod, PedidoData, PedidoVrpagar)" & vbCrLf &
                  "		 Values ('" & p.VendedorCod & "','" & p.PedidoNum & "','" & p.ClienteCod & "','" & p.TabPrecoCod & "','" & p.PedidoData & "', " & p.PedidoVrpagar & ")" & vbCrLf &
                  "WHEN MATCHED" & vbCrLf &
                  "	THEN Update set" & vbCrLf &
                  "			TabPrecoCod    = '" & p.TabPrecoCod & "'" & vbCrLf &
                  "			,PedidoData    = '" & p.PedidoData & "'" & vbCrLf &
                  "			,PedidoVrpagar = " & p.PedidoVrpagar & ";"

            sqls.Add(Sql)

            Dim _CNPJouCPFValido = False

            If Not Funcoes.ValidaCNPJ(p.ClienteCod) AndAlso Not Funcoes.ValidaCPF(p.ClienteCod) Then
                Dim _cliente = clientesEmPreCadastro.Items.FirstOrDefault(Function(c) c.ClienteCod = p.ClienteCod)
                If _cliente IsNot Nothing Then
                    p.ClienteCod = _cliente.ClienteCnpjcpf
                End If
            End If

            If p.ClienteCod.Length = 14 Then
                _CNPJouCPFValido = Funcoes.ValidaCNPJ(p.ClienteCod)
            Else
                _CNPJouCPFValido = Funcoes.ValidaCPF(p.ClienteCod)
            End If

            If _CNPJouCPFValido Then
                Dim _itemsOp2177 = p.PedidoItems.Where(Function(i) Left(i.ProdutoCod, 5) >= 82101 And Left(i.ProdutoCod, 5) <= 82212)
                Dim _itemsOp2180 = p.PedidoItems.Where(Function(i) Left(i.ProdutoCod, 5) = 80101)

                Dim nCliente As Cliente = New Cliente(p.ClienteCod, 0)

                Dim nEmpresa As Cliente = p.VerOrigemEmpresa()

                If _itemsOp2177.Any() Then
                    Dim ped = p.Copy()

                    If p.ClienteCod.Length = 14 AndAlso (nCliente.InscricaoEstadual.Length > 0 AndAlso Not nCliente.InscricaoEstadual = "ISENTO") Then

                        If Left(nEmpresa.Codigo, 8) = "40938762" Then     'BAXI FOODS
                            If nCliente.CodigoEstado = "BA" OrElse
                            nCliente.CodigoEstado = "SC" OrElse
                            nCliente.CodigoEstado = "GO" OrElse
                            nCliente.CodigoEstado = "RN" Then
                                ped.TipoOperacaoCod = "21.78"
                            ElseIf nCliente.CodigoEstado = nEmpresa.CodigoEstado Then
                                ped.TipoOperacaoCod = "21.84"
                            Else
                                ped.TipoOperacaoCod = "21.83"
                            End If
                        ElseIf Left(nEmpresa.Codigo, 8) = "49673784" Then 'BAXI DISTRIBUIDORA
                            If nCliente.CodigoEstado = nEmpresa.CodigoEstado Then
                                ped.TipoOperacaoCod = "21.79"
                            Else
                                ped.TipoOperacaoCod = "21.78"
                            End If
                        Else
                            Throw New Exception("ERRO AO BUSCAR CENTRO DE DISTRIBUICAO")
                        End If
                    Else
                        If Left(nEmpresa.Codigo, 8) = "40938762" Then     'BAXI FOODS
                            ped.TipoOperacaoCod = "21.82"
                        ElseIf Left(nEmpresa.Codigo, 8) = "49673784" Then 'BAXI DISTRIBUIDORA
                            If nCliente.CodigoEstado = nEmpresa.CodigoEstado Then
                                ped.TipoOperacaoCod = "21.79"
                            Else
                                ped.TipoOperacaoCod = "21.78"
                            End If
                        Else
                            Throw New Exception("ERRO AO BUSCAR CENTRO DE DISTRIBUICAO")
                        End If
                    End If

                    ped.PedidoItems = _itemsOp2177.ToList()
                    _pedidosSeparadosPorOperacoes.Add(ped)
                End If

                If _itemsOp2180.Any() Then
                    Dim ped = p.Copy()

                    If Left(nEmpresa.Codigo, 8) = "40938762" Then         'BAXI FOODS
                        If nCliente.CodigoEstado = nEmpresa.CodigoEstado Then
                            ped.TipoOperacaoCod = "21.80"
                        Else
                            ped.TipoOperacaoCod = "21.81"
                        End If
                    ElseIf Left(nEmpresa.Codigo, 8) = "49673784" Then     'BAXI DISTRIBUIDORA
                        If p.ClienteCod.Length = 14 AndAlso (nCliente.InscricaoEstadual.Length > 0 AndAlso Not nCliente.InscricaoEstadual = "ISENTO") Then
                            If nCliente.CodigoEstado = nEmpresa.CodigoEstado Then
                                ped.TipoOperacaoCod = "21.74"
                            Else
                                ped.TipoOperacaoCod = "21.75"
                            End If
                        Else
                            ped.TipoOperacaoCod = "21.82"
                        End If
                    Else
                        Throw New Exception("ERRO AO BUSCAR CENTRO DE DISTRIBUICAO")
                    End If

                    ped.PedidoItems = _itemsOp2180.ToList()
                    _pedidosSeparadosPorOperacoes.Add(ped)
                End If

                If Not _itemsOp2177.Any() AndAlso Not _itemsOp2180.Any() Then
                    _pedidosSeparadosPorOperacoes.Add(p)
                End If
            End If
        Next

        If sqls.Count > 0 Then
            If Not BancoMobile.GravaBanco(sqls) Then
                Throw New Exception("ERRO AO GRAVAR NA TABELA PedidoNAOIntegracaoOnSoft ")
            End If
        End If

        Try
            For Each p In _pedidosSeparadosPorOperacoes

                'DEIXEI AQUI PARA PODER FORÇAR QUANDO PRECISAR - FURLAN - 23/06/2025
                'If p.VendedorCod = "59158924000163" AndAlso p.PedidoNum = "0000000100" Then

                '    Dim pedidoIntegracao As New WsPedidoIntegracao
                '    pedidoIntegracao.VendedorCod = p.VendedorCod
                '    pedidoIntegracao.PedidoNum = p.PedidoNum
                '    pedidoIntegracao.PedidoNumPedCli = "34455"
                '    pedidoIntegracao.PedidoDtEnvio = DateTime.Now
                '    pedidosParaIntegracao.Add(pedidoIntegracao)

                '    Await _http.PostAsync("WsPedidoIntegracao", pedidosParaIntegracao)

                '    Exit Function
                'Else
                '    Continue For
                'End If

                sqls.Clear()

                'Integra o cliente no NGS
                Dim cliente As New WsCliente(p.ClienteCod, 0)

                If (String.IsNullOrWhiteSpace(cliente.ClienteCod)) Then
                    cliente.PreencheCliente(sqls, p.ClienteCod, p.VendedorCod)
                End If

                'Integra os pedido no NGS
                Dim pedidoNGS = p.PreenchePedidoNGS()
                Dim achouEncargos As Boolean = True

                For Each item In pedidoNGS.Itens
                    If item.Encargos Is Nothing OrElse item.Encargos.Count() = 0 Then
                        achouEncargos = False
                        Exit For
                    End If
                Next

                If Not achouEncargos Then
                    Continue For
                End If

                'Numerador Pedido
                Dim num As New NumeradorOnMobile(pedidoNGS.CodigoEmpresa, pedidoNGS.EnderecoEmpresa, 10)
                pedidoNGS.Codigo = (num.Sequencia + 1)
                sqls.Add(num.IncrementarNumeradorSql(False, 1))

                'Numerador Titulo
                Dim numTitulo As New NumeradorOnMobile(1)

                For Each tit In pedidoNGS.Vencimentos
                    tit.Codigo = (numTitulo.Sequencia + 1)
                    sqls.Add(numTitulo.IncrementarNumeradorSql(True, 1))
                Next

                pedidoNGS.SalvarSql(sqls)

                'Insere os pedidos nas tabelas de Pedidos da OnSoft
                p.PedidoNumPedCli = pedidoNGS.Codigo
                p.SalvarSql(sqls)

                For Each item In p.PedidoItems
                    item.SalvarSql(sqls)
                Next

                'Usando para ver os Sqls caso dê erro - Furlan
                'For Each s In sqls
                '    strm.WriteLine(s)
                'Next

                Dim pedidoImportado As Boolean

                If BancoMobile.GravaBanco(sqls) Then
                    pedidoImportado = True
                Else
                    pedidoImportado = False
                End If
            Next

            For Each ped In pedidosOnSoft.Items

                Sql = "SELECT VendedorCod, PedidoNum FROM PedidoXItemIntegracaoOnSoft " & vbCrLf &
                        " WHERE VendedorCod = '" & ped.VendedorCod & "'" & vbCrLf &
                        "   AND PedidoNum   = '" & ped.PedidoNum & "'"

                Dim dsP As New DataSet
                dsP = BancoMobile.ConsultaDataSet(Sql, "IntegracaoDosItens")

                If ped.PedidoItems.Count = dsP.Tables(0).Rows.Count Then
                    'Atualiza o Onsoft que o pedido foi integrado no NGS
                    Dim pedidoIntegracao As New WsPedidoIntegracao
                    pedidoIntegracao.VendedorCod = ped.VendedorCod
                    pedidoIntegracao.PedidoNum = ped.PedidoNum
                    pedidoIntegracao.PedidoNumPedCli = ped.PedidoNumPedCli
                    pedidoIntegracao.PedidoDtEnvio = DateTime.Now
                    pedidosParaIntegracao.Add(pedidoIntegracao)

                    Await _http.PostAsync("WsPedidoIntegracao", pedidosParaIntegracao)
                End If
            Next

        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            'strm.Close()
        End Try
    End Function
End Class
