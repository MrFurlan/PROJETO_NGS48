Imports System.Collections.Generic
Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class RetornoBancario
    Implements IBaseEntity

    '    Dim objBanco As New AcessaBanco
    '    Private Popup As New CarregarPopup


    '#Region "Baixa de Titulos Automatica Recebimento"


    '    Public Function ContabilizarRecebimentoAutomatico()
    '        Dim sql As String
    '        Dim ds_aux As New DataSet
    '        Dim dr As DataRow
    '        Dim ArrySqlAutomatico As New ArrayList
    '        Try

    '            '''Busca os Titulos Agrupados por Data de Baixa
    '            sql = "Select Baixa from ContasApagar where Provisao = 2 and Situacao = 104 group by Baixa"
    '            ds_aux = objBanco.ConsultaDataSet(sql, "Titulos")
    '            For Each dr In ds_aux.Tables(0).Rows
    '                Dim Historico As String = "VALOR REF. COBRANÇA DOS TITULOS "
    '                ''Carrega Titulos na lista com uma determinada data
    '                Dim drTitulo As DataRow
    '                Dim ds_titulos As New DataSet
    '                Dim pTitulosAutomatico As New List(Of FrmTitulosAgrupados.TituloAgrupado)
    '                pTitulosAutomatico.Clear()
    '                sql = "Select Registro_Id from ContasApagar where Provisao = 2 and Situacao = 104 and Baixa = '" & Format(CDate(dr("Baixa")), "yyyy/MM/dd") & "'"
    '                ds_titulos = objBanco.ConsultaDataSet(sql, "Titulos")
    '                For Each drTitulo In ds_titulos.Tables(0).Rows
    '                    Dim i As New FrmTitulosAgrupados.TituloAgrupado
    '                    i.Consultar(drTitulo("Titulo_Id"))
    '                    i.Marcado = False
    '                    i.LiquidoOficial = i.OficialProgramado
    '                    i.LiquidoMoeda = i.MoedaProgramado
    '                    pTitulosAutomatico.Add(i)
    '                    Historico &= i.Titulo & ", "
    '                Next

    '                Dim TituloAuto As New FrmTitulosAgrupados.TituloAgrupado
    '                Dim TituloMestre As New FrmTitulosAgrupados.TituloAgrupado
    '                Dim pLiquidoOficial As Double = 0
    '                Dim pLiquidoMoeda As Double = 0

    '                TituloMestre = GeraTituloMestreRecebimento(pTitulosAutomatico(0))
    '                TituloMestre.Historico = Mid(Historico, 1, 190)
    '                Dim gravar As New GravarTitulo("90", dr("Baixa"), TituloMestre)
    '                'Percorre a lista de titulos 
    '                For Each TituloAuto In pTitulosAutomatico

    '                    sql = "Update ContasAPagar Set "
    '                    sql &= "Provisao=" & CInt(1) & ","
    '                    sql &= "Situacao=" & CInt(110) & ", "
    '                    sql &= "RegistroMestre=" & TituloMestre.Titulo & ""
    '                    sql &= " Where Registro_Id=" & TituloAuto.Titulo & ""
    '                    ArrySqlAutomatico.Add(sql)

    '                    ArrySqlAutomatico.Add(gravar.GravaRazaoCreditoAutomatico(TituloAuto))
    '                    ArrySqlAutomatico.AddRange(GravaDecontosRazao(TituloAuto))

    '                    ''Soma os Titulos para gerar o valor do Titulo Mestre
    '                    pLiquidoOficial += TituloAuto.OficialBaixado
    '                    pLiquidoOficial += TituloAuto.Juros
    '                    pLiquidoOficial += TituloAuto.Acrescimos
    '                    pLiquidoOficial -= TituloAuto.Deducoes
    '                    pLiquidoOficial -= TituloAuto.Descontos


    '                    pLiquidoMoeda += TituloAuto.MoedaBaixado
    '                Next

    '                TituloMestre.LiquidoMoeda = pLiquidoMoeda
    '                TituloMestre.LiquidoOficial = pLiquidoOficial

    '                ArrySqlAutomatico.Add(gravar.GerarTituloMestreAutomarico(TituloMestre, TituloMestre.Moeda, pLiquidoMoeda, pLiquidoOficial))
    '                ArrySqlAutomatico.Add(gravar.GravaRazaoDebitoAutomatico(TituloMestre))

    '            Next



    '            If banco.AtualizaBanco(ArrySqlAutomatico, "Gravar") Then
    '                'MsgBox("Titulos Gravados", MsgBoxStyle.Information, "Titulos Agrupados")
    '            End If
    '        Catch ex As Exception
    '            MsgBox(Me.Page, "Problemas na Baixa de titulos Automáticos")
    '        End Try

    '    End Function

    '    Public Function GeraTituloMestreRecebimento(ByVal ptitulomestre As FrmTitulosAgrupados.TituloAgrupado) As FrmTitulosAgrupados.TituloAgrupado
    '        Dim titulomestre As New FrmTitulosAgrupados.TituloAgrupado
    '        Dim conexao As New Utilitarios.EnderecoConexao
    '        Dim SqlTrans As OleDb.OleDbTransaction
    '        Dim SqlCommand As New OleDb.OleDbCommand
    '        Dim SeqTituloMestre As String
    '        Dim sql As String


    '        '''''''''''''''''''''''''''''''''''''''
    '        ''' Abre a conexao com o banco de dados
    '        '''''''''''''''''''''''''''''''''''''''
    '        Try
    '            If (conexao.SQlConnectionLocal.State = ConnectionState.Closed) Then
    '                conexao.ConectaBanco()
    '            End If
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try

    '        '''''''''''''''''''''''''''''''
    '        '''Gera sequencia de titulos'''
    '        '''''''''''''''''''''''''''''' 
    '        Try
    '            SqlTrans = conexao.SQlConnectionLocal.BeginTransaction(IsolationLevel.ReadCommitted)
    '            SqlCommand.Connection = conexao.SQlConnectionLocal
    '            SqlCommand.Transaction = SqlTrans

    '            sql = "UPDATE Numerador SET "
    '            sql &= "Sequencia = Sequencia +1 "
    '            sql &= " WHERE (Upper(Empresa_Id) = '" & Utilitarios.[Global].NomeServidor.ToUpper & "' and EndEmpresa_Id = '" & 0 & "' And Numerador_Id = " & 1 & ")"
    '            SqlCommand.CommandText = sql
    '            SqlCommand.ExecuteNonQuery()

    '            sql = "Select Sequencia from Numerador where Upper(Empresa_Id) = '" & Utilitarios.[Global].NomeServidor.ToUpper & "' and EndEmpresa_Id = '" & 0 & "' and Numerador_Id = " & 1 & ""
    '            SqlCommand.CommandText = sql
    '            banco.dr = SqlCommand.ExecuteReader
    '            While banco.dr.read()
    '                SeqTituloMestre = banco.dr.GetValue(0)
    '            End While
    '            banco.dr.close()
    '            SqlTrans.Commit()
    '        Catch sqlException As OleDb.OleDbException
    '            SqlTrans.Rollback()
    '            Exit Function
    '            MsgBox(Me.Page, sqlException.Message)
    '        End Try


    '        Try
    '            titulomestre.Pais = Utilitarios.Global.GPais
    '            titulomestre.Titulo = SeqTituloMestre
    '            titulomestre.Provisao = 1
    '            titulomestre.Carteira = 0
    '            titulomestre.Aplicacao = 0
    '            titulomestre.Indexador = 3
    '            titulomestre.TipoPagamento = IIf(ptitulomestre.TipoPagamento = Nothing, 0, ptitulomestre.TipoPagamento)
    '            titulomestre.Situacao = 110
    '            titulomestre.TipoLancamento = IIf(ptitulomestre.TipoLancamento = Nothing, 0, ptitulomestre.TipoLancamento)
    '            titulomestre.Movimento = Format(CDate(ptitulomestre.DataBaixa), "dd/MM/yyyy")
    '            titulomestre.Vencimento = Format(CDate(ptitulomestre.Vencimento), "dd/MM/yyyy")
    '            'If IsDBNull(.Item("Prorrogacao")) Then mdateProrrogacao = #12:00:00 AM# Else mdateProrrogacao = .Item("Prorrogacao")
    '            titulomestre.DataMoeda = Format(CDate(ptitulomestre.DataMoeda), "dd/MM/yyyy")
    '            titulomestre.DataBaixa = Format(CDate(ptitulomestre.DataBaixa), "dd/MM/yyyy")
    '            titulomestre.CreditoNaConta = Format(CDate(ptitulomestre.CreditoNaConta), "dd/MM/yyyy")


    '            titulomestre.EmpresaOrigem = ptitulomestre.EmpresaCredito
    '            titulomestre.EnderecoOrigem = ptitulomestre.EnderecoCredito

    '            titulomestre.EmpresaDestino = ptitulomestre.EmpresaCredito
    '            titulomestre.EnderecoDestino = ptitulomestre.EnderecoCredito
    '            'mstrEmpresaOrigem = .Item("EmpresaOrigem")
    '            'If IsDBNull(.Item("EnderecoOrigem")) Then mintEnderecoOrigem = 0 Else mintEnderecoOrigem = .Item("EnderecoOrigem")
    '            'mstrEmpresaDestino = .Item("EmpresaDestino")
    '            'If IsDBNull(.Item("EnderecoDestino")) Then mintEnderecoDestino = 0 Else mintEnderecoDestino = .Item("EnderecoDestino")

    '            titulomestre.EmpresaDebito = ptitulomestre.EmpresaDebito
    '            titulomestre.EnderecoDebito = ptitulomestre.EnderecoDebito
    '            titulomestre.ContaDebito = ptitulomestre.ContaDebito
    '            titulomestre.CustoDebito = IIf(ptitulomestre.CustoDebito = Nothing, 0, ptitulomestre.CustoDebito)
    '            titulomestre.ClienteDebito = ptitulomestre.ClienteDebito
    '            titulomestre.EndClienteDebito = ptitulomestre.EndClienteDebito

    '            titulomestre.ContaCredito = ptitulomestre.ContaCredito
    '            titulomestre.CustoCredito = IIf(ptitulomestre.CustoCredito = Nothing, 0, ptitulomestre.CustoCredito)
    '            If ptitulomestre.CustoCredito = Nothing Then titulomestre.CustoCredito = 0 Else titulomestre.CustoCredito = ptitulomestre.CustoCredito
    '            titulomestre.EmpresaCredito = ptitulomestre.EmpresaCredito
    '            titulomestre.EnderecoCredito = ptitulomestre.EnderecoCredito
    '            titulomestre.ClienteCredito = ptitulomestre.ClienteCredito
    '            titulomestre.EndClienteCredito = IIf(ptitulomestre.EndClienteCredito = Nothing, 0, ptitulomestre.EndClienteCredito)


    '            titulomestre.ReceberPagar = IIf(ptitulomestre.ReceberPagar = 1, FrmTitulosAgrupados.TituloAgrupado.enumReceberPagar.RECEBER, FrmTitulosAgrupados.TituloAgrupado.enumReceberPagar.PAGAR)
    '            If ptitulomestre.ReceberPagar = 2 Then
    '                titulomestre.Cheque = True
    '            Else
    '                titulomestre.Cheque = False
    '            End If
    '            titulomestre.Slips = True
    '            titulomestre.Recibo = True
    '            titulomestre.Aviso = True
    '            'If IsDBNull(.Item("ReciboDeposito")) Then mblnReciboDeposito = False Else mblnReciboDeposito = (.Item("ReciboDeposito") = "S")
    '            'If IsDBNull(.Item("EmpresaPedido")) Then mstrEmpresaPedido = "" Else mstrEmpresaPedido = .Item("EmpresaPedido")
    '            'If IsDBNull(.Item("EndEmpresaPedido")) Then mintEndEmpresaPedido = -1 Else mintEndEmpresaPedido = .Item("EndEmpresaPedido")
    '            'If IsDBNull(.Item("Pedido")) Then mstrPedido = "" Else mstrPedido = .Item("Pedido")
    '            'If IsDBNull(.Item("PedidoFixacao")) Then mstrPedidoFixacao = "" Else mstrPedidoFixacao = .Item("PedidoFixacao")
    '            titulomestre.OficialBaixado = 0
    '            titulomestre.MoedaBaixado = 0
    '            titulomestre.OficialProgramado = 0
    '            titulomestre.MoedaProgramado = 0

    '            titulomestre.Historico = Mid(ptitulomestre.Historico, 1, 190)

    '            titulomestre.Destinatario = ptitulomestre.EmpresaDebito
    '            titulomestre.EndDestinatario = ptitulomestre.EnderecoDebito

    '            If ptitulomestre.Banco = Nothing Then titulomestre.Banco = 0 Else titulomestre.Banco = ptitulomestre.Banco
    '            If ptitulomestre.Agencia = Nothing Then titulomestre.Agencia = 0 Else titulomestre.Agencia = ptitulomestre.Agencia

    '            titulomestre.DigitoAgencia = ptitulomestre.DigitoAgencia
    '            titulomestre.ContaCorrente = ptitulomestre.ContaCorrente
    '            titulomestre.DigitoConta = ptitulomestre.DigitoConta
    '            titulomestre.Destinacao = ptitulomestre.Destinacao

    '            titulomestre.Produto = "0"
    '            titulomestre.Safra = ""


    '            titulomestre.Moeda = ptitulomestre.Moeda


    '            titulomestre.RetencaoOficial = 0
    '            titulomestre.RetencaoMoeda = 0

    '            titulomestre.FaturaNumero = ""
    '            titulomestre.FaturaSerie = ""


    '            titulomestre.EmpresaCredito = ""
    '            titulomestre.EnderecoCredito = 0
    '            titulomestre.ContaCredito = ""
    '            titulomestre.ClienteCredito = ""
    '            titulomestre.EndClienteCredito = 0
    '            titulomestre.CustoCredito = 0
    '            titulomestre.Provisao = 1


    '            '''''''''''''''''''''''''''''''
    '            '''      Fecha Conexao      '''
    '            ''''''''''''''''''''''''''''''' 
    '            conexao.FecheConexaoBanco()
    '        Catch ex As Exception
    '            MsgBox(Me.Page, "Geração titulo mestre-titulos")
    '        End Try

    '        Return titulomestre
    '    End Function

    '    Public Function GravaDecontosRazao(ByVal ptitulo As FrmTitulosAgrupados.TituloAgrupado) As ArrayList
    '        Dim dsDescontos As DataSet
    '        Dim dr As DataRow
    '        Dim Sql As String
    '        Dim sqlArray As New ArrayList
    '        ''Valida contas de descontos
    '        Try
    '            If ptitulo.Descontos <> 0 Or ptitulo.Deducoes <> 0 Or ptitulo.Juros <> 0 Or ptitulo.Acrescimos <> 0 Then
    '                If Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.PAGAR Then
    '                    Sql = "Select Conta_Id,Descontos,Deducoes,Juros,Acrescimos from ContasXDescontos where Conta_Id ='" & Mid(ptitulo.ContaDebito, 1, 5) & "'"
    '                    dsDescontos = banco.ConsultaDataSet(Sql, "ContasXDescontos")
    '                    For Each dr In dsDescontos.Tables(0).Rows
    '                        If ptitulo.Descontos <> 0 Then
    '                            If dr("Descontos") <> Nothing And IsDBNull(dr("Descontos")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoCredito(ptitulo, "90", dr("Descontos"), ptitulo.Descontos))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de descontos para gravar o valor do desconto / Titulos")

    '                            End If
    '                        End If
    '                        If ptitulo.Deducoes <> 0 Then
    '                            If dr("Deducoes") <> Nothing And IsDBNull(dr("Deducoes")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoCredito(ptitulo, "90", dr("Deducoes"), ptitulo.Deducoes))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de deduçoes para gravar o valor da dedução / Titulos")

    '                            End If

    '                        End If
    '                        If ptitulo.Juros <> 0 Then
    '                            If dr("Juros") <> Nothing And IsDBNull(dr("Juros")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoDebito(ptitulo, "90", dr("Juros"), ptitulo.Juros))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de juros para gravar o valor do juro / Titulos")

    '                            End If
    '                        End If
    '                        If ptitulo.Acrescimos <> 0 Then
    '                            If dr("Acrescimos") <> Nothing And IsDBNull(dr("Acrescimos")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoDebito(ptitulo, "90", dr("Acrescimos"), ptitulo.Acrescimos))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de acrescimos para gravar o valor do acrescimo / Titulos")

    '                            End If
    '                        End If

    '                    Next

    '                ElseIf Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.RECEBER Then
    '                    Sql = "Select Conta_Id,Descontos,Deducoes,Juros,Acrescimos from ContasXDescontos where Conta_Id ='" & ptitulo.ContaCredito & "'"
    '                    dsDescontos = banco.ConsultaDataSet(Sql, "ContasXDescontos")
    '                    For Each dr In dsDescontos.Tables(0).Rows
    '                        If ptitulo.Descontos <> 0 Then
    '                            If dr("Descontos") <> Nothing And IsDBNull(dr("Descontos")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoDebito(ptitulo, "90", dr("Descontos"), ptitulo.Descontos))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de descontos para gravar o valor do desconto / Titulos")

    '                            End If
    '                        End If
    '                        If ptitulo.Deducoes <> 0 Then
    '                            If dr("Deducoes") <> Nothing And IsDBNull(dr("Deducoes")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoDebito(ptitulo, "90", dr("Deducoes"), ptitulo.Deducoes))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de deduções para gravar o valor da dedução / Titulos")

    '                            End If
    '                        End If
    '                        If ptitulo.Juros <> 0 Then
    '                            If dr("Juros") <> Nothing And IsDBNull(dr("Juros")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoCredito(ptitulo, "90", dr("Juros"), ptitulo.Juros))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de juros para gravar o valor do jurro / Titulos")
    '                            End If

    '                        End If
    '                        If ptitulo.Acrescimos <> 0 Then
    '                            If dr("Acrescimos") <> Nothing And IsDBNull(dr("Acrescimos")) = False Then
    '                                sqlArray.Add(GravaDescontosRazaoCredito(ptitulo, "90", dr("Acrescimos"), ptitulo.Acrescimos))
    '                            Else
    '                                MsgBox(Me.Page, "É nescessario informa a conta de acrescimos para gravar o valor do acrescimo / Titulos")

    '                            End If
    '                        End If
    '                    Next


    '                End If
    '            End If
    '            Return sqlArray
    '        Catch ex As Exception
    '            MsgBox(Me.Page, "Problemas com a rotina de descontos do titulo")
    '        End Try
    '    End Function

    '    Public Function GravaDescontosRazaoDebito(ByVal ptitulo As FrmTitulosAgrupados.TituloAgrupado, ByVal pLote As String, ByVal pConta As String, ByVal ValorDesconto As Decimal) As String
    '        Dim conexao As New Utilitarios.EnderecoConexao
    '        Dim SqlTransLoteDebito As OleDb.OleDbTransaction
    '        Dim SqlCommandLoteDebito As New OleDb.OleDbCommand
    '        Dim SeqLoteDebito As String
    '        Dim RealizadoPrevisto As String
    '        Dim ValorCotacao As Double
    '        Dim sql As String


    '        ''''''''''''''''''''''''''''''''''''''''''
    '        ''' VERIFICA SE COTAÇÃO REALIZADA/PREVISTA
    '        ''''''''''''''''''''''''''''''''''''''''''
    '        Try


    '            sql = " Select  Realizado,Cotacao from Cotacoes WHERE (Indexador_Id = " & CInt(3) & " And " & "convert(char (12),Data_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "')"
    '            banco.dr = banco.ConsultaBanco(sql)
    '            If banco.dr.read() = True Then
    '                RealizadoPrevisto = banco.dr.GetValue(0)
    '                ValorCotacao = banco.dr.getvalue(1)
    '            Else
    '                MsgBox(Me.Page, "Nenhuma cotação foi realizada para esta data de cotação!!!")
    '                banco.dr.close()
    '                Return False
    '            End If
    '            banco.dr.close()
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try




    '        '''''''''''''''''''''''''''''''''''''''
    '        ''' Abre a conexao com o banco de dados
    '        '''''''''''''''''''''''''''''''''''''''
    '        Try
    '            If (conexao.SQlConnectionLocal.State = ConnectionState.Closed) Then
    '                conexao.ConectaBanco()
    '            End If
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try


    '        '''''''''''''''''''''''''''''''''''''''''''
    '        'GERA SEQUENCIA DE LOTES EMPRESA DE DEBITO' 
    '        '''''''''''''''''''''''''''''''''''''''''''
    '        Try
    '            SqlTransLoteDebito = conexao.SQlConnectionLocal.BeginTransaction(IsolationLevel.ReadCommitted)
    '            SqlCommandLoteDebito.Connection = conexao.SQlConnectionLocal
    '            SqlCommandLoteDebito.Transaction = SqlTransLoteDebito

    '            sql = "Select Sequencia from SequenciasDeLotes"
    '            sql &= " WHERE (Empresa_Id = '" & ptitulo.EmpresaDebito & "' And EndEmpresa_Id=" & ptitulo.EnderecoDebito & " and convert(char (12),Movimento_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "' and Lote_Id=" & pLote & ")"
    '            SqlCommandLoteDebito.CommandText = sql
    '            banco.dr = SqlCommandLoteDebito.ExecuteReader
    '            If banco.dr.Read() = True Then
    '                SeqLoteDebito = banco.dr.getValue(0) + 1
    '                sql = "UPDATE SequenciasDeLotes SET "
    '                sql &= "Sequencia = Sequencia +1"
    '                sql &= " WHERE (Empresa_Id = '" & ptitulo.EmpresaDebito & "' And EndEmpresa_Id=" & ptitulo.EnderecoDebito & " and convert(char (12),Movimento_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "' and Lote_Id=" & pLote & ")"
    '                banco.dr.close()
    '                SqlCommandLoteDebito.CommandText = sql
    '                SqlCommandLoteDebito.ExecuteNonQuery()
    '            Else
    '                sql = "INSERT INTO SequenciasDeLotes (Empresa_Id,EndEmpresa_Id,Movimento_Id,Lote_Id,Sequencia)"
    '                sql &= " Values ('"
    '                sql &= Utilitarios.[Global].GPais & "','"
    '                sql &= ptitulo.EmpresaDebito & "',"
    '                sql &= ptitulo.EnderecoDebito & ",'"
    '                sql &= Format(CDate(ptitulo.Movimento), "yyyy/MM/dd") & "',"
    '                sql &= pLote & ","
    '                sql &= 1 & ")"
    '                banco.dr.close()
    '                SqlCommandLoteDebito.CommandText = sql
    '                SqlCommandLoteDebito.ExecuteNonQuery()
    '                SeqLoteDebito = 1
    '            End If
    '            banco.dr.close()
    '            SqlTransLoteDebito.Commit()
    '        Catch sqlException As OleDb.OleDbException
    '            SqlTransLoteDebito.Rollback()
    '            MsgBox(Me.Page, sqlException.Message)
    '            Exit Function
    '        End Try
    '        '''''''''''''''''''''''''''''''''''  
    '        'INSERE EMPRESA DE DEBITO NO RAZAO
    '        '''''''''''''''''''''''''''''''''''
    '        sql = "INSERT INTO Razao(Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Produto,"
    '        sql &= "Movimento_Id,Lote_Id,Sequencia_Id,Titulo,Pedido,TipoLancto,CentroDeCustos,Indexador,DataMoeda,"
    '        sql &= "DebitoOficial,CreditoOficial,DebitoMoeda,CreditoMoeda,Historico,Moeda,PrevistoRealizado) Values ('"
    '        sql &= ptitulo.EmpresaDebito & "',"
    '        sql &= ptitulo.EnderecoDebito & ",'"
    '        sql &= pConta & "','"
    '        sql &= ptitulo.ClienteDebito & "',"
    '        sql &= CInt(ptitulo.EndClienteDebito) & ",'"
    '        sql &= ptitulo.Produto & "','"

    '        sql &= Format(CDate(ptitulo.DataBaixa), "yyyy/MM/dd") & "',"
    '        sql &= CInt(pLote) & ","
    '        sql &= CInt(SeqLoteDebito) & ","
    '        sql &= ptitulo.Titulo & ","
    '        If ptitulo.Pedido <> Nothing Then
    '            sql &= ptitulo.Pedido & ","
    '        Else
    '            sql &= 0 & ","
    '        End If

    '        sql &= CInt(ptitulo.TipoLancamento) & ","


    '        sql &= CInt(ptitulo.CustoDebito) & ","

    '        sql &= 3 & ",'"
    '        sql &= Format(CDate(ptitulo.Movimento), "yyyy/MM/dd") & "',"
    '        If Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.PAGAR Then

    '            sql &= Replace(ValorDesconto, ",", ".") & ","

    '        ElseIf Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.RECEBER Then

    '            sql &= Replace(ValorDesconto, ",", ".") & ","

    '        End If
    '        sql &= 0 & ","
    '        If Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.PAGAR Then
    '            If (RealizadoPrevisto = "R" Or RealizadoPrevisto = "P") And ptitulo.TipoLancamento <> "8" Then
    '                Dim valormoeda As Double
    '                valormoeda = CDbl(ValorDesconto) / ValorCotacao
    '                sql &= Replace(valormoeda, ",", ".") & ","
    '            Else
    '                sql &= 0 & ","
    '            End If
    '        ElseIf Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.RECEBER Then
    '            If (RealizadoPrevisto = "R" Or RealizadoPrevisto = "P") And ptitulo.TipoLancamento <> "8" Then
    '                Dim valormoeda As Double
    '                valormoeda = CDbl(ValorDesconto) / ValorCotacao
    '                sql &= Replace(valormoeda, ",", ".") & ","
    '            Else
    '                sql &= 0 & ","
    '            End If
    '        End If
    '        sql &= 0 & ",'"
    '        sql &= UCase(RTrim(ptitulo.Historico)) & "','"

    '        If ptitulo.LiquidoMoeda <> 0 And ptitulo.LiquidoOficial <> 0 Then
    '            sql &= 0 & "','"
    '        ElseIf ptitulo.TipoLancamento = 8 Then
    '            sql &= 0 & "','"
    '        ElseIf ptitulo.LiquidoMoeda <> 0 Then
    '            sql &= 3 & "','"
    '        ElseIf ptitulo.LiquidoOficial <> 0 Then
    '            sql &= 1 & "','"
    '        End If
    '        End If


    '        sql &= "R" & "')"

    '        Return sql

    '        '''Fecha conexao
    '        conexao.FecheConexaoBanco()



    '    End Function

    '    Public Function GravaDescontosRazaoCredito(ByVal ptitulo As FrmTitulosAgrupados.TituloAgrupado, ByVal pLote As String, ByVal pConta As String, ByVal ValorDesconto As Decimal) As String
    '        Dim conexao As New Utilitarios.EnderecoConexao
    '        Dim SqlTransLoteCredito As OleDb.OleDbTransaction
    '        Dim SqlCommandLoteCredito As New OleDb.OleDbCommand
    '        Dim SeqLoteCredito As String
    '        Dim RealizadoPrevisto As String
    '        Dim ValorCotacao As Double
    '        Dim sql As String


    '        ''''''''''''''''''''''''''''''''''''''''''
    '        ''' VERIFICA SE COTAÇÃO REALIZADA/PREVISTA
    '        ''''''''''''''''''''''''''''''''''''''''''
    '        Try


    '            sql = " Select  Realizado,Cotacao from Cotacoes WHERE (Indexador_Id = " & CInt(3) & " And " & "convert(char (12),Data_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "')"
    '            banco.dr = banco.ConsultaBanco(sql)
    '            If banco.dr.read() = True Then
    '                RealizadoPrevisto = banco.dr.GetValue(0)
    '                ValorCotacao = banco.dr.getvalue(1)
    '            Else
    '                MsgBox(Me.Page, "Nenhuma cotação foi realizada para esta data de cotação!!!")
    '                banco.dr.close()
    '                Return False
    '            End If
    '            banco.dr.close()
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try

    '        '''''''''''''''''''''''''''''''''''''''
    '        ''' Abre a conexao com o banco de dados
    '        '''''''''''''''''''''''''''''''''''''''
    '        Try
    '            If (conexao.SQlConnectionLocal.State = ConnectionState.Closed) Then
    '                conexao.ConectaBanco()
    '            End If
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try


    '        ''''''''''''''''''''''''''''''''''''''''''''
    '        'GERA SEQUENCIA DE LOTES EMPRESA DE CREDITO' 
    '        ''''''''''''''''''''''''''''''''''''''''''''
    '        Try
    '            SqlTransLoteCredito = conexao.SQlConnectionLocal.BeginTransaction(IsolationLevel.ReadCommitted)
    '            SqlCommandLoteCredito.Connection = conexao.SQlConnectionLocal
    '            SqlCommandLoteCredito.Transaction = SqlTransLoteCredito

    '            sql = "Select Sequencia from SequenciasDeLotes"
    '            sql &= " WHERE (Empresa_Id = '" & ptitulo.EmpresaCredito & "' And EndEmpresa_Id=" & ptitulo.EnderecoCredito & " and convert(char (12),Movimento_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "' and Lote_Id=" & pLote & ")"
    '            SqlCommandLoteCredito.CommandText = sql
    '            banco.dr = SqlCommandLoteCredito.ExecuteReader
    '            If banco.dr.Read() = True Then
    '                SeqLoteCredito = banco.dr.getValue(0) + 1
    '                sql = "UPDATE SequenciasDeLotes SET "
    '                sql &= "Sequencia = Sequencia +1"
    '                sql &= " WHERE (Empresa_Id = '" & ptitulo.EmpresaCredito & "' And EndEmpresa_Id=" & ptitulo.EnderecoCredito & " and convert(char (12),Movimento_Id,103)='" & Format(CDate(ptitulo.DataBaixa), "dd/MM/yyyy") & "' and Lote_Id=" & pLote & ")"
    '                banco.dr.close()
    '                SqlCommandLoteCredito.CommandText = sql
    '                SqlCommandLoteCredito.ExecuteNonQuery()
    '            Else
    '                sql = "INSERT INTO SequenciasDeLotes (Empresa_Id,EndEmpresa_Id,Movimento_Id,Lote_Id,Sequencia)"
    '                sql &= " Values ('"
    '                sql &= ptitulo.EmpresaCredito & "',"
    '                sql &= ptitulo.EnderecoCredito & ",'"
    '                sql &= Format(CDate(ptitulo.DataBaixa), "yyyy/MM/dd") & "',"
    '                sql &= pLote & ","
    '                sql &= 1 & ")"
    '                banco.dr.close()
    '                SqlCommandLoteCredito.CommandText = sql
    '                SqlCommandLoteCredito.ExecuteNonQuery()
    '                SeqLoteCredito = 1
    '            End If
    '            banco.dr.close()
    '            SqlTransLoteCredito.Commit()
    '        Catch sqlException As OleDb.OleDbException
    '            SqlTransLoteCredito.Rollback()
    '            MsgBox(Me.Page, sqlException.Message)
    '            Exit Function
    '        End Try
    '        '''''''''''''''''''''''''''''''''''
    '        'INSERE EMPRESA DE CREDITO NO RAZAO
    '        '''''''''''''''''''''''''''''''''''
    '        sql = "INSERT INTO Razao(Empresa_Id,EndEmpresa_Id,Conta_Id,Cliente_Id,EndCliente_Id,Produto,"
    '        sql &= "Movimento_Id,Lote_Id,Sequencia_Id,Titulo,Pedido,TipoLancto,CentroDeCustos,Indexador,DataMoeda,"
    '        sql &= "DebitoOficial,CreditoOficial,DebitoMoeda,CreditoMoeda,Historico,Moeda,PrevistoRealizado) Values ('"
    '        sql &= ptitulo.EmpresaCredito & "',"
    '        sql &= ptitulo.EnderecoCredito & ",'"
    '        sql &= pConta & "','"
    '        sql &= ptitulo.ClienteCredito & "',"
    '        sql &= CInt(ptitulo.EnderecoCredito) & ",'"
    '        sql &= ptitulo.Produto & "','"
    '        sql &= Format(CDate(ptitulo.Movimento), "yyyy/MM/dd") & "',"
    '        sql &= CInt(pLote) & ","
    '        sql &= CInt(SeqLoteCredito) & ","
    '        sql &= ptitulo.Titulo & ","
    '        If ptitulo.Pedido <> Nothing Then
    '            sql &= ptitulo.Pedido & ","
    '        Else
    '            sql &= 0 & ","
    '        End If

    '        sql &= CInt(ptitulo.TipoLancamento) & ","

    '        sql &= CInt(ptitulo.CustoCredito) & ","



    '        sql &= 3 & ",'"
    '        sql &= Format(CDate(ptitulo.Movimento), "yyyy/MM/dd") & "',"
    '        sql &= 0 & ","
    '        If Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.PAGAR Then
    '            sql &= Replace(ValorDesconto, ",", ".") & ","
    '        ElseIf Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.RECEBER Then
    '            sql &= Replace(ValorDesconto, ",", ".") & ","
    '        End If
    '        sql &= 0 & ","
    '        If Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.PAGAR Then
    '            If (RealizadoPrevisto = "R" Or RealizadoPrevisto = "P") And ptitulo.TipoLancamento <> "8" Then
    '                Dim valormoedacred As Double
    '                valormoedacred = CDbl(ValorDesconto) / ValorCotacao
    '                sql &= Replace(valormoedacred, ",", ".") & ",'"
    '            Else
    '                sql &= 0 & ",'"
    '            End If
    '        ElseIf Trim(ptitulo.ReceberPagar) = Negocio.Financeiro.Titulo.enumReceberPagar.RECEBER Then
    '            If (RealizadoPrevisto = "R" Or RealizadoPrevisto = "P") And ptitulo.TipoLancamento <> "8" Then
    '                Dim valormoedacred As Double
    '                valormoedacred = CDbl(ValorDesconto) / ValorCotacao
    '                sql &= Replace(valormoedacred, ",", ".") & ",'"
    '            Else
    '                sql &= 0 & ",'"
    '            End If
    '        End If
    '        sql &= UCase(RTrim(ptitulo.Historico)) & "','"


    '        sql &= ptitulo.Moeda & "','"

    '        sql &= "R" & "')"

    '        '''Fecha conexao
    '        conexao.FecheConexaoBanco()
    '        Return sql




    '    End Function


    '#End Region

    '#Region "Baixa de Titulos Automatica Pagamento"
    '    Public Function ContabilizarPagamentoAutomatico() As Boolean
    '        Dim sql As String
    '        Dim ds_aux As New DataSet
    '        Dim dr As DataRow
    '        Dim ArrySqlAutomatico As New ArrayList
    '        Try
    '            ''''Contabiliza titulos que tenham mestre
    '            sql = "Select Registro_Id As Titulo_Id from ContasAPagar where Registro_Id = RegistroMestre and Situacao = 204 "
    '            ds_aux = banco.ConsultaDataSet(sql, "Titulos")
    '            Dim drTitulo As DataRow
    '            For Each drTitulo In ds_aux.Tables(0).Rows
    '                ArrySqlAutomatico.Clear()
    '                Dim tMestre As New FrmTitulosAgrupados.TituloAgrupado
    '                tMestre.Consultar(drTitulo("Titulo_Id"))
    '                tMestre.Marcado = False
    '                tMestre.LiquidoOficial = tMestre.OficialBaixado
    '                tMestre.LiquidoMoeda = tMestre.MoedaBaixado


    '                Dim ds_titulos As New DataSet
    '                Dim drT As DataRow
    '                Dim gravar As New GravarTitulo("95", tMestre.DataBaixa, tMestre)
    '                sql = "Select Registro_Id as Titulo_Id from ContasAPagar where Provisao = 2 and RegistroMestre = '" & tMestre.Titulo & "' and Registro_Id <> RegistroMestre and Situacao = 205 "
    '                ds_titulos = banco.ConsultaDataSet(sql, "Titulos")
    '                For Each drT In ds_titulos.Tables(0).Rows
    '                    Dim tAutomatico As New FrmTitulosAgrupados.TituloAgrupado
    '                    tAutomatico.Consultar(drT("Titulo_Id"))
    '                    tAutomatico.Marcado = False
    '                    tAutomatico.LiquidoOficial = tAutomatico.OficialBaixado
    '                    tAutomatico.LiquidoMoeda = tAutomatico.MoedaBaixado

    '                    sql = "Update ContasAPagar Set "
    '                    sql &= "Provisao=" & CInt(1) & ","
    '                    sql &= "Situacao=" & CInt(210) & " "
    '                    sql &= " Where Titulo_Id=" & tAutomatico.Titulo & ""
    '                    ArrySqlAutomatico.Add(sql)


    '                    ArrySqlAutomatico.Add(gravar.GravaRazaoDebitoAutomatico(tAutomatico))
    '                    ArrySqlAutomatico.AddRange(GravaDecontosRazao(tAutomatico))

    '                Next

    '                sql = "Update Titulos Set "
    '                sql &= "Provisao =" & CInt(1) & ","
    '                sql &= "Situacao =" & CInt(210) & ", "
    '                sql &= "OficialBaixado=OficialProgramado,"
    '                sql &= "MoedaBaixado=MoedaProgramado "
    '                sql &= " Where Titulo_Id=" & tMestre.Titulo & ""
    '                ArrySqlAutomatico.Add(sql)
    '                ArrySqlAutomatico.Add(gravar.GravaRazaoCreditoAutomatico(tMestre))

    '                If Not objBanco.GravaBanco(ArrySqlAutomatico) Then
    '                    'MsgBox("Titulos Gravados", MsgBoxStyle.Information, "Titulos Agrupados")
    '                    Return False
    '                End If
    '            Next

    '            ''Contabiliza titulos sem mestre
    '            Dim dsTitulosInd As New DataSet
    '            Dim drTitulosInd As DataRow
    '            '19-10

    '            sql = "select Registro_Id from ContasAPagar where Situacao = 204 and (RegistroMestre = 0 or RegistroMestre = '' or RegistroMestre is null)"
    '            dsTitulosInd = objBanco.ConsultaDataSet(sql, "Titulos")
    '            For Each drTitulosInd In dsTitulosInd.Tables(0).Rows
    '                ArrySqlAutomatico.Clear()
    '                Dim tIndividual As New FrmTitulosAgrupados.TituloAgrupado
    '                tIndividual.Consultar(drTitulosInd("Pais_Id"), drTitulosInd("Titulo_Id"))
    '                tIndividual.Marcado = False
    '                tIndividual.LiquidoOficial = tIndividual.OficialBaixado
    '                tIndividual.LiquidoMoeda = tIndividual.MoedaBaixado

    '                Dim gravar As New GravarTitulo("95", tIndividual.DataBaixa, tIndividual)

    '                sql = "Update ContasAPagar Set "
    '                sql &= "Provisao=" & CInt(1) & ","
    '                sql &= "Situacao=" & CInt(210) & " "
    '                sql &= " Where Registro_Id=" & tIndividual.Titulo & ""
    '                ArrySqlAutomatico.Add(sql)

    '                tIndividual.LiquidoOficial = tIndividual.OficialBaixado
    '                tIndividual.LiquidoMoeda = tIndividual.MoedaBaixado

    '                ArrySqlAutomatico.Add(gravar.GravaRazaoDebitoAutomatico(tIndividual))
    '                ArrySqlAutomatico.Add(gravar.GravaRazaoCreditoAutomatico(tIndividual))
    '                ArrySqlAutomatico.AddRange(GravaDecontosRazao(tIndividual))

    '                If Not banco.AtualizaBanco(ArrySqlAutomatico, "Gravar") Then
    '                    'MsgBox("Titulos Gravados", MsgBoxStyle.Information, "Titulos Agrupados")
    '                    Return False
    '                End If
    '            Next

    '        Catch ex As Exception
    '            MsgBox(Me.Page, "Problemas na Baixa de titulos Automáticos")
    '            Return False
    '        End Try
    '        Return True
    '    End Function


    '    Public Function GeraTituloMestrePagamento(ByVal ptitulomestre As FrmTitulosAgrupados.TituloAgrupado) As FrmTitulosAgrupados.TituloAgrupado
    '        Dim titulomestre As New FrmTitulosAgrupados.TituloAgrupado
    '        Dim conexao As New Utilitarios.EnderecoConexao
    '        Dim SqlTrans As OleDb.OleDbTransaction
    '        Dim SqlCommand As New OleDb.OleDbCommand
    '        Dim SeqTituloMestre As String
    '        Dim sql As String


    '        '''''''''''''''''''''''''''''''''''''''
    '        ''' Abre a conexao com o banco de dados
    '        '''''''''''''''''''''''''''''''''''''''
    '        Try
    '            If (conexao.SQlConnectionLocal.State = ConnectionState.Closed) Then
    '                conexao.ConectaBanco()
    '            End If
    '        Catch ex As Exception
    '            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '        End Try

    '        '''''''''''''''''''''''''''''''
    '        '''Gera sequencia de titulos'''
    '        '''''''''''''''''''''''''''''' 
    '        Try
    '            SqlTrans = conexao.SQlConnectionLocal.BeginTransaction(IsolationLevel.ReadCommitted)
    '            SqlCommand.Connection = conexao.SQlConnectionLocal
    '            SqlCommand.Transaction = SqlTrans

    '            sql = "UPDATE Numerador SET "
    '            sql &= "Sequencia = Sequencia +1 "
    '            sql &= " WHERE (Upper(Empresa_Id) = '" & Utilitarios.[Global].NomeServidor.ToUpper & "' and EndEmpresa_Id = '" & 0 & "' And Numerador_Id = " & 1 & ")"
    '            SqlCommand.CommandText = sql
    '            SqlCommand.ExecuteNonQuery()

    '            sql = "Select Sequencia from Numerador where Upper(Empresa_Id) = '" & Utilitarios.[Global].NomeServidor.ToUpper & "' and EndEmpresa_Id = '" & 0 & "' and Numerador_Id = " & 1 & ""
    '            SqlCommand.CommandText = sql
    '            banco.dr = SqlCommand.ExecuteReader
    '            While banco.dr.read()
    '                SeqTituloMestre = banco.dr.GetValue(0)
    '            End While
    '            banco.dr.close()
    '            SqlTrans.Commit()
    '        Catch sqlException As OleDb.OleDbException
    '            SqlTrans.Rollback()
    '            Exit Function
    '            MsgBox(Me.Page, sqlException.Message)
    '        End Try


    '        Try
    '            titulomestre.Titulo = SeqTituloMestre
    '            titulomestre.Provisao = 2
    '            titulomestre.Carteira = IIf(ptitulomestre.Carteira = Nothing, 0, ptitulomestre.Carteira)
    '            titulomestre.Aplicacao = 0
    '            titulomestre.Indexador = 3
    '            titulomestre.TipoPagamento = IIf(ptitulomestre.TipoPagamento = Nothing, 0, ptitulomestre.TipoPagamento)
    '            titulomestre.Situacao = 1
    '            titulomestre.TipoLancamento = IIf(ptitulomestre.TipoLancamento = Nothing, 0, ptitulomestre.TipoLancamento)
    '            titulomestre.Movimento = Format(CDate(ptitulomestre.Movimento), "dd/MM/yyyy")
    '            titulomestre.Vencimento = Format(CDate(ptitulomestre.Vencimento), "dd/MM/yyyy")
    '            'If IsDBNull(.Item("Prorrogacao")) Then mdateProrrogacao = #12:00:00 AM# Else mdateProrrogacao = .Item("Prorrogacao")
    '            titulomestre.DataMoeda = Format(CDate(ptitulomestre.DataMoeda), "dd/MM/yyyy")
    '            titulomestre.DataBaixa = Format(CDate(ptitulomestre.DataBaixa), "dd/MM/yyyy")

    '            titulomestre.EmpresaOrigem = ptitulomestre.EmpresaOrigem
    '            titulomestre.EnderecoOrigem = ptitulomestre.EnderecoOrigem

    '            titulomestre.EmpresaDestino = ptitulomestre.EmpresaDestino
    '            titulomestre.EnderecoDestino = ptitulomestre.EnderecoDestino
    '            'mstrEmpresaOrigem = .Item("EmpresaOrigem")
    '            'If IsDBNull(.Item("EnderecoOrigem")) Then mintEnderecoOrigem = 0 Else mintEnderecoOrigem = .Item("EnderecoOrigem")
    '            'mstrEmpresaDestino = .Item("EmpresaDestino")
    '            'If IsDBNull(.Item("EnderecoDestino")) Then mintEnderecoDestino = 0 Else mintEnderecoDestino = .Item("EnderecoDestino")

    '            titulomestre.EmpresaDebito = ptitulomestre.EmpresaDebito
    '            titulomestre.EnderecoDebito = ptitulomestre.EnderecoDebito
    '            titulomestre.ContaDebito = ptitulomestre.ContaDebito
    '            titulomestre.CustoDebito = IIf(ptitulomestre.CustoDebito = Nothing, 0, ptitulomestre.CustoDebito)
    '            titulomestre.ClienteDebito = ptitulomestre.ClienteDebito
    '            titulomestre.EndClienteDebito = ptitulomestre.EndClienteDebito

    '            titulomestre.ContaCredito = ptitulomestre.ContaCredito
    '            titulomestre.CustoCredito = IIf(ptitulomestre.CustoCredito = Nothing, 0, ptitulomestre.CustoCredito)
    '            If ptitulomestre.CustoCredito = Nothing Then titulomestre.CustoCredito = 0 Else titulomestre.CustoCredito = ptitulomestre.CustoCredito
    '            titulomestre.EmpresaCredito = ptitulomestre.EmpresaCredito
    '            titulomestre.EnderecoCredito = ptitulomestre.EnderecoCredito
    '            titulomestre.ClienteCredito = ptitulomestre.ClienteCredito
    '            titulomestre.EndClienteCredito = IIf(ptitulomestre.EndClienteCredito = Nothing, 0, ptitulomestre.EndClienteCredito)


    '            titulomestre.ReceberPagar = IIf(ptitulomestre.ReceberPagar = 1, FrmTitulosAgrupados.TituloAgrupado.enumReceberPagar.RECEBER, FrmTitulosAgrupados.TituloAgrupado.enumReceberPagar.PAGAR)
    '            If ptitulomestre.ReceberPagar = 2 Then
    '                titulomestre.Cheque = True
    '            Else
    '                titulomestre.Cheque = False
    '            End If
    '            titulomestre.Slips = True
    '            titulomestre.Recibo = True
    '            titulomestre.Aviso = True
    '            'If IsDBNull(.Item("ReciboDeposito")) Then mblnReciboDeposito = False Else mblnReciboDeposito = (.Item("ReciboDeposito") = "S")
    '            'If IsDBNull(.Item("EmpresaPedido")) Then mstrEmpresaPedido = "" Else mstrEmpresaPedido = .Item("EmpresaPedido")
    '            'If IsDBNull(.Item("EndEmpresaPedido")) Then mintEndEmpresaPedido = -1 Else mintEndEmpresaPedido = .Item("EndEmpresaPedido")
    '            'If IsDBNull(.Item("Pedido")) Then mstrPedido = "" Else mstrPedido = .Item("Pedido")
    '            'If IsDBNull(.Item("PedidoFixacao")) Then mstrPedidoFixacao = "" Else mstrPedidoFixacao = .Item("PedidoFixacao")
    '            titulomestre.OficialBaixado = 0
    '            titulomestre.MoedaBaixado = 0
    '            titulomestre.OficialProgramado = 0
    '            titulomestre.MoedaProgramado = 0

    '            titulomestre.Historico = ptitulomestre.Historico

    '            titulomestre.Destinatario = ptitulomestre.ClienteDebito
    '            titulomestre.EndDestinatario = ptitulomestre.EndClienteDebito

    '            If ptitulomestre.Banco = Nothing Then titulomestre.Banco = 0 Else titulomestre.Banco = ptitulomestre.Banco
    '            If ptitulomestre.Agencia = Nothing Then titulomestre.Agencia = 0 Else titulomestre.Agencia = ptitulomestre.Agencia

    '            titulomestre.DigitoAgencia = ptitulomestre.DigitoAgencia
    '            titulomestre.ContaCorrente = ptitulomestre.ContaCorrente
    '            titulomestre.DigitoConta = ptitulomestre.DigitoConta
    '            titulomestre.Destinacao = ptitulomestre.Destinacao

    '            titulomestre.Produto = "0"
    '            titulomestre.Safra = ""


    '            titulomestre.Moeda = ptitulomestre.Moeda


    '            titulomestre.RetencaoOficial = 0
    '            titulomestre.RetencaoMoeda = 0

    '            titulomestre.FaturaNumero = ""
    '            titulomestre.FaturaSerie = ""


    '            'titulomestre.EmpresaDebito = ""
    '            'titulomestre.EnderecoDebito = 0
    '            'titulomestre.ContaDebito = ""
    '            'titulomestre.ClienteDebito = ""
    '            'titulomestre.EndClienteDebito = 0
    '            'titulomestre.CustoDebito = 0
    '            'titulomestre.Provisao = 1


    '            '''''''''''''''''''''''''''''''
    '            '''      Fecha Conexao      '''
    '            ''''''''''''''''''''''''''''''' 
    '            conexao.FecheConexaoBanco()
    '        Catch ex As Exception
    '            MsgBox(Me.Page, "Geração titulo mestre-titulos")

    '        End Try

    '        Return titulomestre
    '    End Function



    '#End Region

End Class
