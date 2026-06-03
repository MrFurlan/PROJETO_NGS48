Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'*********************************************************************************************************************************************
'***************************************  List Base Pedidos x Encargos  *********************************************************************
'*********************************************************************************************************************************************
Public Class ListPedidoXEncargo
    Inherits List(Of PedidoXEncargo)

#Region "Construtor"
    Public Sub New(ByVal pItemPedido As PedidoXItem)
        _ItemPedido = pItemPedido
        If ItemPedido.IUD = "I" Then Me.CriaListar() Else Me.RecuperaLista()
    End Sub
#End Region

#Region "Fields"
    Private _ItemPedido As PedidoXItem
#End Region

#Region "Property"
    Public ReadOnly Property ItemPedido As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function CriaListar() As Boolean
        Me.Clear()
        Dim EncLiquido As PedidoXEncargo = Nothing
        Dim EncICMS As PedidoXEncargo = Nothing
        Dim EncICMSDesonerado As PedidoXEncargo = Nothing
        Dim EncIPI As PedidoXEncargo = Nothing
        'Dim EncCustoCofins As PedidoXEncargo = Nothing
        'Dim EncCustoPis As PedidoXEncargo = Nothing
        Dim dblLiquidoOficial As Decimal = 0
        Dim dblLiquidoMoeda As Decimal = 0

        Dim Parametros As New OperacaoXEstado
        Parametros.Empresa = Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8)
        Parametros.CodigoGrupoProduto = Me.ItemPedido.Produto.CodigoGrupo
        Parametros.CodigoProduto = Me.ItemPedido.CodigoProduto
        Parametros.CodigoOperacao = Me.ItemPedido.Pedido.CodigoOperacao
        Parametros.CodigoSubOperacao = Me.ItemPedido.Pedido.CodigoSubOperacao
        Parametros.EstadoOrigem = Me.ItemPedido.Pedido.Empresa.CodigoEstado
        Parametros.EstadoDestino = Me.ItemPedido.Pedido.Cliente.CodigoEstado

        If Me.ItemPedido.Pedido.Cliente.CodigoEstado <> "EX" Then
            Parametros.Retencao = Me.ItemPedido.Retencao
            Parametros.TipoDePessoa = IIf(Me.ItemPedido.Pedido.CodigoCliente.Length = 14, eTipoPessoa.Juridica, eTipoPessoa.Fisica)
            Parametros.EmpresaObriga = Me.ItemPedido.Pedido.Empresa.Empresa.ObrigaEncargo
        End If

        Dim OperacaoxEncargo As New OperacaoXEstado(Parametros)

        If OperacaoxEncargo Is Nothing OrElse OperacaoxEncargo.Codigo = 0 Then Return False

        Me.ItemPedido.CodigoOperacaoXEstado = OperacaoxEncargo.Codigo

        OperacaoxEncargo.Cliente = Me.ItemPedido.Pedido.Cliente

        Dim objProduto As PedidoXEncargo = Nothing
        OperacaoxEncargo.Encargos = Nothing
        'PEGANDO IPI ANTES DE PERCORRER TODOS ENCARGOS PARA SOMAR NO ICMS-ST - FURLAN - 07/08/2024
        For Each enc In OperacaoxEncargo.Encargos
            If enc.CodigoEncargo = "IPI" Then
                Dim objEncargo As New PedidoXEncargo(ItemPedido)

                objEncargo.CodigoEncargo = enc.CodigoEncargo
                objEncargo.BaseOficial = Math.Round(Me.ItemPedido.Lancamentos.TotalOficialPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
                objEncargo.BaseMoeda = Math.Round(Me.ItemPedido.Lancamentos.TotalMoedaPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
                objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (enc.Aliquota / 100), 2, MidpointRounding.AwayFromZero)

                EncIPI = objEncargo
            End If

            'If enc.CodigoEncargo = "CUSTO COFINS" Then
            '    Dim objEncargo As New PedidoXEncargo(ItemPedido)

            '    objEncargo.CodigoEncargo = enc.CodigoEncargo
            '    objEncargo.BaseOficial = Math.Round(Me.ItemPedido.Lancamentos.TotalOficialPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            '    objEncargo.BaseMoeda = Math.Round(Me.ItemPedido.Lancamentos.TotalMoedaPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            '    objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (enc.Aliquota / 100), 2, MidpointRounding.AwayFromZero)

            '    EncCustoCofins = objEncargo
            'End If

            'If enc.CodigoEncargo = "CUSTO PIS" Then
            '    Dim objEncargo As New PedidoXEncargo(ItemPedido)

            '    objEncargo.CodigoEncargo = enc.CodigoEncargo
            '    objEncargo.BaseOficial = Math.Round(Me.ItemPedido.Lancamentos.TotalOficialPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            '    objEncargo.BaseMoeda = Math.Round(Me.ItemPedido.Lancamentos.TotalMoedaPrd * (enc.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            '    objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (enc.Aliquota / 100), 2, MidpointRounding.AwayFromZero)

            '    EncCustoPis = objEncargo
            'End If

        Next

        For Each objOperacaoEncargo As OperacaoXEstadoXEncargo In OperacaoxEncargo.Encargos
            Dim objEncargo As New PedidoXEncargo(ItemPedido)

            objEncargo.CodigoEncargo = objOperacaoEncargo.CodigoEncargo

            If objEncargo.ValorPeso = eValorPeso.Peso Then
                objEncargo.BaseOficial = Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento
                objEncargo.BaseMoeda = Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento
            Else
                objEncargo.BaseOficial = Math.Round(Me.ItemPedido.Lancamentos.TotalOficialPrd * (objOperacaoEncargo.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
                objEncargo.BaseMoeda = Math.Round(Me.ItemPedido.Lancamentos.TotalMoedaPrd * (objOperacaoEncargo.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            End If

            Dim objEncargoXTaxa As New EncargoXTaxa()

            If objEncargo.CodigoEncargo = "ICMS" Then
                EncICMS = objEncargo
            ElseIf objEncargo.CodigoEncargo = "ICMS DESONERADO" Then
                EncICMSDesonerado = objEncargo
            ElseIf (Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "40938762" OrElse Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "49673784") AndAlso objEncargo.CodigoEncargo = "FUNDO FECP" Then
                objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Cliente.CodigoEstado, "ICMS-ST", Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, "")
            ElseIf (Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "40938762" OrElse Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "49673784") AndAlso objEncargo.CodigoEncargo = "ICMS-ST" Then
                objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Cliente.CodigoEstado, objEncargo.CodigoEncargo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, "")
            ElseIf Me.ItemPedido.Pedido.Empresa.CodigoEstado = "MT" AndAlso (objEncargo.CodigoEncargo = "FETHAB" Or objEncargo.CodigoEncargo = "IAGRO") Then
                objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Empresa.CodigoEstado, objEncargo.CodigoEncargo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, Me.ItemPedido.CodigoProduto)
            ElseIf objEncargo.CodigoEncargo = "FUNDERSUL" OrElse objEncargo.CodigoEncargo = "FUNDEMS" OrElse objEncargo.CodigoEncargo = "FUNDEMS" Then
                objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Empresa.CodigoEstado, "UFERMS", Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, Me.ItemPedido.CodigoProduto)
            Else
                objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Empresa.CodigoEstado, objEncargo.CodigoEncargo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento)
            End If

            If objEncargoXTaxa.Estado = Nothing Then
                objEncargo.Percentual = objOperacaoEncargo.Aliquota
            ElseIf (Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "40938762" OrElse Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "49673784") AndAlso (objEncargo.CodigoEncargo = "ICMS-ST" OrElse objEncargo.CodigoEncargo = "FUNDO FECP") Then
                objEncargo.Percentual = objOperacaoEncargo.Aliquota
            ElseIf objEncargo.CodigoEncargo = "FUNDERSUL" OrElse objEncargo.CodigoEncargo = "FUNDEMS" Then
                objEncargo.Percentual = objOperacaoEncargo.Aliquota
            Else : objEncargo.Percentual = objEncargoXTaxa.Percentual
            End If

            Select Case objEncargo.CodigoEncargo
                Case "PRODUTO"
                    objEncargo.ValorOficial = objEncargo.BaseOficial
                    objEncargo.ValorMoeda = objEncargo.BaseMoeda

                    dblLiquidoOficial += objEncargo.ValorOficial
                    dblLiquidoMoeda += objEncargo.ValorMoeda

                    objProduto = objEncargo

                Case "LIQUIDO"
                    EncLiquido = objEncargo
                Case Else
                    'If (Me.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS Or Me.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.COMPRASAORDEM) And Me.ItemPedido.Pedido.Cliente.CodigoCategoria > 3 And (objEncargo.CodigoEncargo.Contains("FACS") Or objEncargo.CodigoEncargo.Contains("FETHAB") Or objEncargo.CodigoEncargo.Contains("FABOV") Or objEncargo.CodigoEncargo.Contains("SENAR") Or objEncargo.CodigoEncargo.Contains("FUNRURAL")) Then
                    '    objEncargo.ValorOficial = 0
                    '    objEncargo.ValorMoeda = 0
                    'Else
                    If objEncargo.ValorPeso = eValorPeso.Peso Then

                        If Not objEncargoXTaxa.Estado Is Nothing AndAlso (objEncargo.CodigoEncargo = "FUNDERSUL" OrElse objEncargo.CodigoEncargo = "FUNDEMS" OrElse (Me.ItemPedido.Pedido.Empresa.CodigoEstado = "MT" AndAlso (objEncargo.CodigoEncargo = "FETHAB" Or objEncargo.CodigoEncargo = "IAGRO"))) Then

                            objEncargo.ValorOficial = Math.Round(((CDec(objEncargoXTaxa.SimplesNacional) * objEncargo.Percentual) / 100) * (Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento / 1000), 2, MidpointRounding.AwayFromZero)

                            If Me.ItemPedido.Pedido.Moeda.Classificacao <> eTiposMoeda.Oficial Then
                                If Me.ItemPedido.Pedido.CodigoIndexador = 99 OrElse Me.ItemPedido.Pedido.IndexadorFixo Then
                                    objEncargo.ValorMoeda = Math.Round(objEncargo.ValorOficial / Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    Dim vlrOficial As Decimal = Math.Round(objEncargo.ValorMoeda * Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    If vlrOficial <> objEncargo.ValorOficial Then objEncargo.ValorOficial = vlrOficial
                                Else
                                    objEncargo.ValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(objEncargo.Percentual * (Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento / 1000), Me.ItemPedido.Pedido.Indexador.Codigo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento), 2, MidpointRounding.AwayFromZero)
                                End If
                            End If

                        ElseIf objEncargo.Encargo.Operador = "%" Then
                            objEncargo.ValorOficial = Math.Round(objEncargo.Percentual * (Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento / 1000), 2, MidpointRounding.AwayFromZero)

                            If Me.ItemPedido.Pedido.Moeda.Classificacao <> eTiposMoeda.Oficial Then
                                If Me.ItemPedido.Pedido.CodigoIndexador = 99 OrElse Me.ItemPedido.Pedido.IndexadorFixo Then
                                    objEncargo.ValorMoeda = Math.Round(objEncargo.ValorOficial / Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    Dim vlrOficial As Decimal = Math.Round(objEncargo.ValorMoeda * Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    If vlrOficial <> objEncargo.ValorOficial Then objEncargo.ValorOficial = vlrOficial
                                Else
                                    objEncargo.ValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(objEncargo.Percentual * (Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento / 1000), Me.ItemPedido.Pedido.Indexador.Codigo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento), 2, MidpointRounding.AwayFromZero)
                                End If
                            End If
                        ElseIf objEncargo.Encargo.Operador = "*" Then
                            objEncargo.ValorOficial = Math.Round(objEncargo.Percentual * Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento, 2, MidpointRounding.AwayFromZero)

                            If Me.ItemPedido.Pedido.Moeda.Classificacao <> eTiposMoeda.Oficial Then
                                If Me.ItemPedido.Pedido.CodigoIndexador = 99 OrElse Me.ItemPedido.Pedido.IndexadorFixo Then
                                    objEncargo.ValorMoeda = Math.Round(objEncargo.ValorOficial / Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    Dim vlrOficial As Decimal = Math.Round(objEncargo.ValorMoeda * Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                    If vlrOficial <> objEncargo.ValorOficial Then objEncargo.ValorOficial = vlrOficial
                                Else
                                    objEncargo.ValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(objEncargo.Percentual * Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento, Me.ItemPedido.Pedido.Indexador.Codigo, Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento), 2, MidpointRounding.AwayFromZero)
                                End If
                            End If
                        End If
                    Else
                        If (Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "40938762" OrElse Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "49673784") AndAlso objEncargo.CodigoEncargo = "FUNDO FECP" Then

                            If objProduto.ValorOficial > 0 Then
                                Dim vst As Decimal = Math.Round((objProduto.ValorOficial * objEncargoXTaxa.SimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                                objEncargo.BaseOficial = objProduto.ValorOficial + vst

                                objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                            Else
                                objEncargo.BaseOficial = 0
                                objEncargo.ValorOficial = 0
                            End If

                        ElseIf (Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "40938762" OrElse Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "49673784") AndAlso objEncargo.CodigoEncargo = "ICMS-ST" Then

                            If objProduto.ValorOficial > 0 Then

                                Dim vst As Decimal = 0

                                If Me.ItemPedido.Pedido.Cliente.CodigoEstado = "MG" Then
                                    objEncargoXTaxa.SelecionarVigente("MG", "ICMS ST MG", Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, Me.ItemPedido.CodigoProduto)

                                    If Me.ItemPedido.UnidadeComercializacaoFatorDeConversao > 0 Then
                                        vst = Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento * Me.ItemPedido.UnidadeComercializacaoFatorDeConversao
                                    Else
                                        vst = Me.ItemPedido.Lancamentos.QuantidadeTotalPrdFaturamento
                                    End If

                                    vst = Math.Round((vst * objEncargoXTaxa.SimplesNacional), 2, MidpointRounding.AwayFromZero)

                                    objEncargo.BaseOficial = vst
                                Else
                                    vst = objProduto.ValorOficial

                                    If Not EncIPI Is Nothing AndAlso EncIPI.Valor > 0 Then
                                        vst += EncIPI.Valor
                                    End If

                                    vst = Math.Round((vst * objEncargoXTaxa.SimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                                    If Not EncIPI Is Nothing AndAlso EncIPI.Valor > 0 Then
                                        objEncargo.BaseOficial = objProduto.ValorOficial + EncIPI.Valor + vst
                                    Else
                                        objEncargo.BaseOficial = objProduto.ValorOficial + vst
                                    End If
                                End If

                                Dim vlrST As Decimal = Math.Round((objEncargo.BaseOficial * objEncargoXTaxa.Percentual) / 100, 2, MidpointRounding.AwayFromZero)

                                Dim perST As Decimal = 0
                                If Not EncICMS Is Nothing AndAlso EncICMS.ValorOficial > 0 Then
                                    perST = Math.Round(((vlrST - EncICMS.ValorOficial) / objEncargo.BaseOficial) * 100, 2, MidpointRounding.AwayFromZero)
                                Else
                                    perST = Math.Round(((objProduto.ValorOficial - vlrST) / objEncargo.BaseOficial) * 100, 2, MidpointRounding.AwayFromZero)
                                End If

                                objEncargo.Percentual = perST

                                If Not EncICMS Is Nothing AndAlso EncICMS.ValorOficial > 0 Then
                                    objEncargo.ValorOficial = vlrST - EncICMS.ValorOficial
                                Else
                                    objEncargo.ValorOficial = vlrST - EncICMSDesonerado.Valor
                                End If
                            Else
                                objEncargo.BaseOficial = 0
                                objEncargo.ValorOficial = 0
                            End If
                        ElseIf objEncargo.CodigoEncargo = "ICMS DIFERENCIAL" Then
                            objEncargo.ValorOficial = Math.Round((objEncargo.BaseOficial * (objEncargo.Percentual / 100)) - EncICMS.ValorOficial, 2, MidpointRounding.AwayFromZero)

                            'ElseIf objEncargo.CodigoEncargo = "ICMS" Then
                            '    If Not EncCustoCofins Is Nothing AndAlso EncCustoCofins.Valor > 0 Then
                            '        objEncargo.BaseOficial += EncCustoCofins.Valor
                            '    End If

                            '    If Not EncCustoPis Is Nothing AndAlso EncCustoPis.Valor > 0 Then
                            '        objEncargo.BaseOficial += EncCustoPis.Valor
                            '    End If

                            '    objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)

                        ElseIf (objEncargo.CodigoEncargo = "CBS" Or objEncargo.CodigoEncargo = "IBS") AndAlso OperacaoxEncargo.ReducaoCBS_Perc > 0 Then

                            Dim percentual = objOperacaoEncargo.AliquotaBase - OperacaoxEncargo.ReducaoCBS_Perc
                            objEncargo.Percentual = Math.Round((objOperacaoEncargo.Aliquota * percentual) / 100, 2, MidpointRounding.AwayFromZero)
                            objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)

                        Else
                            objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                        End If

                        If Me.ItemPedido.Pedido.Moeda.Classificacao <> eTiposMoeda.Oficial Then
                            If Me.ItemPedido.Pedido.CodigoIndexador = 99 OrElse Me.ItemPedido.Pedido.IndexadorFixo Then
                                objEncargo.ValorMoeda = Math.Round(objEncargo.ValorOficial / Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                            Else
                                objEncargo.ValorMoeda = Math.Round(objEncargo.BaseMoeda * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                    'End If

                    If objOperacaoEncargo.Sinal = "-" Then
                        dblLiquidoOficial -= objEncargo.ValorOficial
                        dblLiquidoMoeda -= objEncargo.ValorMoeda
                    ElseIf objOperacaoEncargo.Sinal = "+" Then
                        dblLiquidoOficial += objEncargo.ValorOficial
                        dblLiquidoMoeda += objEncargo.ValorMoeda
                    End If
            End Select

            Me.Add(objEncargo)
        Next


        'VALOR DA BASE DE CÁLCULO DO FUNDO FECP DEVE SER IGUAL AO ICMS ST - FURLAN - 17/01/2025.
        Dim BaseST As Decimal = 0
        Dim temFUNDOFECP As Boolean = False
        For Each enc In Me
            If enc.CodigoEncargo.Contains("ICMS-ST") Then BaseST = enc.BaseOficial
            If enc.CodigoEncargo.Contains("FUNDO FECP") Then temFUNDOFECP = True
        Next

        If temFUNDOFECP AndAlso BaseST > 0 Then
            For Each enc In Me

                If enc.CodigoEncargo.Contains("FUNDO FECP") Then
                    'REMOVE VALOR ANTERIOR
                    dblLiquidoOficial -= enc.ValorOficial
                    dblLiquidoMoeda -= enc.ValorMoeda

                    'RECALCULA
                    enc.BaseOficial = BaseST
                    enc.ValorOficial = Math.Round(enc.BaseOficial * (enc.Percentual / 100), 2, MidpointRounding.AwayFromZero)

                    If Me.ItemPedido.Pedido.Moeda.Classificacao <> eTiposMoeda.Oficial Then
                        If Me.ItemPedido.Pedido.CodigoIndexador = 99 OrElse Me.ItemPedido.Pedido.IndexadorFixo Then
                            enc.ValorMoeda = Math.Round(enc.ValorOficial / Me.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                        Else
                            enc.ValorMoeda = Math.Round(enc.BaseMoeda * (enc.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                        End If
                    End If

                    'ACRESCENTA VALOR RECALCULADO
                    dblLiquidoOficial += enc.ValorOficial
                    dblLiquidoMoeda += enc.ValorMoeda

                End If
            Next
        End If

        EncLiquido.ValorMoeda = dblLiquidoMoeda
        EncLiquido.ValorOficial = dblLiquidoOficial

        Return True
    End Function

    Public Function RecuperaLista() As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String = ""
            sql = "SELECT Encargo_Id, BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda " & vbCrLf & _
                  "  FROM PedidosXEncargos " & vbCrLf & _
                  " WHERE Empresa_Id    ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                  "   AND Pedido_Id     = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                  "   AND Produto_Id    ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                  " Order by Case" & vbCrLf & _
                  "            When Encargo_id = 'PRODUTO' then 1" & vbCrLf & _
                  "            When Encargo_id = 'LIQUIDO' then 3" & vbCrLf & _
                  "            Else 2" & vbCrLf & _
                  "          end, Encargo_Id" & vbCrLf

            Dim dsEncargos As DataSet = objBanco.ConsultaDataSet(sql, "PE")

            Dim objLiquido As PedidoXEncargo = Nothing
            Dim objProduto As PedidoXEncargo = Nothing

            For Each drEncargo As DataRow In dsEncargos.Tables(0).Rows
                Dim objEncargo As New PedidoXEncargo(Me.ItemPedido)

                objEncargo.CodigoEncargo = drEncargo("Encargo_Id")
                objEncargo.BaseOficial = drEncargo("BaseOficial")
                objEncargo.BaseMoeda = drEncargo("BaseMoeda")
                objEncargo.Percentual = drEncargo("Percentual")
                objEncargo.ValorOficial = drEncargo("ValorOficial")
                objEncargo.ValorMoeda = drEncargo("ValorMoeda")

                Select Case objEncargo.CodigoEncargo
                    Case "LIQUIDO" : objLiquido = objEncargo
                    Case "PRODUTO" : objProduto = objEncargo
                End Select
                Me.Add(objEncargo)
            Next

            Return True
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        If Me.ItemPedido.IUD = "" Then Exit Sub

        If Me.ItemPedido.IUD = "D" Or Me.ItemPedido.IUD = "U" Then
            Sqls.Add("DELETE PedidosXEncargos " & _
                     " WHERE Empresa_Id    ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                     "   AND EndEmpresa_Id = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                     "   AND Pedido_Id     = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                     "   AND Produto_Id    ='" & Me.ItemPedido.CodigoProduto & "';")
        End If

        If Me.ItemPedido.IUD <> "D" Then
            For Each enc As PedidoXEncargo In Me
                If ((Me.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial And enc.ValorOficial > 0) Or (Me.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira And enc.ValorMoeda > 0)) Or _
                   enc.CodigoEncargo.ToString.Contains("ICMS") Or enc.CodigoEncargo = "PRODUTO" Or enc.CodigoEncargo = "LIQUIDO" Then
                    enc.IUD = "I"
                    enc.SalvarSql(Sqls)
                End If
            Next
        End If
    End Sub

    Public Sub AtualizaLiquido()
        Dim dblLiquidoOficial As Decimal = 0
        Dim dblLiquidoMoeda As Decimal = 0
        Dim EncLiquido As PedidoXEncargo = Nothing
        Dim EncProduto As PedidoXEncargo = Nothing

        Dim EncICMS As PedidoXEncargo = Nothing
        Dim EncADUANEIRAS As PedidoXEncargo = Nothing
        Dim EncFrete As PedidoXEncargo = Nothing
        Dim EncCustoCofins As PedidoXEncargo = Nothing
        Dim EncCustoPis As PedidoXEncargo = Nothing

        Dim objEncargoXTaxa As New EncargoXTaxa()

        If Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8) = "24450490" AndAlso Me.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.IMPORTACOES AndAlso Me.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then

            For Each objEncargo As PedidoXEncargo In Me
                If objEncargo.CodigoEncargo = "PRODUTO" Then EncProduto = objEncargo
                If objEncargo.CodigoEncargo = "DESP.ADUANEIRAS" Then EncADUANEIRAS = objEncargo
                If objEncargo.CodigoEncargo = "FRETES" Then EncFrete = objEncargo
                If objEncargo.CodigoEncargo = "CUSTO COFINS" Then EncCustoCofins = objEncargo
                If objEncargo.CodigoEncargo = "CUSTO PIS" Then EncCustoPis = objEncargo
            Next

            For Each objEncargo As PedidoXEncargo In Me
                If objEncargo.CodigoEncargo = "ICMS" Then
                    objEncargoXTaxa.SelecionarVigente(Me.ItemPedido.Pedido.Empresa.CodigoEstado, "ICMS", Me.ItemPedido.Lancamentos.LancamentoNormal.Movimento, "")

                    Dim vTotal As Decimal = EncProduto.Valor + EncADUANEIRAS.Valor + EncFrete.Valor

                    If Not EncCustoCofins Is Nothing AndAlso EncCustoCofins.Valor > 0 Then
                        vTotal += EncCustoCofins.Valor
                    End If

                    If Not EncCustoPis Is Nothing AndAlso EncCustoPis.Valor > 0 Then
                        vTotal += EncCustoPis.Valor
                    End If

                    Dim vICMS As Decimal = Math.Round(((vTotal / objEncargoXTaxa.Percentual) * objEncargoXTaxa.SimplesNacional) / 100, 2, MidpointRounding.AwayFromZero)

                    objEncargo.ValorOficial = vICMS

                    objEncargo.BaseOficial = Math.Round((vICMS * 100) / objEncargo.Percentual, 2, MidpointRounding.AwayFromZero)

                    EncICMS = objEncargo
                End If
            Next

            For Each objEncargo As PedidoXEncargo In Me
                If objEncargo.CodigoEncargo = "CUSTO ICMS" Then
                    objEncargo.ValorOficial = EncICMS.ValorOficial
                ElseIf objEncargo.CodigoEncargo = "CUSTO COFINS" OrElse
                    objEncargo.CodigoEncargo = "COFINS" OrElse
                    objEncargo.CodigoEncargo = "PIS" OrElse
                    objEncargo.CodigoEncargo = "CUSTO PIS" Then
                    objEncargo.BaseOficial = EncProduto.Valor + EncFrete.Valor
                    objEncargo.ValorOficial = Math.Round(objEncargo.BaseOficial * (objEncargo.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                End If
            Next
        End If

        For Each objEncargo As PedidoXEncargo In Me
            If objEncargo.CodigoEncargo = "LIQUIDO" Then EncLiquido = objEncargo

            If Me.ItemPedido.Pedido.SubOperacao.Classe = eClassesOperacoes.COMPRAS And Me.ItemPedido.Pedido.Cliente.CodigoCategoria > 3 And (objEncargo.CodigoEncargo.Contains("FACS") Or objEncargo.CodigoEncargo.Contains("FETHAB") Or objEncargo.CodigoEncargo.Contains("SENAR") Or objEncargo.CodigoEncargo.Contains("FUNRURAL")) Then
                objEncargo.ValorOficial = 0
                objEncargo.ValorMoeda = 0
                objEncargo.Percentual = 0
            Else
                If objEncargo.Base = 0 Then
                    objEncargo.Percentual = 0
                ElseIf objEncargo.ValorPeso = eValorPeso.Peso Then
                    If objEncargo.Encargo.Operador = "%" Then
                        objEncargo.Percentual = Math.Round((objEncargo.Valor * 1000) / objEncargo.Base, 9, MidpointRounding.AwayFromZero)
                    ElseIf objEncargo.Encargo.Operador = "*" Then
                        objEncargo.Percentual = Math.Round(objEncargo.Base / objEncargo.Valor, 9, MidpointRounding.AwayFromZero)
                    End If
                Else
                    If objEncargo.CodigoEncargo = "ICMS" OrElse objEncargo.CodigoEncargo = "CUSTO COFINS" OrElse objEncargo.CodigoEncargo = "CUSTO PIS" OrElse objEncargo.CodigoEncargo = "CBS" OrElse objEncargo.CodigoEncargo = "IBS" Then
                        'NÃO FAZ NADA - FURLAN - 25/09/2024
                    Else
                        objEncargo.Percentual = Math.Round((objEncargo.Valor * 100) / objEncargo.Base, 9, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If

            Select Case objEncargo.Sinal
                Case "+"
                    dblLiquidoOficial += Math.Round(objEncargo.ValorOficial, 2, MidpointRounding.AwayFromZero)
                    dblLiquidoMoeda += Math.Round(objEncargo.ValorMoeda, 2, MidpointRounding.AwayFromZero)
                Case "-"
                    dblLiquidoOficial -= Math.Round(objEncargo.ValorOficial, 2, MidpointRounding.AwayFromZero)
                    dblLiquidoMoeda -= Math.Round(objEncargo.ValorMoeda, 2, MidpointRounding.AwayFromZero)
            End Select
        Next

        EncLiquido.BaseOficial = dblLiquidoOficial
        EncLiquido.Percentual = 100
        EncLiquido.ValorOficial = dblLiquidoOficial

        EncLiquido.BaseMoeda = dblLiquidoMoeda
        EncLiquido.Percentual = 100
        EncLiquido.ValorMoeda = dblLiquidoMoeda
    End Sub
#End Region

End Class

'*********************************************************************************************************************************************
'**************************************  Classe Base Pedidos x Encargos  *********************************************************************
'*********************************************************************************************************************************************
Public Class PedidoXEncargo

#Region "Contrutor"
    Public Sub New(ByVal pItemPedido As PedidoXItem)
        _ItemPedido = pItemPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _ItemPedido As PedidoXItem
    Private _CodigoEncargo As String
    Private _Encargo As Encargo
    Private _BaseOficial As Decimal
    Private _BaseMoeda As Decimal
    Private _Percentual As Decimal
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _Quantidade As Integer
#End Region

#Region "Property"
    Public ReadOnly Property ItemPedido() As PedidoXItem
        Get
            Return _ItemPedido
        End Get
    End Property

    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoEncargo() As String
        Get
            Return _CodigoEncargo
        End Get
        Set(ByVal value As String)
            _CodigoEncargo = value
            _Encargo = Nothing
        End Set
    End Property

    Public ReadOnly Property Encargo As Encargo
        Get
            If _Encargo Is Nothing And Me.CodigoEncargo.Length > 0 Then _Encargo = New Encargo(Me.CodigoEncargo)
            Return _Encargo
        End Get
    End Property

    Public Property BaseOficial() As Decimal
        Get
            Return _BaseOficial
        End Get
        Set(ByVal value As Decimal)
            _BaseOficial = value
        End Set
    End Property

    Public Property BaseMoeda() As Decimal
        Get
            Return _BaseMoeda
        End Get
        Set(ByVal value As Decimal)
            _BaseMoeda = value
        End Set
    End Property

    Public Property Percentual() As Decimal
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _Percentual = value
        End Set
    End Property

    Public Property ValorOficial() As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(ByVal value As Decimal)
            _ValorOficial = value
        End Set
    End Property

    Public Property ValorMoeda() As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorMoeda = value
        End Set
    End Property

    Public ReadOnly Property ValorPeso() As eValorPeso
        Get
            Return Encargo.ValorOuPeso
        End Get
    End Property

    Public ReadOnly Property Sinal As String
        Get
            Return ItemPedido.OperacaoxEstado.Encargos.Where(Function(s) s.CodigoEncargo = Me.CodigoEncargo).Single.Sinal
        End Get
    End Property

    'Base e Valor que seguem a moeda do pedido
    Public ReadOnly Property Base As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial OrElse ItemPedido.AjustarValorOficial Then
                Return Me.BaseOficial
            Else
                Return Me.BaseMoeda
            End If
        End Get
    End Property

    Public ReadOnly Property Valor As Decimal
        Get
            If ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial OrElse ItemPedido.AjustarValorOficial Then
                Return Me.ValorOficial
            Else
                Return Me.ValorMoeda
            End If
        End Get
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = "INSERT INTO PedidosXEncargos " & vbCrLf & _
                      "(Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                      " Pedido_Id, Produto_Id," & vbCrLf & _
                      " Encargo_Id, " & vbCrLf & _
                      " BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda) " & vbCrLf & _
                      "VALUES ('" & Me.ItemPedido.Pedido.CodigoEmpresa & "'," & Me.ItemPedido.Pedido.EnderecoEmpresa & "," & vbCrLf & _
                      Me.ItemPedido.Pedido.Codigo & ", '" & Me.ItemPedido.CodigoProduto & "'," & vbCrLf & _
                      "'" & Me.CodigoEncargo & "'," & vbCrLf & _
                      Str(Me.BaseOficial) & "," & vbCrLf & _
                      Str(Me.BaseMoeda) & "," & vbCrLf & _
                      Str(Me.Percentual) & "," & vbCrLf & _
                      Str(Me.ValorOficial) & "," & vbCrLf & _
                      Str(Me.ValorMoeda) & ")" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class

