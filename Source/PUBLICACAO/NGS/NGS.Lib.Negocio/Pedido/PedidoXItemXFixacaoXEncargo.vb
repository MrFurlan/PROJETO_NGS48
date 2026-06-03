Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'***********************************************************************************************************************************
'*************************************  LISTA CLASSE FIXACAO x ENCAGO  *************************************************************
'***********************************************************************************************************************************
Public Class ListFixacaoXEncargo
    Inherits List(Of FixacaoXEncargo)

#Region "Contrutor"
    Public Sub New(ByVal pFixacao As Fixacao)
        Me.FixacaoItemPedido = pFixacao
        If FixacaoItemPedido.IUD = "I" Then
            Listar()
            Exit Sub
        End If

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, Fixacao_Id, Encargo_Id, " & vbCrLf & _
                  "       BaseOficial, BaseMoeda, Percentual, ValorOficial, ValorMoeda " & vbCrLf & _
                  "  FROM VW_PedidosXItensXFixacoesXEncargos " & vbCrLf & _
                  " WHERE Empresa_Id    ='" & Me.FixacaoItemPedido.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & Me.FixacaoItemPedido.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                  "   AND Pedido_Id     = " & Me.FixacaoItemPedido.ItemPedido.Pedido.Codigo & vbCrLf & _
                  "   AND Produto_Id    = " & Me.FixacaoItemPedido.ItemPedido.CodigoProduto & vbCrLf & _
                  "   AND Fixacao_Id    = " & Me.FixacaoItemPedido.Codigo & vbCrLf

            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(sql, "PedidosXItensXFixacoes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objEncargo As New FixacaoXEncargo(Me.FixacaoItemPedido)

                objEncargo.CodigoEncargo = row("Encargo_Id").ToString()
                objEncargo.BaseOficial = row("BaseOficial")
                objEncargo.BaseMoeda = row("BaseMoeda")
                objEncargo.Percentual = row("Percentual")
                objEncargo.ValorOficial = row("ValorOficial")
                objEncargo.ValorMoeda = row("ValorMoeda")

                Me.Add(objEncargo)
            Next
        Catch ex As Exception
        Finally
            Banco = Nothing
        End Try

    End Sub
#End Region

#Region "Fields"
    Private _FixacaoItemPedido As Fixacao
#End Region

#Region "Property"
    Public Property FixacaoItemPedido() As Fixacao
        Get
            Return _FixacaoItemPedido
        End Get
        Set(ByVal value As Fixacao)
            _FixacaoItemPedido = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub AdicionarDadosEncargos(ByVal ItemPedido As PedidoXItem)
        For Each objEncargo As PedidoXEncargo In ItemPedido.Encargos
            Dim objEncargoFixacao As New FixacaoXEncargo(Me.FixacaoItemPedido)

            objEncargoFixacao.IUD = "I"
            objEncargoFixacao.CodigoEncargo = objEncargo.CodigoEncargo
            objEncargoFixacao.BaseOficial = objEncargo.BaseOficial
            objEncargoFixacao.BaseMoeda = objEncargo.BaseMoeda
            objEncargoFixacao.Percentual = objEncargo.Percentual
            objEncargoFixacao.ValorOficial = objEncargo.ValorOficial
            objEncargoFixacao.ValorMoeda = objEncargo.ValorMoeda

            Me.Add(objEncargoFixacao)
        Next
    End Sub

    Public Function Listar() As Boolean
        Me.Clear()
        Dim EncLiquido As FixacaoXEncargo = Nothing
        'Dim EncAFixar As FixacaoXEncargo = Nothing
        Dim dblLiquidoOficial As Decimal = 0
        Dim dblLiquidoMoeda As Decimal = 0

        If Me.FixacaoItemPedido.OperacaoxEstado.Codigo = 0 Then Return False

        If Me.FixacaoItemPedido.ItemPedido.Pedido.Cliente.CodigoEstado <> "EX" Then
            If Me.FixacaoItemPedido.ItemPedido.Pedido.CodigoCliente.Length = 14 Then
                Me.FixacaoItemPedido.OperacaoxEstado.TipoDePessoa = eTipoPessoa.Juridica
            Else
                Me.FixacaoItemPedido.OperacaoxEstado.TipoDePessoa = eTipoPessoa.Fisica
            End If
        End If
        Me.FixacaoItemPedido.OperacaoxEstado.Retencao = Me.FixacaoItemPedido.ItemPedido.Retencao
        Me.FixacaoItemPedido.OperacaoxEstado.EmpresaObriga = Me.FixacaoItemPedido.ItemPedido.Pedido.Empresa.Empresa.ObrigaEncargo

        If Me.FixacaoItemPedido.OperacaoxEstado.Encargos.Count = 0 Then Return False

        For Each OpxEn As OperacaoXEstadoXEncargo In Me.FixacaoItemPedido.OperacaoxEstado.Encargos
            Dim EncFixacao As New FixacaoXEncargo(FixacaoItemPedido)

            EncFixacao.CodigoEncargo = OpxEn.CodigoEncargo

            If EncFixacao.ValorPeso = eValorPeso.Peso Then
                EncFixacao.BaseOficial = Me.FixacaoItemPedido.Quantidade
                EncFixacao.BaseMoeda = Me.FixacaoItemPedido.Quantidade
            Else
                EncFixacao.BaseOficial = Math.Round(Me.FixacaoItemPedido.TotalOficial * (OpxEn.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
                EncFixacao.BaseMoeda = Math.Round(Me.FixacaoItemPedido.TotalMoeda * (OpxEn.AliquotaBase / 100), 2, MidpointRounding.AwayFromZero)
            End If

            Dim objEncargoXTaxa As New EncargoXTaxa()
            objEncargoXTaxa.SelecionarVigente(Me.FixacaoItemPedido.ItemPedido.Pedido.Empresa.CodigoEstado, EncFixacao.CodigoEncargo, Me.FixacaoItemPedido.Movimento)

            If objEncargoXTaxa.Estado = Nothing Then
                EncFixacao.Percentual = OpxEn.Aliquota
            Else
                EncFixacao.Percentual = objEncargoXTaxa.Percentual
            End If

            Select Case EncFixacao.CodigoEncargo
                Case "PRODUTO"
                    EncFixacao.ValorOficial = EncFixacao.BaseOficial
                    EncFixacao.ValorMoeda = EncFixacao.BaseMoeda

                    dblLiquidoOficial += EncFixacao.ValorOficial
                    dblLiquidoMoeda += EncFixacao.ValorMoeda
                Case "LIQUIDO"
                    EncLiquido = EncFixacao
                Case "AFIXAR"
                    Me.FixacaoItemPedido.AFixarDebito = OpxEn.CodigoDebitaConta
                    Me.FixacaoItemPedido.AFixarCredito = OpxEn.CodigoCreditaConta
                    'Case "AFIXAR"
                    '    EncAFixar.BaseOficial = FixacaoItemPedido.Quantidade
                    '    EncAFixar.BaseMoeda = FixacaoItemPedido.Quantidade
                    '    EncAFixar.ValorMoeda = FixacaoItemPedido.TotalMoeda
                    '    EncAFixar.ValorOficial = FixacaoItemPedido.TotalOficial
                Case Else
                    If Not Me.FixacaoItemPedido.ItemPedido.Retencao And _
                        EncFixacao.Encargo.PodeSofreRetencao And _
                        (Me.FixacaoItemPedido.ItemPedido.Pedido.Cliente.CodigoEstado = "EX" Or EncFixacao.Encargo.TipoPessoaRetencao = eTipoPessoa.Ambas Or EncFixacao.Encargo.TipoPessoaRetencao = IIf(Me.FixacaoItemPedido.ItemPedido.Pedido.CodigoCliente.Length = 14, eTipoPessoa.Juridica, eTipoPessoa.Fisica)) Then
                        EncFixacao.ValorOficial = 0
                        EncFixacao.ValorMoeda = 0
                    Else
                        If EncFixacao.ValorPeso = eValorPeso.Peso Then
                            EncFixacao.ValorOficial = Math.Round(EncFixacao.Percentual * (Me.FixacaoItemPedido.Quantidade / 1000), 2, MidpointRounding.AwayFromZero)

                            If Me.FixacaoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                EncFixacao.ValorMoeda = 0
                            Else
                                If Me.FixacaoItemPedido.IndiceFixado > 0 Then
                                    EncFixacao.ValorMoeda = Math.Round(EncFixacao.ValorOficial / Me.FixacaoItemPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                ElseIf Me.FixacaoItemPedido.ItemPedido.Pedido.IndiceFixado > 0 Then
                                    EncFixacao.ValorMoeda = Math.Round(EncFixacao.ValorOficial / Me.FixacaoItemPedido.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                Else
                                    EncFixacao.ValorMoeda = Math.Round(Funcoes.ConverteParaMoedaExtrangeira(EncFixacao.Percentual * (Me.FixacaoItemPedido.Quantidade / 1000), Me.FixacaoItemPedido.ItemPedido.Pedido.Indexador.Codigo, Me.FixacaoItemPedido.Movimento), 2, MidpointRounding.AwayFromZero)
                                End If
                            End If
                        Else
                            EncFixacao.ValorOficial = Math.Round(EncFixacao.BaseOficial * (EncFixacao.Percentual / 100), 2, MidpointRounding.AwayFromZero)

                            If Me.FixacaoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                                EncFixacao.ValorMoeda = 0
                            Else
                                If Me.FixacaoItemPedido.IndiceFixado > 0 Then
                                    EncFixacao.ValorMoeda = Math.Round(EncFixacao.ValorOficial / Me.FixacaoItemPedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                ElseIf Me.FixacaoItemPedido.ItemPedido.Pedido.IndiceFixado > 0 Then
                                    EncFixacao.ValorMoeda = Math.Round(EncFixacao.ValorOficial / Me.FixacaoItemPedido.ItemPedido.Pedido.IndiceFixado, 2, MidpointRounding.AwayFromZero)
                                Else
                                    EncFixacao.ValorMoeda = Math.Round(EncFixacao.BaseMoeda * (EncFixacao.Percentual / 100), 2, MidpointRounding.AwayFromZero)
                                End If
                            End If

                        End If
                    End If

                    If OpxEn.Sinal = "-" Then
                        dblLiquidoOficial -= EncFixacao.ValorOficial
                        dblLiquidoMoeda -= EncFixacao.ValorMoeda
                    ElseIf OpxEn.Sinal = "+" Then
                        dblLiquidoOficial += EncFixacao.ValorOficial
                        dblLiquidoMoeda += EncFixacao.ValorMoeda
                    End If
            End Select

            If EncFixacao.ValorMoeda > 0 Or EncFixacao.ValorOficial > 0 Then Me.Add(EncFixacao)
        Next
        EncLiquido.ValorMoeda = dblLiquidoMoeda
        EncLiquido.ValorOficial = dblLiquidoOficial
        'If Not EncAFixar Is Nothing Then Me.Add(EncAFixar)
        Me.Add(EncLiquido)
        Return True
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)

        'Retirado pois ao fazer um complemento ou estorno no item do pedido
        'os encargos da fixação estavam sendo deletados. Cleberson 04/05/2016
        'If FixacaoItemPedido.IUD = "D" Or FixacaoItemPedido.IUD = "U" Then

        If FixacaoItemPedido.IUD = "D" Then
            Dim sql As String = ""
            sql = "DELETE PedidosXItensXFixacoesXEncargos " & vbCrLf & _
                  " WHERE Empresa_Id    ='" & Me.FixacaoItemPedido.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & Me.FixacaoItemPedido.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                  "   AND Pedido_Id     = " & Me.FixacaoItemPedido.ItemPedido.Pedido.Codigo & vbCrLf & _
                  "   AND Produto_Id    = " & Me.FixacaoItemPedido.ItemPedido.CodigoProduto & vbCrLf & _
                  "   AND Fixacao_Id    = " & Me.FixacaoItemPedido.Codigo & vbCrLf
            Sqls.Add(sql)
            If FixacaoItemPedido.IUD = "D" Then Exit Sub
        End If

        For Each enc As FixacaoXEncargo In Me
            If FixacaoItemPedido.IUD = "I" Then enc.IUD = FixacaoItemPedido.IUD
            'Financeiro Novo
            If FixacaoItemPedido.ItemPedido IsNot Nothing _
                AndAlso FixacaoItemPedido.ItemPedido.Pedido IsNot Nothing _
                AndAlso FixacaoItemPedido.ItemPedido.Pedido.FinanceiroNovo AndAlso _
                FixacaoItemPedido.IUD = "U" Then
                enc.IUD = "I"
            End If
            If (Me.FixacaoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial And enc.ValorOficial > 0) Or _
               (Me.FixacaoItemPedido.ItemPedido.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira And enc.ValorMoeda > 0) Or _
               (enc.CodigoEncargo.Contains("ICMS")) Then
                enc.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'***********************************************************************************************************************************
'*************************************  CLASSE BASE FIXACAO x ENCAGO  **************************************************************
'***********************************************************************************************************************************
Public Class FixacaoXEncargo

#Region "Contrutor"
    Public Sub New(ByVal pFixacao As Fixacao)
        Me.FixacaoPedido = pFixacao
    End Sub
#End Region

#Region "Fields"
    Private _FixacaoPedido As Fixacao
    Private _IUD As String
    Private _CodigoEncargo As String
    Private _BaseOficial As Decimal
    Private _BaseMoeda As Decimal
    Private _Percentual As Decimal
    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _Encargo As Encargo
#End Region

#Region "Property"
    Public Property FixacaoPedido() As Fixacao
        Get
            Return _FixacaoPedido
        End Get
        Set(ByVal value As Fixacao)
            _FixacaoPedido = value
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
            Return Me.Encargo.ValorOuPeso
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "INSERT INTO PedidosXItensXFixacoesXEncargos " & vbCrLf & _
                      "(Empresa_Id, EndEmpresa_Id, Pedido_Id," & vbCrLf & _
                      " Produto_Id, Fixacao_Id, Encargo_Id, BaseOficial, " & vbCrLf & _
                      " BaseMoeda, Percentual, ValorOficial, ValorMoeda) " & vbCrLf & _
                      "Values ('" & Me.FixacaoPedido.ItemPedido.Pedido.CodigoEmpresa & "'," & Me.FixacaoPedido.ItemPedido.Pedido.EnderecoEmpresa & ", " & Me.FixacaoPedido.ItemPedido.Pedido.Codigo & "," & vbCrLf & _
                      "'" & Me.FixacaoPedido.ItemPedido.CodigoProduto & "'," & Me.FixacaoPedido.Codigo & ",'" & Me.CodigoEncargo & "'," & Str(Me.BaseOficial) & "," & vbCrLf & _
                      "" & Str(Me.BaseMoeda) & "," & Str(Me.Percentual) & "," & Str(Me.ValorOficial) & "," & Str(Me.ValorMoeda) & ")"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class