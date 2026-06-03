Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

'******************************************************************************************************************************************************
'**************************************************    LISTA DE PARCELAS DO PEDIDO   ******************************************************************
'******************************************************************************************************************************************************
<Serializable()> _
Public Class ListPedidoXParcela
    Inherits List(Of PedidoXParcela)
#Region "Construtor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
        If Me.Pedido.Codigo = 0 Then Exit Sub

        Dim Banco As New AcessaBanco()
        Try
            Dim strSQL As String = ""
            strSQL = "SELECT Parcela_Id," & vbCrLf & _
                     "       Vencimento," & vbCrLf & _
                     "       Valor" & vbCrLf & _
                     "  FROM PedidoxParcela " & vbCrLf & _
                     " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                     "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                     "   AND Pedido_Id     = " & Me.Pedido.Codigo & vbCrLf & _
                     " Order by Parcela_id" & vbCrLf

            Dim ds As DataSet
            ds = Banco.ConsultaDataSet(strSQL, "PXP")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objParcela As New PedidoXParcela(Me.Pedido)
                objParcela.CodigoParcela = row("Parcela_Id")
                objParcela.DataVencimento = row("Vencimento")
                objParcela.Valor = row("Valor")

                Me.Add(objParcela)
            Next
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _ApuracaoParcela As PedidoxParcelaApuracao
    Private _Msg As String = ""
#End Region

#Region "Property"
    Public ReadOnly Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public ReadOnly Property ApuracaoParcela As PedidoxParcelaApuracao
        Get
            If _ApuracaoParcela Is Nothing Then _ApuracaoParcela = New PedidoxParcelaApuracao(Me.Pedido)
            Return _ApuracaoParcela
        End Get
    End Property

    Public Property MSG() As String
        Get
            Return _Msg
        End Get
        Set(ByVal value As String)
            _Msg = value
        End Set
    End Property

    Public ReadOnly Property Total As Decimal
        Get
            Return (From x In Me Select x.Valor).Sum()
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Parc As PedidoXParcela In Me
            If Parc.Valor = 0 Then Continue For
            If Pedido.IUD = "I" Or Pedido.IUD = "D" Then Parc.IUD = Pedido.IUD
            If Pedido.IUD = "C" Then Parc.IUD = "D"

            Parc.SalvarSql(Sqls)
        Next
    End Sub

    Public Sub Parcelar(pDataInicial As Date)
        Dim ValorParcela As Decimal
        Dim VlrParcelaAjuste As Decimal

        If Me.Pedido.IUD = "I" Then
            Me.Clear()
            ValorParcela = Math.Round(Me.Pedido.Itens.Liquido / Me.Pedido.CondicaoPagamento.Parcelas, 2)
            VlrParcelaAjuste = Me.Pedido.Itens.Liquido - (ValorParcela * Me.Pedido.CondicaoPagamento.Parcelas)

            For i As Integer = 1 To Me.Pedido.CondicaoPagamento.Parcelas
                Dim Parc As New PedidoXParcela(Me.Pedido)
                Parc.CodigoParcela = i
                Parc.DataVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, pDataInicial.AddDays(Me.Pedido.CondicaoPagamento.Periodo(i - 1)))
                Parc.Valor = ValorParcela
                Me.Add(Parc)
            Next
            Me(Me.Count - 1).Valor += VlrParcelaAjuste
        Else
            ValorParcela = Me.Pedido.Itens.Liquido / Me.Pedido.CondicaoPagamento.Parcelas
            VlrParcelaAjuste = Me.Pedido.Itens.Liquido - (ValorParcela * Me.Pedido.CondicaoPagamento.Parcelas)

            Me.RemoveAll(Function(s) s.IUD = "I")
            Me.ForEach(Function(s)
                           If s.CodigoParcela > Me.Pedido.CondicaoPagamento.Parcelas Then
                               s.IUD = "D"
                           End If
                           Return True
                       End Function)

            Dim Parc As PedidoXParcela
            Dim t As Integer = 1
            For x As Integer = 1 To Me.Pedido.CondicaoPagamento.Parcelas
                Parc = Me.Where(Function(S) S.CodigoParcela = t).FirstOrDefault
                If Parc Is Nothing Then
                    Parc = New PedidoXParcela(Me.Pedido)
                    Parc.IUD = "I"
                    Parc.CodigoParcela = x
                Else
                    Parc.IUD = "U"
                End If

                Parc.DataVencimento = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, pDataInicial.AddDays(Me.Pedido.CondicaoPagamento.Periodo(x - 1)))
                Parc.Valor = ValorParcela
                If Parc.IUD = "I" Then Me.Add(Parc)
                t += 1
            Next
            Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += VlrParcelaAjuste
        End If

    End Sub

    Public Sub ModificarParcela(ByVal pCodigoParcela As Integer, ByVal NovoVencimento As DateTime, ByVal NovoValor As Decimal)
        Dim objParcela As PedidoXParcela = Me.Where(Function(s) s.CodigoParcela = pCodigoParcela).First
        Dim Venc As Date = Funcoes.ValidaDataUtil(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, NovoVencimento)

        If Venc > objParcela.DataVencimento And objParcela.CodigoParcela < Me.Pedido.CondicaoPagamento.Parcelas Then
            Dim objParcelaProx As PedidoXParcela = Me.Where(Function(s) s.CodigoParcela = pCodigoParcela + 1).First

            If objParcelaProx IsNot Nothing Then
                If Venc > objParcelaProx.DataVencimento Then
                    MSG = "Vencimento da Parcela não pode ser maior do que o vencimento da próxima parcela"
                    Exit Sub
                End If
            End If
        End If


        Dim conf As Decimal = (From p In Me Where p.CodigoParcela < pCodigoParcela Group By p.Pedido.Codigo Into valorv = Sum(p.Valor) Select valorv).FirstOrDefault

        If conf + NovoValor > Me.Pedido.Itens.Liquido Then
            MSG = "Valor da Parcela somado as demais parcelas anteriores nao pode exceder o valor do pedido."
            Exit Sub
        End If

        objParcela.IUD = "U"
        objParcela.DataVencimento = Venc

        If objParcela.Valor <> NovoValor Then
            Dim VlrParcelaAjuste As Decimal
            Dim AjusteFinal As Decimal
            Dim NumPar As Integer = Me.Pedido.CondicaoPagamento.Parcelas - objParcela.CodigoParcela

            If NumPar = 0 Then
                MSG = "O Valor da ultima parcela nao pode ser alterado."
                Exit Sub
            End If

            If NovoValor > objParcela.Valor Then
                VlrParcelaAjuste = Math.Round((NovoValor - objParcela.Valor) / NumPar, 2)
                AjusteFinal = (NovoValor - objParcela.Valor) - (VlrParcelaAjuste * NumPar)

                For i As Integer = objParcela.CodigoParcela To Me.Pedido.CondicaoPagamento.Parcelas - 1
                    Me(i).IUD = "U"
                    Me(i).Valor -= VlrParcelaAjuste
                Next

                Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += AjusteFinal
            Else
                VlrParcelaAjuste = Math.Round((objParcela.Valor - NovoValor) / NumPar, 2)
                AjusteFinal = (objParcela.Valor - NovoValor) - (VlrParcelaAjuste * NumPar)

                For i As Integer = objParcela.CodigoParcela To Me.Pedido.CondicaoPagamento.Parcelas - 1
                    Me(i).IUD = "U"
                    Me(i).Valor += VlrParcelaAjuste
                Next

                Me(Me.Pedido.CondicaoPagamento.Parcelas - 1).Valor += AjusteFinal
            End If
        End If
        objParcela.Valor = NovoValor

    End Sub

#End Region
End Class

'******************************************************************************************************************************************************
'****************************************************    CLASSE BASE PARCELAS   ***********************************************************************
'******************************************************************************************************************************************************
<Serializable()> _
Public Class PedidoXParcela
#Region "Construtor"
    Public Sub New(pPedido As Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Pedido As Pedido
    Private _CodigoParcela As Integer
    Private _DataVencimento As Date
    Private _Valor As Decimal
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public Property CodigoParcela() As Integer
        Get
            Return _CodigoParcela
        End Get
        Set(ByVal value As Integer)
            _CodigoParcela = value
        End Set
    End Property

    Public Property DataVencimento() As Date
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As Date)
            _DataVencimento = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
        End Set
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
            Case "I", "U"
                Sql = " ;MERGE PedidoXParcela AS Dest" & vbCrLf & _
                      " USING (Select '" & Me.Pedido.CodigoEmpresa & "' as Empresa_Id," & Me.Pedido.EnderecoEmpresa & " as EndEmpresa_Id," & Pedido.Codigo & " as Pedido_Id," & Me.CodigoParcela & " as Parcela_Id) AS Ori" & vbCrLf & _
                      "    ON Dest.Empresa_Id    = Ori.Empresa_Id" & vbCrLf & _
                      "   and Dest.EndEmpresa_Id = Ori.EndEmpresa_Id" & vbCrLf & _
                      "   and Dest.Pedido_Id     = Ori.Pedido_Id" & vbCrLf & _
                      "   and Dest.Parcela_id    = Ori.Parcela_id" & vbCrLf & _
                      "  WHEN NOT MATCHED" & vbCrLf & _
                      "    THEN Insert (Empresa_Id, EndEmpresa_Id, Pedido_Id, Parcela_Id, Vencimento,  Valor)" & vbCrLf & _
                      "		values ('" & Me.Pedido.CodigoEmpresa & "', " & Me.Pedido.EnderecoEmpresa & "," & Pedido.Codigo & "," & Me.CodigoParcela & ",'" & Me.DataVencimento.ToSqlDate & "'," & Str(Me.Valor) & ") " & vbCrLf & _
                      "  WHEN MATCHED " & vbCrLf & _
                      "    THEN Update set " & vbCrLf & _
                      "          Vencimento ='" & Me.DataVencimento.ToSqlDate & "'" & vbCrLf & _
                      "         ,Valor      = " & Str(Me.Valor) & ";" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE PedidoXParcela " & _
                      " WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.Pedido.Codigo & vbCrLf & _
                      "   AND Parcela_Id    = " & Me.CodigoParcela & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region
End Class

'******************************************************************************************************************************************************
'****************************************************    CLASSE BASE PARCELAS APURACAO PEDIDO   *******************************************************
'******************************************************************************************************************************************************
<Serializable()> _
Public Class PedidoxParcelaApuracao

#Region "Construtor"
    Public Sub New(pPedido As Pedido)
        _Pedido = pPedido
        AtualizaFinanceiro()
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Pedido
    Private _CPsnf_AbertoOficial As Decimal
    Private _CPsnf_AbertoMoeda As Decimal
    Private _CPsnf_BaixadoOficial As Decimal
    Private _CPsnf_BaixadoMoeda As Decimal
    Private _CPsnf_AdiantamentoOficial As Decimal
    Private _CPsnf_AdiantamentoMoeda As Decimal
    Private _CPsnf_BXAdiantamentoOficial As Decimal
    Private _CPsnf_BXAdiantamentoMoeda As Decimal
    Private _CPcnf_AbertoOficial As Decimal
    Private _CPcnf_AbertoMoeda As Decimal
    Private _CPcnf_BaixadoOficial As Decimal
    Private _CPcnf_BaixadoMoeda As Decimal
    Private _CPcnf_AdiantamentoOficial As Decimal
    Private _CPcnf_AdiantamentoMoeda As Decimal
    Private _CPcnf_BXAdiantamentoOficial As Decimal
    Private _CPcnf_BXAdiantamentoMoeda As Decimal
    Private _CRsnf_AbertoOficial As Decimal
    Private _CRsnf_AbertoMoeda As Decimal
    Private _CRsnf_BaixadoOficial As Decimal
    Private _CRsnf_BaixadoMoeda As Decimal
    Private _CRsnf_AdiantamentoOficial As Decimal
    Private _CRsnf_AdiantamentoMoeda As Decimal
    Private _CRsnf_BXAdiantamentoOficial As Decimal
    Private _CRsnf_BXAdiantamentoMoeda As Decimal
    Private _CRcnf_AbertoOficial As Decimal
    Private _CRcnf_AbertoMoeda As Decimal
    Private _CRcnf_BaixadoOficial As Decimal
    Private _CRcnf_BaixadoMoeda As Decimal
    Private _CRcnf_AdiantamentoOficial As Decimal
    Private _CRcnf_AdiantamentoMoeda As Decimal
    Private _CRcnf_BXAdiantamentoOficial As Decimal
    Private _CRcnf_BXAdiantamentoMoeda As Decimal
    Private _SaldoOficialAGerar As Decimal
    Private _SaldoMoedaAGerar As Decimal
#End Region

#Region "Property"
    '*****************************************************************
    '************************   Saldo   ******************************
    '*****************************************************************
    Public ReadOnly Property SaldoAdiantamento As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return CP_Adiantamento - CP_BxAdiantamento
            Else
                Return CR_Adiantamento - CR_BxAdiantamento
            End If
        End Get
    End Property

    Public Property SaldoOficialAGerar As Decimal
        Get
            Return _SaldoOficialAGerar
        End Get
        Set(value As Decimal)
            _SaldoOficialAGerar = value
        End Set
    End Property

    Public Property SaldoMoedaAGerar As Decimal
        Get
            Return _SaldoMoedaAGerar
        End Get
        Set(value As Decimal)
            _SaldoMoedaAGerar = value
        End Set
    End Property

    Public ReadOnly Property SaldoAGerar As Decimal
        Get
            If MoedaEstrangeira Then
                Return Me.SaldoMoedaAGerar
            Else
                Return Me.SaldoOficialAGerar
            End If
        End Get
    End Property

    '*****************************************************************
    '*************************  PEDIDO  ******************************
    '*****************************************************************
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido

        End Get
    End Property
    Public ReadOnly Property TotalOficial
        Get
            Return Pedido.Itens.TotalOficial
        End Get
    End Property
    Public ReadOnly Property TotalMoeda
        Get
            Return Pedido.Itens.TotalMoeda
        End Get
    End Property
    Public ReadOnly Property Total As Decimal
        Get
            Return Pedido.Itens.Total
        End Get
    End Property
    Public ReadOnly Property LiquidoOficial
        Get
            Return Pedido.Itens.LiquidoOficial
        End Get
    End Property
    Public ReadOnly Property LiquidoMoeda
        Get
            Return Pedido.Itens.LiquidoMoeda
        End Get
    End Property
    Public ReadOnly Property Liquido As Decimal
        Get
            Return Pedido.Itens.Liquido
        End Get
    End Property
    Public ReadOnly Property MoedaEstrangeira As Boolean
        Get
            Return Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira
        End Get
    End Property

    '*****************************************************************
    '***********************  A PAGAR  *******************************
    '*****************************************************************
    Public Property CPsnf_AbertoOficial As Decimal
        Get
            Return _CPsnf_AbertoOficial
        End Get
        Set(value As Decimal)
            _CPsnf_AbertoOficial = value
        End Set
    End Property
    Public Property CPsnf_AbertoMoeda As Decimal
        Get
            Return _CPsnf_AbertoMoeda
        End Get
        Set(value As Decimal)
            _CPsnf_AbertoMoeda = value
        End Set
    End Property
    Public Property CPsnf_BaixadoOficial As Decimal
        Get
            Return _CPsnf_BaixadoOficial
        End Get
        Set(value As Decimal)
            _CPsnf_BaixadoOficial = value
        End Set
    End Property
    Public Property CPsnf_BaixadoMoeda As Decimal
        Get
            Return _CPsnf_BaixadoMoeda
        End Get
        Set(value As Decimal)
            _CPsnf_BaixadoMoeda = value
        End Set
    End Property
    Public Property CPsnf_AdiantamentoOficial As Decimal
        Get
            Return _CPsnf_AdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CPsnf_AdiantamentoOficial = value
        End Set
    End Property
    Public Property CPsnf_AdiantamentoMoeda As Decimal
        Get
            Return _CPsnf_AdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CPsnf_AdiantamentoMoeda = value
        End Set
    End Property
    Public Property CPsnf_BXAdiantamentoOficial As Decimal
        Get
            Return _CPsnf_BXAdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CPsnf_BXAdiantamentoOficial = value
        End Set
    End Property
    Public Property CPsnf_BXAdiantamentoMoeda As Decimal
        Get
            Return _CPsnf_BXAdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CPsnf_BXAdiantamentoMoeda = value
        End Set
    End Property
    Public Property CPcnf_AbertoOficial As Decimal
        Get
            Return _CPcnf_AbertoOficial
        End Get
        Set(value As Decimal)
            _CPcnf_AbertoOficial = value
        End Set
    End Property
    Public Property CPcnf_AbertoMoeda As Decimal
        Get
            Return _CPcnf_AbertoMoeda
        End Get
        Set(value As Decimal)
            _CPcnf_AbertoMoeda = value
        End Set
    End Property
    Public Property CPcnf_BaixadoOficial As Decimal
        Get
            Return _CPcnf_BaixadoOficial
        End Get
        Set(value As Decimal)
            _CPcnf_BaixadoOficial = value
        End Set
    End Property
    Public Property CPcnf_BaixadoMoeda As Decimal
        Get
            Return _CPcnf_BaixadoMoeda
        End Get
        Set(value As Decimal)
            _CPcnf_BaixadoMoeda = value
        End Set
    End Property
    Public Property CPcnf_AdiantamentoOficial As Decimal
        Get
            Return _CPcnf_AdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CPcnf_AdiantamentoOficial = value
        End Set
    End Property
    Public Property CPcnf_AdiantamentoMoeda As Decimal
        Get
            Return _CPcnf_AdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CPcnf_AdiantamentoMoeda = value
        End Set
    End Property
    Public Property CPcnf_BXAdiantamentoOficial As Decimal
        Get
            Return _CPcnf_BXAdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CPcnf_BXAdiantamentoOficial = value
        End Set
    End Property
    Public Property CPcnf_BXAdiantamentoMoeda As Decimal
        Get
            Return _CPcnf_BXAdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CPcnf_BXAdiantamentoMoeda = value
        End Set
    End Property

    '*******************************************************************************************
    '***********************  A PAGAR SEGUIDO A MOEDA DO PEDIDO  *******************************
    '*******************************************************************************************
    '** SEM NOTA **
    Public ReadOnly Property CPsnf_Aberto As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPsnf_AbertoMoeda
            Else
                Return _CPsnf_AbertoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CPsnf_Baixado As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPsnf_BaixadoMoeda
            Else
                Return _CPsnf_BaixadoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CPsnf_Adiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPsnf_AdiantamentoMoeda
            Else
                Return _CPsnf_AdiantamentoOficial
            End If

        End Get
    End Property
    Public ReadOnly Property CPsnf_BXAdiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPsnf_BXAdiantamentoMoeda
            Else
                Return _CPsnf_BXAdiantamentoOficial
            End If
        End Get
    End Property

    '** COM NOTA **
    Public ReadOnly Property CPcnf_Aberto As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPcnf_AbertoMoeda
            Else
                Return _CPcnf_AbertoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CPcnf_Baixado As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPcnf_BaixadoMoeda
            Else
                Return _CPcnf_BaixadoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CPcnf_Adiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPcnf_AdiantamentoMoeda
            Else
                Return _CPcnf_AdiantamentoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CPcnf_BXAdiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CPcnf_BXAdiantamentoMoeda
            Else
                Return _CPcnf_BXAdiantamentoOficial
            End If
        End Get
    End Property

    '** Resumo CP **
    Public ReadOnly Property CP_Aberto As Decimal
        Get
            Return CPcnf_Aberto + CPsnf_Aberto
        End Get
    End Property
    Public ReadOnly Property CP_Baixado As Decimal
        Get
            Return CPcnf_Baixado + CPsnf_Baixado
        End Get
    End Property
    Public ReadOnly Property CP_Adiantamento As Decimal
        Get
            Return CPcnf_Adiantamento + CPsnf_Adiantamento
        End Get
    End Property
    Public ReadOnly Property CP_BxAdiantamento As Decimal
        Get
            Return CPcnf_BXAdiantamento + CPsnf_BXAdiantamento
        End Get
    End Property

    '*****************************************************************
    '***********************  A RECEBER  *****************************
    '*****************************************************************
    Public Property CRsnf_AbertoOficial As Decimal
        Get
            Return _CRsnf_AbertoOficial
        End Get
        Set(value As Decimal)
            _CRsnf_AbertoOficial = value
        End Set
    End Property
    Public Property CRsnf_AbertoMoeda As Decimal
        Get
            Return _CRsnf_AbertoMoeda
        End Get
        Set(value As Decimal)
            _CRsnf_AbertoMoeda = value
        End Set
    End Property
    Public Property CRsnf_BaixadoOficial As Decimal
        Get
            Return _CRsnf_BaixadoOficial
        End Get
        Set(value As Decimal)
            _CRsnf_BaixadoOficial = value
        End Set
    End Property
    Public Property CRsnf_BaixadoMoeda As Decimal
        Get
            Return _CRsnf_BaixadoMoeda
        End Get
        Set(value As Decimal)
            _CRsnf_BaixadoMoeda = value
        End Set
    End Property
    Public Property CRsnf_AdiantamentoOficial As Decimal
        Get
            Return _CRsnf_AdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CRsnf_AdiantamentoOficial = value
        End Set
    End Property
    Public Property CRsnf_AdiantamentoMoeda As Decimal
        Get
            Return _CRsnf_AdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CRsnf_AdiantamentoMoeda = value
        End Set
    End Property
    Public Property CRsnf_BXAdiantamentoOficial As Decimal
        Get
            Return _CRsnf_BXAdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CRsnf_BXAdiantamentoOficial = value
        End Set
    End Property
    Public Property CRsnf_BXAdiantamentoMoeda As Decimal
        Get
            Return _CRsnf_BXAdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CRsnf_BXAdiantamentoMoeda = value
        End Set
    End Property
    Public Property CRcnf_AbertoOficial As Decimal
        Get
            Return _CRcnf_AbertoOficial
        End Get
        Set(value As Decimal)
            _CRcnf_AbertoOficial = value
        End Set
    End Property
    Public Property CRcnf_AbertoMoeda As Decimal
        Get
            Return _CRcnf_AbertoMoeda
        End Get
        Set(value As Decimal)
            _CRcnf_AbertoMoeda = value
        End Set
    End Property
    Public Property CRcnf_BaixadoOficial As Decimal
        Get
            Return _CRcnf_BaixadoOficial
        End Get
        Set(value As Decimal)
            _CRcnf_BaixadoOficial = value
        End Set
    End Property
    Public Property CRcnf_BaixadoMoeda As Decimal
        Get
            Return _CRcnf_BaixadoMoeda
        End Get
        Set(value As Decimal)
            _CRcnf_BaixadoMoeda = value
        End Set
    End Property
    Public Property CRcnf_AdiantamentoOficial As Decimal
        Get
            Return _CRcnf_AdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CRcnf_AdiantamentoOficial = value
        End Set
    End Property
    Public Property CRcnf_AdiantamentoMoeda As Decimal
        Get
            Return _CRcnf_AdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CRcnf_AdiantamentoMoeda = value
        End Set
    End Property
    Public Property CRcnf_BXAdiantamentoOficial As Decimal
        Get
            Return _CRcnf_BXAdiantamentoOficial
        End Get
        Set(value As Decimal)
            _CRcnf_BXAdiantamentoOficial = value
        End Set
    End Property
    Public Property CRcnf_BXAdiantamentoMoeda As Decimal
        Get
            Return _CRcnf_BXAdiantamentoMoeda
        End Get
        Set(value As Decimal)
            _CRcnf_BXAdiantamentoMoeda = value
        End Set
    End Property

    '*******************************************************************************************
    '***********************  A RECEBER SEGUIDO A MOEDA DO PEDIDO  *******************************
    '*******************************************************************************************
    '** SEM NOTA **
    Public ReadOnly Property CRsnf_Aberto As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRsnf_AbertoMoeda
            Else
                Return _CRsnf_AbertoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRsnf_Baixado As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRsnf_BaixadoMoeda
            Else
                Return _CRsnf_BaixadoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRsnf_Adiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRsnf_AdiantamentoMoeda
            Else
                Return _CRsnf_AdiantamentoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRsnf_BXAdiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRsnf_BXAdiantamentoMoeda
            Else
                Return _CRsnf_BXAdiantamentoOficial
            End If
        End Get
    End Property

    '** COM NOTA **
    Public ReadOnly Property CRcnf_Aberto As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRcnf_AbertoMoeda
            Else
                Return _CRcnf_AbertoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRcnf_Baixado As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRcnf_BaixadoMoeda
            Else
                Return _CRcnf_BaixadoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRcnf_Adiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRcnf_AdiantamentoMoeda
            Else
                Return _CRcnf_AdiantamentoOficial
            End If
        End Get
    End Property
    Public ReadOnly Property CRcnf_BXAdiantamento As Decimal
        Get
            If MoedaEstrangeira Then
                Return _CRcnf_BXAdiantamentoMoeda
            Else
                Return _CRcnf_BXAdiantamentoOficial
            End If
        End Get
    End Property

    '** Resumo CR **
    Public ReadOnly Property CR_Aberto As Decimal
        Get
            Return CRcnf_Aberto + CRsnf_Aberto
        End Get
    End Property
    Public ReadOnly Property CR_Baixado As Decimal
        Get
            Return CRcnf_Baixado + CRsnf_Baixado
        End Get
    End Property
    Public ReadOnly Property CR_Adiantamento As Decimal
        Get
            Return CRcnf_Adiantamento + CRsnf_Adiantamento
        End Get
    End Property
    Public ReadOnly Property CR_BxAdiantamento As Decimal
        Get
            Return CRcnf_BXAdiantamento + CRsnf_BXAdiantamento
        End Get
    End Property

#End Region

#Region "Methods"
    Public Sub AtualizaFinanceiro()
        Dim sql As String
        sql = "Select CPsnf_AbertoOficial," & vbCrLf & _
              "       CPsnf_AbertoMoeda," & vbCrLf & _
              "       CPsnf_BaixadoOficial," & vbCrLf & _
              "       CPsnf_BaixadoMoeda," & vbCrLf & _
              "       CPsnf_AdiantamentoOficial," & vbCrLf & _
              "       CPsnf_AdiantamentoMoeda," & vbCrLf & _
              "       CPsnf_BXAdiantamentoOficial," & vbCrLf & _
              "       CPsnf_BXAdiantamentoMoeda," & vbCrLf & _
              "       CPcnf_AbertoOficial," & vbCrLf & _
              "       CPcnf_AbertoMoeda," & vbCrLf & _
              "       CPcnf_BaixadoOficial," & vbCrLf & _
              "       CPcnf_BaixadoMoeda," & vbCrLf & _
              "       CPcnf_AdiantamentoOficial," & vbCrLf & _
              "       CPcnf_AdiantamentoMoeda," & vbCrLf & _
              "       CPcnf_BXAdiantamentoOficial," & vbCrLf & _
              "       CPcnf_BXAdiantamentoMoeda," & vbCrLf & _
              "       CRsnf_AbertoOficial," & vbCrLf & _
              "       CRsnf_AbertoMoeda," & vbCrLf & _
              "       CRsnf_BaixadoOficial," & vbCrLf & _
              "       CRsnf_BaixadoMoeda," & vbCrLf & _
              "       CRsnf_AdiantamentoOficial," & vbCrLf & _
              "       CRsnf_AdiantamentoMoeda," & vbCrLf & _
              "       CRsnf_BXAdiantamentoOficial," & vbCrLf & _
              "       CRsnf_BXAdiantamentoMoeda," & vbCrLf & _
              "       CRcnf_AbertoOficial," & vbCrLf & _
              "       CRcnf_AbertoMoeda," & vbCrLf & _
              "       CRcnf_BaixadoOficial," & vbCrLf & _
              "       CRcnf_BaixadoMoeda," & vbCrLf & _
              "       CRcnf_AdiantamentoOficial," & vbCrLf & _
              "       CRcnf_AdiantamentoMoeda," & vbCrLf & _
              "       CRcnf_BXAdiantamentoOficial," & vbCrLf & _
              "       CRcnf_BXAdiantamentoMoeda," & vbCrLf & _
              "       SaldoOficialAGerar," & vbCrLf & _
              "       SaldoMoedaAGerar" & vbCrLf & _
              "  From VW_ResumoFinanceiroPedido " & vbCrLf & _
              " Where Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & vbCrLf & _
              "   and Pedido_Id     = " & Me.Pedido.Codigo & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "AP")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim Row As DataRow = ds.Tables(0).Rows(0)

        Me.CPsnf_AbertoOficial = Row("CPsnf_AbertoOficial")
        Me.CPsnf_AbertoMoeda = Row("CPsnf_AbertoMoeda")
        Me.CPsnf_BaixadoOficial = Row("CPsnf_BaixadoOficial")
        Me.CPsnf_BaixadoMoeda = Row("CPsnf_BaixadoMoeda")
        Me.CPsnf_AdiantamentoOficial = Row("CPsnf_AdiantamentoOficial")
        Me.CPsnf_AdiantamentoMoeda = Row("CPsnf_AdiantamentoMoeda")
        Me.CPsnf_BXAdiantamentoOficial = Row("CPsnf_BXAdiantamentoOficial")
        Me.CPsnf_BXAdiantamentoMoeda = Row("CPsnf_BXAdiantamentoMoeda")
        Me.CPcnf_AbertoOficial = Row("CPcnf_AbertoOficial")
        Me.CPcnf_AbertoMoeda = Row("CPcnf_AbertoMoeda")
        Me.CPcnf_BaixadoOficial = Row("CPcnf_BaixadoOficial")
        Me.CPcnf_BaixadoMoeda = Row("CPcnf_BaixadoMoeda")
        Me.CPcnf_AdiantamentoOficial = Row("CPcnf_AdiantamentoOficial")
        Me.CPcnf_AdiantamentoMoeda = Row("CPcnf_AdiantamentoMoeda")
        Me.CPcnf_BXAdiantamentoOficial = Row("CPcnf_BXAdiantamentoOficial")
        Me.CPcnf_BXAdiantamentoMoeda = Row("CPcnf_BXAdiantamentoMoeda")
        Me.CRsnf_AbertoOficial = Row("CRsnf_AbertoOficial")
        Me.CRsnf_AbertoMoeda = Row("CRsnf_AbertoMoeda")
        Me.CRsnf_BaixadoOficial = Row("CRsnf_BaixadoOficial")
        Me.CRsnf_BaixadoMoeda = Row("CRsnf_BaixadoMoeda")
        Me.CRsnf_AdiantamentoOficial = Row("CRsnf_AdiantamentoOficial")
        Me.CRsnf_AdiantamentoMoeda = Row("CRsnf_AdiantamentoMoeda")
        Me.CRsnf_BXAdiantamentoOficial = Row("CRsnf_BXAdiantamentoOficial")
        Me.CRsnf_BXAdiantamentoMoeda = Row("CRsnf_BXAdiantamentoMoeda")
        Me.CRcnf_AbertoOficial = Row("CRcnf_AbertoOficial")
        Me.CRcnf_AbertoMoeda = Row("CRcnf_AbertoMoeda")
        Me.CRcnf_BaixadoOficial = Row("CRcnf_BaixadoOficial")
        Me.CRcnf_BaixadoMoeda = Row("CRcnf_BaixadoMoeda")
        Me.CRcnf_AdiantamentoOficial = Row("CRcnf_AdiantamentoOficial")
        Me.CRcnf_AdiantamentoMoeda = Row("CRcnf_AdiantamentoMoeda")
        Me.CRcnf_BXAdiantamentoOficial = Row("CRcnf_BXAdiantamentoOficial")
        Me.CRcnf_BXAdiantamentoMoeda = Row("CRcnf_BXAdiantamentoMoeda")
        Me.SaldoMoedaAGerar = Row("SaldoMoedaAGerar")
        Me.SaldoOficialAGerar = Row("SaldoOficialAGerar")
    End Sub
#End Region

End Class
