Imports System.Web
Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'**************************************************************************************************************************
'*************************************  LISTA CLASSE FIXACAO  *************************************************************
'**************************************************************************************************************************
<Serializable()>
Public Class ListFixacao
    Inherits List(Of Fixacao)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pItemPedido As PedidoXItem)
        Me.ItemPedido = pItemPedido
        If Me.ItemPedido.Pedido.Codigo = 0 Then Exit Sub

        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Fixacao_Id, Procuracao, Operacao, SubOperacao, Documento, Movimento, Quantidade," & vbCrLf & _
                  "       UnitarioOficial, UnitarioMoeda, TotalOficial, TotalMoeda, CondicaoPagamento, IndiceFixado, OperacaoXEstado" & vbCrLf & _
                  "  FROM VW_PedidosXItensXFixacoes " & vbCrLf & _
                  " WHERE Empresa_Id    ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                  "   AND EndEmpresa_Id = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                  "   AND Pedido_Id     = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                  "   AND Produto_Id    ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                  " ORDER BY Fixacao_Id" & vbCrLf

            Dim dsFixacoes As DataSet = objBanco.ConsultaDataSet(sql, "PedidosXItensXFixacoes")

            For Each drFixacao As DataRow In dsFixacoes.Tables(0).Rows
                Dim objFixacao As New Fixacao(Me.ItemPedido)

                objFixacao.Codigo = Convert.ToInt32(drFixacao("Fixacao_Id"))
                If Not drFixacao.IsNull("Procuracao") Then objFixacao.Procuracao = Convert.ToInt32(drFixacao("Procuracao")) Else objFixacao.Procuracao = -1
                objFixacao.CodigoOperacao = drFixacao("Operacao")
                objFixacao.CodigoSubOperacao = drFixacao("SubOperacao")
                objFixacao.Documento = drFixacao("Documento")
                objFixacao.Movimento = drFixacao("Movimento")
                objFixacao.Quantidade = drFixacao("Quantidade")
                objFixacao.UnitarioOficial = drFixacao("UnitarioOficial")
                objFixacao.UnitarioMoeda = drFixacao("UnitarioMoeda")
                objFixacao.TotalOficial = drFixacao("TotalOficial")
                objFixacao.TotalMoeda = drFixacao("TotalMoeda")
                If Not drFixacao.IsNull("CondicaoPagamento") Then objFixacao.CondicaoPagamento = drFixacao("CondicaoPagamento") Else objFixacao.CondicaoPagamento = -1
                objFixacao.IndiceFixado = drFixacao("IndiceFixado")
                objFixacao.CodigoOperacaoxEstado = drFixacao("OperacaoXEstado")

                Me.Add(objFixacao)
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _ItemPedido As PedidoXItem
    Private _SaldoProcuracao As Decimal
#End Region

#Region "Property"
    Public Property ItemPedido() As PedidoXItem
        Get
            Return _ItemPedido
        End Get
        Set(ByVal value As PedidoXItem)
            _ItemPedido = value
        End Set
    End Property

    Public Property SaldoProcuracao() As Decimal
        Get
            If _SaldoProcuracao = 0 Then VerProcuracao()
            Return _SaldoProcuracao
        End Get
        Set(ByVal value As Decimal)
            _SaldoProcuracao = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each fix As Fixacao In Me
            If ItemPedido.IUD = "D" Or ItemPedido.IUD = "I" Then fix.IUD = ItemPedido.IUD
            fix.SalvarSql(Sqls)
        Next
    End Sub

    Private Sub VerProcuracao()
        Dim sql As String = "SELECT isnull(Sum((isnull(P.Quantidade,0) - ISNULL(sb_fix.Quantidade, 0))),0) AS Saldo " & vbCrLf & _
                            "FROM Procuracoes P    " & vbCrLf & _
                            "		INNER JOIN Clientes C     " & vbCrLf & _
                            "				ON C.Cliente_Id   = P.Cessionario " & vbCrLf & _
                            "				AND C.Endereco_Id = P.EndCessionario " & vbCrLf & _
                            "        LEFT JOIN (Select pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id, pXiXf.Procuracao, " & vbCrLf & _
                            "							sum(isnull(pXiXf.Quantidade,0)) AS Quantidade " & vbCrLf & _
                            "					from PedidosXItensXFixacoes pXiXf " & vbCrLf & _
                            "					Where pXiXf.Empresa_Id    = '" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "                      AND pXiXf.EndEmpresa_Id = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                            "                    Group by pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id, pXiXf.Procuracao) AS sb_fix " & vbCrLf & _
                            "                ON sb_fix.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
                            "               AND sb_fix.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                            "               AND sb_fix.Pedido_Id     = P.PedidoCedente " & vbCrLf & _
                            "               AND sb_fix.Procuracao    = P.Procuracao_ID  " & vbCrLf & _
                            "WHERE P.Empresa_Id = '" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "  AND P.EndEmpresa_Id = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                            "  AND P.PedidoCedente = " & _ItemPedido.Pedido.Codigo & vbCrLf & _
                            "  AND P.Situacao = 1 " & vbCrLf & _
                            "  AND P.Quantidade > 0 "

        Dim objBanco As New AcessaBanco()
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "SaldoDeProcuracao")

        For Each row In ds.Tables(0).Rows
            _SaldoProcuracao += row("Saldo")
        Next

        If _SaldoProcuracao = 0 Then _SaldoProcuracao = -1
    End Sub

    Public Function ListarProcuracao() As DataSet
        Dim sql As String = "SELECT (P.Cedente + ' - ' + C.Nome) AS Cessionario, P.Procuracao_ID AS Procuracao, " & vbCrLf & _
                             "       P.Quantidade, ISNULL(sb_fix.Quantidade, 0) AS Fixado, " & vbCrLf & _
                             "       (P.Quantidade - ISNULL(sb_fix.Quantidade, 0)) AS Saldo" & vbCrLf & _
                             "FROM Procuracoes P    " & vbCrLf & _
                             "		INNER JOIN Clientes C     " & vbCrLf & _
                             "				ON C.Cliente_Id   = P.Cessionario " & vbCrLf & _
                             "				AND C.Endereco_Id = P.EndCessionario " & vbCrLf & _
                             "        LEFT JOIN (Select pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id, pXiXf.Procuracao, " & vbCrLf & _
                             "							sum(isnull(pXiXf.Quantidade,0)) AS Quantidade " & vbCrLf & _
                             "					from PedidosXItensXFixacoes pXiXf " & vbCrLf & _
                             "					Where pXiXf.Empresa_Id    = '" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                             "                      AND pXiXf.EndEmpresa_Id = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                             "                    Group by pXiXf.Empresa_Id, pXiXf.EndEmpresa_Id, pXiXf.Pedido_Id, pXiXf.Procuracao) AS sb_fix " & vbCrLf & _
                             "                ON sb_fix.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
                             "               AND sb_fix.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
                             "               AND sb_fix.Pedido_Id     = P.PedidoCedente " & vbCrLf & _
                             "               AND sb_fix.Procuracao    = P.Procuracao_ID  " & vbCrLf & _
                             "WHERE P.Empresa_Id = '" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                             "  AND P.EndEmpresa_Id = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                             "  AND P.PedidoCedente = " & _ItemPedido.Pedido.Codigo & vbCrLf & _
                             "  AND P.Situacao = 1 " & vbCrLf & _
                             "  AND (P.Quantidade - ISNULL(sb_fix.Quantidade, 0)) > 0 "

        Dim objBanco As New AcessaBanco()
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "SaldoDeProcuracao")

        Return ds
    End Function
#End Region

End Class


'*************************************************************************************************************************
'*************************************  CLASSE BASE FIXACAO  *************************************************************
'*************************************************************************************************************************
<Serializable()>
Public Class Fixacao

#Region "Contrutor"
    Public Sub New(ByVal ItemPedido As PedidoXItem)
        Me.ItemPedido = ItemPedido
    End Sub

    Public Sub New(ByVal ItemPedido As PedidoXItem, ByVal CodigoFixacao As Integer)
        Me.ItemPedido = ItemPedido

        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String = ""
            sql = "SELECT Fixacao_Id, isnull(Procuracao,0) as Procuracao, Operacao, SubOperacao, Documento, Movimento, Quantidade," & _
                  "       UnitarioOficial, UnitarioMoeda, TotalOficial, TotalMoeda, CondicaoPagamento, IndiceFixado, OperacaoXEstado" & _
                  "  FROM VW_PedidosXItensXFixacoes" & _
                  " WHERE Empresa_Id    ='" & ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                  "   AND EndEmpresa_Id = " & ItemPedido.Pedido.EnderecoEmpresa.ToString() & " " & vbCrLf & _
                  "   AND Pedido_Id     = " & ItemPedido.Pedido.Codigo.ToString() & " " & vbCrLf & _
                  "   AND Produto_Id    ='" & ItemPedido.CodigoProduto & "' " & vbCrLf & _
                  "   AND Fixacao_Id    = " & CodigoFixacao & vbCrLf &
                  " ORDER BY Fixacao_Id"

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "Fixacoes")

            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.Codigo = row("Fixacao_Id")
            Me.Procuracao = row("Procuracao")
            Me.CodigoOperacao = row("Operacao")
            Me.CodigoSubOperacao = row("SubOperacao")
            Me.Documento = row("Documento")
            Me.Movimento = row("Movimento")
            Me.Quantidade = row("Quantidade")
            Me.UnitarioOficial = row("UnitarioOficial")
            Me.UnitarioMoeda = row("UnitarioMoeda")
            Me.TotalOficial = row("TotalOficial")
            Me.TotalMoeda = row("TotalMoeda")
            Me.CondicaoPagamento = row("CondicaoPagamento")
            Me.IndiceFixado = row("IndiceFixado")
            Me.CodigoOperacaoxEstado = row("OperacaoXEstado")
        Catch ex As Exception

        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _ItemPedido As PedidoXItem
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Procuracao As Integer
    Private _CodigoOperacao As Integer
    Private _CodigoSubOperacao As Integer
    Private _subOperacao As SubOperacao
    Private _Documento As String
    Private _Movimento As DateTime
    Private _Quantidade As Decimal
    Private _UnitarioOficial As Decimal
    Private _UnitarioMoeda As Decimal
    Private _TotalOficial As Decimal
    Private _TotalMoeda As Decimal
    Private _CondicaoPagamento As Integer
    Private _IndiceFixado As Decimal
    Private _AFixarDebito As String = ""
    Private _AFixarCredito As String = ""
    Private _CodigoOperacaoxEstado As Integer
    Private _OperacaoxEstado As OperacaoXEstado

    '************ Listas ************************
    Private _Encargos As ListFixacaoXEncargo
    Private _NotasFixadas As ListPedidoFixacaoXNotaFiscal
    Private _Titulos As Novo.ListTituloNovo
#End Region

#Region "Property"
    Public Property ItemPedido() As PedidoXItem
        Get
            Return _ItemPedido
        End Get
        Set(ByVal value As PedidoXItem)
            _ItemPedido = value
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Procuracao() As Integer
        Get
            Return _Procuracao
        End Get
        Set(ByVal value As Integer)
            _Procuracao = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
            _subOperacao = Nothing
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _subOperacao = Nothing
        End Set
    End Property

    Public ReadOnly Property SubOperacao As SubOperacao
        Get
            If _subOperacao Is Nothing And _CodigoOperacao > 0 And _CodigoSubOperacao > 0 Then _subOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _subOperacao
        End Get
    End Property

    Public Property Documento() As String
        Get
            Return _Documento
        End Get
        Set(ByVal value As String)
            _Documento = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property UnitarioOficial() As Decimal
        Get
            Return _UnitarioOficial
        End Get
        Set(ByVal value As Decimal)
            _UnitarioOficial = value
        End Set
    End Property

    Public Property UnitarioMoeda() As Decimal
        Get
            Return _UnitarioMoeda
        End Get
        Set(ByVal value As Decimal)
            _UnitarioMoeda = value
        End Set
    End Property

    Public Property TotalOficial() As Decimal
        Get
            Return _TotalOficial
        End Get
        Set(ByVal value As Decimal)
            _TotalOficial = value
        End Set
    End Property

    Public Property TotalMoeda() As Decimal
        Get
            Return _TotalMoeda
        End Get
        Set(ByVal value As Decimal)
            _TotalMoeda = value
        End Set
    End Property

    Public Property CondicaoPagamento() As Integer
        Get
            Return _CondicaoPagamento
        End Get
        Set(ByVal value As Integer)
            _CondicaoPagamento = value
        End Set
    End Property

    Public Property IndiceFixado() As Decimal
        Get
            Return _IndiceFixado
        End Get
        Set(ByVal value As Decimal)
            _IndiceFixado = value
        End Set
    End Property

    Public Property AFixarDebito() As String
        Get
            Return _AFixarDebito
        End Get
        Set(ByVal value As String)
            _AFixarDebito = value
        End Set
    End Property

    Public Property AFixarCredito() As String
        Get
            Return _AFixarCredito
        End Get
        Set(ByVal value As String)
            _AFixarCredito = value
        End Set
    End Property

    Public Property CodigoOperacaoxEstado As Integer
        Get
            Return _CodigoOperacaoxEstado
        End Get
        Set(value As Integer)
            _CodigoOperacaoxEstado = value
            _OperacaoxEstado = Nothing
        End Set
    End Property

    Public ReadOnly Property OperacaoxEstado As OperacaoXEstado
        Get
            If _OperacaoxEstado Is Nothing Then
                Dim Parametros As New OperacaoXEstado
                Parametros.Codigo = Me.CodigoOperacaoxEstado

                If Me.CodigoOperacaoxEstado > 0 Then
                    _OperacaoxEstado = New OperacaoXEstado(Parametros)
                Else
                    Parametros.Empresa = Left(Me.ItemPedido.Pedido.CodigoEmpresa, 8)
                    Parametros.CodigoGrupoProduto = Me.ItemPedido.Produto.CodigoGrupo
                    Parametros.CodigoProduto = Me.ItemPedido.CodigoProduto
                    Parametros.CodigoOperacao = Me.CodigoOperacao
                    Parametros.CodigoSubOperacao = Me.CodigoSubOperacao
                    Parametros.EstadoOrigem = Me.ItemPedido.Pedido.Empresa.CodigoEstado
                    Parametros.EstadoDestino = Me.ItemPedido.Pedido.Cliente.CodigoEstado
                    _OperacaoxEstado = New OperacaoXEstado(Parametros)
                End If
                _CodigoOperacaoxEstado = _OperacaoxEstado.Codigo
            End If
            Return _OperacaoxEstado
        End Get
    End Property

    '**********************************************************************
    '**************  Listas  **********************************************
    '**********************************************************************
    Public Property Encargos() As ListFixacaoXEncargo
        Get
            If _Encargos Is Nothing Then _Encargos = New ListFixacaoXEncargo(Me)
            Return _Encargos
        End Get
        Set(ByVal value As ListFixacaoXEncargo)
            _Encargos = value
        End Set
    End Property

    Public Property NotasFixacao() As ListPedidoFixacaoXNotaFiscal
        Get
            If _NotasFixadas Is Nothing Then
                _NotasFixadas = New ListPedidoFixacaoXNotaFiscal(Me)
                _NotasFixadas.CarregarNotasUsadasNaFixacao()
            End If
            Return _NotasFixadas
        End Get
        Set(ByVal value As ListPedidoFixacaoXNotaFiscal)
            _NotasFixadas = value
        End Set
    End Property

    Public Property Titulos() As Novo.ListTituloNovo
        Get
            If _Titulos Is Nothing Then
                _Titulos = New Novo.ListTituloNovo(Me)
            End If
            Return _Titulos
        End Get
        Set(ByVal value As Novo.ListTituloNovo)
            _Titulos = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function VerNotaFiscal() As Boolean
        Dim sql As String = "Select nXi.Nota_Id " & vbCrLf & _
                            "  From notasfiscaisXitens nXi " & vbCrLf & _
                            " inner join notasfiscais n " & vbCrLf & _
                            "    on n.empresa_id      = nXi.empresa_id " & vbCrLf & _
                            "	and n.endempresa_id   = nXi.endempresa_id " & vbCrLf & _
                            "	and n.Cliente_Id      = nXi.Cliente_Id " & vbCrLf & _
                            "	and n.EndCliente_Id   = nXi.EndCliente_Id " & vbCrLf & _
                            "	and n.EntradaSaida_Id = nXi.EntradaSaida_Id " & vbCrLf & _
                            "	and n.Serie_Id        = nXi.Serie_Id " & vbCrLf & _
                            "	and n.Nota_Id         = nXi.Nota_Id " & vbCrLf & _
                            " where nXi.Empresa_Id    ='" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "   AND nXi.EndEmpresa_Id = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                            "   AND nXi.Pedido        = " & _ItemPedido.Pedido.Codigo & vbCrLf & _
                            "   AND nXi.Fixacao       = " & _Codigo & vbCrLf & _
                            "   AND n.Situacao in(1,4,7) "

        Dim objBanco As New AcessaBanco()
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "NotasXFixacoes")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function VerVencimentos() As Boolean
        Dim sql As String = " Select Provisao, isnull(PedidoFixacao,0) AS PedidoFixacao " & vbCrLf & _
                            "   from ContasAPagar " & vbCrLf & _
                            "  where EmpresaPedido    ='" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "    AND EndEmpresaPedido = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                            "    AND Pedido           = " & _ItemPedido.Pedido.Codigo & vbCrLf & _
                            "    AND PedidoFixacao    = " & _Codigo & vbCrLf & _
                            "    And Situacao         = 1" & vbCrLf & _
                            "    AND Provisao         = 1 " & vbCrLf & _
                            " Union " & vbCrLf & _
                            " Select Provisao, isnull(PedidoFixacao,0) AS PedidoFixacao " & vbCrLf & _
                            "   from ContasAReceber " & vbCrLf & _
                            "  where EmpresaPedido    = '" & _ItemPedido.Pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "    AND EndEmpresaPedido = " & _ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                            "    AND Pedido           = " & _ItemPedido.Pedido.Codigo & vbCrLf & _
                            "    AND PedidoFixacao    = " & _Codigo & vbCrLf & _
                            "    And Situacao         = 1" & vbCrLf & _
                            "    AND Provisao         = 1 " & vbCrLf


        Dim objBanco As New AcessaBanco()
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "VencimentosFixacoes")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        If Me.SubOperacao.Classe <> eClassesOperacoes.COMPLEMENTACOES Then Exit Sub

        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "INSERT INTO PedidosXItensXFixacoes " & vbCrLf & _
                      "       (Empresa_Id, EndEmpresa_Id, Pedido_Id, Produto_Id, " & vbCrLf & _
                      "        Fixacao_Id, Procuracao, Operacao, SubOperacao, Documento, " & vbCrLf & _
                      "        Movimento, Quantidade, UnitarioOficial, UnitarioMoeda, " & vbCrLf & _
                      "        TotalOficial, TotalMoeda, CondicaoPagamento, IndiceFixado, OperacaoXEstado) " & vbCrLf & _
                      " Values ('" & Me.ItemPedido.Pedido.CodigoEmpresa & "'," & Me.ItemPedido.Pedido.EnderecoEmpresa & ", " & Me.ItemPedido.Pedido.Codigo & ",'" & Me.ItemPedido.CodigoProduto & "'," & vbCrLf & _
                      Me.Codigo & "," & Me.Procuracao & "," & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & ",'" & Me.Documento & "'," & vbCrLf & _
                      "'" & Me.Movimento.ToString("yyyy-MM-dd") & "'," & Str(Me.Quantidade) & "," & Str(Me.UnitarioOficial) & "," & Str(Me.UnitarioMoeda) & "," & vbCrLf & _
                      Str(Me.TotalOficial) & "," & Str(Me.TotalMoeda) & "," & Me.CondicaoPagamento & "," & Str(Me.IndiceFixado) & "," & Me.CodigoOperacaoxEstado & ")"

                Sqls.Add(sql)

                SalvarTabelasRelacionadasSql(Sqls)

            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)

                sql = "DELETE PedidosXItensXFixacoes " & _
                      " WHERE Empresa_Id    ='" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                      "   AND Pedido_Id     = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "   AND Produto_Id    ='" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "   AND Fixacao_Id    = " & Me.Codigo & vbCrLf

                Sqls.Add(sql)
            Case Else
                SalvarTabelasRelacionadasSql(Sqls)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Me.Encargos.SalvarSql(Sqls)

        If Not Me.NotasFixacao Is Nothing AndAlso Me.NotasFixacao.Count > 0 Then Me.NotasFixacao.SalvarSql(Sqls)

        If Me.ItemPedido.Pedido.FinanceiroNovo Then
            Me.Titulos.SalvarSql(Sqls, False)
        End If

        ContabilizarAFixar(Sqls)
    End Sub

    Private Sub ContabilizarAFixar(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                Dim LoteSeq As String = ""
                Dim i As Integer = 0

                sql = "Declare" & vbCrLf & _
                      " @SequenciaOriginal int" & vbCrLf & _
                      ",@SequenciaServidor int" & vbCrLf

                sql &= " Set @SequenciaOriginal = (Select Sequencia" & vbCrLf & _
                       "                             from numerador" & vbCrLf & _
                       "                            where UPPER(Empresa_id) = '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "'" & vbCrLf & _
                       "                              and Numerador_id      = 60)" & vbCrLf & _
                       " set @SequenciaServidor = (select isnull(Max(Sequencia_Id),isnull(@SequenciaOriginal,0))" & vbCrLf & _
                       "                             from razao    " & vbCrLf & _
                       "                            Where lote_Id      = 15 " & vbCrLf & _
                       "                              and Movimento_Id = '" & Me.Movimento.ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                       "                              and Sequencia_id between isnull(@SequenciaOriginal,0) and isnull(@SequenciaOriginal,0) + 99999)" & vbCrLf

                i += 1

                LoteSeq = "0015, @SequenciaServidor + " & i  'LOTE / SEQUENCIA

                sql &= " INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                       "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
                       "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
                       "                   Pedido, PedidoFixacao, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito)" & vbCrLf & _
                       "Values ('" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                       ", " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                       ",'" & Me.AFixarDebito & "'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",'" & Me.Movimento.ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                       ", " & LoteSeq & vbCrLf & _
                       ", '" & Me.ItemPedido.Pedido.CodigoUnidadeNegocio & "'" & vbCrLf & _
                       "," & Me.ItemPedido.Pedido.CodigoIndexador & vbCrLf & _
                       ",'" & Me.Movimento.ToString("yyyy/MM/dd") & "'" & vbCrLf
                'VALOR DÉBITO OFICIAL, CREDITO OFICIAL, DEBITO MOEDA, CREDITO MOEDA
                sql &= "," & Str(Me.Encargos.Where(Function(s) s.CodigoEncargo = "AFIXAR").First.ValorOficial) & ", 0.0" & ", " & Str(Me.TotalMoeda) & ", 0.0"

                sql &= ",'REF. FIXACAO " & Me.Codigo & ", PEDIDO: " & Me.ItemPedido.Pedido.Codigo & " - " & Me.ItemPedido.Pedido.Cliente.Nome & "'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",'AFIXAR'" & vbCrLf & _
                       ", " & Me.ItemPedido.Pedido.Codigo.ToSqlNULL & vbCrLf & _
                       ", " & Me.Codigo & vbCrLf & _
                       ",'P'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", '" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''," & 0 & ")"

                i += 1

                LoteSeq = "0015, @SequenciaServidor + " & i  'LOTE / SEQUENCIA

                sql &= " INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, UnidadeDeNegocio, Indexador, DataMoeda, " & vbCrLf & _
                       "                   DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda," & vbCrLf & _
                       "                   Historico, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf, Serie_Nf, Numero_Nf, Produto_NF, CFOP_NF, Sequencia_NF, Encargo_NF," & vbCrLf & _
                       "                   Pedido, PedidoFixacao, PrevistoRealizado, Custo, Produto, Rateado, Deposito, EndDeposito)" & vbCrLf & _
                       "Values ('" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                       ", " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                       ",'" & Me.AFixarCredito & "'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",'" & Me.Movimento.ToString("yyyy/MM/dd") & "'" & vbCrLf & _
                       ", " & LoteSeq & vbCrLf & _
                       ", '" & Me.ItemPedido.Pedido.CodigoUnidadeNegocio & "'" & vbCrLf & _
                       "," & Me.ItemPedido.Pedido.CodigoIndexador & vbCrLf & _
                       ",'" & Me.Movimento.ToString("yyyy/MM/dd") & "'" & vbCrLf
                'VALOR DÉBITO OFICIAL, CREDITO OFICIAL, DEBITO MOEDA, CREDITO MOEDA
                sql &= ", 0.0, " & Str(Me.Encargos.Where(Function(s) s.CodigoEncargo = "AFIXAR").First.ValorOficial) & ", 0.0" & ", " & Str(Me.TotalMoeda)

                sql &= ",'REF. FIXACAO " & Me.Codigo & ", PEDIDO: " & Me.ItemPedido.Pedido.Codigo & " - " & Me.ItemPedido.Pedido.Cliente.Nome & "'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",'AFIXAR'" & vbCrLf & _
                       ", " & Me.ItemPedido.Pedido.Codigo.ToSqlNULL & vbCrLf & _
                       ", " & Me.Codigo & vbCrLf & _
                       ",'P'" & vbCrLf & _
                       ",''" & vbCrLf & _
                       ", '" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                       ", " & 0 & vbCrLf & _
                       ",''," & 0 & ")"

                Sqls.Add(sql)
            Case "D"
                sql = " DELETE From Razao" & vbCrLf & _
                      "  WHERE Empresa_Id      = '" & Me.ItemPedido.Pedido.CodigoEmpresa & "'" & vbCrLf & _
                      "    And EndEmpresa_Id   = " & Me.ItemPedido.Pedido.EnderecoEmpresa & vbCrLf & _
                      "    And Pedido          = " & Me.ItemPedido.Pedido.Codigo & vbCrLf & _
                      "    And PedidoFixacao   = " & Me.Codigo & vbCrLf & _
                      "    And Produto         = '" & Me.ItemPedido.CodigoProduto & "'" & vbCrLf & _
                      "    And Encargo_NF      = 'AFIXAR'" & vbCrLf & _
                      "    And Lote_Id         = 15 "
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class


