Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListProcuracao
    Inherits List(Of Procuracao)
    Private Sql As String
    Private Banco As New AcessaBanco
    Private ds As DataSet

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pPedido As Integer, Optional ByVal zerado As Boolean = True)
        Sql = " SELECT pr.Empresa_Id, pr.EndEmpresa_Id," & vbCrLf & _
              "        pr.Procuracao_ID, pr.Situacao, pr.Documento," & vbCrLf & _
              "		   pr.Cedente, pr.EndCedente, pr.PedidoCedente," & vbCrLf & _
              "		   pr.Cessionario, pr.EndCessionario," & vbCrLf & _
              "        pr.Movimento," & vbCrLf & _
              "        isnull(pr.Observacoes, '') as Observacoes, isnull(pr.UsuarioInclusao, '') AS UsuarioInclusao," & vbCrLf & _
              "        isnull(pr.UsuarioInclusaoDate, '') AS UsuarioInclusaoDate," & vbCrLf & _
              "        isnull(pr.UsuarioAlteracao, '') AS UsuarioAlteracao," & vbCrLf & _
              "        isnull(pr.UsuarioAlteracaoDate, '') AS UsuarioAlteracaoDate," & vbCrLf & _
              "        isnull(pr.UsuarioCancelamento, '') AS UsuarioCancelamento," & vbCrLf & _
              "        isnull(pr.UsuarioCancelamentoDate, '') AS UsuarioCancelamentoDate," & vbCrLf & _
              "        pr.Quantidade," & vbCrLf & _
              "		   ISNULL(sb_Real.Quantidade, 0) AS Realizado," & vbCrLf & _
              "        PR.Quantidade - ISNULL(sb_Real.Quantidade, 0) AS Saldo" & vbCrLf & _
              "   FROM Procuracoes PR" & vbCrLf & _
              "   LEFT OUTER JOIN (Select sb.Empresa_Id," & vbCrLf & _
              "                           sb.EndEmpresa_Id," & vbCrLf & _
              "						      sb.Pedido," & vbCrLf & _
              "						      sb.Procuracao," & vbCrLf & _
              "						      SUM(sb.Quantidade) AS Quantidade" & vbCrLf & _
              "                      from (" & vbCrLf & _
              "							SELECT NF.Empresa_Id," & vbCrLf & _
              "								   NF.EndEmpresa_Id," & vbCrLf & _
              "								   NF.Pedido," & vbCrLf & _
              "								   NF.Procuracao," & vbCrLf & _
              "								   SUM(nfi.QuantidadeFisica) AS Quantidade" & vbCrLf & _
              "							  FROM NotasFiscais NF" & vbCrLf & _
              "							 INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
              "								ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
              "							   AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
              "							   AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
              "							   AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
              "							   AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
              "							   AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
              "							   AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
              "							 Inner join SubOperacoes SO" & vbCrLf & _
              "								on so.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "							   and so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf & _
              "							 Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf & _
              "							   and NF.situacao in (1,4,7)" & vbCrLf & _
              "							 GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Procuracao" & vbCrLf & _
              "							 Union All" & vbCrLf & _
              "							select P.Empresa_Id," & vbCrLf & _
              "									P.EndEmpresa_Id," & vbCrLf & _
              "									P.Pedido_id," & vbCrLf & _
              "									PIF.Procuracao," & vbCrLf & _
              "									sum(pif.Quantidade)" & vbCrLf & _
              "							   from Pedidos P" & vbCrLf & _
              "							  Inner join PedidosXItensXFixacoes PIF" & vbCrLf & _
              "								 on P.Empresa_Id    = PIF.Empresa_Id" & vbCrLf & _
              "								and P.EndEmpresa_Id = PIF.EndEmpresa_Id" & vbCrLf & _
              "								and p.Pedido_Id     = PIF.Pedido_Id" & vbCrLf & _
              "							  Where P.Situacao = 1" & vbCrLf & _
              "							  Group by P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_id, PIF.Procuracao" & vbCrLf & _
              "							  ) sb" & vbCrLf & _
              "						Group by sb.Empresa_Id, sb.EndEmpresa_Id, sb.Pedido, sb.Procuracao" & vbCrLf & _
              "                    ) AS sb_Real" & vbCrLf & _
              "                ON sb_Real.Empresa_Id    = PR.Empresa_Id" & vbCrLf & _
              "               AND sb_Real.EndEmpresa_Id = PR.EndEmpresa_Id" & vbCrLf & _
              "               AND sb_Real.Pedido        = PR.PedidoCedente" & vbCrLf & _
              "               AND sb_Real.Procuracao    = PR.Procuracao_ID" & vbCrLf & _
              "  WHERE PR.Situacao      = 1 " & vbCrLf & _
              "    AND PR.PedidoCedente = " & pPedido & vbCrLf & _
              "  order by PR.Movimento desc" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Procuracoes")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim proc As New Procuracao

            proc.CodigoEmpresa = row("Empresa_Id")
            proc.EnderecoEmpresa = row("EndEmpresa_Id")
            proc.Codigo = row("Procuracao_ID")
            proc.Documento = row("Documento")
            proc.CodigoCedente = row("Cedente")
            proc.EnderecoCedente = row("EndCedente")
            proc.CodigoPedidoCedente = row("PedidoCedente")
            proc.CodigoCessionario = row("Cessionario")
            proc.EnderecoCessionario = row("EndCessionario")
            proc.Movimento = row("Movimento")
            proc.Quantidade = row("Quantidade")
            proc.Observacoes = row("Observacoes")
            proc.UsuarioInclusao = row("UsuarioInclusao")
            proc.UsuarioInclusaoData = row("UsuarioInclusaoDate")
            proc.UsuarioAlteracao = row("UsuarioAlteracao")
            proc.UsuarioAlteracaoData = row("UsuarioAlteracaoDate")
            proc.UsuarioCancelamento = row("UsuarioCancelamento")
            proc.UsuarioCancelamentoData = row("UsuarioCancelamentoDate")
            proc.CodigoSituacao = row("Situacao")

            _QuantidadeProcuracoes += row("Quantidade")
            proc.SetarValores(row("Realizado"), row("Saldo"))
            _QuantidadeEntregueProcuracoes += row("Realizado")
            _QuantidadeSaldoProcuracoes += proc.Saldo

            If zerado Then
                Me.Add(proc)
            ElseIf proc.Saldo > 0 Then
                Me.Add(proc)
            End If
        Next
    End Sub

#End Region

#Region "Fields"
    Private _QuantidadeProcuracoes As Double = 0
    Private _QuantidadeEntregueProcuracoes As Double = 0
    Private _QuantidadeSaldoProcuracoes As Double = 0
#End Region

#Region "Property"
    Public Property QuantidadeProcuracoes() As Double
        Get
            Return _QuantidadeProcuracoes
        End Get
        Set(ByVal value As Double)
            _QuantidadeProcuracoes = value
        End Set
    End Property

    Public Property QuantidadeEntregueProcuracoes() As Double
        Get
            Return _QuantidadeEntregueProcuracoes
        End Get
        Set(ByVal value As Double)
            _QuantidadeEntregueProcuracoes = value
        End Set
    End Property

    Public Property QuantidadeSaldoProcuracoes() As Double
        Get
            Return _QuantidadeSaldoProcuracoes
        End Get
        Set(ByVal value As Double)
            _QuantidadeSaldoProcuracoes = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Shared Function Existe(ByVal Pedido As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT 1 " & _
                                   "  FROM Procuracoes " & _
                                   " WHERE Situacao = 1" & _
                                   "   AND (PedidoCedente     =    " & Pedido.ToString() & _
                                   "       OR" & _
                                   "       PedidoCessionario = " & Pedido.ToString() & ")"

            Dim dsProcuracoes As DataSet = objBanco.ConsultaDataSet(strSQL, "Procuracoes")

            If dsProcuracoes.Tables(0).Rows.Count > 0 Then Return True Else Return False
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function
#End Region

End Class

<Serializable()> _
Public Class Procuracao
    Private Sql As String
    Private Banco As New AcessaBanco

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal Procuracao As Integer)
        Dim ds As New DataSet

        Sql = " SELECT pr.Empresa_Id, pr.EndEmpresa_Id," & vbCrLf & _
              "        pr.Procuracao_ID, pr.Situacao, pr.Documento," & vbCrLf & _
              "		   pr.Cedente, pr.EndCedente, pr.PedidoCedente," & vbCrLf & _
              "		   pr.Cessionario, pr.EndCessionario," & vbCrLf & _
              "        pr.Movimento," & vbCrLf & _
              "        isnull(pr.Observacoes, '') as Observacoes, isnull(pr.UsuarioInclusao, '') AS UsuarioInclusao," & vbCrLf & _
              "        isnull(pr.UsuarioInclusaoDate, '') AS UsuarioInclusaoDate," & vbCrLf & _
              "        isnull(pr.UsuarioAlteracao, '') AS UsuarioAlteracao," & vbCrLf & _
              "        isnull(pr.UsuarioAlteracaoDate, '') AS UsuarioAlteracaoDate," & vbCrLf & _
              "        isnull(pr.UsuarioCancelamento, '') AS UsuarioCancelamento," & vbCrLf & _
              "        isnull(pr.UsuarioCancelamentoDate, '') AS UsuarioCancelamentoDate," & vbCrLf & _
              "        pr.Quantidade," & vbCrLf & _
              "		   ISNULL(sb_Real.Quantidade, 0) AS Realizado," & vbCrLf & _
              "        PR.Quantidade - ISNULL(sb_Real.Quantidade, 0) AS Saldo" & vbCrLf & _
              "   FROM Procuracoes PR" & vbCrLf & _
              "   LEFT OUTER JOIN (Select sb.Empresa_Id," & vbCrLf & _
              "                           sb.EndEmpresa_Id," & vbCrLf & _
              "						      sb.Pedido," & vbCrLf & _
              "						      sb.Procuracao," & vbCrLf & _
              "						      SUM(sb.Quantidade) AS Quantidade" & vbCrLf & _
              "                     from (" & vbCrLf & _
              "							  SELECT NF.Empresa_Id," & vbCrLf & _
              "								     NF.EndEmpresa_Id," & vbCrLf & _
              "								     NF.Pedido," & vbCrLf & _
              "								     NF.Procuracao," & vbCrLf & _
              "								     SUM(nfi.QuantidadeFisica) AS Quantidade" & vbCrLf & _
              "							    FROM NotasFiscais NF" & vbCrLf & _
              "							   INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
              "							      ON NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
              "							     AND NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
              "							     AND NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
              "							     AND NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
              "							     AND NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
              "							     AND NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
              "							     AND NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
              "							   Inner join SubOperacoes SO" & vbCrLf & _
              "								  on so.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "							     and so.SubOperacoes_Id = NF.SubOperacao" & vbCrLf & _
              "							   Where so.Classe <> '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "'" & vbCrLf & _
              "							     and NF.situacao in (1,4,7)" & vbCrLf & _
              "							   GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Pedido, NF.Procuracao" & vbCrLf & _
              "							   Union All" & vbCrLf & _
              "							  Select P.Empresa_Id," & vbCrLf & _
              "								 	 P.EndEmpresa_Id," & vbCrLf & _
              "									 P.Pedido_id," & vbCrLf & _
              "									 PIF.Procuracao," & vbCrLf & _
              "									 sum(pif.Quantidade)" & vbCrLf & _
              "							    from Pedidos P" & vbCrLf & _
              "							   Inner join PedidosXItensXFixacoes PIF" & vbCrLf & _
              "								  on P.Empresa_Id    = PIF.Empresa_Id" & vbCrLf & _
              "								 and P.EndEmpresa_Id = PIF.EndEmpresa_Id" & vbCrLf & _
              "								 and p.Pedido_Id     = PIF.Pedido_Id" & vbCrLf & _
              "							   Where P.Situacao = 1" & vbCrLf & _
              "							   Group by P.Empresa_Id, P.EndEmpresa_Id, P.Pedido_id, PIF.Procuracao" & vbCrLf & _
              "							  ) sb" & vbCrLf & _
              "							 Group by sb.Empresa_Id, sb.EndEmpresa_Id, sb.Pedido, sb.Procuracao" & vbCrLf & _
              "                    ) AS sb_Real" & vbCrLf & _
              "                ON sb_Real.Empresa_Id    = PR.Empresa_Id" & vbCrLf & _
              "               AND sb_Real.EndEmpresa_Id = PR.EndEmpresa_Id" & vbCrLf & _
              "               AND sb_Real.Pedido        = PR.PedidoCedente" & vbCrLf & _
              "               AND sb_Real.Procuracao    = PR.Procuracao_ID" & vbCrLf & _
              "  WHERE PR.Empresa_id    = '" & Empresa & "'" & vbCrLf & _
              "    AND PR.EndEmpresa_id = " & EndEmpresa & vbCrLf & _
              "    AND PR.Procuracao_ID = " & Procuracao & vbCrLf & _
              "  order by PR.Movimento desc" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Procuracao")

        For Each row As DataRow In ds.Tables(0).Rows
            CodigoEmpresa = row("Empresa_Id")
            EnderecoEmpresa = row("EndEmpresa_Id")
            Codigo = row("Procuracao_ID")
            CodigoSituacao = row("Situacao")
            Documento = row("Documento")
            CodigoCedente = row("Cedente")
            EnderecoCedente = row("EndCedente")
            CodigoPedidoCedente = row("PedidoCedente")
            CodigoCessionario = row("Cessionario")
            EnderecoCessionario = row("EndCessionario")
            Movimento = row("Movimento")
            Quantidade = row("Quantidade")
            Observacoes = row("Observacoes")
            UsuarioInclusao = row("UsuarioInclusao")
            UsuarioInclusaoData = row("UsuarioInclusaoDate")
            UsuarioAlteracao = row("UsuarioAlteracao")
            UsuarioAlteracaoData = row("UsuarioAlteracaoDate")
            UsuarioCancelamento = row("UsuarioCancelamento")
            UsuarioCancelamentoData = row("UsuarioCancelamentoDate")
            _Realizado = row("Realizado")
            _Saldo = row("Saldo")
        Next

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente
    Private _Codigo As Integer
    Private _Documento As String = ""
    Private _CodigoCedente As String = ""
    Private _EnderecoCedente As Integer
    Private _Cedente As Cliente
    Private _CodigoPedidoCedente As Integer
    Private _PedidoCedente As Pedido
    Private _CodigoCessionario As String = ""
    Private _EnderecoCessionario As Integer
    Private _Cessionario As Cliente
    Private _Movimento As DateTime = Date.Today
    Private _Quantidade As Double = 0
    Private _Observacoes As String = ""
    Private _UsuarioInclusao As String = ""
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _UsuarioAlteracaoData As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _UsuarioCancelamentoData As DateTime
    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao
    Private _Realizado As Double = 0
    Private _Saldo As Double = 0
    Private _ProdutosValidos As Boolean
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
            If _Empresa Is Nothing AndAlso Me.CodigoEmpresa > 0 Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EnderecoEmpresa)
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
        End Set
    End Property

    Public Property Documento() As String
        Get
            Return _Documento
        End Get
        Set(ByVal value As String)
            _Documento = value
        End Set
    End Property

    Public Property CodigoCedente() As String
        Get
            Return _CodigoCedente
        End Get
        Set(ByVal value As String)
            _CodigoCedente = value
        End Set
    End Property

    Public Property EnderecoCedente() As Integer
        Get
            Return _EnderecoCedente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCedente = value
        End Set
    End Property

    Public Property Cedente() As Cliente
        Get
            If _Cedente Is Nothing AndAlso Me.CodigoCedente > 0 Then _Cedente = New Cliente(Me.CodigoCedente, Me.EnderecoCedente)
            Return _Cedente
        End Get
        Set(ByVal value As Cliente)
            _Cedente = value
        End Set
    End Property

    Public Property CodigoPedidoCedente() As Integer
        Get
            Return _CodigoPedidoCedente
        End Get
        Set(ByVal value As Integer)
            _CodigoPedidoCedente = value
        End Set
    End Property

    Public Property PedidoCedente() As Pedido
        Get
            If (_PedidoCedente Is Nothing OrElse _PedidoCedente.Codigo = 0) AndAlso Me.CodigoPedidoCedente > 0 _
                Then _PedidoCedente = New Pedido(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.CodigoPedidoCedente)
            Return _PedidoCedente
        End Get
        Set(ByVal value As Pedido)

        End Set
    End Property

    Public Property CodigoCessionario() As String
        Get
            Return _CodigoCessionario
        End Get
        Set(ByVal value As String)
            _CodigoCessionario = value
        End Set
    End Property

    Public Property EnderecoCessionario() As Integer
        Get
            Return _EnderecoCessionario
        End Get
        Set(ByVal value As Integer)
            _EnderecoCessionario = value
        End Set
    End Property

    Public Property Cessionario() As Cliente
        Get
            If _Cessionario Is Nothing AndAlso Me.CodigoCessionario.Trim.Length > 0 Then _Cessionario = New Cliente(Me.CodigoCessionario, Me.EnderecoCessionario)
            Return _Cessionario
        End Get
        Set(ByVal value As Cliente)
            _Cessionario = value
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

    Public Property Quantidade() As Double
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Double)
            _Quantidade = value
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

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData() As DateTime
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

    Public Property UsuarioAlteracaoData() As DateTime
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

    Public Property UsuarioCancelamentoData() As DateTime
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioCancelamentoData = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
        End Set
    End Property

    Public Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
        End Set
    End Property

    Public ReadOnly Property Realizado() As Double
        Get
            Return _Realizado
        End Get
    End Property

    Public ReadOnly Property Saldo() As Double
        Get
            Return _Saldo
        End Get
    End Property

    Public ReadOnly Property NomeCessionario() As String
        Get
            Return Cessionario.Nome
        End Get
    End Property

    Public Property ProdutosValidos() As Boolean
        Get
            Return _ProdutosValidos
        End Get
        Set(ByVal value As Boolean)
            _ProdutosValidos = value
        End Set
    End Property
#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Select Case Me.IUD
            Case "I"
                Dim n As Numerador
                n = New Numerador(CodigoEmpresa, EnderecoEmpresa, 9)
                _Codigo = n.Sequencia + 1

                Sqls.Add(n.IncrementarNumeradorSql)

                Sql = " INSERT INTO Procuracoes (Empresa_Id, EndEmpresa_Id, Procuracao_ID, Documento, Cedente, EndCedente, PedidoCedente," & vbCrLf & _
                      "             Cessionario, EndCessionario, Movimento, Quantidade," & vbCrLf & _
                      "             Observacoes, UsuarioInclusao, UsuarioInclusaoDate, Situacao) " & vbCrLf & _
                      " VALUES ('" & CodigoEmpresa & "'," & EnderecoEmpresa & "," & Codigo & ",'" & Documento & "','" & CodigoCedente & "'," & EnderecoCedente & "," & vbCrLf & _
                      "" & CodigoPedidoCedente & ",'" & CodigoCessionario & "'," & EnderecoCessionario & ",'" & CDate(Movimento).ToSqlDate & "'," & Quantidade & ",'" & Observacoes & "'," & vbCrLf & _
                      "'" & UsuarioInclusao & "','" & CDate(UsuarioInclusaoData).ToSqlDate & "',1)"
                Sqls.Add(Sql)
            Case "U"
                Sql = "UPDATE Procuracoes SET" & vbCrLf & _
                      "   Documento            ='" & Documento & "'" & vbCrLf & _
                      "  ,Cedente              ='" & CodigoCedente & "'" & vbCrLf & _
                      "  ,EndCedente           = " & EnderecoCedente & vbCrLf & _
                      "  ,PedidoCedente        = " & CodigoPedidoCedente & vbCrLf & _
                      "  ,Cessionario          ='" & CodigoCessionario & "'" & vbCrLf & _
                      "  ,EndCessionario       = " & EnderecoCessionario & vbCrLf & _
                      "  ,Movimento            ='" & CDate(Movimento).ToSqlDate & "'" & vbCrLf & _
                      "  ,Quantidade           = " & Quantidade & vbCrLf & _
                      "  ,Observacoes          ='" & Observacoes & "'" & vbCrLf & _
                      "  ,UsuarioAlteracao     ='" & UsuarioAlteracao & "'" & vbCrLf & _
                      "  ,UsuarioAlteracaoDate ='" & CDate(UsuarioAlteracaoData).ToSqlDate & "'" & vbCrLf & _
                      " WHERE Empresa_Id    ='" & CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & EnderecoEmpresa & vbCrLf & _
                      "   AND Procuracao_ID = " & Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " UPDATE Procuracoes SET" & vbCrLf & _
                      "    UsuarioCancelamento     ='" & UsuarioCancelamento & "'" & vbCrLf & _
                      "   ,UsuarioCancelamentoDate ='" & CDate(UsuarioCancelamentoData).ToSqlDate & "'" & vbCrLf & _
                      "   ,Situacao = 3 " & vbCrLf & _
                      " WHERE Empresa_Id = '" & CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & EnderecoEmpresa & vbCrLf & _
                      "   AND Procuracao_ID = " & Codigo
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Sub SetarValores(ByVal realizado As Decimal, ByVal saldo As Decimal)
        _Realizado = realizado
        _Saldo = saldo
    End Sub

#End Region

End Class