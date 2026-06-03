Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaDeExportacao
    Inherits List(Of NotaDeExportacao)

#Region "Construtor"
    Public Sub New(ByVal Men As Negocio.MemorandoDeExportacao, Optional ByVal Safra As String = "")
        Dim sql As String
        sql = "Select NFxI.Empresa_Id," & vbCrLf & _
              "       NFxI.EndEmpresa_Id," & vbCrLf & _
              "       NFxI.Cliente_Id," & vbCrLf & _
              "       NFxI.EndCliente_Id," & vbCrLf & _
              "       NFxI.EntradaSaida_Id," & vbCrLf & _
              "       NFxI.Produto_Id," & vbCrLf & _
              "       NFxI.Nota_Id," & vbCrLf & _
              "       NFxI.Serie_Id," & vbCrLf & _
              "       NFxI.QuantidadeFiscal as QtdeNota," & vbCrLf & _
              "       NFxI.Valor as ValorNota," & vbCrLf & _
              "       isnull(sbMem.QtdeComprovadaNesteMemorando,0) as QtdeComprovadaNesteMemorando," & vbCrLf & _
              "       isnull(sbMem.QtdeJaComprovada,0) as QtdeJaComprovada" & vbCrLf & _
              "  from notasfiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " inner join Operacoes OP" & vbCrLf & _
              "    on OP.Operacao_id = NFxI.Operacao" & vbCrLf & _
              " inner Join SubOperacoes SO" & vbCrLf & _
              "    on SO.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "   and SO.SubOperacoes_Id = NFxI.SubOperacao" & vbCrLf & _
              " Inner Join Pedidos P" & vbCrLf & _
              "    on P.Empresa_id    = NF.Empresa_Id" & vbCrLf & _
              "   and P.EndEmpresa_id = NF.EndEmpresa_Id" & vbCrLf & _
              "   and P.Pedido_id     = NF.Pedido" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "			 SELECT M.Empresa," & vbCrLf & _
              "			 	    M.EndEmpresa," & vbCrLf & _
              "					M.Cliente," & vbCrLf & _
              "					M.EndCliente," & vbCrLf & _
              "					M.EntradaSaida," & vbCrLf & _
              "					M.Nota," & vbCrLf & _
              "					M.Serie," & vbCrLf & _
              "					M.Produto," & vbCrLf

        If Not Men Is Nothing Then
            sql &= "					SUM(case" & vbCrLf & _
                   "					      when M.EmpresaMemorando_Id    ='" & Men.CodigoEmpresaMemorando & "'" & vbCrLf & _
                   "						   and M.EndEmpresaMemorando_Id = " & Men.EnderecoEmpresaMemorando & vbCrLf & _
                   "						   and M.Memorando_Id           = '" & Men.CodigoMemorando & "'" & vbCrLf & _
                   "						    Then MxNF.Quantidade" & vbCrLf & _
                   "						    Else 0" & vbCrLf & _
                   "						End) QtdeComprovadaNesteMemorando," & vbCrLf
        Else
            sql &= "                  0 as QtdeComprovadaNesteMemorando," & vbCrLf
        End If


        sql &= "					SUM(MxNF.Quantidade) QtdeJaComprovada" & vbCrLf & _
               "			   FROM MemorandoDeExportacao M" & vbCrLf & _
               "			  INNER JOIN MemorandoDeExportacaoXNotaFiscal MxNF" & vbCrLf & _
               "			     ON M.EmpresaMemorando_Id    = MxNF.EmpresaMemorando_Id" & vbCrLf & _
               "			    AND M.EndEmpresaMemorando_Id = MxNF.EndEmpresaMemorando_Id" & vbCrLf & _
               "			    AND M.Memorando_Id           = MxNF.Memorando_Id" & vbCrLf & _
               "              Group by M.Empresa," & vbCrLf & _
               "			 	       M.EndEmpresa," & vbCrLf & _
               "					   M.Cliente," & vbCrLf & _
               "					   M.EndCliente," & vbCrLf & _
               "					   M.EntradaSaida," & vbCrLf & _
               "					   M.Nota," & vbCrLf & _
               "					   M.Serie," & vbCrLf & _
               "					   M.Produto" & vbCrLf & _
               "             )sbMem" & vbCrLf & _
               "    ON NFxI.Empresa_Id      = sbMem.Empresa " & vbCrLf & _
               "   AND NFxI.EndEmpresa_Id   = sbMem.EndEmpresa" & vbCrLf & _
               "   AND NFxI.Cliente_Id      = sbMem.Cliente" & vbCrLf & _
               "   AND NFxI.EndCliente_Id   = sbMem.EndCliente" & vbCrLf & _
               "   AND NFxI.EntradaSaida_Id = sbMem.EntradaSaida" & vbCrLf & _
               "   AND NFxI.Serie_Id        = sbMem.Serie" & vbCrLf & _
               "   AND NFxI.Nota_Id         = sbMem.Nota" & vbCrLf & _
               "   --AND NFxI.Produto_Id      = sbMem.Produto" & vbCrLf & _
               " where OP.Classe             = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
               "   and SO.Memorando          = 1" & vbCrLf & _
               "   and SO.Classe             ='" & eClassesOperacoes.EXPORTACOES.ToString & "'" & vbCrLf & _
               "   and NF.Situacao            in (1,4,7) " & vbCrLf & _
               "   and NF.TipoDeDocumento    = 1 " & vbCrLf & _
               "   and NFxI.QuantidadeFiscal > 0" & vbCrLf

        If Safra.Length > 0 Then
            sql &= "   and P.Safra = '" & Safra & "'" & vbCrLf
        End If

        If Not Men Is Nothing Then
            sql &= "   and NFxI.Empresa_Id       ='" & Men.CodigoEmpresaMemorando & "'" & vbCrLf & _
                   "   and NFxI.EndEmpresa_Id    = " & Men.EnderecoEmpresaMemorando & vbCrLf & _
                   "   and LEFT(NFxI.Produto_Id,5)       ='" & Men.CodigoProduto.Substring(0, 5) & "'" & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim NE As New Negocio.NotaDeExportacao()
            NE.Memorando = Men
            NE.CodigoEmpresa = row("Empresa_Id")
            NE.EnderecoEmpresa = row("EndEmpresa_Id")
            NE.CodigoCliente = row("Cliente_Id")
            NE.EnderecoCliente = row("EndCliente_Id")
            NE.NumeroNota = row("Nota_Id")
            NE.Serie = row("Serie_Id")
            NE.EntradaSaida = row("EntradaSaida_Id")
            NE.CodigoProduto = row("Produto_Id")
            NE.ValorNota = row("ValorNota")
            NE.QtdeNota = row("QtdeNota")
            NE.QtdeComprovadaNesteMemorando = row("QtdeComprovadaNesteMemorando")
            NE.QtdeJaComprovada = row("QtdeJaComprovada")
            Me.Add(NE)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class NotaDeExportacao

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Men As Negocio.MemorandoDeExportacao)
        _Memorando = Men
        Dim sql As String
        sql = "Select NFxI.Empresa_Id," & vbCrLf & _
              "       NFxI.EndEmpresa_Id," & vbCrLf & _
              "       NFxI.Cliente_Id," & vbCrLf & _
              "       NFxI.EndCliente_Id," & vbCrLf & _
              "       NFxI.EntradaSaida_Id," & vbCrLf & _
              "       NFxI.Produto_Id," & vbCrLf & _
              "       NFxI.Nota_Id," & vbCrLf & _
              "       NFxI.Serie_Id," & vbCrLf & _
              "       NFxI.QuantidadeFiscal as QtdeNota," & vbCrLf & _
              "       NFxI.Valor as ValorNota," & vbCrLf & _
              "       isnull(sbMem.QtdeComprovadaNesteMemorando,0) as QtdeComprovadaNesteMemorando," & vbCrLf & _
              "       isnull(sbMem.QtdeJaComprovada,0) as QtdeJaComprovada" & vbCrLf & _
              "  from notasfiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " inner join Operacoes OP" & vbCrLf & _
              "    on OP.Operacao_id = NFxI.Operacao" & vbCrLf & _
              " inner Join SubOperacoes SO" & vbCrLf & _
              "    on SO.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "   and SO.SubOperacoes_Id = NFxI.SubOperacao" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "			 SELECT M.Empresa," & vbCrLf & _
              "			 	    M.EndEmpresa," & vbCrLf & _
              "					M.Cliente," & vbCrLf & _
              "					M.EndCliente," & vbCrLf & _
              "					M.EntradaSaida," & vbCrLf & _
              "					M.Nota," & vbCrLf & _
              "					M.Serie," & vbCrLf & _
              "					M.Produto," & vbCrLf & _
              "					SUM(case" & vbCrLf & _
              "					      when M.EmpresaMemorando_Id    ='" & Men.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "						   and M.EndEmpresaMemorando_Id = " & Men.EnderecoEmpresaMemorando & vbCrLf & _
              "						   and M.Memorando_Id           = " & Men.CodigoMemorando & vbCrLf & _
              "						    Then MxNF.Quantidade" & vbCrLf & _
              "						    Else 0" & vbCrLf & _
              "						End) QtdeComprovadaNesteMemorando," & vbCrLf & _
              "					SUM(MxNF.Quantidade) QtdeJaComprovada" & vbCrLf & _
              "			   FROM MemorandoDeExportacao M" & vbCrLf & _
              "			  INNER JOIN MemorandoDeExportacaoXNotaFiscal MxNF" & vbCrLf & _
              "			     ON M.EmpresaMemorando_Id    = MxNF.EmpresaMemorando_Id" & vbCrLf & _
              "			    AND M.EndEmpresaMemorando_Id = MxNF.EndEmpresaMemorando_Id" & vbCrLf & _
              "			    AND M.Memorando_Id           = MxNF.Memorando_Id" & vbCrLf & _
              "              Group by M.Empresa," & vbCrLf & _
              "			 	       M.EndEmpresa," & vbCrLf & _
              "					   M.Cliente," & vbCrLf & _
              "					   M.EndCliente," & vbCrLf & _
              "					   M.EntradaSaida," & vbCrLf & _
              "					   M.Nota," & vbCrLf & _
              "					   M.Serie," & vbCrLf & _
              "					   M.Produto" & vbCrLf & _
              "             )sbMem" & vbCrLf & _
              "    ON NFxI.Empresa_Id      = sbMem.Empresa " & vbCrLf & _
              "   AND NFxI.EndEmpresa_Id   = sbMem.EndEmpresa" & vbCrLf & _
              "   AND NFxI.Cliente_Id      = sbMem.Cliente" & vbCrLf & _
              "   AND NFxI.EndCliente_Id   = sbMem.EndCliente" & vbCrLf & _
              "   AND NFxI.EntradaSaida_Id = sbMem.EntradaSaida" & vbCrLf & _
              "   AND NFxI.Serie_Id        = sbMem.Serie" & vbCrLf & _
              "   AND NFxI.Nota_Id         = sbMem.Nota" & vbCrLf & _
              "   AND NFxI.Produto_Id      = sbMem.Produto" & vbCrLf & _
              " Where OP.Classe             = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
              "   and SO.Memorando          = 1" & vbCrLf & _
              "   and SO.Classe             ='" & eClassesOperacoes.EXPORTACOES.ToString & "'" & vbCrLf & _
              "   and NF.Situacao            in (1,4,7) " & vbCrLf & _
              "   and NF.TipoDeDocumento    = 1 " & vbCrLf & _
              "   and NFxI.Empresa_Id       ='" & Men.CodigoEmpresa & "'" & vbCrLf & _
              "   and NFxI.EndEmpresa_Id    = " & Men.EnderecoEmpresa & vbCrLf & _
              "   and NFxI.Cliente_Id       ='" & Men.CodigoCliente & "'" & vbCrLf & _
              "   and NFxI.EndCliente_Id    = " & Men.EnderecoCliente & vbCrLf & _
              "   and NFxI.EntradaSaida_Id  ='" & Men.EntradaSaida & "'" & vbCrLf & _
              "   and NFxI.Nota_Id          = " & Men.NumeroNota & vbCrLf & _
              "   and NFxI.Serie_Id         ='" & Men.Serie & "'" & vbCrLf & _
              "   and NFxI.Produto_Id       ='" & Men.CodigoProduto & "'" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Notas")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoEmpresa = row("Empresa_Id")
            _EnderecoEmpresa = row("EndEmpresa_Id")
            _CodigoCliente = row("Cliente_Id")
            _EnderecoCliente = row("EndCliente_Id")
            _NumeroNota = row("Nota_Id")
            _Serie = row("Serie_Id")
            _EntradaSaida = row("EntradaSaida_Id")
            _CodigoProduto = row("Produto_Id")
            _ValorNota = row("ValorNota")
            _QtdeNota = row("QtdeNota")
            _QtdeComprovadaNesteMemorando = row("QtdeComprovadaNesteMemorando")
            _QtdeJaComprovada = row("QtdeJaComprovada")
        End If

    End Sub
#End Region

#Region "Fields"
    Private _Memorando As Negocio.MemorandoDeExportacao
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Negocio.Cliente

    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Cliente As Negocio.Cliente

    Private _NumeroNota As Integer
    Private _Serie As String = ""
    Private _EntradaSaida As String = ""

    Private _CodigoProduto As String = ""
    Private _Produto As Negocio.Produto

    Private _NotaFiscal As Negocio.NotaFiscal

    Private _QtdeNota As Decimal

    Private _ValorNota As Decimal

    Private _QtdeComprovadaNesteMemorando As Decimal
    Private _QtdeJaComprovada As Decimal
#End Region

#Region "Property"

    Public Property Memorando() As Negocio.MemorandoDeExportacao
        Get
            Return _Memorando
        End Get
        Set(ByVal value As Negocio.MemorandoDeExportacao)
            _Memorando = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
            _Empresa = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public ReadOnly Property Empresa() As Negocio.Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Negocio.Cliente(_CodigoEmpresa, _EnderecoEmpresa)
            Return _Empresa
        End Get
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
            _Cliente = Nothing
            _NotaFiscal = Nothing
        End Set
    End Property

    Public ReadOnly Property Cliente() As Negocio.Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Negocio.Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value.ToUpper
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property NumeroNota() As Integer
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As Integer)
            _NumeroNota = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public ReadOnly Property Produto() As Negocio.Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Negocio.Produto(_CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public ReadOnly Property NomeProduto() As String
        Get
            Return Produto.Nome
        End Get
    End Property

    Public ReadOnly Property NotaFiscal() As Negocio.NotaFiscal
        Get
            If _NotaFiscal Is Nothing Then
                Dim NFConsulta As New Negocio.NotaFiscal
                NFConsulta.CodigoEmpresa = _CodigoEmpresa
                NFConsulta.EnderecoEmpresa = _EnderecoEmpresa
                NFConsulta.CodigoCliente = _CodigoCliente
                NFConsulta.EnderecoCliente = _EnderecoCliente
                NFConsulta.Codigo = _NumeroNota
                NFConsulta.Serie = _Serie
                NFConsulta.EntradaSaida = IIf(_EntradaSaida = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                _NotaFiscal = New Negocio.NotaFiscal(NFConsulta)
            End If
            Return _NotaFiscal
        End Get
    End Property

    Public Property QtdeNota() As Decimal
        Get
            Return _QtdeNota
        End Get
        Set(ByVal value As Decimal)
            _QtdeNota = value
        End Set
    End Property

    Public Property QtdeComprovadaNesteMemorando() As Decimal
        Get
            Return _QtdeComprovadaNesteMemorando
        End Get
        Set(ByVal value As Decimal)
            _QtdeComprovadaNesteMemorando = value
        End Set
    End Property

    Public Property QtdeJaComprovada() As Decimal
        Get
            Return _QtdeJaComprovada
        End Get
        Set(ByVal value As Decimal)
            _QtdeJaComprovada = value
        End Set
    End Property

    Public Property ValorNota() As Decimal
        Get
            Return _ValorNota
        End Get
        Set(ByVal value As Decimal)
            _ValorNota = value
        End Set
    End Property

#End Region

End Class