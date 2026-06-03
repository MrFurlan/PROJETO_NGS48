Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMemorandoDeExportacaoXNotaFiscal
    Inherits List(Of MemorandoDeExportacaoXNotaFiscal)

#Region "Fields"
    Private _Memorando As MemorandoDeExportacao
#End Region

#Region "Property"
    Public Property Memorando() As MemorandoDeExportacao
        Get
            Return _Memorando
        End Get
        Set(ByVal value As MemorandoDeExportacao)
            _Memorando = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New(ByVal Mem As MemorandoDeExportacao)
        _Memorando = Mem

        CarregarNotasParaSelecao("", "", True)

        'Dim Sql As String
        'Sql = "SELECT EmpresaMemorando_Id, EndEmpresaMemorando_Id, Memorando_Id, Empresa_Id, EndEmpresa_Id," & vbCrLf & _
        '      "       Cliente_Id, EndCliente_Id, Nota_Id, Serie_Id, EntradaSaida_Id, Quantidade " & vbCrLf & _
        '      "  FROM MemorandoDeExportacaoXNotaFiscal" & vbCrLf & _
        '      " Where EmpresaMemorando_Id    ='" & Mem.CodigoEmpresaMemorando & "'" & vbCrLf & _
        '      "   and EndEmpresaMemorando_Id = " & Mem.EnderecoEmpresaMemorando & vbCrLf & _
        '      "   and Memorando_Id           = '" & Mem.CodigoMemorando & "'" & vbCrLf
        'Dim Banco As New AcessaBanco
        'Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Notas")

        'For Each row As DataRow In ds.Tables(0).Rows
        '    Dim MMxNF As New MemorandoDeExportacaoXNotaFiscal(Mem)
        '    MMxNF.CodigoEmpresa = row("Empresa_Id")
        '    MMxNF.EnderecoEmpresa = row("EndEmpresa_Id")
        '    MMxNF.CodigoCliente = row("Cliente_Id")
        '    MMxNF.EnderecoCliente = row("EndCliente_Id")
        '    MMxNF.NumeroNota = row("Nota_Id")
        '    MMxNF.Serie = row("Serie_Id")
        '    MMxNF.EntradaSaida = row("EntradaSaida_Id")
        '    MMxNF.QuantidadeMemorando = row("Quantidade")

        '    Me.Add(MMxNF)
        'Next

        'Banco = Nothing
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As MemorandoDeExportacaoXNotaFiscal In Me
            If Memorando.IUD = "D" Or Memorando.IUD = "I" Then item.IUD = Memorando.IUD
            If item.IUD <> "" And item.QuantidadeMemorando > 0 Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Sub CarregarNotasParaSelecao(Optional ByVal Numnotas As String = "", Optional ByVal Safra As String = "", Optional ByVal Alteracao As Boolean = False)
        Dim strSQL As String = ""
        strSQL = "SELECT NF.Empresa_Id," & vbCrLf & _
                 "       NF.EndEmpresa_Id," & vbCrLf & _
                 "       NF.Cliente_Id," & vbCrLf & _
                 "       NF.EndCliente_Id," & vbCrLf & _
                 "       NF.EntradaSaida_Id," & vbCrLf & _
                 "       NF.Serie_Id," & vbCrLf & _
                 "       NF.Nota_Id," & vbCrLf & _
                 "       NF.DataDaNota," & vbCrLf & _
                 "       NFxI.Produto_Id," & vbCrLf & _
                 "       NFxI.Sequencia_Id," & vbCrLf & _
                 "       NFxI.CFOP_Id," & vbCrLf & _
                 "       NFxI.Unitario as UnitarioNota," & vbCrLf & _
                 "       NFxI.QuantidadeFiscal as QuantidadeNota," & vbCrLf & _
                 "       isnull(sbDev.Quantidade,0) as QtdeDevolvida," & vbCrLf & _
                 "       isnull(sbMem.Quantidade,0) as QtdeUsadaMemorando," & vbCrLf & _
                 "       isnull(sbMemNota.Quantidade,0) as QtdeUsadaNesteMemorando" & vbCrLf & _
                 "  FROM NotasFiscais AS NF" & vbCrLf & _
                 " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                 "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                 "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                 "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                 "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                 "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                 "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                 "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                 " INNER JOIN SubOperacoes SO" & vbCrLf & _
                 "    ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf & _
                 "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                 " INNER JOIN PEDIDOS P" & vbCrLf & _
                 "    ON P.Empresa_id    = NF.Empresa_Id" & vbCrLf & _
                 "   AND P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf & _
                 "   AND P.Pedido_id     = NF.Pedido" & vbCrLf & _
                 "  Left Join (" & vbCrLf & _
                 "		 	 SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
                 "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
                 "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
                 "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
                 "                    nfd.EntradaSaida_Id," & vbCrLf & _
                 "                    nfd.Serie_Id," & vbCrLf & _
                 "                    nfd.Nota_Id," & vbCrLf & _
                 "                    nfd.Produto_Id," & vbCrLf & _
                 "                    nfd.Sequencia_Id," & vbCrLf & _
                 "                    nfd.CFOP_Id," & vbCrLf & _
                 "                    sum(nfd.Quantidade) as Quantidade" & vbCrLf & _
                 "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
                 "              Inner Join NotasFiscais nf" & vbCrLf & _
                 "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
                 "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
                 "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
                 "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
                 "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf & _
                 "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf & _
                 "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf & _
                 "              Where nf.situacao        in (1,4,7) " & vbCrLf & _
                 "                and nf.TipoDeDocumento = 1 " & vbCrLf & _
                 "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id," & vbCrLf & _
                 "                       nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf & _
                 "             ) SbDev" & vbCrLf & _
                 "    On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
                 "   and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
                 "   and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
                 "   and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
                 "   and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
                 "   and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
                 "   and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
                 "   and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
                 "   and NFxI.Sequencia_id    = SbDev.Sequencia_Id" & vbCrLf & _
                 "   and NFxI.CFOP_id         = SbDev.CFOP_Id" & vbCrLf & _
                 "  LEFT JOIN (" & vbCrLf & _
                 "             SELECT MExNF.Empresa_Id," & vbCrLf & _
                 "                    MExNF.EndEmpresa_Id," & vbCrLf & _
                 "                    MExNF.Cliente_Id," & vbCrLf & _
                 "                    MExNF.EndCliente_Id," & vbCrLf & _
                 "                    MExNF.EntradaSaida_Id," & vbCrLf & _
                 "                    MExNF.Serie_Id," & vbCrLf & _
                 "                    MExNF.Nota_Id," & vbCrLf & _
                 "                    Mem.Produto as Produto_Id," & vbCrLf & _
                 "                    sum(MExNF.Quantidade) as Quantidade" & vbCrLf & _
                 " 		         FROM MemorandoDeExportacao Mem" & vbCrLf & _
                 "              Inner Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
                 "			 	   ON Mem.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
                 "				  AND Mem.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
                 "			  	  AND Mem.Memorando_Id           = MExNF.Memorando_Id" & vbCrLf & _
                 "              Inner Join NotasFiscais NF" & vbCrLf & _
                 "                 On NF.Empresa_id      = MExNF.Empresa_Id" & vbCrLf & _
                 "                and NF.EndEmpresa_Id   = MExNF.EndEmpresa_Id" & vbCrLf & _
                 "                and NF.Cliente_Id      = MExNF.Cliente_Id" & vbCrLf & _
                 "                and NF.EndCliente_Id   = MExNF.EndCliente_Id" & vbCrLf & _
                 "                and NF.EntradaSaida_Id = MExNF.EntradaSaida_Id" & vbCrLf & _
                 "                and NF.Serie_Id        = MExNF.Serie_Id" & vbCrLf & _
                 "                and NF.Nota_Id         = MExNF.Nota_Id" & vbCrLf & _
                 "              Where NF.situacao        in (1,4,7) " & vbCrLf & _
                 "                and NF.TipoDeDocumento = 1 " & vbCrLf & _
                 "             Group by MExNF.Empresa_Id, MExNF.EndEmpresa_Id, MExNF.Cliente_Id, MExNF.EndCliente_Id, MExNF.EntradaSaida_Id," & vbCrLf & _
                 "                      MExNF.Serie_Id, MExNF.Nota_Id, Mem.Produto" & vbCrLf & _
                 "          ) SbMem" & vbCrLf & _
                 "    On NFxI.Empresa_id      = SbMem.Empresa_Id" & vbCrLf & _
                 "   and NFxI.EndEmpresa_Id   = SbMem.EndEmpresa_Id" & vbCrLf & _
                 "   and NFxI.Cliente_Id      = SbMem.Cliente_Id" & vbCrLf & _
                 "   and NFxI.EndCliente_Id   = SbMem.EndCliente_Id" & vbCrLf & _
                 "   and NFxI.EntradaSaida_Id = SbMem.EntradaSaida_Id" & vbCrLf & _
                 "   and NFxI.Serie_Id        = SbMem.Serie_Id" & vbCrLf & _
                 "   and NFxI.Nota_Id         = SbMem.Nota_Id" & vbCrLf & _
                 "   and NFxI.Produto_Id      = SbMem.Produto_Id" & vbCrLf & _
                 IIf(Alteracao, "  Inner Join", "  left Join") & vbCrLf & _
                 "            (SELECT Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                 "                    Cliente_Id, EndCliente_Id, " & vbCrLf & _
                 "                    Nota_Id, Serie_Id, EntradaSaida_Id," & vbCrLf & _
                 "                    Quantidade" & vbCrLf & _
                 "               FROM MemorandoDeExportacaoXNotaFiscal " & vbCrLf & _
                 "              Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                 "                and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                 "                and Memorando_Id           ='" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                 "            ) sbMemNota" & vbCrLf & _
                 "    On NFxI.Empresa_id      = sbMemNota.Empresa_Id" & vbCrLf & _
                 "   and NFxI.EndEmpresa_Id   = sbMemNota.EndEmpresa_Id" & vbCrLf & _
                 "   and NFxI.Cliente_Id      = sbMemNota.Cliente_Id" & vbCrLf & _
                 "   and NFxI.EndCliente_Id   = sbMemNota.EndCliente_Id" & vbCrLf & _
                 "   and NFxI.EntradaSaida_Id = sbMemNota.EntradaSaida_Id" & vbCrLf & _
                 "   and NFxI.Serie_Id        = sbMemNota.Serie_Id" & vbCrLf & _
                 "   and NFxI.Nota_Id         = sbMemNota.Nota_Id" & vbCrLf


        strSQL &= " Where SO.Memorando       = 1" & vbCrLf & _
                  "   and SO.Devolucao       = 'N'" & vbCrLf & _
                  "   and NF.Situacao        in (1,4,7) " & vbCrLf & _
                  "   and NF.TipoDeDocumento = 1 " & vbCrLf & _
                  "   and NFxI.QuantidadeFiscal + isnull(sbMemNota.Quantidade,0) - isnull(sbDev.Quantidade,0) - isnull(SbMem.Quantidade,0) > 0" & vbCrLf
        '"   and NFxI.Produto_id = '" & Memorando.CodigoProduto & "'" & vbCrLf & _

        If Memorando.NossaEmissao Then
            strSQL &= "   and NF.Empresa_Id      ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and NF.EndEmpresa_Id   = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and NF.Cliente_Id      ='" & Memorando.CodigoClienteMemorando & "'" & vbCrLf & _
                      "   and NF.EndCliente_Id   = " & Memorando.EnderecoClienteMemorando & vbCrLf & _
                      "   and NF.EntradaSaida_Id ='E'" & vbCrLf
        Else
            strSQL &= "   and NF.Empresa_Id      ='" & Memorando.CodigoClienteMemorando & "'" & vbCrLf & _
                      "   and NF.EndEmpresa_Id   = " & Memorando.EnderecoClienteMemorando & vbCrLf & _
                      "   and NF.Cliente_Id      ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and NF.EndCliente_Id   = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and NF.EntradaSaida_Id ='S'" & vbCrLf
        End If

        If Numnotas.Trim.Length > 0 Then
            strSQL &= "   and NF.nota_id in (" & Numnotas & ")"
        End If

        If Safra.Length > 0 Then
            strSQL &= "   and P.Safra = '" & Safra & "'"
        End If

        strSQL &= " order by case" & vbCrLf & _
                  "            when isnull(sbMemNota.Quantidade,0) > 0" & vbCrLf & _
                  "              then 0" & vbCrLf & _
                  "              else 1" & vbCrLf & _
                  "         end," & vbCrLf & _
                  "         NF.nota_id" & vbCrLf


        Dim objBanco As New AcessaBanco
        Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "FixacaoXNotaFiscal")

        Dim adicionar As Boolean = IIf(Me.Count > 0, True, False)
        Dim existe As Boolean

        For Each row As DataRow In ds.Tables(0).Rows
            Dim NFComp As New MemorandoDeExportacaoXNotaFiscal(_Memorando)
            NFComp.CodigoEmpresa = row("Empresa_Id")
            NFComp.EnderecoEmpresa = row("EndEmpresa_Id")
            NFComp.CodigoCliente = row("Cliente_Id")
            NFComp.EnderecoCliente = row("EndCliente_Id")
            NFComp.NumeroNota = row("Nota_Id")
            NFComp.Serie = row("Serie_Id")
            NFComp.EntradaSaida = row("EntradaSaida_id")
            NFComp.QuantidadeNota = row("QuantidadeNota")
            NFComp.QuantidaDevolvida = row("QtdeDevolvida")
            NFComp.QuantidadeJaComprovada = row("QtdeUsadaMemorando")

            NFComp.QuantidadeMemorandoOriginal = row("QtdeUsadaNesteMemorando")
            NFComp.QuantidadeMemorando = row("QtdeUsadaNesteMemorando")
            If adicionar Then
                existe = False
                For i As Integer = 0 To Me.Count - 1
                    If row("Nota_id") = Me(i).NumeroNota And row("Serie_id") = Me(i).Serie Then
                        existe = True
                    End If
                Next
                If Not existe Then Me.Add(NFComp)
            Else
                Me.Add(NFComp)
            End If
        Next
    End Sub
#End Region

End Class


<Serializable()> _
Public Class MemorandoDeExportacaoXNotaFiscal

#Region "Contrutor"

    Public Sub New(ByVal Mem As MemorandoDeExportacao)
        _Memorando = Mem
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Memorando As MemorandoDeExportacao
    Private _CodigoEmpresa As String
    Private _EnderecoEmpresa As Integer
    Private _CodigoCliente As String
    Private _EnderecoCliente As Integer
    Private _NumeroNota As Integer
    Private _Serie As String
    Private _EntradaSaida As String

    Private _NotaFiscal As NotaFiscal

    Private _QuantidadeMemorandoOriginal As Decimal
    Private _QuantidadeMemorando As Decimal

    Private _QuantidadeNota As Decimal
    Private _QuantidaDevolvida As Decimal
    Private _QuantidadeJaComprovada As Decimal


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

    Public Property Memorando() As MemorandoDeExportacao
        Get
            Return _Memorando
        End Get
        Set(ByVal value As MemorandoDeExportacao)
            _Memorando = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _NotaFiscal = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
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

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
            _NotaFiscal = Nothing
        End Set
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



    Public ReadOnly Property NotaFiscal() As NotaFiscal
        Get
            If _NotaFiscal Is Nothing Then
                Dim NF_Consulta As New NotaFiscal
                NF_Consulta.CodigoEmpresa = _CodigoEmpresa
                NF_Consulta.EnderecoEmpresa = _EnderecoEmpresa
                NF_Consulta.CodigoCliente = _CodigoCliente
                NF_Consulta.EnderecoCliente = _EnderecoCliente
                NF_Consulta.Codigo = _NumeroNota
                NF_Consulta.Serie = _Serie
                NF_Consulta.EntradaSaida = IIf(_EntradaSaida = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                _NotaFiscal = New NotaFiscal(NF_Consulta)
            End If
            Return _NotaFiscal
        End Get
    End Property


    Public Property QuantidadeMemorandoOriginal() As Decimal
        Get
            Return _QuantidadeMemorandoOriginal
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeMemorandoOriginal = value
        End Set
    End Property

    Public Property QuantidadeMemorando() As Decimal
        Get
            Return _QuantidadeMemorando
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeMemorando = value
        End Set
    End Property

    Public Property QuantidadeNota() As Decimal
        Get
            Return _QuantidadeNota
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeNota = value
        End Set
    End Property

    Public Property QuantidaDevolvida() As Decimal
        Get
            Return _QuantidaDevolvida
        End Get
        Set(ByVal value As Decimal)
            _QuantidaDevolvida = value
        End Set
    End Property

    Public Property QuantidadeJaComprovada() As Decimal
        Get
            Return _QuantidadeJaComprovada
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeJaComprovada = value
        End Set
    End Property

    Public ReadOnly Property Saldo() As Decimal
        Get
            Return _QuantidadeNota - _QuantidaDevolvida - _QuantidadeJaComprovada + QuantidadeMemorandoOriginal
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = ""
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into MemorandoDeExportacaoXNotaFiscal(EmpresaMemorando_Id, EndEmpresaMemorando_Id, Memorando_Id, Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                      "                                              Cliente_Id, EndCliente_Id, Nota_Id, Serie_Id, EntradaSaida_Id, Quantidade)" & vbCrLf & _
                      " Values('" & Memorando.CodigoEmpresaMemorando & "'," & Memorando.EnderecoEmpresaMemorando & ",'" & Memorando.CodigoMemorando & "','" & _CodigoEmpresa & "'," & _EnderecoEmpresa & "," & vbCrLf & _
                      "        '" & _CodigoCliente & "'," & _EnderecoCliente & "," & _NumeroNota & ",'" & _Serie & "','" & _EntradaSaida & "'," & Str(_QuantidadeMemorando) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update MemorandoDeExportacaoXNotaFiscal set" & vbCrLf & _
                      "   Quantidade        = " & Str(_QuantidadeMemorando) & " " & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and Empresa_Id             ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id          = " & _EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id             ='" & _CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id          = " & _EnderecoCliente & vbCrLf & _
                      "   and Nota_Id                = " & _NumeroNota & vbCrLf & _
                      "   and Serie_Id               ='" & _Serie & "'" & vbCrLf & _
                      "   and EntradaSaida_Id        ='" & _EntradaSaida & "'"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete MemorandoDeExportacaoXNotaFiscal" & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and Empresa_Id             ='" & _CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id          = " & _EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id             ='" & _CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id          = " & _EnderecoCliente & vbCrLf & _
                      "   and Nota_Id                = " & _NumeroNota & vbCrLf & _
                      "   and Serie_Id               ='" & _Serie & "'" & vbCrLf & _
                      "   and EntradaSaida_Id        ='" & _EntradaSaida & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class