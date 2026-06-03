Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListEmissaoReceituarioNota
    Inherits List(Of EmissaoReceituarioNota)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As String, Optional ByVal DataInicial As String = "", Optional ByVal DataFinal As String = "", Optional ByVal Reimprimir As Boolean = False, Optional ByVal Nota As String = "")
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String = ""
        Sql = "SELECT nf.DataDaNota, nfi.Empresa_Id, nfi.EndEmpresa_Id, nfi.Cliente_Id, nfi.EndCliente_Id, C.Nome, C.Complemento," & vbCrLf & _
              "       nfi.EntradaSaida_Id, nfi.Serie_Id, nfi.Nota_Id," & vbCrLf & _
              "       count(distinct nfi.Produto_Id) as itens," & vbCrLf & _
              "       ceiling(convert(float,Count(distinct nfi.Produto_Id)) / 3) as NumReceitas, r.ART, r.RespTecnico" & vbCrLf & _
              "  FROM NotasFiscaisXItens AS nfi " & vbCrLf & _
              " Inner Join NotasFiscais as NF" & vbCrLf & _
              "    on NF.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
              "   and NF.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
              "   and NF.Cliente_Id      = nfi.Cliente_Id" & vbCrLf & _
              "   and NF.EndCliente_Id   = nfi.EndCliente_Id" & vbCrLf & _
              "   and NF.EntradaSaida_Id = nfi.EntradaSaida_Id" & vbCrLf & _
              "   and NF.Nota_Id         = nfi.Nota_Id" & vbCrLf & _
              "   and NF.Serie_Id        = nfi.Serie_Id" & vbCrLf & _
              "  LEFT JOIN Receita AS r " & vbCrLf & _
              "     ON nfi.Receita = r.Receita_Id " & vbCrLf & _
              " INNER JOIN Produtos AS P " & vbCrLf & _
              "    ON nfi.Produto_Id      = P.Produto_Id " & vbCrLf & _
              "   AND P.Fitossanitario    = 'S'" & vbCrLf & _
              "   AND nfi.EntradaSaida_Id = 'S'" & vbCrLf & _
              " Inner Join Clientes C" & vbCrLf & _
              "    ON C.Cliente_Id  = nfi.Cliente_Id" & vbCrLf & _
              "   AND C.Endereco_Id = nfi.EndCliente_Id" & vbCrLf & _
              " Where NF.Situacao = 1 " & vbCrLf

        If Reimprimir Then
            Sql &= " AND nfi.Receita > 0 " & vbCrLf
        Else
            Sql &= " AND (nfi.Receita is null Or nfi.Receita = 0) " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Nota) Then
            Sql &= " AND nfi.Nota_Id = " & Nota & vbCrLf
        End If

        If Empresa.Length > 0 Then
            Sql &= "   And NF.Empresa_Id = '" & Empresa & "' And NF.EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf
        End If

        If DataInicial.Length > 0 Then
            Sql &= " And NF.DataDaNota >= '" & CDate(DataInicial).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        If DataFinal.Length > 0 Then
            Sql &= " And NF.DataDaNota <= '" & CDate(DataFinal).ToString("yyyy/MM/dd") & "'" & vbCrLf
        End If

        Sql &= " group by nf.DataDaNota, nfi.Empresa_Id, nfi.EndEmpresa_Id, nfi.Cliente_Id, nfi.EndCliente_Id, C.Nome, C.Complemento, nfi.EntradaSaida_Id, r.ART, r.RespTecnico, " & vbCrLf & _
              "          nfi.Serie_Id,   nfi.Nota_Id" & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ERN As New EmissaoReceituarioNota
            ERN.DataDaNota = row("DataDaNota")
            ERN.CodigoEmpresa = row("Empresa_Id")
            ERN.EndEmpresa = row("EndEmpresa_Id")
            ERN.CodigoCliente = row("Cliente_Id")
            ERN.EndCliente = row("EndCliente_Id")
            ERN.Nome = row("Nome")
            ERN.Complemento = row("Complemento")
            ERN.EntradaSaida = row("EntradaSaida_Id")
            ERN.Serie = row("Serie_Id")
            ERN.Nota = row("Nota_Id")
            ERN.NumItens = row("itens")
            ERN.NumReceitas = row("NumReceitas")
            If Reimprimir Then
                ERN.CodigoART = row("ART")
                ERN.CodigoRespTecnico = row("RespTecnico")
                ERN.Itens = New ListEmissaoReceituarioProduto(ERN, Reimprimir)
            End If
            Me.Add(ERN)
        Next

    End Sub
#End Region

End Class

<Serializable()> _
Public Class EmissaoReceituarioNota

#Region "Fields"
    Private _DataDaNota As Date

    Private _Empresa As Cliente
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer

    Private _Cliente As Cliente
    Private _CodigoCliente As String
    Private _EndCliente As Integer
    Private _Nome As String
    Private _Complemento As String

    Private _EntradaSaida As String
    Private _Serie As String
    Private _Nota As Integer

    Private _NumItens As Integer
    Private _NumReceitas As Integer

    Private _Itens As ListEmissaoReceituarioProduto

    Private _RespTecnico As Cliente
    Private _CodigoRespTecnico As String
    Private _EndRespTecnico As Integer

    Private _Mensagem As String
    Private _CodigoART As Integer
#End Region

#Region "Property"
    Public Property DataDaNota() As Date
        Get
            Return _DataDaNota
        End Get
        Set(ByVal value As Date)
            _DataDaNota = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
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

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Public Property Complemento() As String
        Get
            Return _Complemento
        End Get
        Set(ByVal value As String)
            _Complemento = value
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property Nota() As Integer
        Get
            Return _Nota
        End Get
        Set(ByVal value As Integer)
            _Nota = value
        End Set
    End Property

    Public Property NumItens() As Integer
        Get
            Return _NumItens
        End Get
        Set(ByVal value As Integer)
            _NumItens = value
        End Set
    End Property

    Public Property NumReceitas() As Integer
        Get
            Return _NumReceitas
        End Get
        Set(ByVal value As Integer)
            _NumReceitas = value
        End Set
    End Property

    Public Property Itens() As ListEmissaoReceituarioProduto
        Get
            If _Itens Is Nothing Then _Itens = New ListEmissaoReceituarioProduto(Me)
            Return _Itens
        End Get
        Set(ByVal value As ListEmissaoReceituarioProduto)
            _Itens = value
        End Set
    End Property

    Public Property RespTecnico() As Cliente
        Get
            If _RespTecnico Is Nothing And _CodigoEmpresa.Length > 0 Then _RespTecnico = New Cliente(_CodigoRespTecnico, _EndRespTecnico)
            Return _RespTecnico
        End Get
        Set(ByVal value As Cliente)
            _RespTecnico = value
        End Set
    End Property

    Public Property CodigoRespTecnico() As String
        Get
            Return _CodigoRespTecnico
        End Get
        Set(ByVal value As String)
            _CodigoRespTecnico = value
            _RespTecnico = Nothing
        End Set
    End Property

    Public Property EndRespTecnico() As Integer
        Get
            Return _EndRespTecnico
        End Get
        Set(ByVal value As Integer)
            _EndRespTecnico = value
            _RespTecnico = Nothing
        End Set
    End Property

    Public Property Mensagem() As String
        Get
            Return _Mensagem
        End Get
        Set(ByVal value As String)
            _Mensagem = value
        End Set
    End Property

    Public Property CodigoART() As Integer
        Get
            Return _CodigoART
        End Get
        Set(ByVal value As Integer)

            _CodigoART = value
        End Set
    End Property
#End Region

#Region "Methods"
    Private Function CriarReceitas(ByRef Sqls As ArrayList) As Boolean
        Dim Receitas As New ListReceita
        Dim NumItensNaReceita As Integer
        Dim NumeradorReceita As New Numerador(CodigoEmpresa, EndEmpresa, 13)

        For Each ProdutoNota As EmissaoReceituarioProduto In Me.Itens

            If Not ProdutoNota.Feito And ProdutoNota.Encerrada = True Then
                NumItensNaReceita = 0
                Dim Receita As New Receita
                Receita.IUD = "I"
                Receita.CodigoReceita = NumeradorReceita.Sequencia + 1 + Receitas.Count
                ProdutoNota.CodigoReceita = Receita.CodigoReceita
                Receita.CodigoArt = ProdutoNota.Nota.CodigoART
                Receita.NumReceita = Receita.ART.ARTAtual + Receitas.Count + 1
                ProdutoNota.NumeroReceita = Receita.NumReceita
                If Receita.ART.ARTAtual + Receitas.Count + 1 > Receita.ART.ARTFinal Then
                    _Mensagem = "Não tem receitas suficiente para a emissão da(s) receita(s) para essa nota."
                    Return False
                End If
                Receita.CodigoCultura = ProdutoNota.CodigoCultura
                Receita.CodigoRespTecnico = ProdutoNota.Nota.CodigoRespTecnico
                Receita.EndRespTecnico = ProdutoNota.Nota.EndRespTecnico
                Receita.DataReceita = ProdutoNota.Nota.DataDaNota

                For Each Produto As EmissaoReceituarioProduto In Me.Itens
                    If Not Produto.Feito And Produto.Encerrada = True And ProdutoNota.CodigoCultura = Produto.CodigoCultura And NumItensNaReceita < 3 Then
                        Produto.Feito = True
                        Produto.CodigoReceita = Receita.CodigoReceita
                        Dim ReceitaProduto As New ReceitaXProduto
                        ReceitaProduto.IUD = "I"
                        ReceitaProduto.CodigoReceita = Produto.CodigoReceita
                        ReceitaProduto.CodigoProduto = Produto.CodigoProduto
                        ReceitaProduto.CodigoCulturaPragaFito = Produto.CodigoCulturaPragaFito
                        ReceitaProduto.CodigoFormaDeAplicacao = Produto.CodigoFormaAplicacao
                        ReceitaProduto.Quantidade = Produto.Quantidade
                        ReceitaProduto.CodigoDosagem = Produto.CodigoDosagem
                        ReceitaProduto.DosagemRecomendada = Math.Round(Produto.DosagemRecomendada, 4)
                        ReceitaProduto.UnidadeDeMedida = Produto.Dosagem.UnidadeDeMedida
                        ReceitaProduto.Vazao = "Aerea: " + Produto.Dosagem.VazaoAerea + " / Terreste: " + Produto.Dosagem.VazaoTerrestre
                        ReceitaProduto.AreaTotal = Produto.AreaTotal
                        ReceitaProduto.AreaTratada = Produto.AreaTratada
                        ReceitaProduto.ModoAplicacao = Produto.ModoDeAplicacao.Descricao.Replace("'", "")
                        ReceitaProduto.EpocaAplicacao = Produto.EpocaDeAplicacao.Replace("'", "")
                        ReceitaProduto.IntervaloDeSeguranca = Produto.Dosagem.IntervaloDeSeguranca.Replace("'", "")
                        ReceitaProduto.InstrucaoDeUso = Produto.Fito.InstrucoesUso.Replace("'", "")
                        ReceitaProduto.NumeroDeAplicacao = Produto.NumeroDeAplicacao
                        ReceitaProduto.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                        ReceitaProduto.UsuarioInclusaoData = Date.Now.ToString("yyyy-MM-dd")
                        Receita.Produtos.Add(ReceitaProduto)
                        NumItensNaReceita += 1
                    End If
                Next
                Receitas.Add(Receita)
            End If
        Next

        Sqls.Add(NumeradorReceita.IncrementarNumeradorSql(False, Receitas.Count))

        For Each Rec As Receita In Receitas
            Rec.SalvarSql(Sqls)
        Next

        For Each ProdutoNota As EmissaoReceituarioProduto In Me.Itens
            If ProdutoNota.Feito = True And ProdutoNota.Encerrada = True Then
                Dim sql As String
                sql = "UPDATE NotasFiscaisXItens SET Receita = " & ProdutoNota.CodigoReceita & vbCrLf & _
                      " WHERE Empresa_Id = '" & ProdutoNota.Nota.CodigoEmpresa & "'" & vbCrLf & _
                      "  AND EndEmpresa_Id = " & ProdutoNota.Nota.EndEmpresa & vbCrLf & _
                      "  AND Cliente_Id = '" & ProdutoNota.Nota.CodigoCliente & "'" & vbCrLf & _
                      "  AND EndCliente_Id = " & ProdutoNota.Nota.EndCliente & vbCrLf & _
                      "  AND EntradaSaida_Id = '" & ProdutoNota.Nota.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                      "  AND Serie_Id = '" & ProdutoNota.Nota.Serie & "'" & vbCrLf & _
                      "  AND Nota_Id = " & ProdutoNota.Nota.Nota & vbCrLf & _
                      "  AND Produto_Id = '" & ProdutoNota.CodigoProduto & "'"

                Sqls.Add(sql)
            End If
        Next

        Return True
    End Function

    Public Function SalvarReceitas() As Boolean
        Dim Sqls As New ArrayList
        Dim Banco As New AcessaBanco

        If Not CriarReceitas(Sqls) Then
            Return False
        End If

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            _Mensagem = HttpContext.Current.Session("ssMessage").ToString
            Return False
        End If
    End Function

#End Region

End Class

<Serializable()> _
Public Class ListEmissaoReceituarioProduto
    Inherits List(Of EmissaoReceituarioProduto)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Nota As EmissaoReceituarioNota, Optional ByVal Reimprimir As Boolean = False)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim Sql As String = ""
        Sql = "SELECT nfi.Produto_Id, P.Nome, P.Unidade, sum(nfi.QuantidadeFisica) as Qtde, nfi.Receita, r.Cultura, r.ARTReceita, r.RespTecnico, cpf.CulturaPragaFito_Id, cpf.Cultura, cpf.Fito, cpf.Praga, rp.Dosagem, rp.FormaDeAplicacao " & vbCrLf & _
              "  FROM NotasFiscaisXItens nfi" & vbCrLf & _
              "  LEFT JOIN Receita AS r " & vbCrLf & _
              "     ON nfi.Receita = r.Receita_Id " & vbCrLf & _
              "  LEFT JOIN ReceitaXProduto rp " & vbCrLf & _
              "     ON r.Receita_Id = rp.Receita_Id " & vbCrLf & _
              "  LEFT JOIN CulturaXPragaXFito cpf " & vbCrLf & _
              "     ON rp.CulturaPragaFito = cpf.CulturaPragaFito_Id " & vbCrLf & _
              "  inner join Produtos P" & vbCrLf & _
              "    on nfi.Produto_id      = P.produto_Id" & vbCrLf & _
              "   AND P.Fitossanitario    = 'S'" & vbCrLf & _
              " where nfi.Empresa_Id      ='" & Nota.CodigoEmpresa & "'" & vbCrLf & _
              "   and nfi.EndEmpresa_Id   = " & Nota.EndEmpresa & vbCrLf & _
              "   and nfi.Cliente_Id      ='" & Nota.CodigoCliente & "'" & vbCrLf & _
              "   and nfi.EndCliente_Id   = " & Nota.EndCliente & vbCrLf & _
              "   and nfi.EntradaSaida_Id ='" & Nota.EntradaSaida & "'" & vbCrLf & _
              "   and nfi.Nota_Id         = " & Nota.Nota & vbCrLf & _
              "   and nfi.Serie_Id        ='" & Nota.Serie & "'" & vbCrLf
        If (Not Reimprimir) Then
            Sql &= "   and (nfi.Receita is null Or nfi.Receita = 0) " & vbCrLf
        End If

        Sql &= " group by nfi.Produto_Id, P.Nome, P.Unidade, nfi.Receita, nfi.Receita, r.Cultura, r.ARTReceita, r.RespTecnico, cpf.CulturaPragaFito_Id, cpf.Cultura, cpf.Fito, cpf.Praga, rp.Dosagem, rp.FormaDeAplicacao " & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim ERP As New EmissaoReceituarioProduto(Nota)
            ERP.CodigoProduto = row("Produto_Id")
            ERP.Nome = row("Nome")
            ERP.Unidade = row("Unidade")
            ERP.Quantidade = row("Qtde")
            If Reimprimir Then
                ERP.Encerrada = True
                ERP.CodigoReceita = row("Receita")
                ERP.NumeroReceita = row("Receita")
                ERP.CodigoCultura = row("Cultura")
                ERP.CodigoPraga = row("Praga")
                ERP.CodigoDosagem = row("Dosagem")
                ERP.CodigoFormaAplicacao = row("FormaDeAplicacao")
            Else
                ERP.Encerrada = False
            End If
            Me.Add(ERP)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class EmissaoReceituarioProduto

    Public Sub New(ByVal pNota As EmissaoReceituarioNota)
        _Nota = pNota
    End Sub

#Region "Fields"
    Private _Nota As EmissaoReceituarioNota
    Private _Produto As Produto
    Private _Feito As Boolean = False
    Private _Encerrada As Boolean = False
    Private _CodigoProduto As String
    Private _Nome As String
    Private _Unidade As String
    Private _Quantidade As Decimal

    Private _Fito As Fito
    Private _CodigoPraga As Integer
    Private _Praga As Praga
    Private _CodigoCultura As Integer
    Private _Cultura As Cultura
    Private _AreaTotal As Decimal
    Private _AreaTratada As Decimal

    Private _CodigoCulturaPragaFito As Integer
    Private _EpocaDeAplicacao As String

    Private _CodigoDosagem As Integer
    Private _Dosagem As Dosagem

    Private _CodigoFormaAplicacao As Integer
    Private _FormaDeAplicacao As FormaDeAplicacao
    Private _dsFormasDeAplicacao As DataSet

    Private _dsCultura As DataSet
    Private _dsPraga As DataSet

    Private _Dosagens As ListDosagem
    Private _DosagemRecomendada As Decimal
    Private _ModoDeAplicacao As ModoDeAplicacao
    Private _NumeroDeAplicacao As Integer

    Private _CodigoReceita As Integer
    Private _NumeroReceita As Integer
#End Region

#Region "Property"
    Public ReadOnly Property Nota() As EmissaoReceituarioNota
        Get
            Return _Nota
        End Get
    End Property

    Public Property Feito() As Boolean
        Get
            Return _Feito
        End Get
        Set(ByVal value As Boolean)
            _Feito = value
        End Set
    End Property

    Public Property Encerrada() As Boolean
        Get
            Return _Encerrada
        End Get
        Set(ByVal value As Boolean)
            _Encerrada = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Fito = New Fito(value, 0)
            _dsCultura = Nothing
            _dsPraga = Nothing
            _Dosagens = Nothing
            _CodigoCulturaPragaFito = 0
            _CodigoFormaAplicacao = 0
            _ModoDeAplicacao = Nothing
            _dsFormasDeAplicacao = Nothing
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Public Property Unidade() As String
        Get
            Return _Unidade
        End Get
        Set(ByVal value As String)
            _Unidade = value
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

    Public ReadOnly Property Fito() As Fito
        Get
            Return _Fito
        End Get
    End Property

    Public ReadOnly Property Culturas() As DataSet
        Get
            If _dsCultura Is Nothing Then
                Dim sql As String
                Dim Banco As New AcessaBanco

                sql = " select distinct Cultura, C.Descricao" & _
                  "   from CulturaXPragaXFito CxP" & _
                  "  inner Join Cultura C" & _
                  "     on C.Cultura_Id = CxP.Cultura" & _
                  "   Where Fito = " & Fito.CodigoFito
                _dsCultura = Banco.ConsultaDataSet(sql, "Cultura")
                _dsPraga = Nothing
            End If
            Return _dsCultura
        End Get
    End Property

    Public Property AreaTotal() As Decimal
        Get
            If _CodigoCultura > 0 Then
                Dim sql As String
                Dim ds As New DataSet
                Dim Banco As New AcessaBanco

                sql = "SELECT AreaPlantada FROM ClienteXCultura " & vbCrLf & _
                      "WHERE Cliente_Id = '" & _Nota.CodigoCliente & "' AND EndCliente_Id = " & _Nota.EndCliente & " AND " & vbCrLf & _
                      "Ano_Id = " & Now.Year & " AND Cultura_Id = " & _CodigoCultura & ""

                ds = Banco.ConsultaDataSet(sql, "ClienteXCultura")

                If ds.Tables(0).Rows.Count > 0 Then
                    _AreaTotal = ds.Tables(0).Rows(0).Item("AreaPlantada")
                Else
                    _AreaTotal = 0
                End If
                Return _AreaTotal
            End If
            Return _AreaTotal
        End Get
        Set(ByVal value As Decimal)
            _AreaTotal = value
        End Set
    End Property

    Public Property AreaTratada() As Decimal
        Get
            Return _AreaTratada
        End Get
        Set(ByVal value As Decimal)
            _AreaTratada = value
        End Set
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
            _dsPraga = Nothing
            _Dosagens = Nothing
            _Cultura = Nothing
            _CodigoCulturaPragaFito = 0
            _CodigoFormaAplicacao = 0
            _ModoDeAplicacao = Nothing
            _dsFormasDeAplicacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And _CodigoCultura > 0 Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura
        End Get
    End Property

    Public ReadOnly Property Pragas() As DataSet
        Get
            If _dsPraga Is Nothing Then
                Dim sql As String
                Dim Banco As New AcessaBanco

                sql = "select distinct CxP.Praga, P.NomeComum" & _
                      "  from CulturaXPragaXFito CxP" & _
                      " inner Join Praga P" & _
                      "    on P.Praga_Id = CxP.Praga" & _
                      " Where Fito    = " & Fito.CodigoFito & _
                      "   and Cultura = " & _CodigoCultura

                _dsPraga = Banco.ConsultaDataSet(sql, "Pragas")
                _Dosagens = Nothing
            End If
            Return _dsPraga
        End Get
    End Property

    Public Property CodigoPraga() As Integer
        Get
            Return _CodigoPraga
        End Get
        Set(ByVal value As Integer)
            _CodigoPraga = value
            _Dosagens = Nothing
            _Praga = Nothing
            _CodigoCulturaPragaFito = 0
            _CodigoFormaAplicacao = 0
            _ModoDeAplicacao = Nothing
            _dsFormasDeAplicacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Praga() As Praga
        Get
            If _Praga Is Nothing And _CodigoPraga > 0 Then _Praga = New Praga(_CodigoPraga)
            Return _Praga
        End Get
    End Property

    Public ReadOnly Property CodigoCulturaPragaFito() As Integer
        Get
            If _CodigoCulturaPragaFito = 0 And Fito.CodigoFito > 0 And _CodigoCultura > 0 And _CodigoPraga > 0 Then
                Dim banco As New AcessaBanco
                Dim ds As DataSet
                Dim Sql As String = "Select CulturaPragaFito_Id as Codigo, EpocaAplicacao from CulturaxPragaXFito where Fito =" & Fito.CodigoFito & " and Cultura =" & _CodigoCultura & " and Praga = " & _CodigoPraga

                ds = banco.ConsultaDataSet(Sql, "CxPxF")
                If ds.Tables(0).Rows.Count > 0 Then
                    _CodigoCulturaPragaFito = ds.Tables(0).Rows(0)("Codigo")
                    _EpocaDeAplicacao = ds.Tables(0).Rows(0)("EpocaAplicacao")
                End If
            End If

            Return _CodigoCulturaPragaFito
        End Get
    End Property

    Public ReadOnly Property EpocaDeAplicacao() As String
        Get
            Return _EpocaDeAplicacao
        End Get
    End Property

    Public ReadOnly Property Dosagens() As ListDosagem
        Get
            If _Dosagens Is Nothing And (Fito.CodigoFito > 0 And _CodigoCultura > 0 And _CodigoPraga > 0) Then _Dosagens = New ListDosagem(0, Fito.CodigoFito, _CodigoCultura, _CodigoPraga)
            Return _Dosagens
        End Get
    End Property

    Public Property CodigoDosagem() As Integer
        Get
            Return _CodigoDosagem
        End Get
        Set(ByVal value As Integer)
            _CodigoDosagem = value
            _Dosagem = Nothing
        End Set
    End Property

    Public ReadOnly Property Dosagem() As Dosagem
        Get
            If _Dosagem Is Nothing And _CodigoDosagem > 0 Then _Dosagem = New Dosagem(CodigoCulturaPragaFito, _CodigoDosagem)
            Return _Dosagem
        End Get
    End Property

    Public Property DosagemRecomendada() As Decimal
        Get
            Return _DosagemRecomendada
        End Get
        Set(ByVal value As Decimal)
            _DosagemRecomendada = value
        End Set
    End Property

    Public Property CodigoFormaAplicacao() As Integer
        Get
            Return _CodigoFormaAplicacao
        End Get
        Set(ByVal value As Integer)
            _CodigoFormaAplicacao = value
            _ModoDeAplicacao = Nothing
        End Set
    End Property

    Public ReadOnly Property FormaDeAplicacao() As FormaDeAplicacao
        Get
            If _FormaDeAplicacao Is Nothing And _CodigoFormaAplicacao > 0 Then _FormaDeAplicacao = New FormaDeAplicacao(_CodigoFormaAplicacao)
            Return _FormaDeAplicacao
        End Get
    End Property

    Public ReadOnly Property FormasDeAplicacao() As DataSet
        Get
            If _Fito.CodigoFito > 0 And _CodigoCultura > 0 And _CodigoPraga > 0 Then
                Dim sql As String
                Dim Banco As New AcessaBanco

                sql = "SELECT FormaDeAplicacao.FormaDeAplicacao_Id, FormaDeAplicacao.Descricao " & _
                      "  FROM         ModoDeAplicacao INNER JOIN " & _
                      " FormaDeAplicacao ON ModoDeAplicacao.FormaDeAplicacao_Id = FormaDeAplicacao.FormaDeAplicacao_Id " & _
                      " WHERE ModoDeAplicacao.CulturaPragaFito_Id = " & CodigoCulturaPragaFito

                _dsFormasDeAplicacao = Banco.ConsultaDataSet(sql, "FormaDeAplicacao")
            End If
            Return _dsFormasDeAplicacao
        End Get
    End Property

    Public ReadOnly Property ModoDeAplicacao() As ModoDeAplicacao
        Get
            If _ModoDeAplicacao Is Nothing And _CodigoFormaAplicacao > 0 And _CodigoCultura > 0 And _CodigoPraga > 0 And Fito.CodigoFito > 0 Then _ModoDeAplicacao = New ModoDeAplicacao(0, _CodigoFormaAplicacao, Fito.CodigoFito, _CodigoCultura, _CodigoPraga)
            Return _ModoDeAplicacao
        End Get
    End Property

    Public Property NumeroDeAplicacao() As Integer
        Get
            Return _NumeroDeAplicacao
        End Get
        Set(ByVal value As Integer)
            _NumeroDeAplicacao = value
        End Set
    End Property

    Public Property CodigoReceita() As Integer
        Get
            Return _CodigoReceita
        End Get
        Set(ByVal value As Integer)
            _CodigoReceita = value
        End Set
    End Property

    Public Property NumeroReceita() As Integer
        Get
            Return _NumeroReceita
        End Get
        Set(ByVal value As Integer)
            _NumeroReceita = value
        End Set
    End Property
#End Region

End Class