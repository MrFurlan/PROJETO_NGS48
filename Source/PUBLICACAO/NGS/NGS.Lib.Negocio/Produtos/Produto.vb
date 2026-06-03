Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListProduto
    Inherits List(Of Produto)

    Public Sub New(Optional ByVal GrupoProduto As String = "", Optional ByVal Agrupar As String = "", Optional ByVal ControlarLote As String = "", Optional ByVal ControlarEmbalagem As String = "", Optional ByVal Fitossanitario As String = "", Optional ByVal Where As String = "", Optional ByVal pSelecionado As Boolean = False, Optional ByVal pProduto As Produto = Nothing)
        Dim objBanco As New AcessaBanco()
        Try
            Dim strSQL As String = "SELECT Produto_Id, Grupo, Unidade, Etapa, Situacao, Embalagem, NCM, Nome, Descricao, " & vbCrLf &
                                   "       DescricaoMapa, PesoQuantidade, EstoqueMinimo, UPPER(isnull(Agrupar,'N')) AS Agrupar, QuantidadeNaCaixa, " & vbCrLf &
                                   "       Gtin8, Gtin12, Gtin13, Gtin14," & vbCrLf &
                                   "       Qualidade, IPI, CarteiraDeCompras, " & vbCrLf &
                                   "       CarteiraDeVendas, ICMS, TipoDoItem, isnull(CodigoDoGenero,0) AS CodigoDoGenero, CodigoEX, CodigoDoServico, " & vbCrLf &
                                   "       UPPER(isnull(controlarEstoque,'N')) as ControlarEstoque, UPPER(isnull(ControlarLote,'')) as ControlarLote, " & vbCrLf &
                                   "       UPPER(isnull(ControlarEmbalagem,'N')) as ControlarEmbalagem, UPPER(isnull(Fitossanitario,'N')) as Fitossanitario, " & vbCrLf &
                                   "       isnull(ProdutoIndea,'') as ProdutoIndea, isnull(Marca,0) as Marca, isnull(ControlarPecas,0) as ControlarPecas, " & vbCrLf &
                                   "       isnull(ControlarPrecoDePauta,0) as ControlarPrecoDePauta, isnull(SubCodigoDoGenero,0) AS SubCodigoDoGenero, " & vbCrLf &
                                   "       isnull(ControlarRomaneio,0) as ControlarRomaneio, isnull(ControlarPesagem,0) as ControlarPesagem, isnull(ControlarDecimais,0) as ControlarDecimais, " & vbCrLf &
                                   "       ISNULL(Cnae,'') Cnae, isnull(Almoxarifado,0) as Almoxarifado, isnull(PrecoDoProduto,0) as PrecoDoProduto, " & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(ControlarNumeroDoLote,0) as ControlarNumeroDoLote, isnull(CodigoEstadoFisico,0) as CodigoEstadoFisico, " & vbCrLf &
                                   "       isnull(InfaDProd,'') as InfaDProd, isnull(AutorizacaoDeRetirada,0) AS AutorizacaoDeRetirada, isnull(RegMinAgr,'') AS RegMinAgr, " & vbCrLf &
                                   "       isnull(CodigoProdutoTerceiro,'') as CodigoProdutoTerceiro, ISNULL(CustoIndireto,0) AS CustoIndireto, " & vbCrLf &
                                   "       isnull(Dashboard,0) as Dashboard, isnull(Seguimento,0) as Seguimento" & vbCrLf &
                                   "  FROM Produtos WHERE 1 = 1 " & vbCrLf

            If GrupoProduto.Length > 0 Then
                strSQL &= " AND left(Grupo," & GrupoProduto.Length & ") = '" & GrupoProduto & "' " & vbCrLf
            End If

            If Agrupar.Length > 0 AndAlso (Agrupar = "S" Or Agrupar = "N") Then
                strSQL &= " AND Agrupar = '" & Agrupar & "' " & vbCrLf
            End If

            If ControlarLote.Length > 0 Then
                strSQL &= " AND ControlarLote = '" & ControlarLote & "' " & vbCrLf
            End If

            If ControlarEmbalagem.Length > 0 Then
                strSQL &= " AND ControlarEmbalagem = '" & ControlarEmbalagem & "' " & vbCrLf
            End If

            If Fitossanitario.Length > 0 Then
                strSQL &= " AND Fitossanitario = '" & Fitossanitario & "' " & vbCrLf
            End If

            If Where.Length > 0 Then
                strSQL &= "AND " & Where & vbCrLf
            End If

            If Not pProduto Is Nothing Then
                If Not pProduto.Codigo Is Nothing AndAlso pProduto.Codigo.Length > 0 Then strSQL &= " AND Produto_Id LIKE '" & pProduto.Codigo & "%'" & vbCrLf
                If Not pProduto.Unidade Is Nothing AndAlso pProduto.Unidade.Length > 0 Then strSQL &= " AND Unidade = '" & pProduto.Unidade & "'" & vbCrLf
                If pProduto.CodigoSituacao > 0 Then strSQL &= " AND Situacao = " & pProduto.CodigoSituacao & vbCrLf
                If Not pProduto.ProdutoIndea Is Nothing AndAlso pProduto.ProdutoIndea.Length > 0 Then strSQL &= " AND ProdutoIndea = '" & pProduto.ProdutoIndea & "'" & vbCrLf
                If Not pProduto.Nome Is Nothing AndAlso pProduto.Nome.Length > 0 Then strSQL &= " AND Nome LIKE '" & pProduto.Nome & "%'" & vbCrLf
                If Not pProduto.Descricao Is Nothing AndAlso pProduto.Descricao.Length > 0 Then strSQL &= " AND Descricao LIKE '" & pProduto.Descricao & "%'" & vbCrLf
                If pProduto.Etapa > 0 Then strSQL &= " AND Etapa = " & pProduto.Etapa & vbCrLf
                If pProduto.CodigoEmbalagem > 0 Then strSQL &= " AND Embalagem = " & pProduto.CodigoEmbalagem & vbCrLf
                If Not pProduto.DescricaoMapa Is Nothing AndAlso pProduto.DescricaoMapa.Length > 0 Then strSQL &= " AND DescricaoMapa LIKE '" & pProduto.DescricaoMapa & "%'" & vbCrLf
                If Not pProduto.NCM Is Nothing AndAlso pProduto.NCM.Length > 0 Then strSQL &= " AND NCM Like '" & pProduto.NCM & "%'" & vbCrLf
                If Not pProduto.PesoQuantidade Is Nothing AndAlso pProduto.PesoQuantidade.Length > 0 Then strSQL &= " AND PesoQuantidade = '" & pProduto.PesoQuantidade & "'" & vbCrLf
                If pProduto.ControlarEstoque Then strSQL &= " AND ControlarEstoque = 'S'" & vbCrLf
                If pProduto.ControlarPrecoDePauta Then strSQL &= " AND ControlarPrecoDePauta = " & CByte(pProduto.ControlarPrecoDePauta) & vbCrLf
                If pProduto.ControlarPecas Then strSQL &= " AND ControlarPecas = " & CByte(pProduto.ControlarPecas) & vbCrLf
                If pProduto.CodigoDaMarca > 0 Then strSQL &= " AND Marca = " & pProduto.CodigoDaMarca & vbCrLf
                If pProduto.CodigoGenero > 0 Then strSQL &= " AND CodigoDoGenero = " & pProduto.CodigoGenero & vbCrLf
                If pProduto.CustoIndireto Then strSQL &= " CustoIndireto = " & CByte(pProduto.CustoIndireto) & vbCrLf
            End If

            strSQL &= "ORDER BY Descricao "

            Dim dsProdutos As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

            For Each drProduto As DataRow In dsProdutos.Tables(0).Rows
                Dim objProduto As New Produto()

                objProduto.Selecionado = pSelecionado

                objProduto.Codigo = drProduto("Produto_Id").ToString()
                objProduto.CodigoGrupo = drProduto("Grupo").ToString()
                objProduto.Unidade = drProduto("Unidade").ToString()
                objProduto.Etapa = Convert.ToInt32(drProduto("Etapa"))
                objProduto.CodigoSituacao = Convert.ToInt32(drProduto("Situacao"))
                objProduto.CodigoEmbalagem = Convert.ToInt32(drProduto("Embalagem"))
                objProduto.NCM = drProduto("NCM").ToString()
                objProduto.CodigoCnae = drProduto("Cnae")
                objProduto.Nome = drProduto("Nome").ToString()
                objProduto.Gtin8 = drProduto("Gtin8").ToString()
                objProduto.Gtin12 = drProduto("Gtin12").ToString()
                objProduto.Gtin13 = drProduto("Gtin13").ToString()
                objProduto.Gtin14 = drProduto("Gtin14").ToString()
                objProduto.Descricao = drProduto("Descricao").ToString()
                objProduto.InfaDProd = drProduto("InfaDProd").ToString
                objProduto.DescricaoMapa = drProduto("DescricaoMapa").ToString()
                objProduto.PesoQuantidade = drProduto("PesoQuantidade").ToString()
                objProduto.EstoqueMinimo = Convert.ToDouble(drProduto("EstoqueMinimo"))
                objProduto.Agrupar = drProduto("Agrupar").ToString()
                objProduto.QuantidadeCaixa = Convert.ToInt32(drProduto("QuantidadeNaCaixa"))
                objProduto.Qualidade = Convert.ToInt32(drProduto("Qualidade"))
                If drProduto.IsNull("IPI") Then objProduto.IPI = 0 Else objProduto.IPI = Convert.ToDouble(drProduto("IPI"))
                'objProduto.IPITributado = drProduto("IPITributado").ToString()
                'objProduto.PisCofinsIntegral = drProduto("PisCofinsIntegral").ToString()
                'objProduto.PisCofinsPresumido = drProduto("PisCofinsPresumido").ToString()
                objProduto.CodigoCarteiraCompra = drProduto("CarteiraDeCompras").ToString()
                objProduto.CodigoCarteiraVenda = drProduto("CarteiraDeVendas").ToString()
                If drProduto.IsNull("TipoDoItem") Then objProduto.TipoItem = 0 Else objProduto.TipoItem = Convert.ToInt32(drProduto("TipoDoItem"))
                objProduto.CodigoGenero = Convert.ToInt32(drProduto("CodigoDoGenero"))
                objProduto.CodigoEX = drProduto("CodigoEX").ToString()
                If drProduto.IsNull("CodigoDoServico") Then objProduto.CodigoServico = 0 Else objProduto.CodigoServico = Convert.ToInt32(drProduto("CodigoDoServico"))
                If drProduto.IsNull("ICMS") Then objProduto.ICMS = 0 Else objProduto.ICMS = Convert.ToDouble(drProduto("ICMS"))
                objProduto.ControlarEstoque = drProduto("ControlarEstoque").Equals("S")
                objProduto.ControlarLote = drProduto("ControlarLote").Equals("S")
                objProduto.ControlarEmbalagem = drProduto("ControlarEmbalagem").Equals("S")
                objProduto.Fitossanitario = drProduto("Fitossanitario").Equals("S")
                objProduto.ProdutoIndea = drProduto("ProdutoIndea")
                objProduto.CodigoDaMarca = drProduto("Marca")
                objProduto.ControlarPecas = drProduto("ControlarPecas") = "True"
                objProduto.ControlarPrecoDePauta = drProduto("ControlarPrecoDePauta") = "True"
                objProduto.SubCodigoGenero = Convert.ToInt32(drProduto("SubCodigoDoGenero"))
                objProduto.ControlarRomaneio = drProduto("ControlarRomaneio") = "True"
                objProduto.ControlarPesagem = drProduto("ControlarPesagem") = "True"
                objProduto.ControlarDecimais = drProduto("ControlarDecimais") = "True"
                objProduto.Almoxarifado = drProduto("Almoxarifado") = "True"
                objProduto.PrecoDoProduto = drProduto("PrecoDoProduto") = "True"
                objProduto.UsuarioInclusao = drProduto("UsuarioInclusao")
                objProduto.UsuarioInclusaoData = drProduto("UsuarioInclusaoData")
                objProduto.UsuarioAlteracao = drProduto("UsuarioAlteracao")
                objProduto.UsuarioAlteracaoData = drProduto("UsuarioAlteracaoData")
                objProduto.ControlarNumeroDoLote = drProduto("ControlarNumeroDoLote") = "True"
                objProduto.CodigoEstadoFisico = drProduto("CodigoEstadoFisico")
                objProduto.AutorizacaoDeRetirada = drProduto("AutorizacaoDeRetirada")
                objProduto.RegistroMinisterioAgricultura = drProduto("RegMinAgr")
                objProduto.CodigoProdutoTerceiro = drProduto("CodigoProdutoTerceiro")
                objProduto.CustoIndireto = drProduto("CustoIndireto") = "True"
                objProduto.Dashboard = drProduto("Dashboard") = "True"
                objProduto.CodigoSeguimento = drProduto("Seguimento")

                Me.Add(objProduto)
            Next
        Finally
            objBanco = Nothing
        End Try
    End Sub

End Class

<Serializable()>
Public Class Produto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal CodigoProduto As String)
        Me.Codigo = CodigoProduto
        Selecionar(CodigoProduto)
    End Sub
#End Region

#Region "Fields"
    Private _Selecionado As Boolean = False
    Private _IUD As String
    Private _Codigo As String = ""
    Private _CodigoGrupo As String = ""
    Private _Grupo As GrupoProduto
    Private _Unidade As String = ""
    Private _Etapa As Integer

    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    Private _CodigoEmbalagem As Integer
    Private _Embalagem As Embalagem
    Private _NCM As String = ""
    Private _Nome As String = ""

    Private _Gtin8 As String
    Private _Gtin12 As String
    Private _Gtin13 As String
    Private _Gtin14 As String

    Private _Descricao As String = ""
    Private _InfaDProd As String = ""
    Private _DescricaoMapa As String = ""
    Private _PesoQuantidade As String = ""
    Private _EstoqueMinimo As Double
    Private _Agrupar As String = ""
    Private _QuantidadeCaixa As Integer
    Private _Qualidade As Integer
    Private _IPI As Double
    Private _CodigoCarteiraCompra As String = ""
    Private _CarteiraCompra As CarteiraFinanceira
    Private _CodigoCarteiraVenda As String = ""
    Private _CarteiraVenda As CarteiraFinanceira

    Private _TipoItem As Integer
    Private _CodigoGenero As Integer
    Private _CodigoEX As String
    Private _CodigoServico As Integer
    Private _ICMS As Double
    Private _ControlarEstoque As Boolean
    Private _ControlarLote As Boolean 'Para Sementes
    Private _Fitossanitario As Boolean
    Private _ControlarEmbalagem As Boolean
    Private _ProdutoIndea As String = ""
    Private _DescricaoTecnica As String = ""
    Private _ControlarPecas As Boolean
    Private _ControlarPrecoDePauta As Boolean
    Private _ControlarRomaneio As Boolean
    Private _ControlarPesagem As Boolean
    Private _ControlarDecimais As Boolean
    Private _SubCodigoGenero As Integer
    Private _DetalhesGenero As GeneroDoProdutoXSub
    Private _CodigoDaMarca As Integer
    Private _Marca As Marca
    Private _CodigoCnae As String
    Private _Almoxarifado As Boolean
    Private _PrecoDoProduto As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
    Private _Dashboard As Boolean
    Private _CodigoSeguimento As Integer
    Private _Seguimento As Seguimento

    Private _AutorizacaoDeRetirada As Boolean
    Private _CustoIndireto As Boolean

    Private _ControlarNumeroDoLote As Boolean 'Para produto Quimico

    'Listas
    Private _ProdutoXEmbalagens As ListProdutoXEmbalagem
    Private _ProdutoXPrecos As ListProdutoXPreco
    Private _ProdutosAgrupados As Negocio.ListProdutoAgrupado
    Private _ProdutoAgrupador As Negocio.ListProdutoAgrupado
    Private _UnidadesDeComercializacao As ListProdutosXUnidadeDeComercializacao
    Private _ProdutoXEspecificacao As ListProdutoXEspecificacao
    Private _ProdutoXEPI As ListProdutoXEPI
    Private _ProdutoXProcedimento As ListProdutoXProcedimento

    Private _CodigoEstadoFisico As Integer
    Private _EstadoFisico As EstadoFisicoIA

    Private _RegistroMinisterioAgricultura As String

    Private _CodigoProdutoTerceiro As String = ""
    Private _NumeroDoLote As Integer
    Private _ValidadeDoLote As Date


#End Region

#Region "Property"
    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
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

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
            _Codigo = value
        End Set
    End Property

    Public Property CodigoGrupo() As String
        Get
            Return _CodigoGrupo
        End Get
        Set(ByVal value As String)
            _CodigoGrupo = value
        End Set
    End Property

    Public Property Grupo() As GrupoProduto
        Get
            If _Grupo Is Nothing And Me.CodigoGrupo.Length > 0 Then _Grupo = New GrupoProduto(Me.CodigoGrupo)
            Return _Grupo
        End Get
        Set(ByVal value As GrupoProduto)
            _Grupo = value
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

    Public Property Etapa() As Integer
        Get
            Return _Etapa
        End Get
        Set(ByVal value As Integer)
            _Etapa = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
            _Situacao = Nothing
        End Set
    End Property

    Public ReadOnly Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
    End Property

    Public ReadOnly Property DescSituacao() As String
        Get
            If Situacao Is Nothing Then
                Return ""
            Else
                Return Situacao.Descricao
            End If
        End Get
    End Property

    Public Property CodigoEmbalagem() As Integer
        Get
            Return _CodigoEmbalagem
        End Get
        Set(ByVal value As Integer)
            _CodigoEmbalagem = value
        End Set
    End Property

    Public Property Embalagem() As Embalagem
        Get
            If _Embalagem Is Nothing And _CodigoEmbalagem > 0 Then _Embalagem = New Embalagem(_CodigoEmbalagem)
            Return _Embalagem
        End Get
        Set(ByVal value As Embalagem)
            _Embalagem = value
        End Set
    End Property

    Public Property NCM() As String
        Get
            Return _NCM
        End Get
        Set(ByVal value As String)
            _NCM = value
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

    'Public ReadOnly Property imgQRcode As String
    '    Get
    '        Return Convert.ToBase64String(_QRcode)
    '    End Get
    'End Property

    'Public ReadOnly Property imgCodBarras As String
    '    Get
    '        Return Convert.ToBase64String(_CodBarras)
    '    End Get
    'End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property InfaDProd() As String
        Get
            Return _InfaDProd
        End Get
        Set(ByVal value As String)
            _InfaDProd = value
        End Set
    End Property

    Public Property DescricaoMapa() As String
        Get
            Return _DescricaoMapa
        End Get
        Set(ByVal value As String)
            _DescricaoMapa = value
        End Set
    End Property

    Public Property PesoQuantidade() As String
        Get
            Return _PesoQuantidade
        End Get
        Set(ByVal value As String)
            _PesoQuantidade = value
        End Set
    End Property

    Public Property EstoqueMinimo() As Double
        Get
            Return _EstoqueMinimo
        End Get
        Set(ByVal value As Double)
            _EstoqueMinimo = value
        End Set
    End Property

    Public Property Agrupar() As String
        Get
            Return _Agrupar
        End Get
        Set(ByVal value As String)
            _Agrupar = value
        End Set
    End Property

    Public Property QuantidadeCaixa() As Integer
        Get
            Return _QuantidadeCaixa
        End Get
        Set(ByVal value As Integer)
            _QuantidadeCaixa = value
        End Set
    End Property

    Public Property Qualidade() As Integer
        Get
            Return _Qualidade
        End Get
        Set(ByVal value As Integer)
            _Qualidade = value
        End Set
    End Property

    Public Property IPI() As Double
        Get
            Return _IPI
        End Get
        Set(ByVal value As Double)
            _IPI = value
        End Set
    End Property

    'Public Property IPITributado() As String
    '    Get
    '        Return _IPITributado
    '    End Get
    '    Set(ByVal value As String)
    '        _IPITributado = value
    '    End Set
    'End Property

    'Public Property PisCofinsIntegral() As String
    '    Get
    '        Return _PisCofinsIntegral
    '    End Get
    '    Set(ByVal value As String)
    '        _PisCofinsIntegral = value
    '    End Set
    'End Property

    'Public Property PisCofinsPresumido() As String
    '    Get
    '        Return _PisCofinsPresumido
    '    End Get
    '    Set(ByVal value As String)
    '        _PisCofinsPresumido = value
    '    End Set
    'End Property

    Public Property CodigoCarteiraCompra() As String
        Get
            Return _CodigoCarteiraCompra
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraCompra = value
            _CarteiraCompra = Nothing
        End Set
    End Property

    Public ReadOnly Property CarteiraCompra() As CarteiraFinanceira
        Get
            If _CarteiraCompra Is Nothing And _CodigoCarteiraCompra.Length > 0 Then _CarteiraCompra = New CarteiraFinanceira(_CodigoCarteiraCompra)
            Return _CarteiraCompra
        End Get
    End Property

    Public Property CodigoCarteiraVenda() As String
        Get
            Return _CodigoCarteiraVenda
        End Get
        Set(ByVal value As String)
            _CodigoCarteiraVenda = value
            _CarteiraVenda = Nothing
        End Set
    End Property

    Public ReadOnly Property CarteiraVenda() As CarteiraFinanceira
        Get
            If _CarteiraVenda Is Nothing And _CodigoCarteiraVenda.Length > 0 Then _CarteiraVenda = New CarteiraFinanceira(_CodigoCarteiraVenda)
            Return _CarteiraVenda
        End Get
    End Property

    Public Property TipoItem() As Integer
        Get
            Return _TipoItem
        End Get
        Set(ByVal value As Integer)
            _TipoItem = value
        End Set
    End Property

    Public Property CodigoGenero() As Integer
        Get
            Return _CodigoGenero
        End Get
        Set(ByVal value As Integer)
            _CodigoGenero = value
        End Set
    End Property

    Public Property CodigoEX() As String
        Get
            Return _CodigoEX
        End Get
        Set(ByVal value As String)
            _CodigoEX = value
        End Set
    End Property

    Public Property CodigoServico() As Integer
        Get
            Return _CodigoServico
        End Get
        Set(ByVal value As Integer)
            _CodigoServico = value
        End Set
    End Property

    Public Property ICMS() As Double
        Get
            Return _ICMS
        End Get
        Set(ByVal value As Double)
            _ICMS = value
        End Set
    End Property

    Public Property ControlarEstoque() As Boolean
        Get
            Return _ControlarEstoque
        End Get
        Set(ByVal value As Boolean)
            _ControlarEstoque = value
        End Set
    End Property

    Public Property ControlarLote() As Boolean
        Get
            Return _ControlarLote
        End Get
        Set(ByVal value As Boolean)
            _ControlarLote = value
        End Set
    End Property

    Public Property Fitossanitario() As Boolean
        Get
            Return _Fitossanitario
        End Get
        Set(ByVal value As Boolean)
            _Fitossanitario = value
        End Set
    End Property

    Public Property ControlarEmbalagem() As Boolean
        Get
            Return _ControlarEmbalagem
        End Get
        Set(ByVal value As Boolean)
            _ControlarEmbalagem = value
        End Set
    End Property

    Public Property ProdutoIndea() As String
        Get
            Return _ProdutoIndea
        End Get
        Set(ByVal value As String)
            _ProdutoIndea = value
        End Set
    End Property

    Public Property ControlarPecas() As Boolean
        Get
            Return _ControlarPecas
        End Get
        Set(ByVal value As Boolean)
            _ControlarPecas = value
        End Set
    End Property

    Public Property ControlarPrecoDePauta() As Boolean
        Get
            Return _ControlarPrecoDePauta
        End Get
        Set(ByVal value As Boolean)
            _ControlarPrecoDePauta = value
        End Set
    End Property

    Public Property ControlarRomaneio() As Boolean
        Get
            Return _ControlarRomaneio
        End Get
        Set(ByVal value As Boolean)
            _ControlarRomaneio = value
        End Set
    End Property

    Public Property ControlarPesagem() As Boolean
        Get
            Return _ControlarPesagem
        End Get
        Set(ByVal value As Boolean)
            _ControlarPesagem = value
        End Set
    End Property

    Public Property ControlarDecimais() As Boolean
        Get
            Return _ControlarDecimais
        End Get
        Set(ByVal value As Boolean)
            _ControlarDecimais = value
        End Set
    End Property

    Public Property SubCodigoGenero() As Integer
        Get
            Return _SubCodigoGenero
        End Get
        Set(ByVal value As Integer)
            _SubCodigoGenero = value
        End Set
    End Property

    Public ReadOnly Property DetalhesGenero() As GeneroDoProdutoXSub
        Get
            If _SubCodigoGenero > 0 Then _DetalhesGenero = New GeneroDoProdutoXSub(_CodigoGenero, _SubCodigoGenero)
            Return _DetalhesGenero
        End Get
    End Property

    Public ReadOnly Property DescricaoTecnica()
        Get
            If _DescricaoTecnica.Length = 0 And _Fitossanitario Then
                Dim sql As String
                sql = "Select 'Nome Tec. ' + NomeTecnico + ', Reg. MA ' + RegistroMA + ', Reg. ONU ' + RegistroONU + ', CL Risco ' + CR.Descricao + ', CL Toxicologica ' + CT.Descricao as DescricaoTecnica" & vbCrLf &
                      "  From ProdutoXFito PF" & vbCrLf &
                      " Inner Join Fito F" & vbCrLf &
                      "    on PF.Fito_Id = F.Fito_Id" & vbCrLf &
                      " Inner Join ClasseDeRisco CR" & vbCrLf &
                      "    on F.ClasseRisco = CR.ClasseDeRisco_Id" & vbCrLf &
                      " Inner Join ClasseToxicologica CT" & vbCrLf &
                      "    on CT.ClasseTox_Id = F.ClasseTox" & vbCrLf &
                      " Where PF.Produto_Id = '" & _Codigo & "'" & vbCrLf

                Dim Banco As New AcessaBanco
                Dim ds As DataSet
                ds = Banco.ConsultaDataSet(sql, "DescricaoTecnica")
                If ds.Tables(0).Rows.Count > 0 Then
                    _DescricaoTecnica = ds.Tables(0).Rows(0)("DescricaoTecnica")
                End If
            End If
            Return _DescricaoTecnica
        End Get
    End Property

    Public Property CodigoDaMarca() As Integer
        Get
            Return _CodigoDaMarca
        End Get
        Set(ByVal value As Integer)
            _CodigoDaMarca = value
        End Set
    End Property

    Public Property Marca() As Marca
        Get
            If _Marca Is Nothing And _CodigoDaMarca > 0 Then _Marca = New Marca(_CodigoDaMarca)
            Return _Marca
        End Get
        Set(ByVal value As Marca)
            _Marca = value
        End Set
    End Property
    Public Property CodigoCnae() As String
        Get
            Return _CodigoCnae
        End Get
        Set(ByVal value As String)
            _CodigoCnae = value
        End Set
    End Property

    Public Property Almoxarifado() As Boolean
        Get
            Return _Almoxarifado
        End Get
        Set(ByVal value As Boolean)
            _Almoxarifado = value
        End Set
    End Property

    Public Property PrecoDoProduto() As Boolean
        Get
            Return _PrecoDoProduto
        End Get
        Set(value As Boolean)
            _PrecoDoProduto = value
        End Set
    End Property

    Public Property AutorizacaoDeRetirada() As Boolean
        Get
            Return _AutorizacaoDeRetirada
        End Get
        Set(value As Boolean)
            _AutorizacaoDeRetirada = value
        End Set
    End Property

    Public Property CustoIndireto() As Boolean
        Get
            Return _CustoIndireto
        End Get
        Set(value As Boolean)
            _CustoIndireto = value
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

    Public Property ControlarNumeroDoLote() As Boolean
        Get
            Return _ControlarNumeroDoLote
        End Get
        Set(ByVal value As Boolean)
            _ControlarNumeroDoLote = value
        End Set
    End Property

    Public ReadOnly Property TemProdutoDeConsumo() As Boolean
        Get
            Return ExisteProdutoDeConsumo()
        End Get
    End Property

    Public Property Dashboard() As Boolean
        Get
            Return _Dashboard
        End Get
        Set(ByVal value As Boolean)
            _Dashboard = value
        End Set
    End Property

    Public Property CodigoSeguimento() As Integer
        Get
            Return _CodigoSeguimento
        End Get
        Set(ByVal value As Integer)
            _CodigoSeguimento = value
        End Set
    End Property

    Public Property Seguimento() As Seguimento
        Get
            If _Seguimento Is Nothing And _CodigoSeguimento > 0 Then _Seguimento = New Seguimento(_CodigoSeguimento)
            Return _Seguimento
        End Get
        Set(ByVal value As Seguimento)
            _Seguimento = value
        End Set
    End Property

    'Listas
    Public Property ProdutoXEmbalagens() As ListProdutoXEmbalagem
        Get
            If _ProdutoXEmbalagens Is Nothing Then _ProdutoXEmbalagens = New ListProdutoXEmbalagem(Me)
            Return _ProdutoXEmbalagens
        End Get
        Set(ByVal value As ListProdutoXEmbalagem)
            _ProdutoXEmbalagens = value
        End Set
    End Property

    Public Property ProdutoXPrecos() As ListProdutoXPreco
        Get
            If _ProdutoXPrecos Is Nothing And _Codigo.Length > 0 Then _ProdutoXPrecos = New ListProdutoXPreco(0, "", 0, _Codigo, "", "")
            Return _ProdutoXPrecos
        End Get
        Set(ByVal value As ListProdutoXPreco)
            _ProdutoXPrecos = value
        End Set
    End Property

    Public Property ProdutosAgrupados As Negocio.ListProdutoAgrupado
        Get
            If _ProdutosAgrupados Is Nothing Then _ProdutosAgrupados = New Negocio.ListProdutoAgrupado(Me, 0)
            Return _ProdutosAgrupados
        End Get
        Set(value As Negocio.ListProdutoAgrupado)
            _ProdutosAgrupados = value
        End Set
    End Property

    Public Property ProdutoAgrupador As Negocio.ListProdutoAgrupado
        Get
            If _ProdutoAgrupador Is Nothing Then _ProdutoAgrupador = New Negocio.ListProdutoAgrupado(Me, 1)
            Return _ProdutoAgrupador
        End Get
        Set(value As Negocio.ListProdutoAgrupado)
            _ProdutoAgrupador = value
        End Set
    End Property

    Public Property UnidadesDeComercializacao As ListProdutosXUnidadeDeComercializacao
        Get
            If _UnidadesDeComercializacao Is Nothing Then _UnidadesDeComercializacao = New Negocio.ListProdutosXUnidadeDeComercializacao(Me)
            Return _UnidadesDeComercializacao
        End Get
        Set(value As ListProdutosXUnidadeDeComercializacao)
            _UnidadesDeComercializacao = value
        End Set
    End Property

    Public Property ProdutoXEspecificacao As ListProdutoXEspecificacao
        Get
            If _ProdutoXEspecificacao Is Nothing Then _ProdutoXEspecificacao = New Negocio.ListProdutoXEspecificacao(Me)
            Return _ProdutoXEspecificacao
        End Get
        Set(value As ListProdutoXEspecificacao)
            _ProdutoXEspecificacao = value
        End Set
    End Property

    Public Property ProdutoXEPI As ListProdutoXEPI
        Get
            If _ProdutoXEPI Is Nothing Then _ProdutoXEPI = New Negocio.ListProdutoXEPI(Me)
            Return _ProdutoXEPI
        End Get
        Set(value As ListProdutoXEPI)
            _ProdutoXEPI = value
        End Set
    End Property

    Public Property ProdutoXProcedimento As ListProdutoXProcedimento
        Get
            If _ProdutoXProcedimento Is Nothing Then _ProdutoXProcedimento = New Negocio.ListProdutoXProcedimento(Me)
            Return _ProdutoXProcedimento
        End Get
        Set(value As ListProdutoXProcedimento)
            _ProdutoXProcedimento = value
        End Set
    End Property

    Public Property CodigoEstadoFisico() As Integer
        Get
            Return _CodigoEstadoFisico
        End Get
        Set(ByVal value As Integer)
            _CodigoEstadoFisico = value
            _EstadoFisico = Nothing
        End Set
    End Property

    Public ReadOnly Property EstadoFisico() As EstadoFisicoIA
        Get
            If _EstadoFisico Is Nothing And _CodigoEstadoFisico > 0 Then _EstadoFisico = New EstadoFisicoIA(_CodigoEstadoFisico)
            Return _EstadoFisico
        End Get
    End Property

    Public Property RegistroMinisterioAgricultura() As String
        Get
            Return _RegistroMinisterioAgricultura
        End Get
        Set(ByVal value As String)
            _RegistroMinisterioAgricultura = value
        End Set
    End Property

    Public Property CodigoProdutoTerceiro() As String
        Get
            Return _CodigoProdutoTerceiro
        End Get
        Set(ByVal value As String)
            _CodigoProdutoTerceiro = value
        End Set
    End Property

    Public Property Gtin8 As String
        Get
            Return _Gtin8
        End Get
        Set(value As String)
            _Gtin8 = value
        End Set
    End Property

    Public Property Gtin12 As String
        Get
            Return _Gtin12
        End Get
        Set(value As String)
            _Gtin12 = value
        End Set
    End Property

    Public Property Gtin13 As String
        Get
            Return _Gtin13
        End Get
        Set(value As String)
            _Gtin13 = value
        End Set
    End Property

    Public Property Gtin14 As String
        Get
            Return _Gtin14
        End Get
        Set(value As String)
            _Gtin14 = value
        End Set
    End Property

    Public Property NumeroDoLote As Integer
        Get
            Return _NumeroDoLote
        End Get
        Set(value As Integer)
            _NumeroDoLote = value
        End Set
    End Property

    Public Property ValidadeDoLote As Date
        Get
            Return _ValidadeDoLote
        End Get
        Set(value As Date)
            _ValidadeDoLote = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Selecionar(ByVal CodigoProduto As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Produto_Id, Grupo, Unidade, Etapa, Situacao, Embalagem, NCM, Nome, Descricao, " & vbCrLf &
                                   "       DescricaoMapa, PesoQuantidade, EstoqueMinimo, UPPER(isnull(Agrupar,'N')) AS Agrupar, QuantidadeNaCaixa, " & vbCrLf &
                                   "       Qualidade, IPI, CarteiraDeCompras, " & vbCrLf &
                                   "       ISNULL(Gtin8, '') AS Gtin8, ISNULL(Gtin12, '') AS Gtin12, ISNULL(Gtin13, '') AS Gtin13, ISNULL(Gtin14, '') AS Gtin14," & vbCrLf &
                                   "       CarteiraDeVendas, ICMS, TipoDoItem, isnull(CodigoDoGenero,0) AS CodigoDoGenero, CodigoEX, CodigoDoServico, " & vbCrLf &
                                   "       UPPER(isnull(controlarEstoque,'N')) as ControlarEstoque, UPPER(isnull(ControlarLote,'')) as ControlarLote, " & vbCrLf &
                                   "       UPPER(isnull(ControlarEmbalagem,'N')) as ControlarEmbalagem, UPPER(isnull(Fitossanitario,'N')) as Fitossanitario, " & vbCrLf &
                                   "       UPPER(isnull(ProdutoIndea,'')) as ProdutoIndea, isnull(Marca,0) as Marca, isnull(ControlarPecas,0) as ControlarPecas, " & vbCrLf &
                                   "       isnull(ControlarPrecoDePauta,0) as ControlarPrecoDePauta, isnull(SubCodigoDoGenero,0) AS SubCodigoDoGenero, " & vbCrLf &
                                   "       isnull(ControlarRomaneio,0) as ControlarRomaneio, isnull(ControlarPesagem,0) as ControlarPesagem, " & vbCrLf &
                                   "       isnull(ControlarDecimais,0) as ControlarDecimais, ISNULL(Cnae,'') Cnae, isnull(Almoxarifado,0) as Almoxarifado, isnull(PrecoDoProduto,0) as PrecoDoProduto, " & vbCrLf &
                                   "       isnull(UsuarioInclusao,'') as UsuarioInclusao, isnull(UsuarioInclusaoData,GetDate()) as UsuarioInclusaoData," & vbCrLf &
                                   "       isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                                   "       isnull(ControlarNumeroDoLote,0) as ControlarNumeroDoLote, isnull(CodigoEstadoFisico,0) as CodigoEstadoFisico, " & vbCrLf &
                                   "       isnull(InfaDProd,'') as InfaDProd, isnull(AutorizacaoDeRetirada,0) AS AutorizacaoDeRetirada, isnull(RegMinAgr,'') AS RegMinAgr, " & vbCrLf &
                                   "       isnull(CodigoProdutoTerceiro,'') as CodigoProdutoTerceiro, isnull(CustoIndireto,0) as CustoIndireto, " & vbCrLf &
                                   "       isnull(Dashboard,0) as Dashboard, isnull(Seguimento,0) as Seguimento" & vbCrLf &
                                   "  FROM Produtos " & vbCrLf &
                                   " WHERE Produto_Id = '" & CodigoProduto & "' "

            Dim dsProdutos As DataSet = objBanco.ConsultaDataSet(strSQL, "Produtos")

            If dsProdutos.Tables(0).Rows.Count > 0 Then
                Dim drProduto As DataRow = dsProdutos.Tables(0).Rows(0)

                Me.Codigo = drProduto("Produto_Id").ToString()
                Me.CodigoGrupo = drProduto("Grupo").ToString()
                Me.Unidade = drProduto("Unidade").ToString()
                Me.Etapa = Convert.ToInt32(drProduto("Etapa"))
                Me.CodigoSituacao = Convert.ToInt32(drProduto("Situacao"))
                Me.CodigoEmbalagem = Convert.ToInt32(drProduto("Embalagem"))
                Me.NCM = drProduto("NCM").ToString()
                Me.CodigoCnae = drProduto("Cnae")
                Me.Nome = drProduto("Nome").ToString()
                Me.Gtin8 = drProduto("Gtin8").ToString()
                Me.Gtin12 = drProduto("Gtin12").ToString()
                Me.Gtin13 = drProduto("Gtin13").ToString()
                Me.Gtin14 = drProduto("Gtin14").ToString()
                Me.Descricao = drProduto("Descricao").ToString()
                Me.InfaDProd = drProduto("InfaDProd").ToString()
                Me.DescricaoMapa = drProduto("DescricaoMapa").ToString()
                Me.PesoQuantidade = drProduto("PesoQuantidade").ToString()
                Me.EstoqueMinimo = Convert.ToDouble(drProduto("EstoqueMinimo"))
                Me.Agrupar = drProduto("Agrupar").ToString()
                Me.QuantidadeCaixa = Convert.ToInt32(drProduto("QuantidadeNaCaixa"))
                Me.Qualidade = Convert.ToInt32(drProduto("Qualidade"))
                If drProduto.IsNull("IPI") Then Me.IPI = 0 Else Me.IPI = Convert.ToDouble(drProduto("IPI"))
                Me.CodigoCarteiraCompra = drProduto("CarteiraDeCompras").ToString()
                Me.CodigoCarteiraVenda = drProduto("CarteiraDeVendas").ToString()
                If drProduto.IsNull("TipoDoItem") Then Me.TipoItem = 0 Else Me.TipoItem = Convert.ToInt32(drProduto("TipoDoItem"))
                Me.CodigoGenero = Convert.ToInt32(drProduto("CodigoDoGenero"))
                Me.CodigoEX = drProduto("CodigoEX").ToString()
                If drProduto.IsNull("CodigoDoServico") Then Me.CodigoServico = 0 Else Me.CodigoServico = Convert.ToInt32(drProduto("CodigoDoServico"))
                If drProduto.IsNull("ICMS") Then Me.ICMS = 0 Else Me.ICMS = Convert.ToDouble(drProduto("ICMS"))
                Me.ControlarEstoque = drProduto("ControlarEstoque").ToString.ToUpper.Equals("S")
                Me.ControlarLote = drProduto("ControlarLote").Equals("S")
                Me.ControlarEmbalagem = drProduto("ControlarEmbalagem").Equals("S")
                Me.Fitossanitario = drProduto("Fitossanitario").Equals("S")
                Me.ProdutoIndea = drProduto("ProdutoIndea")
                Me.CodigoDaMarca = drProduto("Marca")
                Me.ControlarPecas = drProduto("ControlarPecas") = "True"
                Me.ControlarPrecoDePauta = drProduto("ControlarPrecoDePauta") = "True"
                Me.ControlarRomaneio = drProduto("ControlarRomaneio") = "True"
                Me.ControlarPesagem = drProduto("ControlarPesagem") = "True"
                Me.ControlarDecimais = drProduto("ControlarDecimais") = "True"
                Me.SubCodigoGenero = Convert.ToInt32(drProduto("SubCodigoDoGenero"))
                Me.Almoxarifado = drProduto("Almoxarifado") = "True"
                Me.PrecoDoProduto = drProduto("PrecoDoProduto") = "True"
                Me.UsuarioInclusao = drProduto("UsuarioInclusao")
                Me.UsuarioInclusaoData = drProduto("UsuarioInclusaoData")
                Me.UsuarioAlteracao = drProduto("UsuarioAlteracao")
                Me.UsuarioAlteracaoData = drProduto("UsuarioAlteracaoData")
                Me.ControlarNumeroDoLote = drProduto("ControlarNumeroDoLote") = "True"
                Me.CodigoEstadoFisico = drProduto("CodigoEstadoFisico")
                Me.AutorizacaoDeRetirada = drProduto("AutorizacaoDeRetirada")
                Me.RegistroMinisterioAgricultura = drProduto("RegMinAgr")
                Me.CodigoProdutoTerceiro = drProduto("CodigoProdutoTerceiro")
                Me.CustoIndireto = drProduto("CustoIndireto") = "True"
                Me.Dashboard = drProduto("Dashboard") = "True"
                Me.CodigoSeguimento = drProduto("Seguimento")
            End If

            Return True
        Catch ex As Exception
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Sub BuscarSequencia()
        Dim strSQL As String

        strSQL = "Declare" & vbCrLf &
                 "@seq int" & vbCrLf &
                 "Select @seq = ISNULL(max(Produto_Id),'" & Me.CodigoGrupo & "0000') + 1 " & vbCrLf &
                 "   From Produtos " & vbCrLf &
                 "  where left(produto_Id," & Me.CodigoGrupo.Length & ") = '" & Me.CodigoGrupo & "'" & vbCrLf &
                 "if @seq > " & Me.CodigoGrupo & "9999" & vbCrLf &
                 "begin" & vbCrLf &
                 "  select Produto_Id," & vbCrLf &
                 "         ROW_NUMBER () over(order by Produto_Id) as row" & vbCrLf &
                 "    into #seq" & vbCrLf &
                 "    from Produtos" & vbCrLf &
                 "   Where left(produto_Id," & Me.CodigoGrupo.Length & ") = '" & Me.CodigoGrupo & "'" & vbCrLf &
                 "     and len(produto_id)    =" & Me.CodigoGrupo.Length & " + 4" & vbCrLf &
                 "   order by Produto_Id" & vbCrLf &
                 "   select @seq = " & Me.CodigoGrupo & "0000 + min(row)" & vbCrLf &
                 "     from #seq" & vbCrLf &
                 "    where convert(int,RIGHT(produto_id,4)) <> row" & vbCrLf &
                 "  end" & vbCrLf &
                 " Select @seq as Sequencia"

        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "Sequencia")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Me.Codigo = ds.Tables(0).Rows(0).Item("Sequencia")
        Else
            Me.Codigo = Me.CodigoGrupo & "0001"
        End If
    End Sub

    Public Function ExistePedido() As Boolean
        Dim strSQL As String = "Select Top 1 Pedido_Id " & vbCrLf &
                               "  From PedidoXItem " & vbCrLf &
                               " where Produto_id = '" & Me.Codigo & "'"
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "PXI")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ExisteNota() As Boolean
        Dim strSQL As String = "Select Top 1 Nota_Id " & vbCrLf &
                               "  From NotasFiscaisXItens " & vbCrLf &
                               " where Produto_id = '" & Me.Codigo & "'"
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "NotasFiscaisXItens")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ExisteProducao() As Boolean
        Dim strSQL As String = "Select Top 1 Produto_Id " & vbCrLf &
                               "  From Producao " & vbCrLf &
                               " where Produto_id = '" & Me.Codigo & "'"
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "Producao")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ExisteProdutoDeConsumo() As Boolean
        Dim strSQL As String = "Select Top 1 pc.Produto_Id " & vbCrLf &
                               "  From ProdutoDeConsumo pc " & vbCrLf &
                               "      inner join Produtos p on p.Produto_Id = pc.Produto_id" & vbCrLf &
                               " where pc.Produto_id = '" & Me.Codigo & "'"
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "Producao")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function ExisteOperacoesXEncargos() As Boolean
        Dim strSQL As String
        strSQL = "Select Top 1 Produto " & vbCrLf &
                 "  From OperacaoXEstado " & vbCrLf &
                 " where Produto = '" & Me.Codigo & "'" & vbCrLf &
                 " Union " & vbCrLf &
                 "Select Top 1 Produto_id " & vbCrLf &
                 "  From operacoesxencargos " & vbCrLf &
                 " where Produto_id = '" & Me.Codigo & "'" & vbCrLf
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet
        ds = objBanco.ConsultaDataSet(strSQL, "OperacoesXEncargos")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        salvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub salvarSql(ByRef sqls As ArrayList)
        Dim strSql As String = ""

        Select Case IUD
            Case "I"
                strSql = "INSERT INTO Produtos(Produto_Id, Grupo, Unidade, Etapa, Situacao, " & vbCrLf &
                         "Embalagem, NCM, Nome, Descricao, DescricaoMapa, " & vbCrLf &
                         "PesoQuantidade, EstoqueMinimo, Agrupar, Qualidade, " & vbCrLf &
                         "IPI, QuantidadeNaCaixa,  " & vbCrLf &
                         "CarteiraDeCompras, CarteiraDeVendas, TipoDoItem, CodigoDoGenero, " & vbCrLf &
                         "CodigoEX, CodigoDoServico, ControlarLote, ControlarEmbalagem, " & vbCrLf &
                         "Fitossanitario, ProdutoIndea, Marca, ControlarEstoque, ControlarPecas, " & vbCrLf &
                         "ControlarPrecoDePauta, SubCodigoDoGenero, ControlarRomaneio, ControlarPesagem, " & vbCrLf &
                         "ControlarDecimais, Cnae, Almoxarifado, PrecoDoProduto, UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf &
                         "Gtin8, Gtin12, Gtin13, Gtin14," & vbCrLf &
                         "ControlarNumeroDoLote, CodigoEstadoFisico, InfaDProd, AutorizacaoDeRetirada, RegMinAgr, CodigoProdutoTerceiro, CustoIndireto," & vbCrLf &
                         "Dashboard, Seguimento) " & vbCrLf &
                         "Values('" & Me.Codigo & "','" & Me.CodigoGrupo & "','" & Me.Unidade & "'," & Me.Etapa & "," & Me.CodigoSituacao & "," & vbCrLf &
                         Me.CodigoEmbalagem & ",'" & Me.NCM & "','" & Funcoes.EliminarAspasSimples(Me.Nome) & "','" & Funcoes.EliminarAspasSimples(Me.Descricao) & "','" & Funcoes.EliminarAspasSimples(Me.DescricaoMapa) & "', " & vbCrLf &
                         "'" & Me.PesoQuantidade & "'," & Str(Me.EstoqueMinimo) & ",'" & Me.Agrupar & "'," & Me.Qualidade & "," & vbCrLf &
                         Str(Me.IPI) & ", " & Me.QuantidadeCaixa & vbCrLf &
                         ", '" & Me.CodigoCarteiraCompra & "', '" & Me.CodigoCarteiraVenda & "', " & Me.TipoItem & "," & Me.CodigoGenero & "," & vbCrLf &
                         "'" & Me.CodigoEX & "', " & Me.CodigoServico & ",'" & IIf(Me.ControlarLote, "S", "N") & "','" & IIf(Me.ControlarEmbalagem, "S", "N") & "'," & vbCrLf &
                         "'" & IIf(_Fitossanitario, "S", "N") & "','" & _ProdutoIndea & "'," & IIf(Me.CodigoDaMarca > 0, Me.CodigoDaMarca, "NULL") & ",'" & IIf(Me.ControlarEstoque, "S", "N") & "'," & vbCrLf &
                         CByte(_ControlarPecas) & "," & CByte(Me.ControlarPrecoDePauta) & "," & Me.SubCodigoGenero & "," & CByte(Me.ControlarRomaneio) & "," & CByte(Me.ControlarPesagem) & "," & vbCrLf &
                         CByte(Me.ControlarDecimais) & "," & Me.CodigoCnae.ToSqlNULL & "," & CByte(Me.Almoxarifado) & "," & CByte(Me.PrecoDoProduto) & "," & vbCrLf &
                         "'" & UsuarioServidor.NomeUsuario & "','" & Now.ToString("yyyy-MM-dd") & "','" & Gtin8 & "','" & Gtin12 & "','" & Gtin13 & "','" & Gtin14 & "'," & vbCrLf &
                         CByte(Me.ControlarNumeroDoLote) & "," & Me.CodigoEstadoFisico & ",'" & _InfaDProd & "'," & vbCrLf &
                         CByte(Me.AutorizacaoDeRetirada) & ",'" & Me.RegistroMinisterioAgricultura & "','" & Me.CodigoProdutoTerceiro & "'," & CByte(CustoIndireto) & "," &
                         CByte(Me.Dashboard) & "," & Me.CodigoSeguimento & ")"
                sqls.Add(strSql)
                SalvarTabelasRelacionadasSql(sqls)
            Case "U"
                strSql = "update produtos set grupo = '" & Me.CodigoGrupo & "', Unidade = '" & Me.Unidade & "', etapa = " & Me.Etapa & vbCrLf &
                         ", situacao = " & Me.CodigoSituacao & ", embalagem = " & Me.CodigoEmbalagem & ", ncm = '" & Me.NCM & "', nome = '" & Funcoes.EliminarAspasSimples(Me.Nome) & "'" & vbCrLf &
                         ", descricao = '" & Funcoes.EliminarAspasSimples(Me.Descricao) & "', descricaomapa = '" & Funcoes.EliminarAspasSimples(Me.DescricaoMapa) & "'" & vbCrLf &
                         ", pesoquantidade = '" & Me.PesoQuantidade & "', estoqueminimo = " & Str(EstoqueMinimo) & vbCrLf &
                         ", agrupar = '" & Me.Agrupar & "', qualidade = " & Me.Qualidade & ", ipi = " & Str(Me.IPI) & vbCrLf &
                         ", quantidadenacaixa = " & Me.QuantidadeCaixa & vbCrLf &
                         ", carteiradecompras = '" & Me.CodigoCarteiraCompra & "', carteiradevendas = '" & Me.CodigoCarteiraVenda & "'" & vbCrLf &
                         ", tipodoitem = " & Me.TipoItem & ", codigodogenero = " & Me.CodigoGenero & ", codigoex = '" & Me.CodigoEX & "', codigodoservico = " & Me.CodigoServico & vbCrLf &
                         ", controlarlote = '" & IIf(Me.ControlarLote, "S", "N") & "',controlarembalagem = '" & IIf(ControlarEmbalagem, "S", "N") & "'" & vbCrLf &
                         ", fitossanitario = '" & IIf(Me.Fitossanitario, "S", "N") & "', produtoindea = '" & Me.ProdutoIndea & "', marca = " & IIf(Me.CodigoDaMarca > 0, Me.CodigoDaMarca, "NULL") & vbCrLf &
                         ", controlarestoque = '" & IIf(Me.ControlarEstoque, "S", "N") & "', controlarpecas = " & CByte(Me.ControlarPecas) & vbCrLf &
                         ", controlarprecodepauta =" & CByte(Me.ControlarPrecoDePauta) & ", subcodigodogenero = " & Me.SubCodigoGenero & ", controlarromaneio = " & CByte(Me.ControlarRomaneio) & vbCrLf &
                         ", controlarpesagem =" & CByte(Me.ControlarPesagem) & vbCrLf &
                         ", controlardecimais =" & CByte(Me.ControlarDecimais) & vbCrLf &
                         ", Cnae = " & Me.CodigoCnae.ToSqlNULL & vbCrLf &
                         ", Almoxarifado = " & CByte(Me.Almoxarifado) & vbCrLf &
                         ", PrecoDoProduto = " & CByte(Me.PrecoDoProduto) & vbCrLf &
                         ", UsuarioAlteracao ='" & UsuarioServidor.NomeUsuario & "'" & vbCrLf &
                         ", UsuarioAlteracaoData = '" & Now.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         ", ControlarNumeroDoLote = " & CByte(Me.ControlarNumeroDoLote) & vbCrLf &
                         ", CodigoEstadoFisico = " & Me.CodigoEstadoFisico & vbCrLf &
                         ", InfaDProd = '" & Me.InfaDProd & "'" & vbCrLf &
                         ", AutorizacaoDeRetirada = " & CByte(Me.AutorizacaoDeRetirada) & vbCrLf &
                         ", Gtin8 = '" & Me.Gtin8 & "'" & vbCrLf &
                         ", Gtin12 = '" & Me.Gtin12 & "'" & vbCrLf &
                         ", Gtin13 = '" & Me.Gtin13 & "'" & vbCrLf &
                         ", Gtin14 = '" & Me.Gtin14 & "'" & vbCrLf &
                         ", RegMinAgr = '" & Me.RegistroMinisterioAgricultura & "'" & vbCrLf &
                         ",CustoIndireto = " & CByte(CustoIndireto) & vbCrLf &
                         ",CodigoProdutoTerceiro = '" & Me.CodigoProdutoTerceiro & "'" & vbCrLf &
                         ",Dashboard = '" & Me.Dashboard & "'" & vbCrLf &
                         ",Seguimento = " & Me.CodigoSeguimento &
                         " where(produto_id = '" & Me.Codigo & "')"
                sqls.Add(strSql)
                SalvarTabelasRelacionadasSql(sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(sqls)
                strSql = "delete from produtos " & vbCrLf &
                         " where(produto_id = '" & Me.Codigo & "')"
                sqls.Add(strSql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not ProdutoXEmbalagens Is Nothing AndAlso ProdutoXEmbalagens.Count > 0 Then
            For Each emb As ProdutoXEmbalagem In _ProdutoXEmbalagens
                If Me.IUD = "I" Or Me.IUD = "D" Then emb.IUD = Me.IUD
                emb.SalvarSql(Sqls)
            Next
        End If

        UnidadesDeComercializacao.SalvarSql(Sqls)

        If Me.IUD = "I" Or Me.IUD = "D" Then
            ProdutoXEspecificacao.SalvarSql(Sqls)
        End If

        If Me.IUD = "I" Or Me.IUD = "D" Then
            ProdutoXEPI.SalvarSql(Sqls)
        End If

        If Me.IUD = "I" Or Me.IUD = "D" Then
            ProdutoXProcedimento.SalvarSql(Sqls)
        End If

        If Not ProdutosAgrupados Is Nothing AndAlso ProdutosAgrupados.Count > 0 Then
            ProdutosAgrupados.SalvarSQL(Sqls)
        End If
    End Sub
#End Region

End Class