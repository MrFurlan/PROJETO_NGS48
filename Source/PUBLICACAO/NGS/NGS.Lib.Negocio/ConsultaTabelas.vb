Imports System.Data
Imports System.Web
Imports System.Web.UI
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ConsultaTabelas
    Dim Banco As New AcessaBanco

#Region "Selecao de Tabelas"

    Public Function SelecaoTabelas(ByVal Tabela As String)
        Dim Texto As String = String.Empty
        Dim i As Integer
        Dim Sql As String = String.Empty
        Dim Tipo As String

        Tipo = "I"

        Select Case Tabela
            Case "CentrosDeCustos"
                Sql = "SELECT CentroDeCusto_Id as Codigo, Descricao FROM CentrosDeCustos Order By CentroDeCusto_Id"
            Case "TiposDeLancamentos"
                Sql = "SELECT TipoDeLancamento_Id as Codigo, Descricao FROM TiposDeLancamentos Order By TipoDeLancamento_Id"
            Case "Provisoes"
                Sql = "SELECT Provisao_ID as Codigo, Descricao FROM Provisoes Order By Provisao_Id"
            Case "TiposDePagamentos"
                Sql = "SELECT TipoDePagamento_ID as Codigo, Descricao FROM TiposDePagamentos Order By TipoDePagamento_Id"
            Case "ReceberPagar"
                Sql = "SELECT ReceberPagar_ID as Codigo, Descricao FROM ReceberPagar Order By ReceberPagar_Id"
                Tipo = "T"
            Case "Historicos"
                Sql = "SELECT Historico_ID as Codigo, Descricao FROM Historicos Order By Historico_Id"
            Case "Pais"
                Sql = "Select Pais_Id as Codigo, Descricao From Pais Order by Descricao"
            Case "Bancos"
                Sql = "Select Banco_Id as Codigo, Descricao From Bancos Order by Descricao"
            Case "TiposDeClientes"
                Sql = "Select Tipo_Id as Codigo, Descricao From TiposDeClientes Order by Descricao"
            Case "TiposDeLancamentos"
                Sql = "Select TipoDeLancamento_Id as Codigo, Descricao From TiposDeLancamentos Order by Descricao"
            Case "Estados"
                Sql = "Select Estado_Id as Codigo, Descricao From Estados Order by Estado_Id"
                Tipo = "T"
            Case "Situacoes"
                Sql = "Select Situacao_Id as Codigo, Descricao From Situacoes Order by Situacao_Id"
            Case "Categorias"
                Sql = "Select Categoria_Id as Codigo, Descricao From Categorias Order by Categoria_Id"
            Case "Regioes"
                Sql = "Select Regiao_Id as Codigo, Descricao From Regioes Order by Regiao_Id"
            Case "Usuarios"
                Sql = "Select Usuario_Id as Codigo, NomeCompleto as Descricao From Usuarios Order by Usuario_Id"
                Tipo = "T"
            Case "GruposXUsuarios"
                Sql = "SELECT Grupo_Id As Codigo, Usuario_Id As Descricao FROM GruposXUsuarios order by Grupo_Id"
                Tipo = "T"
            Case "Grupos"
                Sql = "SELECT Grupo_Id As Codigo, Descricao FROM Grupos order by Grupo_Id"
            Case "GruposDeEstoques"
                Sql = "SELECT Grupo_Id As Codigo, Descricao FROM GruposDeEstoques order by Grupo_Id"
                Tipo = "T"
            Case "Permissoes"
                Sql = "SELECT Permissao_Id As Codigo, Descricao FROM Permissoes order by Permissao_Id"
                Tipo = "T"
            Case "Produtos"
                Sql = "SELECT Produto_Id As Codigo, Nome as Descricao FROM Produtos order by Produto_Id"
                Tipo = "T"
            Case "RotinasDeImpressao"
                Sql = "SELECT Rotina_Id As Codigo, Descricao FROM RotinasDeImpressao order by Rotina_Id"
                Tipo = "T"
            Case "ComprasXProdutos"
                Sql = "SELECT Produto_Id As Codigo, Descricao FROM ComprasXProdutos order by Descricao"
                Tipo = "T"
            Case "TitulosXFamiliaDeProdutos"
                Sql = "SELECT Produto_Id As Codigo, Descricao FROM TitulosXFamiliaDeProdutos order by Descricao"
                Tipo = "T"
            Case "Autorizantes"
                Sql = "SELECT Autorizante_Id As Codigo, Nivel as Descricao FROM Autorizantes order by Autorizante_Id"
                Tipo = "T"
            Case "Analises"
                Sql = "Select Analise_Id as Codigo, Descricao as Descricao FROM Analises order by Analise_Id"
            Case "Operacoes"
                Sql = "Select Operacao_Id as Codigo, Descricao FROM Operacoes order by Operacao_Id"
                Tipo = "T"
            Case "SituacaoTributaria"
                Sql = "Select SituacaoTributaria_Id as Codigo, Descricao FROM SituacaoTributaria order by SituacaoTributaria_Id"
            Case "CfopTitulo"
                Sql = "Select GrupoCfop_Id as Codigo, Descricao FROM CfopTitulo order by GrupoCfop_Id"
                Tipo = "T"
            Case "Encargos"
                Sql = "Select Encargo_Id as Codigo, Descricao FROM Encargos order by Encargo_Id"
                Tipo = "T"
            Case "TiposDeVeiculos"
                Sql = "Select Codigo_Id as Codigo, Descricao FROM TiposDeVeiculos order by Codigo_Id"
            Case "ViaDeTransportes"
                Sql = "Select Codigo_Id as Codigo, Descricao FROM ViaDeTransportes order by Codigo_Id"
            Case "TabelaDeClassificacoes"
                Sql = "Select Codigo_Id as Codigo, Descricao FROM TabelaDeClassificacoes order by Codigo_Id"
        End Select

        i = 0
        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Tabelas").Tables(0).Rows

            If Tipo = "T" Then
                Texto &= RTrim(Dr("Codigo")) & "|"
                Texto &= Dr("Descricao") & "]"
            Else
                Texto &= Format(Dr("Codigo"), "000") & "|"
                Texto &= Dr("Descricao") & "]"
            End If

            i = i + 1

        Next

        Return Texto

    End Function

    Public Function LerTabelas(ByVal Codigo As String, ByVal Tabela As String)
        Dim Texto As String = String.Empty
        Dim i As Integer = 0
        Dim Sql As String = String.Empty

        If Tabela = "ReceberPagar" And Codigo = "1" Then
            Codigo = "R"
        End If

        If Tabela = "ReceberPagar" And Codigo = "2" Then
            Codigo = "P"
        End If

        Select Case Tabela
            Case "CentrosDeCustos"
                Sql = "SELECT CentroDeCusto_Id as Codigo, Descricao FROM CentrosDeCustos Where CentroDeCusto_Id = " & Codigo
            Case "TiposDeLancamentos"
                Sql = "SELECT TipoDeLancamento_Id as Codigo, Descricao FROM TiposDeLancamentos Where TipoDeLancamento_Id = " & Codigo
            Case "Provisoes"
                Sql = "SELECT Provisao_ID as Codigo, Descricao FROM Provisoes Where Provisao_Id = " & Codigo
            Case "TiposDePagamentos"
                Sql = "SELECT TipoDePagamento_ID as Codigo, Descricao FROM TiposDePagamentos Where TipoDePagamento_Id = " & Codigo
            Case "ReceberPagar"
                Sql = "SELECT ReceberPagar_ID as Codigo, Descricao FROM ReceberPagar Where ReceberPagar_Id = '" & Codigo & "'"
            Case "Historicos"
                Sql = "SELECT Historico_ID as Codigo, Descricao FROM Historicos Where Historico_Id = " & Codigo
            Case "Produtos"
                Sql = "SELECT Produto_ID as Codigo, Nome as Descricao FROM Produtos Where Produto_Id = '" & Codigo & "'"
            Case "Autorizantes"
                Sql = "SELECT Autorizante_ID as Codigo, Solicitante_ID as Descricao  FROM AutorizantesXSolicitantes Where Solicitante_ID = " & Codigo
            Case "Operacoes"
                Sql = "SELECT Operacao_Id as Codigo, Descricao FROM Operacoes Where Operacao_Id = " & Codigo
            Case "CfopTitulo"
                Sql = "SELECT GrupoCfop_Id as Codigo, Descricao FROM CfopTitulo Where GrupoCfop_Id = " & Codigo
            Case "GruposDeEstoques"
                Sql = "SELECT Grupo_Id as Codigo, Descricao FROM GruposDeEstoques Where Grupo_Id = '" & Codigo & "'"
            Case "Encargos"
                Sql = "SELECT Encargo_Id as Codigo, Descricao FROM Encargos Where Encargo_Id = '" & Codigo & "'"
            Case "SituacaoTributaria"
                Sql = "SELECT SituacaoTributaria_Id as Codigo, Descricao FROM SituacaoTributaria Where SituacaoTributaria_Id = " & Codigo
            Case "Estados"
                Sql = "SELECT Estado_Id as Codigo, Descricao FROM Estados Where Estado_Id = '" & Codigo & "'"
            Case "TiposDeVeiculos"
                Sql = "SELECT Codigo_Id as Codigo, Descricao FROM TiposDeVeiculos Where Codigo_Id = " & Codigo & ""
            Case "ViaDeTransportes"
                Sql = "SELECT Codigo_Id as Codigo, Descricao FROM ViaDeTransportes Where Codigo_Id = " & Codigo & ""
            Case "TabelaDeClassificacoes"
                Sql = "SELECT Codigo_Id as Codigo, Descricao FROM TabelaDeClassificacoes Where Codigo_Id = " & Codigo & ""
        End Select

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Tabelas").Tables(0).Rows

            Texto &= Dr("Codigo") & "|"
            Texto &= Dr("Descricao")

        Next

        Return Texto

    End Function

#End Region

#Region "Retorna Autorizante ou Solicitante / Nível"

    Public Function LerAutorizantes()
        Dim Texto As String
        Dim Sql As String

        Sql = "SELECT Autorizante_Id, Nivel From Autorizantes Where Autorizante_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Autorizantes").Tables(0).Rows
            Texto &= RTrim(Dr("Autorizante_id")) & "|"
            Texto &= Dr("Nivel")
        Next

        Return Texto

    End Function

    Public Function SelecaoAutorizantes()
        Dim Texto As String = String.Empty
        Dim Sql As String

        Sql = "SELECT AutorizantesXSolicitantes.Autorizante_Id, Autorizantes.Nivel "
        Sql &= "FROM AutorizantesXSolicitantes INNER JOIN "
        Sql &= "Autorizantes ON AutorizantesXSolicitantes.Solicitante_Id = Autorizantes.Autorizante_Id "
        Sql &= "WHERE AutorizantesXSolicitantes.Solicitante_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "' "
        Sql &= "ORDER BY AutorizantesXSolicitantes.Autorizante_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "AutorizantesXSolicitantes").Tables(0).Rows
            Texto &= RTrim(Dr("Autorizante_id")) & "|"
            Texto &= Dr("Nivel") & "]"
        Next

        Return Texto

    End Function

    Public Function SelecaoTodosAutorizantes()
        Dim Texto As String = String.Empty
        Dim Sql As String = String.Empty

        Sql = " SELECT Autorizante_Id, Nivel "
        Sql &= " FROM Autorizantes"
        Sql &= " ORDER BY Autorizante_Id"

        Dim dsAutorizantes As DataSet = Banco.ConsultaDataSet(Sql, "Autorizantes")

        For Each Dr As DataRow In dsAutorizantes.Tables(0).Rows
            Texto &= RTrim(Dr("Autorizante_id")) & "|"
            Texto &= Dr("Nivel") & "]"
        Next

        Return Texto

    End Function

#End Region

#Region "Selecao de Carteiras"

    Public Function SelecaoCarteiras(ByVal Tabela As String)
        Dim Sql As String = String.Empty
        Dim Texto As String = String.Empty

        Sql = "SELECT Produto_Id As Codigo, Descricao FROM ComprasXProdutos Where Situacao = 'A' order by Produto_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Tabelas").Tables(0).Rows
            Texto &= Dr("Codigo") & "|"
            Texto &= Dr("Descricao") & "]"
        Next

        Return Texto

    End Function

#End Region

#Region "Selecao de Clientes"

    Public Function SelecaoClientes(ByVal CNPJ As String, ByVal Nome As String, ByVal Tipo As String)
        Dim Texto As String
        Dim Sql As String = ""
        Dim CnpjCpf As String
        Dim strCliente As String
        Dim strNome As String
        Dim strCidade As String
        Dim strEnd As String
        Dim strComplemento As String

        If Tipo <> "" Then
            Sql = "SELECT DISTINCT Clientes.Cliente_Id, Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = " & Tipo & ")"
        End If

        If CNPJ <> "" Then
            Sql = "SELECT DISTINCT Clientes.Cliente_Id, Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = " & Tipo & " And Clientes.Cliente_Id Like '" & CNPJ & "%')"
        End If

        If Nome <> "" Then
            Sql = "SELECT DISTINCT Top 100 Clientes.Cliente_Id, Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = " & Tipo & " And Clientes.Nome Like '" & Nome & "%')"
        End If

        If CNPJ = "" And Nome = "" Then
            Sql = "SELECT DISTINCT Top 100 Clientes.Cliente_Id , Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE ClientesXTipos.Tipo_Id = " & Tipo
        End If

        If Tipo = "999" And Nome = "" Then
            Sql = "SELECT DISTINCT Top 100 Clientes.Cliente_Id , Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
        End If

        If Tipo = "999" And Nome <> "" Then
            Sql = "SELECT DISTINCT Top 100 Clientes.Cliente_Id, Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE Clientes.Nome Like '" & Nome & "%'"
        End If

        If Tipo = "999" And CNPJ <> "" Then
            Sql = "SELECT DISTINCT Top 100 Clientes.Cliente_Id, Clientes.Endereco_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Endereco as Endereco, Clientes.Complemento as Complemento, Clientes.Cidade as Cidade, Clientes.Estado as Estado, Clientes.Reduzido as Reduzido, Clientes.Inscricao as Inscricao, Clientes.Imagem as Imagem, isnull(Clientes.Cep,'') AS Cep "
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
            Sql &= " WHERE Clientes.Cliente_Id Like '" & CNPJ & "%'"
        End If

        Sql &= " Order by Clientes.Nome"

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            strCliente = RTrim(Dr("Cliente_Id"))
            strNome = Left(RTrim(Dr("Fantasia")), 40)
            strComplemento = Left(RTrim(Dr("Complemento")), 20)
            strCidade = Left(RTrim(Dr("Cidade")), 25)
            strEnd = Dr("Endereco_Id")
            CnpjCpf = strCliente

            If Len(strCliente) > 11 Then
                CnpjCpf = Left(strCliente, 2) & "." & Mid(strCliente, 3, 3) & "." & Mid(strCliente, 6, 3) & "/" & Mid(strCliente, 9, 4) & "-" & Mid(strCliente, 13, 2)
            ElseIf Len(strCliente) = 11 Then
                CnpjCpf = Left(strCliente, 3) & "." & Mid(strCliente, 4, 3) & "." & Mid(strCliente, 7, 3) & "-" & Mid(strCliente, 10, 2)
            End If

            Texto &= Dr("Reduzido") & StrDup(11 - Len(Dr("Reduzido")), " ") & "|"
            Texto &= CnpjCpf & StrDup(18 - Len(CnpjCpf), " ") & "|"
            Texto &= strEnd & StrDup(3 - Len(strEnd), " ") & "|"
            Texto &= strNome & StrDup(40 - Len(strNome), " ") & "|"
            Texto &= strCidade & StrDup(25 - Len(strCidade), " ") & "|"
            Texto &= Dr("Estado") & "|"
            Texto &= strComplemento & StrDup(20 - Len(strComplemento), " ") & "|"
            Texto &= Dr("Inscricao") & "|"
            Texto &= Dr("Cep") & "]"
        Next

        Return Texto

    End Function

    Public Function LerCliente(ByVal Reduzido As String)
        Dim Texto As String
        Dim Sql As String

        Sql = "SELECT Cliente_Id, Endereco_ID, Nome, isnull(Cep,'') AS Cep "
        Sql &= " FROM Clientes"
        Sql &= " WHERE Reduzido = '" & Reduzido & "'"

        Texto = ""
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome") & "|"
            Texto &= Dr("Cep")
        Next

        Return Texto

    End Function

    Public Function LerClientePorTipo(ByVal Reduzido As String, ByVal Tipo As String)
        Dim Texto As String
        Dim Sql As String

        Sql = "  SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome"
        Sql &= " FROM   Clientes INNER JOIN"
        Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id"
        Sql &= " WHERE  Reduzido = '" & Reduzido & "' AND ClientesXTipos.Tipo_Id = " & Tipo

        Texto = ""
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome")
        Next

        Return Texto

    End Function

#End Region

#Region "Unidade de Negocio X Empresas"

    Public Function UnidadesDeNegociosXEmpresas(ByVal CNPJ As String, ByVal Nome As String, ByVal Tipo As String)
        Dim Texto As String = String.Empty
        Dim Sql As String = String.Empty
        Dim CnpjCpf As String
        Dim strCliente As String
        Dim strNome As String
        Dim strCidade As String
        Dim strEnd As String

        If (CNPJ = "" And Nome = "") Or (CNPJ <> "" And Nome <> "") Then
            Sql = "  SELECT GruposXEmpresas.Empresa_Id as Cliente_Id, GruposXEmpresas.EndEmpresa_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido as Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= "      GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Empresa_Id AND "
            Sql &= "      Clientes.Endereco_Id = GruposXEmpresas.EndEmpresa_Id "
            Sql &= "      RIGHT OUTER JOIN"
            Sql &= "        GruposXEmpresasXUsuarios ON GruposXEmpresas.Empresa_Id = GruposXEmpresasXUsuarios.Empresa_Id"
            Sql &= " WHERE GruposXEmpresasXUsuarios.Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
            Sql &= " GROUP BY GruposXEmpresasXUsuarios.Usuario_Id, GruposXEmpresas.Empresa_Id, GruposXEmpresas.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido"
        ElseIf CNPJ <> "" Then
            Sql = " SELECT GruposXEmpresas.Empresa_Id as Cliente_Id, GruposXEmpresas.EndEmpresa_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido as Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= "      GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Empresa_Id AND "
            Sql &= "      Clientes.Endereco_Id = GruposXEmpresas.EndEmpresa_Id "
            Sql &= "      RIGHT OUTER JOIN"
            Sql &= "        GruposXEmpresasXUsuarios ON GruposXEmpresas.Empresa_Id = GruposXEmpresasXUsuarios.Empresa_Id"
            Sql &= " WHERE Clientes.Cliente_Id Like '" & CNPJ & "%' And GruposXEmpresasXUsuarios.Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
            Sql &= " GROUP BY GruposXEmpresasXUsuarios.Usuario_Id, GruposXEmpresas.Empresa_Id, GruposXEmpresas.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido"
        ElseIf Nome <> "" Then
            Sql = " SELECT GruposXEmpresas.Empresa_Id as Cliente_Id, GruposXEmpresas.EndEmpresa_Id as Endereco_Id, Clientes.Nome as Fantasia, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido as Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= "      GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Empresa_Id AND "
            Sql &= "      Clientes.Endereco_Id = GruposXEmpresas.EndEmpresa_Id "
            Sql &= "      RIGHT OUTER JOIN"
            Sql &= "        GruposXEmpresasXUsuarios ON GruposXEmpresas.Empresa_Id = GruposXEmpresasXUsuarios.Empresa_Id"
            Sql &= " WHERE Clientes.Nome Like '" & Nome & "%' And GruposXEmpresasXUsuarios.Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"
            Sql &= " GROUP BY GruposXEmpresasXUsuarios.Usuario_Id, GruposXEmpresas.Empresa_Id, GruposXEmpresas.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, "
            Sql &= " Clientes.Estado, Clientes.Reduzido"
        End If

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            strCliente = RTrim(Dr("Cliente_Id"))
            strNome = Left(RTrim(Dr("Fantasia")), 40)
            strCidade = Left(RTrim(Dr("Cidade")), 25)
            strEnd = Dr("Endereco_Id")

            If Len(strCliente) > 11 Then
                CnpjCpf = Left(strCliente, 2) & "." & Mid(strCliente, 3, 3) & "." & Mid(strCliente, 6, 3) & "/" & Mid(strCliente, 9, 4) & "-" & Mid(strCliente, 13, 2)
            ElseIf Len(strCliente) = 11 Then
                CnpjCpf = Left(strCliente, 3) & "." & Mid(strCliente, 4, 3) & "." & Mid(strCliente, 7, 3) & "-" & Mid(strCliente, 10, 2)
            Else
                CnpjCpf = strCliente
            End If

            Texto &= Dr("Reduzido") & StrDup(11 - Len(Dr("Reduzido")), " ") & "|"
            Texto &= CnpjCpf & StrDup(18 - Len(CnpjCpf), " ") & "|"
            Texto &= strEnd & StrDup(3 - Len(strEnd), " ") & "|"
            Texto &= strNome & StrDup(40 - Len(strNome), " ") & "|"
            Texto &= strCidade & StrDup(25 - Len(strCidade), " ") & "|"
            Texto &= Dr("Estado") & "]"
        Next

        Return Texto

    End Function

    Public Function LerUnidadesDeNegociosXEmpresas(ByVal Reduzido As String)
        Dim Texto As String = String.Empty
        Dim Sql As String = String.Empty
        Dim Aquisicao As Date = Nothing

        Sql = " SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome"
        Sql &= " FROM Clientes INNER JOIN"
        Sql &= "      GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Empresa_Id AND "
        Sql &= "      Clientes.Endereco_Id = GruposXEmpresas.EndEmpresa_Id "
        Sql &= "      RIGHT OUTER JOIN"
        Sql &= "        GruposXEmpresasXUsuarios ON GruposXEmpresas.Empresa_Id = GruposXEmpresasXUsuarios.Empresa_Id"
        Sql &= " WHERE Clientes.Reduzido = '" & Reduzido & "' And GruposXEmpresasXUsuarios.Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome")
        Next

        Return Texto

    End Function

#End Region

#Region "Empresas X Unidades de Negocios"

    Public Function EmpresasXUnidadesDeNegocios(ByVal CNPJ As String, ByVal Nome As String, ByVal Tipo As String)
        Dim Texto As String = String.Empty
        Dim Sql As String = String.Empty
        Dim CnpjCpf As String
        Dim strCliente As String
        Dim strNome As String
        Dim strCidade As String
        Dim strEnd As String

        If (CNPJ = "" And Nome = "") Or (CNPJ <> "" And Nome <> "") Then
            Sql = " SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome AS Fantasia, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id"
        ElseIf CNPJ <> "" Then
            Sql = " SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome AS Fantasia, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id"
            Sql &= " WHERE GruposXEmpresas.Empresa_Id LIKE '" & Microsoft.VisualBasic.RTrim(CNPJ) & "%'"
        ElseIf Nome <> "" Then
            Sql = " SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome AS Fantasia, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id"
            Sql &= " WHERE GruposXEmpresas.Empresa_Id LIKE '" & Microsoft.VisualBasic.RTrim(CNPJ) & "%'"
            Sql &= " WHERE Clientes.Nome Like '" & Nome & "%'"
        End If

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            strCliente = RTrim(Dr("Cliente_Id"))
            strNome = Left(RTrim(Dr("Fantasia")), 40)
            strCidade = Left(RTrim(Dr("Cidade")), 25)
            strEnd = Dr("Endereco_Id")

            If Len(strCliente) > 11 Then
                CnpjCpf = Left(strCliente, 2) & "." & Mid(strCliente, 3, 3) & "." & Mid(strCliente, 6, 3) & "/" & Mid(strCliente, 9, 4) & "-" & Mid(strCliente, 13, 2)
            Else
                CnpjCpf = Left(strCliente, 3) & "." & Mid(strCliente, 4, 3) & "." & Mid(strCliente, 7, 3) & "-" & Mid(strCliente, 10, 2)
            End If

            Texto &= Dr("Reduzido") & StrDup(11 - Len(Dr("Reduzido")), " ") & "|"
            Texto &= CnpjCpf & StrDup(18 - Len(CnpjCpf), " ") & "|"
            Texto &= strEnd & StrDup(3 - Len(strEnd), " ") & "|"
            Texto &= strNome & StrDup(40 - Len(strNome), " ") & "|"
            Texto &= strCidade & StrDup(25 - Len(strCidade), " ") & "|"
            Texto &= Dr("Estado") & "]"
        Next

        Return Texto

    End Function

    Public Function LerUnidadesDeNegocios(ByVal Reduzido As String)
        Dim Texto As String = String.Empty
        Dim Sql As String = String.Empty
        Dim Aquisicao As Date = Nothing

        Sql = "  SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Nome"
        Sql &= " FROM   Clientes INNER JOIN"
        Sql &= "        GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND "
        Sql &= "        Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id RIGHT OUTER JOIN"
        Sql &= "        GruposXEmpresasXUsuarios ON GruposXEmpresas.Empresa_Id = GruposXEmpresasXUsuarios.Empresa_Id"
        Sql &= " WHERE Clientes.Reduzido = '" & Reduzido & "' And GruposXEmpresasXUsuarios.Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome")
        Next

        Return Texto

    End Function

#End Region

#Region "Empresa do Menu"

    Public Function LerEmpresaDoMenu()
        Dim Texto As String
        Dim Sql As String

        Sql = "SELECT Reduzido, Cliente_Id, Endereco_ID, Nome,Cidade,Estado"
        Sql &= " FROM Clientes"
        Sql &= " WHERE Cliente_Id = '" & HttpContext.Current.Session("ssEmpresa") & "' and Endereco_Id = 0"

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Reduzido") & "|"
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome") & "|"
            Texto &= Dr("Cidade") & "|"
            Texto &= Dr("Estado")
        Next

        If Texto = "" Then
            Texto = "ERRO"
        End If

        Return Texto

    End Function

#End Region

#Region "Selecao Plano de Contas"

    Public Function SelecaoPlanoDeContas(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Grupo As String)
        Dim Texto As String
        Dim Sql As String

        Empresa = Replace(Empresa, ".", "")
        Empresa = Replace(Empresa, "/", "")
        Empresa = Replace(Empresa, "-", "")

        If Grupo = "" Then
            Sql = "SELECT Conta_Id, Titulo"
            Sql &= " FROM PlanoDeContas"
            Sql &= " WHERE LEN(Conta_Id) = 7 "
            'Sql &= " Empresa_Id = '" & Empresa & "' AND EndEmpresa_Id = " & EndEmpresa
        Else
            Sql = "SELECT SUBSTRING(Conta_Id, 8, 2) AS Conta_Id, Titulo"
            Sql &= " FROM PlanoDeContas"
            Sql &= " WHERE LEN(Conta_Id)>7  AND Conta_Id LIKE '" & Grupo & "%' "
            'Sql &= " Empresa_Id = '" & Empresa & "' AND EndEmpresa_Id = " & EndEmpresa
        End If

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "PlanoDeContas").Tables(0).Rows

            Texto &= Dr("Conta_Id") & "|"
            Texto &= Dr("Titulo") & "]"

        Next

        Return Texto

    End Function

    Public Function LerConta(ByVal Empresa As String, ByVal Endereco As String, ByVal Grupo As String, ByVal Conta As String, ByVal Tipo As String) As String
        Dim Texto As String
        Dim Sql As String
        Dim Aquisicao As Date = Nothing
        Dim NL As Integer = 0

        Empresa = Replace(Empresa, ".", "")
        Empresa = Replace(Empresa, "/", "")
        Empresa = Replace(Empresa, "-", "")

        'If Len(Grupo) = 4 Then
        '    Grupo = Left(Grupo, 3) & "0" & Mid(Grupo, 4, 1)
        'End If

        If Conta <> "" Then
            'Select Case Len(Conta)
            '    Case 1
            '        Conta = "000" & Left(Conta, 1)
            '    Case 2
            '        Conta = "00" & Left(Conta, 2)
            '    Case 3
            '        Conta = "0" & Left(Conta, 3)
            'End Select
            Grupo &= Conta
        End If

        If Tipo = "Grupo" Then
            Sql = "SELECT Titulo, Cliente, Produto, CentroDeCusto, TipoDeCliente"
            Sql &= " FROM PlanoDeContas"
            Sql &= " WHERE Conta_Id = " & Grupo & " "
        Else
            Sql = "SELECT Titulo, Cliente, Produto, CentroDeCusto, TipoDeCliente"
            Sql &= " FROM PlanoDeContas"
            Sql &= " WHERE Conta_Id = " & Grupo & " "
        End If

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            If Tipo = "Grupo" Then
                Texto &= Grupo & "|"
            Else
                Texto &= Conta & "|"
            End If
            Texto &= Dr("Titulo") & "|"
            Texto &= Dr("Cliente") & "|"
            Texto &= Dr("Produto") & "|"
            Texto &= Dr("CentroDeCusto") & "|"
            Texto &= Dr("TipoDeCliente")
        Next

        If Texto = "" Then
            Texto = "ERRO"
        End If

        Return Texto

    End Function

    Public Function ConsultaContas(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Conta As String)
        Dim Texto As String
        Dim Sql As String
        Dim strTipo As String
        Dim Titulo As String

        Empresa = Replace(Empresa, ".", "")
        Empresa = Replace(Empresa, "/", "")
        Empresa = Replace(Empresa, "-", "")

        Sql = "SELECT PlanoDeContas.Empresa_Id, PlanoDeContas.EndEmpresa_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, isnull(PlanoDeContas.Cliente, '') as Cliente,"
        Sql &= "isnull(PlanoDeContas.Produto, '') as Produto, isnull(PlanoDeContas.CentroDeCusto, '') as CentroDeCusto, isnull(PlanoDeContas.TipoDeCliente, '0') as TipoDeCliente, isnull(TiposdeClientes.Descricao, '') as Descricao"
        Sql &= " FROM PlanoDeContas INNER JOIN TiposDeClientes ON PlanoDeContas.TipodeCliente = TiposDeClientes.Tipo_Id"
        Sql &= " WHERE (PlanoDeContas.Conta_Id LIKE '" & Conta & "') AND (PlanoDeContas.Empresa_Id = '" & Empresa & "') AND (PlanoDeContas.EndEmpresa_Id = " & EndEmpresa & ")"

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "PlanoDeContas").Tables(0).Rows
            strTipo = Dr("TipoDeCliente")

            Titulo = RTrim(Dr("Titulo"))
            Titulo = Funcoes.EliminarCaracteresEspeciais(Titulo)

            Texto &= Dr("Empresa_Id") & "|"
            Texto &= Dr("EndEmpresa_Id") & "|"
            Texto &= Dr("Conta_Id") & StrDup(9 - Len(Dr("Conta_Id")), " ") & "|"
            Texto &= Titulo & StrDup(80 - Len(Titulo), " ") & "|"
            Texto &= Dr("Cliente") & StrDup(1 - Len(Dr("Cliente")), " ") & "|"
            Texto &= Dr("Produto") & StrDup(1 - Len(Dr("Produto")), " ") & "|"
            Texto &= Dr("CentroDeCusto") & StrDup(1 - Len(Dr("CentroDeCusto")), " ") & "|"
            Texto &= strTipo & StrDup(1 - Len(strTipo), " ") & "|"
            Texto &= Dr("Descricao") & "]"
        Next

        Return Texto

    End Function

#End Region

#Region "Ler Empresa"

    Public Function LerEmpresa(ByVal Cnpj As String, ByVal Endereco As String)
        Dim Texto As String
        Dim Sql As String

        Sql = "SELECT Reduzido, Cliente_Id, Endereco_ID, Nome,Cidade,Estado"
        Sql &= " FROM Clientes"
        Sql &= " WHERE Cliente_Id = '" & Cnpj & "' and Endereco_Id = " & Endereco & ""

        Texto = ""

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Texto &= Dr("Reduzido") & "|"
            Texto &= Dr("Cliente_Id") & "|"
            Texto &= Dr("Endereco_Id") & "|"
            Texto &= Dr("Nome") & "|"
            Texto &= Dr("Cidade") & "|"
            Texto &= Dr("Estado")
        Next

        If Texto = "" Then
            Texto = "ERRO"
        End If

        Return Texto

    End Function

#End Region

#Region "Empresa Matriz"

    Public Function SelecaoEmpresaMatriz(ByVal CNPJ As String, ByVal Nome As String)
        Dim Texto As String
        Dim i As Integer
        Dim Sql As String = ""

        If CNPJ <> "" Then
            Sql = "  SELECT SUBSTRING(Clientes.Cliente_Id, 1, 8) AS Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            Sql &= " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = 1) AND (ClientesXEmpresas.Matriz = 'S') And (Clientes.Cliente_Id Like '" & CNPJ & "%')"
        End If

        If Nome <> "" Then
            Sql = "  SELECT SUBSTRING(Clientes.Cliente_Id, 1, 8) AS Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            Sql &= " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = 1) AND (ClientesXEmpresas.Matriz = 'S') And (Clientes.Nome Like '" & Nome & "%')"
        End If

        If CNPJ = "" And Nome = "" Then
            Sql = "  SELECT SUBSTRING(Clientes.Cliente_Id, 1, 8) AS Cliente_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado"
            Sql &= " FROM Clientes INNER JOIN"
            Sql &= " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            Sql &= " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN"
            Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id"
            Sql &= " WHERE (ClientesXTipos.Tipo_Id = 1) AND (ClientesXEmpresas.Matriz = 'S')"
        End If

        i = 0

        Texto = "<TABLE style='WIDTH: 753px' cellSpacing='-1' cellPadding='-1'>"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

            If i Mod 2 = 0 Then
                Texto &= "<TR class='clLinhaTabela1' onmouseover=""this.className='clLinhaTabela4'"" onmouseout=""this.className='clLinhaTabela1'"">"
            Else
                Texto &= "<TR class='clLinhaTabela3' onmouseover=""this.className='clLinhaTabela4'"" onmouseout=""this.className='clLinhaTabela3'"">"
            End If

            If Mid(HttpContext.Current.Session("ssNomeEmpresa"), 1, 5) = "INSOL" Then
                Texto &= "<TD WIDTH='31px' HEIGHT='21'><IMG height='11' src='images/seta1.gif' width='11' id='BtnSel' style=""Cursor: hand"" onCLICK = ""EmpresaSelecionada('" + Dr("Cliente_Id") + "','" & Dr("Nome") & "','" + Dr("Cidade") + "','" + Dr("Estado") + "')""></TD>"
            Else
                Texto &= "<TD WIDTH='31px'><i class='fa fa-arrow-right seta' id='BtnSel' style=""Cursor: hand"" onCLICK = ""EmpresaSelecionada('" + Dr("Cliente_Id") + "','" & Dr("Nome") & "','" + Dr("Cidade") + "','" + Dr("Estado") + "')""></i></TD>"
            End If

            Texto &= "<TD WIDTH='100px'>" & Dr("Cliente_Id") & "</TD>"
            Texto &= "<TD WIDTH='360px'>" & Dr("Nome") & "</TD>"
            Texto &= "<TD colspan='2'>" & Dr("Cidade") & "</TD>"
            Texto &= "<TD WIDTH='37px'>" & Dr("Estado") & "</TD>"
            Texto &= "</TR>"

            i = i + 1

        Next

        Texto += "</TABLE>"

        Return Texto

    End Function

#End Region

End Class
