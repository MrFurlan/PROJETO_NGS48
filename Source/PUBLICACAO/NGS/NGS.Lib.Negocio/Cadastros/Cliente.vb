Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis
Imports Newtonsoft.Json
Imports System.Text
Imports System.Net

<Serializable()>
Public Class ListCliente
    Inherits List(Of Cliente)
    Implements IBaseEntity

    Public Erro As Exception

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Unidade As String, Optional ByVal Tipos As List(Of eTipoCliente) = Nothing, Optional ByVal Cliente As String = "", Optional ByVal Nome As String = "", Optional ByVal NomeFantasia As String = "", Optional ByVal Reduzido As String = "", Optional ByVal categoria As String = "")
        Dim Banco As New AcessaBanco()
        Try
            Dim Sql As String
            Dim blnVirgula As Boolean = False
            Dim strAndWhere As String = "WHERE"

            Sql = "SELECT C.Cliente_Id, C.Endereco_Id, C.Regiao, isnull(C.MicroRegiao,0) as MicroRegiao, C.Categoria, C.Estado, isnull(C.Pais,1058) AS Pais, C.Nome, C.Fantasia, C.Endereco, C.Numero, C.Complemento, C.Bairro, isnull(C.Cep,'') AS Cep, C.Cidade, C.Inscricao, isnull(C.Telefone,'') as Telefone, isnull(C.Fax,'') as Fax, isnull(C.Email,'') as Email, " & vbCrLf &
                  "       C.Imagem, C.Reduzido, isnull(C.CodigoDoMunicipio,0) as CodigoDoMunicipio, C.Situacao, isnull(C.Habilitacao,'') as Habilitacao, C.EmailNFE, isnull(C.OrgaoRegCategoria,'') as OrgaoRegCategoria, isnull(C.UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
                  "       isnull(C.UsuarioInclusaoData,Getdate()) as UsuarioInclusaoData, isnull(C.UsuarioAlteracao,'') as UsuarioAlteracao, isnull(C.UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                  "       isnull(C.OutrosTelefones, '') AS OutrosTelefones, isnull(C.RG, '') AS RG, isnull(C.Site, '') AS Site, isnull(C.Sexo,'I') as Sexo, isnull(C.NascimentoConstituicao,GetDate()) as NascimentoConstituicao, isnull(C.NaturalidadeCidade,'') as NaturalidadeCidade, isnull(C.NaturalidadeEstado,'') as NaturalidadeEstado, isnull(C.ClienteDesde,GetDate()) as ClienteDesde, " & vbCrLf &
                  "       isnull(C.EstadoCivil,'') as EstadoCivil, isnull(C.Suframa, '') AS Suframa, isnull(C.DesdobrarFornecedor, 0) AS DesdobrarFornecedor, isnull(C.RNTRCTransportador,'') AS RNTRCTransportador, " & vbCrLf &
                  "       isnull(ClienteCorrespondencia,'') as ClienteCorrespondencia, isnull(EndClienteCorrespondencia,0) as EndClienteCorrespondencia, C.ApenasAVista " & vbCrLf &
                  "  FROM Clientes C " & vbCrLf


            If Cliente.Length > 0 Then
                Sql &= strAndWhere & " C.Cliente_Id LIKE '" & Cliente & "%' " & vbCrLf
                strAndWhere = " AND "
            End If

            If Nome.Length > 0 Then
                Sql &= strAndWhere & " C.Nome LIKE '" & Nome & "%' " & vbCrLf
                strAndWhere = " AND "
            End If

            If Not String.IsNullOrWhiteSpace(categoria) Then
                Sql &= strAndWhere & " C.Categoria = " & categoria & vbCrLf
                strAndWhere = " AND "
            End If

            If NomeFantasia.Length > 0 Then
                Sql &= strAndWhere & " C.Fantasia LIKE '" & NomeFantasia & "%' " & vbCrLf
                strAndWhere = " AND "
            End If

            If Reduzido.Length > 0 Then
                Sql &= strAndWhere & " C.Reduzido = '" & Reduzido & "' " & vbCrLf
            End If

            If Unidade.Trim.Length > 0 Then
                Sql &= " INNER Join GruposXEmpresas GE " &
                       "    ON GE.Cliente_Id    = C.Cliente_Id " &
                       "   AND GE.EndCliente_Id = C.Endereco_Id " &
                       " WHERE GE.Empresa_Id = '" & Unidade & "'"
            End If

            If Not Tipos Is Nothing Then
                Sql &= " INNER JOIN ClientesXTipos CT " & vbCrLf &
                    "    ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
                    " WHERE CT.Tipo_Id IN ("

                For Each objTipo As Integer In Tipos
                    Sql &= IIf(blnVirgula, ", ", "") & Convert.ToInt32(objTipo)
                    blnVirgula = True
                Next
                Sql &= ") "
            End If

            Sql &= " ORDER BY C.Nome ASC"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cli As New Cliente
                Cli.Codigo = row("Cliente_Id")
                Cli.IsCnpj = IIf(row("Cliente_Id").ToString.Length = 14, True, False)
                Cli.CodigoEndereco = row("Endereco_Id")
                Cli.CodigoRegiao = row("Regiao")
                Cli.CodigoMicroRegiao = row("MicroRegiao")
                Cli.CodigoCategoria = row("Categoria")
                Cli.CodigoEstado = row("Estado")
                Cli.CodigoPais = row("Pais")
                Cli.Nome = row("Nome")
                Cli.Fantasia = row("Fantasia")
                Cli.Endereco = row("Endereco")
                Cli.Numero = row("Numero")
                Cli.Complemento = row("Complemento")
                Cli.Bairro = row("Bairro")
                Cli.CEP = row("Cep")
                Cli.Cidade = row("Cidade")
                Cli.InscricaoEstadual = row("Inscricao")
                Cli.Telefone = row("Telefone")
                Cli.Fax = row("Fax")
                Cli.Email = row("Email")
                Cli.Imagem = row("Imagem")
                Cli.Reduzido = row("Reduzido")
                Cli.CodigoMunicipio = row("CodigoDoMunicipio")
                Cli.CodigoSituacao = row("Situacao")
                Cli.Habilitacao = row("Habilitacao")
                Cli.EmailNFE = row("EmailNFE")
                Cli.OrgaoRegCategoria = row("OrgaoRegCategoria")
                Cli.UsuarioInclusao = row("UsuarioInclusao")
                Cli.UsuarioInclusaoData = row("UsuarioInclusaoData")
                Cli.UsuarioAlteracao = row("UsuarioAlteracao")
                Cli.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
                Cli.OutrosTelefones = row("OutrosTelefones")
                Cli.RG = row("RG")
                Cli.Site = row("Site")
                Cli.Sexo = row("Sexo")
                Cli.NascimentoConstituicao = row("NascimentoConstituicao")
                Cli.NaturalidadeCidade = row("NaturalidadeCidade")
                Cli.NaturalidadeEstado = row("NaturalidadeEstado")
                Cli.ClienteDesde = row("ClienteDesde")
                Cli.EstadoCivil = row("EstadoCivil")
                Cli.Suframa = row("Suframa")
                Cli.DesdobrarFornecedor = row("DesdobrarFornecedor") = "True"
                Cli.RNTRCTransportador = row("RNTRCTransportador")
                Cli.CodigoClienteCorrespondencia = row("ClienteCorrespondencia")
                Cli.EndClienteCorrespondencia = row("EndClienteCorrespondencia")
                Cli.ApenasAVista = row("ApenasAVista")
                Me.Add(Cli)
            Next
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Cli As Cliente In Me
            If Cli.IUD <> "" Then
                Cli.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function FindCliente(ByVal pCodigo As String) As ListCliente
        Dim Banco As New AcessaBanco()
        Dim listCliente As New ListCliente()
        Try
            Dim Sql As String

            Sql = " SELECT Unidade.Cliente_Id AS CodigoUnidade, Unidade.Nome AS Unidade, C.Cliente_Id, C.Endereco_Id, C.Regiao, isnull(C.MicroRegiao,0) as MicroRegiao, C.Categoria, C.Estado, isnull(C.Pais,1058) AS Pais, C.Nome, C.Fantasia, C.Endereco, C.Numero, C.Complemento, C.Bairro, isnull(C.Cep,'') AS Cep, C.Cidade, C.Inscricao, isnull(C.Telefone,'') as Telefone, isnull(C.Fax,'') as Fax, isnull(C.Email,'') as Email, " & vbCrLf &
                    " C.Imagem, C.Reduzido, isnull(C.CodigoDoMunicipio,0) as CodigoDoMunicipio, C.Situacao, isnull(C.Habilitacao,'') as Habilitacao, C.EmailNFE, isnull(C.OrgaoRegCategoria,'') as OrgaoRegCategoria, isnull(C.UsuarioInclusao,'') as UsuarioInclusao, " & vbCrLf &
                    " isnull(C.UsuarioInclusaoData,Getdate()) as UsuarioInclusaoData, isnull(C.UsuarioAlteracao,'') as UsuarioAlteracao, isnull(C.UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData, " & vbCrLf &
                    " isnull(C.OutrosTelefones,'') as OutrosTelefones, isnull(C.RG,'') as RG, isnull(C.Site,'') as Site, isnull(C.Sexo,'I') as Sexo, isnull(C.NascimentoConstituicao,GetDate()) as NascimentoConstituicao, isnull(C.NaturalidadeCidade,'') as NaturalidadeCidade, isnull(C.NaturalidadeEstado,'') as NaturalidadeEstado, isnull(C.ClienteDesde,GetDate()) as ClienteDesde, isnull(C.EstadoCivil,'') as EstadoCivil, " & vbCrLf &
                    " isnull(C.Suframa, '') AS Suframa, isnull(C.DesdobrarFornecedor, 0) AS DesdobrarFornecedor, isnull(C.RNTRCTransportador,'') AS RNTRCTransportador,  " & vbCrLf &
                    " isnull(C.ClienteCorrespondencia,'') as ClienteCorrespondencia, isnull(C.EndClienteCorrespondencia,0) as EndClienteCorrespondencia, C.ApenasAVista " & vbCrLf &
                    " FROM Clientes AS C " & vbCrLf &
                    " LEFT JOIN GruposXEmpresas " & vbCrLf &
                    "    ON GruposXEmpresas.Cliente_Id = C.Cliente_Id " & vbCrLf &
                    "    AND GruposXEmpresas.EndCliente_Id = C.Endereco_Id " & vbCrLf &
                    " LEFT JOIN Clientes AS Unidade " & vbCrLf &
                    "    ON GruposXEmpresas.Empresa_Id = Unidade.Cliente_Id " & vbCrLf &
                    "    AND GruposXEmpresas.EndEmpresa_Id = Unidade.Endereco_Id " & vbCrLf &
                    " WHERE C.Situacao = 1 " & vbCrLf &
                    " AND C.Cliente_Id = '" & pCodigo & "'" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")

            For Each row In ds.Tables(0).Rows
                Dim Cliente As New Cliente()
                Cliente.Codigo = row("Cliente_Id")
                Cliente.CodigoEndereco = row("Endereco_Id")
                If Not IsDBNull(row("CodigoUnidade")) Then Cliente.CodigoUnidadeDeNegocio = row("CodigoUnidade")
                Cliente.CodigoRegiao = row("Regiao")
                Cliente.CodigoMicroRegiao = row("MicroRegiao")
                Cliente.CodigoCategoria = row("Categoria")
                Cliente.CodigoEstado = row("Estado")
                Cliente.CodigoPais = row("Pais")
                Cliente.Nome = row("Nome")
                Cliente.Fantasia = row("Fantasia")
                Cliente.Endereco = row("Endereco")
                Cliente.Numero = row("Numero")
                Cliente.Complemento = row("Complemento")
                Cliente.Bairro = row("Bairro")
                Cliente.CEP = row("Cep")
                Cliente.Cidade = row("Cidade")
                Cliente.InscricaoEstadual = row("Inscricao")
                Cliente.Telefone = row("Telefone")
                Cliente.Fax = row("Fax")
                Cliente.Email = row("Email")
                Cliente.Imagem = row("Imagem")
                Cliente.Reduzido = row("Reduzido")
                Cliente.CodigoMunicipio = row("CodigoDoMunicipio")
                Cliente.CodigoSituacao = row("Situacao")
                Cliente.Habilitacao = row("Habilitacao")
                Cliente.EmailNFE = row("EmailNFE")
                Cliente.OrgaoRegCategoria = row("OrgaoRegCategoria")
                Cliente.UsuarioInclusao = row("UsuarioInclusao")
                Cliente.UsuarioInclusaoData = row("UsuarioInclusaoData")
                Cliente.UsuarioAlteracao = row("UsuarioAlteracao")
                Cliente.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
                Cliente.OutrosTelefones = row("OutrosTelefones")
                Cliente.RG = row("RG")
                Cliente.Site = row("Site")
                Cliente.Sexo = row("Sexo")
                Cliente.NascimentoConstituicao = row("NascimentoConstituicao")
                Cliente.NaturalidadeCidade = row("NaturalidadeCidade")
                Cliente.NaturalidadeEstado = row("NaturalidadeEstado")
                Cliente.ClienteDesde = row("ClienteDesde")
                Cliente.EstadoCivil = row("EstadoCivil")
                Cliente.Suframa = row("Suframa")
                Cliente.DesdobrarFornecedor = row("DesdobrarFornecedor") = "True"
                Cliente.RNTRCTransportador = row("RNTRCTransportador")
                Cliente.CodigoClienteCorrespondencia = row("ClienteCorrespondencia")
                Cliente.EndClienteCorrespondencia = row("EndClienteCorrespondencia")
                Cliente.ApenasAVista = row("ApenasAVista")
                listCliente.Add(Cliente)
            Next
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
        Return listCliente
    End Function
#End Region

End Class

<Serializable()>
Public Class Cliente
    Implements IBaseEntity

#Region "Construtores"

    Public Sub New(ByVal Cnpj As String)

        ServicePointManager.SecurityProtocol = CType(3072, SecurityProtocolType)

        Dim url As String = "https://www.receitaws.com.br/v1/cnpj/" + Cnpj
        Dim webClient As New System.Net.WebClient
        webClient.Encoding = Encoding.UTF8
        Dim res As String = webClient.DownloadString(url)
        Dim cliente As ReceitaFederal = JsonConvert.DeserializeObject(Of ReceitaFederal)(res)

        If cliente.Status = "ERROR" Then
            Exit Sub
        ElseIf cliente.Nome.Length = 0 Then
            Exit Sub
        Else
            If cliente.Situacao.ToUpper = "ATIVA" Then
                Me.CodigoSituacao = 1
            Else
                Me.CodigoSituacao = 4
            End If

            Me.IsCnpj = True
            Me.Codigo = Cnpj
            Me.CodigoEndereco = 0
            Me.Nome = cliente.Nome
            If cliente.Fantasia.Length > 0 Then
                Me.Fantasia = cliente.Fantasia
            Else
                Me.Fantasia = cliente.Nome
            End If

            Me.CEP = Funcoes.EliminarCaracteresEspeciais(cliente.Cep)
            Me.Endereco = cliente.Logradouro

            Dim nrStr As String = Funcoes.RemoverLetrasDoNumero(cliente.Numero.ToUpper)
            nrStr = Funcoes.EliminarCaracteresEspeciais(nrStr)
            If nrStr.Length > 0 Then
                Me.Numero = nrStr
            Else
                Me.Numero = 0
            End If

            Me.Complemento = cliente.Complemento
            Me.Cidade = Funcoes.EliminarCaracteresEspeciais(cliente.Municipio).ToUpper()
            Me.Bairro = cliente.Bairro
            Me.CodigoEstado = cliente.Uf
            Dim _municipio As New Municipio(cliente.Uf, Funcoes.EliminarCaracteresEspeciais(cliente.Municipio))
            Me.CodigoMunicipio = _municipio.CodigoIbge
            Me.Municipio = _municipio
            Me.CodigoPais = 1058

            'Me.Tipos
            Dim data = cliente.Abertura.Split("/")
            Me.NascimentoConstituicao = New Date(data(2), data(1), data(0))
            Me.Telefone = Funcoes.EliminarCaracteresEspeciais(cliente.Telefone)
            Me.Email = cliente.Email
            Me.EmailNFE = cliente.Email
        End If
    End Sub

    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String, ByVal pCodigoEndereco As Integer)

        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT Cliente_Id, Endereco_Id, Regiao, isnull(MicroRegiao,0) as MicroRegiao, Categoria, Estado, isnull(Pais,1058) AS Pais, Nome, Fantasia, Endereco, Numero, Complemento, Bairro, isnull(Cep,'') AS Cep, Cidade, Inscricao, isnull(Telefone,'') as Telefone, isnull(Fax,'') as Fax, isnull(Email,'') as Email, " & vbCrLf &
                  "       Imagem, Reduzido, isnull(CodigoDoMunicipio,0) as CodigoDoMunicipio, Situacao, isnull(Habilitacao,'') as Habilitacao, EmailNFE, isnull(OrgaoRegCategoria,'') as OrgaoRegCategoria, isnull(UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
                  "       isnull(UsuarioInclusaoData,Getdate()) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                  "       isnull(OutrosTelefones,'') as OutrosTelefones, isnull(RG,'') as RG, isnull(Site,'') as Site, isnull(Sexo,'I') as Sexo, isnull(NascimentoConstituicao,GetDate()) as NascimentoConstituicao, isnull(NaturalidadeCidade,'') as NaturalidadeCidade, isnull(NaturalidadeEstado,'') as NaturalidadeEstado, isnull(ClienteDesde,GetDate()) as ClienteDesde, isnull(EstadoCivil,'') as EstadoCivil, " & vbCrLf &
                  "       isnull(Suframa, '') AS Suframa, isnull(DesdobrarFornecedor, 0) AS DesdobrarFornecedor, isnull(RNTRCTransportador,'') AS RNTRCTransportador, " & vbCrLf &
                  "       isnull(ClienteCorrespondencia,'') as ClienteCorrespondencia, isnull(EndClienteCorrespondencia,0) as EndClienteCorrespondencia, ApenasAVista " & vbCrLf &
                  "  FROM Clientes" & vbCrLf &
                  " WHERE Cliente_Id  = '" & pCodigo & "'" & vbCrLf &
                  "   AND Endereco_Id = " & pCodigoEndereco

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.Codigo = row("Cliente_Id")
            Me.IsCnpj = IIf(row("Cliente_Id").ToString.Length = 14, True, False)
            Me.CodigoEndereco = row("Endereco_Id")
            Me.CodigoRegiao = row("Regiao")
            Me.CodigoMicroRegiao = row("MicroRegiao")
            Me.CodigoCategoria = row("Categoria")
            Me.CodigoEstado = row("Estado")
            Me.CodigoPais = row("Pais")
            Me.Nome = row("Nome")
            Me.Fantasia = row("Fantasia")
            Me.Endereco = row("Endereco")
            Me.Numero = row("Numero")
            Me.Complemento = row("Complemento")
            Me.Bairro = row("Bairro")
            Me.CEP = row("Cep")
            Me.Cidade = row("Cidade")
            Me.InscricaoEstadual = row("Inscricao")
            Me.Telefone = row("Telefone")
            Me.Fax = row("Fax")
            Me.Email = row("Email")
            Me.Imagem = row("Imagem")
            Me.Reduzido = row("Reduzido")
            Me.CodigoMunicipio = row("CodigoDoMunicipio")
            Me.CodigoSituacao = row("Situacao")
            Me.Habilitacao = row("Habilitacao")
            Me.EmailNFE = row("EmailNFE")
            Me.OrgaoRegCategoria = row("OrgaoRegCategoria")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.UsuarioInclusaoData = row("UsuarioInclusaoData")
            Me.UsuarioAlteracao = row("UsuarioAlteracao")
            Me.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Me.OutrosTelefones = row("OutrosTelefones")
            Me.RG = row("RG")
            Me.Site = row("Site")
            Me.Sexo = row("Sexo")
            Me.NascimentoConstituicao = row("NascimentoConstituicao")
            Me.NaturalidadeCidade = row("NaturalidadeCidade")
            Me.NaturalidadeEstado = row("NaturalidadeEstado")
            Me.ClienteDesde = row("ClienteDesde")
            Me.EstadoCivil = row("EstadoCivil")
            Me.Suframa = row("Suframa")
            Me.DesdobrarFornecedor = row("DesdobrarFornecedor") = "True"
            Me.RNTRCTransportador = row("RNTRCTransportador")
            Me.CodigoClienteCorrespondencia = row("ClienteCorrespondencia")
            Me.EndClienteCorrespondencia = row("EndClienteCorrespondencia")
            Me.ApenasAVista = row("ApenasAVista") = "True"

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub

    Public Sub BuscarClientePelaInscricaoEstadual()

        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT Cliente_Id, Endereco_Id, Regiao, isnull(MicroRegiao,0) as MicroRegiao, Categoria, Estado, isnull(Pais,1058) AS Pais, Nome, Fantasia, Endereco, Numero, Complemento, Bairro, isnull(Cep,'') AS Cep, Cidade, Inscricao, isnull(Telefone,'') as Telefone, isnull(Fax,'') as Fax, isnull(Email,'') as Email, " & vbCrLf &
                  "       Imagem, Reduzido, isnull(CodigoDoMunicipio,0) as CodigoDoMunicipio, Situacao, isnull(Habilitacao,'') as Habilitacao, EmailNFE, isnull(OrgaoRegCategoria,'') as OrgaoRegCategoria, isnull(UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
                  "       isnull(UsuarioInclusaoData,Getdate()) as UsuarioInclusaoData, isnull(UsuarioAlteracao,'') as UsuarioAlteracao, isnull(UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData," & vbCrLf &
                  "       isnull(OutrosTelefones,'') as OutrosTelefones, isnull(RG,'') as RG, isnull(Site,'') as Site, isnull(Sexo,'I') as Sexo, isnull(NascimentoConstituicao,GetDate()) as NascimentoConstituicao, isnull(NaturalidadeCidade,'') as NaturalidadeCidade, isnull(NaturalidadeEstado,'') as NaturalidadeEstado, isnull(ClienteDesde,GetDate()) as ClienteDesde, isnull(EstadoCivil,'') as EstadoCivil, " & vbCrLf &
                  "       isnull(Suframa, '') AS Suframa, isnull(DesdobrarFornecedor, 0) AS DesdobrarFornecedor, isnull(RNTRCTransportador,'') AS RNTRCTransportador, " & vbCrLf &
                  "       isnull(ClienteCorrespondencia,'') as ClienteCorrespondencia, isnull(EndClienteCorrespondencia,0) as EndClienteCorrespondencia, ApenasAVista " & vbCrLf &
                  "  FROM Clientes" & vbCrLf &
                  " WHERE Cliente_Id  = '" & Me.Codigo & "'" & vbCrLf &
                  "   AND Inscricao = '" & Me.InscricaoEstadual & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.Codigo = row("Cliente_Id")
            Me.IsCnpj = IIf(row("Cliente_Id").ToString.Length = 14, True, False)
            Me.CodigoEndereco = row("Endereco_Id")
            Me.CodigoRegiao = row("Regiao")
            Me.CodigoMicroRegiao = row("MicroRegiao")
            Me.CodigoCategoria = row("Categoria")
            Me.CodigoEstado = row("Estado")
            Me.CodigoPais = row("Pais")
            Me.Nome = row("Nome")
            Me.Fantasia = row("Fantasia")
            Me.Endereco = row("Endereco")
            Me.Numero = row("Numero")
            Me.Complemento = row("Complemento")
            Me.Bairro = row("Bairro")
            Me.CEP = row("Cep")
            Me.Cidade = row("Cidade")
            Me.InscricaoEstadual = row("Inscricao")
            Me.Telefone = row("Telefone")
            Me.Fax = row("Fax")
            Me.Email = row("Email")
            Me.Imagem = row("Imagem")
            Me.Reduzido = row("Reduzido")
            Me.CodigoMunicipio = row("CodigoDoMunicipio")
            Me.CodigoSituacao = row("Situacao")
            Me.Habilitacao = row("Habilitacao")
            Me.EmailNFE = row("EmailNFE")
            Me.OrgaoRegCategoria = row("OrgaoRegCategoria")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.UsuarioInclusaoData = row("UsuarioInclusaoData")
            Me.UsuarioAlteracao = row("UsuarioAlteracao")
            Me.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Me.OutrosTelefones = row("OutrosTelefones")
            Me.RG = row("RG")
            Me.Site = row("Site")
            Me.Sexo = row("Sexo")
            Me.NascimentoConstituicao = row("NascimentoConstituicao")
            Me.NaturalidadeCidade = row("NaturalidadeCidade")
            Me.NaturalidadeEstado = row("NaturalidadeEstado")
            Me.ClienteDesde = row("ClienteDesde")
            Me.EstadoCivil = row("EstadoCivil")
            Me.Suframa = row("Suframa")
            Me.DesdobrarFornecedor = row("DesdobrarFornecedor") = "True"
            Me.RNTRCTransportador = row("RNTRCTransportador")
            Me.CodigoClienteCorrespondencia = row("ClienteCorrespondencia")
            Me.EndClienteCorrespondencia = row("EndClienteCorrespondencia")
            Me.ApenasAVista = row("ApenasAVista") = "True"

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

#Region "Fields"
    Public Erro As Exception

    Private _IUD As String
    Private _Codigo As String = ""
    Private _CodigoFormatado As String
    Private _CodigoEndereco As Integer
    Private _IsCnpj As Boolean
    Public Property IsCnpj() As Boolean
        Get
            Return _IsCnpj
        End Get
        Set(ByVal value As Boolean)
            _IsCnpj = value
        End Set
    End Property

    Private _CodigoRegiao As Integer
    Private _Regiao As Regioes
    Private _CodigoMicroRegiao As Integer
    Private _MicroRegiao As MicroRegiao

    Private _CodigoCategoria As Integer
    Private _Categoria As Categorias
    Private _CodigoEstado As String
    Private _Estado As Estado
    Private _CodigoPais As Integer
    Private _Pais As Pais
    Private _Nome As String
    Private _Fantasia As String
    Private _Endereco As String
    Private _Numero As Integer
    Private _FisicaJuridica As String
    Private _Complemento As String
    Private _Bairro As String
    Private _Cep As String
    Private _Cidade As String
    Private _InscricaoEstadual As String
    Private _Telefone As String
    Private _Fax As String
    Private _Email As String
    Private _Imagem As String
    Private _Reduzido As String
    Private _CodigoMunicipio As Integer
    Private _Municipio As Municipio

    Private _Situacao As Situacao
    Private _CodigoSituacao As Integer

    Private _Habilitacao As String
    Private _EmailNFE As String
    Private _OrgaoRegCategoria As String
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
    Private _OutrosTelefones As String
    Private _RG As String
    Private _Site As String
    Private _Sexo As String
    Private _NascimentoConstituicao As DateTime
    Private _NaturalidadeCidade As String
    Private _NaturalidadeEstado As String
    Private _ClienteDesde As DateTime
    Private _EstadoCivil As String
    Private _Suframa As String
    Private _DesdobrarFornecedor As Boolean

    Private _RNTRCTransportador As String = ""
    Private _CodigoClienteCorrespondencia As String = ""
    Private _EndCorrespondencia As String = ""
    Private _EndClienteCorrespondencia As Integer
    Private _ClienteCorrespondencia As Cliente

    Private _CodigoUnidadeDeNegocio As String
    Private _EnderecoUnidadeDeNegocio As Integer
    Private _UnidadeDeNegocio As Cliente

    Private _Foto As Image

    Private _ApenasAVista As Boolean

    Private _Empresa As ClienteXEmpresa

    Private _ContasBancarias As ListClienteXContaBancaria
    Private _Contatos As ListClienteXContato
    Private _Tipos As ListClientexTipo
    Private _Dependentes As ListClientexDependente
    Private _Veiculos As ListClienteXVeiculo
    Private _Equipamentos As ListClienteXEquipamento
    Private _Imoveis As ListClienteXImovel
    Private _Arrendantes As ListClienteXArrendante
    Private _Representantes As ListClienteXRepresentante
    Private _Safras As ListClienteXSafra
    Private _Socios As ListClienteXSocio
    Private _Matriculas As ListClienteXMatricula
    Private _Financiamentos As ListClienteXFinanciamento
    Private _ReceitasDespesas As ListClienteXReceitasDespesas
    Private _CentrosDeCustos As ListClienteXCentroDeCusto
    Private _Documentos As ListClienteXDocumento
    Private _SubstitutoTributario As ListClienteXSubstitutoTributario
    Private _TabelasDePrecos As ListClienteXTabelaDePreco



#End Region

#Region "Propriedades"
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
            _CodigoFormatado = Funcoes.FormatarCpfCnpj(_Codigo)
        End Set
    End Property

    Public ReadOnly Property CodigoFormatado() As String
        Get
            If _CodigoFormatado = "" And _Codigo <> "" Then _CodigoFormatado = Funcoes.FormatarCpfCnpj(_Codigo)
            Return _CodigoFormatado
        End Get
    End Property

    Public Property CodigoEndereco() As Integer
        Get
            Return _CodigoEndereco
        End Get
        Set(ByVal value As Integer)
            _CodigoEndereco = value
        End Set
    End Property

    Public Property CodigoRegiao() As Integer
        Get
            Return _CodigoRegiao
        End Get
        Set(ByVal value As Integer)
            _CodigoRegiao = value
        End Set
    End Property

    Public Property Regiao() As Regioes
        Get
            If _Regiao Is Nothing And _CodigoRegiao > 0 Then _Regiao = New Regioes(_CodigoRegiao)
            Return _Regiao
        End Get
        Set(ByVal value As Regioes)
            _Regiao = value
        End Set
    End Property

    Public Property CodigoMicroRegiao() As Integer
        Get
            Return _CodigoMicroRegiao
        End Get
        Set(ByVal value As Integer)
            _CodigoMicroRegiao = value
        End Set
    End Property

    Public Property MicroRegiao() As MicroRegiao
        Get
            If _MicroRegiao Is Nothing And _CodigoMicroRegiao > 0 Then _MicroRegiao = New MicroRegiao(Me.Regiao)
            Return _MicroRegiao
        End Get
        Set(ByVal value As MicroRegiao)
            _MicroRegiao = value
        End Set
    End Property

    Public Property CodigoCategoria() As Integer
        Get
            Return _CodigoCategoria
        End Get
        Set(ByVal value As Integer)
            _CodigoCategoria = value
        End Set
    End Property

    Public Property Categoria() As Categorias
        Get
            If _Categoria Is Nothing And _CodigoCategoria > 0 Then _Categoria = New Categorias(_CodigoCategoria)
            Return _Categoria
        End Get
        Set(ByVal value As Categorias)
            _Categoria = value
        End Set
    End Property

    Public Property CodigoEstado() As String
        Get
            Return _CodigoEstado
        End Get
        Set(ByVal value As String)
            _CodigoEstado = value
        End Set
    End Property

    Public Property Estado() As Estado
        Get
            If _Estado Is Nothing AndAlso Not _CodigoEstado Is Nothing AndAlso _CodigoEstado.Length > 0 Then _Estado = New Estado(_CodigoEstado)
            Return _Estado
        End Get
        Set(ByVal value As Estado)
            _Estado = value
        End Set
    End Property

    Public Property CodigoPais() As Integer
        Get
            Return _CodigoPais
        End Get
        Set(ByVal value As Integer)
            _CodigoPais = value
        End Set
    End Property

    Public Property Pais() As Pais
        Get
            If _Pais Is Nothing And _CodigoPais > 0 Then _Pais = New Pais(_CodigoPais)
            Return _Pais
        End Get
        Set(ByVal value As Pais)
            _Pais = value
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

    Public Property Fantasia() As String
        Get
            Return _Fantasia
        End Get
        Set(ByVal value As String)
            _Fantasia = value
        End Set
    End Property

    Public Property Endereco() As String
        Get
            Return _Endereco
        End Get
        Set(ByVal value As String)
            _Endereco = value
        End Set
    End Property

    Public Property Numero() As Integer
        Get
            Return _Numero
        End Get
        Set(ByVal value As Integer)
            _Numero = value
        End Set
    End Property

    Public Property FisicaJuridica As String
        Get
            Return _FisicaJuridica
        End Get
        Set(ByVal value As String)
            _FisicaJuridica = value
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

    Public Property Bairro() As String
        Get
            Return _Bairro
        End Get
        Set(ByVal value As String)
            _Bairro = value
        End Set
    End Property

    Public Property CEP() As String
        Get
            Return _Cep
        End Get
        Set(ByVal value As String)
            _Cep = value
        End Set
    End Property

    Public Property Cidade() As String
        Get
            Return _Cidade
        End Get
        Set(ByVal value As String)
            _Cidade = value
        End Set
    End Property

    Public Property Municipio() As Municipio
        Get
            If _Municipio Is Nothing And _Cidade.Length > 0 And _CodigoEstado.Length > 0 Then _Municipio = New Municipio(_CodigoEstado, _Cidade)
            Return _Municipio
        End Get
        Set(ByVal value As Municipio)
            _Municipio = value
        End Set
    End Property

    Public Property InscricaoEstadual() As String
        Get
            Return _InscricaoEstadual
        End Get
        Set(ByVal value As String)
            _InscricaoEstadual = value
        End Set
    End Property

    Public Property Telefone() As String
        Get
            Return _Telefone
        End Get
        Set(ByVal value As String)
            _Telefone = value
        End Set
    End Property

    Public Property Fax() As String
        Get
            Return _Fax
        End Get
        Set(ByVal value As String)
            _Fax = value
        End Set
    End Property

    Public Property Email() As String
        Get
            Return _Email
        End Get
        Set(ByVal value As String)
            _Email = value
        End Set
    End Property

    Public Property Imagem() As String
        Get
            Return _Imagem
        End Get
        Set(ByVal value As String)
            _Imagem = value
        End Set
    End Property

    Public Property Reduzido() As String
        Get
            Return _Reduzido
        End Get
        Set(ByVal value As String)
            _Reduzido = value
        End Set
    End Property

    Public Property CodigoMunicipio() As Integer
        Get
            Return _CodigoMunicipio
        End Get
        Set(ByVal value As Integer)
            _CodigoMunicipio = value
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
            If _Situacao Is Nothing And CodigoSituacao > 0 Then _Situacao = New Situacao(CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
        End Set
    End Property

    Public Property Habilitacao() As String
        Get
            Return _Habilitacao
        End Get
        Set(ByVal value As String)
            _Habilitacao = value
        End Set
    End Property

    Public Property EmailNFE() As String
        Get
            Return _EmailNFE
        End Get
        Set(ByVal value As String)
            _EmailNFE = value
        End Set
    End Property

    Public Property OrgaoRegCategoria() As String
        Get
            Return _OrgaoRegCategoria
        End Get
        Set(ByVal value As String)
            _OrgaoRegCategoria = value
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

    Public Property OutrosTelefones() As String
        Get
            Return _OutrosTelefones
        End Get
        Set(ByVal value As String)
            _OutrosTelefones = value
        End Set
    End Property

    Public Property RG() As String
        Get
            Return _RG
        End Get
        Set(ByVal value As String)
            _RG = value
        End Set
    End Property

    Public Property Site() As String
        Get
            Return _Site
        End Get
        Set(ByVal value As String)
            _Site = value
        End Set
    End Property

    Public Property Sexo() As String
        Get
            Return _Sexo
        End Get
        Set(ByVal value As String)
            _Sexo = value
        End Set
    End Property

    Public Property NascimentoConstituicao() As DateTime
        Get
            Return _NascimentoConstituicao
        End Get
        Set(ByVal value As DateTime)
            _NascimentoConstituicao = value
        End Set
    End Property

    Public Property NaturalidadeCidade() As String
        Get
            Return _NaturalidadeCidade
        End Get
        Set(ByVal value As String)
            _NaturalidadeCidade = value
        End Set
    End Property

    Public Property NaturalidadeEstado() As String
        Get
            Return _NaturalidadeEstado
        End Get
        Set(ByVal value As String)
            _NaturalidadeEstado = value
        End Set
    End Property

    Public Property ClienteDesde() As DateTime
        Get
            Return _ClienteDesde
        End Get
        Set(ByVal value As DateTime)
            _ClienteDesde = value
        End Set
    End Property

    Public Property EstadoCivil() As String
        Get
            Return _EstadoCivil
        End Get
        Set(ByVal value As String)
            _EstadoCivil = value
        End Set
    End Property

    Public Property Suframa() As String
        Get
            Return _Suframa
        End Get
        Set(ByVal value As String)
            _Suframa = value
        End Set
    End Property

    Public Property DesdobrarFornecedor() As Boolean
        Get
            Return _DesdobrarFornecedor
        End Get
        Set(ByVal value As Boolean)
            _DesdobrarFornecedor = value
        End Set
    End Property

    Public Property CodigoClienteCorrespondencia() As String
        Get
            Return _CodigoClienteCorrespondencia
        End Get
        Set(ByVal value As String)
            _CodigoClienteCorrespondencia = value
            _ClienteCorrespondencia = Nothing
        End Set
    End Property

    Public Property EndCorrespondencia() As String
        Get
            Return _EndCorrespondencia
        End Get
        Set(ByVal value As String)
            _EndCorrespondencia = value
        End Set
    End Property

    Public Property EndClienteCorrespondencia() As Integer
        Get
            Return _EndClienteCorrespondencia
        End Get
        Set(ByVal value As Integer)
            _EndClienteCorrespondencia = value
            _ClienteCorrespondencia = Nothing
        End Set
    End Property

    Public Property ClienteCorrespondencia() As Cliente
        Get
            If _ClienteCorrespondencia Is Nothing And _CodigoClienteCorrespondencia.Length > 0 Then _ClienteCorrespondencia = New Cliente(_CodigoClienteCorrespondencia, _EndClienteCorrespondencia)
            Return _ClienteCorrespondencia
        End Get
        Set(ByVal value As Cliente)
            _ClienteCorrespondencia = value
        End Set
    End Property

    Public Property Foto() As Image
        Get
            Return Nothing
        End Get
        Set(ByVal value As Image)
            _Foto = value
        End Set
    End Property

    Public Property ApenasAVista() As Boolean
        Get
            Return _ApenasAVista
        End Get
        Set(ByVal value As Boolean)
            _ApenasAVista = value
        End Set
    End Property

    Public Property Empresa() As ClienteXEmpresa
        Get
            If _Empresa Is Nothing Then _Empresa = New ClienteXEmpresa(Me.Codigo, Me.CodigoEndereco)
            Return _Empresa
        End Get
        Set(ByVal value As ClienteXEmpresa)
            _Empresa = value
        End Set
    End Property

    Public Property ContasBancarias() As ListClienteXContaBancaria
        Get
            If _ContasBancarias Is Nothing Then _ContasBancarias = New ListClienteXContaBancaria(Me)
            Return _ContasBancarias
        End Get
        Set(ByVal value As ListClienteXContaBancaria)
            _ContasBancarias = value
        End Set
    End Property

    Public Property Contatos() As ListClienteXContato
        Get
            If _Contatos Is Nothing And _Codigo.Length > 0 Then _Contatos = New ListClienteXContato(Me)
            Return _Contatos
        End Get
        Set(ByVal value As ListClienteXContato)
            _Contatos = value
        End Set
    End Property

    Public Property Tipos() As ListClientexTipo
        Get
            If _Tipos Is Nothing Then _Tipos = New ListClientexTipo(Me)
            Return _Tipos
        End Get
        Set(ByVal value As ListClientexTipo)
            _Tipos = value
        End Set
    End Property

    Public Property Dependentes() As ListClientexDependente
        Get
            If _Dependentes Is Nothing Then _Dependentes = New ListClientexDependente(Me)
            Return _Dependentes
        End Get
        Set(ByVal value As ListClientexDependente)
            _Dependentes = value
        End Set
    End Property

    Public Property Veiculos() As ListClienteXVeiculo
        Get
            If _Veiculos Is Nothing Then _Veiculos = New ListClienteXVeiculo(Me)
            Return _Veiculos
        End Get
        Set(ByVal value As ListClienteXVeiculo)
            _Veiculos = value
        End Set
    End Property

    Public Property Equipamentos() As ListClienteXEquipamento
        Get
            If _Equipamentos Is Nothing Then _Equipamentos = New ListClienteXEquipamento(Me)
            Return _Equipamentos
        End Get
        Set(ByVal value As ListClienteXEquipamento)
            _Equipamentos = value
        End Set
    End Property

    Public Property Imoveis() As ListClienteXImovel
        Get
            If _Imoveis Is Nothing Then _Imoveis = New ListClienteXImovel(Me)
            Return _Imoveis
        End Get
        Set(ByVal value As ListClienteXImovel)
            _Imoveis = value
        End Set
    End Property

    Public Property Arrendantes() As ListClienteXArrendante
        Get
            If _Arrendantes Is Nothing Then _Arrendantes = New ListClienteXArrendante(Me)
            Return _Arrendantes
        End Get
        Set(ByVal value As ListClienteXArrendante)
            _Arrendantes = value
        End Set
    End Property

    Public Property Representantes() As ListClienteXRepresentante
        Get
            If _Representantes Is Nothing Then _Representantes = New ListClienteXRepresentante(Me)
            Return _Representantes
        End Get
        Set(ByVal value As ListClienteXRepresentante)
            _Representantes = value
        End Set
    End Property

    Public Property Safras() As ListClienteXSafra
        Get
            If _Safras Is Nothing Then _Safras = New ListClienteXSafra(Me)
            Return _Safras
        End Get
        Set(ByVal value As ListClienteXSafra)
            _Safras = value
        End Set
    End Property

    Public Property Socios() As ListClienteXSocio
        Get
            If _Socios Is Nothing Then _Socios = New ListClienteXSocio(Me)
            Return _Socios
        End Get
        Set(ByVal value As ListClienteXSocio)
            _Socios = value
        End Set
    End Property

    Public Property Matriculas() As ListClienteXMatricula
        Get
            If _Matriculas Is Nothing Then _Matriculas = New ListClienteXMatricula(Me)
            Return _Matriculas
        End Get
        Set(ByVal value As ListClienteXMatricula)
            _Matriculas = value
        End Set
    End Property

    Public Property Financiamentos() As ListClienteXFinanciamento
        Get
            If _Financiamentos Is Nothing Then _Financiamentos = New ListClienteXFinanciamento(Me)
            Return _Financiamentos
        End Get
        Set(ByVal value As ListClienteXFinanciamento)
            _Financiamentos = value
        End Set
    End Property

    Public Property ReceitasDespesas() As ListClienteXReceitasDespesas
        Get
            If _ReceitasDespesas Is Nothing Then _ReceitasDespesas = New ListClienteXReceitasDespesas(Me)
            Return _ReceitasDespesas
        End Get
        Set(ByVal value As ListClienteXReceitasDespesas)
            _ReceitasDespesas = value
        End Set
    End Property

    Public Property SubstitutoTributario() As ListClienteXSubstitutoTributario
        Get
            If _SubstitutoTributario Is Nothing Then _SubstitutoTributario = New ListClienteXSubstitutoTributario(Me)
            Return _SubstitutoTributario
        End Get
        Set(ByVal value As ListClienteXSubstitutoTributario)
            _SubstitutoTributario = value
        End Set
    End Property

    Public Property RNTRCTransportador() As String
        Get
            Return _RNTRCTransportador
        End Get
        Set(ByVal value As String)
            _RNTRCTransportador = value
        End Set
    End Property

    Public Property CodigoUnidadeDeNegocio() As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDeNegocio = value
        End Set
    End Property

    Public Property EnderecoUnidadeDeNegocio() As Integer
        Get
            Return _EnderecoUnidadeDeNegocio
        End Get
        Set(ByVal value As Integer)
            _EnderecoUnidadeDeNegocio = value
        End Set
    End Property

    Public Property UnidadeDeNegocio() As Cliente
        Get
            If _UnidadeDeNegocio Is Nothing Then Return New Cliente(_CodigoUnidadeDeNegocio, _EnderecoUnidadeDeNegocio)
            Return _UnidadeDeNegocio
        End Get
        Set(ByVal value As Cliente)
            _UnidadeDeNegocio = value
        End Set
    End Property

    Public Property Documentos() As ListClienteXDocumento
        Get
            If _Documentos Is Nothing Then _Documentos = New ListClienteXDocumento(Me)
            Return _Documentos
        End Get
        Set(ByVal value As ListClienteXDocumento)
            _Documentos = value
        End Set
    End Property

    Public Property CentrosDeCustos() As ListClienteXCentroDeCusto
        Get
            If _CentrosDeCustos Is Nothing Then _CentrosDeCustos = New ListClienteXCentroDeCusto(Me)
            Return _CentrosDeCustos
        End Get
        Set(ByVal value As ListClienteXCentroDeCusto)
            _CentrosDeCustos = value
        End Set
    End Property

    Public Property TabelasDePrecos() As ListClienteXTabelaDePreco
        Get
            If _TabelasDePrecos Is Nothing Then _TabelasDePrecos = New ListClienteXTabelaDePreco(Me)
            Return _TabelasDePrecos
        End Get
        Set(ByVal value As ListClienteXTabelaDePreco)
            _TabelasDePrecos = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function RelatorioDeClientes(ByVal Parametros As Cliente) As DataSet
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco()
        Dim strAndWhere As String = "WHERE"
        Dim Sql As String

        Sql = "SELECT C.Cliente_Id AS Codigo, C.Endereco_Id AS CodigoEndereco, C.Regiao, C.Categoria, C.Estado AS CodigoEstado, isnull(C.Pais,1058) AS Pais, C.Nome, " & vbCrLf &
              "       C.Fantasia, C.Endereco, C.Numero, C.Complemento, C.Bairro, isnull(C.Cep,'') AS Cep, C.Cidade, C.Inscricao AS InscricaoEstadual, " & vbCrLf &
              "       isnull(C.Telefone,'') as Telefone, isnull(C.Fax,'') as Fax, isnull(C.Email,'') as Email, C.Imagem, C.Reduzido, " & vbCrLf &
              "       isnull(C.CodigoDoMunicipio,0) as CodigoMunicipio, C.Situacao, isnull(C.Habilitacao,'') as Habilitacao, C.EmailNFE, " & vbCrLf &
              "       isnull(C.OrgaoRegCategoria,'') as OrgaoRegCategoria, isnull(C.UsuarioInclusao,'') as UsuarioInclusao, " & vbCrLf &
              "       isnull(C.UsuarioInclusaoData,Getdate()) as UsuarioInclusaoData, isnull(C.UsuarioAlteracao,'') as UsuarioAlteracao, " & vbCrLf &
              "       isnull(C.UsuarioAlteracaoData,GetDate()) as UsuarioAlteracaoData, isnull(C.OutrosTelefones, '') AS OutrosTelefones, " & vbCrLf &
              "       isnull(C.RG, '') AS RG, isnull(C.Site, '') AS Site, isnull(C.Sexo,'I') as Sexo, isnull(C.NascimentoConstituicao,GetDate()) as NascimentoConstituicao, " & vbCrLf &
              "       isnull(C.NaturalidadeCidade,'') as NaturalidadeCidade, isnull(C.NaturalidadeEstado,'') as NaturalidadeEstado, " & vbCrLf &
              "       isnull(C.ClienteDesde,GetDate()) as ClienteDesde, isnull(C.EstadoCivil,'') as EstadoCivil, isnull(C.Suframa, '') AS Suframa, " & vbCrLf &
              "       isnull(C.DesdobrarFornecedor, 0) AS DesdobrarFornecedor, isnull(C.RNTRCTransportador,'') AS RNTRCTransportador, " & vbCrLf &
              "       isnull(C.ClienteCorrespondencia,'') as ClienteCorrespondencia, isnull(C.EndClienteCorrespondencia,0) as EndClienteCorrespondencia, C.ApenasAVista " & vbCrLf &
              "  FROM Clientes C " & vbCrLf

        If Parametros.Codigo.Length > 0 Then
            Sql &= strAndWhere & " C.Cliente_Id LIKE '" & Parametros.Codigo & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoRegiao > 0 Then
            Sql &= strAndWhere & " C.Regiao = " & Parametros.CodigoRegiao & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoCategoria > 0 Then
            Sql &= strAndWhere & " C.Categoria = " & Parametros.CodigoCategoria & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.CodigoEstado Is Nothing AndAlso Parametros.CodigoEstado.Length > 0 Then
            Sql &= strAndWhere & " C.Estado = '" & Parametros.CodigoEstado & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoPais > 0 Then
            Sql &= strAndWhere & " C.Pais = " & Parametros.CodigoPais & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Nome Is Nothing AndAlso Parametros.Nome.Length > 0 Then
            Sql &= strAndWhere & " C.Nome LIKE '" & Parametros.Nome & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Fantasia Is Nothing AndAlso Parametros.Fantasia.Length > 0 Then
            Sql &= strAndWhere & " C.Fantasia LIKE '" & Parametros.Fantasia & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Endereco Is Nothing AndAlso Parametros.Endereco.Length > 0 Then
            Sql &= strAndWhere & " C.Endereco LIKE '" & Parametros.Endereco & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Complemento Is Nothing AndAlso Parametros.Complemento.Length > 0 Then
            Sql &= strAndWhere & " C.Complemento LIKE '" & Parametros.Complemento & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Bairro Is Nothing AndAlso Parametros.Bairro.Length > 0 Then
            Sql &= strAndWhere & " C.Bairro LIKE '" & Parametros.Bairro & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Cidade Is Nothing AndAlso Parametros.Cidade.Length > 0 Then
            Sql &= strAndWhere & " C.Cidade LIKE '" & Parametros.Cidade & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoSituacao > 0 Then
            Sql &= strAndWhere & " C.Situacao = " & Parametros.CodigoSituacao & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Sexo Is Nothing AndAlso Parametros.Sexo.Length > 0 Then
            Sql &= strAndWhere & " C.Sexo = '" & Parametros.Sexo & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.EstadoCivil Is Nothing AndAlso Parametros.EstadoCivil.Length > 0 Then
            Sql &= strAndWhere & " C.EstadoCivil = '" & Parametros.EstadoCivil & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        Return Banco.ConsultaDataSet(Sql, "Clientes")
    End Function

    Public Function RelatorioDeClientesExcel(ByVal Parametros As Cliente) As DataSet
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco()
        Dim strAndWhere As String = "WHERE"
        Dim Sql As String

        Sql = " SELECT C.Cliente_Id AS Codigo, C.Endereco_Id AS CodigoEndereco, C.Nome," & vbCrLf &
              "   C.Fantasia, C.Endereco, C.Numero, C.Complemento, C.Bairro, isnull(C.Cep,'') AS Cep, C.Cidade, C.Estado AS CodigoEstado, C.Inscricao AS InscricaoEstadual,  " & vbCrLf &
              "       isnull(C.Telefone,'') as Telefone, isnull(C.Email,'') as Email, C.EmailNFE  " & vbCrLf &
              "          FROM Clientes C  " & vbCrLf &
              "          inner join ClientesXTipos ct" & vbCrLf &
              "       on ct.Cliente_Id   = C.Cliente_Id " & vbCrLf &
              "  and ct.Endereco_Id = C.Endereco_Id  " & vbCrLf

        If Parametros.Codigo.Length > 0 Then
            Sql &= strAndWhere & " C.Cliente_Id LIKE '" & Parametros.Codigo & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoRegiao > 0 Then
            Sql &= strAndWhere & " C.Regiao = " & Parametros.CodigoRegiao & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoCategoria > 0 Then
            Sql &= strAndWhere & " C.Categoria = " & Parametros.CodigoCategoria & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.CodigoEstado Is Nothing AndAlso Parametros.CodigoEstado.Length > 0 Then
            Sql &= strAndWhere & " C.Estado = '" & Parametros.CodigoEstado & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoPais > 0 Then
            Sql &= strAndWhere & " C.Pais = " & Parametros.CodigoPais & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Nome Is Nothing AndAlso Parametros.Nome.Length > 0 Then
            Sql &= strAndWhere & " C.Nome LIKE '" & Parametros.Nome & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Fantasia Is Nothing AndAlso Parametros.Fantasia.Length > 0 Then
            Sql &= strAndWhere & " C.Fantasia LIKE '" & Parametros.Fantasia & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Endereco Is Nothing AndAlso Parametros.Endereco.Length > 0 Then
            Sql &= strAndWhere & " C.Endereco LIKE '" & Parametros.Endereco & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Complemento Is Nothing AndAlso Parametros.Complemento.Length > 0 Then
            Sql &= strAndWhere & " C.Complemento LIKE '" & Parametros.Complemento & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Bairro Is Nothing AndAlso Parametros.Bairro.Length > 0 Then
            Sql &= strAndWhere & " C.Bairro LIKE '" & Parametros.Bairro & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Cidade Is Nothing AndAlso Parametros.Cidade.Length > 0 Then
            Sql &= strAndWhere & " C.Cidade LIKE '" & Parametros.Cidade & "%' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Parametros.CodigoSituacao > 0 Then
            Sql &= strAndWhere & " C.Situacao = " & Parametros.CodigoSituacao & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.Sexo Is Nothing AndAlso Parametros.Sexo.Length > 0 Then
            Sql &= strAndWhere & " C.Sexo = '" & Parametros.Sexo & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        If Not Parametros.EstadoCivil Is Nothing AndAlso Parametros.EstadoCivil.Length > 0 Then
            Sql &= strAndWhere & " C.EstadoCivil = '" & Parametros.EstadoCivil & "' " & vbCrLf
            strAndWhere = " AND "
        End If

        Return Banco.ConsultaDataSet(Sql, "Clientes")
    End Function


    'DEPOIS QUE FIZER O NOVO PEDIDO EXCLUIR
    Public Function ToDataTableToDsPedidos(ByVal NomeTabela As String) As DataTable
        Dim dtCliente As New DataTable(NomeTabela)
        dtCliente.Columns.Add("Nome", GetType(String))
        dtCliente.Columns.Add("Codigo", GetType(String))
        dtCliente.Columns.Add("Inscricao", GetType(String))
        dtCliente.Columns.Add("Endereco", GetType(String))
        dtCliente.Columns.Add("Telefone", GetType(String))
        dtCliente.Columns.Add("CEP", GetType(String))
        dtCliente.Columns.Add("Cidade", GetType(String))
        dtCliente.Columns.Add("Estado", GetType(String))

        Dim drCliente As DataRow = dtCliente.NewRow()

        drCliente("Nome") = Me.Nome
        drCliente("Codigo") = Me.Codigo
        drCliente("Inscricao") = Me.InscricaoEstadual
        drCliente("Endereco") = Me.Endereco
        drCliente("Telefone") = Me.Telefone
        drCliente("CEP") = Me.CEP
        drCliente("Cidade") = Me.Cidade
        drCliente("Estado") = Me.CodigoEstado
        dtCliente.Rows.Add(drCliente)

        Return dtCliente
    End Function

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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "Insert Into Clientes(Cliente_Id, Endereco_Id, Regiao, MicroRegiao, Categoria, Estado, Pais, Nome, Fantasia, Endereco, Numero, Complemento, Bairro, Cep, Cidade, Inscricao, Telefone, Fax, Email, " & vbCrLf &
                      "                     Imagem, Reduzido, CodigoDoMunicipio, Situacao, Habilitacao, EmailNFE, OrgaoRegCategoria, UsuarioInclusao, UsuarioInclusaoData," & vbCrLf &
                      "                   OutrosTelefones, Rg, Site, Sexo,  NascimentoConstituicao, NaturalidadeCidade, NaturalidadeEstado, ClienteDesde, EstadoCivil, Suframa, DesdobrarFornecedor, RNTRCTransportador, ClienteCorrespondencia, EndClienteCorrespondencia, ApenasAVista)" & vbCrLf &
                      " Values ('" & _Codigo & "'," & _CodigoEndereco & "," & _CodigoRegiao & "," & Me.CodigoMicroRegiao & "," & _CodigoCategoria & ",'" & _CodigoEstado & "'," & _CodigoPais & ",'" & _Nome & "','" & _Fantasia & "','" & _Endereco & "'," & _Numero & ",'" & _Complemento & "','" & _Bairro & "','" & _Cep & "','" & _Cidade & "','" & _InscricaoEstadual & "','" & _Telefone & "','" & _Fax & "','" & _Email & "'," & vbCrLf &
                      "         '" & Me.Imagem & "', '" & _Reduzido & "'," & CodigoMunicipio & "," & CodigoSituacao & ",'" & _Habilitacao & "','" & _EmailNFE & "','" & _OrgaoRegCategoria & "','" & _UsuarioInclusao & "','" & Now.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "'" & _OutrosTelefones & "','" & _RG & "','" & _Site & "','" & _Sexo & "','" & _NascimentoConstituicao.ToString("yyyy-MM-dd") & "','" & _NaturalidadeCidade & "','" & _NaturalidadeEstado & "','" & _ClienteDesde.ToString("yyyy-MM-dd") & "','" & _EstadoCivil & "', '" & _Suframa & "', " & IIf(_DesdobrarFornecedor, 1, 0) & ", '" & _RNTRCTransportador & "','" & Me.CodigoClienteCorrespondencia & "'," & Me.EndClienteCorrespondencia & ", " & IIf(_ApenasAVista, 1, 0) & ")"

                Sqls.Add(sql)
                '***********************************************************************
                '********* Procedimento para Salvar as tabelas relacionadas   **********
                '***********************************************************************
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = " Update Clientes set " & vbCrLf &
                      "    Cliente_Id                ='" & _Codigo & "'" & vbCrLf &
                      "   ,Endereco_Id               = " & _CodigoEndereco & vbCrLf &
                      "   ,Regiao                    = " & _CodigoRegiao & vbCrLf &
                      "   ,MicroRegiao               = " & Me.CodigoMicroRegiao & vbCrLf &
                      "   ,Categoria                 = " & _CodigoCategoria & vbCrLf &
                      "   ,Estado                    ='" & _CodigoEstado & "'" & vbCrLf &
                      "   ,Pais                      = " & _CodigoPais & vbCrLf &
                      "   ,Nome                      ='" & _Nome & "'" & vbCrLf &
                      "   ,Fantasia                  ='" & _Fantasia & "'" & vbCrLf &
                      "   ,Endereco                  ='" & _Endereco & "'" & vbCrLf &
                      "   ,Numero                    = " & _Numero & vbCrLf &
                      "   ,Complemento               ='" & _Complemento & "'" & vbCrLf &
                      "   ,Bairro                    ='" & _Bairro & "'" & vbCrLf &
                      "   ,CEP                       ='" & _Cep & "'" & vbCrLf &
                      "   ,Cidade                    ='" & _Cidade & "'" & vbCrLf &
                      "   ,Inscricao                 ='" & _InscricaoEstadual & "'" & vbCrLf &
                      "   ,Telefone                  ='" & _Telefone & "'" & vbCrLf &
                      "   ,Fax                       ='" & _Fax & "'" & vbCrLf &
                      "   ,Email                     ='" & _Email & "'" & vbCrLf &
                      "   ,Imagem                    ='" & _Imagem & "'" & vbCrLf &
                      "   ,Reduzido                  ='" & _Reduzido & "'" & vbCrLf &
                      "   ,CodigoDoMunicipio         = " & CodigoMunicipio & vbCrLf &
                      "   ,Situacao                  = " & CodigoSituacao & vbCrLf &
                      "   ,Habilitacao               ='" & _Habilitacao & "'" & vbCrLf &
                      "   ,EmailNFE                  ='" & _EmailNFE & "'" & vbCrLf &
                      "   ,OrgaoRegCategoria         ='" & _OrgaoRegCategoria & "'" & vbCrLf &
                      "   ,UsuarioAlteracao          ='" & _UsuarioAlteracao & "'" & vbCrLf &
                      "   ,UsuarioAlteracaoData      ='" & Now.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "   ,OutrosTelefones           ='" & _OutrosTelefones & "'" & vbCrLf &
                      "   ,Rg                        ='" & _RG & "'" & vbCrLf &
                      "   ,Site                      ='" & _Site & "'" & vbCrLf &
                      "   ,Sexo                      ='" & _Sexo & "'" & vbCrLf &
                      "   ,NascimentoConstituicao    ='" & _NascimentoConstituicao.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "   ,NaturalidadeCidade        ='" & _NaturalidadeCidade & "'" & vbCrLf &
                      "   ,NaturalidadeEstado        ='" & _NaturalidadeEstado & "'" & vbCrLf &
                      "   ,ClienteDesde              ='" & _ClienteDesde.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "   ,EstadoCivil               ='" & _EstadoCivil & "'" & vbCrLf &
                      "   ,Suframa                   ='" & _Suframa & "'" & vbCrLf &
                      "   ,DesdobrarFornecedor       = " & IIf(_DesdobrarFornecedor, 1, 0) & "" & vbCrLf &
                      "   ,ClienteCorrespondencia    ='" & Me.CodigoClienteCorrespondencia & "'" & vbCrLf &
                      "   ,EndClienteCorrespondencia = " & Me.EndClienteCorrespondencia & vbCrLf &
                      "   ,RNTRCTransportador        ='" & _RNTRCTransportador & "'" & vbCrLf &
                      "   ,ApenasAVista              = " & IIf(_ApenasAVista, 1, 0) & "" & vbCrLf &
                      "	Where Cliente_Id  ='" & _Codigo & "'" & vbCrLf &
                      "   and Endereco_Id = " & _CodigoEndereco
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"

                SalvarTabelasRelacionadasSql(Sqls)
                sql = " Delete Clientes" & vbCrLf &
                      "	 Where Cliente_Id  = '" & _Codigo & "'" & vbCrLf &
                      "    and Endereco_Id = " & _CodigoEndereco
                Sqls.Add(sql)

            Case "U_TIPO"
                SalvarTabelasRelacionadasSql(Sqls)
        End Select

    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not Contatos Is Nothing Then Contatos.SalvarSql(Sqls)
        If Not ContasBancarias Is Nothing Then ContasBancarias.SalvarSql(Sqls)
        If Not Tipos Is Nothing Then Tipos.SalvarSql(Sqls)
        If Not Dependentes Is Nothing Then Dependentes.SalvarSql(Sqls)
        If Not Veiculos Is Nothing Then Veiculos.SalvarSql(Sqls)
        If Not Equipamentos Is Nothing Then Equipamentos.SalvarSql(Sqls)
        If Not Imoveis Is Nothing Then Imoveis.SalvarSql(Sqls)
        If Not Arrendantes Is Nothing Then Arrendantes.SalvarSql(Sqls)
        If Not Representantes Is Nothing Then Representantes.SalvarSql(Sqls)
        If Not Safras Is Nothing Then Safras.SalvarSql(Sqls)
        If Not Socios Is Nothing Then Socios.SalvarSql(Sqls)
        If Not Matriculas Is Nothing Then Matriculas.SalvarSql(Sqls)
        If Not Financiamentos Is Nothing Then Financiamentos.SalvarSql(Sqls)
        If Not ReceitasDespesas Is Nothing Then ReceitasDespesas.SalvarSql(Sqls)
        If Not CentrosDeCustos Is Nothing Then CentrosDeCustos.SalvarSql(Sqls)
        If Not Documentos Is Nothing Then Documentos.SalvarSql(Sqls)
        If Not SubstitutoTributario Is Nothing Then SubstitutoTributario.SalvarSql(Sqls)
        If Not TabelasDePrecos Is Nothing Then TabelasDePrecos.SalvarSql(Sqls)


        If _CodigoEstado = "EX" Then
            Dim objMunicipio As New Municipio(_CodigoEstado, _Cidade)

            If objMunicipio Is Nothing OrElse objMunicipio.CodigoIbge = 0 Then
                objMunicipio.IUD = "I"
            Else
                objMunicipio.IUD = "U"
            End If

            objMunicipio.CodigoEstado = _CodigoEstado
            objMunicipio.CodigoMunicipio = _Cidade
            objMunicipio.CodigoIbge = 9999999
            objMunicipio.EstadoIbge = 99

            objMunicipio.SalvarSql(Sqls)
        End If
    End Sub

#End Region

End Class