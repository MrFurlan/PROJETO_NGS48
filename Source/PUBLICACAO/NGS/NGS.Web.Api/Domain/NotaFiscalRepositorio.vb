Imports NGS.Lib.Negocio
Imports System.IO
Imports System.Net
Imports System.Web
Imports System.Web.Http
Imports System.Net.Http
Imports System.Web.Hosting
Imports System.Web.Mvc
Imports NGS.Lib.Uteis
Imports System.Linq

Public Class NotaFiscalRepositorio
    Implements INotaFiscalRepositorio

    Public Sub New()
    End Sub

    Public Function EmitirNotaFiscal(ByVal fileName As String, ByVal usarProdutoXML As Boolean, ByVal notaDeTerceiro As Boolean, ByVal pedido As String) As String Implements INotaFiscalRepositorio.EmitirNotaFiscal

        Try

            Dim DsXml As New DataSet
            Dim objNotaFiscal As New [Lib].Negocio.NotaFiscal
            Dim filePath As String
            Dim data As DataModel

            If notaDeTerceiro Then
                'filePath = HostingEnvironment.MapPath(String.Format("~/Files/{0}", fileName))
                filePath = String.Format("C:\Alfasig\LeituraNFe\-download\{0}", fileName)
            Else
                filePath = String.Format("C:\Alfasig\LeituraNFe\-xmldestinatario\-nfe\{0}", fileName)
            End If

            If File.Exists(filePath) Then
                If Not Path.GetExtension(filePath).ToLower() = ".xml" Then
                    Throw New Exception("O arquivo existe, mas não é um arquivo XML!")
                End If
            Else
                Throw New Exception("Arquivo não existe! Arquivo" & filePath)
            End If

            data = DataAcess.ReadDataFromXmlFile()
            UsuarioServidor.Conexao = UsuarioLocal.GetStringConexao(data.HashedUsuario, data.NomeServidor, data.Banco, data.Host)

            Dim strConexao As String() = UsuarioServidor.Conexao.ToString.Split(New Char() {";"}, StringSplitOptions.RemoveEmptyEntries)


            If strConexao.Count > 0 Then

                UsuarioServidor.NomeServidor = strConexao(0).Replace("Data Source=", String.Empty).Trim()
                'UsuarioServidor.EnderecoLocal = UsuarioServidor.Conexao.ToString()
                'UsuarioServidor.BancoDeDados = strConexao(1).Replace("Initial Catalog=", String.Empty).Trim()

                ''PARA IMPORTAÇÃO MANUAL PELO VS
                'UsuarioServidor.EnderecoLocal = "Data Source=SRVNGS; Initial Catalog=Horus; User Id=n@$; Password=pwd_ngs123"
                'UsuarioServidor.BancoDeDados = "Horus"

                'response.Cookies("conexao").Value = FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao)
                'Dim response = New HttpResponseMessage(HttpStatusCode.OK) ' Cria uma resposta HTTP com status 200 OK
                Dim cookie As New HttpCookie("conexao", FuncoesStrings.CodificarPara64Bits(UsuarioServidor.Conexao)) ' Cria o cookie com o nome "conexao" e valor "some_value"
                cookie.Expires = DateTime.Now.AddDays(1) ' Define a expiração do cookie para um dia
                HttpContext.Current.Response.Cookies.Add(cookie) ' Adiciona o cookie à resposta

                If String.IsNullOrEmpty(HttpContext.Current.Response.Cookies("conexao").Value) Then

                    Throw New Exception("Usuário não cadastrado para nenhum banco de dados!")

                End If

            Else

                Throw New Exception("Este usuário não existe no banco de dados de Usuários!")

            End If

            'If pNomeArquivo.Contains("-nfe") Then pNomeArquivo = pNomeArquivo.Replace("-nfe", "")
            'DsXml.ReadXml(filePath)

            If Not Funcoes.LimparErroLeituraXML(filePath, DsXml) Then
                Throw New Exception("Erro na leitura do XML!")
            End If

            objNotaFiscal.NotaDeTerceiro = notaDeTerceiro

            'da api, como passar os parametros na url
            DocumentoEletronico.PreencherNFeComXML(objNotaFiscal, DsXml, True, usarProdutoXML, notaDeTerceiro, pedido, False, False)

            HttpContext.Current.Session("ssNomeUsuario") = "AUTOMATICO"

            Dim erroMsg As String

            objNotaFiscal.XMLImportado = True

            'Alterado para fazer importação manual NUBA
            'Ex. chamada na API
            'localhost:44323/api/NotaFiscal/EmitirNotaFiscal?xml=51250353267147000290550010000000491468380371-nfe.xml&usarProdutoXML=false

            If Not String.IsNullOrEmpty(pedido) Then
                'Alterar dados para importar na NUBA
                objNotaFiscal.MicSistemas = True
                objNotaFiscal.CodigoPedido = pedido
                objNotaFiscal.CodigoPedidoMIC = pedido
            End If


            If objNotaFiscal.MicSistemas = False AndAlso objNotaFiscal.NotaDeTerceiro = False Then
                Throw New Exception("Não é uma nota emitida pela MIC!")
            End If

            If objNotaFiscal.DataNota < DateTime.ParseExact("01/07/2024 00:00:00", "dd/MM/yyyy HH:mm:ss", Nothing) Then
                Throw New Exception("A data da nota tem que ser maior que 01/07/2024!")
            End If

            If objNotaFiscal.MicSistemas Then
                objNotaFiscal.Pedido = New Pedido(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.CodigoPedidoMIC)
            Else
                objNotaFiscal.Pedido = New Pedido(objNotaFiscal.Empresa.Codigo, objNotaFiscal.Empresa.CodigoEndereco, objNotaFiscal.CodigoPedido)

                If objNotaFiscal.NotaDeTerceiro Then
                    objNotaFiscal.CarregarDadosNotaDeTerceiro(HttpContext.Current.Session("ssNomeUsuario"))
                End If
            End If

            objNotaFiscal.CarregarDetalhesPedido()

            If objNotaFiscal.VerificarLiberaCarregamento(objNotaFiscal.Pedido.Codigo) = False Then
                Throw New Exception("Pedido aguardando Liberação para Carregamento!")
            End If

            Dim ddlCarteira As New DropDownList
            Dim ddlCondicaoPagamento As New DropDownList
            Dim ddlTipoDePagto As New DropDownList
            Dim parameters = New Dictionary(Of String, Object)
            parameters.Add("tipo", "FRT")
            parameters.Add("Origem", "NF")
            parameters.Add("Indice", objNotaFiscal.IndiceNota)

            'Vencimento
            If Not objNotaFiscal.Pedido.CondicaoPagamento Is Nothing AndAlso objNotaFiscal.Pedido.CondicaoPagamento.Codigo > 0 Then
                objNotaFiscal.CarregarVencimentosDaNota("NF", ddlCarteira, ddlCondicaoPagamento, ddlTipoDePagto, parameters)
            End If

            CadastrarTransportador(objNotaFiscal)

            objNotaFiscal.IUD = "I"
            objNotaFiscal.Usuario = HttpContext.Current.Session("ssNomeUsuario")
            objNotaFiscal.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
            objNotaFiscal.DataInclusao = Format(Date.Now, "yyyy-MM-dd HH:mm:ss")

            If objNotaFiscal.CodigoPedido = 0 And objNotaFiscal.CodigoPedidoMIC > 0 Then
                objNotaFiscal.CodigoPedido = objNotaFiscal.CodigoPedidoMIC
            End If

            If objNotaFiscal.Itens(0).Produto.Agrupar = "N" AndAlso objNotaFiscal.SubOperacao.QuantidadeFisico Then

                Dim Rom As New Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoRomaneio)
                Rom.Codigo = -1
                Rom.CodigoPedido = objNotaFiscal.CodigoPedido
                Rom.Pedido = objNotaFiscal.Pedido
                Rom.CodigoProduto = objNotaFiscal.Itens(0).CodigoProduto
                Rom.EntradaSaida = Left(objNotaFiscal.SubOperacao.EntradaSaida.ToString, 1)

                objNotaFiscal.Romaneio = Rom
                objNotaFiscal.CodigoRomaneio = -1

                objNotaFiscal.CriarRomaneio = True

                Rom.PesoLiquido = objNotaFiscal.Itens(0).QuantidadeFiscal
                Rom.PesoBruto = objNotaFiscal.Itens(0).QuantidadeFiscal

                objNotaFiscal.Itens(0).QuantidadeFisica = objNotaFiscal.Itens(0).QuantidadeFiscal
                objNotaFiscal.Itens(0).PesoBruto = objNotaFiscal.Itens(0).QuantidadeFiscal
                objNotaFiscal.Itens(0).PesoLiquido = objNotaFiscal.Itens(0).QuantidadeFiscal
                objNotaFiscal.Itens(0).PesoFiscal = objNotaFiscal.Itens(0).QuantidadeFiscal

                For Each analise In Rom.DescontosAnalises
                    If Left(objNotaFiscal.Itens(0).CodigoProduto, 5) = "10101" Then
                        If analise.CodigoAnalise = 1 Then
                            analise.Percentual = 14.0
                        ElseIf analise.CodigoAnalise = 2 Then
                            analise.Percentual = 1.0
                        ElseIf analise.CodigoAnalise = 3 Then
                            analise.Percentual = 8.0
                        ElseIf analise.CodigoAnalise = 12 Then
                            analise.Percentual = 3.0
                        End If

                    ElseIf Left(objNotaFiscal.Itens(0).CodigoProduto, 5) = "10102" Then
                        If analise.CodigoAnalise = 1 Then
                            analise.Percentual = 14.0
                        ElseIf analise.CodigoAnalise = 2 Then
                            analise.Percentual = 1.0
                        ElseIf analise.CodigoAnalise = 3 Then
                            analise.Percentual = 5.0
                        End If

                    End If
                Next

            End If

            If objNotaFiscal.ValidarNotaFiscal(DsXml, False, erroMsg, True) Then

                If objNotaFiscal.NossaEmissao AndAlso objNotaFiscal.Eletronica Then
                    objNotaFiscal.CodigoSituacao = 4
                Else
                    objNotaFiscal.CodigoSituacao = 1
                End If

                If Not objNotaFiscal.NotaDeTerceiro AndAlso (objNotaFiscal.Pedido Is Nothing OrElse objNotaFiscal.Pedido.Codigo = 0) Then
                    Throw New Exception("Não foi informado o número do pedido na tag NR_PEDIDO!")
                End If

                If Not objNotaFiscal.NotaDeTerceiro Then
                    objNotaFiscal.Devolucao()
                End If

                objNotaFiscal.ObservacoesEmbarque()

                objNotaFiscal.ObservacoesControleInterno = "INCLUSÃO PELO USUARIO AUTOMATICO"

                If objNotaFiscal.Salvar() Then
                    Return String.Format("Nota: {0} cadastrada com sucesso! XML: {1}", objNotaFiscal.Codigo, fileName)
                Else
                    Throw New Exception(String.Format("Erro ao salvar a nota! Erro: {0}", objNotaFiscal.MsgRetorno))
                End If

            Else
                If erroMsg.Length > 0 Then
                    Throw New Exception(String.Format("Nota: {0} com erro, é necessário depurar! XML: {1} Erro: {2}", objNotaFiscal.Codigo, fileName, erroMsg))
                Else
                    Throw New Exception(String.Format("Nota: {0} com erro, é necessário depurar! XML: {1}", objNotaFiscal.Codigo, fileName))
                End If
            End If

        Catch ex As Exception

            Throw

        End Try

    End Function


    Public Function ValidaCpfCnpj(ByVal documento As String) As Boolean
        If String.IsNullOrWhiteSpace(documento) Then Return False

        ' Remove tudo que não for número
        documento = New String(documento.Where(AddressOf Char.IsDigit).ToArray())

        ' CPF = 11 dígitos | CNPJ = 14 dígitos
        If documento.Length = 11 Then
            Return ValidaCPF(documento)
        ElseIf documento.Length = 14 Then
            Return ValidaCNPJ(documento)
        Else
            Return False
        End If
    End Function

    Private Function ValidaCPF(cpf As String) As Boolean
        If cpf.Length <> 11 Then Return False
        If cpf.Distinct().Count() = 1 Then Return False ' todos os dígitos iguais

        Dim soma As Integer
        Dim resto As Integer

        ' Primeiro dígito
        For i = 1 To 9
            soma += CInt(cpf.Substring(i - 1, 1)) * (11 - i)
        Next
        resto = (soma * 10) Mod 11
        If resto = 10 Then resto = 0
        If resto <> CInt(cpf.Substring(9, 1)) Then Return False

        ' Segundo dígito
        soma = 0
        For i = 1 To 10
            soma += CInt(cpf.Substring(i - 1, 1)) * (12 - i)
        Next
        resto = (soma * 10) Mod 11
        If resto = 10 Then resto = 0
        Return resto = CInt(cpf.Substring(10, 1))
    End Function

    Private Function ValidaCNPJ(cnpj As String) As Boolean
        If cnpj.Length <> 14 Then Return False
        If cnpj.Distinct().Count() = 1 Then Return False ' todos iguais ex: 000000...

        Dim multiplicador1 As Integer() = {5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2}
        Dim multiplicador2 As Integer() = {6, 5, 4, 3, 2, 9, 8, 7, 6, 5, 4, 3, 2}
        Dim tempCnpj As String = cnpj.Substring(0, 12)
        Dim soma As Integer = 0

        For i = 0 To 11
            soma += CInt(tempCnpj(i).ToString()) * multiplicador1(i)
        Next

        Dim resto As Integer = soma Mod 11
        Dim digito As String = If(resto < 2, "0", (11 - resto).ToString())
        tempCnpj &= digito
        soma = 0

        For i = 0 To 12
            soma += CInt(tempCnpj(i).ToString()) * multiplicador2(i)
        Next

        resto = soma Mod 11
        digito &= If(resto < 2, "0", (11 - resto).ToString())

        Return cnpj.EndsWith(digito)
    End Function


    Public Sub CadastrarTransportador(objNotaFiscal As [Lib].Negocio.NotaFiscal)

        Try

            'PESQUISA NA RECEITA FEDERAL DADOS DA EMPRESA
            If ValidaCpfCnpj(RemoveMask(objNotaFiscal.CodigoTransportador)) AndAlso (RemoveMask(objNotaFiscal.CodigoTransportador).Length = 14) Then

                Dim transportador As New Cliente(RemoveMask(objNotaFiscal.CodigoTransportador), 0)

                If transportador.Codigo Is Nothing OrElse transportador.Codigo.Length = 0 Then

                    transportador = New Cliente(RemoveMask(objNotaFiscal.CodigoTransportador))

                    'Remover mascara CEP
                    transportador.CEP = New String(transportador.CEP.Where(AddressOf Char.IsDigit).ToArray())

                    If transportador.Nome.Length = 0 Then
                        objNotaFiscal.Transportador.Nome = "A Receita Federal não disponibilizou as informações desse CNPJ."
                    End If

                    objNotaFiscal.Transportador.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                    objNotaFiscal.Transportador.UsuarioInclusaoData = Now()

                Else
                    'Transportador já existe, não precisamos seguir com o cadastro
                    objNotaFiscal.Transportador = transportador
                End If

                Dim listTiposCliente As New [Lib].Negocio.ListClientexTipo(objNotaFiscal.Transportador)

                If listTiposCliente.Where(Function(x) x.CodigoTipo = 7).Count() = 0 Then

                    objNotaFiscal.Transportador = transportador
                    objNotaFiscal.Transportador.IUD = "I"
                    objNotaFiscal.Transportador.ClienteDesde = String.Format("{0}/{1}/{2}", Now.Day.ToString.PadLeft(2, "0"), Now.Month.ToString.PadLeft(2, "0"), Now.Year)
                    objNotaFiscal.Transportador.CodigoCategoria = 4

                    Dim objTipo As New [Lib].Negocio.ClientexTipo(objNotaFiscal.Transportador)
                    objTipo.IUD = "I"
                    objTipo.CodigoTipo = 7
                    objNotaFiscal.Transportador.Tipos.Add(objTipo)
                    'objTipo = New [Lib].Negocio.ClientexTipo(objNotaFiscal.Transportador)
                    'objTipo.IUD = "I"
                    'objTipo.CodigoTipo = 7
                    'objNotaFiscal.Transportador.Tipos.Add(objTipo)

                End If

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

    Public Function RemoveMask(ByVal pCpfCnpj As String) As String
        'Return pCpfCnpj.Replace(".", "").Replace("-", "").Replace("/", "").Replace("(", "").Replace(")", "").Replace(" ", "")
        Return String.Join("", System.Text.RegularExpressions.Regex.Split(pCpfCnpj, "[^\d]"))
    End Function

End Class
