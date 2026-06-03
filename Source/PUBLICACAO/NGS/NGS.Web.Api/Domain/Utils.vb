Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Namespace Utils

    ''' <summary>
    ''' Enumerador do Tipo de Agrupamento para detalhamento do Cliente
    ''' </summary>
    Public Enum TipoRelatorioCliente
        TotalCliente = 1
        TotalPorUnidadeCliente = 2
    End Enum

    ''' <summary>
    ''' Enumerador para o Tipo de PDF a ser Gerado
    ''' </summary>
    Public Enum TipoImpressaoPDF
        AgroindustriaListaPosicaoPedidos = 1
        AgroindustriaListaPosicaoClientes = 2
        AgroindustriaExtratoPorPedido = 3
        AgroindustriaExtratoPorCliente = 4
        AgroindustriaPosicaoClienteExternal = 5
        AgroindustriaExtratoPorPedidoExternal = 6
    End Enum

    ''' <summary>
    ''' Situação das Requisições
    ''' </summary>
    Public Enum SituacaoRequisicao
        Todas = 0
        EnviadoTriagem = 301
        DevolvidoRequisitante = 302
        RealizadaTriagemEncaminhadaAutorizacao = 305
        DevolvidoTriagem = 306
        Autorizadas = 310
        DevolvidoAutorizacao = 311
        SolicitadoRevisao = 314
        RequisicaoAutorizadaEnviadaCotacao = 315
        Pendentes = 316
        CompraEfetivadaAguardandoEntrega = 320
        OrdemEntregue = 325
        DevolvidaAlmoxarifadoEstoque = 330
        Cancelado = 331
    End Enum

    ''' <summary>
    ''' Enumerador para os Acessos de Áreas de Menu
    ''' </summary>
    Public Enum AreaAcessoMenu
        Padrao = 0
        Agroindustria = 1
        Agropecuarias = 2
        Concessionarias = 3
        Portos = 4
        Shopping = 5
        Usinas = 6
    End Enum

    Public NotInheritable Class SituacaoRequisicaoLabel
        Public Shared ReadOnly Property Label As Dictionary(Of Integer, String)
            Get
                Dim labels As New Dictionary(Of Integer, String) From {
                    {SituacaoRequisicao.Todas, ""},
                    {SituacaoRequisicao.EnviadoTriagem, "Enviado para Triagem"},
                    {SituacaoRequisicao.DevolvidoRequisitante, "Devolvido para o Requisitante"},
                    {SituacaoRequisicao.RealizadaTriagemEncaminhadaAutorizacao, "Realizada triagem e encaminhada para Autorização"},
                    {SituacaoRequisicao.DevolvidoTriagem, "Devolvido para Triagem"},
                    {SituacaoRequisicao.Autorizadas, "Autorizada"},
                    {SituacaoRequisicao.DevolvidoAutorizacao, "Devolvido para Autorização"},
                    {SituacaoRequisicao.SolicitadoRevisao, "Solicitado Revisão"},
                    {SituacaoRequisicao.RequisicaoAutorizadaEnviadaCotacao, "Aguardando Cotação"},
                    {SituacaoRequisicao.Pendentes, "Pendente"},
                    {SituacaoRequisicao.CompraEfetivadaAguardandoEntrega, "Compra Efetivada - Aguardando Entrega"},
                    {SituacaoRequisicao.OrdemEntregue, "Ordem Entregue"},
                    {SituacaoRequisicao.DevolvidaAlmoxarifadoEstoque, "Devolvida para Almoxarifado/Estoque"},
                    {SituacaoRequisicao.Cancelado, "Cancelada"}
                }
                Return labels
            End Get
        End Property
    End Class

    Public Enum TipoDocumento
        Todos = -1
        PDF = 1
        Email = 2
        Html = 3
        Xml = 4
        AnexoEmail = 5
        AnexoOrcamento = 6
        Indefinido = 0
    End Enum

    Public Enum SituacaoRegistro
        Normal = 1
        Cancelado = 2
        Liquidado = 6
        Todos = 0
    End Enum

    Public NotInheritable Class SituacaoRegistroLabel
        Public Shared ReadOnly Property Label As Dictionary(Of Integer, String)
            Get
                Dim labels As New Dictionary(Of Integer, String) From {
                    {SituacaoRegistro.Normal, "ABERTO"},
                    {SituacaoRegistro.Cancelado, "CANCELADO"},
                    {SituacaoRegistro.Liquidado, "LIQUIDADO"},
                    {SituacaoRegistro.Todos, "TODOS"}
                }
                Return labels
            End Get
        End Property
    End Class

End Namespace

Public Module ColorUtils

    Public Function HslToHex(hue As Integer, saturation As Integer, lightness As Integer) As String

        Dim s = saturation / 100.0
        Dim l = lightness / 100.0
        Dim c = (1 - Math.Abs(2 * l - 1)) * s
        Dim x = c * (1 - Math.Abs((hue / 60.0) Mod 2 - 1))
        Dim m = l - c / 2

        Dim r As Double, g As Double, b As Double

        Select Case hue
            Case < 60
                r = c : g = x : b = 0
            Case < 120
                r = x : g = c : b = 0
            Case < 180
                r = 0 : g = c : b = x
            Case < 240
                r = 0 : g = x : b = c
            Case < 300
                r = x : g = 0 : b = c
            Case Else
                r = c : g = 0 : b = x
        End Select

        Dim rHex = CInt((r + m) * 255)
        Dim gHex = CInt((g + m) * 255)
        Dim bHex = CInt((b + m) * 255)

        Return $"#{rHex:X2}{gHex:X2}{bHex:X2}"

    End Function

End Module

Public Module SessionHelper

    Public Function ObterProdutoDashboardFiltro(request As HttpRequestBase, session As HttpSessionStateBase, ByVal repository As ProdutoDashboardRepositorio) As ProdutoDashboardFiltro

        Dim filtro As ProdutoDashboardFiltro


        If UsuarioServidor.Usuario Is Nothing Then

            If UsuarioServidor.Conexao Is Nothing Then
                Dim cookie As HttpCookie = request.Cookies("conexao")
                If cookie IsNot Nothing Then
                    UsuarioServidor.Conexao = FuncoesStrings.DecodificarDe64Bits(cookie.Value)
                End If
            End If

            UsuarioServidor.CarregarInformacaoParaUsuarioServidor(HttpContext.Current.User.Identity.Name)

        End If

        If session("ProdutoDashboardFiltro") Is Nothing Then

            Dim produto As New ProdutoDashboard
            Dim empresas = repository.ConsultarEmpresas
            Dim seguimentos = repository.ConsultarSeguimentos
            Dim inicio = If(String.IsNullOrEmpty(request("dataInicio")), Date.Now.AddDays(-30).ToString(), request("dataInicio"))
            Dim fim = If(String.IsNullOrEmpty(request("dataFim")), Date.Now.ToString(), request("dataFim"))
            Dim empresaSelecionada = If(String.IsNullOrEmpty(UsuarioServidor.EmpresasDashboard), UsuarioServidor.CodigoEmpresa, UsuarioServidor.EmpresasDashboard)
            Dim seguimentoSelecionado = If(String.IsNullOrEmpty(UsuarioServidor.SeguimentoDashboard), "TODOS", UsuarioServidor.SeguimentoDashboard)
            Dim empresaUsuario = String.Format("{0}|{1}", UsuarioServidor.CodigoEmpresa, UsuarioServidor.EnderecoEmpresa)
            Dim dataInicioUsuario = UsuarioServidor.DataInicioDashboard
            Dim dataFimUsuario = UsuarioServidor.DataFimDashboard

            If Not String.IsNullOrEmpty(empresaSelecionada) Then
                empresaSelecionada = empresaSelecionada
            Else
                empresaSelecionada = empresaUsuario
            End If

            If String.IsNullOrEmpty(seguimentoSelecionado) Then
                seguimentoSelecionado = "0"
            End If

            If dataInicioUsuario.Year > 1 Then
                inicio = dataInicioUsuario
            End If

            If dataFimUsuario.Year > 1 Then
                fim = dataFimUsuario
            End If

            filtro = New ProdutoDashboardFiltro With {
                .Empresas = empresas,
                .EmpresasSelecionada = empresaSelecionada,
                .Seguimentos = seguimentos,
                .SeguimentoSelecionado = seguimentoSelecionado,
                .DataInicio = DateTime.Parse(inicio),
                .DataFim = DateTime.Parse(fim)
            }

            session("ProdutoDashboardFiltro") = filtro
            validaKeyCode(seguimentoSelecionado)
            validaKeyCode(empresaSelecionada)

        Else

            Dim bAjustaDadosUsuario As Boolean = False

            filtro = DirectCast(session("ProdutoDashboardFiltro"), ProdutoDashboardFiltro)

            If request("Filtro.EmpresasSelecionadas") Is Nothing Or
               request("Filtro.SeguimentoSelecionado") Is Nothing Or
               request("dataInicio") Is Nothing Or
               request("dataFim") Is Nothing Then
                bAjustaDadosUsuario = False
            ElseIf Not filtro.EmpresasSelecionada = request("Filtro.EmpresasSelecionadas") Or
                    Not filtro.SeguimentoSelecionado = request("Filtro.SeguimentoSelecionado") Or
                    Not filtro.DataInicio = request("dataInicio") Or
                    Not filtro.DataFim = request("dataFim") Then
                bAjustaDadosUsuario = True
            End If

            If bAjustaDadosUsuario Then

                filtro.EmpresasSelecionada = request("Filtro.EmpresasSelecionadas")
                filtro.SeguimentoSelecionado = request("Filtro.SeguimentoSelecionado")
                filtro.DataInicio = DateTime.Parse(request("dataInicio"))
                filtro.DataFim = DateTime.Parse(request("dataFim"))
                session("ProdutoDashboardFiltro") = filtro
                validaKeyCode(filtro.SeguimentoSelecionado)
                validaKeyCode(filtro.EmpresasSelecionada)

                If bAjustaDadosUsuario Then
                    AtualizarDadosUsuario(filtro)
                End If

            End If

        End If

        Return filtro
    End Function

    Private Sub validaKeyCode(ByVal CodigoEmpresa As String)
        Dim objValidateKey As New ValidationKey(Left(CodigoEmpresa, 8))
        If objValidateKey.AtivoLocal Then
            UsuarioServidor.KeyCodeActive = True
        Else
            UsuarioServidor.KeyCodeActive = False
        End If
    End Sub

    Public Function ObterProdutoDashboardFiltro(session As HttpSessionStateBase) As ProdutoDashboardFiltro
        Return DirectCast(session("ProdutoDashboardFiltro"), ProdutoDashboardFiltro)
    End Function

    Public Sub AtualizarProdutoDashboard(session As HttpSessionStateBase, filtro As ProdutoDashboardFiltro)
        session("ProdutoDashboardFiltro") = filtro
    End Sub

    Public Sub AtualizarDadosUsuario(ByVal filtro As ProdutoDashboardFiltro)

        Try

            'Não pode atualizar os dados se a empresa não for selecionada
            If String.IsNullOrEmpty(filtro.EmpresasSelecionada) Then
                Exit Sub
            End If

            Dim strEmpresa As String() = filtro.EmpresasSelecionada.Split(",")
            Dim objEmpresa As [Lib].Negocio.Cliente

            If strEmpresa.Length = 0 Then
                strEmpresa = filtro.EmpresasSelecionada.Split("|")
                objEmpresa = New [Lib].Negocio.Cliente(strEmpresa(0), Convert.ToInt32(strEmpresa(1)))
            Else
                objEmpresa = New [Lib].Negocio.Cliente(UsuarioServidor.CodigoEmpresa, Convert.ToInt32(UsuarioServidor.EnderecoEmpresa))
            End If

            If objEmpresa IsNot Nothing Then

                HttpContext.Current.Session("Empresa") = objEmpresa
                HttpContext.Current.Session("ssEmpresa") = objEmpresa.Codigo
                HttpContext.Current.Session("ssEndEmpresa") = objEmpresa.CodigoEndereco
                HttpContext.Current.Session("ssNomeEmpresa") = objEmpresa.Nome
                HttpContext.Current.Session("ssEnderecoEmpresa") = objEmpresa.Endereco
                HttpContext.Current.Session("ssCidadeEmpresa") = objEmpresa.Cidade
                HttpContext.Current.Session("ssEstadoEmpresa") = objEmpresa.CodigoEstado
                HttpContext.Current.Session("ssReduzidoEmpresa") = objEmpresa.Reduzido
                HttpContext.Current.Session("ssImagemEmpresa") = objEmpresa.Imagem

                Dim Sqls As New ArrayList
                If UsuarioServidor.Usuario IsNot Nothing Then
                    Dim objUsuario As New [Lib].Negocio.Usuario(UsuarioServidor.Usuario.Usuario_Id)
                    If objUsuario IsNot Nothing Then
                        objUsuario.IUD = "U_DASHBOARD"
                        objUsuario.DataInicioDashboard = filtro.DataInicio
                        objUsuario.DataFimDashboard = filtro.DataFim
                        objUsuario.SeguimentoDashboard = filtro.SeguimentoSelecionado
                        objUsuario.EmpresasDashboard = filtro.EmpresasSelecionada
                        objUsuario.SalvarSql(Sqls)

                        Dim banco As New [Lib].Negocio.AcessaBanco()
                        banco.GravaBanco(Sqls)

                    End If
                End If

            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try

    End Sub

End Module
