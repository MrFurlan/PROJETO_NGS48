Imports System.IO
Imports System.Data.SqlClient
Imports System.Reflection
Imports System.Web.Services
Imports System.Web.Script.Serialization
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Principal
    Inherits System.Web.UI.MasterPage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Me.Page.Header.DataBind()
            If HttpContext.Current.Session("ssEmpresa") IsNot Nothing Then SetMenuCss()
            pageLoad()
            'CarregarEmpresa()
            'HIDmaster.Value = Guid.NewGuid().ToString
            'ucTrocaEmpresa.SetarHID(HIDmaster.Value)
        End If
    End Sub

    Protected Sub lbtSair_Click(ByVal sender As Object, ByVal e As EventArgs)
        Response.Redirect("~/Logout.aspx")
    End Sub

    Private Sub pageLoad()
        Try
            Using bp As New BasePage()
                If String.IsNullOrWhiteSpace(UsuarioServidor.NomeUsuario) OrElse String.IsNullOrWhiteSpace(UsuarioServidor.Conexao) Then
                    Response.Redirect(String.Format("~/Logout.aspx?url={0}", Request.Path), False)
                    Exit Sub
                ElseIf (Not String.IsNullOrWhiteSpace(UsuarioServidor.NomeUsuario)) AndAlso (Not String.IsNullOrWhiteSpace(UsuarioServidor.Conexao)) Then
                    Dim objAssembly As Assembly = Assembly.GetExecutingAssembly()
                    Dim objName As AssemblyName = objAssembly.GetName()
                    Dim objVersao As Version = objName.Version

                    loginform.Visible = False
                    lblVersao.Text = String.Format("{0} {1}", "Versão: ", objVersao.ToString())

                    If Not IsNothing(Session("Conexao")) Then
                        loginform.Visible = True
                        Dim strCumprimento As String = String.Empty

                        Select Case DateTime.Now.Hour
                            Case 6 To 12 : strCumprimento = "Bom dia"
                            Case 12 To 18 : strCumprimento = "Boa tarde"
                            Case Else : strCumprimento = "Boa noite"
                        End Select

                        lblMensagem.Text = String.Format("{0}, {1}!", strCumprimento, Session("ssNomeUsuario"))

                        Dim vlrCompra = String.Empty
                        Dim vlrPTAX = String.Empty
                        Dim vlrVenda = String.Empty
                        'Dim sql As String = "SELECT * FROM Cotacoes WHERE Data_Id = '" & DateTime.Now.ToString("yyyy-MM-dd") & "' order by Data_Id desc"
                        Dim sql As String = "SELECT * FROM Cotacoes WHERE Data_Id = CONVERT(datetime,'" & DateTime.Now.ToString("yyyy-MM-dd") & "', 120) order by Data_Id desc"
                        Dim ds As DataSet = New [Lib].Negocio.AcessaBanco().ConsultaDataSet(sql, "Cotacoes")
                        If (ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0) Then
                            For Each row As DataRow In ds.Tables(0).Rows
                                If (CInt(row("Indexador_Id")) = 1) Then
                                    vlrCompra = String.Format("{0:N4}", CDec(row("Indice")))
                                ElseIf (CInt(row("Indexador_Id")) = 2) Then
                                    vlrPTAX = String.Format("{0:N4}", CDec(row("Indice")))
                                ElseIf (CInt(row("Indexador_Id")) = 3) Then
                                    vlrVenda = String.Format("{0:N4}", CDec(row("Indice")))
                                End If
                            Next
                        End If

                        If Session("Empresa") Is Nothing Then
                            Dim obj As Usuario = UsuarioServidor.Carregar(Session("ssNomeUsuario"))
                            Session("Empresa") = obj.Empresa
                            HIDLiberaJanela.Value = obj.LiberaJanela

                            If Not UsuarioServidor.KeyCodeActive Then validaKeyCode(obj.Empresa.Codigo)

                            If Not obj.Empresa Is Nothing Then
                                With obj.Empresa
                                    lblEmpresa.Text = String.Format("<b>{0} - {1} ({2})</b><br />{3}, {4}<br />{5} - {6}-{7}{8}<br />DÓLAR COMPRA: {9} | VENDA: {10} | PTAX: {11}",
                                                                    .Reduzido, .Nome.ToUpper(), [Lib].Negocio.Funcoes.FormatarCpfCnpj(.Codigo),
                                                                    .Endereco.ToUpper(), .Numero.ToString(),
                                                                    .CEP, .Cidade, .CodigoEstado,
                                                                    (IIf(String.IsNullOrWhiteSpace(.Complemento), String.Empty, (" - " + .Complemento.ToUpper()))),
                                                                    vlrCompra, vlrVenda, vlrPTAX)
                                End With
                            End If
                        Else
                            Dim objEmpresa As Cliente = CType(Session("Empresa"), Cliente)

                            HIDLiberaJanela.Value = UsuarioServidor.LiberaJanela

                            If Not UsuarioServidor.KeyCodeActive Then validaKeyCode(objEmpresa.Codigo)

                            With objEmpresa
                                lblEmpresa.Text = String.Format("<b>{0} - {1} ({2})</b><br />{3}, {4}<br />{5} - {6}-{7}{8}<br />DÓLAR COMPRA: {9} | VENDA: {10} | PTAX: {11}",
                                                                    .Reduzido, .Nome.ToUpper(), [Lib].Negocio.Funcoes.FormatarCpfCnpj(.Codigo),
                                                                    .Endereco.ToUpper(), .Numero.ToString(),
                                                                    .CEP, .Cidade, .CodigoEstado,
                                                                    (IIf(String.IsNullOrWhiteSpace(.Complemento), String.Empty, (" - " + .Complemento.ToUpper()))),
                                                                    vlrCompra, vlrVenda, vlrPTAX)
                            End With
                        End If

                        Dim strDivisoes As String() = Session("conexao").ToString().Split(";")

                        For Each strDivisao As String In strDivisoes
                            If strDivisao.IndexOf("Data Source") > -1 Then
                                lblHost.Text = strDivisao.Replace("Data Source", "").Replace("=", "").Trim()
                            ElseIf strDivisao.IndexOf("Initial Catalog") > -1 Then
                                lblBanco.Text = strDivisao.Replace("Initial Catalog", "").Replace("=", "").Trim()
                            End If
                        Next
                    End If
                End If
            End Using
        Catch ex As Exception
            'MsgBox(Me.Page, ex.Message)
            'Response.Redirect(String.Format("~/Logout.aspx?url={0}", Request.Path), False)
            ScriptManager.RegisterClientScriptBlock(Page, GetType(Page), Guid.NewGuid().ToString(), String.Format("alert('{0}'); window.location.href = '{1}';", "Licença Expirada, entre em contato com o Suporte.", ResolveUrl(String.Format("~/Logout.aspx?url={0}", Request.Path))), True)
        End Try
    End Sub

    Private Sub validaKeyCode(ByVal CodigoEmpresa As String)
        Dim objValidateKey As New ValidationKey(Left(CodigoEmpresa, 8))
        If objValidateKey.AtivoLocal Then
            lnkAlert.Visible = False
            UsuarioServidor.KeyCodeActive = True
        Else
            lnkAlert.Visible = True
            UsuarioServidor.KeyCodeActive = False
        End If
    End Sub

    Public Sub setMenu(ByVal modulo As [Lib].Negocio.eModulo)
        Dim pnlMenu As Panel = CType(Me.Page.FindControlRecursive("pnlMenu"), Panel)
        If pnlMenu IsNot Nothing Then
            pnlMenu.Controls.Clear()
            Select Case modulo
                Case [Lib].Negocio.eModulo.Adiantamento
                    Dim ctrl As mnuAdiantamento = CType(LoadControl("~/Menus/mnuAdiantamento.ascx"), mnuAdiantamento)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Comercial
                    Dim ctrl As mnuComercial = CType(LoadControl("~/Menus/mnuComercial.ascx"), mnuComercial)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Compras
                    Dim ctrl As mnuCompras = CType(LoadControl("~/Menus/mnuCompras.ascx"), mnuCompras)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Contabil
                    Dim ctrl As mnuContabil = CType(LoadControl("~/Menus/mnuContabil.ascx"), mnuContabil)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Custos
                    Dim ctrl As mnuCustos = CType(LoadControl("~/Menus/mnuCustos.ascx"), mnuCustos)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Expedicao
                    Dim ctrl As mnuExpedicao = CType(LoadControl("~/Menus/mnuExpedicao.ascx"), mnuExpedicao)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Financeiro
                    Dim ctrl As mnuFinanceiro = CType(LoadControl("~/Menus/mnuFinanceiro.ascx"), mnuFinanceiro)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Fiscal
                    Dim ctrl As mnuFiscal = CType(LoadControl("~/Menus/mnuFiscal.ascx"), mnuFiscal)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Gestao
                    Dim ctrl As mnuGestao = CType(LoadControl("~/Menus/mnuGestao.ascx"), mnuGestao)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Gerencial
                    Dim ctrl As mnuGerencial = CType(LoadControl("~/Menus/mnuGerencial.ascx"), mnuGerencial)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Patrimonio
                    Dim ctrl As mnuPatrimonio = CType(LoadControl("~/Menus/mnuPatrimonio.ascx"), mnuPatrimonio)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.ProducaoEstoque
                    Dim ctrl As mnuProducaoEstoque = CType(LoadControl("~/Menus/mnuProducaoEstoque.ascx"), mnuProducaoEstoque)
                    pnlMenu.Controls.Add(ctrl)
                Case [Lib].Negocio.eModulo.Revenda
                    Dim ctrl As mnuRevenda = CType(LoadControl("~/Menus/mnuRevenda.ascx"), mnuRevenda)
                    pnlMenu.Controls.Add(ctrl)
                Case Else
                    Dim ctrl As mnuGestao = CType(LoadControl("~/Menus/mnuGestao.ascx"), mnuGestao)
                    pnlMenu.Controls.Add(ctrl)
            End Select
            setBreadcrumbs(modulo)
        End If
    End Sub

    Public Sub SetMenuCss()

        Dim cssTheme As HtmlLink = CType(Me.Page.FindControlRecursive("cssTheme"), HtmlLink)

        cssTheme.Attributes.Remove("href")
        cssTheme.Attributes.Add("href", "~/Styles/styledefault.css")
        imgLogo.ImageUrl = "~/Images/logo_ngs.png"

        'Busca CSS personalizado para Empresa
        If File.Exists(Server.MapPath("~/Styles/" + Left(HttpContext.Current.Session("ssEmpresa"), 8) + ".css")) Then
            cssTheme.Attributes.Remove("href")
            cssTheme.Attributes.Add("href", "~/Styles/" + Left(HttpContext.Current.Session("ssEmpresa"), 8) + ".css")
        End If

        'Busca Logo personalizado para Empresa
        If File.Exists(Server.MapPath("~/Images/" + Left(HttpContext.Current.Session("ssEmpresa"), 8) + ".png")) Then
            imgLogo.ImageUrl = "~/Images/" + Left(HttpContext.Current.Session("ssEmpresa"), 8) + ".png"
        End If
    End Sub

    Public Sub setBreadcrumbs(ByVal modulo As [Lib].Negocio.eModulo)
        Dim pageName As String = Path.GetFileName(Request.Path).ToUpper().Trim()
        Dim obj As Breadcrumbs = Serializador.Deserializar()
        Dim lst As New List(Of BreadcrumbsNode)
        If obj IsNot Nothing AndAlso obj.Nodes IsNot Nothing AndAlso obj.Nodes.Count > 0 Then
            getRoot(pageName, obj.Nodes(0))
            getNodes(pageName, obj.Nodes, lst)
        End If
        lst = lst.OrderByDescending(Function(s) s.Index).ToList()
        rpt.DataSource = lst
        rpt.DataBind()
    End Sub

    Private Sub getNodes(ByVal pageName As String, ByVal lstNodes As List(Of BreadcrumbsNode), ByRef lst As List(Of BreadcrumbsNode))
        Dim index As Integer = 1
        Dim obj As BreadcrumbsNode = getNodeByName(pageName, lstNodes)
        If obj IsNot Nothing Then
            obj.Index = index
            lst.Add(obj)
            index += 1
            If obj.Root IsNot Nothing Then
                obj.Root.Index = index
                lst.Add(obj.Root)
                index += 1
                getIndex(obj.Root, lst, index)
            End If
        End If
    End Sub

    Private Sub getIndex(ByVal obj As BreadcrumbsNode, ByRef lst As List(Of BreadcrumbsNode), ByRef index As Integer)
        If obj.Root IsNot Nothing Then
            obj.Root.Index = index
            lst.Add(obj.Root)
            index += 1
            getIndex(obj.Root, lst, index)
        End If
    End Sub

    Private Function getNodeByName(ByVal pageName As String, ByVal lstNodes As List(Of BreadcrumbsNode)) As BreadcrumbsNode
        For Each node As BreadcrumbsNode In lstNodes
            If node IsNot Nothing AndAlso Path.GetFileName(node.Url).ToUpper().Trim().Equals(pageName) Then
                Return node
            ElseIf node.Nodes IsNot Nothing AndAlso node.Nodes.Count > 0 AndAlso node.Nodes.Any(Function(s) Path.GetFileName(s.Url).ToUpper().Trim().Equals(pageName)) Then
                Return node.Nodes.Where(Function(s) Path.GetFileName(s.Url).ToUpper().Equals(pageName)).FirstOrDefault()
            ElseIf node.Nodes IsNot Nothing AndAlso node.Nodes.Count > 0 AndAlso node.Nodes.Any(Function(s) s.Nodes IsNot Nothing AndAlso s.Nodes.Count > 0) Then
                For Each no As BreadcrumbsNode In node.Nodes
                    If no.Nodes IsNot Nothing AndAlso no.Nodes.Count > 0 Then
                        Dim obj As BreadcrumbsNode = getNodeByName(pageName, no.Nodes)
                        If obj Is Nothing Then
                            Continue For
                        End If
                        Return obj
                    End If
                Next
            End If
        Next
        Return Nothing
    End Function

    Private Sub getRoot(ByVal pageName As String, ByVal node As BreadcrumbsNode)
        If node IsNot Nothing AndAlso node.Nodes IsNot Nothing AndAlso node.Nodes.Count > 0 Then
            For Each n As BreadcrumbsNode In node.Nodes
                n.Root = node
                If n IsNot Nothing AndAlso n.Nodes IsNot Nothing AndAlso n.Nodes.Count > 0 Then
                    getRoot(pageName, n)
                End If
            Next
        End If
    End Sub

    Protected Sub lnkTrocaEmpresa_Click(sender As Object, e As EventArgs) Handles lnkTrocaEmpresa.Click
        Try
            If UsuarioServidor.TrocaEmpresa Then
                Dim Popup As New CarregarPopup
                ucTrocaEmpresa.Limpar()
                Popup.TrocaDeEmpresa(Me.Page, "objSelecaoDeEmpresas" & HIDmaster.Value)
            Else
                ScriptManager.RegisterStartupScript(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "TrocaBloqueada();", True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class