Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class ucTrocaEmpresa
    Inherits BaseUserControl

    Dim sql As String = String.Empty

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            CarregarUnidadesNegocio()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        CarregarUnidadesNegocio()
        GridClientes.DataSource = New List(Of Object)
        GridClientes.DataBind()
        GridClientes.SelectedIndex = -1
        Session.Remove("_MainUserControl")
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim strEmpresa As String() = args.Split(";")
            Dim objEmpresa As New [Lib].Negocio.Cliente(strEmpresa(0), Convert.ToInt32(strEmpresa(1)))
            objEmpresa.UnidadeDeNegocio = New [Lib].Negocio.Cliente(LbxUnidades.SelectedValue, 0)
            If objEmpresa IsNot Nothing Then
                Carrega(objEmpresa)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Carrega(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            Dim objEmpresa As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
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
                    objUsuario.IUD = "U"
                    objUsuario.Empresa_Id = objEmpresa.Codigo
                    objUsuario.EnderecoEmpresa = objEmpresa.CodigoEndereco
                    objUsuario.AcessoUnidade = objEmpresa.UnidadeDeNegocio.Codigo
                    objUsuario.AcessoEmpresa = objEmpresa.Codigo
                    objUsuario.AcessoEnderecoEmpresa = objEmpresa.CodigoEndereco
                    objUsuario.SalvarSql(Sqls)
                End If
            End If

            Dim banco As New [Lib].Negocio.AcessaBanco()
            banco.GravaBanco(Sqls)

            Response.Redirect(Request.Path)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarUnidadesNegocio()
        LbxUnidades.Items.Clear()
        Dim Tipos As New List(Of [Lib].Negocio.eTipoCliente)
        Tipos.Add([Lib].Negocio.eTipoCliente.UnidadesNegocio)
        Dim objUnidadesNegocio As New [Lib].Negocio.ListCliente("", Tipos)
        If objUnidadesNegocio.Count > 0 Then
            For Each objUnidade As [Lib].Negocio.Cliente In objUnidadesNegocio
                LbxUnidades.Items.Add(New ListItem(objUnidade.Nome, objUnidade.Codigo.Trim()))
            Next
        Else : MsgBox(Me.Page, objUnidadesNegocio.Erro.Message, eTitulo.Info)
        End If
        objUnidadesNegocio = Nothing
    End Sub

    Protected Sub LbxUnidades_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles LbxUnidades.SelectedIndexChanged
        Dim objEmpresas As New [Lib].Negocio.ListCliente(LbxUnidades.SelectedValue)
        If objEmpresas.Count > 0 Then
            GridClientes.DataSource = objEmpresas.ToArray()
            GridClientes.DataBind()
            GridClientes.SelectedIndex = -1
        End If
    End Sub

    Protected Sub GridClientes_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)
        Selecionar(e.CommandArgument.ToString())
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divTrocaEmpresa")
    End Sub

End Class
