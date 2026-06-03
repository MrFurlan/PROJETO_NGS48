Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Usuarios
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("Usuarios", "ACESSAR") Then
                CarregarUsuarios()
                CarregarBancoDeDados()
                CargaUnidadeDeNegocio()
                Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                CarregarMenuDeAcesso()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUsuarios()
        Sql = " SELECT Usuarios.Usuario_Id AS Usuario, Usuarios.NomeCompleto, upper(BU.NomeDoBanco) as BancoDeDados, upper(ISNULL(Usuarios.MenuDeAcesso, '')) AS MenuDeAcesso, " & vbCrLf &
              "        ISNULL(Usuarios.Email, 'nao informado') AS Email, AcessoUnidade, AcessoEmpresa, AcessoEndEmpresa, isnull(Ativo,0) as Ativo, ISNULL(Usuarios.Cliente, '') AS Cliente, " & vbCrLf &
              "        ISNULL(Usuarios.EndCliente, 0) as EndCliente, Cl.Nome AS NomeCliente, isnull(ImprimirDanfe,0) as ImprimirDanfe, ISNULL(LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
              "        ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, ISNULL(LiberaJanela, 0) as LiberaJanela" & vbCrLf &
              "   FROM   Usuarios " & vbCrLf &
              "  INNER JOIN Usuarios.dbo.Usuarios BU " & vbCrLf &
              "     ON Usuarios.Usuario_Id COLLATE DATABASE_DEFAULT = BU.Usuario_Id COLLATE DATABASE_DEFAULT" & vbCrLf &
              "   LEFT JOIN Clientes Cl" & vbCrLf &
              "     ON Usuarios.Cliente = Cl.Cliente_Id" & vbCrLf &
              "    AND Usuarios.EndCliente = Cl.Endereco_Id" & vbCrLf

        gridUsuarios.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        gridUsuarios.DataBind()
    End Sub

    Private Sub CarregarBancoDeDados()
        ddlBancoDeDados.Items.Clear()
        Sql = "Select Banco_Id as Banco From Usuarios.dbo.Bancos Order By Banco_Id"
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlBancoDeDados.Items.Add(New ListItem(UCase(Dr("Banco")), UCase(Dr("Banco"))))
        Next
        ddlBancoDeDados.Items.Insert(0, "")
        ddlBancoDeDados.SelectedIndex = 0
    End Sub

    Private Sub CarregarMenuDeAcesso()
        Sql = "Select Menu_Id as Menu From MenuDeAcesso Order By Menu_Id"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlMenuDeAcesso.Items.Add(New ListItem(UCase(Dr("Menu")), UCase(Dr("Menu"))))
        Next
        ddlMenuDeAcesso.Items.Insert(0, "")
        ddlMenuDeAcesso.SelectedIndex = 0
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Protected Sub btnConsultarCliente_Click(sender As Object, e As EventArgs) Handles btnConsultarCliente.Click
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me, "objClienteUsuario" & HID.Value.ToString, "txtNome")
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteUsuario" & HID.Value) Is Nothing Then
            Dim cli As Cliente = Session("objClienteUsuario" & HID.Value)
            txtCodigoCliente.Value = cli.Codigo
            'txtEnderecoCliente.Value = cli.CodigoEndereco

            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value

            Session.Remove("objClienteUsuario" & HID.Value)
        End If

    End Sub

    Private Sub Limpar()
        txtUsuario.Text = ""
        chkAtivo.Checked = False
        ChkImprimirDanfe.Checked = False
        chkLiberaEmpresa.Checked = False
        chkTrocaEmpresa.Checked = False
        chkLiberaJanela.Checked = False

        If Session("ssNomeUsuario") = "FURLAN" OrElse Session("ssNomeUsuario") = "FELIPE" OrElse Session("ssNomeUsuario") = "KAYNÃ" Then
            chkLiberaJanela.Enabled = True
        Else
            chkLiberaJanela.Enabled = False
        End If

        txtNomeCompleto.Text = ""
        ddlBancoDeDados.SelectedIndex = 0
        ddlMenuDeAcesso.SelectedIndex = 0
        ddlUnidadeDeNegocio.SelectedIndex = 0
        ddlEmpresa.Items.Clear()
        txtEmail.Text = ""
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtUsuario.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        'lnkExcluir.Parent.Visible = False
    End Sub

    Private Function Validar_Campos() As Boolean
        If txtUsuario.Text.Length = 0 Then
            Mensagem = "Usuário não foi informado"
            Return False
        ElseIf txtNomeCompleto.Text.Length = 0 Then
            Mensagem = "Nome Completo não foi informado"
            Return False
        ElseIf ddlBancoDeDados.SelectedIndex = 0 Then
            Mensagem = "Bando de Dados não foi selecionado"
            Return False
        ElseIf ddlUnidadeDeNegocio.SelectedIndex = 0 Then
            Mensagem = "Unidade de Negócio não foi selecionada"
            Return False
        ElseIf ddlEmpresa.SelectedIndex = 0 Then
            Mensagem = "Empresa não foi selecionada"
            Return False
        ElseIf ddlMenuDeAcesso.SelectedIndex = 0 Then
            Mensagem = "Menu de Acesso não foi selecionado"
            Return False
        Else
            Return True
        End If
    End Function

    Private Function getParam() As String
        Dim Texto As String = ""
        If Not String.IsNullOrWhiteSpace(txtUsuario.Text) Then
            Texto &= "Usuários: " & txtUsuario.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtNomeCompleto.Text) Then
            Texto &= "Nome Completo: " & txtNomeCompleto.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlBancoDeDados.SelectedValue) Then
            Texto &= "Banco de Dados: " & ddlBancoDeDados.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlUnidadeDeNegocio.SelectedValue) Then
            Texto &= "Unidade de Negócios: " & ddlUnidadeDeNegocio.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Texto &= "Empresa: " & ddlEmpresa.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlMenuDeAcesso.SelectedValue) Then
            Texto &= "Menu de Acesso: " & ddlMenuDeAcesso.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtEmail.Text) Then
            Texto &= "Email: " & txtEmail.Text & vbCrLf
        End If

        Return Texto
    End Function

    Protected Sub gridUsuarios_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridUsuarios.SelectedIndexChanged
        Try
            If Funcoes.VerificaPermissao("Usuarios", "LEITURA") Then
                Limpar()
                txtUsuario.Text = gridUsuarios.SelectedRow.Cells(1).Text()
                txtNomeCompleto.Text = gridUsuarios.SelectedRow.Cells(2).Text()
                ddlBancoDeDados.SelectedValue = gridUsuarios.SelectedRow.Cells(3).Text()
                ddlMenuDeAcesso.SelectedValue = gridUsuarios.SelectedRow.Cells(4).Text()
                txtEmail.Text = Server.HtmlDecode(gridUsuarios.SelectedRow.Cells(5).Text())
                ddlUnidadeDeNegocio.SelectedValue = gridUsuarios.SelectedRow.Cells(6).Text()
                CargaEmpresa()
                ddlEmpresa.SelectedValue = gridUsuarios.SelectedRow.Cells(7).Text() & "-" & gridUsuarios.SelectedRow.Cells(8).Text()
                chkAtivo.Enabled = True
                ChkImprimirDanfe.Enabled = True
                chkAtivo.Checked = CType(gridUsuarios.SelectedRow.FindControl("chkAtivogrid"), CheckBox).Checked
                ChkImprimirDanfe.Checked = CType(gridUsuarios.SelectedRow.FindControl("gridchkImprimir"), CheckBox).Checked
                chkLiberaEmpresa.Checked = CType(gridUsuarios.SelectedRow.FindControl("gridchkLiberaEmpresa"), CheckBox).Checked
                chkTrocaEmpresa.Checked = CType(gridUsuarios.SelectedRow.FindControl("gridchkTrocaEmpresa"), CheckBox).Checked
                chkLiberaJanela.Checked = CType(gridUsuarios.SelectedRow.FindControl("gridchkLiberaJanela"), CheckBox).Checked

                Dim objCliente As New [Lib].Negocio.Cliente(gridUsuarios.SelectedRow.Cells(9).Text(), Convert.ToInt32(gridUsuarios.SelectedRow.Cells(10).Text()))

                If Not String.IsNullOrWhiteSpace(objCliente.Codigo) Then
                    Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objCliente)
                    txtCliente.Text = itemCliente.Text
                    txtCodigoCliente.Value = itemCliente.Value
                End If

                txtNomeCompleto.Focus()
                txtUsuario.Enabled = False
                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True
                'lnkExcluir.Parent.Visible = True
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresa()
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("Usuarios", "GRAVAR") Then
            If Validar_Campos() Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                Sql = "INSERT Into Usuarios(Usuario_id, NomeCompleto, Empresa, EndEmpresa, MenuDeAcesso, Email, AcessoUnidade, AcessoEmpresa, AcessoEndEmpresa, " & vbCrLf &
                      "Ativo, Cliente, EndCliente, ImprimirDanfe, LiberaEmpresa, TrocaEmpresa, LiberaJanela) " & vbCrLf &
                " Values('" & UCase(txtUsuario.Text) & "' " & vbCrLf &
                ",'" & UCase(txtNomeCompleto.Text) & "'" & vbCrLf &
                ",'" & Empresa(0) & "'" & vbCrLf &
                "," & Empresa(1) & vbCrLf &
                ",'" & UCase(ddlMenuDeAcesso.SelectedValue) & "'" & vbCrLf &
                ",'" & txtEmail.Text & "'" & vbCrLf &
                ",'" & ddlUnidadeDeNegocio.SelectedValue & "'" & vbCrLf &
                ",'" & Empresa(0) & "'" & vbCrLf &
                "," & Empresa(1) & vbCrLf &
                "," & IIf(chkAtivo.Checked, "1", "0") & vbCrLf

                If String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                    Sql &= ", null, null"
                Else
                    Sql &= ",'" & txtCodigoCliente.Value.Split("-")(0) & "', " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
                End If

                Sql &= "," & IIf(ChkImprimirDanfe.Checked, "1", "0") & vbCrLf &
                       "," & IIf(chkLiberaEmpresa.Checked, "1", "0") & vbCrLf &
                       "," & IIf(chkTrocaEmpresa.Checked, "1", "0") & vbCrLf &
                       "," & IIf(chkLiberaJanela.Checked, "1", "0") & ")"

                SqlArray.Add(Sql)

                Sql = "if exists (select * from Usuarios.dbo.Usuarios where Usuario_Id = '" & UCase(txtUsuario.Text) & "')" '" and HostDoServidor = '" & HttpContext.Current.Session("ssNomeServidor") & "')"
                Sql &= " update Usuarios.dbo.Usuarios set NomeDoBanco = '" & UCase(ddlBancoDeDados.SelectedValue) & "' where Usuario_id = '" & UCase(txtUsuario.Text) & "'"
                Sql &= " else Insert into Usuarios.dbo.Usuarios(Usuario_id, NomeDoBanco, HostDoServidor) values('" & UCase(txtUsuario.Text) & "', '" & UCase(ddlBancoDeDados.SelectedValue) & "', '" & HttpContext.Current.Session("ssNomeServidor") & "')"

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarUsuarios()
                Else
                    MsgBox(Me.Page, "Erro ao Incluir: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Else : MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        If Funcoes.VerificaPermissao("Usuarios", "ALTERAR") Then
            If Validar_Campos() Then
                Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                Sql = "UPDATE Usuarios" & vbCrLf &
                      " Set NomeCompleto     ='" & txtNomeCompleto.Text & "' " & vbCrLf &
                      ",    MenuDeAcesso     ='" & UCase(ddlMenuDeAcesso.SelectedValue) & "'" & vbCrLf &
                      ",    Empresa          ='" & Empresa(0) & "'" & vbCrLf &
                      ",    EndEmpresa       = " & Empresa(1) & vbCrLf &
                      ",    Email            ='" & txtEmail.Text & "' " & vbCrLf &
                      ",    AcessoUnidade    ='" & ddlUnidadeDeNegocio.SelectedValue & "' " & vbCrLf &
                      ",    AcessoEmpresa    ='" & Empresa(0) & "' " & vbCrLf &
                      ",    AcessoEndEmpresa = " & Empresa(1) & " " & vbCrLf &
                      ",    Ativo            = " & IIf(chkAtivo.Checked, "1", "0") & vbCrLf &
                      ",    ImprimirDanfe    = " & IIf(ChkImprimirDanfe.Checked, "1", "0") & vbCrLf &
                      ",    LiberaEmpresa    = " & IIf(chkLiberaEmpresa.Checked, "1", "0") & vbCrLf &
                      ",    TrocaEmpresa     = " & IIf(chkTrocaEmpresa.Checked, "1", "0") & vbCrLf &
                      ",    LiberaJanela     = " & IIf(chkLiberaJanela.Checked, "1", "0") & vbCrLf
                If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                    Sql &= ",  Cliente          = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                           ",  EndCliente       = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf

                Else
                    Sql &= ",  Cliente          =  NULL " & vbCrLf & _
                           ",  EndCliente       =  NULL " & vbCrLf
                End If

                Sql &= " WHERE Usuario_Id = '" & txtUsuario.Text & "'"

                SqlArray.Add(Sql)

                Sql = "if exists (select * from Usuarios.dbo.Usuarios where Usuario_Id = '" & UCase(txtUsuario.Text) & "')" '" and HostDoServidor = '" & HttpContext.Current.Session("ssNomeServidor") & "')"
                Sql &= " update Usuarios.dbo.Usuarios set NomeDoBanco = '" & UCase(ddlBancoDeDados.SelectedValue) & "' where Usuario_id = '" & UCase(txtUsuario.Text) & "'"
                Sql &= " else Insert into Usuarios.dbo.Usuarios(Usuario_id, NomeDoBanco, HostDoServidor) values('" & UCase(txtUsuario.Text) & "', '" & UCase(ddlBancoDeDados.SelectedValue) & "', '" & HttpContext.Current.Session("ssNomeServidor") & "')"


                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    CarregarUsuarios()
                Else
                    MsgBox(Me.Page, "Erro ao Alterar: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                End If
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Else : MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
        End If
    End Sub

    'Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
    '    If Funcoes.VerificaPermissao("Usuarios", "EXCLUIR") Then
    '        Sql = "DELETE FROM Usuarios"
    '        Sql &= " WHERE Usuario_Id = '" & txtUsuario.Text & "' "
    '        SqlArray.Add(Sql)

    '        If Banco.GravaBanco(SqlArray) Then
    '            Limpar()
    '            CarregarUsuarios()
    '        Else
    '            MsgBox(Me.Page, "Erro ao Excluir: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
    '        End If
    '    Else : MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
    '    End If
    'End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Usuarios", "RELATORIO") Then
                Dim ds As New DataSet
                Sql = " SELECT Usuarios.Usuario_Id AS Usuario, Usuarios.NomeCompleto, upper(BU.NomeDoBanco) as BancoDeDados, upper(ISNULL(Usuarios.MenuDeAcesso, '')) AS MenuDeAcesso, ISNULL(Usuarios.Email, " & vbCrLf &
                      " 'nao informado') AS Email" & vbCrLf &
                      " FROM   Usuarios INNER JOIN Usuarios.dbo.Usuarios BU " & vbCrLf &
                      "        ON Usuarios.Usuario_Id COLLATE DATABASE_DEFAULT = BU.Usuario_Id COLLATE DATABASE_DEFAULT" & vbCrLf
                If ddlEmpresa.SelectedValue.ToString.Length > 0 Then
                    Dim objEmpresa = ddlEmpresa.SelectedValue.Split("-")
                    Sql += "        Where Empresa = '" & objEmpresa(0) & "'" & vbCrLf &
                        "        And EndEmpresa = '" & objEmpresa(1) & "'" & vbCrLf &
                        "        And AcessoUnidade = '" & ddlUnidadeDeNegocio.SelectedValue & "'" & vbCrLf
                End If

                ds = Banco.ConsultaDataSet(Sql, "Usuarios")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("ConsultaParam", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_Usuarios", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Usuarios")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class