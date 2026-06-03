Imports NGS.Lib.Negocio

Public Class UsuariosXProcessos
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then

                If Funcoes.VerificaPermissao("LiberarPermissoes", "ACESSAR") Then
                    CarregarUsuarios()
                    CarregaCopia()
                    CarregarGrupos()
                    LimparCampos(True)
                    CarregarProcessos()
                    LimparCamposGrupo()
                    CarregarProcessosGrupo()
                    HID.Value = Guid.NewGuid().ToString
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal str As String)
        If Session("Processos_Grupo" & HID.Value.ToString()) IsNot Nothing Then
            CarregarProcessos()
            CarregarProcessosGrupo()
            ddlProcessos.SelectedValue = Session("Processos_Grupo" & HID.Value.ToString())
            Session.Remove("Processos_Grupo" & HID.Value.ToString())
        ElseIf Session("Processos_Usuario" & HID.Value.ToString()) IsNot Nothing Then
            CarregarProcessos()
            CarregarProcessosGrupo()
            ddlProcessoGrupo.SelectedValue = Session("Processos_Usuario" & HID.Value.ToString())
            Session.Remove("Processos_Usuario" & HID.Value.ToString())
        ElseIf Session("Grupos" & HID.Value.ToString()) IsNot Nothing Then
            CarregarGrupos()
            Session.Remove("Grupos" & HID.Value.ToString())
        End If
    End Sub

#Region "Por Usuário"

#Region "Events"

    Protected Sub grdUsuarios_SelectedIndexChanged(sender As Object, e As EventArgs) Handles grdUsuarios.SelectedIndexChanged
        Try
            ddlUsuarios.SelectedValue = String.Empty
            ddlUsuarios.SelectedValue = Server.HtmlDecode(grdUsuarios.SelectedRow.Cells(1).Text())
            ddlProcessos.SelectedValue = Server.HtmlDecode(grdUsuarios.SelectedRow.Cells(2).Text())
            CarregaPermissoes()
            lnkExcluir.Parent.Visible = True
            lnkNovo.Parent.Visible = False
            ddlUsuarios.Enabled = False
            ddlProcessos.Enabled = False
            lnkAddProcesso.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUsuarios_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUsuarios.SelectedIndexChanged
        Try
            CarregarUsuariosXProcessos()
            getGruposDoUsuario()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "GRAVAR") Then
                If ValidaCampos() Then
                    Dim Sql = "INSERT Into UsuariosXProcessos (Usuario_Id, Processo_id) " & vbCrLf & _
                      "             Values('" & UCase(ddlUsuarios.SelectedValue) & "', '" & UCase(ddlProcessos.SelectedValue) & "')" & vbCrLf

                    If Banco.GravaBanco(Sql) Then
                        MsgBox(Me.Page, "Processo incluído ao usuário.")
                        LimparCampos(False)
                        CarregarUsuariosXProcessos()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "EXCLUIR") Then
                If grdUsuarios.SelectedIndex <> -1 Then
                    If ValidaCampos() Then
                        Dim sql As String = "DELETE FROM UsuariosXProcessosxPermissoes" & vbCrLf & _
                                            "        WHERE Usuario_Id = '" & UCase(ddlUsuarios.SelectedValue) & "' And Processo_ID = '" & UCase(ddlProcessos.SelectedValue) & "'" & vbCrLf & _
                                            "DELETE FROM UsuariosXProcessos" & vbCrLf & _
                                            "        WHERE Usuario_Id = '" & UCase(ddlUsuarios.SelectedValue) & "' And Processo_ID = '" & UCase(ddlProcessos.SelectedValue) & "'" & vbCrLf

                        If Banco.GravaBanco(sql) Then
                            MsgBox(Me.Page, "Processo excluído com Sucesso.", eTitulo.Sucess)
                            LimparCampos(False)
                            CarregarUsuariosXProcessos()
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Selecione um registro, para exclusão.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "RELATORIO") Then
                Dim sql As String = ""

                sql = " select u.NomeCompleto as Usuario, up.Processo_Id as Processo, upp.Permissao_Id as Permissao from UsuariosxProcessos up " & vbCrLf & _
                      " Inner Join Usuarios as u " & vbCrLf & _
                      " on u.Usuario_Id = up.Usuario_Id" & vbCrLf & _
                      " Inner Join UsuariosXProcessosXPermissoes upp " & vbCrLf & _
                      " On upp.Usuario_Id = up.Usuario_Id" & vbCrLf & _
                      " and upp.Processo_Id = up.Processo_Id" & vbCrLf

                Dim DS As DataSet = Banco.ConsultaDataSet(sql, "UsuariosXProcessos")

                Funcoes.BindReport(Me.Page, DS, "Cr_UsuariosXProcessos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPermitido_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chk.NamingContainer, GridViewRow)
            Dim sql As String = ""

            If chk.Checked Then
                sql = "Insert into UsuariosXProcessosXPermissoes (Usuario_Id, Processo_Id, Permissao_Id) values ('" & ddlUsuarios.SelectedValue & "', '" & ddlProcessos.SelectedValue & "', '" & grdPermissoes.Rows(row.RowIndex).Cells(1).Text & "')"
            Else
                sql = "Delete UsuariosXProcessosXPermissoes where Usuario_Id = '" & ddlUsuarios.SelectedValue & _
                    "' and Processo_Id = '" & ddlProcessos.SelectedValue & "' and Permissao_Id = '" & grdPermissoes.Rows(row.RowIndex).Cells(1).Text & "'"
            End If
            Banco.GravaBanco(sql)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAddProcesso_Click(sender As Object, e As EventArgs) Handles lnkAddProcesso.Click
        Try
            ucCadastrarProcesso.Limpar()
            ucCadastrarProcesso.SetarHID(HID.Value)
            Popup.CadastroDeProcessos(Me.Page, "Processos_Usuario" & HID.Value.ToString, "txtProcesso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregarUsuarios()
        Dim Sql = "Select Usuario_Id as Usuario From Usuarios Order By Usuario_Id"

        ddlUsuarios.Items.Clear()

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUsuarios.Items.Add(New ListItem(UCase(Dr("Usuario")), UCase(Dr("Usuario"))))
        Next

        ddlUsuarios.Items.Insert(0, "")
        ddlUsuarios.SelectedIndex = 0
    End Sub

    Private Sub CarregarProcessos()
        ddl.Carregar(ddlProcessos, CarregarDDL.Tabela.Processos, "", True)
    End Sub

    Private Sub CarregarUsuariosXProcessos()
        Dim sql = "Select upper(Usuario_Id) as Usuarios, upper(Processo_Id) as Processos From UsuariosXProcessos " & vbCrLf & _
                  " Where Usuario_Id = '" & UCase(ddlUsuarios.SelectedValue) & "' Order By Processo_Id"

        grdUsuarios.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        grdUsuarios.DataBind()
    End Sub

    Private Sub CarregaPermissoes()
        Dim sql = "   Select case when upp.Permissao_Id IS NULL then 0 else 1 end as Permitido, p.Permissao_Id from Permissoes p    " & vbCrLf & _
                "   	Left JOin UsuariosXProcessosXPermissoes upp                                                               " & vbCrLf & _
                "   		on upp.Permissao_Id = p.Permissao_Id                                                                  " & vbCrLf & _
                "   	    And (upp.Processo_Id is null or upp.Processo_Id = '" & ddlProcessos.SelectedValue & "')                 " & vbCrLf & _
                "   	    And (upp.Usuario_Id is null or upp.Usuario_Id = '" & ddlUsuarios.SelectedValue & "')                      " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Permissao")
        grdPermissoes.DataSource = ds
        grdPermissoes.DataBind()
    End Sub

    Private Sub LimparCampos(ByVal limpartudo As Boolean)
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
        ddlUsuarios.Enabled = True
        ddlProcessos.Enabled = True
        ddlProcessos.SelectedIndex = 0
        lnkAddProcesso.Enabled = True

        If limpartudo Then
            ddlUsuarios.SelectedIndex = 0
            grdUsuarios.DataSource = Nothing
            grdUsuarios.DataBind()
            lblGruposDoUsuario.Text = String.Empty
        End If

        grdPermissoes.DataSource = Nothing
        grdPermissoes.DataBind()

    End Sub

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlUsuarios.SelectedValue) Then
            MsgBox(Me.Page, "Informe o Usuário.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlProcessos.SelectedValue) Then
            MsgBox(Me.Page, "Informe o Processo.")
            Return False
        End If
        Return True
    End Function

#End Region

#End Region

#Region "Por Grupo"

#Region "Events"

    Protected Sub ddlGrupo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupo.SelectedIndexChanged
        Try
            CarregarGruposXProcessos()
            getUsuariosDoGrupo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grdProcessoGrupo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles grdProcessoGrupo.SelectedIndexChanged
        Try
            ddlGrupo.SelectedValue = String.Empty
            ddlGrupo.SelectedValue = Server.HtmlDecode(grdProcessoGrupo.SelectedRow.Cells(1).Text())
            ddlProcessoGrupo.SelectedValue = Server.HtmlDecode(grdProcessoGrupo.SelectedRow.Cells(2).Text())
            CarregaPermissoesGrupo()
            lnkExcluirGrupo.Parent.Visible = True
            lnkNovoGrupo.Parent.Visible = False
            ddlGrupo.Enabled = False
            ddlProcessoGrupo.Enabled = False
            lnkAddProcessoGrupo.Enabled = False
            lnkAddGrupo.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPermissaoGrupo_CheckedChanged(sender As Object, e As EventArgs)
        Try
            Dim chk As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chk.NamingContainer, GridViewRow)
            Dim sql As String = ""

            If chk.Checked Then
                sql = "Insert into GruposXProcessosXPermissoes " & vbCrLf & _
                      " (Grupo_Id, Processo_Id, Permissao_Id) values ('" & ddlGrupo.SelectedValue & "', '" & ddlProcessoGrupo.SelectedValue & "', '" & grdPermissoesGrupo.Rows(row.RowIndex).Cells(1).Text & "')"
            Else
                sql = "Delete GruposXProcessosXPermissoes where Grupo_Id = '" & ddlGrupo.SelectedValue & _
                    "' and Processo_Id = '" & ddlProcessoGrupo.SelectedValue & "' and Permissao_Id = '" & grdPermissoesGrupo.Rows(row.RowIndex).Cells(1).Text & "'"
            End If
            Banco.GravaBanco(sql)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovoGrupo_Click(sender As Object, e As EventArgs) Handles lnkNovoGrupo.Click
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "GRAVAR") Then
                If ValidaCamposGrupo() Then
                    Dim Sql = "INSERT Into GruposXProcessos (Grupo_Id, Processo_id) " & vbCrLf & _
                      "             Values('" & UCase(ddlGrupo.SelectedValue) & "', '" & UCase(ddlProcessoGrupo.SelectedValue) & "')" & vbCrLf

                    If Banco.GravaBanco(Sql) Then
                        MsgBox(Me.Page, "Processo incluído ao Grupo.")
                        LimparCamposGrupo()
                        CarregarGruposXProcessos()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluirGrupo_Click(sender As Object, e As EventArgs) Handles lnkExcluirGrupo.Click
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "EXCLUIR") Then
                If grdProcessoGrupo.SelectedIndex <> -1 Then
                    If ValidaCamposGrupo() Then
                        Dim sql As String = "DELETE FROM GruposXProcessosxPermissoes" & vbCrLf & _
                                            "        WHERE Grupo_Id = '" & UCase(ddlGrupo.SelectedValue) & "' And Processo_ID = '" & UCase(ddlProcessoGrupo.SelectedValue) & "'" & vbCrLf & _
                                            "DELETE FROM GruposXProcessos" & vbCrLf & _
                                            "        WHERE Grupo_Id = '" & UCase(ddlGrupo.SelectedValue) & "' And Processo_ID = '" & UCase(ddlProcessoGrupo.SelectedValue) & "'" & vbCrLf

                        If Banco.GravaBanco(sql) Then
                            MsgBox(Me.Page, "Permissoes excluída com Sucesso.", eTitulo.Sucess)
                            LimparCamposGrupo()
                            CarregarGruposXProcessos()
                        End If
                    End If
                Else
                    MsgBox(Me.Page, "Selecione um registro, para exclusão.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimparGrupo_Click(sender As Object, e As EventArgs) Handles lnkLimparGrupo.Click
        Try
            LimparCamposGrupo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorioGrupo_Click(sender As Object, e As EventArgs) Handles lnkRelatorioGrupo.Click
        Try
            Dim sql As String = "select	g.grupo_id + ' - (' + g.Descricao + ')' as Grupo,                                                           " & vbCrLf & _
                                "		isnull(dbo.UsuariosDoGrupo(g.Grupo_id), 'NADA CONSTA') as Usuario,  " & vbCrLf & _
                                "		ISNULL(dbo.ProcessosDoGrupo(g.Grupo_Id), 'NADA CONSTA') as Processo, isnull(grup.contProcesso,0) as contProcesso, " & vbCrLf & _
                                "       isnull(gruu.contUsuario, 0) as contUsuario" & vbCrLf & _
                                "	from Grupos	g                                                            " & vbCrLf & _
                                "    left join (select Grupo_Id, COUNT(Processo_Id) as contProcesso from GruposXProcessos  " & vbCrLf & _
                                "		group by Grupo_Id) as grup                                                         " & vbCrLf & _
                                "		on g.Grupo_Id = grup.Grupo_Id                                                      " & vbCrLf & _
                                "	left join (select grupo_id, COUNT(Usuario_Id) as contUsuario from GruposXUsuarios gu   " & vbCrLf & _
                                "				group by Grupo_Id) as gruu                                             " & vbCrLf & _
                                "		on g.Grupo_Id = gruu.Grupo_Id                                                      " & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then sql &= " Where g.grupo_id = '" & ddlGrupo.SelectedValue & "'" & vbCrLf

            sql &= "	order by g.Grupo_Id" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "GruposXProcessosXUsuarios")

            Funcoes.BindReport(Me.Page, ds, "Cr_GruposXProcessosXUsuarios", eExportType.PDF)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAddGrupo_Click(sender As Object, e As EventArgs) Handles lnkAddGrupo.Click
        Try
            ucCadastrarGrupo.Limpar()
            ucCadastrarGrupo.SetarHID(HID.Value)
            Popup.CadastroDeGrupos(Me.Page, "Grupos" & HID.Value.ToString(), "txtGrupo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAddProcessoGrupo_Click(sender As Object, e As EventArgs) Handles lnkAddProcessoGrupo.Click
        Try
            ucCadastrarProcesso.Limpar()
            ucCadastrarProcesso.SetarHID(HID.Value)
            Popup.CadastroDeProcessos(Me.Page, "Processos_Grupo" & HID.Value.ToString(), "txtProcesso")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregarGrupos()
        Dim Sql = "Select Grupo_Id from Grupos order by Grupo_Id"

        ddlGrupo.Items.Clear()

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlGrupo.Items.Add(New ListItem(UCase(Dr("Grupo_Id")), UCase(Dr("Grupo_Id"))))
        Next

        ddlGrupo.Items.Insert(0, "")
        ddlGrupo.SelectedIndex = 0
    End Sub

    Private Sub CarregarProcessosGrupo()
        ddl.Carregar(ddlProcessoGrupo, CarregarDDL.Tabela.Processos, "", True)
    End Sub

    Private Sub CarregarGruposXProcessos()

        Dim sql = "Select upper(Grupo_Id) as Grupo, upper(Processo_Id) as Processo From GruposXProcessos " & vbCrLf & _
                  " Where Grupo_Id = '" & UCase(ddlGrupo.SelectedValue) & "' Order By Processo_Id"

        grdProcessoGrupo.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        grdProcessoGrupo.DataBind()
    End Sub

    Private Sub CarregaPermissoesGrupo()
        Dim sql = "   Select case when gpp.Permissao_Id IS NULL then 0 else 1 end as Permitido, p.Permissao_Id from Permissoes p    " & vbCrLf & _
                "   	Left JOin GruposXProcessosXPermissoes gpp                                                               " & vbCrLf & _
                "   		on gpp.Permissao_Id = p.Permissao_Id                                                                  " & vbCrLf & _
                "   	    And (gpp.Processo_Id is null or gpp.Processo_Id = '" & ddlProcessoGrupo.SelectedValue & "')                 " & vbCrLf & _
                "   	    And (gpp.Grupo_id is null or gpp.Grupo_id = '" & ddlGrupo.SelectedValue & "')                      " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Permissao")
        grdPermissoesGrupo.DataSource = ds
        grdPermissoesGrupo.DataBind()
    End Sub

    Private Sub LimparCamposGrupo()
        lnkNovoGrupo.Parent.Visible = True
        lnkExcluirGrupo.Parent.Visible = False
        ddlGrupo.Enabled = True
        ddlGrupo.SelectedIndex = 0
        ddlProcessoGrupo.Enabled = True
        ddlProcessoGrupo.SelectedIndex = 0
        grdProcessoGrupo.DataSource = Nothing
        grdProcessoGrupo.DataBind()
        grdPermissoesGrupo.DataSource = Nothing
        grdPermissoesGrupo.DataBind()
        lnkAddProcessoGrupo.Enabled = True
        lnkAddGrupo.Enabled = True
        lblUsuariosDoGrupo.Text = String.Empty
    End Sub

    Private Function ValidaCamposGrupo() As Boolean
        If String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
            MsgBox(Me.Page, "Informe o Grupo.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlProcessoGrupo.SelectedValue) Then
            MsgBox(Me.Page, "Informe o Processo.")
            Return False
        End If
        Return True
    End Function

    Private Sub getUsuariosDoGrupo()
        Dim sql As String = " SELECT isnull(dbo.usuariosdogrupo('" & ddlGrupo.SelectedValue & "'), '') as Usuario_Id  "

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Usuarios")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            lblUsuariosDoGrupo.Text = ds.Tables(0).Rows(0)(0).ToString()
        Else
            lblUsuariosDoGrupo.Text = String.Empty
        End If
    End Sub

    Private Sub getGruposDoUsuario()
        Dim sql As String = " SELECT isnull(dbo.gruposdousuario('" & ddlUsuarios.SelectedValue & "'), '') as Usuario_Id  "

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Usuarios")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            lblGruposDoUsuario.Text = ds.Tables(0).Rows(0)(0).ToString()
        Else
            lblGruposDoUsuario.Text = String.Empty
        End If
    End Sub

#End Region

#End Region

#Region "Copia Usuario"

#Region "Events"
    Protected Sub ddlUsuarioOrigem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUsuarioOrigem.SelectedIndexChanged
        CarregarUsuarioOrigemXProcessos()
    End Sub

    Protected Sub ddlUsuarioCopia_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUsuarioCopia.SelectedIndexChanged
        CarregarUsuarioCopiaXProcessos()
    End Sub

    Protected Sub lnkCopiaUsuario_Click(sender As Object, e As EventArgs) Handles lnkCopiaUsuario.Click
        Try
            CopiarUsuario()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkCopiaLimpar_Click(sender As Object, e As EventArgs) Handles lnkCopiaLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

#Region "Methods"
    Private Sub CarregaCopia()
        ddl.Carregar(ddlUsuarioOrigem, CarregarDDL.Tabela.Usuarios, "", True)
        ddl.Carregar(ddlUsuarioCopia, CarregarDDL.Tabela.Usuarios, "", True)
    End Sub

    Private Sub CarregarUsuarioOrigemXProcessos()
        Dim sql = "Select upper(Usuario_Id) as Usuarios, upper(Processo_Id) as Processos From UsuariosXProcessos " & vbCrLf & _
          " Where Usuario_Id = '" & UCase(ddlUsuarioOrigem.SelectedValue) & "' Order By Processo_Id"

        grdOrigem.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        grdOrigem.DataBind()
    End Sub

    Private Sub CarregarUsuarioCopiaXProcessos()
        Dim sql = "Select upper(Usuario_Id) as Usuarios, upper(Processo_Id) as Processos From UsuariosXProcessos " & vbCrLf & _
          " Where Usuario_Id = '" & UCase(ddlUsuarioCopia.SelectedValue) & "' Order By Processo_Id"

        grdCopia.DataSource = Banco.ConsultaDataSet(sql, "Consulta")
        grdCopia.DataBind()
    End Sub

    Private Sub CopiarUsuario()
        Try
            If Funcoes.VerificaPermissao("UsuariosXProcessos", "GRAVAR") Then
                If ddlUsuarioCopia.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Selecione um usario destino da cópia.")
                Else
                    Dim SqlArray As New ArrayList
                    Dim Sql As String

                    For i = 0 To grdOrigem.Rows.Count - 1
                        Sql = " DELETE FROM UsuariosXProcessosXPermissoes" & vbCrLf &
                              " WHERE Usuario_Id  = '" & UCase(ddlUsuarioCopia.SelectedValue) & "'" & vbCrLf &
                              " AND   Processo_Id = '" & UCase(grdOrigem.Rows(i).Cells(0).Text) & "'"
                        SqlArray.Add(Sql)

                        Sql = " DELETE FROM UsuariosXProcessos" & vbCrLf &
                              " WHERE Usuario_Id  = '" & UCase(ddlUsuarioCopia.SelectedValue) & "'" & vbCrLf &
                              " AND   Processo_id = '" & UCase(grdOrigem.Rows(i).Cells(0).Text) & "'" & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " INSERT INTO UsuariosXProcessos (Usuario_Id, Processo_id) " & vbCrLf &
                              " VALUES ('" & UCase(ddlUsuarioCopia.SelectedValue) & "', '" & UCase(grdOrigem.Rows(i).Cells(0).Text) & "')" & vbCrLf
                        SqlArray.Add(Sql)

                        Sql = " INSERT INTO UsuariosXProcessosXPermissoes (Usuario_Id, Processo_Id, Permissao_Id)" & vbCrLf &
                              " SELECT '" & UCase(ddlUsuarioCopia.SelectedValue) & "' , Processo_Id, Permissao_Id " & vbCrLf &
                              " FROM UsuariosXProcessosXPermissoes " & vbCrLf &
                              " WHERE Usuario_Id  = '" & UCase(ddlUsuarioOrigem.SelectedValue) & "'" & vbCrLf &
                              " AND   Processo_Id = '" & UCase(grdOrigem.Rows(i).Cells(0).Text) & "'"
                        SqlArray.Add(Sql)
                    Next

                    'Sql = " DELETE FROM GruposXUsuarios" & vbCrLf &
                    '      " WHERE Usuario_Id = '" & UCase(ddlUsuarioCopia.SelectedValue) & "'"
                    'SqlArray.Add(Sql)

                    'Sql = " INSERT INTO GruposXUsuarios" & vbCrLf &
                    '      " VALUES ((SELECT Grupo_Id FROM GruposXUsuarios WHERE Usuario_Id = '" & UCase(ddlUsuarioOrigem.SelectedValue) & "')," & vbCrLf &
                    '      " '" & UCase(ddlUsuarioCopia.SelectedValue) & "')"
                    'SqlArray.Add(Sql)

                    If SqlArray.Count > 0 Then
                        If Banco.GravaBanco(SqlArray) Then
                            MsgBox(Me.Page, "Usuário transferido com Sucesso.", eTitulo.Sucess)
                            CarregarUsuarioCopiaXProcessos()
                        Else
                            MsgBox(Me.Page, "Erro ao incluir permissão.", eTitulo.Erro)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub

    Private Sub Limpar()
        ddlUsuarioOrigem.SelectedIndex = 0
        ddlUsuarioCopia.SelectedIndex = 0

        grdOrigem.DataSource = Nothing
        grdOrigem.DataBind()

        grdCopia.DataSource = Nothing
        grdCopia.DataBind()
    End Sub
#End Region

#End Region

    Protected Sub lnkAJudaGrupo_Click(sender As Object, e As EventArgs) Handles lnkAJudaGrupo.Click
        Try
            Funcoes.Ajuda(Me.Page, "LiberarPermissoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkCopiaAjuda_Click(sender As Object, e As EventArgs) Handles lnkCopiaAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LiberarPermissoes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class