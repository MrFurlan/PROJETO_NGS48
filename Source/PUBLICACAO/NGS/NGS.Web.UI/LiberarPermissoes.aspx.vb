Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LiberarPermissoes
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("LiberarPermissoes", "ACESSAR") Then
                BuscarGrupos()
                BuscaPermissoes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx", eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

    Public Sub BuscarGrupos()
        Sql = "select Grupo_Id from Grupos order by Grupo_Id"
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "grupo")
        lstGrupos.DataSource = ds
        lstGrupos.DataBind()
    End Sub

    Public Sub BuscaPermissoes()
        Sql = "select Permissao_Id from permissoes order by Permissao_Id"
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Permissao")
        lstPermissoes.DataSource = ds
        lstPermissoes.DataBind()
    End Sub

    Public Sub AtualizaPermissoes()
        If Not String.IsNullOrWhiteSpace(hdfGrupoId.Value) And Not String.IsNullOrWhiteSpace(hdfProcessoId.Value) Then
            Try
                'exclui as permissoes
                ExcluirPermissoes(hdfGrupoId.Value, hdfProcessoId.Value)
                'busca as novas permissoes checadas
                For Each item As ListViewDataItem In lstPermissoes.Items
                    Dim chk As CheckBox = item.FindControl("chkPermissoes")
                    If chk.Checked Then
                        IncluirPermissoes(hdfGrupoId.Value, hdfProcessoId.Value, chk.Text)
                    End If
                Next

                MsgBox(Me.Page, "Permissões Atualizadas com Sucesso.", eTitulo.Sucess)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message.ToString)
            End Try
        Else
            MsgBox(Me.Page, "Por favor selecione um Grupo e um processo !")
        End If

    End Sub

    Protected Sub lstGrupos_OnItemCommand(ByVal sender As Object, ByVal e As ListViewCommandEventArgs)
        If String.Equals(e.CommandName, "AddToList") Then
            Dim dataItem As ListViewDataItem = CType(e.Item, ListViewDataItem)
            Dim GrupoId As String = lstGrupos.DataKeys(dataItem.DisplayIndex).Value.ToString()

            'dar o bind nos processos de acordo com o grupo selecionado
            Dim Sql As String = "Select Processo_id From GruposXProcessos where Grupo_Id = '"
            Sql &= GrupoId + "' order by Processo_Id"
            hdfGrupoId.Value = GrupoId.ToString()

            'limpando as opcões selecionadas anteriormente
            hdfProcessoId.Value = String.Empty
            txtProcesso.Text = String.Empty
            lstParticipantes.DataSource = New DataTable()
            lstParticipantes.DataBind()

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Processos")
            lstProcessos.DataSource = ds
            lstProcessos.DataBind()

            'dar o bind nos usuario de acordo com o grupo selecionado
            Dim SqlUsuario As String = "Select Usuario_id From GruposXUsuarios where Grupo_Id = '"
            SqlUsuario &= GrupoId + "' order by Usuario_Id"

            Dim dsuser As DataSet = Banco.ConsultaDataSet(SqlUsuario, "Usuarios")
            lstUsuarios.DataSource = dsuser
            lstUsuarios.DataBind()

            For Each item As ListViewDataItem In lstPermissoes.Items
                Dim chk As CheckBox = item.FindControl("chkPermissoes")
                chk.Checked = False
            Next
            txtGrupo.Text = GrupoId
        End If
    End Sub

    Protected Sub lstProcessos_OnItemCommand(ByVal sender As Object, ByVal e As ListViewCommandEventArgs)
        'bind nas permissoes do grupo naquele processo
        If String.Equals(e.CommandName, "AddToList") Then
            Dim dataItem As ListViewDataItem = CType(e.Item, ListViewDataItem)
            Dim ProcessoId As String = lstProcessos.DataKeys(dataItem.DisplayIndex).Value.ToString()
            hdfProcessoId.Value = ProcessoId.ToString()

            Sql = "Select Permissao_Id From GruposXProcessosXPermissoes where Grupo_Id = '"
            Sql &= hdfGrupoId.Value + "' and Processo_Id = '" + ProcessoId + "' order by Permissao_Id"

            For Each item As ListViewDataItem In lstPermissoes.Items
                Dim chk As CheckBox = item.FindControl("chkPermissoes")
                chk.Checked = False
            Next
            For Each item As ListViewDataItem In lstPermissoes.Items
                For Each row As DataRow In Banco.ConsultaDataSet(Sql, "GruposXProcessosXPermissao").Tables(0).Rows
                    Dim chk As CheckBox = item.FindControl("chkPermissoes")
                    If chk.Text = row.Item("Permissao_Id").ToString() Then
                        chk.Checked = True
                        Exit For
                    End If
                Next
            Next
            txtProcesso.Text = ProcessoId
        End If

    End Sub

    Protected Sub lstUsuarios_OnItemCommand(ByVal sender As Object, ByVal e As ListViewCommandEventArgs)
        'bind na lista que informa em quais grupos o usuario pertence
        If String.Equals(e.CommandName, "AddToList") Then
            Dim dataItem As ListViewDataItem = CType(e.Item, ListViewDataItem)
            Dim UsuarioId As String = lstUsuarios.DataKeys(dataItem.DisplayIndex).Value.ToString()

            Sql = "Select Grupo_id From GruposXUsuarios where Usuario_Id = '"
            Sql &= UsuarioId + "' order by Grupo_Id"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Participantes")
            lstParticipantes.DataSource = ds
            lstParticipantes.DataBind()
            txtUsuario.Text = UsuarioId
        End If
    End Sub

    Public Function ExcluirPermissoes(ByVal Grupo As String, ByVal Processo As String)
        Dim SqlArray As New ArrayList

        Sql = "Delete From GruposXProcessosXPermissoes Where Grupo_Id = '" & Grupo & "' and "
        Sql &= "Processo_Id = '" & Processo & "'"

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) = False Then
            Return HttpContext.Current.Session("ssMessage")
        Else
            Return ""
        End If

    End Function

    Public Function IncluirPermissoes(ByVal Grupo As String, ByVal Processo As String, ByVal Permissao As String)
        Dim SqlArray As New ArrayList

        Sql = "Insert into GruposXProcessosXPermissoes(Grupo_Id, Processo_Id, Permissao_Id) "
        Sql &= "Values('" & Grupo & "','" & Processo & "','" & Permissao & "')"

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) = False Then
            Return HttpContext.Current.Session("ssMessage")
        Else
            Return ""
        End If

    End Function

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdAjuda.Click
        Dim NomeArquivo As String = "Manual/LiberarPermissoes.mht"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        ScriptManager.RegisterClientScriptBlock(Me, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
    End Sub

    Protected Sub btnSair_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSair.Click
        Response.Redirect("~/Gestao.aspx")
    End Sub

    Protected Sub btnAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnAtualizar.Click
        AtualizaPermissoes()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LiberarPermissoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class