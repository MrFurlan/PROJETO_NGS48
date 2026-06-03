Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PlanoDeRateio
    Inherits BasePage

    Private Mensagem As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Custos)
                If Funcoes.VerificaPermissao("PlanoDeRateio", "ACESSAR") Then
                    HabilitarCampos()
                    ListarGrupoDeProduto()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Custos.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Private Sub HabilitarCampos()

        btnEmpresa.Enabled = True
        ddlGrupoProduto.Enabled = True
        ddlProduto.Enabled = True
        txtPercentual.Enabled = True
        lnkLimpar.Enabled = True
        imgIncluir.Enabled = True
    End Sub

    Private Sub Limpar()
        ddlGrupoProduto.SelectedIndex = 0
        ddlProduto.Items.Clear()
        txtPercentual.Text = ""
        gridPlanoDeRateio.DataBind()

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        Session.Remove("ssPlanoDeRateio")
        Dim dtPlanoDeRateio As New DataTable("PlanoDeRateio")
        dtPlanoDeRateio.Columns.Add("Indice", GetType(String))
        dtPlanoDeRateio.Columns.Add("Empresa", Type.GetType("System.String"))
        dtPlanoDeRateio.Columns.Add("Grupo", Type.GetType("System.String"))
        dtPlanoDeRateio.Columns.Add("Produto", Type.GetType("System.String"))
        dtPlanoDeRateio.Columns.Add("Percentual", Type.GetType("System.Double"))
        Session("ssPlanoDeRateio") = dtPlanoDeRateio

        If txtEmpresa.Text = "" Then
            ListarEmpresas()
        End If

        HID.Value = Guid.NewGuid().ToString
        ucConsultaEmpresas.SetarHID(HID.Value)

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            btnEmpresa.Enabled = False
            txtEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objEmpresaPDR" & HID.Value) Is Nothing Then
            Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmpresaPDR" & HID.Value), [Lib].Negocio.Cliente))
            txtEmpresa.Text = itemEmpresa.Text
            txtCodigoEmpresa.Value = itemEmpresa.Value
            Session.Remove("objEmpresaPDR" & HID.Value)
        End If
    End Sub

    Private Sub ListarEmpresas()
        Dim objEmpresa As New [Lib].Negocio.Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
        Dim itemEmpresa As ListItem = Funcoes.FormatarListItemCliente(objEmpresa)
        txtEmpresa.Text = itemEmpresa.Text
        txtCodigoEmpresa.Value = itemEmpresa.Value
    End Sub

    Private Sub ListarGrupoDeProduto()
        ddl.Carregar(ddlGrupoProduto, CarregarDDL.Tabela.NivelGrupoProduto, "len(grupo_id) = 1 or len(grupo_id) = 3 or len(grupo_id) > 3", True)
    End Sub

    Protected Sub ddlGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlGrupoProduto.SelectedIndex > 0 AndAlso ddlGrupoProduto.SelectedValue.ToString.Length > 3 Then
                ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = " & ddlGrupoProduto.SelectedValue, True)
            Else
                ddlProduto.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidarCampos() As Boolean
        If txtCodigoEmpresa.Value.ToString.Length = 0 Then
            Mensagem = "Empresa não foi selecionada."
            Return False
        ElseIf ddlGrupoProduto.SelectedIndex = 0 Then
            Mensagem = "Grupo do produto não foi selecionado."
            Return False
        ElseIf txtPercentual.Text.Length = 0 Then
            Mensagem = "Percentual não foi Informado."
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub btnEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "Livre"
            ucConsultaEmpresas.Limpar()
            Popup.ConsultaDeEmpresas(Me.Page, "objEmpresaPDR" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgIncluir_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If ValidarCampos() Then
                Dim percentual As Double = 0
                Dim tem As Boolean = False
                For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                    If CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Empresa") = txtCodigoEmpresa.Value And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Grupo") = ddlGrupoProduto.SelectedValue And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Produto") = ddlProduto.SelectedValue And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual") = txtPercentual.Text Then
                        tem = True
                    End If
                    percentual += CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual")
                Next

                If tem Then
                    MsgBox(Me.Page, "Item já existente.")
                ElseIf percentual + CDbl(txtPercentual.Text) > 100 Then
                    MsgBox(Me.Page, "A soma dos itens mais o item atual não pode ultrapassar 100%.")
                Else
                    Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
                    Dim drRow As DataRow = CType(Session("ssPlanoDeRateio"), DataTable).NewRow()
                    drRow("Indice") = Empresa(0) & ";" & Empresa(1) & ";" & ddlGrupoProduto.SelectedValue & ";" & IIf(ddlProduto.SelectedIndex = 0, "", ddlProduto.SelectedValue) & ";" & txtPercentual.Text
                    drRow("Empresa") = Empresa(0) & "-" & Empresa(1)
                    drRow("Grupo") = ddlGrupoProduto.SelectedValue
                    drRow("Produto") = ddlProduto.SelectedValue
                    drRow("Percentual") = txtPercentual.Text
                    CType(Session("ssPlanoDeRateio"), DataTable).Rows.Add(drRow)

                    gridPlanoDeRateio.DataSource = CType(Session("ssPlanoDeRateio"), DataTable)
                    gridPlanoDeRateio.DataBind()

                    ddlGrupoProduto.SelectedIndex = 0
                    ddlProduto.Items.Clear()
                    txtPercentual.Text = ""
                End If

                If gridPlanoDeRateio.Rows.Count > 0 And lnkNovo.Enabled = False Then
                    lnkNovo.Enabled = True
                End If
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridPlanoDeRateio_RowCommand(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewCommandEventArgs)
        Try
            If e.CommandName = "Select" Then
                Dim strArrayItem As String() = e.CommandArgument.ToString().Split(";")

                For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                    If CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Empresa") = (strArrayItem(0) & "-" & strArrayItem(1)).ToString And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Grupo") = strArrayItem(2) And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Produto") = strArrayItem(3) And CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual") = strArrayItem(4) Then
                        CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Delete()
                        Exit For
                    End If
                Next

                gridPlanoDeRateio.DataSource = CType(Session("ssPlanoDeRateio"), DataTable)
                gridPlanoDeRateio.DataBind()

                lnkNovo.Parent.Visible = False
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeRateio", "GRAVAR") Then
                Dim percentual As Double = 0
                For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                    percentual += CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual")
                Next

                If percentual < 100 Then
                    MsgBox(Me.Page, "A soma do percentual dos itens informados deve ser 100%.")
                Else
                    Dim arraySql As New ArrayList
                    Dim sql As String
                    Dim Empresa() As String

                    For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                        Empresa = CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Empresa").ToString.Split("-")

                        sql = "INSERT INTO PlanoDeRateio (Empresa_id, EndEmpresa_id, Grupo_id, Produto_id, Percentual) " & vbCrLf &
                              "VALUES " & vbCrLf &
                              "('" & Empresa(0) & "', " & Empresa(1) & ", " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Grupo") & "', " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Produto") & "', " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual") & "')" & vbCrLf
                        arraySql.Add(sql)
                    Next

                    If Banco.GravaBanco(arraySql) Then
                        MsgBox(Me.Page, "Registro incluido com Sucesso.", eTitulo.Sucess)
                        Limpar()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeRateio", "ALTERAR") Then
                Dim percentual As Double = 0
                Dim Empresa() As String = Nothing
                For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                    percentual += CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual")
                    Empresa = CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Empresa").ToString.Split("-")
                Next
                If percentual < 100 Then
                    MsgBox(Me.Page, "A soma do Percentual dos itens informados deve ser 100%")
                Else
                    Dim arraySql As New ArrayList
                    Dim sql As String
                    sql = "DELETE FROM PlanoDeRateio " & vbCrLf &
                          "WHERE " & vbCrLf &
                          "Empresa_id = '" & Empresa(0) & "' AND EndEmpresa_id = " & Empresa(1)
                    arraySql.Add(sql)
                    For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                        sql = "INSERT INTO PlanoDeRateio (Empresa_id, EndEmpresa_id, Grupo_id, Produto_id, Percentual) " & vbCrLf &
                              "VALUES " & vbCrLf &
                              "('" & Empresa(0) & "', " & Empresa(1) & ", " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Grupo") & "', " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Produto") & "', " & vbCrLf &
                              "'" & CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Percentual") & "')" & vbCrLf
                        arraySql.Add(sql)
                    Next
                    If Banco.GravaBanco(arraySql) Then
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("PlanoDeRateio", "EXCLUIR") Then
                Dim Empresa() As String = Nothing
                For i = 0 To CType(Session("ssPlanoDeRateio"), DataTable).Rows.Count - 1
                    Empresa = CType(Session("ssPlanoDeRateio"), DataTable).Rows(i).Item("Empresa").ToString.Split("-")
                Next

                Dim arraySql As New ArrayList
                Dim sql As String

                sql = "DELETE FROM PlanoDeRateio " & vbCrLf &
                      "WHERE " & vbCrLf &
                      "Empresa_id = '" & Empresa(0) & "' AND EndEmpresa_id = " & Empresa(1) & vbCrLf
                arraySql.Add(sql)

                If Banco.GravaBanco(arraySql) Then
                    MsgBox(Me.Page, "Registro excluido com Sucesso.", eTitulo.Sucess)
                    Limpar()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If txtEmpresa.Text.Length > 0 Then
                Dim sql As String
                Dim ds As New DataSet
                Dim Empresa() As String = txtCodigoEmpresa.Value.Split("-")
                sql = "SELECT Empresa_id, EndEmpresa_id, Grupo_id, Produto_id, Percentual " & vbCrLf &
                      "FROM PlanoDeRateio " & vbCrLf &
                      "WHERE " & vbCrLf &
                      "Empresa_id = '" & Empresa(0) & "' AND EndEmpresa_id = 0" & vbCrLf
                ds = Banco.ConsultaDataSet(sql, "PlanoDeRateio")

                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Não existem Registros da Empresa selecionada.")
                Else
                    For Each dr As DataRow In ds.Tables(0).Rows
                        Dim drRow As DataRow = CType(Session("ssPlanoDeRateio"), DataTable).NewRow()
                        drRow("Indice") = dr("Empresa_Id") & ";" & dr("EndEmpresa_id") & ";" & dr("Grupo_id") & ";" & dr("Produto_id") & ";" & dr("Percentual")
                        drRow("Empresa") = dr("Empresa_Id") & "-" & dr("EndEmpresa_id")
                        drRow("Grupo") = dr("Grupo_id")
                        drRow("Produto") = dr("Produto_id")
                        drRow("Percentual") = dr("Percentual")
                        CType(Session("ssPlanoDeRateio"), DataTable).Rows.Add(drRow)
                    Next
                    gridPlanoDeRateio.DataSource = CType(Session("ssPlanoDeRateio"), DataTable)
                    gridPlanoDeRateio.DataBind()

                    ddlGrupoProduto.SelectedIndex = 0
                    ddlProduto.Items.Clear()
                    txtPercentual.Text = ""

                    lnkNovo.Enabled = False
                    lnkAtualizar.Enabled = True
                    lnkExcluir.Enabled = True
                End If
            Else
                MsgBox(Me.Page, "Empresa não foi selecionada.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PlanoDeRateio")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class