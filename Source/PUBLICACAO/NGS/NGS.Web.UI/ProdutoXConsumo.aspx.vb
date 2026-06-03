Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ProdutoXConsumo
    Inherits BasePage

    Private ds As DataSet
    Private objPrd As [Lib].Negocio.Produto

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ProdutoXConsumo", "ACESSAR") Then

                ddl.Carregar(ddlGrupoProdutoProducao, CarregarDDL.Tabela.GrupoProduto)
                ddl.Carregar(ddlGrupoProdutoConsumo, CarregarDDL.Tabela.GrupoProduto)
                ddl.Carregar(ddlGrupoProdutoInsumo, CarregarDDL.Tabela.GrupoProduto)

                Limpar()

                AtualizarGrid(True)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoxPRD" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque Then
                ddlGrupoProdutoProducao.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutoProducao()
                ddlProdutoProducao.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxPRD" & HID.Value)
                ProdutoProducao()
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        ElseIf Session("objProdutoxCON" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque OrElse objProduto.TemProdutoDeConsumo Then
                ddlGrupoProdutoConsumo.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutoConsumo()
                ddlProdutosConsumo.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxCON" & HID.Value)
                ProdutosConsumo()
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        ElseIf Session("objProdutoxINSU" & HID.Value) IsNot Nothing Then

            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            If objProduto.ControlarEstoque Then
                ddlGrupoProdutoInsumo.SelectedValue = objProduto.CodigoGrupo
                CarregarProdutoInsumo()
                ddlProdutosInsumo.SelectedValue = objProduto.Codigo
                Session.Remove("objProdutoxINSU" & HID.Value)
                ProdutosInsumo()
            Else
                MsgBox(Me.Page, "Produto selecionado não está marcado para CONTROLAR ESTOQUE.", eTitulo.Info)
            End If
        End If
    End Sub

    Private Sub AtualizarGrid(ByVal inicio As Boolean)
        Dim objProducao = New ListProdutoXConsumos()

        objProducao.CarregarTudo()

        If inicio Then
            If objProducao.Count > 0 Then
                TabContainer1.ActiveTabIndex = 0
            Else
                TabContainer1.ActiveTabIndex = 1
            End If
        End If

        Dim prd As String = String.Empty

        If objProducao.Count > 0 Then
            For Each row In objProducao

                If Not row.Produto.Codigo = prd Then
                    Dim drItem As DataRow = CType(Session("objProducaoXConsumo" & HID.Value), DataTable).NewRow()

                    drItem("Produto") = row.Produto.Codigo
                    drItem("Nome") = row.Produto.Nome

                    CType(Session("objProducaoXConsumo" & HID.Value), DataTable).Rows.Add(drItem)

                    prd = row.Produto.Codigo
                End If
            Next

            Dim ds1 As DataView
            ds1 = New DataView(CType(Session("objProducaoXConsumo" & HID.Value), DataTable))
            ds1.Sort = "Nome ASC"

            gridProdutoXConsumo.DataSource = ds1
        Else
            gridProdutoXConsumo.DataSource = Nothing
        End If

        gridProdutoXConsumo.DataBind()
    End Sub

    Protected Sub gridProdutoXConsumo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        lnkNovo.Parent.Visible = False
        lnkExcluir.Parent.Visible = True

        CType(Session("objProducao" & HID.Value), DataTable).Rows.Clear()
        CType(Session("objConsumo" & HID.Value), DataTable).Rows.Clear()
        CType(Session("objInsumo" & HID.Value), DataTable).Rows.Clear()

        Dim objProducao = New ListProdutoXConsumos()
        objProducao.CarregarProduto(gridProdutoXConsumo.SelectedRow.Cells(1).Text())

        Dim objProducaoInsumo = New ListProdutoXInsumos()
        objProducaoInsumo.CarregarProduto(gridProdutoXConsumo.SelectedRow.Cells(1).Text())

        Dim primeira = True

        For Each row In objProducao
            If primeira Then
                Dim drItem As DataRow = CType(Session("objProducao" & HID.Value), DataTable).NewRow()

                drItem("Produto") = row.Produto.Codigo
                drItem("NomeProduto") = row.Produto.Nome
                CType(Session("objProducao" & HID.Value), DataTable).Rows.Add(drItem)

                primeira = False
            End If

            Dim drItemConsumo As DataRow = CType(Session("objConsumo" & HID.Value), DataTable).NewRow()

            drItemConsumo("Produto") = row.ProdutoConsumo.Codigo
            drItemConsumo("NomeProduto") = row.ProdutoConsumo.Nome
            drItemConsumo("Percentual") = row.Percentual
            CType(Session("objConsumo" & HID.Value), DataTable).Rows.Add(drItemConsumo)
        Next

        ddlGrupoProdutoProducao.Enabled = False
        lnkBuscaProdutoProducao.Enabled = False
        ddlProdutoProducao.Enabled = False

        gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
        gridProducao.DataBind()

        gridConsumo.DataSource = CType(Session("objConsumo" & HID.Value), DataTable)
        gridConsumo.DataBind()

        lblPercentual.Text = "TOTAL = " & CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") & "%"

        For Each dr In objProducaoInsumo
            Dim drItemInsumo As DataRow = CType(Session("objInsumo" & HID.Value), DataTable).NewRow()

            drItemInsumo("Produto") = dr.ProdutoInsumo.Codigo
            drItemInsumo("NomeProduto") = dr.ProdutoInsumo.Nome
            drItemInsumo("Base") = CInt(dr.Base)
            CType(Session("objInsumo" & HID.Value), DataTable).Rows.Add(drItemInsumo)
        Next

        gridInsumo.DataSource = CType(Session("objInsumo" & HID.Value), DataTable)
        gridInsumo.DataBind()

        TabContainer1.ActiveTabIndex = 1
    End Sub

    Protected Sub ddlGrupoProdutoProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoProducao.SelectedIndexChanged
        CarregarProdutoProducao()
    End Sub

    Private Sub CarregarProdutoProducao()
        Try
            ddl.Carregar(ddlProdutoProducao, CarregarDDL.Tabela.Produto, " ControlarEstoque = 'S' and Grupo ='" & ddlGrupoProdutoProducao.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoProducao_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoProducao.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxPRD" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProdutoProducao_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutoProducao.SelectedIndexChanged
         ProdutoProducao()
    End Sub

    Private Sub ProdutoProducao()
        Try
            If ddlProdutoProducao.SelectedIndex = 0 Then
                ddlProdutoProducao.SelectedIndex = 0
                MsgBox(Me.Page, "Produto para Produção não foi informado.", eTitulo.Info)
            ElseIf gridProducao.Rows.Count > 0 Then
                MsgBox(Me.Page, "Produto para Produção já foi informado.", eTitulo.Info)
            Else
                Dim objProducao = New ListProdutoXConsumos()

                objProducao.CarregarProduto(ddlProdutoProducao.SelectedValue)

                If objProducao.Count > 0 Then
                    ddlProdutoProducao.SelectedIndex = 0
                    MsgBox(Me.Page, "Produto para Produção já está cadastrado.", eTitulo.Info)
                    Exit Sub
                End If

                objPrd = New Produto(ddlProdutoProducao.SelectedValue)

                Dim drItem As DataRow = CType(Session("objProducao" & HID.Value), DataTable).NewRow()

                drItem("Produto") = objPrd.Codigo
                drItem("NomeProduto") = objPrd.Nome
                CType(Session("objProducao" & HID.Value), DataTable).Rows.Add(drItem)

                gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
                gridProducao.DataBind()

                ddlGrupoProdutoProducao.Enabled = False
                lnkBuscaProdutoProducao.Enabled = False
                ddlProdutoProducao.Enabled = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemoverProducao_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim objProducao = New ListProdutoXConsumos()

        objProducao.CarregarProduto(gridProducao.Rows(row.RowIndex).Cells(0).Text)

        If objProducao.Count > 0 Then
            MsgBox(Me.Page, "Produto da Produção só pode ser removido atráves do botão Excluir.", eTitulo.Info)
        ElseIf gridConsumo.Rows.Count > 0 Then
            MsgBox(Me.Page, "Produto da Produção só pode ser removido após a remoção dos itens de Consumo.", eTitulo.Info)
        Else
            CType(Session("objProducao" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

            gridProducao.DataSource = CType(Session("objProducao" & HID.Value), DataTable)
            gridProducao.DataBind()

            ddlGrupoProdutoProducao.Enabled = True
            lnkBuscaProdutoProducao.Enabled = True
            ddlProdutoProducao.Enabled = True

            ddlGrupoProdutoProducao.SelectedIndex = 0
            ddlProdutoProducao.Items.Clear()
        End If
    End Sub

    Protected Sub ddlGrupoProdutoConsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoConsumo.SelectedIndexChanged
        CarregarProdutoConsumo()
    End Sub

    Private Sub CarregarProdutoConsumo()
        Try
            ddl.Carregar(ddlProdutosConsumo, CarregarDDL.Tabela.Produto, " ((ControlarEstoque = 'S') or Produto_Id in ( select pc.Produto_id from ProdutoDeConsumo pc inner join Produtos p on p.Produto_Id = pc.Produto_id where  p.Grupo = '" & ddlGrupoProdutoConsumo.SelectedValue & "')) and Grupo ='" & ddlGrupoProdutoConsumo.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoConsumo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoConsumo.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxCON" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProdutosConsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutosConsumo.SelectedIndexChanged
        ProdutosConsumo()
    End Sub

    Private Sub ProdutosConsumo()
        Try
            If String.IsNullOrWhiteSpace(txtPercentual.Text) OrElse CDec(txtPercentual.Text) = 0 Then
                ddlProdutosConsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Percentual para Consumo não foi informado.", eTitulo.Info)
            ElseIf CDec(txtPercentual.Text) > 100 Then
                ddlProdutosConsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Percentual para Consumo não pode ser maior que 100%.", eTitulo.Info)
            ElseIf ddlProdutosConsumo.SelectedIndex = 0 Then
                ddlProdutosConsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Produto para Consumo não foi informado.", eTitulo.Info)
            Else
                If CType(Session("objConsumo" & HID.Value), DataTable).Rows.Count > 0 Then

                    If Not IsDBNull(CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "Produto = " & ddlProdutosConsumo.SelectedValue)) Then
                        ddlProdutosConsumo.SelectedIndex = 0
                        MsgBox(Me.Page, "Produto já existe na lista para Consumo.", eTitulo.Info)
                        Exit Sub
                    ElseIf (CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") + CDec(txtPercentual.Text)) > 100 Then
                        ddlProdutosConsumo.SelectedIndex = 0
                        MsgBox(Me.Page, "A soma do Percentual para Consumo dos Produtos não pode ser maior que 100%.", eTitulo.Info)
                        Exit Sub
                    End If
                End If

                objPrd = New Produto(ddlProdutosConsumo.SelectedValue)

                Dim objProducao = New ListProdutoXConsumos()

                objProducao.CarregarProduto(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"))

                If objProducao.Count > 0 Then
                    If Funcoes.VerificaPermissao("ProdutoXConsumo", "GRAVAR") Then
                        Dim objPxC = New ProdutoXConsumos()
                        objPxC.IUD = "I"
                        objPxC.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                        objPxC.CodigoProdutoConsumo = objPrd.Codigo
                        objPxC.Percentual = CDec(txtPercentual.Text)

                        If Not objPxC.Salvar Then
                            MsgBox(Me.Page, "Erro na inclusão do Produto para Consumo.", eTitulo.Info)
                            Exit Sub
                        End If
                    Else
                        MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
                        Exit Sub
                    End If
                End If

                Dim drItem As DataRow = CType(Session("objConsumo" & HID.Value), DataTable).NewRow()

                drItem("Produto") = objPrd.Codigo
                drItem("NomeProduto") = objPrd.Nome
                drItem("Percentual") = txtPercentual.Text
                CType(Session("objConsumo" & HID.Value), DataTable).Rows.Add(drItem)

                gridConsumo.DataSource = CType(Session("objConsumo" & HID.Value), DataTable)
                gridConsumo.DataBind()

                lblPercentual.Text = "TOTAL = " & CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") & "%"

                txtPercentual.Text = String.Empty
                ddlProdutosConsumo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridConsumo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        txtPercentual.Text = CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Percentual")

        objPrd = New Produto(CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Produto"))
        ddlGrupoProdutoConsumo.SelectedValue = objPrd.CodigoGrupo
        ddl.Carregar(ddlProdutosConsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoConsumo.SelectedValue & "'")

        With ddlProdutosConsumo
            ddlProdutosConsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objPrd.Codigo))
        End With

        ddlGrupoProdutoConsumo.Enabled = False
        lnkBuscaProdutoConsumo.Enabled = False
        ddlProdutosConsumo.Enabled = False

        btnPercentual.Visible = True

    End Sub

    Protected Sub imgRemoverConsumo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim objProducao = New ListProdutoXConsumos()

        objProducao.CarregarProduto(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"))

        If objProducao.Count > 0 Then
            If Funcoes.VerificaPermissao("ProdutoXConsumo", "EXCLUIR") Then
                Dim objPxC = New ProdutoXConsumos()
                objPxC.IUD = "D"
                objPxC.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                objPxC.CodigoProdutoConsumo = CType(Session("objConsumo" & HID.Value), DataTable).Rows(row.RowIndex).Item("Produto")

                If objPxC.TemConsumo Then
                    MsgBox(Me.Page, "Registro está sendo usado na Ordem de Produção.")
                    Exit Sub
                Else
                    If Not objPxC.Salvar Then
                        MsgBox(Me.Page, "Erro ao remover Produto para Consumo.", eTitulo.Info)
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.")
            End If
        End If

        CType(Session("objConsumo" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

        gridConsumo.DataSource = CType(Session("objConsumo" & HID.Value), DataTable)
        gridConsumo.DataBind()

        lblPercentual.Text = "TOTAL = " & CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") & "%"

        txtPercentual.Text = String.Empty
        ddlGrupoProdutoConsumo.SelectedIndex = 0
        ddlProdutosConsumo.Items.Clear()
    End Sub

    Protected Sub btnPercentual_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Dim total As Decimal = CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") _
                                   - CDec(CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Percentual")) _
                                   + CDec(txtPercentual.Text)

            If total > 100 Then
                MsgBox(Me.Page, "Total do(s) Produto(s) para Consumo não pode ultrapassar 100%.", eTitulo.Info)
                Exit Sub
            End If

            Dim objProducao = New ListProdutoXConsumos()

            objProducao.CarregarConsumo(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"), CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Produto"))

            If objProducao.Count > 0 Then
                If Funcoes.VerificaPermissao("ProdutoXConsumo", "ALTERAR") Then
                    Dim objPxC = New ProdutoXConsumos()
                    objPxC.IUD = "U"
                    objPxC.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                    objPxC.CodigoProdutoConsumo = CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Produto")
                    objPxC.Percentual = CDec(txtPercentual.Text)

                    If Not objPxC.Salvar Then
                        MsgBox(Me.Page, "Erro ao atualizar o Percentual do Produto para Consumo.", eTitulo.Info)
                        Exit Sub
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para aterar registro.")
                End If
            End If

            CType(Session("objConsumo" & HID.Value), DataTable).Rows(gridConsumo.SelectedIndex).Item("Percentual") = txtPercentual.Text

            gridConsumo.DataSource = CType(Session("objConsumo" & HID.Value), DataTable)
            gridConsumo.DataBind()

            lblPercentual.Text = "TOTAL = " & CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") & "%"

            btnPercentual.Visible = False
            ddlGrupoProdutoConsumo.Enabled = True
            lnkBuscaProdutoConsumo.Enabled = True
            ddlProdutosConsumo.Enabled = True
            txtPercentual.Text = String.Empty
            ddlProdutosConsumo.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupoProdutoInsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlGrupoProdutoInsumo.SelectedIndexChanged
        CarregarProdutoInsumo()
    End Sub

    Private Sub CarregarProdutoInsumo()
        Try
            ddl.Carregar(ddlProdutosInsumo, CarregarDDL.Tabela.Produto, " ((ControlarEstoque = 'S') or Produto_Id in ( select pc.Produto_id from ProdutoDeConsumo pc inner join Produtos p on p.Produto_Id = pc.Produto_id where  p.Grupo = '" & ddlGrupoProdutoInsumo.SelectedValue & "')) and Grupo ='" & ddlGrupoProdutoInsumo.SelectedValue & "'")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBuscaProdutoInsumo_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProdutoInsumo.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxINSU" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProdutosInsumo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlProdutosInsumo.SelectedIndexChanged
        ProdutosInsumo()
    End Sub

    Private Sub ProdutosInsumo()
        Try
            If ddlProdutosInsumo.SelectedIndex = 0 Then
                ddlProdutosInsumo.SelectedIndex = 0
                MsgBox(Me.Page, "Produto para Insumo não foi informado.", eTitulo.Info)
            Else
                If CType(Session("objInsumo" & HID.Value), DataTable).Rows.Count > 0 Then

                    If Not IsDBNull(CType(Session("objInsumo" & HID.Value), DataTable).Compute("SUM(Base)", "Produto = " & ddlProdutosInsumo.SelectedValue)) Then
                        ddlProdutosInsumo.SelectedIndex = 0
                        MsgBox(Me.Page, "Produto já existe na lista para Insumo.", eTitulo.Info)
                        Exit Sub
                    End If
                End If

                objPrd = New Produto(ddlProdutosInsumo.SelectedValue)

                Dim objProducao = New ListProdutoXConsumos()

                objProducao.CarregarProduto(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"))

                If objProducao.Count > 0 Then
                    If Funcoes.VerificaPermissao("ProdutoXConsumo", "GRAVAR") Then
                        Dim objPxI = New ProdutoXInsumos()
                        objPxI.IUD = "I"
                        objPxI.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                        objPxI.CodigoProdutoInsumo = objPrd.Codigo
                        objPxI.Base = 0
                        If txtBase.Text.Length > 0 Then objPxI.Base = CDec(txtBase.Text)

                        If Not objPxI.Salvar Then
                            MsgBox(Me.Page, "Erro na inclusão do Produto para Insumo.", eTitulo.Info)
                            Exit Sub
                        End If
                    Else
                        MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
                        Exit Sub
                    End If
                End If

                Dim drItem As DataRow = CType(Session("objInsumo" & HID.Value), DataTable).NewRow()

                drItem("Produto") = objPrd.Codigo
                drItem("NomeProduto") = objPrd.Nome
                drItem("Base") = 0
                If txtBase.Text.Length > 0 Then drItem("Base") = CInt(txtBase.Text)

                CType(Session("objInsumo" & HID.Value), DataTable).Rows.Add(drItem)

                gridInsumo.DataSource = CType(Session("objInsumo" & HID.Value), DataTable)
                gridInsumo.DataBind()

                ddlProdutosInsumo.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridInsumo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

        txtBase.Text = CType(Session("objInsumo" & HID.Value), DataTable).Rows(gridInsumo.SelectedIndex).Item("Base")

        objPrd = New Produto(CType(Session("objInsumo" & HID.Value), DataTable).Rows(gridInsumo.SelectedIndex).Item("Produto"))
        ddlGrupoProdutoInsumo.SelectedValue = objPrd.CodigoGrupo
        ddl.Carregar(ddlProdutosInsumo, CarregarDDL.Tabela.Produto, " Situacao = 1 and Grupo ='" & ddlGrupoProdutoInsumo.SelectedValue & "'")

        With ddlProdutosInsumo
            ddlProdutosInsumo.SelectedIndex = .Items.IndexOf(.Items.FindByValue(objPrd.Codigo))
        End With

        ddlGrupoProdutoInsumo.Enabled = False
        lnkBuscaProdutoInsumo.Enabled = False
        ddlProdutosInsumo.Enabled = False

        btnBase.Visible = True

    End Sub

    Protected Sub imgRemoverInsumo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim img As ImageButton = CType(sender, ImageButton)
        Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        Dim objProducao = New ListProdutoXInsumos()

        objProducao.CarregarProduto(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"))

        If objProducao.Count > 0 Then
            If Funcoes.VerificaPermissao("ProdutoXConsumo", "EXCLUIR") Then
                Dim objPxI = New ProdutoXInsumos()
                objPxI.IUD = "D"
                objPxI.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                objPxI.CodigoProdutoInsumo = CType(Session("objInsumo" & HID.Value), DataTable).Rows(row.RowIndex).Item("Produto")

                If objPxI.TemInsumo Then
                    MsgBox(Me.Page, "Registro está sendo usado na Ordem de Produção.")
                    Exit Sub
                Else
                    If Not objPxI.Salvar Then
                        MsgBox(Me.Page, "Erro ao remover Produto para Insumo.", eTitulo.Info)
                        Exit Sub
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.")
            End If
        End If

        CType(Session("objInsumo" & HID.Value), DataTable).Rows(row.RowIndex).Delete()

        gridInsumo.DataSource = CType(Session("objInsumo" & HID.Value), DataTable)
        gridInsumo.DataBind()

        ddlGrupoProdutoInsumo.SelectedIndex = 0
        ddlProdutosInsumo.Items.Clear()
    End Sub


    Protected Sub btnBase_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim objProducao = New ListProdutoXInsumos()

            objProducao.CarregarInsumo(CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto"), CType(Session("objInsumo" & HID.Value), DataTable).Rows(gridInsumo.SelectedIndex).Item("Produto"))

            If objProducao.Count > 0 Then
                If Funcoes.VerificaPermissao("ProdutoXConsumo", "ALTERAR") Then
                    Dim objPxI = New ProdutoXInsumos()
                    objPxI.IUD = "U"
                    objPxI.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                    objPxI.CodigoProdutoInsumo = CType(Session("objInsumo" & HID.Value), DataTable).Rows(gridInsumo.SelectedIndex).Item("Produto")
                    objPxI.Base = 0
                    If txtBase.Text.Length > 0 Then objPxI.Base = CDec(txtBase.Text)

                    If Not objPxI.Salvar Then
                        MsgBox(Me.Page, "Erro ao atualizar a Base do Produto para Insumo.", eTitulo.Info)
                        Exit Sub
                    End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para aterar registro.")
                    Exit Sub
                End If
            End If

            CType(Session("objInsumo" & HID.Value), DataTable).Rows(gridInsumo.SelectedIndex).Item("Base") = txtBase.Text

            gridInsumo.DataSource = CType(Session("objInsumo" & HID.Value), DataTable)
            gridInsumo.DataBind()

            btnBase.Visible = False
            ddlGrupoProdutoInsumo.Enabled = True
            lnkBuscaProdutoInsumo.Enabled = True
            ddlProdutosInsumo.Enabled = True
            txtBase.Text = String.Empty
            ddlProdutosInsumo.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function ValidarSelecao() As Boolean
        If CType(Session("objProducao" & HID.Value), DataTable).Rows.Count = 0 Then
            MsgBox(Me.Page, "Produto para Produção não foi informado.", eTitulo.Info)
            Return False
        ElseIf CType(Session("objConsumo" & HID.Value), DataTable).Rows.Count = 0 Then
            MsgBox(Me.Page, "Produto para Consumo não foi informado.", eTitulo.Info)
            Return False
        ElseIf CType(Session("objConsumo" & HID.Value), DataTable).Compute("SUM(Percentual)", "") <> 100 Then
            MsgBox(Me.Page, "Quantidade total do Percentual para Consumo não pode ser diferente de 100%.", eTitulo.Info)
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub Limpar()
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False

        Session.Remove("objProducaoXConsumo" & HID.Value)
        Session.Remove("objProducao" & HID.Value)
        Session.Remove("objConsumo" & HID.Value)
        Session.Remove("objInsumo" & HID.Value)

        HID.Value = Guid.NewGuid().ToString

        Dim dtProdutoXConsumo As New DataTable("ItemProdutoXConsumo")
        dtProdutoXConsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtProdutoXConsumo.Columns.Add("Nome", Type.GetType("System.String"))
        Session("objProducaoXConsumo" & HID.Value) = dtProdutoXConsumo

        Dim dtProducao As New DataTable("ItemProducao")
        dtProducao.Columns.Add("Produto", Type.GetType("System.String"))
        dtProducao.Columns.Add("NomeProduto", Type.GetType("System.String"))
        Session("objProducao" & HID.Value) = dtProducao

        Dim dtConsumo As New DataTable("ItemConsumo")
        dtConsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtConsumo.Columns.Add("Percentual", Type.GetType("System.Decimal"))
        Session("objConsumo" & HID.Value) = dtConsumo

        Dim dtInsumo As New DataTable("ItemInsumo")
        dtInsumo.Columns.Add("Produto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("NomeProduto", Type.GetType("System.String"))
        dtInsumo.Columns.Add("Base", Type.GetType("System.Int64"))
        Session("objInsumo" & HID.Value) = dtInsumo

        ddlGrupoProdutoProducao.Enabled = True
        lnkBuscaProdutoProducao.Enabled = True
        ddlProdutoProducao.Enabled = True
        btnPercentual.Visible = False
        txtPercentual.Enabled = True
        ddlGrupoProdutoConsumo.Enabled = True
        lnkBuscaProdutoConsumo.Enabled = True
        ddlProdutosConsumo.Enabled = True

        ddlGrupoProdutoProducao.SelectedIndex = 0
        ddlProdutoProducao.Items.Clear()
        gridProducao.DataSource = Nothing
        gridProducao.DataBind()

        txtPercentual.Text = String.Empty
        lblPercentual.Text = String.Empty
        ddlGrupoProdutoConsumo.SelectedIndex = 0
        ddlProdutosConsumo.Items.Clear()
        gridConsumo.DataSource = Nothing
        gridConsumo.DataBind()

        btnBase.Visible = False
        txtBase.Enabled = True
        txtBase.Text = String.Empty
        ddlGrupoProdutoInsumo.Enabled = True
        lnkBuscaProdutoInsumo.Enabled = True
        ddlProdutosInsumo.Enabled = True
        ddlGrupoProdutoInsumo.SelectedIndex = 0
        ddlProdutosInsumo.Items.Clear()
        gridInsumo.DataSource = Nothing
        gridInsumo.DataBind()
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        If Funcoes.VerificaPermissao("ProdutoXConsumo", "GRAVAR") Then
            If ValidarSelecao() Then
                Dim SQls As New ArrayList
                Dim objBanco As New AcessaBanco()

                Dim objProducao = New ListProdutoXConsumos()

                For Each pXc In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                    Dim objPxC = New ProdutoXConsumos()
                    objPxC.IUD = "I"
                    objPxC.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                    objPxC.CodigoProdutoConsumo = pXc("Produto")
                    objPxC.Percentual = pXc("Percentual")

                    objProducao.Add(objPxC)
                Next

                objProducao.SalvarSql(SQls)

                Dim objProducaoXI = New ListProdutoXInsumos()

                For Each pXi In CType(Session("objInsumo" & HID.Value), DataTable).Rows
                    Dim objPxI = New ProdutoXInsumos()
                    objPxI.IUD = "I"
                    objPxI.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                    objPxI.CodigoProdutoInsumo = pXi("Produto")
                    objPxI.Base = CDec(pXi("Base"))

                    objProducaoXI.Add(objPxI)
                Next

                objProducaoXI.SalvarSql(SQls)

                If objBanco.GravaBanco(SQls) Then
                    MsgBox(Me.Page, "Registro incluído com sucesso.", eTitulo.Sucess)
                    Limpar()
                    AtualizarGrid(False)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString, eTitulo.Erro)
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        If Funcoes.VerificaPermissao("ProdutoXConsumo", "EXCLUIR") Then
            Dim SQls As New ArrayList
            Dim objBanco As New AcessaBanco()

            Dim objProducaoXC = New ListProdutoXConsumos()

            For Each pXc In CType(Session("objConsumo" & HID.Value), DataTable).Rows
                Dim objPxC = New ProdutoXConsumos()
                objPxC.IUD = "D"
                objPxC.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                objPxC.CodigoProdutoConsumo = pXc("Produto")
                objPxC.Percentual = pXc("Percentual")

                objProducaoXC.Add(objPxC)
            Next

            objProducaoXC.SalvarSql(SQls)

            Dim objProducaoXI = New ListProdutoXInsumos()

            For Each pXi In CType(Session("objInsumo" & HID.Value), DataTable).Rows
                Dim objPxI = New ProdutoXInsumos()
                objPxI.IUD = "D"
                objPxI.CodigoProduto = CType(Session("objProducao" & HID.Value), DataTable).Rows(0).Item("Produto")
                objPxI.CodigoProdutoInsumo = pXi("Produto")
                objPxI.Base = CDec(pXi("Base"))

                objProducaoXI.Add(objPxI)
            Next

            objProducaoXI.SalvarSql(SQls)

            If objBanco.GravaBanco(SQls) Then
                MsgBox(Me.Page, "Registro excluído com sucesso.", eTitulo.Sucess)
                Limpar()
                AtualizarGrid(False)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString, eTitulo.Erro)
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir registro.")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutoXConsumo")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class