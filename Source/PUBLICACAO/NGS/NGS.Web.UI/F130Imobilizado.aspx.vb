Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class F130Imobilizado
    Inherits BasePage

    Private lstF130 As [Lib].Negocio.lstPisCofinsRegistrosF130

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("F130Imobilizado", "ACESSAR") Then
                ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio)
                Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa)
                ddl.Carregar(ddlCstPisCofins, CarregarDDL.Tabela.SituacaoTributariaPISCOFINS, "")
                CarregarGridF130Imobilizado()
                CargaCentroDeCusto()
                limpar()
                txtMovimentoId.Text = Format(Today, "dd/MM/yyyy")
                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
            End If
        End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            CarregarEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If txtDataInicial.Text = String.Empty OrElse txtDataFinal.Text = String.Empty Then
                MsgBox(Me.Page, "Período não está correto para consulta.", eTitulo.Erro)
                Exit Sub
            End If

            lstF130 = New lstPisCofinsRegistrosF130(ddlEmpresa.SelectedValue.ToString.Split("-")(0), ddlEmpresa.SelectedValue.ToString.Split("-")(1),
                                                    txtDataInicial.Text, txtDataFinal.Text)
            If lstF130 IsNot Nothing AndAlso lstF130.Count > 0 Then
                gridF130Imobilizado.DataSource = lstF130
                gridF130Imobilizado.DataBind()
                TabContainer1.ActiveTabIndex = 1
            Else
                MsgBox(Me.Page, "Não foram encontrados registros no período determinado.", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("F130Imobilizado", "ACESSAR") Then
                If ValidarCampos() Then

                    Dim objF130 As New PisCofinsRegistroF130()
                    Dim MesOpeAquis = txtMovimentoId.Text.Split("/")(1) & txtMovimentoId.Text.Split("/")(2)

                    objF130.EmpresaId = ddlEmpresa.SelectedValue.ToString.Split("-")(0)
                    objF130.MovimentoId = CDate(Trim(txtMovimentoId.Text))
                    objF130.Situacao = Trim(txtSituacao.Text)
                    objF130.NatBcCred = Trim(txtNatBcCred.Text)
                    objF130.IdentBemImob = Trim(ddlIdentBemImob.SelectedValue)
                    objF130.IndOrigCred = Trim(ddlIndOrigCred.SelectedValue)
                    objF130.IndUtilBemImob = Trim(ddlIndUtilBemImob.SelectedValue)
                    objF130.MesOpeAquis = Trim(txtMovimentoId.Text.Split("/")(1) & txtMovimentoId.Text.Split("/")(2))
                    objF130.ParcOperNaoBcCred = Replace(Trim(txtParcOperNaoBcCred.Text), ",", ".")

                    objF130.VlOperAquis = Replace(Replace(txtValor.Text, ".", ""), ",", ".")

                    objF130.VlBcCred = Replace(Replace(txtVlBcCred.Text, ".", ""), ",", ".")

                    objF130.Cst = Trim(ddlCstPisCofins.SelectedValue.Split("-")(0))

                    objF130.AliqPis = Replace(Replace(txtAliqPis.Text, ".", ""), ",", ".")
                    objF130.VlPis = Replace(Replace(txtVlPis.Text, ".", ""), ",", ".")

                    objF130.AliqCofins = Replace(Replace(txtAliqCofins.Text, ".", ""), ",", ".")
                    objF130.VlCofins = Replace(Replace(txtVlCofins.Text, ".", ""), ",", ".")

                    objF130.CodCta = Trim(txtConta.Text.Split("-")(0))

                    If Not String.IsNullOrWhiteSpace(ddlCodCcus.SelectedValue) Then
                        objF130.CodCcus = Trim(ddlCodCcus.SelectedValue)
                    End If
                    If Not String.IsNullOrWhiteSpace(txtDescBemImob.Text) Then
                        objF130.DescBemImob = Trim(txtDescBemImob.Text)
                    End If

                    objF130.IUD = "I"
                    If objF130.Salvar Then
                        MsgBox(Me.Page, "Informação inserida com sucesso.", eTitulo.Sucess)
                        limpar()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.", "~/Producao.aspx")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("F130Imobilizado", "ACESSAR") Then
                If ValidarCampos() Then
                    Dim objF130 As New PisCofinsRegistroF130()

                    objF130.EmpresaId = ddlEmpresa.SelectedValue.ToString.Split("-")(0)
                    objF130.MovimentoId = txtMovimentoId.Text

                    If txtNatBcCred.Text <> "" Then
                        objF130.NatBcCred = Trim(txtNatBcCred.Text)
                    End If
                    If ddlIdentBemImob.SelectedValue IsNot Nothing Then
                        objF130.IdentBemImob = Trim(ddlIdentBemImob.SelectedValue)
                    End If
                    If ddlIndOrigCred.SelectedValue IsNot Nothing Then
                        objF130.IndOrigCred = Trim(ddlIndOrigCred.SelectedValue)
                    End If
                    If ddlIndUtilBemImob.SelectedValue IsNot Nothing Then
                        objF130.IndUtilBemImob = Trim(ddlIndUtilBemImob.SelectedValue)
                    End If
                    If txtParcOperNaoBcCred.Text <> "" Then
                        objF130.ParcOperNaoBcCred = Replace(Trim(txtParcOperNaoBcCred.Text), ",", ".")
                    End If
                    If ddlCstPisCofins.SelectedValue IsNot Nothing Then
                        objF130.Cst = Trim(ddlCstPisCofins.SelectedValue.Split("-")(0))
                    End If
                    If txtAliqPis.Text <> "" Then
                        objF130.AliqPis = Replace(Trim(txtAliqPis.Text), ",", ".")
                    End If
                    If txtVlPis.Text <> "" Then
                        objF130.VlPis = Replace(Trim(txtVlPis.Text), ",", ".")
                    End If
                    If txtAliqCofins.Text <> "" Then
                        objF130.AliqCofins = Replace(Trim(txtAliqCofins.Text), ",", ".")
                    End If
                    If txtVlCofins.Text <> "" Then
                        objF130.VlCofins = Replace(txtVlCofins.Text, ",", ".")
                    End If
                    If ddlCodCcus.SelectedValue IsNot Nothing Then
                        objF130.CodCcus = Trim(ddlCodCcus.SelectedValue)
                    End If
                    If txtDescBemImob.Text <> "" Then
                        objF130.DescBemImob = Trim(txtDescBemImob.Text)
                    End If

                    objF130.VlOperAquis = Replace(Replace(txtValor.Text, ".", ""), ",", ".")

                    objF130.VlBcCred = Replace(Replace(txtVlBcCred.Text, ".", ""), ",", ".")

                    objF130.IUD = "U"
                    If objF130.Salvar Then
                        MsgBox(Me.Page, "Informação atualizada com sucesso.", eTitulo.Sucess)
                        limpar()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", "~/Producao.aspx")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("F130Imobilizado", "ACESSAR") Then
                If ValidarCampos() Then
                    Dim objF130 = New PisCofinsRegistroF130()

                    objF130.EmpresaId = ddlEmpresa.SelectedValue.ToString.Split("-")(0)
                    objF130.MovimentoId = txtMovimentoId.Text
                    objF130.CodCta = txtConta.Text
                    objF130.IUD = "D"
                    If objF130.Salvar Then
                        MsgBox(Me.Page, "Informação excluída com sucesso.", eTitulo.Sucess)
                        limpar()
                    End If
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
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtaliqPis_TextChanged(sender As Object, e As EventArgs) Handles txtAliqPis.TextChanged
        txtVlPis.Text = CDec(txtVlBcCred.Text) * CDec(txtAliqPis.Text)
    End Sub

    Protected Sub txtaliqCofins_TextChanged(sender As Object, e As EventArgs) Handles txtAliqCofins.TextChanged
        txtVlCofins.Text = CDec(txtVlBcCred.Text) * CDec(txtAliqCofins.Text)
    End Sub

    Private Sub CarregarEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Private Sub CargaCentroDeCusto()
        ddl.Carregar(ddlCodCcus, CarregarDDL.Tabela.CentroDeCusto, "Len(CentroDeCusto_Id) = 5", True)
    End Sub

    Private Sub limpar()
        Session.Remove("objContaEx" & HID.Value)

        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        ddlEmpresa.Enabled = False
        ddlUnidadeDeNegocio.Enabled = False

        txtMovimentoId.Enabled = True
        btnConsultaConta.Enabled = True

        txtSituacao.Text = String.Empty
        txtValor.Text = String.Empty
        txtNatBcCred.Text = String.Empty
        ddlIdentBemImob.SelectedValue = String.Empty
        ddlIndOrigCred.SelectedValue = String.Empty
        ddlIndUtilBemImob.SelectedValue = String.Empty
        txtParcOperNaoBcCred.Text = "0,00"
        ddlCstPisCofins.SelectedValue = String.Empty
        txtAliqPis.Text = "1,65"
        txtVlPis.Text = String.Empty
        txtVlBcCred.Text = String.Empty
        txtAliqCofins.Text = "7,60"
        txtVlCofins.Text = String.Empty
        txtConta.Text = String.Empty
        ddlCodCcus.SelectedValue = String.Empty
        txtDescBemImob.Text = String.Empty
        txtSituacao.Text = 1

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPlanoDeContas.SetarHID(HID.Value)
        CarregarGridF130Imobilizado()
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtMovimentoId.Text) Then
            MsgBox(Me.Page, "Campo movimento não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Campo empresa não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtConta.Text) Then
            MsgBox(Me.Page, "Campo Conta não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtnatBcCred.Text) Then
            MsgBox(Me.Page, "Campo natBcCred não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlIdentBemImob.Text) Then
            MsgBox(Me.Page, "Campo IdentBemImob não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlIndOrigCred.Text) Then
            MsgBox(Me.Page, "Campo IndOrigCred não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlIndUtilBemImob.Text) Then
            MsgBox(Me.Page, "Campo IndUtilBemImob não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtparcOperNaoBcCred.Text) Then
            MsgBox(Me.Page, "Campo parcOperNaoBcCred não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtaliqPis.Text) Then
            MsgBox(Me.Page, "Campo aliqPis não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtvlPis.Text) Then
            MsgBox(Me.Page, "Campo vlPis não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlCstPisCofins.SelectedValue) Then
            MsgBox(Me.Page, "Campo CstPisCofins não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtVlBcCred.Text) Then
            MsgBox(Me.Page, "Campo txtVlBcCred não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtaliqCofins.Text) Then
            MsgBox(Me.Page, "Campo aliqCofins não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtvlCofins.Text) Then
            MsgBox(Me.Page, "Campo vlCofins não foi informado!", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Private Sub CarregarGridF130Imobilizado()
        Try
            lstF130 = New lstPisCofinsRegistrosF130()

            gridF130Imobilizado.DataSource = lstF130
            gridF130Imobilizado.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        Try
            If Session("objContaEx" & HID.Value) IsNot Nothing Then
                Dim objConta As [Lib].Negocio.PlanoDeConta = CType(Session("objContaEx" & HID.Value), [Lib].Negocio.PlanoDeConta)

                If objConta.Conta.Length = 9 Then
                    hdfConta.Value = objConta.Conta
                    txtConta.Text = objConta.Conta & " - " & objConta.Titulo
                Else
                    hdfConta.Value = String.Empty
                    txtConta.Text = String.Empty
                    MsgBox(Me.Page, "É obrigatório uma conta analítica.")
                End If
                Session.Remove("objContaEx" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridF130Imobilizado_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridF130Imobilizado.SelectedIndexChanged
        Try
            CarregarDadosF130Imobilizado()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarDadosF130Imobilizado()
        Try
            txtMovimentoId.Text = gridF130Imobilizado.SelectedRow.Cells(1).Text()
            txtConta.Text = gridF130Imobilizado.SelectedRow.Cells(2).Text()
            If gridF130Imobilizado.SelectedRow.Cells(3).Text() = "1" Then
                txtSituacao.Text = "NORMAL"
            End If
            txtNatBcCred.Text = gridF130Imobilizado.SelectedRow.Cells(4).Text()
            ddlIdentBemImob.Text = gridF130Imobilizado.SelectedRow.Cells(5).Text()
            ddlIndOrigCred.Text = gridF130Imobilizado.SelectedRow.Cells(6).Text()
            ddlIndUtilBemImob.Text = gridF130Imobilizado.SelectedRow.Cells(7).Text()
            'MesOpeAquis
            'ddlMesOpeAquis.Text = gridF130Imobilizado.SelectedRow.Cells(6).Text()
            'VlOperAquis
            'ddlVlOperAquis.Text = gridF130Imobilizado.SelectedRow.Cells(7).Text()
            txtParcOperNaoBcCred.Text = gridF130Imobilizado.SelectedRow.Cells(10).Text()
            'VlBcCred
            'txtVlBcCred.Text = gridF130Imobilizado.SelectedRow.Cells(9).Text()
            'IndNrParc
            'txtIndNrParc.Text = gridF130Imobilizado.SelectedRow.Cells(10).Text()
            ddlCstPisCofins.SelectedValue = gridF130Imobilizado.SelectedRow.Cells(13).Text()
            'VlBcPis
            'txtVlBcPis.Text = gridF130Imobilizado.SelectedRow.Cells(12).Text()
            txtAliqPis.Text = gridF130Imobilizado.SelectedRow.Cells(15).Text()
            txtVlPis.Text = gridF130Imobilizado.SelectedRow.Cells(16).Text()
            txtVlBcCred.Text = gridF130Imobilizado.SelectedRow.Cells(18).Text()
            txtAliqCofins.Text = gridF130Imobilizado.SelectedRow.Cells(19).Text()
            txtVlCofins.Text = gridF130Imobilizado.SelectedRow.Cells(20).Text()
            ddlCodCcus.SelectedValue = gridF130Imobilizado.SelectedRow.Cells(21).Text()
            txtDescBemImob.Text = gridF130Imobilizado.SelectedRow.Cells(22).Text()

            gridF130Imobilizado.DataBind()
            TabContainer1.ActiveTabIndex = 0

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtMovimentoId.Enabled = False
            btnConsultaConta.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaConta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultaConta.Click
        Try
            Dim txtConta As TextBox = ucConsultaPlanoDeContas.FindControl("txtConta")
            If txtConta IsNot Nothing Then
                txtConta.Text = String.Empty
                txtConta.Focus()
            End If

            Session("TipoConta" & HID.Value) = "Credito"
            Session("Codigo" & HID.Value) = hdfConta.ClientID
            Session("Descricao" & HID.Value) = txtConta.ClientID
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objContaEx" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class