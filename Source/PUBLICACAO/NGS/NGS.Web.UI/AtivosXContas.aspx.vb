Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AtivosXContas
    Inherits BasePage

#Region "Events"
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Patrimonio)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AtivosXContas", "ACESSAR") Then
                ddl.Carregar(DdlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeNegocio, DdlEmpresa)
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Patrimonio.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub DdlUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidadeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridAtivosXContas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridAtivosXContas.SelectedIndexChanged
        Try
            Dim ds As DataSet = getDataSet()
            txtConta.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Conta_Id")
            txtNomeConta.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Titulo")

            txtContaDepreciacaoDebito.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("DepreciacaoDebito")
            txtNomeDepreciacaoDebito.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Titulo_DepreDeb")
            txtContaDepreciacaoCredito.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("DepreciacaoCredito")
            txtNomeDepreciacaoCredito.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Titulo_DepreCre")

            txtContaDeBaixa.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("ContaDeBaixa")
            txtNomeContaDeBaixa.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Titulo_Baixa")

            txtContaTransferencia.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("ContaDeTransferencia")
            txtNomeContaTransferencia.Text = ds.Tables(0).Rows(GridAtivosXContas.SelectedIndex).Item("Titulo_Transferencia")

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgConta_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgConta.Click
        Try
            If GridAtivosXContas.Rows.Count > 0 AndAlso GridAtivosXContas.SelectedIndex > 0 Then
                MsgBox(Me.Page, "Conta do ítem selecionado não pode ser alterada")
            Else
                ucConsultaPlanoDeContas.Limpar()
                ucConsultaPlanoDeContas.BindGridView(True)
                Popup.ConsultaDePlanoDeContas(Me, "objAtivosXContas" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgLimparConta_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        txtConta.Text = String.Empty
        txtNomeConta.Text = String.Empty
    End Sub

    Protected Sub ImgDepDebito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objAtivosXContasDepreciacaoDebito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgLimparDepDebito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        txtContaDepreciacaoDebito.Text = String.Empty
        txtNomeDepreciacaoDebito.Text = String.Empty
    End Sub

    Protected Sub ImgDepCredito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objAtivosXContasDepreciacaoCredito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ImgLimparDepCredito_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        txtContaDepreciacaoCredito.Text = String.Empty
        txtNomeDepreciacaoCredito.Text = String.Empty
    End Sub

    Protected Sub imgContaDeBaixa_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgContaDeBaixa.Click
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objAtivoXContasDeBaixa" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparContaDeBaixa_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparContaDeBaixa.Click
        Try
            txtContaDeBaixa.Text = String.Empty
            txtNomeContaDeBaixa.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgContaTransferencia_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgContaTransferencia.Click
        Try
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me, "objAtivoXContasDeTransferencia" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgLimparContaTransferencia_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs) Handles imgLimparContaTransferencia.Click
        Try
            txtContaTransferencia.Text = String.Empty
            txtNomeContaTransferencia.Text = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("AtivosXContas", "GRAVAR") Then
                If ValidaCampos() Then
                    Dim sql As String = "INSERT INTO AtivosXContas (Empresa_Id, EndEmpresa_Id, Conta_Id, DepreciacaoDebito, DepreciacaoCredito, ContaDeBaixa, ContaDeTransferencia)" & vbCrLf & _
                                        "                    Values( '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'," & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                                        "                           ,'" & txtConta.Text & "','" & txtContaDepreciacaoDebito.Text & "','" & txtContaDepreciacaoCredito.Text & "'" & vbCrLf & _
                                        "                           ,'" & txtContaDeBaixa.Text & "', '" & txtContaTransferencia.Text & "')" & vbCrLf
                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
                        LimparContas()
                        AtualizarLista()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("AtivosXContas", "ALTERAR") Then
                If GridAtivosXContas.Rows.Count > 0 AndAlso GridAtivosXContas.SelectedIndex = -1 Then
                    MsgBox(Me.Page, "Selecione um Registro para alteração")
                ElseIf ValidaCampos() Then
                    Dim sql = " Update AtivosXContas " & vbCrLf & _
                              "    Set DepreciacaoDebito    = '" & txtContaDepreciacaoDebito.Text & "'," & vbCrLf & _
                              "        DepreciacaoCredito   = '" & txtContaDepreciacaoCredito.Text & "'," & vbCrLf & _
                              "        ContaDeBaixa         = '" & txtContaDeBaixa.Text & "'," & vbCrLf & _
                              "        ContaDeTransferencia = '" & txtContaTransferencia.Text & "'" & vbCrLf & _
                              "  Where Empresa_Id    = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                              "    and EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                              "    and Conta_Id      = '" & txtConta.Text & "'"

                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro atualizado com Sucesso.", eTitulo.Sucess)
                        LimparContas()
                        AtualizarLista()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Alterar Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("AtivosXContas", "EXCLUIR") Then
                If GridAtivosXContas.Rows.Count > 0 AndAlso GridAtivosXContas.SelectedIndex > 0 OrElse String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                    MsgBox(Me.Page, "Selecione um Registro para exclusão")
                Else
                    Dim sql As String = "   Delete AtivosXContas " & vbCrLf & _
                                        "    Where Empresa_Id    = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                                        "      and EndEmpresa_Id =  " & DdlEmpresa.SelectedValue.Split("-")(1) & vbCrLf & _
                                        "      and Conta_Id      = '" & txtConta.Text & "'"
                    If Banco.GravaBanco(sql) Then
                        MsgBox(Me.Page, "Registro excluído com Sucesso.", eTitulo.Sucess)
                        LimparCampos()
                        AtualizarLista()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para Excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("AtivosXContas", "RELATORIO") Then

            Else
                MsgBox(Me.Page, "Usuário sem permissão para Emitir relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "AtivosXContas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            LimparContas()
            AtualizarLista()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub LimparContas()
        txtConta.Text = String.Empty
        txtNomeConta.Text = String.Empty
        txtContaDepreciacaoDebito.Text = String.Empty
        txtNomeDepreciacaoDebito.Text = String.Empty
        txtContaDepreciacaoCredito.Text = String.Empty
        txtNomeDepreciacaoCredito.Text = String.Empty
        txtContaDeBaixa.Text = String.Empty
        txtNomeContaDeBaixa.Text = String.Empty
        txtContaTransferencia.Text = String.Empty
        txtNomeContaTransferencia.Text = String.Empty
    End Sub

    Private Sub LimparCampos()
        txtConta.Text = String.Empty
        txtNomeConta.Text = String.Empty
        txtContaDepreciacaoDebito.Text = String.Empty
        txtNomeDepreciacaoDebito.Text = String.Empty
        txtContaDepreciacaoCredito.Text = String.Empty
        txtNomeDepreciacaoCredito.Text = String.Empty
        txtContaDeBaixa.Text = String.Empty
        txtNomeContaDeBaixa.Text = String.Empty
        txtContaTransferencia.Text = String.Empty
        txtNomeContaTransferencia.Text = String.Empty

        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeNegocio, DdlEmpresa)
        GridAtivosXContas.DataBind()

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        HID.Value = Guid.NewGuid().ToString
        ucConsultaPlanoDeContas.SetarHID(HID.Value)

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeNegocio.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "SELECT axc.Conta_Id, axc.DepreciacaoDebito, axc.DepreciacaoCredito," & vbCrLf & _
                            "       isnull(axc.ContaDeBaixa, '') as ContaDeBaixa, isnull(axc.ContaDeTransferencia, '') as ContaDeTransferencia," & vbCrLf & _
                            "       ISNULL((SELECT Titulo from PlanoDeContas Where Conta_Id = axc.Conta_Id), '') As Titulo, " & vbCrLf & _
                            "       ISNULL((SELECT Titulo from PlanoDeContas Where Conta_Id = LTRIM(axc.DepreciacaoDebito)),'') As Titulo_DepreDeb, " & vbCrLf & _
                            "       ISNULL((SELECT Titulo from PlanoDeContas Where Conta_Id = LTRIM(axc.DepreciacaoCredito)),'') As Titulo_DepreCre, " & vbCrLf & _
                            "       ISNULL((SELECT Titulo from PlanoDeContas Where Conta_Id = LTRIM(axc.ContaDeBaixa)),'') As Titulo_Baixa,                 " & vbCrLf & _
                            "       ISNULL((SELECT Titulo from PlanoDeContas Where Conta_Id = LTRIM(axc.ContaDeTransferencia)),'') As Titulo_Transferencia  " & vbCrLf & _
                            "  FROM AtivosXContas axc" & vbCrLf & _
                            " WHERE (axc.Empresa_Id    = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "') " & vbCrLf & _
                            "   AND (axc.EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1) & ")" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXContas")

        Return ds
    End Function

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        Try
            If Session("objAtivosXContas" & HID.Value) IsNot Nothing Then
                txtConta.Text = CType(HttpContext.Current.Session("objAtivosXContas" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeConta.Text = CType(HttpContext.Current.Session("objAtivosXContas" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
                Session.Remove("objAtivosXContas" & HID.Value)
            ElseIf Session("objAtivosXContasDepreciacaoDebito" & HID.Value) IsNot Nothing Then
                txtContaDepreciacaoDebito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoDebito" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeDepreciacaoDebito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoDebito" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
                Session.Remove("objAtivosXContasDepreciacaoDebito" & HID.Value)
            ElseIf Session("objAtivosXContasDepreciacaoCredito" & HID.Value) IsNot Nothing Then
                txtContaDepreciacaoCredito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoCredito" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeDepreciacaoCredito.Text = CType(HttpContext.Current.Session("objAtivosXContasDepreciacaoCredito" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
                Session.Remove("objAtivosXContasDepreciacaoCredito" & HID.Value)
            ElseIf Not Session("objAtivoXContasDeBaixa" & HID.Value) Is Nothing Then
                txtContaDeBaixa.Text = CType(HttpContext.Current.Session("objAtivoXContasDeBaixa" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeContaDeBaixa.Text = CType(HttpContext.Current.Session("objAtivoXContasDeBaixa" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
                Session.Remove("objAtivoXContasDeBaixa" & HID.Value)
            ElseIf Not Session("objAtivoXContasDeTransferencia" & HID.Value) Is Nothing Then
                txtContaTransferencia.Text = CType(HttpContext.Current.Session("objAtivoXContasDeTransferencia" & HID.Value), [Lib].Negocio.PlanoDeConta).Conta
                txtNomeContaTransferencia.Text = CType(HttpContext.Current.Session("objAtivoXContasDeTransferencia" & HID.Value), [Lib].Negocio.PlanoDeConta).Titulo
                Session.Remove("objAtivoXContasDeTransferencia" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizarLista()
        Dim ds As DataSet = getDataSet()
        GridAtivosXContas.DataSource = ds
        GridAtivosXContas.DataBind()
    End Sub

    Private Function IsContaGrupoEContemCliente() As Boolean
        Dim sql As String = "select Conta_Id from PlanoDeContas where LEN(conta_id) = 7 and Cliente = 'S' and Conta_Id = '" & txtContaTransferencia.Text & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ContaGrupo")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Empresa não foi selecionada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtConta.Text) Then
            MsgBox(Me.Page, "Conta não foi selecionada.")
            Return False
        ElseIf Left(txtConta.Text, 1) <> 1 Then
            MsgBox(Me.Page, "Conta não foi selecionada não é uma conta do Ativo.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtContaDepreciacaoDebito.Text) AndAlso Left(txtContaDepreciacaoDebito.Text, 1) <= 2 Then
            MsgBox(Me.Page, "Depreciação Débito deve ter a primeira posição maior que 2.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtContaDepreciacaoCredito.Text) AndAlso Left(txtContaDepreciacaoCredito.Text, 7) <> Left(txtConta.Text, 7) Then
            MsgBox(Me.Page, "Depreciação Crédito deve conter as 7 primeiras posições igual as 7 primeiras posições da conta.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtContaDeBaixa.Text) AndAlso Left(txtContaDeBaixa.Text, 1) <= 2 Then
            MsgBox(Me.Page, "Conta de Baixa deve ter a primeira posição maior que 2.")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtContaTransferencia.Text) AndAlso Left(txtContaTransferencia.Text, 1) <> 1 Then
            MsgBox(Me.Page, "Conta de Transferência não é uma conta do Ativo")
            Return False
        ElseIf Not String.IsNullOrWhiteSpace(txtContaTransferencia.Text) AndAlso Not IsContaGrupoEContemCliente() Then
            MsgBox(Me.Page, "Conta de Transferência deve ser uma conta de grupo que contenha Clientes")
            Return False
        End If
        Return True
    End Function

#End Region

End Class