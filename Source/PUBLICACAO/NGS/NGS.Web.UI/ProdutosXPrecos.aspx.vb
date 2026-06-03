Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ProdutosXPrecos
    Inherits BasePage

    Private Sql As String
    Private Empresa() As String
    Private Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Revenda)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PRODUTOSXPRECOS", "ACESSAR") Then
                    CargaGrupo()
                    ddl.Carregar(ddlTabelaDePreco, CarregarDDL.Tabela.TabelaDePreco, " ATIVO = 1 ", True)
                    ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "", True)
                    Limpar(True)
                    txtMovimento.Text = Now.Date
                    txtData.Text = Now.Date
                    ddl.Carregar(ddlGrupoPreco, CarregarDDL.Tabela.NivelGrupoProduto, "len(Grupo_Id) = 2")
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoPxP" & HID.Value) IsNot Nothing Then
            Dim objProduto As Produto = CType(obj, Produto)
            ddlGrupo.SelectedValue = objProduto.CodigoGrupo
            CargaProduto()
            ddlProduto.SelectedValue = objProduto.Codigo
            Session.Remove("objProdutoPxP" & HID.Value)
            validarMoedas()
        End If
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            txtCodigoCliente.Value = CType(Session("objCliente" & HID.Value), Cliente).Codigo & "-" & CType(Session("objCliente" & HID.Value), Cliente).CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(CType(Session("objCliente" & HID.Value), Cliente).Codigo) & " - " & CType(Session("objCliente" & HID.Value), Cliente).Nome
            Session.Remove("objCliente" & HID.Value)

            CarregarListaProdutos(ddlTabelaDePreco.SelectedValue.ToString(), txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1), "", "", "")
        End If
    End Sub

    Private Sub CargaGrupo()
        Dim lstGrupoProduto As New ListGrupoProduto(True, "", "", "")
        Dim strGrupoProduto As String

        ddlGrupo.Items.Clear()

        For Each dr As GrupoProduto In lstGrupoProduto
            strGrupoProduto = Funcoes.AlinharEsquerda(dr.Descricao, 50, ".") & " - " & dr.Codigo
            ddlGrupo.Items.Add(New ListItem(strGrupoProduto, dr.Codigo))
        Next

        Funcoes.InserirLinhaEmBranco(ddlGrupo)
    End Sub

    Private Sub CargaProduto()
        Dim lstProduto As New ListProduto(ddlGrupo.SelectedValue)
        Dim strProduto As String
        ddlProduto.Items.Clear()
        For Each dr As Produto In lstProduto
            strProduto = Funcoes.AlinharEsquerda(dr.Nome, 50, ".") & " - " & dr.Codigo
            ddlProduto.Items.Add(New ListItem(strProduto, dr.Codigo.ToString()))
        Next
        Funcoes.InserirLinhaEmBranco(ddlProduto)
    End Sub

    Private Sub CargaValorDolar()
        Dim objBanco As New AcessaBanco()

        Dim sql As String = "SELECT indice FROM Cotacoes " & vbCrLf &
                            "WHERE Indexador_Id = 3 " & vbCrLf &
                            "AND Data_id = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf
        Dim dsValorDolar As DataSet = objBanco.ConsultaDataSet(sql, "Cotacoes")

        For Each row As ProdutoXPreco In CType(Session("ListaPxP" & HID.Value), ListProdutoXPreco)
            If Not row.CodigoMoeda = 1 Then
                For Each dolar As DataRow In dsValorDolar.Tables(0).Rows
                    row.Valor = row.Valor * dolar("indice")
                Next
            End If
        Next

    End Sub

    Private Sub Limpar(ByVal fLimpar As Boolean)
        lnkConsultar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        btnBuscaCliente.Enabled = False
        btnProduto.Enabled = True

        ddlGrupo.Enabled = True
        ddlProduto.Enabled = True
        ddlMoeda.Enabled = False
        ddlGrupo.SelectedIndex = 0
        ddlProduto.Items.Clear()
        ddlMoeda.SelectedIndex = 1

        If fLimpar Then
            ddlTabelaDePreco.SelectedIndex = 0
            txtCliente.Text = ""
            txtCodigoCliente.Value = ""
            btnBuscaCliente.Enabled = True
        End If

        txtMovimento.Enabled = True
        TxtValor.Enabled = False
        TxtValor.Text = ""
        txtDolar.Enabled = False
        txtDolar.Text = ""
        txtFixoOperacional.Text = "0"
        txtMargemMenor.Text = "0"
        txtValorMargemMenor.Text = ""
        txtMargemMaior.Text = "0"
        txtValorMargemMaior.Text = ""

        lblHistoricoProduto.Text = "Histórico do Produto: "
        'TxtValor.Text = "1"
        'txtFixoOperacional.Text = "1,3"
        'txtMargemMenor.Text = "015,00"
        'txtMargemMaior.Text = "020,00"

        gridPxP.DataSource = Nothing
        gridPxP.DataBind()
        Session.Remove("ListaPxP" & HID.Value)

        gridHistorico.DataSource = Nothing
        gridHistorico.DataBind()
        Session.Remove("ListaPxPHistorico" & HID.Value)
    End Sub

    Private Sub Carregar_Grid(ByVal Produto As String, ByVal Data As Date)
        'Dim Lista As [Lib].Negocio.ListProdutosXPrecos

        'If Produto.Length > 0 Then
        '    Lista = New [Lib].Negocio.ListProdutosXPrecos(Produto, Nothing, "Data_Id")
        'Else
        '    Lista = New [Lib].Negocio.ListProdutosXPrecos(Data)
        'End If

        'gridPxP.DataSource = Lista.ToArray()
        'gridPxP.DataBind()
        'Session("ListaPxP") = Lista
    End Sub

    Private Sub IniciarPxP(ByVal Tipo As String)
        If ValidarCampos() Then

            'If CDate(txtMovimento.Text) < Date.Now Then
            '    MsgBox(Me.Page, "Não Permitido em data Retroativa")
            '    Exit Sub
            'End If

            If Tipo = "I" And Funcoes.VerificaPermissao("ProdutosXPrecos", "GRAVAR") Or
               Tipo = "U" And Funcoes.VerificaPermissao("ProdutosXPrecos", "ALTERAR") Or
               Tipo = "D" And Funcoes.VerificaPermissao("ProdutosXPrecos", "EXCLUIR") Then

                'If Tipo = "I" And CDate(txtMovimento.Text) < Now.Date Then
                '    MsgBox(Me.Page, "Não é Permitido a Entrada de Valores com dara retroativa")
                '    Exit Sub
                'End If

                Dim PxP As New ProdutoXPreco()
                PxP.CodigoTabelaDePreco = ddlTabelaDePreco.SelectedValue
                PxP.CodigoProduto = ddlProduto.SelectedValue
                PxP.CodigoCliente = txtCodigoCliente.Value.Split("-")(0)
                PxP.EnderecoCliente = txtCodigoCliente.Value.Split("-")(1)
                PxP.Movimento = CDate(txtMovimento.Text)
                PxP.CodigoMoeda = ddlMoeda.SelectedValue
                If ddlMoeda.SelectedValue = "1" Then
                    PxP.Valor = TxtValor.Text
                Else
                    PxP.Valor = txtDolar.Text
                End If
                PxP.FixoOperacional = txtFixoOperacional.Text
                PxP.MargemMenor = txtMargemMenor.Text
                PxP.MargemMaior = txtMargemMaior.Text
                PxP.IUD = Tipo
                If PxP.Salvar() Then
                    Limpar(False)
                    Carregar_Grid("", CDate(txtData.Text))
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Produto incluido!")
                        Case "U"
                            MsgBox(Me.Page, "Produto altualizado!")
                        Case "D"
                            MsgBox(Me.Page, "Produto excluido!")
                    End Select
                Else
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Erro ao Incluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "U"
                            MsgBox(Me.Page, "Erro ao Alterar Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                        Case "D"
                            MsgBox(Me.Page, "Erro ao Excluir Registro: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                    End Select
                End If
            Else
                Select Case Tipo
                    Case "I"
                        MsgBox(Me.Page, "Usuário sem permissão para incluir registro.", eTitulo.Info)
                    Case "U"
                        MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
                    Case "D"
                        MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
                End Select
            End If

            CarregarListaProdutos(ddlTabelaDePreco.SelectedValue.ToString(), txtCodigoCliente.Value.Split("-")(0), txtCodigoCliente.Value.Split("-")(1), ddlProduto.SelectedValue.ToString(), CDate(txtData.Text).ToString("yyyy-MM-dd"), "")
        Else
            MsgBox(Me.Page, Mensagem)
        End If
    End Sub

    Private Sub CarregarListaProdutos(Optional ByVal tabelaDePreco As String = "", Optional ByVal codigoCliente As String = "", Optional ByVal enderecoCliente As String = "", Optional ByVal Produto As String = "", Optional ByVal Movimento As String = "", Optional ByVal pOrderby As String = "")

        Dim Lista As ListProdutoXPreco = New ListProdutoXPreco(tabelaDePreco, codigoCliente, enderecoCliente, Produto, Movimento, pOrderby)

        Session("ListaPxP") = Lista
        CargaValorDolar()

        gridPxP.DataSource = Lista.ToArray()
        gridPxP.DataBind()

        TabContainer1.ActiveTabIndex = 1
    End Sub

    Function ValidarCampos() As Boolean
        If ddlTabelaDePreco.SelectedValue = "" Then
            Mensagem = "Tabela de Preços não foi selecionada."
            Return False
        ElseIf txtCliente.Text = "" Then
        Mensagem = "Cliente não foi selecionado."
            Return False
        ElseIf ddlProduto.SelectedValue = "" Then
            Mensagem = "Produto não foi informado."
            Return False
            'ElseIf TxtValor.Text.Length = 0 Then
            '    Mensagem = "Valor oficial não foi informado."
            '    Return False
        Else
            Return True
        End If
    End Function

    Function validarMoedas() As Boolean
        ddlMoeda.Enabled = True
        If ddlMoeda.SelectedValue = 1 Then
            TxtValor.Enabled = True
            txtDolar.Enabled = False
            txtDolar.Text = ""
        Else
            txtDolar.Enabled = True
            TxtValor.Enabled = False
            TxtValor.Text = ""
        End If
    End Function

    Function validaRegras() As Boolean
        lnkConsultar.Parent.Visible = False
        lnkNovo.Parent.Visible = True
        btnBuscaCliente.Enabled = False
        btnProduto.Enabled = False
        ddlGrupo.Enabled = False
        ddlProduto.Enabled = False
        txtMovimento.Enabled = False
    End Function

    Protected Sub gridPxP_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim PxP As ProdutoXPreco = New ProdutoXPreco(CType(Session("ListaPxP"), ListProdutoXPreco)(gridPxP.SelectedIndex).CodigoTabelaDePreco, CType(Session("ListaPxP"), ListProdutoXPreco)(gridPxP.SelectedIndex).CodigoCliente, CType(Session("ListaPxP"), [Lib].Negocio.ListProdutoXPreco)(gridPxP.SelectedIndex).EnderecoCliente, CType(Session("ListaPxP"), [Lib].Negocio.ListProdutoXPreco)(gridPxP.SelectedIndex).CodigoProduto, CType(Session("ListaPxP"), [Lib].Negocio.ListProdutoXPreco)(gridPxP.SelectedIndex).Movimento.ToString("yyyy-MM-dd"))

            lblHistoricoProduto.Text = "Histórico do Produto: " & PxP.CodigoProduto & " - " & PxP.NomeProduto

            Dim Lista As ListProdutoXPreco = New ListProdutoXPreco(PxP.CodigoTabelaDePreco, PxP.CodigoCliente, PxP.EnderecoCliente, PxP.CodigoProduto, "", "")

            Session("ListaPxPHistorico") = Lista
            CargaValorDolar()

            gridHistorico.DataSource = Lista.ToArray()
            gridHistorico.DataBind()

            '******** Carrega Cadastro  ******************
            validarMoedas()
            validaRegras()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            ddlGrupo.Enabled = False
            ddlProduto.Enabled = False
            txtMovimento.Enabled = False
            ddlMoeda.Enabled = False

            Dim pCliente As Cliente = New Cliente(PxP.CodigoCliente, PxP.EnderecoCliente)
            txtCodigoCliente.Value = pCliente.Codigo & "-" & pCliente.CodigoEndereco
            txtCliente.Text = Funcoes.FormatarCpfCnpj(pCliente.Codigo) & " - " & pCliente.Nome

            ddlGrupo.SelectedValue = PxP.Produto.CodigoGrupo
            CargaProduto()
            ddlProduto.SelectedValue = PxP.CodigoProduto
            txtMovimento.Text = PxP.Movimento
            ddlMoeda.SelectedValue = PxP.CodigoMoeda

            If ddlMoeda.SelectedValue = 1 Then
                TxtValor.Text = PxP.Valor
                txtDolar.Text = String.Empty

            Else
                TxtValor.Text = String.Empty
                txtDolar.Text = PxP.Valor
            End If
            txtFixoOperacional.Text = PxP.FixoOperacional
            txtMargemMenor.Text = PxP.MargemMenor
            txtMargemMaior.Text = PxP.MargemMaior
            AtualizaValorMargemMenor()
            AtualizaValorMargemMaior()
            TabContainer1.ActiveTabIndex = 2
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlTabelaDePreco_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlTabelaDePreco.SelectedValue.Length > 0 Then
                'CarregarListaProdutos(ddlTabelaDePreco.SelectedValue.ToString(), "", "", "", "", "")
            Else
                Limpar(True)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            validarMoedas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMoeda_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            validarMoedas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridHistorico_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            lnkNovo.Enabled = True
            lnkAtualizar.Enabled = True
            lnkExcluir.Enabled = True
            ddlGrupo.Enabled = False
            ddlProduto.Enabled = False
            'txtMovimento.Enabled = False

            Dim PxP As ProdutoXPreco = CType(Session("ListaPxPHistorico"), ListProdutoXPreco)(gridHistorico.SelectedIndex)

            '******** Carrega Cadastro  ******************
            lnkNovo.Enabled = True
            lnkAtualizar.Enabled = True
            lnkExcluir.Enabled = True
            ddlGrupo.Enabled = False
            ddlProduto.Enabled = False
            txtMovimento.Enabled = False
            ddlGrupo.SelectedValue = PxP.Produto.CodigoGrupo
            CargaProduto()
            ddlProduto.SelectedValue = PxP.CodigoProduto
            txtMovimento.Text = PxP.Movimento
            ddlMoeda.SelectedValue = PxP.CodigoMoeda

            If ddlMoeda.SelectedValue = 1 Then
                TxtValor.Text = PxP.Valor
                txtDolar.Text = String.Empty

            Else
                TxtValor.Text = String.Empty
                txtDolar.Text = PxP.Valor
            End If
            txtFixoOperacional.Text = PxP.FixoOperacional
            txtMargemMenor.Text = PxP.MargemMenor
            txtMargemMaior.Text = PxP.MargemMaior
            AtualizaValorMargemMenor()
            AtualizaValorMargemMaior()
            TabContainer1.ActiveTabIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtDataTabela_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Carregar_Grid("", CDate(txtData.Text))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridPxP_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        Try
            gridPxP.PageIndex = e.NewPageIndex
            gridPxP.DataSource = CType(Session("ListaPxP"), ListProdutoXPreco)
            gridPxP.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtMargemMenor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            AtualizaValorMargemMenor()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtMargemMaior_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            AtualizaValorMargemMaior()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub AtualizaValorMargemMenor()
        Dim Valor As Decimal = CDec(IIf(IsNumeric(TxtValor.Text), TxtValor.Text, "0"))
        Dim FixoOperacional As Decimal = 1 + CDec(IIf(IsNumeric(txtFixoOperacional.Text), txtFixoOperacional.Text, "0")) / 100
        Dim Margem As Decimal = 1 + CDec(IIf(IsNumeric(txtMargemMenor.Text), txtMargemMenor.Text, "0")) / 100

        txtValorMargemMenor.Text = Math.Round((Valor * FixoOperacional) * Margem, 2)
    End Sub

    Private Sub AtualizaValorMargemMaior()
        Dim Valor As Decimal = CDec(IIf(IsNumeric(TxtValor.Text), TxtValor.Text, "0"))
        Dim FixoOperacional As Decimal = 1 + CDec(IIf(IsNumeric(txtFixoOperacional.Text), txtFixoOperacional.Text, "0")) / 100
        Dim Margem As Decimal = 1 + CDec(IIf(IsNumeric(txtMargemMaior.Text), txtMargemMaior.Text, "0")) / 100

        txtValorMargemMaior.Text = Math.Round((Valor * FixoOperacional) * Margem, 2)
    End Sub

    Protected Sub txtFixoOperacional_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            AtualizaValorMargemMenor()
            AtualizaValorMargemMaior()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub TxtValor_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            AtualizaValorMargemMenor()
            AtualizaValorMargemMaior()
            validaRegras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub TxtDolar_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ValidarCampos() Then
                validaRegras()
                lnkConsultar.Parent.Visible = False
            Else
                MsgBox(Me.Page, Mensagem)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoPxP" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnBuscaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnBuscaCliente.Click
        Try
            If ddlTabelaDePreco.SelectedValue.Length > 0 Then
                Dim strJScript As String = ""
                HttpContext.Current.Session("ssCampo") = "Livre"
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value)
            Else
                MsgBox(Me.Page, "Tabela de Preço não foi selecionada.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            Dim tabelaDePreco As String = String.Empty
            Dim cliente As String = String.Empty
            Dim endCliente As String = String.Empty
            Dim produto As String = String.Empty

            If ddlTabelaDePreco.SelectedValue.Length > 0 Then tabelaDePreco = ddlTabelaDePreco.SelectedValue.ToString()

            If txtCodigoCliente.Value.Length > 0 Then
                cliente = txtCodigoCliente.Value.Split("-")(0)
                endCliente = txtCodigoCliente.Value.Split("-")(1)
            End If

            If ddlProduto.SelectedValue.Length > 0 Then produto = ddlProduto.SelectedValue.ToString()

            CarregarListaProdutos(tabelaDePreco, cliente, endCliente, produto, "", "")

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            IniciarPxP("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            IniciarPxP("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            IniciarPxP("D")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutosXPrecos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjudaTabela_Click(sender As Object, e As EventArgs) Handles lnkAjudaTabela.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutosXPrecos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class