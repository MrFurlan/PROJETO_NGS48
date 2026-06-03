Imports Microsoft.VisualBasic
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class MemorandoDeExportacao
    Inherits BasePage

#Region "Variáveis locais / Auxiliares"
    Private objMemorando As [Lib].Negocio.MemorandoDeExportacao
#End Region

#Region "Load"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("MemorandoDeExportacao", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If
            Limpar()
            txtDataInicial.Text = ("01/01/" & Now.Year)
            txtDataFinal.Text = ("31/12/" & Now.Year)
            ddl.Carregar(ddlPais, CarregarDDL.Tabela.Pais, "", True)
            ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
            ddl.Carregar(ddlSafraSelecaoNotas, CarregarDDL.Tabela.Safra, "")
            ddlTipoConhec.SelectedValue = 10
            Session("objMemorando") = New [Lib].Negocio.MemorandoDeExportacao
            BuscarMoedas()
            BuscarIndexadores()
            carregarEstado()
        End If

        If CarregarEmitente() Then Exit Sub
        If CarregarComprovando() Then Exit Sub
        If CarregarExpEquiparado() Then Exit Sub
        If CarregarRepresentante() Then Exit Sub
        If CarregarProduto() Then Exit Sub
        If CarregarProdutoConsulta() Then Exit Sub
        If CarregarMemoEquipExport() Then Exit Sub
    End Sub

    Private Function CarregarMemoEquipExport() As Boolean
        If Not Session("objMemorandoSelecionado" & HID.Value) Is Nothing Then
            SessaoRecuperaMemorando()
            Dim objConsultaMemoEquipExport As New [Lib].Negocio.MemorandoDeExportacao(CType(Session("objMemorandoSelecionado" & HID.Value), [Lib].Negocio.MemorandoDeExportacao).CodigoEmpresaMemorando, CType(Session("objMemorandoSelecionado" & HID.Value), [Lib].Negocio.MemorandoDeExportacao).EnderecoEmpresaMemorando, CType(Session("objMemorandoSelecionado" & HID.Value), [Lib].Negocio.MemorandoDeExportacao).CodigoMemorando)
            objMemorando.NumeroNota = objConsultaMemoEquipExport.NumeroNota
            objMemorando.Serie = objConsultaMemoEquipExport.Serie
            objMemorando.EntradaSaida = objConsultaMemoEquipExport.EntradaSaida
            objMemorando.CodigoMemorandoEquiparado = objConsultaMemoEquipExport.CodigoMemorando
            txtNota.Text = objMemorando.NumeroNota
            txtSerie.Text = objMemorando.Serie
            txtMemorandoEquiparadoExp.Visible = True
            txtMemorandoEquiparadoExp.Text = objConsultaMemoEquipExport.CodigoMemorando
            Session.Remove("objMemorandoSelecionado" & HID.Value)
            SessaoSalvaMemorando()
            Return True
        End If
        Return False
    End Function

    Private Function CarregarEmitente() As Boolean
        If Not Session("objClienteEmitente" & HID.Value) Is Nothing Then
            SessaoRecuperaMemorando()
            objMemorando.CodigoEmpresaMemorando = CType(Session("objClienteEmitente" & HID.Value), [Lib].Negocio.Cliente).Codigo
            objMemorando.EnderecoEmpresaMemorando = CType(Session("objClienteEmitente" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco

            With objMemorando.EmpresaMemorando
                txtEmitente.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
            End With

            objMemorando.CodigoEmpresa = ""
            objMemorando.EnderecoEmpresa = 0
            objMemorando.CodigoCliente = ""
            objMemorando.EnderecoCliente = 0
            objMemorando.NumeroNota = 0
            objMemorando.Serie = ""
            objMemorando.EntradaSaida = ""
            txtNota.Text = ""
            txtSerie.Text = ""

            Session.Remove("objClienteEmitente" & HID.Value)
            SessaoSalvaMemorando()
            NossaEmissao()
            Return True
        End If
        Return False
    End Function

    Private Function CarregarExpEquiparado() As Boolean
        If Not Session("objClienteExpEquiparado" & HID.Value) Is Nothing Then
            SessaoRecuperaMemorando()
            objMemorando.CodExportadorEquiparado = CType(Session("objClienteExpEquiparado" & HID.Value), [Lib].Negocio.Cliente).Codigo
            objMemorando.EnderecoExportadorEquiparado = CType(Session("objClienteExpEquiparado" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco

            With objMemorando.ExportadorEquiparadoMemorando
                TxtExpEquiparado.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                If TxtExpEquiparado.Text <> "" Then
                    Dim strJScript As String = ""
                    strJScript += "var x = (screen.height / 2) - 250; "
                    strJScript += "var y = (screen.width / 2) - 400; "
                    strJScript += "window.open(""ConsultaMemorandoEquiparado.aspx?ExpEqui=" & objMemorando.CodExportadorEquiparado & "&EndExpEqui=" & objMemorando.EnderecoExportadorEquiparado & "&tipo=v"", """", ""resizable=no, menubar=no, scrollbars=yes, width=800, height=500, top="" + x + "", left="" + y + """");"
                    ScriptManager.RegisterClientScriptBlock(Me, BtnCadExpEquiparado.GetType(), "ConsMem", strJScript, True)
                End If
            End With

            Session.Remove("objClienteExpEquiparado" & HID.Value)
            SessaoSalvaMemorando()
            Return True
        End If
        Return False
    End Function

    Private Function CarregarComprovando() As Boolean
        If Not Session("objClienteComprovando" & HID.Value) Is Nothing Then
            SessaoRecuperaMemorando()
            objMemorando.CodigoClienteMemorando = CType(Session("objClienteComprovando" & HID.Value), [Lib].Negocio.Cliente).Codigo
            objMemorando.EnderecoClienteMemorando = CType(Session("objClienteComprovando" & HID.Value), [Lib].Negocio.Cliente).CodigoEndereco
            With objMemorando.ClienteMemorando
                txtComprovando.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
            End With
            Session.Remove("objClienteComprovando" & HID.Value)
            SessaoSalvaMemorando()
            Return True
        End If
        Return False
    End Function

    Private Function CarregarRepresentante() As Boolean
        If Not Session("objClienteRepresentante" & HID.Value) Is Nothing Then
            SessaoRecuperaMemorando()
            txtRepresentante.Text = CType(Session("objClienteRepresentante" & HID.Value), [Lib].Negocio.Cliente).Nome
            objMemorando.Representante = txtRepresentante.Text
            Dim Representante() As String = HRepresentante.Value.Split("-")
            objMemorando.CodigoRepresentante = Representante(0)
            Session.Remove("objClienteRepresentante" & HID.Value)
            SessaoSalvaMemorando()
            Return True
        End If
        Return False
    End Function

    Private Function CarregarComprovandoConsulta() As Boolean
        If Not Session("objClienteComprovCons" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteComprovCons" & HID.Value), [Lib].Negocio.Cliente))
            txtComprovandoConsulta.Text = itemCliente.Text
            HComprovandoConsulta.Value = itemCliente.Value
            Session.Remove("objClienteComprovCons" & HID.Value)
            Return True
        End If
        Return False
    End Function

    Public Function CarregarProduto() As Boolean

    End Function

    Public Function CarregarProdutoConsulta() As Boolean

        Return False
    End Function

    Public Sub NossaEmissao()
        SessaoRecuperaMemorando()
        If objMemorando.EmpresaMemorando Is Nothing OrElse objMemorando.EmpresaMemorando.Empresa.Empresa_id Is Nothing Then
            objMemorando.NossaEmissao = False
            BtnNotaDeComprovacao.Visible = False
            txtQuantidadeSaldoNota.Visible = False
            txtNota.Enabled = True
            txtSerie.Enabled = True
        Else
            objMemorando.NossaEmissao = True
            BtnNotaDeComprovacao.Visible = True
            txtQuantidadeSaldoNota.Visible = True
            txtNota.Enabled = False
            txtSerie.Enabled = False
        End If
    End Sub

    Private Sub BuscarMoedas()
        Dim objMoedas As New [Lib].Negocio.Moedas()

        ddlMoeda.Items.Clear()

        ddlMoeda.Items.Add(New ListItem("", 0))

        For Each objMoeda As [Lib].Negocio.Moeda In objMoedas
            ddlMoeda.Items.Add(New ListItem(objMoeda.Codigo.ToString() & "-" & objMoeda.Descricao, objMoeda.Codigo.ToString()))
        Next
    End Sub

    Private Sub BuscarIndexadores()
        Dim objIndexadores As New [Lib].Negocio.Indexadores()
        If objIndexadores.Selecionar() Then
            ddlIndexador.Items.Add(New ListItem("", 0))
            For Each objIndexador As [Lib].Negocio.Indexador In objIndexadores

                ddlIndexador.Items.Add(New ListItem(objIndexador.Codigo.ToString() & "-" & objIndexador.Descricao, _
                                                    objIndexador.Codigo.ToString()))
            Next
        End If
    End Sub

    Private Sub carregarEstado()
        ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)
    End Sub

#End Region

#Region "Sessão"

    Private Sub SessaoSalvaMemorando()
        Session("objMemorando") = objMemorando
    End Sub

    Private Sub SessaoRecuperaMemorando()
        objMemorando = CType(Session("objMemorando"), [Lib].Negocio.MemorandoDeExportacao)
    End Sub

    Private Function indiceInicial() As Integer
        Return gridNotas.PageSize * gridNotas.PageIndex
    End Function

#End Region

#Region "Botões"

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objClienteMeM" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteMeM" & HID.Value), [Lib].Negocio.Cliente))
            txtEmitenteConsulta.Text = itemCliente.Text
            HEmitenteConsulta.Value = itemCliente.Value
            Session.Remove("objClienteMeM" & HID.Value)
        ElseIf Session("objEmitente" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objEmitente" & HID.Value), [Lib].Negocio.Cliente))
            txtEmitente.Text = itemCliente.Text
            HEmitente.Value = itemCliente.Value
            Session.Remove("objEmitente" & HID.Value)
        ElseIf Session("objExpEquiparado" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objExpEquiparado" & HID.Value), [Lib].Negocio.Cliente))
            TxtExpEquiparado.Text = itemCliente.Text
            HExportadorEquiparado.Value = itemCliente.Value
            Session.Remove("objExpEquiparado" & HID.Value)
        ElseIf Session("objComprovando" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objComprovando" & HID.Value), [Lib].Negocio.Cliente))
            txtComprovando.Text = itemCliente.Text
            HComprovando.Value = itemCliente.Value
            Session.Remove("objComprovando" & HID.Value)
        ElseIf Session("objRepresentante" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objRepresentante" & HID.Value), [Lib].Negocio.Cliente))
            txtRepresentante.Text = itemCliente.Text
            HRepresentante.Value = itemCliente.Value
            Session.Remove("objRepresentante" & HID.Value)
        ElseIf Session("objClienteComprovCons" & HID.Value) IsNot Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteComprovCons" & HID.Value), [Lib].Negocio.Cliente))
            txtComprovandoConsulta.Text = itemCliente.Text
            HComprovandoConsulta.Value = itemCliente.Value
            Session.Remove("objClienteComprovCons" & HID.Value)
        ElseIf Session("objProdutoMeM" & HID.Value) IsNot Nothing Then
            SessaoRecuperaMemorando()
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoMeM" & HID.Value)
            objMemorando.CodigoProduto = objProduto.Codigo
            txtProduto.Text = objProduto.Codigo & " - " & objProduto.Nome
            Session.Remove("objProdutoMeM" & HID.Value)
        ElseIf Session("objProdutoMemCons" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = Session("objProdutoMemCons" & HID.Value)
            txtProdutoConsulta.Text = objProduto.Codigo & " - " & objProduto.Nome
            Session.Remove("objProdutoMemCons" & HID.Value)
        End If
    End Sub

    Protected Sub btnEmitente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteMeM" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnCadEmitente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objEmitente" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnExpEquiparado_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objExpEquiparado" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnComprovando_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objComprovando" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Popup.ConsultarClientes(Me, "Representante", txtRepresentante, HRepresentante, "", True)
        ucConsultaClientes.SetarTipoCliente(eTipoCliente.Representantes)
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objRepresentante" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaProduto.Limpar()
        Session("Where" & HID.Value) = "Situacao = 1"
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoMeM" & HID.Value, txtNome.ClientID, True)
    End Sub

#End Region

#Region "GridNota"

    Protected Sub btnQtde_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnQtde As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btnQtde.NamingContainer, GridViewRow)

        SessaoRecuperaMemorando()
        objMemorando.NotasComprovadas(indiceInicial() + row.RowIndex).QuantidadeMemorando = objMemorando.NotasComprovadas(indiceInicial() + row.RowIndex).Saldo
        SessaoSalvaMemorando()

        Dim QtdeMem As Decimal = objMemorando.QuantidadeMemorando
        txtQuantidadeMemorando.Text = QtdeMem.ToString("N0")
        If objMemorando.NossaEmissao Then
            txtQuantidadeSaldoNota.Text = (objMemorando.NotaDeComprovacao.QtdeNota - objMemorando.NotaDeComprovacao.QtdeJaComprovada + objMemorando.NotaDeComprovacao.QtdeComprovadaNesteMemorando - QtdeMem).ToString("N0")
        End If

        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub btnZerar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim btnZerar As Button = CType(sender, Button)
        Dim row As GridViewRow = CType(btnZerar.NamingContainer, GridViewRow)

        SessaoRecuperaMemorando()
        objMemorando.NotasComprovadas(indiceInicial() + row.RowIndex).QuantidadeMemorando = 0
        SessaoSalvaMemorando()

        Dim QtdeMem As Decimal = objMemorando.QuantidadeMemorando
        txtQuantidadeMemorando.Text = QtdeMem.ToString("N0")
        If objMemorando.NossaEmissao Then
            txtQuantidadeSaldoNota.Text = (objMemorando.NotaDeComprovacao.QtdeNota - objMemorando.NotaDeComprovacao.QtdeJaComprovada + objMemorando.NotaDeComprovacao.QtdeComprovadaNesteMemorando - QtdeMem).ToString("N0")
        End If

        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub txtQtde_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim txtQtde As TextBox = CType(sender, TextBox)
        Dim row As GridViewRow = CType(txtQtde.NamingContainer, GridViewRow)
        SessaoRecuperaMemorando()
        If objMemorando.NotasComprovadas(indiceInicial() + row.RowIndex).Saldo >= CDec(txtQtde.Text) Then
            objMemorando.NotasComprovadas(indiceInicial() + row.RowIndex).QuantidadeMemorando = CDec(txtQtde.Text)
            txtQuantidadeMemorando.Text = objMemorando.QuantidadeMemorando
        End If
        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub gridNotas_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        SessaoRecuperaMemorando()
        gridNotas.PageIndex = e.NewPageIndex
        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

#End Region

    Protected Sub txtMemorando_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        SessaoRecuperaMemorando()

        If objMemorando.EmpresaMemorando Is Nothing Then
            txtMemorando.Text = ""
            MsgBox(Me.Page, "Informe a Empresa do Memorando")
            Exit Sub
        End If
        objMemorando.CodigoMemorando = txtMemorando.Text

        Dim mem As New [Lib].Negocio.MemorandoDeExportacao(objMemorando.CodigoEmpresaMemorando, objMemorando.EnderecoEmpresaMemorando, objMemorando.CodigoMemorando)
        If mem.NotasComprovadas.Count > 0 Then
            objMemorando = mem
            objMemorando.IUD = "U"
            lnkNovo.Parent.Visible = False
            lnkAlterar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            SessaoSalvaMemorando()
            CarregarFormularioComAClasse()
            MsgBox(Me.Page, "Este Memorando já Existe")
            Exit Sub
        End If

        'objMemorando.CodigoMemorando = txtMemorando.Text
        SessaoSalvaMemorando()
    End Sub

    Public Sub CarregarFormularioComAClasse()
        SessaoRecuperaMemorando()


        If objMemorando.CodigoEmpresaMemorando.Length > 0 Then
            With objMemorando.EmpresaMemorando
                txtEmitente.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                HEmitente.Value = .Codigo & "-" & .CodigoEndereco
            End With
        Else
            txtEmitente.Text = ""
            HEmitente.Value = ""
        End If

        If objMemorando.CodExportadorEquiparado.Length > 0 Then
            With objMemorando.ExportadorEquiparadoMemorando
                TxtExpEquiparado.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                HExportadorEquiparado.Value = .Codigo & "-" & .CodigoEndereco
            End With
        Else
            TxtExpEquiparado.Text = ""
            HExportadorEquiparado.Value = ""
        End If

        NossaEmissao()

        If objMemorando.CodigoClienteMemorando.Length > 0 Then
            With objMemorando.ClienteMemorando
                txtComprovando.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                HComprovando.Value = .Codigo & "-" & .CodigoEndereco
            End With
        Else
            txtComprovando.Text = ""
            HComprovando.Value = ""
        End If

        'If objMemorando.Representante.Length > 0 Then
        '    txtRepresentante.Text = objMemorando.Representante
        'Else
        '    txtRepresentante.Text = ""
        '    HRepresentante.Value = ""
        'End If
        'Representante
        If Not String.IsNullOrWhiteSpace(objMemorando.CodigoRepresentante) Then
            With objMemorando.RepresentanteMemorando
                txtRepresentante.Text = .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado & " - " & .CodigoFormatado
                HRepresentante.Value = .Codigo & "-" & .CodigoEndereco
            End With
        Else
            txtRepresentante.Text = String.Empty
            HRepresentante.Value = String.Empty
        End If

        If objMemorando.CodigoProduto.Length > 0 Then
            txtProduto.Text = objMemorando.Produto.Codigo & " - " & objMemorando.Produto.Nome
        Else
            objMemorando.CodigoProduto = 0
            txtProduto.Text = ""
        End If

        If objMemorando.ConhecimentosDeEmbarque Is Nothing OrElse objMemorando.ConhecimentosDeEmbarque.Count = 0 Then
            txtNumeroConhecimento.Text = ""
            txtDataConhecimento.Text = ""
            ddlTipoConhec.SelectedValue = 10
        End If

        If objMemorando.RegistrosDeExportacao Is Nothing OrElse objMemorando.RegistrosDeExportacao.Count = 0 Then
            TxtRegExportacao.Text = ""
            TxtDataRegExp.Text = ""
        End If

        If objMemorando.TipoDocumento.Trim <> "" Then
            ddlTipoDocumento.SelectedValue = objMemorando.TipoDocumento
        Else
            ddlTipoDocumento.SelectedValue = 0
        End If

        With objMemorando
            txtMemorando.Text = .CodigoMemorando
            txtDataMemorando.Text = .DataMemorando.ToString("dd/MM/yyyy")
            txtNumeroDespacho.Text = .NrDespachoExp
            txtDataDespacho.Text = .DataDespachoExp.ToString("dd/MM/yyyy")
            txtDataAverba.Text = .DataAverbacao.ToString("dd/MM/yyyy")
            ddlTipoDocumento.Text = .TipoDocumento
            ddlPais.SelectedValue = IIf(.CodigoPaisDestino = 0, "", .CodigoPaisDestino)
            txtNavio.Text = .Navio
            txtNota.Text = .NumeroNota
            txtSerie.Text = .Serie
            txtQuantidadeMemorando.Text = .QuantidadeMemorando
            txtValorNota.Text = .ValorNota
            TxtDataDaNota.Text = .DataNota.ToString("dd/MM/yyyy")
            ddlMoeda.SelectedValue = .Moeda
            ddlIndexador.SelectedValue = .Indexador
            TxtNumAtoConcessorio.Text = .NumAtoConcessorio
            TxtDtaRegAtoConcessorio.Text = .DtaRegAtoConcessorio.ToString("dd/MM/yyyy")
            TxtDtaValidAtoConcessorio.Text = .DtaValidAtoConcessorio.ToString("dd/MM/yyyy")
            If .CodigoMemorandoEquiparado <> 0 Then
                txtMemorandoEquiparadoExp.Text = .CodigoMemorandoEquiparado
                txtMemorandoEquiparadoExp.Visible = True
            Else
                txtMemorandoEquiparadoExp.Visible = False
            End If

            gridNotas.DataSource = .NotasComprovadas
            gridNotas.DataBind()
            GridConhecimento.DataSource = .ConhecimentosDeEmbarque
            GridConhecimento.DataBind()
            GridRegistroExp.DataSource = .RegistrosDeExportacao
            GridRegistroExp.DataBind()
        End With
    End Sub

    Public Function ValidarCarregarMemorando() As Boolean
        SessaoRecuperaMemorando()

        If String.IsNullOrWhiteSpace(objMemorando.CodigoEmpresaMemorando) Then
            MsgBox(Me.Page, "Informe a empresa do Memorando")
            Return False
        End If

        If String.IsNullOrWhiteSpace(objMemorando.CodigoClienteMemorando) Then
            MsgBox(Me.Page, "Informe qual empresa esta sendo comprovada pelo memorando")
            Return False
        End If

        If String.IsNullOrWhiteSpace(objMemorando.CodigoMemorando) Then
            MsgBox(Me.Page, "Informe o Numero do memorando")
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtDataMemorando.Text) Then
            MsgBox(Me.Page, "Informe a data do memorando")
            Return False
        Else
            objMemorando.DataMemorando = CDate(txtDataMemorando.Text)
        End If
        'Representante
        If (Not String.IsNullOrWhiteSpace(HRepresentante.Value)) Then
            objMemorando.CodigoRepresentante = HRepresentante.Value.Split("-").ToArray(0)
            objMemorando.EnderecoRepresentante = HRepresentante.Value.Split("-").ToArray(1)
        Else
            objMemorando.CodigoRepresentante = "NULL"
            objMemorando.EnderecoRepresentante = "NULL"
        End If
        'objMemorando.CodigoRepresentante = IIf(String.IsNullOrWhiteSpace(txtRepresentante.Text), "NULL", HRepresentante.Value.Split("-").ToArray(0))
        'objMemorando.EnderecoRepresentante = IIf(String.IsNullOrWhiteSpace(txtRepresentante.Text), "NULL", HRepresentante.Value.Split("-").ToArray(1))

        objMemorando.NrDespachoExp = txtNumeroDespacho.Text
        If objMemorando.NrDespachoExp.Length = 0 And objMemorando.NossaEmissao Then
            MsgBox(Me.Page, "Informe o Numero do Despacho Exportacao")
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtDataDespacho.Text) Then
            If objMemorando.NossaEmissao Then
                MsgBox(Me.Page, "Informe a Data do Despacho Exportacao")
                Return False
            End If
        Else
            objMemorando.DataDespachoExp = CDate(txtDataDespacho.Text)
        End If

        If String.IsNullOrWhiteSpace(txtDataAverba.Text) Then
            MsgBox(Me.Page, "Informe a data de Averbação")
            Return False
        Else
            objMemorando.DataAverbacao = CDate(txtDataAverba.Text)
        End If

        If objMemorando.ConhecimentosDeEmbarque Is Nothing OrElse objMemorando.ConhecimentosDeEmbarque.Count = 0 Then
            MsgBox(Me.Page, "Informe o Numero do Conhecimento De Embarque e Clique Na Figura (mais) Para Adicionar")
            Return False
        End If

        If objMemorando.RegistrosDeExportacao Is Nothing OrElse objMemorando.RegistrosDeExportacao.Count = 0 Then
            MsgBox(Me.Page, "Informe o Numero do Registro de Exportação e Clique Na Figura (mais) Para Adicionar")
            Return False
        End If

        If String.IsNullOrWhiteSpace(ddlTipoDocumento.SelectedValue) Then
            objMemorando.TipoDocumento = 0
        Else
            objMemorando.TipoDocumento = ddlTipoDocumento.SelectedValue
        End If

        If String.IsNullOrWhiteSpace(ddlPais.SelectedValue) Then
            objMemorando.CodigoPaisDestino = 0
        Else
            objMemorando.CodigoPaisDestino = ddlPais.SelectedValue
        End If

        If objMemorando.CodigoPaisDestino = 0 And objMemorando.NossaEmissao Then
            MsgBox(Me.Page, "Informe o Pais")
            Return False
        End If

        objMemorando.Navio = txtNavio.Text
        If objMemorando.Navio.Length = 0 And objMemorando.NossaEmissao Then
            MsgBox(Me.Page, "Informe o Navio")
            Return False
        End If

        If Not objMemorando.NossaEmissao Then
            objMemorando.NumeroNota = IIf(txtNota.Text.Length = 0, 0, txtNota.Text)
            objMemorando.Serie = IIf(txtSerie.Text.Length = 0, 0, txtSerie.Text)
            objMemorando.DataNota = IIf(TxtDataDaNota.Text.Length = 0, "", CDate(TxtDataDaNota.Text))
            objMemorando.Moeda = IIf(ddlMoeda.Text.Length = 0, 0, ddlMoeda.SelectedValue)
            objMemorando.Indexador = IIf(ddlIndexador.Text.Length = 0, 0, ddlIndexador.SelectedValue)
            objMemorando.ValorNota = IIf(txtValorNota.Text.Length = 0, 0, txtValorNota.Text)

        End If

        If objMemorando.NumeroNota = 0 Or objMemorando.Serie = "" Then
            MsgBox(Me.Page, "Informe o Numero e a serie da Nota")
            Return False
        End If

        If Not String.IsNullOrWhiteSpace(TxtNumAtoConcessorio.Text) AndAlso CDec(TxtNumAtoConcessorio.Text) > 0 Then
            objMemorando.NumAtoConcessorio = TxtNumAtoConcessorio.Text
            If TxtDtaRegAtoConcessorio.Text.Length = 0 Then
                MsgBox(Me.Page, "Informe a Data de registro Do Ato Concessório.")
                Return False
            Else
                objMemorando.DtaRegAtoConcessorio = CDate(TxtDtaRegAtoConcessorio.Text)
            End If
            If TxtDtaValidAtoConcessorio.Text.Length = 0 Then
                MsgBox(Me.Page, "Informe a Data de Validade Do Ato Concessório.")
                Return False
            Else
                objMemorando.DtaValidAtoConcessorio = CDate(TxtDtaValidAtoConcessorio.Text)
            End If
        End If
        Return True
    End Function

    Public Sub Limpar()
        objMemorando = New [Lib].Negocio.MemorandoDeExportacao
        objMemorando.IUD = "I"
        lnkNovo.Enabled = True
        lnkAlterar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkNovo.Parent.Visible = False
        txtMemorandoEquiparadoExp.Visible = False
        SessaoSalvaMemorando()
        CarregarFormularioComAClasse()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaProduto.SetarHID(HID.Value)
    End Sub

    Protected Sub gridConsulta_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridNotasComprovacaoPendentes.SelectedIndexChanged, gridConsulta.SelectedIndexChanged
        Limpar()
        Dim Emp As String = gridConsulta.SelectedRow.Cells(2).Text
        Dim EndEmp As Integer = gridConsulta.SelectedRow.Cells(3).Text
        Dim NumMem As String = gridConsulta.SelectedRow.Cells(6).Text
        Dim Memorando As New [Lib].Negocio.MemorandoDeExportacao(Emp, EndEmp, NumMem)
        Memorando.IUD = "U"
        lnkNovo.Parent.Visible = False
        lnkAlterar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True
        objMemorando = Memorando
        SessaoSalvaMemorando()
        'Session("MemorandoOriginal") = New ClonarObjeto(Of MemorandoDeExportacao)().Clonar(objMemorando)
        Session("MemorandoOriginal") = New [Lib].Negocio.MemorandoDeExportacao(Emp, EndEmp, NumMem)
        CarregarFormularioComAClasse()
        ddlSafraSelecaoNotas.SelectedValue = ""
        TBMemorando.ActiveTabIndex = 1
        lnkNovo.Parent.Visible = False
        lnkAlterar.Parent.Visible = True
    End Sub

    Private Function ValidaConhecimento() As Boolean
        If txtNumeroConhecimento.Text.Trim = "" Then
            Return False
        ElseIf txtDataConhecimento.Text.Trim = "" Then
            Return False
        ElseIf ddlTipoConhec.SelectedValue.Trim = "" Then
            Return False
        End If
        Return True

    End Function

    Private Function ValidaRegExp() As Boolean
        If TxtRegExportacao.Text.Trim = "" Then
            Return False
        ElseIf TxtDataRegExp.Text.Trim = "" Then
            Return False
        ElseIf ddlEstado.SelectedValue.Trim = "" Then
            Return False
        End If
        Return True

    End Function

    Protected Sub BtnNotaDeComprovacao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gridNotas.Visible = False
        gridNotaSaida.Visible = True
        SessaoRecuperaMemorando()
        Dim notasdeSaida As New [Lib].Negocio.ListNotaDeExportacao(objMemorando)
        Session("NotasDeSaida") = notasdeSaida
        gridNotaSaida.DataSource = notasdeSaida
        gridNotaSaida.DataBind()
    End Sub

    Protected Sub gridNotaSaida_PageIndexChanging(ByVal sender As Object, ByVal e As System.Web.UI.WebControls.GridViewPageEventArgs)
        gridNotaSaida.PageIndex = e.NewPageIndex
        gridNotaSaida.DataSource = CType(Session("NotasDeSaida"), [Lib].Negocio.ListNotaDeExportacao)
        gridNotaSaida.DataBind()
    End Sub

    Protected Sub gridNotaSaida_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridNotaSaida.SelectedIndexChanged
        Dim notasdeSaida As [Lib].Negocio.ListNotaDeExportacao = Session("NotasDeSaida")
        SessaoRecuperaMemorando()
        'Alimenta OBJMemorando (classe)
        Dim i As Integer = gridNotaSaida.PageSize * gridNotaSaida.PageIndex + gridNotaSaida.SelectedIndex
        objMemorando.NotaDeComprovacao = notasdeSaida(i)
        objMemorando.CodigoEmpresa = notasdeSaida(i).CodigoEmpresa
        objMemorando.EnderecoEmpresa = notasdeSaida(i).EnderecoEmpresa
        objMemorando.CodigoCliente = notasdeSaida(i).CodigoCliente
        objMemorando.EnderecoCliente = notasdeSaida(i).EnderecoCliente
        objMemorando.NumeroNota = notasdeSaida(i).NumeroNota
        objMemorando.Serie = notasdeSaida(i).Serie
        objMemorando.EntradaSaida = notasdeSaida(i).EntradaSaida
        objMemorando.DataNota = notasdeSaida(i).NotaFiscal.DataNota
        objMemorando.ValorNota = notasdeSaida(i).NotaFiscal.TotalNota
        objMemorando.Moeda = notasdeSaida(i).NotaFiscal.Pedido.CodigoMoeda
        objMemorando.Indexador = notasdeSaida(i).NotaFiscal.Pedido.CodigoIndexador
        If Not notasdeSaida(i).NotaFiscal.DadosDaExportacao Is Nothing Then
            objMemorando.Navio = notasdeSaida(i).NotaFiscal.DadosDaExportacao.Navio
            objMemorando.NrDespachoExp = notasdeSaida(i).NotaFiscal.DadosDaExportacao.NrDespachoExp
            objMemorando.DataDespachoExp = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataDespachoExp
            If notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataAverbacao IsNot Nothing Then
                objMemorando.DataAverbacao = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataAverbacao
            End If
            objMemorando.CodigoPaisDestino = notasdeSaida(i).NotaFiscal.DadosDaExportacao.CodigoPaisDestino

            objMemorando.NumAtoConcessorio = notasdeSaida(i).NotaFiscal.DadosDaExportacao.NumAtoConcessorio
            objMemorando.DtaRegAtoConcessorio = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DtaRegAtoConcessorio
            objMemorando.DtaValidAtoConcessorio = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DtaValidAtoConcessorio

            For Each re As NotaFiscalXRE In notasdeSaida(i).NotaFiscal.DadosDaExportacaoRE
                Dim objMemoExpXRegExp As New [Lib].Negocio.MemorandoDeExportacaoXRegistroDeExportacao(objMemorando)
                objMemoExpXRegExp.CodRegistroDeExportacao = re.RegistroDeExportacao
                objMemoExpXRegExp.DataRegExportacao = re.DataRegistroDeExportacao
                objMemoExpXRegExp.UfProdutor = re.UfProdutor
                objMemorando.RegistrosDeExportacao.Add(objMemoExpXRegExp)
            Next

        Else
            objMemorando.Navio = ""
            objMemorando.NrDespachoExp = ""
            objMemorando.DataDespachoExp = Now
            objMemorando.DataAverbacao = Now
            objMemorando.CodigoPaisDestino = 0

            objMemorando.RegistrosDeExportacao.Clear()

        End If
        SessaoSalvaMemorando()

        'Alimenta TXTs(campos)
        txtQuantidadeSaldoNota.Text = notasdeSaida(i).QtdeNota - notasdeSaida(i).QtdeJaComprovada + notasdeSaida(i).QtdeComprovadaNesteMemorando - objMemorando.QuantidadeMemorando
        txtNota.Text = notasdeSaida(i).NumeroNota
        txtSerie.Text = notasdeSaida(i).Serie
        txtValorNota.Text = notasdeSaida(i).ValorNota
        ddlMoeda.SelectedValue = notasdeSaida(i).NotaFiscal.Pedido.CodigoMoeda
        ddlIndexador.SelectedValue = notasdeSaida(i).NotaFiscal.Pedido.CodigoIndexador
        If Not notasdeSaida(i).NotaFiscal.DadosDaExportacao Is Nothing Then
            txtNavio.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.Navio
            txtNumeroDespacho.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.NrDespachoExp
            txtDataDespacho.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataDespachoExp

            If notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataAverbacao IsNot Nothing Then
                txtDataAverba.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DataAverbacao
            End If

            ddlPais.SelectedValue = notasdeSaida(i).NotaFiscal.DadosDaExportacao.CodigoPaisDestino

            TxtNumAtoConcessorio.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.NumAtoConcessorio
            TxtDtaRegAtoConcessorio.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DtaRegAtoConcessorio
            TxtDtaValidAtoConcessorio.Text = notasdeSaida(i).NotaFiscal.DadosDaExportacao.DtaValidAtoConcessorio

            GridRegistroExp.DataSource = objMemorando.RegistrosDeExportacao.ToArray
            GridRegistroExp.DataBind()
        Else
            txtNavio.Text = ""
            txtNumeroDespacho.Text = ""
            txtDataDespacho.Text = Now
            txtNumeroConhecimento.Text = ""
            txtDataConhecimento.Text = Now
            txtDataAverba.Text = ""
            ddlPais.SelectedValue = ""

            TxtNumAtoConcessorio.Text = ""
            TxtDtaRegAtoConcessorio.Text = ""
            TxtDtaValidAtoConcessorio.Text = ""

            GridRegistroExp.DataSource = ""
            GridRegistroExp.DataBind()
        End If
        TxtDataDaNota.Text = notasdeSaida(i).NotaFiscal.DataNota
        gridNotas.Visible = True
        gridNotaSaida.Visible = False
        Session.Remove("NotasDeSaida")
    End Sub

    Protected Sub btnAtualizarConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        gridClientesPendentes.DataSource = Nothing
        gridClientesPendentes.DataBind()
        gridNotasPendentes.DataSource = Nothing
        gridNotasPendentes.DataBind()

        Dim notasdeSaidaPendentes As New [Lib].Negocio.ListNotaDeExportacao(Nothing, ddlSafra.SelectedValue)
        gridNotasComprovacaoPendentes.DataSource = notasdeSaidaPendentes
        gridNotasComprovacaoPendentes.DataBind()


        Dim sql As String
        sql = "SELECT NF.Empresa_id as Empresa," & vbCrLf & _
              "       NF.EndEmpresa_id as EndEmpresa," & vbCrLf & _
              "       NF.Cliente_Id as Cliente," & vbCrLf & _
              "       NF.EndCliente_Id as EndCliente," & vbCrLf & _
              "       C.Nome as NomeCliente," & vbCrLf & _
              "       C.Cidade + '-' + C.Estado as Cidade," & vbCrLf & _
              "       NF.EntradaSaida_Id as ES," & vbCrLf & _
              "       Case when Nf.EntradaSaida_Id = 'S' then 'COMPROVAÇÃO A RECEBER' Else 'COMPROVAÇÃO A EMITIR' END AS DescSituacao, " & vbCrLf & _
              "       NFxI.Produto_Id as Produto," & vbCrLf & _
              "       Prd.Nome as NomeProduto, " & vbCrLf & _
              "       sum(NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(sbMem.Quantidade,0)) as Saldo" & vbCrLf & _
              "  FROM NotasFiscais AS NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " INNER JOIN Pedidos P " & vbCrLf & _
              "    ON P.Empresa_id    = NF.Empresa_id" & vbCrLf & _
              "   AND P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf & _
              "   AND P.Pedido_id     = NF.Pedido" & vbCrLf & _
              " INNER JOIN Clientes C" & vbCrLf & _
              "    ON C.Cliente_id  = NF.Cliente_Id" & vbCrLf & _
              "   AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
              " INNER JOIN Produtos Prd" & vbCrLf & _
              "    on Prd.Produto_Id = NFxI.Produto_id" & vbCrLf & _
              " INNER JOIN Operacoes OP" & vbCrLf & _
              "    ON NFxI.Operacao    = OP.Operacao_Id " & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf & _
              "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "		 	 SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
              "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
              "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
              "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
              "                    nfd.EntradaSaida_Id," & vbCrLf & _
              "                    nfd.Serie_Id," & vbCrLf & _
              "                    nfd.Nota_Id," & vbCrLf & _
              "                    nfd.Produto_Id," & vbCrLf & _
              "                    nfd.Sequencia_Id," & vbCrLf & _
              "                    nfd.CFOP_Id," & vbCrLf & _
              "                    sum(nfd.Quantidade) as Quantidade" & vbCrLf & _
              "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
              "              Inner Join NotasFiscais nf" & vbCrLf & _
              "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
              "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
              "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
              "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
              "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf & _
              "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf & _
              "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf & _
              "              Where nf.situacao        in (1, 4)" & vbCrLf & _
              "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id," & vbCrLf & _
              "                       nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf & _
              "             ) SbDev" & vbCrLf & _
              "    On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
              "   and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
              "   and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
              "   and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
              "   and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
              "   and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
              "   and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
              "   and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
              "   and NFxI.Sequencia_id    = SbDev.Sequencia_Id" & vbCrLf & _
              "   and NFxI.CFOP_id         = SbDev.CFOP_Id" & vbCrLf & _
              "  LEFT JOIN (" & vbCrLf & _
              "             SELECT MExNF.Empresa_Id," & vbCrLf & _
              "                    MExNF.EndEmpresa_Id," & vbCrLf & _
              "                    MExNF.Cliente_Id," & vbCrLf & _
              "                    MExNF.EndCliente_Id," & vbCrLf & _
              "                    MExNF.EntradaSaida_Id," & vbCrLf & _
              "                    MExNF.Serie_Id," & vbCrLf & _
              "                    MExNF.Nota_Id," & vbCrLf & _
              "                    Mem.Produto as Produto_Id," & vbCrLf & _
              "                    sum(MExNF.Quantidade) as Quantidade" & vbCrLf & _
              " 		      FROM MemorandoDeExportacao Mem" & vbCrLf & _
              "              Inner Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
              "			 	    ON Mem.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
              "				   AND Mem.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
              "			  	   AND Mem.Memorando_Id           = MExNF.Memorando_Id" & vbCrLf & _
              "              Inner Join NotasFiscais NF" & vbCrLf & _
              "                 On NF.Empresa_id      = MExNF.Empresa_Id" & vbCrLf & _
              "                and NF.EndEmpresa_Id   = MExNF.EndEmpresa_Id" & vbCrLf & _
              "                and NF.Cliente_Id      = MExNF.Cliente_Id" & vbCrLf & _
              "                and NF.EndCliente_Id   = MExNF.EndCliente_Id" & vbCrLf & _
              "                and NF.EntradaSaida_Id = MExNF.EntradaSaida_Id" & vbCrLf & _
              "                and NF.Serie_Id        = MExNF.Serie_Id" & vbCrLf & _
              "                and NF.Nota_Id         = MExNF.Nota_Id" & vbCrLf & _
              "              Where NF.situacao        in (1,4)" & vbCrLf & _
              "             Group by MExNF.Empresa_Id, MExNF.EndEmpresa_Id, MExNF.Cliente_Id, MExNF.EndCliente_Id, MExNF.EntradaSaida_Id," & vbCrLf & _
              "                      MExNF.Serie_Id, MExNF.Nota_Id, Mem.Produto" & vbCrLf & _
              "          ) SbMem" & vbCrLf & _
              "    On NFxI.Empresa_id      = SbMem.Empresa_Id" & vbCrLf & _
              "   and NFxI.EndEmpresa_Id   = SbMem.EndEmpresa_Id" & vbCrLf & _
              "   and NFxI.Cliente_Id      = SbMem.Cliente_Id" & vbCrLf & _
              "   and NFxI.EndCliente_Id   = SbMem.EndCliente_Id" & vbCrLf & _
              "   and NFxI.EntradaSaida_Id = SbMem.EntradaSaida_Id" & vbCrLf & _
              "   and NFxI.Serie_Id        = SbMem.Serie_Id" & vbCrLf & _
              "   and NFxI.Nota_Id         = SbMem.Nota_Id" & vbCrLf & _
              "   --and NFxI.Produto_Id      = SbMem.Produto_Id" & vbCrLf & _
              " Where SO.Memorando    = 1" & vbCrLf & _
              "   and SO.Devolucao    = 'N'" & vbCrLf & _
              "   and not(OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and SO.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf & _
              "   and NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(SbMem.Quantidade,0) > 0" & vbCrLf

        If rbEmissaoMem.Checked Then
            sql &= "   and Nf.EntradaSaida_Id = 'E' "
        End If

        If rbComprovacaoMem.Checked Then
            sql &= "   and Nf.EntradaSaida_Id = 'S' "
        End If

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "   and P.Safra = '" & ddlSafra.SelectedValue & "'"
        End If

        '"  and nf.datadanota > getdate() - 180" & vbCrLf & _

        sql &= " Group by NF.Empresa_Id," & vbCrLf & _
               "          NF.EndEmpresa_Id," & vbCrLf & _
               "          NF.Cliente_Id, " & vbCrLf & _
               "          NF.EndCliente_Id," & vbCrLf & _
               "          C.Nome," & vbCrLf & _
               "          C.Cidade + '-' + C.Estado," & vbCrLf & _
               "          NF.EntradaSaida_Id," & vbCrLf & _
               "          Case when Nf.EntradaSaida_Id = 'S' then 'COMPROVAÇÃO A RECEBER' Else 'COMPROVAÇÃO A EMITIR' END, " & vbCrLf & _
               "          NFxI.Produto_Id," & vbCrLf & _
               "          Prd.Nome" & vbCrLf
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "ClientesPendentes")

        gridClientesPendentes.DataSource = ds.Tables("ClientesPendentes")
        gridClientesPendentes.DataBind()
        Session("DsMemorando") = ds
    End Sub

    Protected Sub gridClientesPendentes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dsMen As New DataSet
        dsMen = Session("DsMemorando")

        Dim row As DataRow = dsMen.Tables("ClientesPendentes").Rows(gridClientesPendentes.SelectedIndex)

        Dim sql As String
        sql = "SELECT NF.Empresa_id as Empresa," & vbCrLf & _
              "       NF.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
              "       NF.Cliente_Id as Cliente," & vbCrLf & _
              "       NF.EndCliente_Id as EndCliente," & vbCrLf & _
              "       NF.EntradaSaida_Id as ES," & vbCrLf & _
              "       NF.Serie_Id as Serie," & vbCrLf & _
              "       NF.Nota_Id as Nota," & vbCrLf & _
              "       NF.DataDaNota," & vbCrLf & _
              "       convert(int,getdate() - NF.DataDaNota) as DiasDecorridos," & vbCrLf & _
              "       NFxI.QuantidadeFiscal as QuantidadeNota," & vbCrLf & _
              "       isnull(sbDev.Quantidade,0) as QtdeDevolvida," & vbCrLf & _
              "       isnull(sbMem.Quantidade,0) as QtdeJaComprovada," & vbCrLf & _
              "       NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(sbMem.Quantidade,0) as Saldo" & vbCrLf & _
              "  FROM NotasFiscais AS NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " INNER JOIN Pedidos P " & vbCrLf & _
              "    ON P.Empresa_id    = NF.Empresa_id" & vbCrLf & _
              "   AND P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf & _
              "   AND P.Pedido_id     = NF.Pedido" & vbCrLf & _
              " INNER JOIN Operacoes OP" & vbCrLf & _
              "    ON NFxI.Operacao    = OP.Operacao_Id " & vbCrLf & _
              " INNER JOIN SubOperacoes SO" & vbCrLf & _
              "    ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf & _
              "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "		 	 SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
              "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
              "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
              "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
              "                    nfd.EntradaSaida_Id," & vbCrLf & _
              "                    nfd.Serie_Id," & vbCrLf & _
              "                    nfd.Nota_Id," & vbCrLf & _
              "                    nfd.Produto_Id," & vbCrLf & _
              "                    nfd.Sequencia_Id," & vbCrLf & _
              "                    nfd.CFOP_Id," & vbCrLf & _
              "                    sum(nfd.Quantidade) as Quantidade" & vbCrLf & _
              "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
              "              Inner Join NotasFiscais nf" & vbCrLf & _
              "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
              "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
              "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
              "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
              "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf & _
              "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf & _
              "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf & _
              "              Where nf.situacao        in (1, 4)" & vbCrLf & _
              "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id," & vbCrLf & _
              "                       nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf & _
              "             ) SbDev" & vbCrLf & _
              "    On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
              "   and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
              "   and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
              "   and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
              "   and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
              "   and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
              "   and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
              "   and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
              "   and NFxI.Sequencia_id    = SbDev.Sequencia_Id" & vbCrLf & _
              "   and NFxI.CFOP_id         = SbDev.CFOP_Id" & vbCrLf & _
              "  LEFT JOIN (" & vbCrLf & _
              "             SELECT MExNF.Empresa_Id," & vbCrLf & _
              "                    MExNF.EndEmpresa_Id," & vbCrLf & _
              "                    MExNF.Cliente_Id," & vbCrLf & _
              "                    MExNF.EndCliente_Id," & vbCrLf & _
              "                    MExNF.EntradaSaida_Id," & vbCrLf & _
              "                    MExNF.Serie_Id," & vbCrLf & _
              "                    MExNF.Nota_Id," & vbCrLf & _
              "                    Mem.Produto as Produto_Id," & vbCrLf & _
              "                    sum(MExNF.Quantidade) as Quantidade" & vbCrLf & _
              " 		         FROM MemorandoDeExportacao Mem" & vbCrLf & _
              "              Inner Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
              "			 	   ON Mem.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
              "				  AND Mem.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
              "			  	  AND Mem.Memorando_Id           = MExNF.Memorando_Id" & vbCrLf & _
              "              Inner Join NotasFiscais NF" & vbCrLf & _
              "                 On NF.Empresa_id      = MExNF.Empresa_Id" & vbCrLf & _
              "                and NF.EndEmpresa_Id   = MExNF.EndEmpresa_Id" & vbCrLf & _
              "                and NF.Cliente_Id      = MExNF.Cliente_Id" & vbCrLf & _
              "                and NF.EndCliente_Id   = MExNF.EndCliente_Id" & vbCrLf & _
              "                and NF.EntradaSaida_Id = MExNF.EntradaSaida_Id" & vbCrLf & _
              "                and NF.Serie_Id        = MExNF.Serie_Id" & vbCrLf & _
              "                and NF.Nota_Id         = MExNF.Nota_Id" & vbCrLf & _
              "              Where NF.situacao        in (1,4)" & vbCrLf & _
              "             Group by MExNF.Empresa_Id, MExNF.EndEmpresa_Id, MExNF.Cliente_Id, MExNF.EndCliente_Id, MExNF.EntradaSaida_Id," & vbCrLf & _
              "                      MExNF.Serie_Id, MExNF.Nota_Id, Mem.Produto" & vbCrLf & _
              "          ) SbMem" & vbCrLf & _
              "    On NFxI.Empresa_id      = SbMem.Empresa_Id" & vbCrLf & _
              "   and NFxI.EndEmpresa_Id   = SbMem.EndEmpresa_Id" & vbCrLf & _
              "   and NFxI.Cliente_Id      = SbMem.Cliente_Id" & vbCrLf & _
              "   and NFxI.EndCliente_Id   = SbMem.EndCliente_Id" & vbCrLf & _
              "   and NFxI.EntradaSaida_Id = SbMem.EntradaSaida_Id" & vbCrLf & _
              "   and NFxI.Serie_Id        = SbMem.Serie_Id" & vbCrLf & _
              "   and NFxI.Nota_Id         = SbMem.Nota_Id" & vbCrLf & _
              "   and NFxI.Produto_Id      = SbMem.Produto_Id" & vbCrLf & _
              " Where SO.Memorando    = 1" & vbCrLf & _
              "   and SO.Devolucao    = 'N'" & vbCrLf & _
              "   and not(OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and SO.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf & _
              "   and NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(SbMem.Quantidade,0) > 0" & vbCrLf

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "   and P.Safra = '" & ddlSafra.SelectedValue & "'"
        End If

        '      "   and nf.datadanota      > getdate() - 180 " & vbCrLf & _

        sql &= "   and nf.Empresa_id      ='" & row("Empresa") & "'" & vbCrLf & _
               "   and nf.EndEmpresa_Id   = " & row("EndEmpresa") & vbCrLf & _
               "   and nf.Cliente_id      ='" & row("Cliente") & "'" & vbCrLf & _
               "   and nf.EndCliente_Id   = " & row("EndCliente") & vbCrLf & _
               "   and nf.EntradaSaida_Id ='" & row("ES") & "'" & vbCrLf & _
               " order by NF.nota_id" & vbCrLf

        If dsMen.Tables.Contains("NotasPendentes") Then dsMen.Tables("NotasPendentes").Rows.Clear()
        dsMen.Merge(Banco.ConsultaDataSet(sql, "NotasPendentes"))

        gridNotasPendentes.DataSource = dsMen.Tables("NotasPendentes")
        gridNotasPendentes.DataBind()
        Session("DsMemorando") = dsMen
    End Sub

    Protected Sub btnIncluirDoGrid_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Limpar()
        Dim btnIncluir As Button = CType(sender, Button)
        Dim Gridrow As GridViewRow = CType(btnIncluir.NamingContainer, GridViewRow)

        Dim ds As DataSet = Session("DsMemorando")
        Dim row As DataRow = ds.Tables("ClientesPendentes").Rows(Gridrow.RowIndex)

        Dim Emitente As [Lib].Negocio.Cliente
        Dim Comprovando As [Lib].Negocio.Cliente

        If row("ES") = "S" Then
            Emitente = New [Lib].Negocio.Cliente(row("Cliente"), row("EndCliente"))
            Comprovando = New [Lib].Negocio.Cliente(row("Empresa"), row("EndEmpresa"))

        Else
            Emitente = New [Lib].Negocio.Cliente(row("Empresa"), row("EndEmpresa"))
            Comprovando = New [Lib].Negocio.Cliente(row("Cliente"), row("EndCliente"))
        End If

        Session("objClienteEmitente" & HID.Value) = Emitente
        Session("objClienteComprovando" & HID.Value) = Comprovando
        Session("objProdutoMeM" & HID.Value) = New [Lib].Negocio.Produto(row("Produto"))

        CarregarEmitente()
        CarregarComprovando()
        CarregarProduto()
        ddlSafraSelecaoNotas.SelectedValue = ddlSafra.SelectedValue
        'lnkAlterar.Parent.Visible = False
        'lnkExcluir.Parent.Visible = False
        lnkNovo.Parent.Visible = True
        TBMemorando.ActiveTabIndex = 1
    End Sub

    Protected Sub btnConsultaProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ucConsultaProduto.Limpar()
        Session("Where" & HID.Value) = "Situacao = 1"
        Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
        Popup.ConsultaDeProduto(Me.Page, "objProdutoMemCons" & HID.Value, txtNome.ClientID, True)
    End Sub

    Protected Sub btnComprovandoConsulta_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'Popup.ConsultarClientes(Me, "Comp", txtComprovandoConsulta, HComprovandoConsulta, "")
        ucConsultaClientes.Limpar()
        Popup.ConsultaDeClientes(Me.Page, "objClienteComprovCons" & HID.Value, "txtNome")
    End Sub

    Protected Sub ImgBuscaMais_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaMemorando()
        If objMemorando.NossaEmissao And objMemorando.NumeroNota = 0 Then
            MsgBox(Me.Page, "Selecione a nota de Comprovacao")
            Exit Sub
        End If
        objMemorando.NotasComprovadas.CarregarNotasParaSelecao(txtProcuraMaisNotas.Text, ddlSafraSelecaoNotas.SelectedValue)
        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub imgImpressao_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Limpar()
        Dim btnImprimir As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(btnImprimir.NamingContainer, GridViewRow)

        Dim ds As DataSet = Session("DsMemorandoPesquisa")
        Dim row As DataRow = ds.Tables("Memorandos").Rows(Gridrow.RowIndex)

        Dim Memorando As New [Lib].Negocio.MemorandoDeExportacao(row("Emitente"), row("EndEmitente"), row("memorando"))
        objMemorando = Memorando
        SessaoSalvaMemorando()
        Session("MemorandoOriginal") = New [Lib].Negocio.MemorandoDeExportacao(objMemorando.CodigoEmpresa, objMemorando.EnderecoEmpresa, objMemorando.CodigoMemorando)
        CarregarFormularioComAClasse()

        SessaoRecuperaMemorando()

        'Carregar memorando
        If ValidarCarregarMemorando() Then
            Dim Relatorio As New [Lib].Negocio.MemorandoDeExportacaoEspelho
            Relatorio.ExibirEspelho(Me, objMemorando)
        End If

    End Sub

    Protected Sub imgImpressaoPend_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Limpar()
        Dim sql As String
        Dim dsRel As DataSet

        Dim btnImprimir As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(btnImprimir.NamingContainer, GridViewRow)

        Dim ds As DataSet = Session("DsMemorando")

        Dim rowCliente As DataRow = ds.Tables("ClientesPendentes").Rows(Gridrow.RowIndex)

        'CLIENTE PENDENTES E DADOS DO MESMO
        sql = "SELECT NF.Empresa_id as Empresa," & vbCrLf & _
                   "       NF.EndEmpresa_id as EndEmpresa," & vbCrLf & _
                   "       Emp.Nome as NomeEmpresa," & vbCrLf & _
                   "       Emp.Cidade + '-' + Emp.Estado as CidadeEmpresa," & vbCrLf & _
                   "       NF.Cliente_Id as Cliente," & vbCrLf & _
                   "       NF.EndCliente_Id as EndCliente," & vbCrLf & _
                   "       C.Nome as NomeCliente," & vbCrLf & _
                   "       C.Cidade + '-' + C.Estado as Cidade," & vbCrLf & _
                   "       NF.EntradaSaida_Id as ES," & vbCrLf & _
                   "       NFxI.Produto_Id as Produto," & vbCrLf & _
                   "       Prd.Nome as NomeProduto, " & vbCrLf & _
                   "       sum(NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(sbMem.Quantidade,0)) as Saldo" & vbCrLf & _
                   "  FROM NotasFiscais AS NF" & vbCrLf & _
                   " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                   "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                   "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                   "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                   "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                   "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                   "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                   "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                   " INNER JOIN Pedidos P " & vbCrLf & _
                   "    ON P.Empresa_id    = NF.Empresa_id" & vbCrLf & _
                   "   AND P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf & _
                   "   AND P.Pedido_id     = NF.Pedido" & vbCrLf & _
                   " INNER JOIN Clientes C" & vbCrLf & _
                   "    ON C.Cliente_id  = NF.Cliente_Id" & vbCrLf & _
                   "   AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
                    " INNER JOIN Clientes Emp " & vbCrLf & _
                   "    ON Emp.Cliente_id  = NF.Empresa_Id" & vbCrLf & _
                   "   AND Emp.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf & _
                   " INNER JOIN Produtos Prd" & vbCrLf & _
                   "    on Prd.Produto_Id = NFxI.Produto_id" & vbCrLf & _
                   " INNER JOIN Operacoes OP" & vbCrLf & _
                   "    ON NFxI.Operacao    = OP.Operacao_Id " & vbCrLf & _
                   " INNER JOIN SubOperacoes SO" & vbCrLf & _
                   "    ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf & _
                   "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                   "  Left Join (" & vbCrLf & _
                   "		 	 SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
                   "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
                   "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
                   "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
                   "                    nfd.EntradaSaida_Id," & vbCrLf & _
                   "                    nfd.Serie_Id," & vbCrLf & _
                   "                    nfd.Nota_Id," & vbCrLf & _
                   "                    nfd.Produto_Id," & vbCrLf & _
                   "                    nfd.Sequencia_Id," & vbCrLf & _
                   "                    nfd.CFOP_Id," & vbCrLf & _
                   "                    sum(nfd.Quantidade) as Quantidade" & vbCrLf & _
                   "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
                   "              Inner Join NotasFiscais nf" & vbCrLf & _
                   "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
                   "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
                   "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
                   "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
                   "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf & _
                   "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf & _
                   "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf & _
                   "              Where nf.situacao        in (1, 4)" & vbCrLf & _
                   "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id," & vbCrLf & _
                   "                       nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf & _
                   "             ) SbDev" & vbCrLf & _
                   "    On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
                   "   and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
                   "   and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
                   "   and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
                   "   and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
                   "   and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
                   "   and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
                   "   and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
                   "   and NFxI.Sequencia_id    = SbDev.Sequencia_Id" & vbCrLf & _
                   "   and NFxI.CFOP_id         = SbDev.CFOP_Id" & vbCrLf & _
                   "  LEFT JOIN (" & vbCrLf & _
                   "             SELECT MExNF.Empresa_Id," & vbCrLf & _
                   "                    MExNF.EndEmpresa_Id," & vbCrLf & _
                   "                    MExNF.Cliente_Id," & vbCrLf & _
                   "                    MExNF.EndCliente_Id," & vbCrLf & _
                   "                    MExNF.EntradaSaida_Id," & vbCrLf & _
                   "                    MExNF.Serie_Id," & vbCrLf & _
                   "                    MExNF.Nota_Id," & vbCrLf & _
                   "                    Mem.Produto as Produto_Id," & vbCrLf & _
                   "                    sum(MExNF.Quantidade) as Quantidade" & vbCrLf & _
                   " 		      FROM MemorandoDeExportacao Mem" & vbCrLf & _
                   "              Inner Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
                   "			 	    ON Mem.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
                   "				   AND Mem.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
                   "			  	   AND Mem.Memorando_Id           = MExNF.Memorando_Id" & vbCrLf & _
                   "              Inner Join NotasFiscais NF" & vbCrLf & _
                   "                 On NF.Empresa_id      = MExNF.Empresa_Id" & vbCrLf & _
                   "                and NF.EndEmpresa_Id   = MExNF.EndEmpresa_Id" & vbCrLf & _
                   "                and NF.Cliente_Id      = MExNF.Cliente_Id" & vbCrLf & _
                   "                and NF.EndCliente_Id   = MExNF.EndCliente_Id" & vbCrLf & _
                   "                and NF.EntradaSaida_Id = MExNF.EntradaSaida_Id" & vbCrLf & _
                   "                and NF.Serie_Id        = MExNF.Serie_Id" & vbCrLf & _
                   "                and NF.Nota_Id         = MExNF.Nota_Id" & vbCrLf & _
                   "              Where NF.situacao        in (1,4)" & vbCrLf & _
                   "             Group by MExNF.Empresa_Id, MExNF.EndEmpresa_Id, MExNF.Cliente_Id, MExNF.EndCliente_Id, MExNF.EntradaSaida_Id," & vbCrLf & _
                   "                      MExNF.Serie_Id, MExNF.Nota_Id, Mem.Produto" & vbCrLf & _
                   "          ) SbMem" & vbCrLf & _
                   "    On NFxI.Empresa_id      = SbMem.Empresa_Id" & vbCrLf & _
                   "   and NFxI.EndEmpresa_Id   = SbMem.EndEmpresa_Id" & vbCrLf & _
                   "   and NFxI.Cliente_Id      = SbMem.Cliente_Id" & vbCrLf & _
                   "   and NFxI.EndCliente_Id   = SbMem.EndCliente_Id" & vbCrLf & _
                   "   and NFxI.EntradaSaida_Id = SbMem.EntradaSaida_Id" & vbCrLf & _
                   "   and NFxI.Serie_Id        = SbMem.Serie_Id" & vbCrLf & _
                   "   and NFxI.Nota_Id         = SbMem.Nota_Id" & vbCrLf & _
                   "   and NFxI.Produto_Id      = SbMem.Produto_Id" & vbCrLf & _
                   " Where SO.Memorando    = 1" & vbCrLf & _
                   "   and SO.Devolucao    = 'N'" & vbCrLf & _
                   "   and not(OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and SO.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf & _
                   "   and NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(SbMem.Quantidade,0) > 0" & vbCrLf

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "   and P.Safra = '" & ddlSafra.SelectedValue & "'"
        End If

        sql &= "   and nf.Empresa_id      ='" & rowCliente("Empresa") & "'" & vbCrLf & _
             "   and nf.EndEmpresa_Id   = " & rowCliente("EndEmpresa") & vbCrLf & _
             "   and nf.Cliente_id      ='" & rowCliente("Cliente") & "'" & vbCrLf & _
             "   and nf.EndCliente_Id   = " & rowCliente("EndCliente") & vbCrLf & _
             "   and nf.EntradaSaida_Id ='" & rowCliente("ES") & "'" & vbCrLf


        sql &= " Group by NF.Empresa_Id," & vbCrLf & _
               "          NF.EndEmpresa_Id," & vbCrLf & _
               "          Emp.Nome, " & vbCrLf & _
               "          Emp.Cidade + '-' + Emp.Estado, " & vbCrLf & _
               "          NF.Cliente_Id, " & vbCrLf & _
               "          NF.EndCliente_Id," & vbCrLf & _
               "          C.Nome," & vbCrLf & _
               "          C.Cidade + '-' + C.Estado," & vbCrLf & _
               "          NF.EntradaSaida_Id," & vbCrLf & _
               "          NFxI.Produto_Id," & vbCrLf & _
               "          Prd.Nome" & vbCrLf

        dsRel = Banco.ConsultaDataSet(sql, "ClientesPendentes")

        'NOTAS PENDENTES DO CLIENTE SELECIONADO
        sql = "SELECT NF.Empresa_id as Empresa," & vbCrLf & _
                      "       NF.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
                      "       NF.Cliente_Id as Cliente," & vbCrLf & _
                      "       NF.EndCliente_Id as EndCliente," & vbCrLf & _
                      "       NF.EntradaSaida_Id as ES," & vbCrLf & _
                      "       NF.Serie_Id as Serie," & vbCrLf & _
                      "       NF.Nota_Id as Nota," & vbCrLf & _
                      "       NF.DataDaNota," & vbCrLf & _
                      "       convert(int,getdate() - NF.DataDaNota) as DiasDecorridos," & vbCrLf & _
                      "       NFxI.QuantidadeFiscal as QuantidadeNota," & vbCrLf & _
                      "       isnull(sbDev.Quantidade,0) as QtdeDevolvida," & vbCrLf & _
                      "       isnull(sbMem.Quantidade,0) as QtdeJaComprovada," & vbCrLf & _
                      "       NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(sbMem.Quantidade,0) as Saldo" & vbCrLf & _
                      "  FROM NotasFiscais AS NF" & vbCrLf & _
                      " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                      "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                      "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                      "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                      "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                      "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                      "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                      "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                      " INNER JOIN Pedidos P " & vbCrLf & _
                      "    ON P.Empresa_id    = NF.Empresa_id" & vbCrLf & _
                      "   AND P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf & _
                      "   AND P.Pedido_id     = NF.Pedido" & vbCrLf & _
                      " INNER JOIN Operacoes OP" & vbCrLf & _
                      "    ON NFxI.Operacao    = OP.Operacao_Id " & vbCrLf & _
                      " INNER JOIN SubOperacoes SO" & vbCrLf & _
                      "    ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf & _
                      "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
                      "  Left Join (" & vbCrLf & _
                      "		 	 SELECT nfd.EmpresaDevolucao_Id as Empresa_Id," & vbCrLf & _
                      "                    nfd.EndEmpresaDevolucao_Id as EndEmpresa_Id," & vbCrLf & _
                      "                    nfd.ClienteDevolucao_Id as Cliente_Id," & vbCrLf & _
                      "                    nfd.EndClienteDevolucao_Id as EndCliente_Id," & vbCrLf & _
                      "                    nfd.EntradaSaida_Id," & vbCrLf & _
                      "                    nfd.Serie_Id," & vbCrLf & _
                      "                    nfd.Nota_Id," & vbCrLf & _
                      "                    nfd.Produto_Id," & vbCrLf & _
                      "                    nfd.Sequencia_Id," & vbCrLf & _
                      "                    nfd.CFOP_Id," & vbCrLf & _
                      "                    sum(nfd.Quantidade) as Quantidade" & vbCrLf & _
                      "			      FROM NotaFiscalDevolucaoXNotaFiscal nfd" & vbCrLf & _
                      "              Inner Join NotasFiscais nf" & vbCrLf & _
                      "                 On nf.Empresa_id      = nfd.EmpresaDevolucao_Id" & vbCrLf & _
                      "                and nf.EndEmpresa_Id   = nfd.EndEmpresaDevolucao_Id" & vbCrLf & _
                      "                and nf.Cliente_Id      = nfd.ClienteDevolucao_Id" & vbCrLf & _
                      "                and nf.EndCliente_Id   = nfd.EndClienteDevolucao_Id" & vbCrLf & _
                      "                and nf.EntradaSaida_Id = nfd.EntradaSaidaDevolucao_Id" & vbCrLf & _
                      "                and nf.Serie_Id        = nfd.SerieDevolucao_Id" & vbCrLf & _
                      "                and nf.Nota_Id         = nfd.NotaDevolucao_Id" & vbCrLf & _
                      "              Where nf.situacao        in (1, 4)" & vbCrLf & _
                      "              Group by nfd.EmpresaDevolucao_Id, nfd.EndEmpresaDevolucao_Id, nfd.ClienteDevolucao_Id, nfd.EndClienteDevolucao_Id," & vbCrLf & _
                      "                       nfd.EntradaSaida_Id, nfd.Serie_Id, nfd.Nota_Id, nfd.Produto_Id, nfd.Sequencia_Id, nfd.cfop_id" & vbCrLf & _
                      "             ) SbDev" & vbCrLf & _
                      "    On NFxI.Empresa_id      = SbDev.Empresa_Id" & vbCrLf & _
                      "   and NFxI.EndEmpresa_Id   = SbDev.EndEmpresa_Id" & vbCrLf & _
                      "   and NFxI.Cliente_Id      = SbDev.Cliente_Id" & vbCrLf & _
                      "   and NFxI.EndCliente_Id   = SbDev.EndCliente_Id" & vbCrLf & _
                      "   and NFxI.EntradaSaida_Id = SbDev.EntradaSaida_Id" & vbCrLf & _
                      "   and NFxI.Serie_Id        = SbDev.Serie_Id" & vbCrLf & _
                      "   and NFxI.Nota_Id         = SbDev.Nota_Id" & vbCrLf & _
                      "   and NFxI.Produto_Id      = SbDev.Produto_Id" & vbCrLf & _
                      "   and NFxI.Sequencia_id    = SbDev.Sequencia_Id" & vbCrLf & _
                      "   and NFxI.CFOP_id         = SbDev.CFOP_Id" & vbCrLf & _
                      "  LEFT JOIN (" & vbCrLf & _
                      "             SELECT MExNF.Empresa_Id," & vbCrLf & _
                      "                    MExNF.EndEmpresa_Id," & vbCrLf & _
                      "                    MExNF.Cliente_Id," & vbCrLf & _
                      "                    MExNF.EndCliente_Id," & vbCrLf & _
                      "                    MExNF.EntradaSaida_Id," & vbCrLf & _
                      "                    MExNF.Serie_Id," & vbCrLf & _
                      "                    MExNF.Nota_Id," & vbCrLf & _
                      "                    Mem.Produto as Produto_Id," & vbCrLf & _
                      "                    sum(MExNF.Quantidade) as Quantidade" & vbCrLf & _
                      " 		         FROM MemorandoDeExportacao Mem" & vbCrLf & _
                      "              Inner Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
                      "			 	   ON Mem.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
                      "				  AND Mem.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
                      "			  	  AND Mem.Memorando_Id           = MExNF.Memorando_Id" & vbCrLf & _
                      "              Inner Join NotasFiscais NF" & vbCrLf & _
                      "                 On NF.Empresa_id      = MExNF.Empresa_Id" & vbCrLf & _
                      "                and NF.EndEmpresa_Id   = MExNF.EndEmpresa_Id" & vbCrLf & _
                      "                and NF.Cliente_Id      = MExNF.Cliente_Id" & vbCrLf & _
                      "                and NF.EndCliente_Id   = MExNF.EndCliente_Id" & vbCrLf & _
                      "                and NF.EntradaSaida_Id = MExNF.EntradaSaida_Id" & vbCrLf & _
                      "                and NF.Serie_Id        = MExNF.Serie_Id" & vbCrLf & _
                      "                and NF.Nota_Id         = MExNF.Nota_Id" & vbCrLf & _
                      "              Where NF.situacao        in (1,4)" & vbCrLf & _
                      "             Group by MExNF.Empresa_Id, MExNF.EndEmpresa_Id, MExNF.Cliente_Id, MExNF.EndCliente_Id, MExNF.EntradaSaida_Id," & vbCrLf & _
                      "                      MExNF.Serie_Id, MExNF.Nota_Id, Mem.Produto" & vbCrLf & _
                      "          ) SbMem" & vbCrLf & _
                      "    On NFxI.Empresa_id      = SbMem.Empresa_Id" & vbCrLf & _
                      "   and NFxI.EndEmpresa_Id   = SbMem.EndEmpresa_Id" & vbCrLf & _
                      "   and NFxI.Cliente_Id      = SbMem.Cliente_Id" & vbCrLf & _
                      "   and NFxI.EndCliente_Id   = SbMem.EndCliente_Id" & vbCrLf & _
                      "   and NFxI.EntradaSaida_Id = SbMem.EntradaSaida_Id" & vbCrLf & _
                      "   and NFxI.Serie_Id        = SbMem.Serie_Id" & vbCrLf & _
                      "   and NFxI.Nota_Id         = SbMem.Nota_Id" & vbCrLf & _
                      "   and NFxI.Produto_Id      = SbMem.Produto_Id" & vbCrLf & _
                      " Where SO.Memorando    = 1" & vbCrLf & _
                      "   and SO.Devolucao    = 'N'" & vbCrLf & _
                      "   and not(OP.Classe = '" & eClassesOperacoes.VENDAS.ToString & "' and SO.Classe = '" & eClassesOperacoes.EXPORTACOES.ToString & "')" & vbCrLf & _
                      "   and NFxI.QuantidadeFiscal - isnull(sbDev.Quantidade,0) - isnull(SbMem.Quantidade,0) > 0" & vbCrLf

        If ddlSafra.SelectedIndex > 0 Then
            sql &= "   and P.Safra = '" & ddlSafra.SelectedValue & "'"
        End If


        sql &= "   and nf.Empresa_id      ='" & rowCliente("Empresa") & "'" & vbCrLf & _
               "   and nf.EndEmpresa_Id   = " & rowCliente("EndEmpresa") & vbCrLf & _
               "   and nf.Cliente_id      ='" & rowCliente("Cliente") & "'" & vbCrLf & _
               "   and nf.EndCliente_Id   = " & rowCliente("EndCliente") & vbCrLf & _
               "   and nf.EntradaSaida_Id ='" & rowCliente("ES") & "'" & vbCrLf & _
               " order by NF.nota_id" & vbCrLf

        dsRel.Merge(Banco.ConsultaDataSet(sql, "NotasPendentes"))

        'Imagem
        Dim dtImagem As DataTable = dsRel.Tables.Add("Images")
        dtImagem.Columns.Add("path", GetType(String))
        dtImagem.Columns.Add("image", GetType(System.Byte()))

        Dim drImagem As DataRow = dtImagem.NewRow()
        Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

        drImagem("path") = strCaminhoImagem
        drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        dtImagem.Rows.Add(drImagem)

        Dim crpt As New ReportDocument()
        crpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_MemorandoDeExportacaoPendencias.rpt")
        crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/" & rowCliente("Empresa") & "-" & rowCliente("Cliente") & "-Ctrl" & (New Random).Next & ".PDF"
        Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        Try
            crpt.SetDataSource(dsRel)

            'Dim crparametervalues As ParameterValues
            'Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            'Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition
            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub imgImpPenComprov_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Limpar()
        Dim Safra As String = ddlSafra.SelectedValue
        Dim sql As String

        sql = "Select NFxI.Empresa_Id," & vbCrLf & _
              "       NFxI.EndEmpresa_Id," & vbCrLf & _
              "       Emp.Nome as NomeEmpresa," & vbCrLf & _
              "       Emp.Cidade + '-' + Emp.Estado as CidadeEmpresa," & vbCrLf & _
              "       NFxI.Cliente_Id," & vbCrLf & _
              "       NFxI.EndCliente_Id," & vbCrLf & _
              "       C.Nome as NomeCliente," & vbCrLf & _
              "       C.Cidade + '-' + C.Estado as CidadeCliente," & vbCrLf & _
              "       NFxI.EntradaSaida_Id," & vbCrLf & _
              "       NFxI.Produto_Id," & vbCrLf & _
              "       Prd.Nome as NomeProduto, " & vbCrLf & _
              "       NFxI.Nota_Id," & vbCrLf & _
              "       NFxI.Serie_Id," & vbCrLf & _
              "       NFxI.QuantidadeFiscal as QtdeNota," & vbCrLf & _
              "       isnull(sbMem.QtdeComprovadaNesteMemorando,0) as QtdeComprovadaNesteMemorando," & vbCrLf & _
              "       isnull(sbMem.QtdeJaComprovada,0) as QtdeJaComprovada" & vbCrLf & _
              "  from notasfiscais NF" & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
              " inner join Operacoes OP" & vbCrLf & _
              "    on OP.Operacao_id = NFxI.Operacao" & vbCrLf & _
              " inner Join SubOperacoes SO" & vbCrLf & _
              "    on SO.Operacao_Id     = NF.Operacao" & vbCrLf & _
              "   and SO.SubOperacoes_Id = NFxI.SubOperacao" & vbCrLf & _
              " Inner Join Pedidos P" & vbCrLf & _
              "    on P.Empresa_id    = NF.Empresa_Id" & vbCrLf & _
              "   and P.EndEmpresa_id = NF.EndEmpresa_Id" & vbCrLf & _
              "   and P.Pedido_id     = NF.Pedido" & vbCrLf & _
              " INNER JOIN Clientes C" & vbCrLf & _
              "    ON C.Cliente_id  = NF.Cliente_Id" & vbCrLf & _
              "   AND C.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
              " INNER JOIN Clientes Emp " & vbCrLf & _
              "    ON Emp.Cliente_id  = NF.Empresa_Id" & vbCrLf & _
              "   AND Emp.Endereco_Id = NF.EndEmpresa_Id" & vbCrLf & _
              " INNER JOIN Produtos Prd" & vbCrLf & _
              "    on Prd.Produto_Id = NFxI.Produto_id" & vbCrLf & _
              "  Left Join (" & vbCrLf & _
              "			 SELECT M.Empresa," & vbCrLf & _
              "			 	    M.EndEmpresa," & vbCrLf & _
              "					M.Cliente," & vbCrLf & _
              "					M.EndCliente," & vbCrLf & _
              "					M.EntradaSaida," & vbCrLf & _
              "					M.Nota," & vbCrLf & _
              "					M.Serie," & vbCrLf & _
              "					M.Produto," & vbCrLf

        sql &= "                  0 as QtdeComprovadaNesteMemorando," & vbCrLf


        sql &= "					SUM(MxNF.Quantidade) QtdeJaComprovada" & vbCrLf & _
               "			   FROM MemorandoDeExportacao M" & vbCrLf & _
               "			  INNER JOIN MemorandoDeExportacaoXNotaFiscal MxNF" & vbCrLf & _
               "			     ON M.EmpresaMemorando_Id    = MxNF.EmpresaMemorando_Id" & vbCrLf & _
               "			    AND M.EndEmpresaMemorando_Id = MxNF.EndEmpresaMemorando_Id" & vbCrLf & _
               "			    AND M.Memorando_Id           = MxNF.Memorando_Id" & vbCrLf & _
               "              Group by M.Empresa," & vbCrLf & _
               "			 	       M.EndEmpresa," & vbCrLf & _
               "					   M.Cliente," & vbCrLf & _
               "					   M.EndCliente," & vbCrLf & _
               "					   M.EntradaSaida," & vbCrLf & _
               "					   M.Nota," & vbCrLf & _
               "					   M.Serie," & vbCrLf & _
               "					   M.Produto" & vbCrLf & _
               "             )sbMem" & vbCrLf & _
               "    ON NFxI.Empresa_Id      = sbMem.Empresa " & vbCrLf & _
               "   AND NFxI.EndEmpresa_Id   = sbMem.EndEmpresa" & vbCrLf & _
               "   AND NFxI.Cliente_Id      = sbMem.Cliente" & vbCrLf & _
               "   AND NFxI.EndCliente_Id   = sbMem.EndCliente" & vbCrLf & _
               "   AND NFxI.EntradaSaida_Id = sbMem.EntradaSaida" & vbCrLf & _
               "   AND NFxI.Serie_Id        = sbMem.Serie" & vbCrLf & _
               "   AND NFxI.Nota_Id         = sbMem.Nota" & vbCrLf & _
               "   AND NFxI.Produto_Id      = sbMem.Produto" & vbCrLf & _
               " where OP.Classe             = '" & eClassesOperacoes.VENDAS.ToString & "'" & vbCrLf & _
               "   and SO.Memorando          = 1" & vbCrLf & _
               "   and SO.Classe             ='" & eClassesOperacoes.EXPORTACOES.ToString & "'" & vbCrLf & _
               "   and NFxI.QuantidadeFiscal > 0" & vbCrLf

        If Safra.Length > 0 Then
            sql &= "   and P.Safra = '" & Safra & "'" & vbCrLf
        End If

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Notas")

        gridNotasComprovacaoPendentes.DataSource = ds
        gridNotasComprovacaoPendentes.DataBind()

        'Imagem
        Dim dtImagem As DataTable = ds.Tables.Add("Images")
        dtImagem.Columns.Add("path", GetType(String))
        dtImagem.Columns.Add("image", GetType(System.Byte()))

        Dim drImagem As DataRow = dtImagem.NewRow()
        Dim strCaminhoImagem As String = HttpContext.Current.Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

        drImagem("path") = strCaminhoImagem
        drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
        dtImagem.Rows.Add(drImagem)

        Dim crpt As New ReportDocument()
        crpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_NossasPendenciasCompDeExportacao.rpt")
        crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/NossasPendComprovExp" & HttpContext.Current.Session("ssEmpresa") & "-Ctrl" & (New Random).Next & ".PDF"
        Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        Try
            crpt.SetDataSource(ds)

            'Dim crparametervalues As ParameterValues
            'Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            'Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub ImgAddRegExp_Click1(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgAddRegExp.Click
        If ValidaRegExp() Then
            SessaoRecuperaMemorando()
            Dim objMemoExpXRegExp As New [Lib].Negocio.MemorandoDeExportacaoXRegistroDeExportacao(objMemorando)
            objMemoExpXRegExp.CodRegistroDeExportacao = TxtRegExportacao.Text
            objMemoExpXRegExp.DataRegExportacao = TxtDataRegExp.Text
            objMemoExpXRegExp.UfProdutor = ddlEstado.SelectedValue
            objMemorando.RegistrosDeExportacao.Add(objMemoExpXRegExp)
            SessaoSalvaMemorando()


            GridRegistroExp.DataSource = objMemorando.RegistrosDeExportacao.ToArray
            GridRegistroExp.DataBind()

            TxtRegExportacao.Text = ""
            TxtDataRegExp.Text = ""
        End If
    End Sub

    Protected Sub ImgExcluirRegExp_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaMemorando()

        Dim ImgExcluirRegExp As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(ImgExcluirRegExp.NamingContainer, GridViewRow)

        objMemorando.RegistrosDeExportacao.RemoveAt(Gridrow.RowIndex)
        SessaoSalvaMemorando()


        GridRegistroExp.DataSource = objMemorando.RegistrosDeExportacao.ToArray
        GridRegistroExp.DataBind()

        TxtRegExportacao.Text = ""
        TxtDataRegExp.Text = ""
    End Sub

    Protected Sub ImgBuscaNotasEquipExp_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Dim mem As New [Lib].Negocio.MemorandoDeExportacao(objMemorando.CodExportadorEquiparado, objMemorando.EnderecoExportadorEquiparado, "")

        mem.NotasComprovadas.CarregarNotasParaSelecao(txtProcuraMaisNotas.Text, ddlSafraSelecaoNotas.SelectedValue)
        gridNotas.DataSource = mem.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub ImgAddConhec_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles ImgAddConhec.Click
        If ValidaConhecimento() Then
            SessaoRecuperaMemorando()
            Dim objMemoExpXConhec As New [Lib].Negocio.MemorandoDeExportacaoXConhecimento(objMemorando)
            objMemoExpXConhec.NumConhecimentoDeEmbarque = txtNumeroConhecimento.Text
            objMemoExpXConhec.DataConhecimento = txtDataConhecimento.Text
            objMemoExpXConhec.TipoConhecimento = ddlTipoConhec.SelectedValue
            objMemorando.ConhecimentosDeEmbarque.Add(objMemoExpXConhec)
            SessaoSalvaMemorando()


            GridConhecimento.DataSource = objMemorando.ConhecimentosDeEmbarque.ToArray
            GridConhecimento.DataBind()

            txtNumeroConhecimento.Text = ""
            txtDataConhecimento.Text = ""
            ddlTipoConhec.SelectedValue = 10
        End If
    End Sub

    Protected Sub ImgExcluirConhec_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        SessaoRecuperaMemorando()

        Dim ImgExcluirConhec As ImageButton = CType(sender, ImageButton)
        Dim Gridrow As GridViewRow = CType(ImgExcluirConhec.NamingContainer, GridViewRow)

        objMemorando.ConhecimentosDeEmbarque.RemoveAt(Gridrow.RowIndex)
        SessaoSalvaMemorando()


        GridConhecimento.DataSource = objMemorando.ConhecimentosDeEmbarque.ToArray
        GridConhecimento.DataBind()

        txtNumeroConhecimento.Text = ""
        txtDataConhecimento.Text = ""
        ddlTipoConhec.SelectedValue = 10
    End Sub

    Protected Sub btnConsultaNota_Click(sender As Object, e As EventArgs) Handles btnConsultaNota.Click
        SessaoRecuperaMemorando()
        If objMemorando.NossaEmissao And objMemorando.NumeroNota = 0 Then
            MsgBox(Me.Page, "Selecione a nota de Comprovacao")
            Exit Sub
        End If
        objMemorando.NotasComprovadas.Clear()
        objMemorando.NotasComprovadas.CarregarNotasParaSelecao(txtProcura.Text, ddlSafraSelecaoNotas.SelectedValue)
        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub btnBuscaMaisNotas_Click(sender As Object, e As EventArgs) Handles btnBuscaMaisNotas.Click
        SessaoRecuperaMemorando()
        If objMemorando.NossaEmissao And objMemorando.NumeroNota = 0 Then
            MsgBox(Me.Page, "Selecione a nota de Comprovacao")
            Exit Sub
        End If
        objMemorando.NotasComprovadas.CarregarNotasParaSelecao(txtProcuraMaisNotas.Text, ddlSafraSelecaoNotas.SelectedValue)
        gridNotas.DataSource = objMemorando.NotasComprovadas
        gridNotas.DataBind()
    End Sub

    Protected Sub btnAddRegExportacao_Click(sender As Object, e As EventArgs) Handles btnAddRegExportacao.Click
        If ValidaRegExp() Then
            SessaoRecuperaMemorando()
            Dim objMemoExpXRegExp As New [Lib].Negocio.MemorandoDeExportacaoXRegistroDeExportacao(objMemorando)
            objMemoExpXRegExp.CodRegistroDeExportacao = TxtRegExportacao.Text
            objMemoExpXRegExp.DataRegExportacao = TxtDataRegExp.Text
            objMemoExpXRegExp.UfProdutor = ddlEstado.SelectedValue
            objMemorando.RegistrosDeExportacao.Add(objMemoExpXRegExp)
            SessaoSalvaMemorando()


            GridRegistroExp.DataSource = objMemorando.RegistrosDeExportacao.ToArray
            GridRegistroExp.DataBind()

            TxtRegExportacao.Text = String.Empty
            TxtDataRegExp.Text = String.Empty
        End If
    End Sub

    Protected Sub btnAddConhecimento_Click(sender As Object, e As EventArgs) Handles btnAddConhecimento.Click
        If ValidaConhecimento() Then
            SessaoRecuperaMemorando()
            Dim objMemoExpXConhec As New [Lib].Negocio.MemorandoDeExportacaoXConhecimento(objMemorando)
            objMemoExpXConhec.NumConhecimentoDeEmbarque = txtNumeroConhecimento.Text
            objMemoExpXConhec.DataConhecimento = txtDataConhecimento.Text
            objMemoExpXConhec.TipoConhecimento = ddlTipoConhec.SelectedValue
            objMemorando.ConhecimentosDeEmbarque.Add(objMemoExpXConhec)
            SessaoSalvaMemorando()


            GridConhecimento.DataSource = objMemorando.ConhecimentosDeEmbarque.ToArray
            GridConhecimento.DataBind()

            txtNumeroConhecimento.Text = ""
            txtDataConhecimento.Text = ""
            ddlTipoConhec.SelectedValue = 10
        End If
    End Sub

    Protected Sub lnkConsulta_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsulta.Click
        Dim sql As String
        sql = "select ME.EmpresaMemorando_Id as Emitente," & vbCrLf & _
              "       ME.EndEmpresaMemorando_Id as EndEmitente," & vbCrLf & _
              "       C.Nome as NomeEmitente," & vbCrLf & _
              "       C.Cidade + '-' + C.Estado as CidadeUfEmitente," & vbCrLf & _
              "       ME.Memorando_id as Memorando," & vbCrLf & _
              "       ME.DataMemorando," & vbCrLf & _
              "       SUM(MExNF.Quantidade) as Quantidade" & vbCrLf & _
              "  from MemorandoDeExportacao ME" & vbCrLf & _
              " Left Join MemorandoDeExportacaoXNotaFiscal MExNF" & vbCrLf & _
              "    on ME.EmpresaMemorando_Id    = MExNF.EmpresaMemorando_Id" & vbCrLf & _
              "   and ME.EndEmpresaMemorando_Id = MExNF.EndEmpresaMemorando_Id" & vbCrLf & _
              "   and ME.Memorando_id           = MExNF.Memorando_id" & vbCrLf & _
              " Inner Join Clientes C" & vbCrLf & _
              "    on C.Cliente_id  = ME.EmpresaMemorando_Id" & vbCrLf & _
              "   and C.Endereco_Id = ME.EndEmpresaMemorando_Id" & vbCrLf & _
              " Where ME.DataMemorando between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf

        If txtEmitenteConsulta.Text.Length > 0 Then
            Dim Em As String() = HEmitenteConsulta.Value.Split("-")
            If chkConsolidarEmitente.Checked Then
                sql &= "   and left(ME.EmpresaMemorando_Id,8)    = '" & Em(0).Substring(0, 8) & "'"
            Else
                sql &= "   and ME.EmpresaMemorando_Id    ='" & Em(0) & "'" & vbCrLf & _
                       "   and ME.EndEmpresaMemorando_Id = " & Em(1)
            End If
        End If

        If txtComprovandoConsulta.Text.Length > 0 Then
            Dim com As String() = HComprovandoConsulta.Value.Split("-")
            If chkConsolidaComprovando.Checked Then
                sql &= "   and left(ME.ClienteMemorando,8)    = '" & com(0).Substring(0, 8) & "'"
            Else
                sql &= "   and ME.ClienteMemorando    ='" & com(0) & "'" & vbCrLf & _
                       "   and ME.EndClienteMemorando = " & com(1)
            End If
        End If

        If txtProdutoConsulta.Text.Length > 0 Then
            sql &= "   and ME.Produto    ='" & txtProdutoConsulta.Text.Split("-")(0) & "'" & vbCrLf
        End If

        sql &= " Group by ME.EmpresaMemorando_Id," & vbCrLf & _
               "          ME.EndEmpresaMemorando_Id," & vbCrLf & _
               "          C.Nome," & vbCrLf & _
               "          C.Cidade + '-' + C.Estado," & vbCrLf & _
               "          ME.Memorando_id," & vbCrLf & _
               "          ME.DataMemorando" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Memorandos")
        gridConsulta.DataSource = ds.Tables(0)
        gridConsulta.DataBind()

        Session("DsMemorandoPesquisa") = ds

        If ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Não foi Encontrado Nenhum Registro")
        End If
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        txtEmitenteConsulta.Text = ""
        HEmitenteConsulta.Value = ""
        txtComprovandoConsulta.Text = ""
        HComprovandoConsulta.Value = ""
        txtProdutoConsulta.Text = ""
        txtDataInicial.Text = ("01/01/" & Now.Year)
        txtDataFinal.Text = ("31/12/" & Now.Year)
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        SessaoRecuperaMemorando()
        If ValidarCarregarMemorando() Then
            objMemorando.IUD = "I"
            objMemorando.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")        'Usuario Que Esta Incluindo
            objMemorando.DataInclusao = Format(CDate(Today), "yyyy/MM/dd")                     'Data da desta Inclusao
            If objMemorando.Salvar Then
                MsgBox(Me.Page, "Memorando Salvo com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, "Erro ao salvar o Memorando")
            End If

        End If
    End Sub

    Protected Sub lnkAlterar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAlterar.Click
        SessaoRecuperaMemorando()
        If ValidarCarregarMemorando() Then
            objMemorando.IUD = "U"
            objMemorando.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")        'Usuario Que Esta Incluindo
            objMemorando.DataAlteracao = Format(CDate(Today), "yyyy/MM/dd")                     'Data da desta Inclusao

            If objMemorando.Salvar Then
                MsgBox(Me.Page, "Memorando alterado com Sucesso.", eTitulo.Sucess)
                Limpar()
            Else
                MsgBox(Me.Page, "Erro ao alterar o Memorando")
            End If
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        SessaoRecuperaMemorando()
        If ValidarCarregarMemorando() Then
            Dim Relatorio As New [Lib].Negocio.MemorandoDeExportacaoEspelho
            Relatorio.ExibirEspelho(Me, objMemorando)
        End If
    End Sub

    Protected Sub lnkLimparCadastro_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimparCadastro.Click
        Limpar()
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        SessaoRecuperaMemorando()
        objMemorando.IUD = "D"
        If objMemorando.Salvar Then
            MsgBox(Me.Page, "Memorando excluído com Sucesso.", eTitulo.Sucess)
            Limpar()
            lnkConsulta_Click(lnkConsulta, Nothing)
            TBMemorando.ActiveTabIndex = 0
        End If
    End Sub

    Protected Sub btnExcluirRegExp_Click(sender As Object, e As EventArgs)
        SessaoRecuperaMemorando()

        Dim btnExcluirReg As Button = CType(sender, Button)
        Dim Gridrow As GridViewRow = CType(btnExcluirReg.NamingContainer, GridViewRow)

        objMemorando.RegistrosDeExportacao.RemoveAt(Gridrow.RowIndex)
        SessaoSalvaMemorando()


        GridRegistroExp.DataSource = objMemorando.RegistrosDeExportacao.ToArray
        GridRegistroExp.DataBind()

        TxtRegExportacao.Text = ""
        TxtDataRegExp.Text = ""
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "MemorandoDeExportacao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class
