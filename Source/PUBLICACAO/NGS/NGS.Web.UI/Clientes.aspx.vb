Imports System.Runtime.InteropServices
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports System.Drawing
Imports OfficeOpenXml.Style

Public Class Clientes
    Inherits BasePage

#Region "Fields"
    Private objCliente As New [Lib].Negocio.Cliente
    Private objMunicipio As [Lib].Negocio.Municipio
    Private objClienteXContato As [Lib].Negocio.ClienteXContato
    Private objClienteXDependente As [Lib].Negocio.ClientexDependente
    Private objClienteXContaBancaria As [Lib].Negocio.ClienteXContaBancaria
    Private objClienteXSafra As [Lib].Negocio.ClienteXSafra
    Private objClienteXVeiculo As [Lib].Negocio.ClienteXVeiculo
    Private objClienteXEquipamento As [Lib].Negocio.ClienteXEquipamento
    Private objClienteXImovel As [Lib].Negocio.ClienteXImovel
    Private objClienteXMatricula As [Lib].Negocio.ClienteXMatricula
    Private objClienteXArrendante As [Lib].Negocio.ClienteXArrendante
    Private objClienteXRepresentante As [Lib].Negocio.ClienteXRepresentante
    Private objClienteXSocio As [Lib].Negocio.ClienteXSocio
    Private objClienteXFinanciamento As [Lib].Negocio.ClienteXFinanciamento
    Private ObjClienteXReceitaDespesa As [Lib].Negocio.ClienteXReceitasDespesas
#End Region

#Region "Events"

    'LOAD *****************************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Gestao.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("Clientes", "ACESSAR") Then
                ddl.Carregar(ddlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "", False)
                ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "Situacao_id not in (7, 8, 9, 10)")
                ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlEstadoContaBancaria, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlEstadoImovel, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlEstadoMatricula, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlEstadoST, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlNaturalidade, CarregarDDL.Tabela.Estados, "", True)
                ddl.Carregar(ddlRegiao, CarregarDDL.Tabela.Regiao, "")
                ddl.Carregar(ddlPais, CarregarDDL.Tabela.Pais, "", True)
                ddl.Carregar(ddlCategoria, CarregarDDL.Tabela.CategoriaCliente, "", True)
                CarregarTiposDeCliente()
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "", True)
                ddl.Carregar(ddlSafraFinanciamento, CarregarDDL.Tabela.Safra, "", True)
                ddl.Carregar(ddlCultura, CarregarDDL.Tabela.Cultura, "", True)
                ddl.Carregar(ddlUnidadeImovel, CarregarDDL.Tabela.UnidadeDeMedida, "", True)
                ddl.Carregar(ddlGrupoProdutoFinanciamento, CarregarDDL.Tabela.GrupoProduto, "")
                ddl.Carregar(ddlMoedaFinanciamento, CarregarDDL.Tabela.Moeda, "")
                ddl.Carregar(ddlTipoVeiculo, CarregarDDL.Tabela.TipoDeVeiculo, "")
                ddl.Carregar(ddlTipoEquipamento, CarregarDDL.Tabela.TipoDeVeiculo, "")
                ddl.Carregar(ddlBancoContaBancaria, CarregarDDL.Tabela.Bancos, "")
                ddl.Carregar(ddlBancoContato, CarregarDDL.Tabela.Bancos, "")


                HID.Value = Guid.NewGuid().ToString()
                ucArrendante.SetarHID(HID.Value)
                ucConsultaCep.SetarHID(HID.Value)
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaCodMunicipios.SetarHID(HID.Value)
                ucConsultaEstados.SetarHID(HID.Value)
                ucConsultaCadastro.SetarHID(HID.Value)

                LimparCampos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
            End If
        Else
            If Not Request("__EVENTARGUMENT") Is Nothing Then
                If Request("__EVENTARGUMENT") = "MostrarImagem" Then verFoto()
            End If
        End If
    End Sub


    'CADASTRO *****************************************************************************************************************
    '**************************************************************************************************************************
    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            IniciarProcesso("I")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            IniciarProcesso("U")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            IniciarProcesso("D")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFichaCadastral_Click(sender As Object, e As EventArgs) Handles lnkFichaCadastral.Click
        Try
            'FICHA CADASTRAL EM BRANCO
            RecuperarSessionCliente()

            'CLIENTE
            Dim ds As New DataSet
            Dim tbClientes As New DataTable("Clientes")
            tbClientes.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Endereco_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Regiao_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Regiao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Categoria_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Categoria", Type.GetType("System.String"))
            tbClientes.Columns.Add("Estado_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Estado", Type.GetType("System.String"))
            tbClientes.Columns.Add("Pais", Type.GetType("System.String"))
            tbClientes.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientes.Columns.Add("Fantasia", Type.GetType("System.String"))
            tbClientes.Columns.Add("Endereco", Type.GetType("System.String"))
            tbClientes.Columns.Add("Numero", Type.GetType("System.String"))
            tbClientes.Columns.Add("Complemento", Type.GetType("System.String"))
            tbClientes.Columns.Add("Bairro", Type.GetType("System.String"))
            tbClientes.Columns.Add("CEP", Type.GetType("System.String"))
            tbClientes.Columns.Add("Cidade", Type.GetType("System.String"))
            tbClientes.Columns.Add("Inscricao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Telefone", Type.GetType("System.String"))
            tbClientes.Columns.Add("Fax", Type.GetType("System.String"))
            tbClientes.Columns.Add("email", Type.GetType("System.String"))
            tbClientes.Columns.Add("Reduzido", Type.GetType("System.String"))
            tbClientes.Columns.Add("CodigoMunicipio", Type.GetType("System.String"))
            tbClientes.Columns.Add("Situacao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Habilitacao", Type.GetType("System.String"))
            tbClientes.Columns.Add("EmailNFE", Type.GetType("System.String"))
            tbClientes.Columns.Add("RegistroAuxiliar", Type.GetType("System.String"))
            tbClientes.Columns.Add("FisicaJuridica", Type.GetType("System.String"))
            tbClientes.Columns.Add("RG", Type.GetType("System.String"))
            tbClientes.Columns.Add("Site", Type.GetType("System.String"))
            tbClientes.Columns.Add("Sexo", Type.GetType("System.String"))
            tbClientes.Columns.Add("NascimentoConstituicao", Type.GetType("System.String"))
            tbClientes.Columns.Add("NaturalidadeCidade", Type.GetType("System.String"))
            tbClientes.Columns.Add("NaturalidadeEstado", Type.GetType("System.String"))
            tbClientes.Columns.Add("ClienteDesde", Type.GetType("System.String"))
            ds.Tables.Add(tbClientes)

            'LOGOTIPO
            Dim tbLogotipo As DataTable = ds.Tables.Add("Logotipo")
            tbLogotipo.Columns.Add("path", GetType(String))
            tbLogotipo.Columns.Add("Imagem", GetType(System.Byte()))
            tbLogotipo.Columns.Add("Nome", GetType(String))
            tbLogotipo.Columns.Add("Cidade", GetType(String))
            tbLogotipo.Columns.Add("Estado_Id", GetType(String))
            Dim drImagem As DataRow = tbLogotipo.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            drImagem("path") = strCaminhoImagem
            drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            drImagem("Nome") = HttpContext.Current.Session("ssNomeEmpresa")
            drImagem("Cidade") = HttpContext.Current.Session("ssCidadeEmpresa")
            drImagem("Estado_Id") = HttpContext.Current.Session("ssEstadoEmpresa")
            tbLogotipo.Rows.Add(drImagem)

            'ARRENDAMENTOS
            Dim tbClientesXArrendamentos As New DataTable("ClientesXArrendamentos")
            tbClientesXArrendamentos.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("EndCliente_id", Type.GetType("System.Int32"))
            tbClientesXArrendamentos.Columns.Add("Proprietario_Id", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("EndProprietario_Id", Type.GetType("System.Int32"))
            tbClientesXArrendamentos.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("DataContrato", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Vencimento", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Matricula", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Observacoes", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXArrendamentos)

            'FINANCIAMENTOS
            Dim tbClientesXFinanciamentos As New DataTable("ClientesXFinanciamentos")
            tbClientesXFinanciamentos.Columns.Add("Financiador_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("EndFinanciador_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Fantasia", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Descricao", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Safra_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("DataFinanciamento", Type.GetType("System.DateTime"))
            tbClientesXFinanciamentos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
            tbClientesXFinanciamentos.Columns.Add("NumeroDeParcelas", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Produto", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("Moeda", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("Observacoes", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXFinanciamentos)

            'IMÓVEIS
            Dim tbClientesXImoveis As New DataTable("ClientesXImoveis")
            tbClientesXImoveis.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXImoveis.Columns.Add("Descricao_Id", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Onerado", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Cidade", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Estado", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("AreaTotal", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("UnidadeDeMedida", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("AreaConstruida", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("NumeroDoRegistro", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Cartorio", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("Observacoes", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("DataAvaliacao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXImoveis)

            'SAFRAS
            Dim tbClientesXSafras As New DataTable("ClientesXSafras")
            tbClientesXSafras.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXSafras.Columns.Add("Safra_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("Cultura_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("AreaPlantada", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Produtividade", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("ConsumoProprio", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Comprometimento", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("EstimativaDeEntrega", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Observacoes", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXSafras)

            'VEÍCULOS
            Dim tbClientesXVeiculos As New DataTable("ClientesXVeiculos")
            tbClientesXVeiculos.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXVeiculos.Columns.Add("Placa_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("TipoDeVeiculo_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("Ano", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("MarcaModelo", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("Fabricante", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXVeiculos.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXVeiculos.Columns.Add("DataAvaliacao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXVeiculos)

            'REPRESENTANTES
            Dim tbClientesXRepresentantes As New DataTable("ClientesXRepresentantes")
            tbClientesXRepresentantes.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXRepresentantes.Columns.Add("Endereco_Id", Type.GetType("System.String"))
            tbClientesXRepresentantes.Columns.Add("Nome", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXRepresentantes)

            'DEPENDENTES
            Dim tbClientesXDependentes As New DataTable("ClientesXDependentes")
            tbClientesXDependentes.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("TipoDeDependente", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("Identidade", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("CPF", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("DataNascimento", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("Profissao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXDependentes)

            'CONTATOS
            Dim tbContatos As New DataTable("Contatos")
            tbContatos.Columns.Add("NomeContato", Type.GetType("System.String"))
            tbContatos.Columns.Add("Funcao", Type.GetType("System.String"))
            tbContatos.Columns.Add("Telefone", Type.GetType("System.String"))
            tbContatos.Columns.Add("Email", Type.GetType("System.String"))
            ds.Tables.Add(tbContatos)

            'TIPOS DE CLIENTE
            Dim tbTiposDeClientes As New DataTable("TiposDeClientes")
            tbTiposDeClientes.Columns.Add("Tipo_Id", Type.GetType("System.String"))
            tbTiposDeClientes.Columns.Add("Descricao", Type.GetType("System.String"))
            ds.Tables.Add(tbTiposDeClientes)

            Dim param As New Dictionary(Of String, Object)
            param.Add("titulo", "Ficha Cadastral")
            param.Add("folha", "Folha.:")

            Funcoes.BindReport(Me.Page, ds, "Cr_FichaCadastralBranco", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF), param)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkDuplicar_Click(sender As Object, e As EventArgs) Handles lnkDuplicar.Click
        Try
            RecuperarSessionCliente()
            If objCliente Is Nothing Then
                MsgBox(Me.Page, "É necessário selecionar um registro para duplicar!")
            Else
                objCliente.CodigoEndereco = getEnderecoId(objCliente.Codigo)
                SalvarSessionCliente()
                CarregarCliente(objCliente)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "LEITURA") Then
                Dim verCliente As String = Funcoes.EliminarCaracteresEspeciais(txtCpfCnpj.Text)

                If String.IsNullOrWhiteSpace(txtReduzido.Text) AndAlso String.IsNullOrWhiteSpace(txtCpfCnpj.Text) AndAlso String.IsNullOrWhiteSpace(txtNome.Text) AndAlso String.IsNullOrWhiteSpace(txtNomeFantasia.Text) AndAlso String.IsNullOrWhiteSpace(ddlCategoria.SelectedValue) Then
                    MsgBox(Me.Page, "Informe um dos campos ('Reduzido', 'Cpf/Cnpj', 'Nome', 'Fantasia' ou 'categoria') para efetuar a Consulta.")
                ElseIf verCliente.Length = 11 AndAlso Not Funcoes.ValidaCPF(verCliente) Then
                    MsgBox(Me.Page, "Cpf " & verCliente & " inválido.")
                    txtCpfCnpj.Text = String.Empty
                ElseIf verCliente.Length = 14 AndAlso Not Funcoes.ValidaCNPJ(verCliente) Then
                    MsgBox(Me.Page, "CNPJ " & verCliente & " inválido.")
                    txtCpfCnpj.Text = String.Empty
                Else
                    Dim ListaClientes As New [Lib].Negocio.ListCliente("", Nothing, verCliente, txtNome.Text, txtNomeFantasia.Text, txtReduzido.Text, ddlCategoria.SelectedValue)
                    If ListaClientes.Count > 0 Then
                        Session("ssListClientes") = ListaClientes
                        gridConsultaCliente.DataSource = ListaClientes.ToArray()
                        gridConsultaCliente.DataBind()
                        btnImpressaoClientes.Visible = True
                        txtCodigoEndereco.Text = ListaClientes.Max(Function(x) x.CodigoEndereco) + IIf(verCliente.Length = 14, 0, 1)
                        TabContainer1.ActiveTabIndex = 19
                    Else
                        txtCodigoEndereco.Text = 0
                        btnImpressaoClientes.Visible = False

                        If verCliente.Length = 14 Then
                            MsgBox(Me.Page, "Não foi encontrado o CNPJ com o(s) parâmetro(s) informado(s) no Sistema. Vamos verificar se a Receita Federal disponibiliza essas informações.")
                        ElseIf verCliente.Length = 11 Then
                            MsgBox(Me.Page, "Não foi encontrado o CPF com o(s) parâmetro(s) informado(s) no Sistema.")
                        Else
                            MsgBox(Me.Page, "Não foi encontrado o CÓDIGO com o(s) parâmetro(s) informado(s) no Sistema.")
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            Relatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            Relatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlEmpresas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresas.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlEmpresas.SelectedValue) Then
                CarregarCliente(ddlEmpresas.SelectedValue.Split("-")(0), ddlEmpresas.SelectedValue.Split("-")(1))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgFicha_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgFicha.Click
        Try
            RecuperarSessionCliente()

            'CLIENTE'
            Dim ds As New DataSet
            Dim tbClientes As New DataTable("Clientes")
            tbClientes.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Endereco_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Regiao_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Regiao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Categoria_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Categoria", Type.GetType("System.String"))
            tbClientes.Columns.Add("Estado_Id", Type.GetType("System.String"))
            tbClientes.Columns.Add("Desc_Estado", Type.GetType("System.String"))
            tbClientes.Columns.Add("Pais", Type.GetType("System.String"))
            tbClientes.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientes.Columns.Add("Fantasia", Type.GetType("System.String"))
            tbClientes.Columns.Add("Endereco", Type.GetType("System.String"))
            tbClientes.Columns.Add("Numero", Type.GetType("System.String"))
            tbClientes.Columns.Add("Complemento", Type.GetType("System.String"))
            tbClientes.Columns.Add("Bairro", Type.GetType("System.String"))
            tbClientes.Columns.Add("CEP", Type.GetType("System.String"))
            tbClientes.Columns.Add("Cidade", Type.GetType("System.String"))
            tbClientes.Columns.Add("Inscricao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Telefone", Type.GetType("System.String"))
            tbClientes.Columns.Add("Fax", Type.GetType("System.String"))
            tbClientes.Columns.Add("Email", Type.GetType("System.String"))
            tbClientes.Columns.Add("Reduzido", Type.GetType("System.String"))
            tbClientes.Columns.Add("CodigoMunicipio", Type.GetType("System.String"))
            tbClientes.Columns.Add("Situacao", Type.GetType("System.String"))
            tbClientes.Columns.Add("Habilitacao", Type.GetType("System.String"))
            tbClientes.Columns.Add("EmailNFE", Type.GetType("System.String"))
            tbClientes.Columns.Add("RegistroAuxiliar", Type.GetType("System.String"))
            tbClientes.Columns.Add("FisicaJuridica", Type.GetType("System.String"))
            tbClientes.Columns.Add("RG", Type.GetType("System.String"))
            tbClientes.Columns.Add("Site", Type.GetType("System.String"))
            tbClientes.Columns.Add("Sexo", Type.GetType("System.String"))
            tbClientes.Columns.Add("NascimentoConstituicao", Type.GetType("System.String"))
            tbClientes.Columns.Add("NaturalidadeCidade", Type.GetType("System.String"))
            tbClientes.Columns.Add("NaturalidadeEstado", Type.GetType("System.String"))
            tbClientes.Columns.Add("ClienteDesde", Type.GetType("System.String"))
            ds.Tables.Add(tbClientes)

            Dim objCorrespondencia = getEndCorrespondencia(objCliente.CodigoClienteCorrespondencia, objCliente.EndClienteCorrespondencia, objCliente.Codigo, objCliente.CodigoEndereco)

            Dim drCliente As DataRow = tbClientes.NewRow()
            drCliente("Cliente_Id") = Funcoes.FormatarCpfCnpj(objCliente.Codigo)
            drCliente("Endereco_Id") = objCliente.CodigoEndereco
            drCliente("Regiao_Id") = objCliente.CodigoRegiao

            If objCliente.CodigoRegiao > 0 Then
                drCliente("Desc_Regiao") = objCliente.Regiao.Descricao
            End If

            drCliente("Categoria_Id") = objCliente.CodigoCategoria
            drCliente("Desc_Categoria") = objCliente.Categoria.Descricao
            drCliente("Estado_Id") = objCliente.CodigoEstado
            drCliente("Desc_Estado") = objCliente.Estado.Descricao
            drCliente("Pais") = objCliente.CodigoPais
            drCliente("Nome") = objCliente.Nome
            drCliente("Fantasia") = objCliente.Fantasia
            drCliente("Endereco") = objCliente.Endereco
            drCliente("Numero") = objCliente.Numero
            drCliente("Complemento") = objCliente.Complemento
            drCliente("Bairro") = objCliente.Bairro
            drCliente("CEP") = objCliente.CEP
            drCliente("Cidade") = objCliente.Cidade
            drCliente("Inscricao") = IIf(objCorrespondencia Is Nothing, objCliente.InscricaoEstadual, objCorrespondencia.InscricaoEstadual)
            drCliente("Telefone") = objCliente.Telefone
            drCliente("Fax") = objCliente.Fax
            drCliente("Email") = IIf(objCorrespondencia Is Nothing, objCliente.Email, objCorrespondencia.Email)
            drCliente("Reduzido") = objCliente.Reduzido
            drCliente("CodigoMunicipio") = objCliente.CodigoMunicipio
            drCliente("Situacao") = objCliente.CodigoSituacao
            drCliente("Habilitacao") = objCliente.Habilitacao
            drCliente("EmailNFE") = objCliente.EmailNFE
            drCliente("RegistroAuxiliar") = objCliente.OrgaoRegCategoria
            drCliente("FisicaJuridica") = IIf(objCorrespondencia Is Nothing, IIf(objCliente.Codigo.Length = "11", "F", "J"), objCorrespondencia.FisicaJuridica)
            drCliente("RG") = IIf(objCorrespondencia Is Nothing, objCliente.RG, objCorrespondencia.RG)
            drCliente("Site") = IIf(objCorrespondencia Is Nothing, objCliente.Site, objCorrespondencia.Site)
            drCliente("Sexo") = IIf(objCorrespondencia Is Nothing, objCliente.Sexo, objCorrespondencia.Sexo)
            drCliente("NascimentoConstituicao") = objCliente.NascimentoConstituicao
            drCliente("NaturalidadeCidade") = IIf(String.IsNullOrWhiteSpace(objCliente.NaturalidadeCidade), "-", objCliente.NaturalidadeCidade)
            drCliente("NaturalidadeEstado") = IIf(String.IsNullOrWhiteSpace(objCliente.NaturalidadeEstado), "-", objCliente.NaturalidadeEstado)
            drCliente("ClienteDesde") = objCliente.ClienteDesde
            tbClientes.Rows.Add(drCliente)

            'LOGOTIPO
            Dim tbLogotipo As DataTable = ds.Tables.Add("Logotipo")
            tbLogotipo.Columns.Add("path", GetType(String))
            tbLogotipo.Columns.Add("Imagem", GetType(System.Byte()))
            tbLogotipo.Columns.Add("Nome", GetType(String))
            tbLogotipo.Columns.Add("Cidade", GetType(String))
            tbLogotipo.Columns.Add("Estado_Id", GetType(String))
            Dim drImagem As DataRow = tbLogotipo.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
            drImagem("path") = strCaminhoImagem
            drImagem("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            drImagem("Nome") = HttpContext.Current.Session("ssNomeEmpresa")
            drImagem("Cidade") = HttpContext.Current.Session("ssCidadeEmpresa")
            drImagem("Estado_Id") = HttpContext.Current.Session("ssEstadoEmpresa")
            tbLogotipo.Rows.Add(drImagem)

            'ARRENDAMENTOS
            Dim tbClientesXArrendamentos As New DataTable("ClientesXArrendamentos")
            tbClientesXArrendamentos.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("EndCliente_id", Type.GetType("System.Int32"))
            tbClientesXArrendamentos.Columns.Add("Proprietario_Id", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("EndProprietario_Id", Type.GetType("System.Int32"))
            tbClientesXArrendamentos.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("DataContrato", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Vencimento", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Matricula", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Observacoes", Type.GetType("System.String"))
            tbClientesXArrendamentos.Columns.Add("Area", Type.GetType("System.Decimal"))
            ds.Tables.Add(tbClientesXArrendamentos)

            If objCliente.Arrendantes.Count > 0 Then
                For Each rowArrendante As [Lib].Negocio.ClienteXArrendante In objCliente.Arrendantes
                    Dim drArrendante As DataRow = tbClientesXArrendamentos.NewRow()
                    drArrendante("Cliente_Id") = objCliente.Codigo
                    drArrendante("EndCliente_id") = objCliente.CodigoEndereco
                    drArrendante("Proprietario_Id") = rowArrendante.CodigoArrendante
                    drArrendante("EndProprietario_Id") = rowArrendante.EndArrendante
                    drArrendante("Nome") = rowArrendante.NomeArrendante
                    drArrendante("DataContrato") = rowArrendante.DataContrato
                    drArrendante("Vencimento") = rowArrendante.DataVencimento
                    drArrendante("Matricula") = rowArrendante.Matricula
                    drArrendante("Observacoes") = rowArrendante.Observacao
                    drArrendante("Area") = rowArrendante.Area
                    tbClientesXArrendamentos.Rows.Add(drArrendante)
                Next
            End If

            'FINANCIAMENTOS
            Dim tbClientesXFinanciamentos As New DataTable("ClientesXFinanciamentos")
            tbClientesXFinanciamentos.Columns.Add("Financiador_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("EndFinanciador_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Fantasia", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Descricao", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Safra_Id", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("DataFinanciamento", Type.GetType("System.DateTime"))
            tbClientesXFinanciamentos.Columns.Add("Vencimento", Type.GetType("System.DateTime"))
            tbClientesXFinanciamentos.Columns.Add("NumeroDeParcelas", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Produto", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("Quantidade", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("Moeda", Type.GetType("System.String"))
            tbClientesXFinanciamentos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXFinanciamentos.Columns.Add("Observacoes", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXFinanciamentos)

            If objCliente.Financiamentos.Count > 0 Then
                For Each rowFinanciamento As [Lib].Negocio.ClienteXFinanciamento In objCliente.Financiamentos
                    Dim drFinanciamento As DataRow = tbClientesXFinanciamentos.NewRow()
                    drFinanciamento("Financiador_Id") = rowFinanciamento.CodigoFinanciador
                    drFinanciamento("EndFinanciador_Id") = rowFinanciamento.EndFinanciador
                    drFinanciamento("Fantasia") = rowFinanciamento.NomeFinanciador
                    drFinanciamento("Descricao") = rowFinanciamento.NomeProduto
                    drFinanciamento("Safra_Id") = rowFinanciamento.CodigoSafra
                    drFinanciamento("DataFinanciamento") = rowFinanciamento.DataFinanciamento
                    drFinanciamento("Vencimento") = rowFinanciamento.DataVencimento
                    drFinanciamento("NumeroDeParcelas") = rowFinanciamento.NumeroDeParcelas
                    drFinanciamento("Produto") = rowFinanciamento.NomeProduto
                    drFinanciamento("Quantidade") = rowFinanciamento.Quantidade
                    drFinanciamento("Moeda") = rowFinanciamento.Moeda.Descricao
                    drFinanciamento("ValorOficial") = rowFinanciamento.ValorOficial
                    drFinanciamento("ValorMoeda") = rowFinanciamento.ValorMoeda
                    drFinanciamento("Observacoes") = rowFinanciamento.Observacao
                    tbClientesXFinanciamentos.Rows.Add(drFinanciamento)
                Next
            End If

            'IMÓVEIS
            Dim tbClientesXImoveis As New DataTable("ClientesXImoveis")
            tbClientesXImoveis.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXImoveis.Columns.Add("Descricao_Id", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Onerado", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Cidade", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Estado", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("AreaTotal", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("UnidadeDeMedida", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("AreaConstruida", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("NumeroDoRegistro", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("Cartorio", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXImoveis.Columns.Add("Observacoes", Type.GetType("System.String"))
            tbClientesXImoveis.Columns.Add("DataAvaliacao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXImoveis)

            If objCliente.Imoveis.Count > 0 Then
                For Each rowImovel As [Lib].Negocio.ClienteXImovel In objCliente.Imoveis
                    Dim drImovel As DataRow = tbClientesXImoveis.NewRow()
                    drImovel("Cliente_Id") = objCliente.Codigo
                    drImovel("EndCliente_Id") = objCliente.CodigoEndereco
                    drImovel("Descricao_Id") = rowImovel.Descricao
                    drImovel("Onerado") = rowImovel.Onerado
                    drImovel("Cidade") = rowImovel.CodigoCidade
                    drImovel("Estado") = rowImovel.CodigoEstado
                    drImovel("AreaTotal") = rowImovel.AreaTotal
                    drImovel("UnidadeDeMedida") = rowImovel.CodigoUnidadeDeMedida
                    drImovel("AreaConstruida") = rowImovel.AreaConstruida
                    drImovel("NumeroDoRegistro") = rowImovel.NumeroRegistro
                    drImovel("Cartorio") = rowImovel.Cartorio
                    drImovel("ValorOficial") = rowImovel.ValorOficial
                    drImovel("ValorMoeda") = rowImovel.ValorMoeda
                    drImovel("Observacoes") = rowImovel.Observacao
                    drImovel("DataAvaliacao") = rowImovel.DataAvaliacao
                    tbClientesXImoveis.Rows.Add(drImovel)
                Next
            End If

            'SAFRAS
            Dim tbClientesXSafras As New DataTable("ClientesXSafras")
            tbClientesXSafras.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXSafras.Columns.Add("Safra_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("Cultura_Id", Type.GetType("System.String"))
            tbClientesXSafras.Columns.Add("AreaPlantada", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Produtividade", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("ConsumoProprio", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Comprometimento", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("EstimativaDeEntrega", Type.GetType("System.Decimal"))
            tbClientesXSafras.Columns.Add("Observacoes", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXSafras)

            If objCliente.Safras.Count > 0 Then
                For Each rowSafra As [Lib].Negocio.ClienteXSafra In objCliente.Safras
                    Dim drSafra As DataRow = tbClientesXSafras.NewRow()
                    drSafra("Cliente_Id") = objCliente.Codigo
                    drSafra("EndCliente_Id") = objCliente.CodigoEndereco
                    drSafra("Safra_Id") = rowSafra.CodigoSafra
                    drSafra("Cultura_Id") = rowSafra.NomeCultura
                    drSafra("AreaPlantada") = rowSafra.AreaPlantada
                    drSafra("Produtividade") = rowSafra.Produtividade
                    drSafra("ConsumoProprio") = rowSafra.ConsumoProprio
                    drSafra("Comprometimento") = rowSafra.Comprometido
                    drSafra("EstimativaDeEntrega") = rowSafra.EstimativaEntrega
                    drSafra("Observacoes") = rowSafra.Observacao
                    tbClientesXSafras.Rows.Add(drSafra)
                Next
            End If

            'VEÍCULOS
            Dim tbClientesXVeiculos As New DataTable("ClientesXVeiculos")
            tbClientesXVeiculos.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("EndCliente_Id", Type.GetType("System.Int32"))
            tbClientesXVeiculos.Columns.Add("Placa_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("TipoDeVeiculo_Id", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("Ano", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("MarcaModelo", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("Fabricante", Type.GetType("System.String"))
            tbClientesXVeiculos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tbClientesXVeiculos.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tbClientesXVeiculos.Columns.Add("DataAvaliacao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXVeiculos)

            If objCliente.Veiculos.Count > 0 Then
                For Each rowVeiculo As [Lib].Negocio.ClienteXVeiculo In objCliente.Veiculos
                    Dim drVeiculo As DataRow = tbClientesXVeiculos.NewRow()
                    drVeiculo("Cliente_Id") = objCliente.Codigo
                    drVeiculo("EndCliente_Id") = objCliente.CodigoEndereco
                    drVeiculo("Placa_Id") = rowVeiculo.CodigoPlaca
                    drVeiculo("TipoDeVeiculo_Id") = rowVeiculo.DescricaoTipoDeVeiculo
                    drVeiculo("Ano") = rowVeiculo.Ano
                    drVeiculo("MarcaModelo") = rowVeiculo.MarcaModelo
                    drVeiculo("Fabricante") = rowVeiculo.Fabricante
                    drVeiculo("ValorOficial") = rowVeiculo.ValorOficial
                    drVeiculo("ValorMoeda") = rowVeiculo.ValorMoeda
                    drVeiculo("DataAvaliacao") = rowVeiculo.DataAvaliacao
                    tbClientesXVeiculos.Rows.Add(drVeiculo)
                Next
            End If

            'EQUIPAMENTOS
            Dim tblClientesxEquipamentos As New DataTable("ClientesxEquipamentos")
            tblClientesxEquipamentos.Columns.Add("Cliente_id", Type.GetType("System.String"))
            tblClientesxEquipamentos.Columns.Add("EndCliente_id", Type.GetType("System.Int32"))
            tblClientesxEquipamentos.Columns.Add("Registro_Id", Type.GetType("System.Int32"))
            tblClientesxEquipamentos.Columns.Add("TipoDeVeiculo", Type.GetType("System.Int32"))
            tblClientesxEquipamentos.Columns.Add("DescricaoTipoVeiculo", Type.GetType("System.String"))
            tblClientesxEquipamentos.Columns.Add("Ano", Type.GetType("System.Int32"))
            tblClientesxEquipamentos.Columns.Add("Marca", Type.GetType("System.String"))
            tblClientesxEquipamentos.Columns.Add("Modelo", Type.GetType("System.String"))
            tblClientesxEquipamentos.Columns.Add("Fabricante", Type.GetType("System.String"))
            tblClientesxEquipamentos.Columns.Add("ValorOficial", Type.GetType("System.Decimal"))
            tblClientesxEquipamentos.Columns.Add("ValorMoeda", Type.GetType("System.Decimal"))
            tblClientesxEquipamentos.Columns.Add("DataAvaliacao", Type.GetType("System.DateTime"))
            tblClientesxEquipamentos.Columns.Add("Onerado", Type.GetType("System.Boolean"))
            ds.Tables.Add(tblClientesxEquipamentos)

            If objCliente.Equipamentos.Count > 0 Then
                For Each rowEquipamento As [Lib].Negocio.ClienteXEquipamento In objCliente.Equipamentos
                    Dim drEquipamento As DataRow = tblClientesxEquipamentos.NewRow()
                    drEquipamento("Cliente_Id") = objCliente.Codigo
                    drEquipamento("EndCliente_Id") = objCliente.CodigoEndereco
                    drEquipamento("Registro_Id") = rowEquipamento.Codigo
                    drEquipamento("TipoDeVeiculo") = rowEquipamento.CodigoTipoDeEquipamento
                    drEquipamento("DescricaoTipoVeiculo") = rowEquipamento.DescricaoTipoDeEquipamento
                    drEquipamento("Ano") = rowEquipamento.Ano
                    drEquipamento("Marca") = rowEquipamento.Marca
                    drEquipamento("Modelo") = rowEquipamento.Modelo
                    drEquipamento("Fabricante") = rowEquipamento.Fabricante
                    drEquipamento("ValorOficial") = rowEquipamento.ValorOficial
                    drEquipamento("ValorMoeda") = rowEquipamento.ValorMoeda
                    drEquipamento("DataAvaliacao") = rowEquipamento.DataAvaliacao
                    drEquipamento("Onerado") = rowEquipamento.Onerado
                    tblClientesxEquipamentos.Rows.Add(drEquipamento)
                Next
            End If

            'REPRESENTANTES
            Dim tbClientesXRepresentantes As New DataTable("ClientesXRepresentantes")
            tbClientesXRepresentantes.Columns.Add("Cliente_Id", Type.GetType("System.String"))
            tbClientesXRepresentantes.Columns.Add("Endereco_Id", Type.GetType("System.String"))
            tbClientesXRepresentantes.Columns.Add("Nome", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXRepresentantes)

            If objCliente.Representantes.Count > 0 Then
                For Each rowRepresentante As [Lib].Negocio.ClienteXRepresentante In objCliente.Representantes
                    Dim drRepresentante As DataRow = tbClientesXRepresentantes.NewRow()
                    drRepresentante("Cliente_Id") = rowRepresentante.CodigoRepresentante
                    drRepresentante("Endereco_Id") = rowRepresentante.EndRepresentante
                    drRepresentante("Nome") = rowRepresentante.NomeRepresentante
                    tbClientesXRepresentantes.Rows.Add(drRepresentante)
                Next
            End If

            'DEPENDENTES
            Dim tbClientesXDependentes As New DataTable("ClientesXDependentes")
            tbClientesXDependentes.Columns.Add("Nome", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("TipoDeDependente", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("Identidade", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("CPF", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("DataNascimento", Type.GetType("System.String"))
            tbClientesXDependentes.Columns.Add("Profissao", Type.GetType("System.String"))
            ds.Tables.Add(tbClientesXDependentes)

            If objCliente.Dependentes.Count > 0 Then
                For Each rowDependente As [Lib].Negocio.ClientexDependente In objCliente.Dependentes
                    Dim drDependente As DataRow = tbClientesXDependentes.NewRow()
                    drDependente("Nome") = rowDependente.Nome
                    drDependente("TipoDeDependente") = rowDependente.TipoDeDependente
                    drDependente("Identidade") = rowDependente.RG
                    drDependente("CPF") = rowDependente.CPF
                    drDependente("DataNascimento") = rowDependente.DataNascimento
                    drDependente("Profissao") = rowDependente.Profissao
                    tbClientesXDependentes.Rows.Add(drDependente)
                Next
            End If

            'CONTAS BANCARIAS
            Dim tblContasBancarias As New DataTable("ClientesXContasBancarias")
            tblContasBancarias.Columns.Add("Banco_Id", Type.GetType("System.Int32"))
            tblContasBancarias.Columns.Add("NomeBanco", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("Agencia_Id", Type.GetType("System.Int32"))
            tblContasBancarias.Columns.Add("DigitoAgencia_Id", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("ContaCorrente_Id", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("DigitoConta_Id", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("TipoConta", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("Observacoes", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("Praca", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("Cidade", Type.GetType("System.String"))
            tblContasBancarias.Columns.Add("Estado", Type.GetType("System.String"))

            ds.Tables.Add(tblContasBancarias)

            If objCliente.ContasBancarias.Count > 0 Then
                For Each rowContaBancaria As [Lib].Negocio.ClienteXContaBancaria In objCliente.ContasBancarias
                    Dim drContasBancarias As DataRow = tblContasBancarias.NewRow()
                    drContasBancarias("Banco_Id") = rowContaBancaria.CodigoBanco
                    drContasBancarias("NomeBanco") = rowContaBancaria.NomeBanco
                    drContasBancarias("Agencia_Id") = rowContaBancaria.CodigoAgencia
                    drContasBancarias("DigitoAgencia_Id") = rowContaBancaria.DigitoAgencia
                    drContasBancarias("ContaCorrente_Id") = rowContaBancaria.ContaCorrente
                    drContasBancarias("DigitoConta_Id") = rowContaBancaria.DigitoConta
                    drContasBancarias("TipoConta") = rowContaBancaria.TipoConta
                    drContasBancarias("Observacoes") = rowContaBancaria.Observacoes
                    drContasBancarias("Praca") = rowContaBancaria.Praca
                    drContasBancarias("Cidade") = rowContaBancaria.Cidade
                    drContasBancarias("Estado") = rowContaBancaria.Estado
                    tblContasBancarias.Rows.Add(drContasBancarias)
                Next
            End If

            'CONTATOS
            Dim tbContatos As New DataTable("Contatos")
            tbContatos.Columns.Add("NomeContato", Type.GetType("System.String"))
            tbContatos.Columns.Add("Funcao", Type.GetType("System.String"))
            tbContatos.Columns.Add("Telefone", Type.GetType("System.String"))
            tbContatos.Columns.Add("Email", Type.GetType("System.String"))
            ds.Tables.Add(tbContatos)

            If objCliente.Contatos.Count > 0 Then
                For Each rowContato As [Lib].Negocio.ClienteXContato In objCliente.Contatos
                    Dim drContato As DataRow = tbContatos.NewRow()
                    drContato("NomeContato") = rowContato.NomeContato
                    drContato("Funcao") = rowContato.Funcao
                    drContato("Telefone") = rowContato.Telefone
                    drContato("Email") = rowContato.email
                    tbContatos.Rows.Add(drContato)
                Next
            End If

            'TIPOS DE CLIENTE
            Dim tbTiposDeClientes As New DataTable("TiposDeClientes")
            tbTiposDeClientes.Columns.Add("Tipo_Id", Type.GetType("System.String"))
            tbTiposDeClientes.Columns.Add("Descricao", Type.GetType("System.String"))
            ds.Tables.Add(tbTiposDeClientes)

            If objCliente.Tipos.Count > 0 Then
                For Each rowTipo As [Lib].Negocio.ClientexTipo In objCliente.Tipos
                    Dim drTipo As DataRow = tbTiposDeClientes.NewRow()
                    drTipo("Tipo_Id") = rowTipo.CodigoTipo
                    drTipo("Descricao") = rowTipo.DescricaoTipo
                    tbTiposDeClientes.Rows.Add(drTipo)
                Next
            End If

            Dim rpt As New ReportDocument()
            rpt.FileName = Server.MapPath("~/Reports/Cr_FichaCadastralCompleta.rpt")
            rpt.Load(rpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
            Dim arquivo As String = NomeArquivo

            Try
                rpt.SetDataSource(ds)

                Dim param As New Dictionary(Of String, Object)
                param.Add("TemDependentes", ds.Tables("ClientesXDependentes").Rows.Count > 0)
                param.Add("TemRepresentantes", ds.Tables("ClientesXRepresentantes").Rows.Count > 0)
                param.Add("TemContatos", ds.Tables("Contatos").Rows.Count > 0)
                param.Add("TemContasBancarias", ds.Tables("ClientesXContasBancarias").Rows.Count > 0)
                param.Add("TemSafras", ds.Tables("ClientesXSafras").Rows.Count > 0)
                param.Add("TemFinanciamentos", ds.Tables("ClientesXFinanciamentos").Rows.Count > 0)
                param.Add("TemArrendamentos", ds.Tables("ClientesXArrendamentos").Rows.Count > 0)
                param.Add("TemTipoCliente", ds.Tables("TiposDeClientes").Rows.Count > 0)
                param.Add("TemImoveis", ds.Tables("ClientesXImoveis").Rows.Count > 0)
                param.Add("TemVeiculos", ds.Tables("ClientesXVeiculos").Rows.Count > 0)
                param.Add("TemEquipamentos", ds.Tables("ClientesXEquipamentos").Rows.Count > 0)
                param.Add("EndCorrespondencia", IIf(objCorrespondencia Is Nothing, String.Empty, objCorrespondencia.EndCorrespondencia))
                Funcoes.BindParameters(rpt, param)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                Try
                    rpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message.ToString)
                Finally
                    rpt.Close()
                    rpt.Dispose()
                    If IO.File.Exists(arquivo) Then
                        Funcoes.AbrirArquivo(Page, NomeArquivo2)
                    End If
                End Try
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message.ToString)
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtCpfCnpj_TextChanged(sender As Object, e As EventArgs) Handles txtCpfCnpj.TextChanged
        Try
            lnkConsultar_Click(Nothing, Nothing)
            'PESQUISA NA RECEITA FEDERAL DADOS DA EMPRESA
            If (txtCpfCnpj.Text.RemoveMask.Length = 14) AndAlso gridConsultaCliente.Rows.Count = 0 Then
                Dim novoCliente As New Cliente(txtCpfCnpj.Text.RemoveMask)

                If novoCliente.Nome.Length = 0 Then
                    MsgBox(Me.Page, "A Receita Federal não disponibilizou as informações desse CNPJ.")
                    Exit Sub
                End If


                LimparCampos()

                RecuperarSessionCliente()

                objCliente = novoCliente

                objCliente.IUD = "I"

                txtReduzido.Text = objCliente.Reduzido
                txtCpfCnpj.Text = objCliente.Codigo
                txtCodigoEndereco.Text = objCliente.CodigoEndereco

                txtNome.Text = objCliente.Nome
                txtNomeFantasia.Text = objCliente.Fantasia

                txtCep.Text = objCliente.CEP

                txtEndereco.Text = objCliente.Endereco
                txtNumero.Text = objCliente.Numero
                txtComplemento.Text = objCliente.Complemento
                txtBairro.Text = objCliente.Bairro
                ddlEstado.SelectedValue = objCliente.CodigoEstado
                txtCodMunicipio.Text = objCliente.CodigoMunicipio

                txtCorrespondencia.Text = String.Empty

                txtCidade.Text = objCliente.Cidade
                txtCidade.Enabled = False

                ddlPais.SelectedValue = objCliente.CodigoPais
                ddlPais_SelectedIndexChanged(Nothing, Nothing)

                txtDataNascimento.Text = objCliente.NascimentoConstituicao.ToShortDateString()

                objCliente.ClienteDesde = String.Format("{0}/{1}/{2}", Now.Day.ToString.PadLeft(2, "0"), Now.Month.ToString.PadLeft(2, "0"), Now.Year)
                txtDataCliente.Text = objCliente.ClienteDesde.ToShortDateString()

                txtTelefone.Text = objCliente.Telefone
                txtEmail.Text = objCliente.Email
                txtEmailNFE.Text = objCliente.EmailNFE
                ddlSituacao.SelectedValue = objCliente.CodigoSituacao

                If objCliente.CodigoSituacao = 1 Then
                    ddlSituacao.Enabled = True
                    imgBloqueio.Visible = False
                Else
                    ddlSituacao.Enabled = False
                    imgBloqueio.Visible = True
                End If

                ddlCategoria.SelectedValue = 4
                objCliente.CodigoCategoria = ddlCategoria.SelectedValue

                Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)
                objTipo.IUD = "I"
                objTipo.CodigoTipo = 4
                objCliente.Tipos.Add(objTipo)
                'chkListTipoDeCliente.Items(1).Selected = True

                objTipo = New [Lib].Negocio.ClientexTipo(objCliente)
                objTipo.IUD = "I"
                objTipo.CodigoTipo = 5
                objCliente.Tipos.Add(objTipo)
                'chkListTipoDeCliente.Items(5).Selected = True

                For Each row As [Lib].Negocio.ClientexTipo In objCliente.Tipos
                    For i As Integer = 0 To chkListTipoDeCliente.Items.Count - 1
                        If chkListTipoDeCliente.Items(i).Value = row.CodigoTipo Then
                            chkListTipoDeCliente.Items(i).Selected = True
                            Exit For
                        End If
                    Next
                Next

                objCliente.UsuarioInclusao = Session("ssNomeUsuario")
                objCliente.UsuarioInclusaoData = Now()
                ddlUsuarios.Items.Clear()
                ddlUsuarios.Items.Add("Inc.- " & objCliente.UsuarioInclusao)

                SalvarSessionCliente()

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgBloqueio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgBloqueio.Click
        Try
            If Funcoes.VerificaPermissao("ALTERARSITUACAOCLIENTE", "ALTERAR") Then
                ddlSituacao.Enabled = True
            Else
                MsgBox(Me.Page, "Usuário Sem permissão para alterar a situação do cliente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnCorrespondencia_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCliXCorr" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgLimparEndCorrespondencia_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            RecuperarSessionCliente()
            objCliente.CodigoClienteCorrespondencia = String.Empty
            objCliente.EndClienteCorrespondencia = 0
            txtCorrespondencia.Text = String.Empty
            SalvarSessionCliente()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Function consultaMunicipio(estado As String, codigo As String, municipio As String)
        Dim sql As String = String.Empty
        Dim text = Textos.RetirarAcentos(municipio)

        '" AND Municipio_id LIKE '" & Mid(text, 1, 5) & "%'"
        sql = " SELECT Municipio_id FROM Municipios " &
              " WHERE Estado_id = '" & estado & "'" &
              " AND Codigo_id = " & codigo &
              " AND Municipio_id = '" & text.ToUpper & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Municipios")
        Return ds.Tables(0).Rows(0).Item("Municipio_id")
    End Function

    Protected Sub imgCep_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgCep.Click
        Try
            If txtCep.Text.RemoveMask.Length <> 8 Then
                MsgBox(Me.Page, "CEP deve ser informado com 8 dígitos.")
            Else
                If Funcoes.VerificaConexaoInternet Then
                    Dim ds As New DataSet
                    ds = Funcoes.BuscaCep(txtCep.Text.RemoveMask)

                    If Not IsNothing(ds) AndAlso (ds.Tables(0).Rows.Count > 0) Then
                        If ds.Tables(0).Columns.Contains("erro") Then
                            MsgBox(Me.Page, "Cep informado é inexistente.")
                            txtCep.Text = String.Empty
                            txtCep.Focus()
                        Else
                            'retorno da consulta de cep do endereço https://viacep.com.br/ 
                            ddlPais.SelectedValue = "1058"
                            ddlPais_SelectedIndexChanged(Nothing, Nothing)

                            Dim row As DataRow = ds.Tables(0).Rows(0)
                            Dim estado As String = row("uf").ToString().ToUpper()
                            Dim codigo As String = row("ibge").ToString().Substring(2, row("ibge").ToString.Length - 2)
                            Dim municipio As String = Funcoes.SubstituirCaracteresEspeciaisMunicipio(row("localidade").ToString().ToUpper())

                            Dim cidade As String = consultaMunicipio(estado, codigo, municipio)

                            If String.IsNullOrWhiteSpace(cidade) OrElse Not cidade = municipio Then
                                ddlEstado.SelectedIndex = 0
                                txtCidade.Text = String.Empty
                                txtCodMunicipio.Text = String.Empty
                                txtBairro.Text = row("bairro").ToString().ToUpper()
                                txtEndereco.Text = row("logradouro").ToString().ToUpper()
                                MsgBox(Me.Page, "Cidade não foi encontrada no Cadastro de Municípios, entre em contato com o Suporte.")
                            Else
                                ddlEstado.SelectedValue = estado
                                txtCidade.Text = cidade
                                txtBairro.Text = row("bairro").ToString().ToUpper()
                                txtEndereco.Text = row("logradouro").ToString().ToUpper()
                                txtCodMunicipio.Text = codigo
                            End If

                            If String.IsNullOrWhiteSpace(txtEndereco.Text) Then
                                txtEndereco.Focus()
                            Else
                                txtNumero.Focus()
                            End If

                            If Not String.IsNullOrWhiteSpace(row("logradouro")) Then txtEndereco.Focus()
                        End If
                    Else
                        MsgBox(Me.Page, "O serviço de 'BuscaCep' esta fora do ar.")
                    End If
                Else
                    MsgBox(Me.Page, "No momento não é possível executar esse processo. Verifique sua conexão com a Internet.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultaCep_Click(sender As Object, e As ImageClickEventArgs) Handles imgConsultaCep.Click
        Try
            ucConsultaCep.Limpar()
            Popup.ConsultaDeCep(Me.Page, "objCep" & HID.Value.ToString, "txtCep")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub consultarClienteEX()
        Dim Sql As String

        Sql = "Select * from Clientes" & vbCrLf &
              "Where Estado = 'EX'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        Dim objCnpj As String = ""

        For Each row As DataRow In ds.Tables(0).Rows
            objCnpj = row("Cliente_Id")
        Next
        Session("cnpj") = CStr(objCnpj + 1)
    End Sub

    Protected Sub ddlPais_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlPais.SelectedIndexChanged
        Try
            If String.IsNullOrWhiteSpace(ddlPais.SelectedValue) Then
                ddlEstado.SelectedIndex = 0
                txtCidade.Text = String.Empty
                txtCodMunicipio.Text = String.Empty
                txtCep.Text = String.Empty
            ElseIf ddlPais.SelectedValue <> "1058" Then
                ddlEstado.Enabled = False
                ddlEstado.SelectedValue = "EX"
                txtCodMunicipio.Text = "9999999"
                txtCep.Text = "99999999"
                txtCidade.Enabled = True
                ddlRegiao.SelectedValue = 11

                If txtCpfCnpj.Text.Length = 0 Then
                    consultarClienteEX()
                    txtCpfCnpj.Text = Session("cnpj")
                    txtCodigoEndereco.Text = "0"
                End If

                txtCpfCnpj.Enabled = False
            Else
                ddlEstado.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlEstado_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEstado.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlEstado.SelectedValue) Then
                If ddlPais.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "País deve ser selecionado.")
                ElseIf ddlEstado.SelectedValue = "EX" Then
                    ddlEstado.SelectedIndex = 0
                    MsgBox(Me.Page, "EXTERIOR não pode ser selecionado.")
                Else
                    txtCidade.Text = String.Empty
                    txtCodMunicipio.Text = 0
                    txtCep.Text = String.Empty
                    Session("ssUF" & HID.Value) = ddlEstado.SelectedValue
                    ucConsultaCodMunicipios.Limpar()
                    Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioCli" & HID.Value)
                End If
            Else
                txtCidade.Text = String.Empty
                txtCodMunicipio.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlRegiao_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlRegiao.SelectedIndexChanged
        Try
            If ddlRegiao.SelectedValue > 0 Then
                ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, " Regiao_Id= " & ddlRegiao.SelectedValue)
            Else
                ddlMicroRegiao.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlNaturalidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(ddlNaturalidade.SelectedValue) Then
                Session("ssUF" & HID.Value) = ddlNaturalidade.SelectedValue
                ucConsultaCodMunicipios.Limpar()
                Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioNATCLI" & HID.Value)
            Else
                txtCidadeNatural.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtInscricao_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ValidarInscricao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub chkIsento_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkIsento.CheckedChanged
        Try
            If chkIsento.Checked Then
                txtInscricao.Text = "ISENTO"
                txtCodigoInscricao.Value = String.Empty
                imgInscricao.ImageUrl = "~/images/certo.jpg"
                imgInscricao.ToolTip = "Inscrição Válida"
            Else
                txtInscricao.Text = String.Empty
                txtCodigoInscricao.Value = String.Empty
            End If
            txtInscricao.Enabled = Not chkIsento.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnConsultaCadastro_Click(sender As Object, e As EventArgs) Handles btnConsultaCadastro.Click
        Try
            Dim txtCpf As TextBox = CType(ucConsultaCadastro.FindControlRecursive("txtCpfCnpj"), TextBox)
            ucConsultaCadastro.Limpar()
            ucConsultaCadastro.Origem = eTipoDeDocumento.Nota
            Popup.ConsultaCadastro(Me.Page, "objCadastro" & HID.Value, txtCpf.ClientID, True, 100)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgBtAlteraCidade_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgBtAlteraCidade.Click
        Try
            ddlEstado_SelectedIndexChanged(Nothing, Nothing)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'TIPO DE CLIENTES *********************************************************************************************************
    '**************************************************************************************************************************
    Protected Sub chkListTipoDeCliente_DataBound(sender As Object, e As EventArgs) Handles chkListTipoDeCliente.DataBound
        Try
            For i As Integer = 0 To CType(sender, CheckBoxList).Items.Count - 1
                If CType(sender, CheckBoxList).Items(i).Text = "EMPRESA" OrElse
                    CType(sender, CheckBoxList).Items(i).Text = "UNIDADE DE NEGOCIO" OrElse
                    CType(sender, CheckBoxList).Items(i).Text = "CENTRO DE DISTRIBUICAO" OrElse
                    CType(sender, CheckBoxList).Items(i).Text = "SERVIDOR" Then
                    CType(sender, CheckBoxList).Items(i).Enabled = False
                End If
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub chkListTipoDeCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            RecuperarSessionCliente()

            Dim result As String = Request.Form("__EVENTTARGET")
            Dim checkedBox As String() = result.Split("$"c)

            Dim index As Integer = Integer.Parse(checkedBox(checkedBox.Length - 1))


            Dim codigo As String = chkListTipoDeCliente.Items(index).Value
            Dim descricao As String = chkListTipoDeCliente.Items(index).Text

            If Not String.IsNullOrWhiteSpace(txtCpfCnpj.Text) Then
                Dim objTipo As New [Lib].Negocio.ClientexTipo(objCliente)
                objTipo.IUD = "I"
                objTipo.CodigoTipo = codigo

                If objCliente.IUD = "U" AndAlso Not chkListTipoDeCliente.Items(index).Selected Then
                    objTipo.IUD = "D"
                    If objTipo.Salvar() Then objCliente.Tipos.Remove(objTipo)
                Else
                    If objCliente.IUD = "U" Then objTipo.Salvar()

                    If objCliente.Tipos.Where(Function(x) x.CodigoTipo = objTipo.CodigoTipo).Count() = 0 Then
                        objCliente.Tipos.Add(objTipo)
                    End If
                End If
                SalvarSessionCliente()
            Else
                MsgBox(Me.Page, "Informe o cnpj.")
                If chkListTipoDeCliente.Items(index).Selected Then chkListTipoDeCliente.Items(index).Selected = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'CONTATO*****************************************************************************************************************
    '************************************************************************************************************************
    Protected Sub imgConsultarContato_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgContato As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgContato.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexContato.Value = row.RowIndex

            txtNomeContato.Text = objCliente.Contatos(row.RowIndex).NomeContato
            txtFuncaoContato.Text = objCliente.Contatos(row.RowIndex).Funcao
            txtTelefoneContato.Text = objCliente.Contatos(row.RowIndex).Telefone
            txtEmailContato.Text = objCliente.Contatos(row.RowIndex).email
            ddlBancoContato.SelectedIndex = objCliente.Contatos(row.RowIndex).CodigoBanco
            txtAgenciaContato.Text = objCliente.Contatos(row.RowIndex).CodigoAgencia
            txtDGAgContato.Text = objCliente.Contatos(row.RowIndex).DigitoAgencia
            txtCtaContato.Text = objCliente.Contatos(row.RowIndex).CodigoContaCorrente
            txtDGCtaContato.Text = objCliente.Contatos(row.RowIndex).DigitoConta
            txtObservacaoContato.Text = objCliente.Contatos(row.RowIndex).Observacao

            txtNomeContato.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoCONT_Click(sender As Object, e As EventArgs) Handles lnkNovoCONT.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtNomeContato.Text.Length = 0 Then
                    MsgBox(Me.Page, "Nome do Contato não foi informado.")
                ElseIf ddlBancoContato.SelectedIndex > 0 AndAlso (txtAgenciaContato.Text.Length = 0 OrElse txtDGAgContato.Text.Length = 0 OrElse txtCtaContato.Text.Length = 0 OrElse txtDGCtaContato.Text.Length = 0) Then
                    MsgBox(Me.Page, "Caso o Banco seja selecionado, Agência-DG e Conta-DG são obrigatórios.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text.Trim
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXContato = New [Lib].Negocio.ClienteXContato()
                    objClienteXContato.Cliente = objCliente
                    objClienteXContato.NomeContato = UCase(Funcoes.EliminarCaracteresEspeciais(txtNomeContato.Text.Trim))
                    objClienteXContato.Funcao = UCase(Funcoes.EliminarCaracteresEspeciais(txtFuncaoContato.Text.Trim))
                    objClienteXContato.Telefone = txtTelefoneContato.Text.RemoveMask
                    objClienteXContato.email = txtEmailContato.Text.Trim
                    objClienteXContato.CodigoBanco = IIf(ddlBancoContato.SelectedIndex = 0, 0, ddlBancoContato.SelectedValue)
                    objClienteXContato.NomeBanco = IIf(ddlBancoContato.SelectedIndex = 0, "", ddlBancoContato.SelectedItem.Text())
                    objClienteXContato.CodigoAgencia = IIf(txtAgenciaContato.Text.Length = 0, 0, txtAgenciaContato.Text.Trim)
                    objClienteXContato.DigitoAgencia = txtDGAgContato.Text.Trim
                    objClienteXContato.CodigoContaCorrente = txtCtaContato.Text.Trim
                    objClienteXContato.DigitoConta = txtDGCtaContato.Text.Trim
                    objClienteXContato.Observacao = txtObservacaoContato.Text.ToUpper.Trim

                    If objCliente.IUD = "I" Then
                        objCliente.Contatos.Add(objClienteXContato)
                    Else
                        objClienteXContato.IUD = "U"
                        If txtNomeContato.Enabled Then objClienteXContato.IUD = "I"
                        If objClienteXContato.Salvar() Then
                            If objClienteXContato.IUD = "I" Then
                                objCliente.Contatos.Add(objClienteXContato)
                            Else
                                objCliente.Contatos(HgridRowIndexContato.Value).NomeContato = UCase(Funcoes.EliminarCaracteresEspeciais(txtNomeContato.Text.Trim))
                                objCliente.Contatos(HgridRowIndexContato.Value).Funcao = UCase(Funcoes.EliminarCaracteresEspeciais(txtFuncaoContato.Text.Trim))
                                objCliente.Contatos(HgridRowIndexContato.Value).Telefone = txtTelefoneContato.Text.RemoveMask
                                objCliente.Contatos(HgridRowIndexContato.Value).email = txtEmailContato.Text.Trim
                                objCliente.Contatos(HgridRowIndexContato.Value).CodigoBanco = IIf(ddlBancoContato.SelectedIndex = 0, 0, ddlBancoContato.SelectedValue)
                                objCliente.Contatos(HgridRowIndexContato.Value).NomeBanco = IIf(ddlBancoContato.SelectedIndex = 0, "", ddlBancoContato.SelectedItem.Text())
                                objCliente.Contatos(HgridRowIndexContato.Value).CodigoAgencia = IIf(txtAgenciaContato.Text.Trim.Length = 0, 0, txtAgenciaContato.Text.Trim)
                                objCliente.Contatos(HgridRowIndexContato.Value).DigitoAgencia = txtDGAgContato.Text.Trim
                                objCliente.Contatos(HgridRowIndexContato.Value).CodigoContaCorrente = txtCtaContato.Text.Trim
                                objCliente.Contatos(HgridRowIndexContato.Value).DigitoConta = txtDGCtaContato.Text.Trim
                                objCliente.Contatos(HgridRowIndexContato.Value).Observacao = txtObservacaoContato.Text.ToUpper.Trim
                            End If
                        End If
                    End If

                    SalvarSessionCliente()
                    Limpar_Contatos()
                    Carregar_Contatos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirContato_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                RecuperarSessionCliente()
                Dim imgContato As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgContato.NamingContainer, GridViewRow)

                objClienteXContato = New [Lib].Negocio.ClienteXContato(ddlEmpresas.SelectedValue.Split("-")(0), ddlEmpresas.SelectedValue.Split("-")(1), row.Cells(1).Text)
                objClienteXContato.IUD = "D"

                If Not objCliente.IUD = "I" Then objClienteXContato.Salvar()
                objCliente.Contatos.RemoveAt(row.RowIndex)

                SalvarSessionCliente()
                Limpar_Contatos()
                Carregar_Contatos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparCONT_Click(sender As Object, e As EventArgs) Handles lnkLimparCONT.Click
        Try
            Limpar_Contatos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'DEPENDENTES***************************************************************************************************************
    '**************************************************************************************************************************
    Protected Sub imgConsultarDependente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgDependente As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgDependente.NamingContainer, GridViewRow)

            RecuperarSessionCliente()
            HgridRowIndexDependente.Value = row.RowIndex

            txtNomeDependente.Text = objCliente.Dependentes(row.RowIndex).Nome
            ddlTipoDependente.SelectedValue = objCliente.Dependentes(row.RowIndex).TipoDeDependente
            txtRGDependente.Text = objCliente.Dependentes(row.RowIndex).RG
            txtCPFDependente.Text = objCliente.Dependentes(row.RowIndex).CPF
            txtNascimentoDependente.Text = objCliente.Dependentes(row.RowIndex).DataNascimento
            txtProfissaoDependente.Text = objCliente.Dependentes(row.RowIndex).Profissao
            txtCustoDependente.Text = objCliente.Dependentes(row.RowIndex).CustoAno

            txtNomeDependente.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirDependente_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgDependente As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgDependente.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXD As [Lib].Negocio.ClientexDependente In objCliente.Dependentes
                    If rowCXD.Nome = row.Cells(1).Text Then

                        objClienteXDependente = New [Lib].Negocio.ClientexDependente()
                        objClienteXDependente.Cliente = objCliente
                        objClienteXDependente.Nome = rowCXD.Nome
                        objClienteXDependente.TipoDeDependente = rowCXD.TipoDeDependente
                        objClienteXDependente.RG = rowCXD.RG
                        objClienteXDependente.CPF = rowCXD.CPF
                        objClienteXDependente.DataNascimento = rowCXD.DataNascimento
                        objClienteXDependente.Profissao = rowCXD.Profissao
                        objClienteXDependente.CustoAno = rowCXD.CustoAno

                        objClienteXDependente.IUD = "D"
                        objCliente.Dependentes.RemoveAt(row.RowIndex)

                        If Not objCliente.IUD = "I" Then objClienteXDependente.Salvar()
                    End If
                Next
                SalvarSessionCliente()
                Limpar_Dependentes()
                Carregar_Dependentes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoD_Click(sender As Object, e As EventArgs) Handles lnkNovoD.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtNomeDependente.Text.Length = 0 Then
                    MsgBox(Me.Page, "Nome do Dependente não foi informado.")
                ElseIf ddlTipoDependente.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Tipo do Dependente não foi selecionado.")
                ElseIf txtRGDependente.Text.Length = 0 Then
                    MsgBox(Me.Page, "RG do Dependente não foi informado.")
                ElseIf txtCPFDependente.Text.Length > 0 AndAlso Not Funcoes.ValidaCPF(txtCPFDependente.Text.RemoveMask) Then
                    MsgBox(Me.Page, "CPF do Dependente informado não é válido.")
                ElseIf txtNascimentoDependente.Text.Length = 0 Then
                    MsgBox(Me.Page, "Data de Nascimento do Dependente não foi informada.")
                Else
                    RecuperarSessionCliente()

                    If Not IsNumeric(txtCustoDependente.Text) Then txtCustoDependente.Text = 0

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXDependente = New [Lib].Negocio.ClientexDependente()
                    objClienteXDependente.Cliente = objCliente
                    objClienteXDependente.Nome = UCase(Funcoes.EliminarCaracteresEspeciais(txtNomeDependente.Text.Trim))
                    objClienteXDependente.TipoDeDependente = ddlTipoDependente.SelectedValue
                    objClienteXDependente.RG = txtRGDependente.Text.RemoveMask
                    objClienteXDependente.CPF = txtCPFDependente.Text.RemoveMask
                    objClienteXDependente.DataNascimento = txtNascimentoDependente.Text
                    objClienteXDependente.Profissao = UCase(txtProfissaoDependente.Text.Trim)
                    objClienteXDependente.CustoAno = IIf(IsNumeric(txtCustoDependente.Text), CDec(txtCustoDependente.Text), 0)

                    If objCliente.IUD = "I" Then
                        objCliente.Dependentes.Add(objClienteXDependente)
                    Else
                        objClienteXDependente.IUD = "U"
                        If txtNomeDependente.Enabled Then objClienteXDependente.IUD = "I"

                        If objClienteXDependente.Salvar() Then
                            If objClienteXDependente.IUD = "I" Then
                                objCliente.Dependentes.Add(objClienteXDependente)
                            Else
                                objCliente.Dependentes(HgridRowIndexDependente.Value).Nome = UCase(Funcoes.EliminarCaracteresEspeciais(txtNomeDependente.Text))
                                objCliente.Dependentes(HgridRowIndexDependente.Value).TipoDeDependente = ddlTipoDependente.SelectedValue
                                objCliente.Dependentes(HgridRowIndexDependente.Value).RG = txtRGDependente.Text.RemoveMask
                                objCliente.Dependentes(HgridRowIndexDependente.Value).CPF = txtCPFDependente.Text.RemoveMask
                                objCliente.Dependentes(HgridRowIndexDependente.Value).DataNascimento = txtNascimentoDependente.Text
                                objCliente.Dependentes(HgridRowIndexDependente.Value).Profissao = UCase(txtProfissaoDependente.Text.Trim)
                                objCliente.Dependentes(HgridRowIndexDependente.Value).CustoAno = IIf(IsNumeric(txtCustoDependente.Text), CDec(txtCustoDependente.Text), 0)
                            End If
                        End If
                    End If

                    SalvarSessionCliente()
                    Limpar_Dependentes()
                    Carregar_Dependentes()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparD_Click(sender As Object, e As EventArgs) Handles lnkLimparD.Click
        Try
            Limpar_Dependentes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'CONTAS BANCARIAS**********************************************************************************************************
    '**************************************************************************************************************************
    Protected Sub ddlEstadoContaBancaria_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(ddlEstadoContaBancaria.SelectedValue) Then
                Session("ssUF" & HID.Value) = ddlEstadoContaBancaria.SelectedValue
                ucConsultaCodMunicipios.Limpar()
                Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioCLIxCTA" & HID.Value)
            Else
                txtCidadeContaBancaria.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarContaBancaria_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgContaBancaria As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgContaBancaria.NamingContainer, GridViewRow)

            RecuperarSessionCliente()
            HgridRowIndexContaBancaria.Value = row.RowIndex

            ddlBancoContaBancaria.SelectedValue = objCliente.ContasBancarias(row.RowIndex).CodigoBanco
            txtAgenciaContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).CodigoAgencia
            txtDGAgContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).DigitoAgencia
            txtCtaContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).ContaCorrente
            txtDGCtaContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).DigitoConta
            ddlTipoConta.SelectedValue = objCliente.ContasBancarias(row.RowIndex).TipoConta
            ddlEstadoContaBancaria.SelectedValue = objCliente.ContasBancarias(row.RowIndex).Estado
            txtCidadeContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).Cidade
            txtPracaContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).Praca
            txtObservacaoContaBancaria.Text = objCliente.ContasBancarias(row.RowIndex).Observacoes
            chkAtivo.Checked = objCliente.ContasBancarias(row.RowIndex).Ativo

            ddlBancoContaBancaria.Enabled = False
            txtAgenciaContaBancaria.Enabled = False
            txtDGAgContaBancaria.Enabled = False
            txtCtaContaBancaria.Enabled = False
            txtDGCtaContaBancaria.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirContaBancaria_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then

                Dim imgContaBancaria As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgContaBancaria.NamingContainer, GridViewRow)

                If ContemVinculoTitulo(row) Then
                    MsgBox(Me.Page, "Conta não pode ser excluída, pois está associado a Título(s).")
                    Exit Sub
                End If

                RecuperarSessionCliente()

                For Each objCXCB As [Lib].Negocio.ClienteXContaBancaria In objCliente.ContasBancarias
                    Dim dAgencia As String = row.Cells(4).Text
                    dAgencia = dAgencia.Replace("&nbsp;", "")

                    Dim dConta As String = row.Cells(6).Text
                    dConta = dConta.Replace("&nbsp;", "")

                    If objCXCB.CodigoBanco = Server.HtmlDecode(row.Cells(1).Text) AndAlso
                       objCXCB.CodigoAgencia = Server.HtmlDecode(row.Cells(3).Text) AndAlso
                       objCXCB.DigitoAgencia = dAgencia AndAlso
                       objCXCB.ContaCorrente = Server.HtmlDecode(row.Cells(5).Text) AndAlso
                       objCXCB.DigitoConta = dConta Then

                        objClienteXContaBancaria = New [Lib].Negocio.ClienteXContaBancaria()
                        objClienteXContaBancaria.Cliente = objCliente
                        objClienteXContaBancaria.CodigoBanco = objCXCB.CodigoBanco
                        objClienteXContaBancaria.NomeBanco = objCXCB.NomeBanco
                        objClienteXContaBancaria.CodigoAgencia = objCXCB.CodigoAgencia
                        objClienteXContaBancaria.DigitoAgencia = objCXCB.DigitoAgencia
                        objClienteXContaBancaria.ContaCorrente = objCXCB.ContaCorrente
                        objClienteXContaBancaria.DigitoConta = objCXCB.DigitoConta
                        objClienteXContaBancaria.TipoConta = objCXCB.TipoConta
                        objClienteXContaBancaria.Estado = objCXCB.Estado
                        objClienteXContaBancaria.Cidade = objCXCB.Cidade
                        objClienteXContaBancaria.Praca = objCXCB.Praca
                        objClienteXContaBancaria.Observacoes = objCXCB.Observacoes

                        objClienteXContaBancaria.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.ContasBancarias.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXContaBancaria.Salvar() Then
                                objCliente.ContasBancarias.RemoveAt(row.RowIndex)
                                MsgBox(Me.Page, "Conta Excluída com Sucesso!")
                            End If
                            Exit For
                        End If
                    End If
                Next

                SalvarSessionCliente()
                Limpar_ContasBancarias()
                Carregar_ContasBancarias()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovoCB_Click(sender As Object, e As EventArgs) Handles lnkNovoCB.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf ddlBancoContaBancaria.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Banco não foi selecionado.")
                ElseIf txtAgenciaContaBancaria.Text.Length = 0 Then
                    MsgBox(Me.Page, "Agência não foi informada.")
                ElseIf txtCtaContaBancaria.Text.Length = 0 Then
                    MsgBox(Me.Page, "Conta não foi informada.")
                ElseIf txtDGCtaContaBancaria.Text.Length = 0 Then
                    MsgBox(Me.Page, "Dígito da Conta não foi informado.")
                ElseIf ddlEstadoContaBancaria.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Estado da Agência não foi selecionado.")
                ElseIf txtCidadeContaBancaria.Text.Length = 0 Then
                    MsgBox(Me.Page, "Cidade da Agência não foi informada.")
                ElseIf txtPracaContaBancaria.Text.Length = 0 Then
                    MsgBox(Me.Page, "Praça da Agência não foi informada.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXContaBancaria = New [Lib].Negocio.ClienteXContaBancaria()
                    objClienteXContaBancaria.Cliente = objCliente
                    objClienteXContaBancaria.CodigoBanco = ddlBancoContaBancaria.SelectedValue
                    objClienteXContaBancaria.NomeBanco = ddlBancoContaBancaria.SelectedItem.Text()
                    Dim objAgencias As New [Lib].Negocio.Agencia(ddlBancoContaBancaria.SelectedValue, txtAgenciaContaBancaria.Text.Trim, txtDGAgContaBancaria.Text.Trim)
                    objClienteXContaBancaria.Agencia = objAgencias
                    objClienteXContaBancaria.CodigoAgencia = txtAgenciaContaBancaria.Text.Trim
                    objClienteXContaBancaria.DigitoAgencia = txtDGAgContaBancaria.Text.Trim
                    objClienteXContaBancaria.ContaCorrente = txtCtaContaBancaria.Text.Trim
                    objClienteXContaBancaria.DigitoConta = txtDGCtaContaBancaria.Text.Trim
                    objClienteXContaBancaria.TipoConta = ddlTipoConta.SelectedValue
                    objClienteXContaBancaria.Estado = ddlEstadoContaBancaria.SelectedValue
                    objClienteXContaBancaria.Cidade = txtCidadeContaBancaria.Text
                    objClienteXContaBancaria.Praca = txtPracaContaBancaria.Text
                    objClienteXContaBancaria.Observacoes = txtObservacaoContaBancaria.Text
                    objClienteXContaBancaria.Ativo = chkAtivo.Checked

                    If objCliente.IUD = "I" Then
                        objCliente.ContasBancarias.Add(objClienteXContaBancaria)
                    Else
                        objClienteXContaBancaria.IUD = "U"
                        If ddlBancoContaBancaria.Enabled Then objClienteXContaBancaria.IUD = "I"
                        If objClienteXContaBancaria.Salvar() Then
                            If objClienteXContaBancaria.IUD = "I" Then
                                objCliente.ContasBancarias.Add(objClienteXContaBancaria)
                            Else
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).CodigoBanco = ddlBancoContaBancaria.SelectedValue
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).NomeBanco = ddlBancoContaBancaria.SelectedItem.Text()
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).CodigoAgencia = txtAgenciaContaBancaria.Text.Trim
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).DigitoAgencia = txtDGAgContaBancaria.Text.Trim
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).ContaCorrente = txtCtaContaBancaria.Text.Trim
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).DigitoConta = txtDGCtaContaBancaria.Text.Trim
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).TipoConta = ddlTipoConta.SelectedValue
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).Estado = ddlEstadoContaBancaria.SelectedValue
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).Cidade = txtCidadeContaBancaria.Text
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).Praca = txtPracaContaBancaria.Text
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).Observacoes = txtObservacaoContaBancaria.Text
                                objCliente.ContasBancarias(HgridRowIndexContaBancaria.Value).Ativo = chkAtivo.Checked
                            End If
                        End If
                    End If

                    SalvarSessionCliente()
                    Limpar_ContasBancarias()
                    Carregar_ContasBancarias()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparBC_Click(sender As Object, e As EventArgs) Handles lnkLimparBC.Click
        Try
            Limpar_ContasBancarias()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'CULTURA**********************************************************************************************************
    '*****************************************************************************************************************
    Protected Sub imgConsultarCultura_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgCultura As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgCultura.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexCultura.Value = row.RowIndex
            ddlSafra.SelectedValue = objCliente.Safras(row.RowIndex).CodigoSafra
            ddlCultura.SelectedValue = objCliente.Safras(row.RowIndex).CodigoCultura
            txtArea.Text = objCliente.Safras(row.RowIndex).AreaPlantada.ToString("N2")
            txtProdutividade.Text = objCliente.Safras(row.RowIndex).Produtividade.ToString("N2")
            txtConsumoProprio.Text = objCliente.Safras(row.RowIndex).ConsumoProprio.ToString("N2")
            txtComprometido.Text = objCliente.Safras(row.RowIndex).Comprometido.ToString("N2")
            txtEstimativaDeEntrega.Text = objCliente.Safras(row.RowIndex).EstimativaEntrega.ToString("N2")
            txtObservacaoSafra.Text = objCliente.Safras(row.RowIndex).Observacao

            ddlSafra.Enabled = False
            ddlCultura.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirCultura_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgCultura As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgCultura.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXS As [Lib].Negocio.ClienteXSafra In objCliente.Safras
                    If rowCXS.CodigoSafra = row.Cells(1).Text AndAlso
                       rowCXS.CodigoCultura = row.Cells(2).Text Then

                        objClienteXSafra = New [Lib].Negocio.ClienteXSafra()
                        objClienteXSafra.Cliente = objCliente
                        objClienteXSafra.CodigoSafra = rowCXS.CodigoSafra
                        objClienteXSafra.CodigoCultura = rowCXS.CodigoCultura
                        objClienteXSafra.AreaPlantada = rowCXS.AreaPlantada
                        objClienteXSafra.Comprometido = rowCXS.Comprometido
                        objClienteXSafra.ConsumoProprio = rowCXS.ConsumoProprio
                        objClienteXSafra.EstimativaEntrega = rowCXS.EstimativaEntrega
                        objClienteXSafra.Observacao = rowCXS.Observacao
                        objClienteXSafra.Produtividade = rowCXS.Produtividade
                        objClienteXSafra.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Safras.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXSafra.Salvar() Then
                                objCliente.Safras.RemoveAt(row.RowIndex)
                                MsgBox(Me.Page, "Sucesso na exclusão.")
                            End If
                            Exit For
                        End If
                    End If
                Next

                SalvarSessionCliente()
                Limpar_Cultura()
                Carregar_Safras()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_ClickCULT(sender As Object, e As EventArgs) Handles lnkNovoCULT.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf ddlSafra.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Safra não foi selecionada.")
                ElseIf ddlCultura.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Cultura não foi selecionada.")
                ElseIf txtArea.Text.Length = 0 Then
                    MsgBox(Me.Page, "Área não foi informada.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXSafra = New [Lib].Negocio.ClienteXSafra()
                    objClienteXSafra.Cliente = objCliente
                    objClienteXSafra.CodigoSafra = ddlSafra.SelectedValue
                    objClienteXSafra.CodigoCultura = ddlCultura.SelectedValue
                    objClienteXSafra.NomeCultura = ddlCultura.SelectedItem.Text
                    objClienteXSafra.AreaPlantada = txtArea.Text
                    objClienteXSafra.Produtividade = IIf(txtProdutividade.Text.Length = 0, 0, txtProdutividade.Text)
                    objClienteXSafra.ConsumoProprio = IIf(txtConsumoProprio.Text.Length = 0, 0, txtConsumoProprio.Text)
                    objClienteXSafra.Comprometido = IIf(txtComprometido.Text.Length = 0, 0, txtComprometido.Text)
                    objClienteXSafra.EstimativaEntrega = IIf(txtEstimativaDeEntrega.Text.Length = 0, 0, txtEstimativaDeEntrega.Text)
                    objClienteXSafra.Observacao = txtObservacaoSafra.Text

                    If objCliente.IUD = "I" Then
                        objCliente.Safras.Add(objClienteXSafra)
                    Else
                        objClienteXSafra.IUD = "U"
                        If ddlSafra.Enabled Then objClienteXSafra.IUD = "I"
                        If objClienteXSafra.Salvar() Then
                            If objClienteXSafra.IUD = "I" Then
                                objCliente.Safras.Add(objClienteXSafra)
                            Else
                                objCliente.Safras(HgridRowIndexCultura.Value).AreaPlantada = txtArea.Text
                                objCliente.Safras(HgridRowIndexCultura.Value).Produtividade = IIf(txtProdutividade.Text.Length = 0, 0, txtProdutividade.Text)
                                objCliente.Safras(HgridRowIndexCultura.Value).ConsumoProprio = IIf(txtConsumoProprio.Text.Length = 0, 0, txtConsumoProprio.Text)
                                objCliente.Safras(HgridRowIndexCultura.Value).Comprometido = IIf(txtComprometido.Text.Length = 0, 0, txtComprometido.Text)
                                objCliente.Safras(HgridRowIndexCultura.Value).EstimativaEntrega = IIf(txtEstimativaDeEntrega.Text.Length = 0, 0, txtEstimativaDeEntrega.Text)
                                objCliente.Safras(HgridRowIndexCultura.Value).Observacao = txtObservacaoSafra.Text
                            End If
                            Limpar_Cultura()
                        End If
                    End If

                    SalvarSessionCliente()

                    Carregar_Safras()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparCULT_Click(sender As Object, e As EventArgs) Handles lnkLimparCULT.Click
        Try
            Limpar_Cultura()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'VEICULOS *************************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub txtVlrOficialVeiculo_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtDataAvaliacaoVeiculo.Text.Length = 0 Then
                txtVlrOficialVeiculo.Text = String.Empty
                MsgBox(Me.Page, "Data de Avaliação do Veículo não foi informada.")
            ElseIf txtVlrOficialVeiculo.Text.Length = 0 Then
                MsgBox(Me.Page, "Valor Oficial do Veículo não foi informado.")
            Else
                txtVlrMoedaVeiculo.Text = Funcoes.ConverteParaMoedaExtrangeira(txtVlrOficialVeiculo.Text, 3, txtDataAvaliacaoVeiculo.Text.ToSqlDate()).ToString("N2")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarVeiculo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgVeiculo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgVeiculo.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexVeiculo.Value = row.RowIndex

            txtPlacaVeiculo.Text = objCliente.Veiculos(row.RowIndex).CodigoPlaca
            ddlTipoVeiculo.SelectedValue = objCliente.Veiculos(row.RowIndex).CodigoTipoDeVeiculo
            txtAnoVeiculo.Text = objCliente.Veiculos(row.RowIndex).Ano
            txtFabricanteVeiculo.Text = objCliente.Veiculos(row.RowIndex).Fabricante
            txtModeloVeiculo.Text = objCliente.Veiculos(row.RowIndex).MarcaModelo
            txtDataAvaliacaoVeiculo.Text = objCliente.Veiculos(row.RowIndex).DataAvaliacao
            txtVlrOficialVeiculo.Text = objCliente.Veiculos(row.RowIndex).ValorOficial.ToString("N2")
            txtVlrMoedaVeiculo.Text = objCliente.Veiculos(row.RowIndex).ValorMoeda.ToString("N2")
            chkVeiculoOnerado.Checked = objCliente.Veiculos(row.RowIndex).Onerado

            txtPlacaVeiculo.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirVeiculo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgVeiculo As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgVeiculo.NamingContainer, GridViewRow)

                RecuperarSessionCliente()
                For Each rowCXV As [Lib].Negocio.ClienteXVeiculo In objCliente.Veiculos
                    If rowCXV.CodigoPlaca = row.Cells(1).Text Then
                        objClienteXVeiculo = New [Lib].Negocio.ClienteXVeiculo()
                        objClienteXVeiculo.Cliente = objCliente
                        objClienteXVeiculo.IUD = "D"
                        objClienteXVeiculo.CodigoPlaca = rowCXV.CodigoPlaca
                        objClienteXVeiculo.CodigoTipoDeVeiculo = rowCXV.CodigoTipoDeVeiculo
                        objClienteXVeiculo.DescricaoTipoDeVeiculo = rowCXV.DescricaoTipoDeVeiculo
                        objClienteXVeiculo.Ano = rowCXV.Ano
                        objClienteXVeiculo.Fabricante = rowCXV.Fabricante
                        objClienteXVeiculo.MarcaModelo = rowCXV.MarcaModelo
                        objClienteXVeiculo.DataAvaliacao = rowCXV.DataAvaliacao
                        objClienteXVeiculo.ValorOficial = rowCXV.ValorOficial
                        objClienteXVeiculo.ValorMoeda = rowCXV.ValorMoeda

                        If objCliente.IUD = "I" Then
                            objCliente.Veiculos.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXVeiculo.Salvar() Then objCliente.Veiculos.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next

                SalvarSessionCliente()
                Limpar_Veiculos()
                Carregar_Veiculos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_ClickV(sender As Object, e As EventArgs) Handles lnkNovoV.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtPlacaVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Placa do Veículo não foi informada.")
                ElseIf ddlTipoVeiculo.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Tipo do Veículo não foi selecionado.")
                ElseIf txtAnoVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Ano do Veículo não foi informado.")
                ElseIf txtAnoVeiculo.Text.Length > 0 AndAlso txtAnoVeiculo.Text.Length < 4 Then
                    MsgBox(Me.Page, "Ano do Veículo deve ter 4 dígitos.")
                ElseIf txtAnoVeiculo.Text.Length > 0 AndAlso CInt(txtAnoVeiculo.Text) < 1950 Then
                    MsgBox(Me.Page, "Ano do Veículo informado não é válido.")
                ElseIf txtFabricanteVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Fabricante do Veículo não foi informado.")
                ElseIf txtModeloVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Modelo/Marca do Veículo não foi informado.")
                ElseIf txtDataAvaliacaoVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Data de Avaliação do Veículo não foi informada.")
                ElseIf txtVlrOficialVeiculo.Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Veículo não foi informado.")
                ElseIf txtVlrOficialVeiculo.Text.Length > 0 AndAlso CDbl(txtVlrOficialVeiculo.Text) = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Veículo não pode ser Zero.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXVeiculo = New [Lib].Negocio.ClienteXVeiculo()
                    objClienteXVeiculo.Cliente = objCliente
                    objClienteXVeiculo.CodigoPlaca = txtPlacaVeiculo.Text.Trim
                    objClienteXVeiculo.CodigoTipoDeVeiculo = ddlTipoVeiculo.SelectedValue
                    objClienteXVeiculo.DescricaoTipoDeVeiculo = ddlTipoVeiculo.SelectedItem.Text()
                    objClienteXVeiculo.Ano = txtAnoVeiculo.Text
                    objClienteXVeiculo.Fabricante = txtFabricanteVeiculo.Text.Trim
                    objClienteXVeiculo.MarcaModelo = txtModeloVeiculo.Text.Trim
                    objClienteXVeiculo.DataAvaliacao = txtDataAvaliacaoVeiculo.Text
                    objClienteXVeiculo.ValorOficial = txtVlrOficialVeiculo.Text
                    objClienteXVeiculo.ValorMoeda = txtVlrMoedaVeiculo.Text
                    objClienteXVeiculo.Onerado = chkVeiculoOnerado.Checked

                    If objCliente.IUD = "I" Then
                        objCliente.Veiculos.Add(objClienteXVeiculo)
                    Else
                        objClienteXVeiculo.IUD = "U"
                        If txtPlacaVeiculo.Enabled Then objClienteXVeiculo.IUD = "I"
                        If objClienteXVeiculo.Salvar() Then
                            If objClienteXVeiculo.IUD = "I" Then
                                objCliente.Veiculos.Add(objClienteXVeiculo)
                            Else
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).CodigoPlaca = txtPlacaVeiculo.Text.Trim
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).CodigoTipoDeVeiculo = ddlTipoVeiculo.SelectedValue
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).DescricaoTipoDeVeiculo = ddlTipoVeiculo.SelectedItem.Text()
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).Ano = txtAnoVeiculo.Text
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).Fabricante = txtFabricanteVeiculo.Text.Trim
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).MarcaModelo = txtModeloVeiculo.Text.Trim
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).DataAvaliacao = txtDataAvaliacaoVeiculo.Text
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).ValorOficial = txtVlrOficialVeiculo.Text
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).ValorMoeda = txtVlrMoedaVeiculo.Text
                                objCliente.Veiculos(HgridRowIndexVeiculo.Value).Onerado = chkVeiculoOnerado.Checked
                            End If
                        End If
                    End If

                    SalvarSessionCliente()
                    Limpar_Veiculos()
                    Carregar_Veiculos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparV_Click(sender As Object, e As EventArgs) Handles lnkLimparV.Click
        Try
            Limpar_Veiculos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'EQUIPAMENTOS *********************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub txtVlrOficialEquipamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
        If txtDataAvaliacaoEquipamento.Text.Length = 0 Then
            txtVlrOficialEquipamento.Text = String.Empty
            MsgBox(Me.Page, "Data de Avaliação do Equipamento não foi informada.")
        ElseIf txtVlrOficialEquipamento.Text.Length = 0 Then
            MsgBox(Me.Page, "Valor Oficial do Equipamento não foi informado.")
        Else
            txtVlrMoedaEquipamento.Text = Funcoes.ConverteParaMoedaExtrangeira(txtVlrOficialEquipamento.Text, 3, txtDataAvaliacaoEquipamento.Text.ToSqlDate()).ToString("N2")
        End If
    End Sub

    Protected Sub imgConsultarEquipamento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgEquipamento As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgEquipamento.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexEquipamento.Value = row.RowIndex

            txtRegistroEquipamento.Text = objCliente.Equipamentos(row.RowIndex).Codigo
            ddlTipoEquipamento.SelectedValue = objCliente.Equipamentos(row.RowIndex).CodigoTipoDeEquipamento
            txtAnoEquipamento.Text = objCliente.Equipamentos(row.RowIndex).Ano
            txtFabricanteEquipamento.Text = objCliente.Equipamentos(row.RowIndex).Fabricante
            txtMarcaEquipamento.Text = objCliente.Equipamentos(row.RowIndex).Marca
            txtModeloEquipamento.Text = objCliente.Equipamentos(row.RowIndex).Modelo
            txtDataAvaliacaoEquipamento.Text = objCliente.Equipamentos(row.RowIndex).DataAvaliacao
            txtVlrOficialEquipamento.Text = objCliente.Equipamentos(row.RowIndex).ValorOficial.ToString("N2")
            txtVlrMoedaEquipamento.Text = objCliente.Equipamentos(row.RowIndex).ValorMoeda.ToString("N2")
            chkEquipamentoOnerado.Checked = objCliente.Equipamentos(row.RowIndex).Onerado

            txtRegistroEquipamento.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirEquipamento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgEquipamento As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgEquipamento.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXV As [Lib].Negocio.ClienteXEquipamento In objCliente.Equipamentos
                    If rowCXV.Codigo = row.Cells(1).Text Then

                        objClienteXEquipamento = New [Lib].Negocio.ClienteXEquipamento()
                        objClienteXEquipamento.Cliente = objCliente
                        objClienteXEquipamento.Codigo = rowCXV.Codigo
                        objClienteXEquipamento.CodigoTipoDeEquipamento = rowCXV.CodigoTipoDeEquipamento
                        objClienteXEquipamento.DescricaoTipoDeEquipamento = rowCXV.DescricaoTipoDeEquipamento
                        objClienteXEquipamento.Ano = rowCXV.Ano
                        objClienteXEquipamento.Fabricante = rowCXV.Fabricante
                        objClienteXEquipamento.Marca = rowCXV.Marca
                        objClienteXEquipamento.Modelo = rowCXV.Modelo
                        objClienteXEquipamento.DataAvaliacao = rowCXV.DataAvaliacao
                        objClienteXEquipamento.ValorOficial = rowCXV.ValorOficial
                        objClienteXEquipamento.ValorMoeda = rowCXV.ValorMoeda

                        objClienteXEquipamento.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Equipamentos.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXEquipamento.Salvar() Then objCliente.Equipamentos.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next

                SalvarSessionCliente()
                Limpar_Equipamentos()
                Carregar_Equipamentos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoE_Click(sender As Object, e As EventArgs) Handles lnkNovoE.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtRegistroEquipamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Registro do Equipamento não foi informada.")
                ElseIf ddlTipoEquipamento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Tipo do Equipamento não foi selecionado.")
                ElseIf txtAnoEquipamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Ano do Equipamento não foi informado.")
                ElseIf txtAnoEquipamento.Text.Length > 0 AndAlso txtAnoEquipamento.Text.Length < 4 Then
                    MsgBox(Me.Page, "Ano do Equipamento deve ter 4 dígitos.")
                ElseIf txtAnoEquipamento.Text.Length > 0 AndAlso CInt(txtAnoEquipamento.Text) < 1950 Then
                    MsgBox(Me.Page, "Ano do Equipamento informado não é válido.")
                ElseIf txtDataAvaliacaoEquipamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Data de Avaliação do Equipamento não foi informada.")
                ElseIf txtVlrOficialEquipamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Equipamento não foi informado.")
                ElseIf txtVlrOficialEquipamento.Text.Length > 0 AndAlso CDbl(txtVlrOficialEquipamento.Text) = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Equipamento não pode ser Zero.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXEquipamento = New [Lib].Negocio.ClienteXEquipamento()
                    objClienteXEquipamento.Cliente = objCliente
                    objClienteXEquipamento.Codigo = txtRegistroEquipamento.Text.Trim
                    objClienteXEquipamento.CodigoTipoDeEquipamento = ddlTipoEquipamento.SelectedValue
                    objClienteXEquipamento.DescricaoTipoDeEquipamento = ddlTipoEquipamento.SelectedItem.Text()
                    objClienteXEquipamento.Ano = txtAnoEquipamento.Text
                    objClienteXEquipamento.Fabricante = txtFabricanteEquipamento.Text.Trim
                    objClienteXEquipamento.Marca = txtMarcaEquipamento.Text.Trim
                    objClienteXEquipamento.Modelo = txtModeloEquipamento.Text.Trim
                    objClienteXEquipamento.DataAvaliacao = txtDataAvaliacaoEquipamento.Text
                    objClienteXEquipamento.ValorOficial = txtVlrOficialEquipamento.Text
                    objClienteXEquipamento.ValorMoeda = txtVlrMoedaEquipamento.Text
                    objClienteXEquipamento.Onerado = chkEquipamentoOnerado.Checked

                    If objCliente.IUD = "I" Then
                        objCliente.Equipamentos.Add(objClienteXEquipamento)
                    Else
                        objClienteXEquipamento.IUD = "U"
                        If txtRegistroEquipamento.Enabled Then objClienteXEquipamento.IUD = "I"
                        If objClienteXEquipamento.Salvar() Then
                            If objClienteXEquipamento.IUD = "I" Then
                                objCliente.Equipamentos.Add(objClienteXEquipamento)
                            Else
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Codigo = txtRegistroEquipamento.Text.Trim
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).CodigoTipoDeEquipamento = ddlTipoEquipamento.SelectedValue
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).DescricaoTipoDeEquipamento = ddlTipoEquipamento.SelectedItem.Text()
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Ano = txtAnoEquipamento.Text.Trim
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Fabricante = txtFabricanteEquipamento.Text.Trim
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Marca = txtMarcaEquipamento.Text.Trim
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Modelo = txtModeloEquipamento.Text.Trim
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).DataAvaliacao = txtDataAvaliacaoEquipamento.Text
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).ValorOficial = txtVlrOficialEquipamento.Text
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).ValorMoeda = txtVlrMoedaEquipamento.Text
                                objCliente.Equipamentos(HgridRowIndexEquipamento.Value).Onerado = chkEquipamentoOnerado.Checked
                            End If
                        End If
                    End If

                    SalvarSessionCliente()
                    Limpar_Equipamentos()
                    Carregar_Equipamentos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparE_Click(sender As Object, e As EventArgs) Handles lnkLimparE.Click
        Try
            Limpar_Equipamentos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'IMOVEIS **************************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub ddlEstadoImovel_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(ddlEstadoImovel.SelectedValue) Then
                Session("ssUF" & HID.Value) = ddlEstadoContaBancaria.SelectedValue
                ucConsultaCodMunicipios.Limpar()
                Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioCLIxIMO" & HID.Value)
            Else
                txtCidadeImovel.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtVlrOficialImovel_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtDataAvaliacaoImovel.Text.Length = 0 Then
                txtVlrOficialImovel.Text = String.Empty
                MsgBox(Me.Page, "Data de Avaliação do Imóvel não foi informada.")
            ElseIf txtVlrOficialImovel.Text.Length = 0 Then
                MsgBox(Me.Page, "Valor Oficial do Imóvel não foi informado.")
            Else
                txtVlrMoedaImovel.Text = Funcoes.ConverteParaMoedaExtrangeira(txtVlrOficialImovel.Text, 3, txtDataAvaliacaoImovel.Text.ToSqlDate()).ToString("N2")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarImovel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgImovel As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgImovel.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexImovel.Value = row.RowIndex
            txtCodigoImovel.Text = objCliente.Imoveis(row.RowIndex).CodigoImovel
            txtDescricaoImovel.Text = objCliente.Imoveis(row.RowIndex).Descricao
            chkOneradoImovel.Checked = objCliente.Imoveis(row.RowIndex).Onerado
            ddlEstadoImovel.SelectedValue = objCliente.Imoveis(row.RowIndex).CodigoEstado
            txtCidadeImovel.Text = objCliente.Imoveis(row.RowIndex).CodigoCidade
            txtAreaTotalImovel.Text = objCliente.Imoveis(row.RowIndex).AreaTotal.ToString("N2")
            txtAreaConstruidaImovel.Text = objCliente.Imoveis(row.RowIndex).AreaConstruida.ToString("N2")
            ddlUnidadeImovel.SelectedValue = objCliente.Imoveis(row.RowIndex).CodigoUnidadeDeMedida
            txtRegistroImovel.Text = objCliente.Imoveis(row.RowIndex).NumeroRegistro
            txtCartorioImovel.Text = objCliente.Imoveis(row.RowIndex).Cartorio
            txtDataAvaliacaoImovel.Text = objCliente.Imoveis(row.RowIndex).DataAvaliacao
            txtVlrOficialImovel.Text = objCliente.Imoveis(row.RowIndex).ValorOficial.ToString("N2")
            txtVlrMoedaImovel.Text = objCliente.Imoveis(row.RowIndex).ValorMoeda.ToString("N2")
            txtObservacaoImovel.Text = objCliente.Imoveis(row.RowIndex).Observacao

            txtCodigoImovel.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirImovel_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgImovel As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgImovel.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXI As [Lib].Negocio.ClienteXImovel In objCliente.Imoveis
                    If rowCXI.CodigoImovel = row.Cells(1).Text Then

                        objClienteXImovel = New [Lib].Negocio.ClienteXImovel()
                        objClienteXImovel.Cliente = objCliente
                        objClienteXImovel.CodigoImovel = rowCXI.CodigoImovel
                        objClienteXImovel.Descricao = rowCXI.Descricao
                        objClienteXImovel.Onerado = rowCXI.Onerado
                        objClienteXImovel.CodigoCidade = rowCXI.CodigoCidade
                        objClienteXImovel.CodigoEstado = rowCXI.CodigoEstado
                        objClienteXImovel.AreaTotal = rowCXI.AreaTotal
                        objClienteXImovel.AreaConstruida = rowCXI.AreaConstruida
                        objClienteXImovel.CodigoUnidadeDeMedida = rowCXI.CodigoUnidadeDeMedida
                        objClienteXImovel.NumeroRegistro = rowCXI.NumeroRegistro
                        objClienteXImovel.DataAvaliacao = rowCXI.DataAvaliacao
                        objClienteXImovel.ValorOficial = rowCXI.ValorOficial
                        objClienteXImovel.ValorMoeda = rowCXI.ValorMoeda
                        objClienteXImovel.Observacao = rowCXI.Observacao
                        objClienteXImovel.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Imoveis.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXImovel.Salvar() Then
                                objCliente.Imoveis.RemoveAt(row.RowIndex)
                            Else
                                MsgBox(Me.Page, "Erro ao excluir Imóvel: " & Funcoes.EliminarCaracteresEspeciais(RTrim(HttpContext.Current.Session("ssMessage").ToString)) & ". Entre em contato com o Suporte de TI")
                            End If
                            Exit For
                        End If
                    End If
                Next
                SalvarSessionCliente()
                Limpar_Imoveis()
                Carregar_Imoveis()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoI_Click(sender As Object, e As EventArgs) Handles lnkNovoI.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtCodigoImovel.Text.Length = 0 Then
                    MsgBox(Me.Page, "Código do Imóvel não foi informado.")
                ElseIf txtDescricaoImovel.Text.Length = 0 Then
                    MsgBox(Me.Page, "Descricão do Imóvel não foi informada.")
                ElseIf ddlEstadoImovel.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Estado do Imóvel não foi selecionado.")
                ElseIf txtCidadeImovel.Text.Length = 0 Then
                    MsgBox(Me.Page, "Cidade do Imóvel não foi informada.")
                ElseIf txtVlrOficialImovel.Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Imóvel não foi informado.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                    objClienteXImovel = New [Lib].Negocio.ClienteXImovel()
                    objClienteXImovel.Cliente = objCliente
                    objClienteXImovel.CodigoImovel = txtCodigoImovel.Text.Trim
                    objClienteXImovel.Descricao = Funcoes.EliminarCaracteresEspeciais(txtDescricaoImovel.Text.Trim)
                    objClienteXImovel.Onerado = chkOneradoImovel.Checked
                    objClienteXImovel.CodigoCidade = txtCidadeImovel.Text.Trim
                    objClienteXImovel.CodigoEstado = ddlEstadoImovel.SelectedValue
                    objClienteXImovel.AreaTotal = IIf(txtAreaTotalImovel.Text.Length = 0, 0, txtAreaTotalImovel.Text)
                    objClienteXImovel.AreaConstruida = IIf(txtAreaConstruidaImovel.Text.Length = 0, 0, txtAreaConstruidaImovel.Text)
                    objClienteXImovel.CodigoUnidadeDeMedida = ddlUnidadeImovel.SelectedValue
                    objClienteXImovel.NumeroRegistro = IIf(txtRegistroImovel.Text.Length = 0, "", txtRegistroImovel.Text)
                    objClienteXImovel.Cartorio = IIf(txtCartorioImovel.Text.Length = 0, "", Funcoes.EliminarCaracteresEspeciais(txtCartorioImovel.Text))
                    objClienteXImovel.DataAvaliacao = IIf(txtDataAvaliacaoImovel.Text.Length = 0, Now(), txtDataAvaliacaoImovel.Text)
                    objClienteXImovel.ValorOficial = IIf(txtVlrOficialImovel.Text.Length = 0, 0, txtVlrOficialImovel.Text)
                    objClienteXImovel.ValorMoeda = IIf(txtVlrMoedaImovel.Text.Length = 0, 0, txtVlrMoedaImovel.Text)
                    objClienteXImovel.Observacao = IIf(txtObservacaoImovel.Text.Length = 0, "", Funcoes.EliminarCaracteresEspeciais(txtObservacaoImovel.Text))

                    If objCliente.IUD = "I" Then
                        objCliente.Imoveis.Add(objClienteXImovel)
                    Else
                        objClienteXImovel.IUD = "U"
                        If txtCodigoImovel.Enabled Then objClienteXImovel.IUD = "I"
                        If objClienteXImovel.Salvar() Then
                            If objClienteXImovel.IUD = "I" Then
                                objCliente.Imoveis.Add(objClienteXImovel)
                            Else
                                objCliente.Imoveis(HgridRowIndexImovel.Value).Descricao = txtDescricaoImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).Onerado = chkOneradoImovel.Checked
                                objCliente.Imoveis(HgridRowIndexImovel.Value).CodigoCidade = txtCidadeImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).CodigoEstado = ddlEstadoImovel.SelectedValue
                                objCliente.Imoveis(HgridRowIndexImovel.Value).AreaTotal = txtAreaTotalImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).AreaConstruida = txtAreaConstruidaImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).CodigoUnidadeDeMedida = ddlUnidadeImovel.SelectedValue
                                objCliente.Imoveis(HgridRowIndexImovel.Value).NumeroRegistro = txtRegistroImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).Cartorio = txtCartorioImovel.Text.Trim
                                objCliente.Imoveis(HgridRowIndexImovel.Value).DataAvaliacao = txtDataAvaliacaoImovel.Text
                                objCliente.Imoveis(HgridRowIndexImovel.Value).ValorOficial = txtVlrOficialImovel.Text
                                objCliente.Imoveis(HgridRowIndexImovel.Value).ValorMoeda = txtVlrMoedaImovel.Text
                                objCliente.Imoveis(HgridRowIndexImovel.Value).Observacao = txtObservacaoImovel.Text.Trim.ToUpper()
                            End If
                        End If
                    End If
                    SalvarSessionCliente()
                    Limpar_Imoveis()
                    Carregar_Imoveis()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparI_Click(sender As Object, e As EventArgs) Handles lnkLimparI.Click
        Try
            Limpar_Imoveis()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'ARRENDANTES **********************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub imgExcluirArrendante_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgArrendante As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgArrendante.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXA As [Lib].Negocio.ClienteXArrendante In objCliente.Arrendantes
                    If rowCXA.CodigoArrendante = row.Cells(0).Text AndAlso rowCXA.EndArrendante = row.Cells(1).Text AndAlso rowCXA.DataContrato = Format(CDate(row.Cells(3).Text), "yyyy-MM-dd") Then
                        objClienteXArrendante = New [Lib].Negocio.ClienteXArrendante
                        objClienteXArrendante.Cliente = objCliente
                        objClienteXArrendante.CodigoArrendante = rowCXA.CodigoArrendante
                        objClienteXArrendante.EndArrendante = rowCXA.EndArrendante
                        objClienteXArrendante.DataContrato = rowCXA.DataContrato
                        objClienteXArrendante.DataVencimento = rowCXA.DataVencimento
                        objClienteXArrendante.Matricula = rowCXA.Matricula
                        objClienteXArrendante.Observacao = rowCXA.Observacao
                        objClienteXArrendante.Area = rowCXA.Area

                        objClienteXArrendante.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Arrendantes.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXArrendante.Salvar() Then objCliente.Arrendantes.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next
                SalvarSessionCliente()
                Carregar_Arrendantes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoA_Click(sender As Object, e As EventArgs) Handles lnkNovoA.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                ucArrendante.Limpar()
                Popup.ConsultaDeArrendantes(Me.Page, "objArrendanteCli")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'REPRESENTANTES *******************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub btnRepresentante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCLIXR" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarRepresentante_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgRepresentante As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgRepresentante.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexRepresentante.Value = row.RowIndex
            Dim objRepresentante As New [Lib].Negocio.Cliente(objCliente.Representantes(row.RowIndex).CodigoRepresentante, objCliente.Representantes(row.RowIndex).EndRepresentante)
            Dim itemRepresentante As ListItem = Funcoes.FormatarListItemCliente(objRepresentante)
            txtRepresentante.Text = itemRepresentante.Text
            txtCodigoRepresentante.Value = itemRepresentante.Value & "-" & objRepresentante.Nome

            chkPrincipalRepresentante.Checked = objCliente.Representantes(row.RowIndex).Principal

            btnRepresentante.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirRepresentante_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgRepresentante As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgRepresentante.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXR As [Lib].Negocio.ClienteXRepresentante In objCliente.Representantes
                    If rowCXR.CodigoRepresentante = row.Cells(1).Text AndAlso rowCXR.EndRepresentante = row.Cells(2).Text Then
                        objClienteXRepresentante = New [Lib].Negocio.ClienteXRepresentante
                        objClienteXRepresentante.Cliente = objCliente
                        objClienteXRepresentante.CodigoRepresentante = rowCXR.CodigoRepresentante
                        objClienteXRepresentante.EndRepresentante = rowCXR.EndRepresentante
                        objClienteXRepresentante.Principal = rowCXR.Principal

                        objClienteXRepresentante.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Representantes.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXRepresentante.Salvar() Then objCliente.Representantes.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next
                SalvarSessionCliente()
                Limpar_Representantes()
                Carregar_Representantes()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoR_Click(sender As Object, e As EventArgs) Handles lnkNovoR.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtCodigoRepresentante.Value.ToString.Length = 0 Then
                    MsgBox(Me.Page, "Representante não foi informado.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)
                    Dim strRepresentante() As String = txtCodigoRepresentante.Value.ToString.Split("-")

                    objClienteXRepresentante = New [Lib].Negocio.ClienteXRepresentante()
                    objClienteXRepresentante.Cliente = objCliente
                    objClienteXRepresentante.CodigoRepresentante = strRepresentante(0)
                    objClienteXRepresentante.EndRepresentante = strRepresentante(1)
                    objClienteXRepresentante.NomeRepresentante = strRepresentante(2)
                    objClienteXRepresentante.Principal = CByte(chkPrincipalRepresentante.Checked)

                    If objCliente.IUD = "I" Then
                        objCliente.Representantes.Add(objClienteXRepresentante)
                    Else
                        objClienteXRepresentante.IUD = "U"
                        If btnRepresentante.Enabled Then objClienteXRepresentante.IUD = "I"
                        If objClienteXRepresentante.Salvar() Then
                            If objClienteXRepresentante.IUD = "I" Then
                                objCliente.Representantes.Add(objClienteXRepresentante)
                            Else
                                objCliente.Representantes(HgridRowIndexRepresentante.Value).CodigoRepresentante = strRepresentante(0)
                                objCliente.Representantes(HgridRowIndexRepresentante.Value).EndRepresentante = strRepresentante(1)
                                objCliente.Representantes(HgridRowIndexRepresentante.Value).NomeRepresentante = strRepresentante(2)
                                objCliente.Representantes(HgridRowIndexRepresentante.Value).Principal = CByte(chkPrincipalRepresentante.Checked)
                            End If
                        End If
                    End If
                    SalvarSessionCliente()
                    Limpar_Representantes()
                    Carregar_Representantes()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparR_Click(sender As Object, e As EventArgs) Handles lnkLimparR.Click
        Try
            Limpar_Representantes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'SOCIOS ***************************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub btnSocio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCLIXS" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarSocio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgSocio As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgSocio.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexSocio.Value = row.RowIndex
            Dim objSocio As New [Lib].Negocio.Cliente(objCliente.Socios(row.RowIndex).CodigoSocio, objCliente.Socios(row.RowIndex).EndSocio)
            Dim itemSocio As ListItem = Funcoes.FormatarListItemCliente(objSocio)
            txtSocio.Text = itemSocio.Text
            txtCodigoSocio.Value = itemSocio.Value & "-" & objSocio.Nome

            txtParticipacaoSocio.Text = objCliente.Socios(row.RowIndex).Participacao.ToString("N2")

            btnSocio.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirSocio_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgSocio As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgSocio.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXS As [Lib].Negocio.ClienteXSocio In objCliente.Socios
                    If rowCXS.CodigoSocio = row.Cells(1).Text AndAlso rowCXS.EndSocio = row.Cells(2).Text Then
                        objClienteXSocio = New [Lib].Negocio.ClienteXSocio
                        objClienteXSocio.Cliente = objCliente
                        objClienteXSocio.CodigoSocio = rowCXS.CodigoSocio
                        objClienteXSocio.EndSocio = rowCXS.EndSocio
                        objClienteXSocio.Participacao = rowCXS.Participacao

                        objClienteXSocio.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Socios.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXSocio.Salvar() Then objCliente.Socios.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next

                SalvarSessionCliente()

                Limpar_Socios()

                Carregar_Socios()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoS_Click(sender As Object, e As EventArgs) Handles lnkNovoS.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtCodigoSocio.Value.ToString.Length = 0 Then
                    MsgBox(Me.Page, "Sócio não foi informado.")
                ElseIf txtParticipacaoSocio.Text.Length = 0 Then
                    MsgBox(Me.Page, "Participação do Sócio não foi informada.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)
                    Dim strSocio() As String = txtCodigoSocio.Value.ToString.Split("-")

                    objClienteXSocio = New [Lib].Negocio.ClienteXSocio()
                    objClienteXSocio.Cliente = objCliente
                    objClienteXSocio.CodigoSocio = strSocio(0)
                    objClienteXSocio.EndSocio = strSocio(1)
                    objClienteXSocio.NomeSocio = strSocio(2)
                    objClienteXSocio.Participacao = txtParticipacaoSocio.Text

                    If objCliente.IUD = "I" Then
                        objCliente.Socios.Add(objClienteXSocio)
                    Else
                        objClienteXSocio.IUD = "U"
                        If btnSocio.Enabled Then objClienteXSocio.IUD = "I"
                        If objClienteXSocio.Salvar() Then
                            If objClienteXSocio.IUD = "I" Then
                                objCliente.Socios.Add(objClienteXSocio)
                            Else
                                objCliente.Socios(HgridRowIndexSocio.Value).CodigoSocio = strSocio(0)
                                objCliente.Socios(HgridRowIndexSocio.Value).EndSocio = strSocio(1)
                                objCliente.Socios(HgridRowIndexSocio.Value).NomeSocio = strSocio(2)
                                objCliente.Socios(HgridRowIndexSocio.Value).Participacao = txtParticipacaoSocio.Text
                            End If
                        End If
                    End If
                    SalvarSessionCliente()
                    Limpar_Socios()
                    Carregar_Socios()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparS_Click(sender As Object, e As EventArgs) Handles lnkLimparS.Click
        Try
            Limpar_Socios()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'MATRICULAS ***********************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub imgConsultarMatricula_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgMatricula As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgMatricula.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexMatricula.Value = row.RowIndex
            txtNumeroMatricula.Text = objCliente.Matriculas(row.RowIndex).CodigoMatricula
            txtAreaMatricula.Text = objCliente.Matriculas(row.RowIndex).Area
            txtDataAvaliacaoMatricula.Text = objCliente.Matriculas(row.RowIndex).DataAvaliacao
            txtAreaMatricula.Text = objCliente.Matriculas(row.RowIndex).Area.ToString("N2")
            txtVlrOficialMatricula.Text = objCliente.Matriculas(row.RowIndex).ValorOficial.ToString("N2")

            txtRegistroMatricula.Text = objCliente.Matriculas(row.RowIndex).Registro
            txtLivroMatricula.Text = objCliente.Matriculas(row.RowIndex).Livro
            txtFolhaMatricula.Text = objCliente.Matriculas(row.RowIndex).Folha
            txtMunicipioMatricula.Text = objCliente.Matriculas(row.RowIndex).CodigoMunicipio
            ddlEstadoMatricula.SelectedValue = objCliente.Matriculas(row.RowIndex).CodigoEstado

            txtNumeroMatricula.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirMatricula_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgMatricula As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgMatricula.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXM As [Lib].Negocio.ClienteXMatricula In objCliente.Matriculas
                    If rowCXM.CodigoMatricula = row.Cells(1).Text Then
                        objClienteXMatricula = New [Lib].Negocio.ClienteXMatricula
                        objClienteXMatricula.Cliente = objCliente
                        objClienteXMatricula.CodigoMatricula = rowCXM.CodigoMatricula
                        objClienteXMatricula.Area = rowCXM.Area
                        objClienteXMatricula.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Matriculas.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXMatricula.Salvar() Then objCliente.Matriculas.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next
                SalvarSessionCliente()
                Limpar_Matricula()
                Carregar_Matriculas()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlEstadoMatricula_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(ddlEstadoMatricula.SelectedValue) Then
                Session("ssUF" & HID.Value) = ddlEstadoMatricula.SelectedValue
                ucConsultaCodMunicipios.Limpar()
                Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioCLIxA" & HID.Value)
            Else
                txtMunicipioMatricula.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoM_Click(sender As Object, e As EventArgs) Handles lnkNovoM.Click
        If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
            If Trim(txtCpfCnpj.Text).Length = 0 Then
                MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
            ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
            ElseIf Trim(txtNumeroMatricula.Text).Length = 0 Then
                MsgBox(Me.Page, "Matricula não foi informado.")
            ElseIf Trim(txtAreaMatricula.Text).Length = 0 Then
                MsgBox(Me.Page, "Área da Matrícula não foi informada.")
            ElseIf Trim(txtAreaMatricula.Text).Length > 0 AndAlso CDbl(txtAreaMatricula.Text) = 0 Then
                MsgBox(Me.Page, "Área da Matrícula não pode ser Zero(0).")
            ElseIf Trim(txtVlrOficialMatricula.Text).Length = 0 Then
                MsgBox(Me.Page, "Valor Oficial da Matrícula não foi informado.")
            ElseIf Trim(txtVlrOficialMatricula.Text).Length > 0 AndAlso CDbl(txtVlrOficialMatricula.Text) = 0 Then
                MsgBox(Me.Page, "Valor Oficial da Matrícula não pode ser Zero(0).")
            Else
                RecuperarSessionCliente()

                objCliente.Codigo = txtCpfCnpj.Text
                objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)

                objClienteXMatricula = New [Lib].Negocio.ClienteXMatricula
                objClienteXMatricula.Cliente = objCliente
                objClienteXMatricula.CodigoMatricula = txtNumeroMatricula.Text
                objClienteXMatricula.DataAvaliacao = IIf(txtDataAvaliacaoMatricula.Text.Length = 0, Now(), txtDataAvaliacaoMatricula.Text)
                objClienteXMatricula.Area = txtAreaMatricula.Text
                objClienteXMatricula.ValorOficial = txtVlrOficialMatricula.Text
                objClienteXMatricula.Registro = txtRegistroMatricula.Text
                objClienteXMatricula.Livro = txtLivroMatricula.Text
                objClienteXMatricula.Folha = txtFolhaMatricula.Text
                objClienteXMatricula.CodigoMunicipio = txtMunicipioMatricula.Text
                objClienteXMatricula.CodigoEstado = ddlEstadoMatricula.SelectedValue

                If objCliente.IUD = "I" Then
                    objCliente.Matriculas.Add(objClienteXMatricula)
                Else
                    objClienteXMatricula.IUD = "U"
                    If txtNumeroMatricula.Enabled Then objClienteXMatricula.IUD = "I"
                    If objClienteXMatricula.Salvar() Then
                        If objClienteXMatricula.IUD = "I" Then
                            objCliente.Matriculas.Add(objClienteXMatricula)
                        Else
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).CodigoMatricula = txtNumeroMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).DataAvaliacao = IIf(txtDataAvaliacaoMatricula.Text.Length = 0, Now(), txtDataAvaliacaoMatricula.Text)
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).Area = txtAreaMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).ValorOficial = txtVlrOficialMatricula.Text

                            objCliente.Matriculas(HgridRowIndexMatricula.Value).Registro = txtRegistroMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).Livro = txtLivroMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).Folha = txtFolhaMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).CodigoMunicipio = txtMunicipioMatricula.Text
                            objCliente.Matriculas(HgridRowIndexMatricula.Value).CodigoEstado = ddlEstadoMatricula.SelectedValue
                        End If
                    End If
                End If

                SalvarSessionCliente()

                Limpar_Matricula()

                Carregar_Matriculas()
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
        End If
    End Sub

    Protected Sub lnkLimparM_Click(sender As Object, e As EventArgs) Handles lnkLimparM.Click
        Try
            Limpar_Matricula()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'FINANCIAMENTOS *******************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub btnFinanciador_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Session("ssCampo") = "LivreClasse"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteCLIXF" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlGrupoProdutoFinanciamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(ddlProdutoFinanciamento, CarregarDDL.Tabela.Produto, " isnull(Situacao,1) = 1 AND Grupo = '" & ddlGrupoProdutoFinanciamento.SelectedValue & "'", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlMoedaFinanciamento_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlMoedaFinanciamento.SelectedValue = 1 Then
                txtOficialFinanciamento.Enabled = True
                txtMoedaFinanciamento.Enabled = False
            Else
                txtOficialFinanciamento.Enabled = False
                txtMoedaFinanciamento.Enabled = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtOficialFinanciamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtDataFinanciamento.Text.Length = 0 Then
                MsgBox(Me.Page, "Data do Financiamento não foi informado.")
            ElseIf txtOficialFinanciamento.Text.Length = 0 Then
                MsgBox(Me.Page, "Valor Oficial do Financiamento não foi informado.")
            Else
                txtMoedaFinanciamento.Text = Funcoes.ConverteParaMoedaExtrangeira(txtOficialFinanciamento.Text, ddlMoedaFinanciamento.SelectedValue, txtDataFinanciamento.Text.ToSqlDate()).ToString("N2")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub txtMoedaFinanciamento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtDataFinanciamento.Text.Length = 0 Then
                MsgBox(Me.Page, "Data do Financiamento não foi informado.")
            ElseIf txtMoedaFinanciamento.Text.Length = 0 Then
                MsgBox(Me.Page, "Valor Moeda do Financiamento não foi informado.")
            Else
                txtOficialFinanciamento.Text = Funcoes.ConverteParaMoedaExtrangeira(txtMoedaFinanciamento.Text, ddlMoedaFinanciamento.SelectedValue, txtDataFinanciamento.Text.ToSqlDate()).ToString("N2")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgConsultarFinanciamento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgFinanciamento As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgFinanciamento.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            HgridRowIndexFinanciamento.Value = row.RowIndex
            Dim objFinanciador As New [Lib].Negocio.Cliente(objCliente.Financiamentos(row.RowIndex).CodigoFinanciador, objCliente.Financiamentos(row.RowIndex).EndFinanciador)
            Dim itemFinanciador As ListItem = Funcoes.FormatarListItemCliente(objFinanciador)
            txtFinanciador.Text = itemFinanciador.Text
            txtCodigoFinanciador.Value = itemFinanciador.Value & "-" & objFinanciador.Nome

            txtCodigoFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).CodigoFinanciamento
            txtTipoFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).TipoFinanciamento
            ddlSafraFinanciamento.SelectedValue = objCliente.Financiamentos(row.RowIndex).CodigoSafra
            ddlGrupoProdutoFinanciamento.SelectedValue = objCliente.Financiamentos(row.RowIndex).Produto.CodigoGrupo
            ddl.Carregar(ddlProdutoFinanciamento, CarregarDDL.Tabela.Produto, " isnull(Situacao,1) = 1 AND Grupo = '" & ddlGrupoProdutoFinanciamento.SelectedValue & "'", True)
            ddlProdutoFinanciamento.SelectedValue = objCliente.Financiamentos(row.RowIndex).CodigoProduto
            txtDataFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).DataFinanciamento
            txtVencimentoFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).DataVencimento
            txtParcelaFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).NumeroDeParcelas
            txtQuantidadeFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).Quantidade.ToString("N2")
            ddlMoedaFinanciamento.SelectedValue = objCliente.Financiamentos(row.RowIndex).CodigoMoeda
            txtOficialFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).ValorOficial.ToString("N2")
            txtMoedaFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).ValorMoeda.ToString("N2")
            txtObservacaoFinanciamento.Text = objCliente.Financiamentos(row.RowIndex).Observacao

            If ddlMoedaFinanciamento.SelectedValue = 1 Then
                txtOficialFinanciamento.Enabled = True
                txtMoedaFinanciamento.Enabled = False
            Else
                txtOficialFinanciamento.Enabled = False
                txtMoedaFinanciamento.Enabled = True
            End If

            btnFinanciador.Enabled = False
            txtCodigoFinanciamento.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirFinanciamento_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgFinanciamento As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgFinanciamento.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                For Each rowCXF As [Lib].Negocio.ClienteXFinanciamento In objCliente.Financiamentos
                    If rowCXF.CodigoFinanciador = row.Cells(1).Text AndAlso rowCXF.EndFinanciador = row.Cells(2).Text Then
                        objClienteXFinanciamento = New [Lib].Negocio.ClienteXFinanciamento()
                        objClienteXFinanciamento.Cliente = objCliente
                        objClienteXFinanciamento.CodigoFinanciador = rowCXF.CodigoFinanciador
                        objClienteXFinanciamento.EndFinanciador = rowCXF.EndFinanciador
                        objClienteXFinanciamento.CodigoFinanciamento = rowCXF.CodigoFinanciamento
                        objClienteXFinanciamento.TipoFinanciamento = rowCXF.TipoFinanciamento
                        objClienteXFinanciamento.CodigoSafra = rowCXF.CodigoSafra
                        objClienteXFinanciamento.DataFinanciamento = rowCXF.DataFinanciamento
                        objClienteXFinanciamento.DataVencimento = rowCXF.DataVencimento
                        objClienteXFinanciamento.NumeroDeParcelas = rowCXF.NumeroDeParcelas
                        objClienteXFinanciamento.CodigoProduto = rowCXF.CodigoProduto
                        objClienteXFinanciamento.Quantidade = rowCXF.Quantidade
                        objClienteXFinanciamento.CodigoMoeda = rowCXF.CodigoMoeda
                        objClienteXFinanciamento.ValorOficial = rowCXF.ValorOficial
                        objClienteXFinanciamento.ValorMoeda = rowCXF.ValorMoeda
                        objClienteXFinanciamento.Observacao = rowCXF.Observacao

                        objClienteXFinanciamento.IUD = "D"

                        If objCliente.IUD = "I" Then
                            objCliente.Financiamentos.RemoveAt(row.RowIndex)
                        Else
                            If objClienteXFinanciamento.Salvar() Then objCliente.Financiamentos.RemoveAt(row.RowIndex)
                            Exit For
                        End If
                    End If
                Next
                SalvarSessionCliente()
                Limpar_Financiamentos()
                Carregar_Financiamentos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoF_Click(sender As Object, e As EventArgs) Handles lnkNovoF.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Trim(txtCpfCnpj.Text).Length = 0 Then
                    MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
                ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
                    MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto não foi informado.")
                ElseIf txtCodigoFinanciador.Value.ToString.Length = 0 Then
                    MsgBox(Me.Page, "Financiador não foi informado.")
                ElseIf txtCodigoFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Código do Financiamento não foi informado.")
                ElseIf txtTipoFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Tipo do Financiamento não foi informado.")
                ElseIf ddlSafraFinanciamento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Safra do Financiamento não foi selecionada.")
                ElseIf ddlProdutoFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Produto do Financiamento não foi selecionado.")
                ElseIf txtParcelaFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Parcela(s) do Financiamento não foi informada.")
                ElseIf txtDataFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Data do Financiamento não foi informada.")
                ElseIf txtVencimentoFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Vencimento do Financiamento não foi informado.")
                ElseIf ddlMoedaFinanciamento.SelectedIndex = 0 Then
                    MsgBox(Me.Page, "Moeda do Financiamento não foi informada.")
                ElseIf txtQuantidadeFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Quantidade do Financiamento não foi informado.")
                ElseIf txtOficialFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor Oficial do Financiamento não foi informado.")
                ElseIf txtMoedaFinanciamento.Text.Length = 0 Then
                    MsgBox(Me.Page, "Valor Moeda do Financiamento não foi informado.")
                Else
                    RecuperarSessionCliente()

                    objCliente.Codigo = txtCpfCnpj.Text
                    objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)
                    Dim strFinanciador() As String = txtCodigoFinanciador.Value.ToString.Split("-")

                    objClienteXFinanciamento = New [Lib].Negocio.ClienteXFinanciamento()
                    objClienteXFinanciamento.Cliente = objCliente
                    objClienteXFinanciamento.CodigoFinanciador = strFinanciador(0)
                    objClienteXFinanciamento.EndFinanciador = strFinanciador(1)
                    objClienteXFinanciamento.NomeFinanciador = strFinanciador(2)
                    objClienteXFinanciamento.CodigoFinanciamento = txtCodigoFinanciamento.Text
                    objClienteXFinanciamento.TipoFinanciamento = Funcoes.EliminarCaracteresEspeciais(txtTipoFinanciamento.Text)
                    objClienteXFinanciamento.CodigoSafra = ddlSafraFinanciamento.SelectedValue
                    objClienteXFinanciamento.DataFinanciamento = txtDataFinanciamento.Text
                    objClienteXFinanciamento.DataVencimento = txtVencimentoFinanciamento.Text
                    objClienteXFinanciamento.NumeroDeParcelas = txtParcelaFinanciamento.Text
                    objClienteXFinanciamento.CodigoProduto = ddlProdutoFinanciamento.SelectedValue
                    objClienteXFinanciamento.Quantidade = txtQuantidadeFinanciamento.Text
                    objClienteXFinanciamento.CodigoMoeda = ddlMoedaFinanciamento.SelectedValue
                    objClienteXFinanciamento.ValorOficial = txtOficialFinanciamento.Text
                    objClienteXFinanciamento.ValorMoeda = txtMoedaFinanciamento.Text
                    objClienteXFinanciamento.Observacao = IIf(txtObservacaoFinanciamento.Text.Length = 0, "", txtObservacaoFinanciamento.Text)

                    If objCliente.IUD = "I" Then
                        objCliente.Financiamentos.Add(objClienteXFinanciamento)
                    Else
                        objClienteXFinanciamento.IUD = "U"
                        If txtCodigoFinanciamento.Enabled Then objClienteXFinanciamento.IUD = "I"
                        If objClienteXFinanciamento.Salvar() Then
                            If objClienteXFinanciamento.IUD = "I" Then
                                objCliente.Financiamentos.Add(objClienteXFinanciamento)
                            Else
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).CodigoFinanciador = strFinanciador(0)
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).EndFinanciador = strFinanciador(1)
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).NomeFinanciador = strFinanciador(2)
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).CodigoFinanciamento = txtCodigoFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).TipoFinanciamento = Funcoes.EliminarCaracteresEspeciais(txtTipoFinanciamento.Text)
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).CodigoSafra = ddlSafraFinanciamento.SelectedValue
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).DataFinanciamento = txtDataFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).DataVencimento = txtVencimentoFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).NumeroDeParcelas = txtParcelaFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).CodigoProduto = ddlProdutoFinanciamento.SelectedValue
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).Quantidade = txtQuantidadeFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).CodigoMoeda = ddlMoedaFinanciamento.SelectedValue
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).ValorOficial = txtOficialFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).ValorMoeda = txtMoedaFinanciamento.Text
                                objCliente.Financiamentos(HgridRowIndexFinanciamento.Value).Observacao = IIf(txtObservacaoFinanciamento.Text.Length = 0, "", txtObservacaoFinanciamento.Text)
                            End If
                        End If
                    End If
                    SalvarSessionCliente()
                    Limpar_Financiamentos()
                    Carregar_Financiamentos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparF_Click(sender As Object, e As EventArgs) Handles lnkLimparF.Click
        Try
            Limpar_Financiamentos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'RECEITAS/DESPESAS ****************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub imgConsultarReceitaDespesa_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            Dim imgReceitaDespesa As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgReceitaDespesa.NamingContainer, GridViewRow)

            RecuperarSessionCliente()

            ObjClienteXReceitaDespesa = objCliente.ReceitasDespesas(row.RowIndex)

            HgridIndexRecDes.Value = ObjClienteXReceitaDespesa.Sequencia
            ddlAno.SelectedValue = ObjClienteXReceitaDespesa.Ano
            ddlReceitaDespesa.SelectedValue = ObjClienteXReceitaDespesa.ReceitaDespesa
            ddlTipoReceitaDespesa.SelectedValue = ObjClienteXReceitaDespesa.TipoReceitaDespesa
            txtValorAnoRecDes.Text = ObjClienteXReceitaDespesa.ValorAno
            txtDescricaoRecDes.Text = ObjClienteXReceitaDespesa.Descricao

            btnSocio.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirReceitaDespesa_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgRecDes As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgRecDes.NamingContainer, GridViewRow)

                RecuperarSessionCliente()


                objCliente.ReceitasDespesas(row.RowIndex).IUD = "D"

                If objCliente.IUD = "I" Then
                    objCliente.ReceitasDespesas.Remove(objCliente.ReceitasDespesas(row.RowIndex))
                Else
                    If objCliente.ReceitasDespesas(row.RowIndex).Salvar() Then
                        objCliente.ReceitasDespesas.Remove(objCliente.ReceitasDespesas(row.RowIndex))
                    End If
                End If

                SalvarSessionCliente()
                Limpar_ReceitasDespesas()
                Carregar_ReceitasDespesas()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovoRD_Click(sender As Object, e As EventArgs) Handles lnkNovoRD.Click
        Try
            If Funcoes.VerificaPermissao("Clientes", "GRAVAR") Then
                If Not IsNumeric(txtValorAnoRecDes.Text) Then
                    MsgBox(Me.Page, "Informe um Valor.")
                ElseIf CDec(txtValorAnoRecDes.Text) <= 0 Then
                    MsgBox(Me.Page, "Informe Um Valor Maior que zero")
                Else
                    RecuperarSessionCliente()

                    ObjClienteXReceitaDespesa = New [Lib].Negocio.ClienteXReceitasDespesas(objCliente)
                    ObjClienteXReceitaDespesa.Ano = ddlAno.SelectedValue
                    ObjClienteXReceitaDespesa.ReceitaDespesa = ddlReceitaDespesa.SelectedValue
                    ObjClienteXReceitaDespesa.TipoReceitaDespesa = ddlTipoReceitaDespesa.SelectedValue
                    ObjClienteXReceitaDespesa.ValorAno = CDec(txtValorAnoRecDes.Text)
                    ObjClienteXReceitaDespesa.Descricao = txtDescricaoRecDes.Text

                    If HgridIndexRecDes.Value = -1 Then
                        If objCliente.ReceitasDespesas.Count = 0 Then
                            ObjClienteXReceitaDespesa.Sequencia = 0
                        Else
                            Dim maior As Integer
                            For Each cxrd As [Lib].Negocio.ClienteXReceitasDespesas In objCliente.ReceitasDespesas
                                If cxrd.Sequencia > maior Then maior = cxrd.Sequencia
                            Next
                            ObjClienteXReceitaDespesa.Sequencia = maior + 1
                        End If
                    End If

                    If objCliente.IUD = "I" Then
                        objCliente.ReceitasDespesas.Add(ObjClienteXReceitaDespesa)
                    Else
                        ObjClienteXReceitaDespesa.IUD = "U"
                        If HgridIndexRecDes.Value = -1 Then ObjClienteXReceitaDespesa.IUD = "I"
                        If ObjClienteXReceitaDespesa.Salvar() Then
                            If ObjClienteXReceitaDespesa.IUD = "I" Then
                                objCliente.ReceitasDespesas.Add(ObjClienteXReceitaDespesa)
                            Else
                                objCliente.ReceitasDespesas(HgridIndexRecDes.Value).Ano = ObjClienteXReceitaDespesa.Ano
                                objCliente.ReceitasDespesas(HgridIndexRecDes.Value).ReceitaDespesa = ObjClienteXReceitaDespesa.ReceitaDespesa
                                objCliente.ReceitasDespesas(HgridIndexRecDes.Value).TipoReceitaDespesa = ObjClienteXReceitaDespesa.TipoReceitaDespesa
                                objCliente.ReceitasDespesas(HgridIndexRecDes.Value).ValorAno = ObjClienteXReceitaDespesa.ValorAno
                                objCliente.ReceitasDespesas(HgridIndexRecDes.Value).Descricao = ObjClienteXReceitaDespesa.Descricao
                            End If
                        End If
                    End If
                    SalvarSessionCliente()
                    Limpar_ReceitasDespesas()
                    Carregar_ReceitasDespesas()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para incluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub

    Protected Sub lnkLimparRD_Click(sender As Object, e As EventArgs) Handles lnkLimparRD.Click
        Try
            Limpar_ReceitasDespesas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub



    'CENTROS DE CUSTOS ****************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub lnkNovoCC_Click(sender As Object, e As EventArgs) Handles lnkNovoCC.Click
        Try
            If ddlCentroDeCusto.SelectedValue.Length > 5 Then
                MsgBox(Me.Page, "Selecione um Centro de Custo Valido")
                Exit Sub
            End If

            If CDec(txtPercResponsavelCC.Text) = 0 Then
                MsgBox(Me.Page, "Informe um valor valido para o percentual do centro de custo")
                Exit Sub
            End If

            RecuperarSessionCliente()
            Dim ClixCC As New [Lib].Negocio.ClienteXCentroDeCusto(objCliente)
            ClixCC.IUD = "I"
            ClixCC.CodigoCentroDeCusto = ddlCentroDeCusto.SelectedValue
            ClixCC.PercentualFixo = txtPercResponsavelCC.Text
            objCliente.CentrosDeCustos.Add(ClixCC)
            SalvarSessionCliente()
            gridCC.DataSource = objCliente.CentrosDeCustos.ToArray
            gridCC.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparCC_Click(sender As Object, e As EventArgs) Handles lnkLimparCC.Click
        Try
            ddlCentroDeCusto.SelectedIndex = 0
            txtPercResponsavelCC.Text = "0,00"
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirCC_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgCC As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgCC.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                objCliente.CentrosDeCustos(row.RowIndex).IUD = "D"

                If objCliente.IUD = "I" Then
                    objCliente.CentrosDeCustos.Remove(objCliente.CentrosDeCustos(row.RowIndex))
                Else
                    If objCliente.CentrosDeCustos(row.RowIndex).Salvar() Then objCliente.CentrosDeCustos.Remove(objCliente.CentrosDeCustos(row.RowIndex))
                End If

                SalvarSessionCliente()
                lnkLimparCC_Click(Nothing, Nothing)

                gridCC.DataSource = objCliente.CentrosDeCustos.ToArray
                gridCC.DataBind()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub


    'CONSULTAS ************************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub gridConsultaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles gridConsultaCliente.SelectedIndexChanged
        Try
            CarregarCliente(gridConsultaCliente.SelectedRow.Cells(1).Text(), gridConsultaCliente.SelectedRow.Cells(2).Text())
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnImpressaoClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Relatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    'SUBSTITUTO TRIBUTÁRIO ************************************************************************************************************
    '**********************************************************************************************************************************

    Protected Sub lnkNovoST_Click(sender As Object, e As EventArgs) Handles lnkNovoST.Click
        Try
            If txtSubTributario.Text.Length = 0 Then
                MsgBox(Me.Page, "Substituto tributário não informado.")
                Exit Sub
            ElseIf ddlEstadoST.SelectedValue.Length = 0 Then
                MsgBox(Me.Page, "Estado não informado.")
                Exit Sub
            End If

            RecuperarSessionCliente()
            Dim cXst As New [Lib].Negocio.ClienteXSubstitutoTributario(objCliente)
            cXst.IUD = "I"
            cXst.IESubstitutoTributario = txtSubTributario.Text
            cXst.Estado_Id = ddlEstadoST.SelectedValue
            objCliente.SubstitutoTributario.Add(cXst)
            SalvarSessionCliente()

            gridST.DataSource = objCliente.SubstitutoTributario.ToArray
            gridST.DataBind()
            lnkLimparST_Click(Nothing, Nothing)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirST_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgCC As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgCC.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                objCliente.SubstitutoTributario(row.RowIndex).IUD = "D"

                If objCliente.IUD = "I" Then
                    objCliente.SubstitutoTributario.Remove(objCliente.SubstitutoTributario(row.RowIndex))
                Else
                    If objCliente.SubstitutoTributario(row.RowIndex).Salvar() Then objCliente.SubstitutoTributario.Remove(objCliente.SubstitutoTributario(row.RowIndex))
                End If

                SalvarSessionCliente()
                lnkLimparCC_Click(Nothing, Nothing)

                gridST.DataSource = objCliente.SubstitutoTributario.ToArray
                gridST.DataBind()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparST_Click(sender As Object, e As EventArgs) Handles lnkLimparST.Click
        Try
            txtSubTributario.Text = String.Empty
            ddlEstadoST.SelectedIndex = 0

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarSubstitutoTributario()
        RecuperarSessionCliente()

        If objCliente.Tipos.Count > 0 Then
            For Each row As [Lib].Negocio.ClientexTipo In objCliente.Tipos
                If row.CodigoTipo = 1 Then
                    TabSubstitutoTributario.Visible = True
                    If objCliente IsNot Nothing AndAlso objCliente.SubstitutoTributario IsNot Nothing AndAlso objCliente.SubstitutoTributario.Count > 0 Then
                        gridST.DataSource = objCliente.SubstitutoTributario.ToArray
                        gridST.DataBind()
                    Else
                        gridST.DataSource = Nothing
                        gridST.DataBind()
                    End If
                End If
            Next
        End If
    End Sub

    'TABELA DE PREÇOS *****************************************************************************************************
    '**********************************************************************************************************************
    Protected Sub lnkNovoTP_Click(sender As Object, e As EventArgs) Handles lnkNovoTP.Click
        Try
            RecuperarSessionCliente()

            Dim ClixTP = New [Lib].Negocio.ClienteXTabelaDePreco(objCliente)

            ClixTP.IUD = "I"
            ClixTP.CodigoTabela = ddlTabelasDePrecos.SelectedValue
            ClixTP.DescricaoTipo = ddlTabelasDePrecos.SelectedItem.Text.Split("-")(1)

            If objCliente.IUD = "I" Then
                objCliente.TabelasDePrecos.Add(ClixTP)
            Else
                If ClixTP.Salvar Then
                    objCliente.TabelasDePrecos.Add(ClixTP)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            End If
            SalvarSessionCliente()
            CarregarTabelaDePrecos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparTP_Click(sender As Object, e As EventArgs) Handles lnkLimparTP.Click
        Try
            ddlTabelasDePrecos.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExcluirTP_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                Dim imgCC As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(imgCC.NamingContainer, GridViewRow)

                RecuperarSessionCliente()

                objCliente.TabelasDePrecos(row.RowIndex).IUD = "D"

                If objCliente.IUD = "I" Then
                    objCliente.TabelasDePrecos.Remove(objCliente.TabelasDePrecos(row.RowIndex))
                Else
                    If objCliente.TabelasDePrecos(row.RowIndex).Salvar() Then objCliente.TabelasDePrecos.Remove(objCliente.TabelasDePrecos(row.RowIndex))
                End If

                SalvarSessionCliente()
                lnkLimparTP_Click(Nothing, Nothing)
                CarregarTabelaDePrecos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir Registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub CarregarTabelaDePrecos()
        RecuperarSessionCliente()

        If objCliente.Tipos.Count > 0 Then
            For Each row As [Lib].Negocio.ClientexTipo In objCliente.Tipos
                If row.CodigoTipo = 6 Then
                    TabTabelaPrecos.Visible = True
                    ddl.Carregar(ddlTabelasDePrecos, CarregarDDL.Tabela.TabelaDePreco, "", True)
                    txtRepresentanteTP.Text = objCliente.Codigo & " - " & objCliente.Nome
                    txtRepresentanteTP.Enabled = False

                    If objCliente IsNot Nothing AndAlso objCliente.TabelasDePrecos IsNot Nothing AndAlso objCliente.TabelasDePrecos.Count > 0 Then
                        gridTP.DataSource = objCliente.TabelasDePrecos.Where(Function(s) s.IUD <> "D")
                    Else
                        gridTP.DataSource = Nothing
                    End If
                    gridTP.DataBind()

                End If
            Next
        End If
    End Sub


#End Region

#Region "Methods"

    'LOAD ******************************************************************************************************************
    '***********************************************************************************************************************
    Private Sub SalvarSessionCliente()
        ViewState("objSCliente") = objCliente
    End Sub

    Private Sub RecuperarSessionCliente()
        objCliente = CType(ViewState("objSCliente"), [Lib].Negocio.Cliente)
    End Sub


    'CADASTRO *****************************************************************************************************************
    '**************************************************************************************************************************
    Private Sub LimparCampos()
        ddlEmpresas.Items.Clear()

        SetStringEmptyinControls({txtReduzido, txtCpfCnpj, txtCodigoEndereco, txtNome, txtNomeFantasia, txtCorrespondencia,
                       txtCep, txtEndereco, txtNumero, txtComplemento, txtBairro, txtCidade, txtCodMunicipio,
                       txtDataNascimento, txtCidadeNatural, txtInscricao, txtRG, txtTelefone, txtFax, txtOutrosTelefones,
                       txtEmail, txtHabilitacao, txtEmailNFE, txtRegistros, txtSite, txtUltimaAtualizacao, txtRNTRCTransportador})
        SetStringEmptyinControls({ddlSituacao, ddlCategoria, ddlPais, ddlEstado, ddlRegiao, ddlEstadoCivil, ddlNaturalidade})
        SetStringEmptyinControls({radMasculino, radFeminino})
        SetStringEmptyinControls({txtCodigoInscricao, HiddenUpload})

        TabSubstitutoTributario.Visible = False

        ddlMicroRegiao.Items.Clear()

        SetarData(txtDataCliente)

        imgFoto.ImageUrl = "~/images/foto.jpg"
        imgInscricao.ImageUrl = "~/images/question.jpg"

        chkApenasAVista.Enabled = True
        chkApenasAVista.Checked = False

        lnkFichaCadastral.Parent.Visible = False
        imgFicha.Parent.Visible = False
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkDuplicar.Parent.Visible = False
        lnkConsultar.Parent.Visible = True
        txtCpfCnpj.ReadOnly = False
        txtCidade.Enabled = False
        txtCpfCnpj.Enabled = True
        CarregarTiposDeCliente()
        Limpar_Contatos()
        Limpar_Dependentes()
        Limpar_ContasBancarias()
        Limpar_Cultura()
        Limpar_Veiculos()
        Limpar_Equipamentos()
        Limpar_Imoveis()
        Limpar_Matricula()
        Limpar_Representantes()
        Limpar_Socios()
        Limpar_Financiamentos()
        lnkLimparCC_Click(Nothing, Nothing)
        lnkLimparST_Click(Nothing, Nothing)
        TabTabelaPrecos.Visible = False




        '*************************************************************************
        '**************************** Tab Arquivos ******************************
        '*************************************************************************
        txtDescricaoDocumento.Text = String.Empty
        txtNomeDoArquivo.Text = String.Empty
        gridDocumentos.DataSource = Nothing
        gridDocumentos.DataBind()

        '*************************************************************************
        '*************************************************************************

        SetClearValue({gridCC, gridContatos, gridDependentes, gridContasBancarias, gridCultura, gridVeiculos, gridEquipamentos, gridImoveis,
                       gridArrendantes, gridRepresentantes, gridSocios, gridMatriculas, gridFinanciamentos, gridReceitasDespesas, gridConsultaCliente})

        objCliente = New [Lib].Negocio.Cliente()
        objCliente.IUD = "I"
        SalvarSessionCliente()

        TabContainer1.ActiveTabIndex = 0

        ddlUsuarios.Items.Clear()
        ddlUsuarios.Items.Add("Inc.- " & Session("ssNomeUsuario"))

        txtCpfCnpj.Focus()
    End Sub

    Public Overrides Sub Carregar(obj As String)
        Try
            If Not Session("objCep" & HID.Value) Is Nothing Then
                Dim cep As Array = obj.Split("-")
                txtCep.Text = cep(0)
                txtEndereco.Text = cep(1)
                txtBairro.Text = cep(2)
                txtCidade.Text = cep(3)
                ddlEstado.SelectedValue = cep(4)
                txtCodMunicipio.Text = cep(5).ToString().Substring(2, cep(5).ToString.Length - 2)
                ddlPais.SelectedValue = "1058" 'BRASIL
                ddlPais_SelectedIndexChanged(Nothing, Nothing)
                txtNumero.Focus()
            ElseIf Session("ObjCadastro" & HID.Value) IsNot Nothing Then

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClienteCliXCorr" & HID.Value) Is Nothing Then
                RecuperarSessionCliente()
                Dim objCorrespondencia As [Lib].Negocio.Cliente = CType(Session("objClienteCliXCorr" & HID.Value), [Lib].Negocio.Cliente)
                objCliente.CodigoClienteCorrespondencia = objCorrespondencia.Codigo
                objCliente.EndClienteCorrespondencia = objCorrespondencia.CodigoEndereco
                txtCorrespondencia.Text = objCorrespondencia.Nome & " / " & objCorrespondencia.Complemento & " / " & objCorrespondencia.Cidade & "-" & objCorrespondencia.CodigoEstado
                Session.Remove("objClienteCliXCorr" & HID.Value)
                SalvarSessionCliente()
            ElseIf Not Session("objClienteCLIXF" & HID.Value) Is Nothing Then
                Dim objFinanciador = CType(Session("objClienteCLIXF" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemFinanciador As ListItem = Funcoes.FormatarListItemCliente(objFinanciador)
                txtFinanciador.Text = itemFinanciador.Text
                txtCodigoFinanciador.Value = itemFinanciador.Value & "-" & CType(Session("objClienteCLIXF" & HID.Value), [Lib].Negocio.Cliente).Nome
                Session.Remove("objClienteCLIXF" & HID.Value)
            ElseIf Not Session("objClienteCLIXR" & HID.Value) Is Nothing Then
                Dim objRepresentante = CType(Session("objClienteCLIXR" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemRepresentante As ListItem = Funcoes.FormatarListItemCliente(objRepresentante)
                txtRepresentante.Text = itemRepresentante.Text
                txtCodigoRepresentante.Value = itemRepresentante.Value & "-" & CType(Session("objClienteCLIXR" & HID.Value), [Lib].Negocio.Cliente).Nome
                Session.Remove("objClienteCLIXR" & HID.Value)
            ElseIf Not Session("objClienteCLIXS" & HID.Value) Is Nothing Then
                Dim objSocio = CType(Session("objClienteCLIXS" & HID.Value), [Lib].Negocio.Cliente)
                Dim itemSocio As ListItem = Funcoes.FormatarListItemCliente(objSocio)
                txtSocio.Text = itemSocio.Text
                txtCodigoSocio.Value = itemSocio.Value & "-" & CType(Session("objClienteCLIXS" & HID.Value), [Lib].Negocio.Cliente).Nome
                Session.Remove("objClienteCLIXS" & HID.Value)
            ElseIf Not Session("objMunicipioCli" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioCli" & HID.Value), [Lib].Negocio.Municipio)
                txtCodMunicipio.Text = objMunicipio.CodigoIbge
                txtCidade.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioCli" & HID.Value)
            ElseIf Not Session("objMunicipioNATCLI" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioNATCLI" & HID.Value), [Lib].Negocio.Municipio)
                txtCidadeNatural.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioNATCLI" & HID.Value)
            ElseIf Not Session("objMunicipioCLIxIMO" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioCLIxIMO" & HID.Value), [Lib].Negocio.Municipio)
                txtCidadeImovel.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioCLIxIMO" & HID.Value)
            ElseIf Not Session("objMunicipioCLIxCTA" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioCLIxCTA" & HID.Value), [Lib].Negocio.Municipio)
                txtCidadeContaBancaria.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioCLIxCTA" & HID.Value)
            ElseIf Not Session("objMunicipioCLIxA" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioCLIxA" & HID.Value), [Lib].Negocio.Municipio)
                txtMunicipioMatricula.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioCLIxA" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    '<DllImport("c:\windows\system\DllInscE32.dll", ExactSpelling:=True, SetLastError:=True)>
    'Public Shared Function ConsisteInscricaoEstadual(ByVal vInsc As String, ByVal vUF As String) As Integer
    'End Function

    Private Sub ValidarInscricao()
        If ddlEstado.SelectedIndex = 0 Then
            txtInscricao.Text = String.Empty
            MsgBox(Me.Page, "Estado não foi selecionado.")
        Else
            Try
                If Funcoes.ValidarIE(txtInscricao.Text, ddlEstado.SelectedValue) Then
                    imgInscricao.ImageUrl = "~/images/certo.jpg"
                    imgInscricao.ToolTip = "Inscrição Válida"
                    txtCodigoInscricao.Value = txtInscricao.Text
                Else
                    MsgBox(Me.Page, txtInscricao.Text & " - IE Inválida.")
                    imgInscricao.ImageUrl = "~/images/erro.jpg"
                    imgInscricao.ToolTip = "Inscrição Inválida"
                    txtCodigoInscricao.Value = txtInscricao.Text
                End If
            Catch ex As Exception
                MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
            End Try
        End If
    End Sub

    Function validarCampos() As Boolean
        RecuperarSessionCliente()
        objCliente.Codigo = txtCpfCnpj.Text.Trim
        objCliente.CodigoEndereco = IIf(String.IsNullOrWhiteSpace(txtCodigoEndereco.Text), 0, txtCodigoEndereco.Text)
        objCliente.Telefone = txtTelefone.Text.RemoveMask

        Dim temTipo As Boolean = False
        Dim tipos As Integer = 0
        Dim tipoServidor As Boolean = False
        Dim tipoEmpresa As Boolean = False
        Dim tipoTransportador As Boolean = False
        Dim i As Integer = 0

        For i = 0 To chkListTipoDeCliente.Items.Count - 1
            If chkListTipoDeCliente.Items(i).Selected = True Then
                If chkListTipoDeCliente.Items(i).Value = 1 Then tipoEmpresa = True
                If chkListTipoDeCliente.Items(i).Value = 100 Then tipoServidor = True
                If chkListTipoDeCliente.Items(i).Value = 7 Then tipoTransportador = True
                temTipo = True
                tipos += 1
            End If
        Next

        If ddlEstado.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Estado não foi selecionado.")
            Return False
        ElseIf imgInscricao.ImageUrl = "~/images/erro.jpg" Then
            MsgBox(Me.Page, "Inscrição Estadual Inválida.")
            Return False
        ElseIf Not temTipo Then
            MsgBox(Me.Page, "Tipo de Cliente não foi selecionado.")
            Return False
        ElseIf objCliente.Codigo.Length = 0 Then
            MsgBox(Me.Page, "CPF/CNPJ não foi informado.")
            Return False
        ElseIf Not tipoServidor AndAlso Not IsNumeric(objCliente.Codigo) Then
            MsgBox(Me.Page, "Código do Cliente não pode conter Letras")
            Return False
        ElseIf Not tipoServidor AndAlso Not ddlEstado.SelectedValue = "EX" AndAlso IsNumeric(objCliente.Codigo) AndAlso Not objCliente.Codigo.Length = 11 AndAlso Not objCliente.Codigo.Length = 14 Then
            MsgBox(Me.Page, "Código do Cliente deve conter 11(CPF) ou 14(CNPJ) díditos")
            Return False
        ElseIf objCliente.Codigo.Length = 11 AndAlso Not Funcoes.ValidaCPF(objCliente.Codigo) Then
            MsgBox(Me.Page, "Cpf inválido.")
            Return False
        ElseIf objCliente.Codigo.Length = 14 AndAlso Not Funcoes.ValidaCNPJ(objCliente.Codigo) Then
            MsgBox(Me.Page, "CNPJ inválido.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtBairro.Text.Trim) Then
            MsgBox(Me.Page, "Bairro é obrigatório.")
            Return False
        ElseIf txtBairro.Text.Trim.Length <= 3 Then
            MsgBox(Me.Page, "Bairro deve conter no mínimo 3 caracteres.")
            Return False
        ElseIf txtBairro.Text.Length > 50 Then
            MsgBox(Me.Page, "Bairro deve conter no máximo 50 caracteres.")
            Return False
        ElseIf tipoServidor AndAlso tipos > 1 Then
            MsgBox(Me.Page, "Cliente do Tipo Servidor não pode ter outros Tipos")
            Return False
        ElseIf tipoEmpresa AndAlso txtReduzido.Text.Length = 0 Then
            MsgBox(Me.Page, "Reduzido para Tipo de Cliente Empresa é obrigatório")
            Return False
        ElseIf tipoTransportador AndAlso txtRNTRCTransportador.Text.Length = 0 Then
            MsgBox(Me.Page, "RNTRC do Transportador para Tipo de Cliente Transportadores é obrigatório")
            Return False
        ElseIf tipoTransportador AndAlso txtRNTRCTransportador.Text.Length < 8 Then
            MsgBox(Me.Page, "RNTRC do Transportador deve ter 8 dígitos")
            Return False
        ElseIf Trim(txtCodigoEndereco.Text).Length = 0 Then
            MsgBox(Me.Page, "Código do Cliente/Sub-Filial/Entreposto é obrigatório")
            Return False
        ElseIf Trim(txtNome.Text).Length = 0 Then
            MsgBox(Me.Page, "Nome não foi informado.")
            Return False
        ElseIf Trim(txtNomeFantasia.Text).Length = 0 Then
            MsgBox(Me.Page, "Nome Fantasia não foi informado.")
            Return False
        ElseIf Trim(txtEndereco.Text).Length = 0 Then
            MsgBox(Me.Page, "Endereço não foi informado.")
            Return False
        ElseIf Trim(txtCidade.Text).Length = 0 Then
            MsgBox(Me.Page, "Cidade não foi informada.")
            Return False
        ElseIf ddlPais.SelectedIndex = 0 Then
            MsgBox(Me.Page, "País não foi selecionado.")
            Return False
        ElseIf ddlRegiao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Região não foi selecionada.")
            Return False
        ElseIf ddlCategoria.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Categoria não foi selecionada.")
            Return False
        ElseIf ddlSituacao.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Situação não foi selecionada.")
            Return False
        ElseIf txtCep.Text.Length = 0 AndAlso Not ddlEstado.SelectedValue = "EX" Then
            MsgBox(Me.Page, "CEP não foi informado.")
            Return False
        ElseIf txtCep.Text.RemoveMask.Length < 8 AndAlso Not ddlEstado.SelectedValue = "EX" Then
            MsgBox(Me.Page, "CEP deve ter 8 dígitos.")
            Return False
        ElseIf objCliente.Codigo.Length = 11 AndAlso (radMasculino.Checked = False AndAlso radFeminino.Checked = False) Then
            MsgBox(Me.Page, "Selecione Masculino ou Feminino.")
            Return False
        ElseIf objCliente.Codigo.Length = 11 AndAlso ddlEstadoCivil.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Estado Civil não foi selecionado.")
            Return False
        ElseIf objCliente.Telefone.Length > 0 AndAlso objCliente.Telefone.Length > 11 Then
            MsgBox(Me.Page, "Verifique o número do Telefone pois está fora do tamanho Padrão.")
            Return False
        ElseIf objCliente.Codigo.Length = 11 AndAlso ddlCategoria.SelectedValue < 4 AndAlso String.IsNullOrWhiteSpace(txtInscricao.Text) Then
            MsgBox(Me.Page, "Inscrição não foi informada.")
            Return False
        ElseIf Not IsNumeric(objCliente.Numero) Then
            MsgBox(Me.Page, "Número do endereço inválido, apenas números!")
            Return False
        ElseIf txtNome.Text.Length > 60 Then
            MsgBox(Me.Page, "Nome deve conter no máximo 60 caracteres.")
            Return False
        ElseIf txtNomeFantasia.Text.Length > 60 Then
            MsgBox(Me.Page, "Nome Fantasia deve conter no máximo 60 caracteres.")
            Return False
        ElseIf txtEndereco.Text.Length > 60 Then
            MsgBox(Me.Page, "Endereço deve conter no máximo 60 caracteres.")
            Return False
        ElseIf txtComplemento.Text.Length > 60 Then
            MsgBox(Me.Page, "Complemento deve conter no máximo 60 caracteres.")
            Return False

        Else
            SalvarSessionCliente()
            Return True
        End If
    End Function

    Private Sub verFoto()
        Dim infoarquivo As New IO.FileInfo(filUpload.PostedFile.FileName)
        Dim strCaminho As String = Server.MapPath("FotosClientes/") & infoarquivo.Name
        If IO.File.Exists(strCaminho) Then
            MsgBox(Me.Page, "Nome de arquivo já existe na base, renomeie o arquivo e tente novamente.")
        Else
            If Upload(infoarquivo) Then
                imgFoto.ImageUrl = "~/FotosClientes/" & infoarquivo.Name
                HiddenUpload.Value = infoarquivo.Name
            End If
        End If
    End Sub

    Private Function Upload(ByVal infoarquivo As Object) As Boolean
        Try
            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(filUpload.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                'Definimos onde ele será salvo 
                Dim strCaminho As String = Server.MapPath("FotosClientes/") & infoarquivo.Name
                'Salvamos o mesmo 
                filUpload.PostedFile.SaveAs(strCaminho)
                'Mensagem de confirmacao 
                Return True
            Else
                MsgBox(Me.Page, "Selecione um arquivo!")
                Return False
            End If
        Catch ex As Exception
            'Se der algum erro, exibimos a mensagem 
            MsgBox(Me.Page, "Há erros!. " & ex.Message)
            Return False
        End Try
    End Function

    Private Sub IniciarProcesso(ByVal Tipo As String)
        If Tipo = "I" AndAlso Funcoes.VerificaPermissao("Clientes", "GRAVAR") OrElse
           Tipo = "U" AndAlso Funcoes.VerificaPermissao("Clientes", "ALTERAR") OrElse
           Tipo = "D" AndAlso Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
            If validarCampos() Then
                Dim Sqls As New ArrayList
                RecuperarSessionCliente()
                objCliente.IUD = Tipo
                objCliente.CodigoSituacao = ddlSituacao.SelectedValue
                objCliente.Reduzido = txtReduzido.Text
                objCliente.Nome = Left(Funcoes.EliminarCaracteresEspeciais(txtNome.Text.Replace("-", " ")).ToUpper(), 60)
                objCliente.Fantasia = Left(Funcoes.EliminarCaracteresEspeciais(txtNomeFantasia.Text.Replace("-", " ")).ToUpper(), 60)
                objCliente.Endereco = Left(Funcoes.EliminarCaracteresEspeciais(txtEndereco.Text).ToUpper(), 60)
                objCliente.Numero = IIf(txtNumero.Text.Length = 0, 0, txtNumero.Text)
                objCliente.Complemento = Left(Funcoes.EliminarCaracteresEspeciais(txtComplemento.Text).ToUpper(), 60)
                objCliente.Bairro = Left(Funcoes.EliminarCaracteresEspeciais(txtBairro.Text).ToUpper(), 50)
                objCliente.CodigoEstado = ddlEstado.SelectedValue
                objCliente.CodigoMunicipio = txtCodMunicipio.Text
                objCliente.Cidade = Funcoes.EliminarCaracteresEspeciais(txtCidade.Text).ToUpper()
                objCliente.Municipio.CodigoEstado = objCliente.CodigoEstado
                objCliente.Municipio.CodigoMunicipio = objCliente.Cidade
                objCliente.CodigoRegiao = ddlRegiao.SelectedValue

                objCliente.ApenasAVista = chkApenasAVista.Checked

                If ddlMicroRegiao.SelectedIndex > 0 Then objCliente.CodigoMicroRegiao = ddlMicroRegiao.SelectedValue

                objCliente.CEP = txtCep.Text.RemoveMask
                objCliente.CodigoPais = ddlPais.SelectedValue
                objCliente.ClienteDesde = IIf(txtDataCliente.Text.Length = 0, Now(), txtDataCliente.Text)

                If imgFoto.ImageUrl.ToString.Length > 0 Then
                    objCliente.Imagem = HiddenUpload.Value
                End If

                If txtCpfCnpj.Text.Trim.Length = 11 Then
                    If radMasculino.Checked = True Then
                        objCliente.Sexo = "M"
                    ElseIf radFeminino.Checked = True Then
                        objCliente.Sexo = "F"
                    End If
                    objCliente.EstadoCivil = ddlEstadoCivil.SelectedValue
                    objCliente.RG = txtRG.Text.RemoveMask
                    objCliente.Habilitacao = txtHabilitacao.Text
                End If
                objCliente.NascimentoConstituicao = IIf(txtDataNascimento.Text.Length = 0, Now(), txtDataNascimento.Text)

                If ddlNaturalidade.SelectedIndex > 0 Then
                    objCliente.NaturalidadeEstado = ddlNaturalidade.SelectedValue
                End If

                If txtCidadeNatural.Text.Length > 0 Then
                    objCliente.NaturalidadeCidade = txtCidadeNatural.Text.ToUpper()
                End If

                If txtInscricao.Text.Length > 0 Then
                    objCliente.InscricaoEstadual = txtInscricao.Text
                Else
                    objCliente.InscricaoEstadual = String.Empty
                End If

                objCliente.Telefone = txtTelefone.Text.RemoveMask

                'Conforme solicitação email 26/05/2021 - Furlan
                If Left(Session("ssEmpresa"), 8) = "05366261" OrElse Left(Session("ssEmpresa"), 8) = "38198213" OrElse Left(Session("ssEmpresa"), 8) = "40938762" OrElse Left(Session("ssEmpresa"), 8) = "62747840" OrElse Left(Session("ssEmpresa"), 8) = "62780383" OrElse Left(Session("ssEmpresa"), 8) = "63358210" Then
                    If objCliente.Telefone.Length = 0 Then
                        MsgBox(Me.Page, "Telefone não foi informado.")
                        Exit Sub
                    ElseIf objCliente.Telefone.Length < 10 Then
                        MsgBox(Me.Page, "Telefone não foi informado corretamente.")
                        Exit Sub
                    End If
                End If

                objCliente.Fax = txtFax.Text.RemoveMask
                objCliente.RNTRCTransportador = txtRNTRCTransportador.Text
                objCliente.OutrosTelefones = txtOutrosTelefones.Text
                objCliente.Email = Trim(txtEmail.Text)
                objCliente.EmailNFE = Trim(txtEmailNFE.Text)
                objCliente.OrgaoRegCategoria = txtRegistros.Text
                objCliente.CodigoCategoria = ddlCategoria.SelectedValue
                objCliente.Site = txtSite.Text

                Select Case Tipo
                    Case "I"
                        objCliente.UsuarioInclusao = Session("ssNomeUsuario")
                        objCliente.UsuarioInclusaoData = Now()
                    Case "U"
                        objCliente.UsuarioAlteracao = Session("ssNomeUsuario")
                        objCliente.UsuarioAlteracaoData = Now()
                End Select

                If objCliente.Salvar Then
                    LimparCampos()
                    Select Case Tipo
                        Case "I"
                            MsgBox(Me.Page, "Registro incluído com Sucesso.", eTitulo.Sucess)
                        Case "U"
                            MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                    End Select
                Else
                    TabContainer1.ActiveTabIndex = 0
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                TabContainer1.ActiveTabIndex = 0
            End If
        Else
            TabContainer1.ActiveTabIndex = 0
            Select Case Tipo
                Case "I"
                    MsgBox(Me.Page, "Usuário sem permissão para incluir registro!")
                Case "U"
                    MsgBox(Me.Page, "Usuário sem permissão para alterar registro!")
                Case "D"
                    MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End Select
        End If
    End Sub

    Private Function getEnderecoId(ByVal cliente As String) As Integer
        Dim endereco As Integer = 0
        Dim sql As String = "SELECT coalesce((max(Endereco_Id) + 1),'0') as Endereco_Id FROM Clientes WHERE Cliente_Id = '" & cliente & "'"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Clientes")
        If (ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            Return CInt(ds.Tables(0).Rows(0).Item("Endereco_Id"))
        End If
        Return endereco
    End Function

    Private Sub CarregarEmpresas(ByVal ddl As DropDownList, ByVal Empresa As String, ByVal endEmpresa As Integer)
        Dim sql As String = " SELECT DISTINCT c.Cliente_Id as Codigo, c.Endereco_Id, c.Reduzido, c.Nome, c.Cidade, c.Estado  " & vbCrLf &
                            "   FROM Clientes c                                                                       " & vbCrLf
        If Not String.IsNullOrWhiteSpace(Empresa) Then
            sql &= " WHERE  Left(c.Cliente_Id, 8) = '" & Left(Empresa, 8) & "'"
        End If

        Dim ds = Banco.ConsultaDataSet(sql, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido"), Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))))
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddl, 0)
        If Not String.IsNullOrWhiteSpace(Empresa) Then ddlEmpresas.SelectedValue = Empresa & "-" & endEmpresa
    End Sub

    Private Sub CarregarCliente(ByVal pCliente As String, ByVal pEndereco As String)
        CarregarEmpresas(ddlEmpresas, pCliente, pEndereco)
        objCliente = New [Lib].Negocio.Cliente(pCliente, pEndereco)
        objCliente.IUD = "U"

        imgFicha.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
        lnkDuplicar.Parent.Visible = True
        lnkConsultar.Parent.Visible = False
        txtCpfCnpj.ReadOnly = True

        txtReduzido.Text = objCliente.Reduzido
        txtCpfCnpj.Text = objCliente.Codigo

        chkApenasAVista.Checked = objCliente.ApenasAVista

        If Funcoes.VerificaPermissao("Clientes", "LIBERAR") Then
            chkApenasAVista.Enabled = True
        Else
            chkApenasAVista.Enabled = False
        End If

        txtCodigoEndereco.Text = objCliente.CodigoEndereco
        txtNome.Text = objCliente.Nome
        txtNomeFantasia.Text = objCliente.Fantasia
        txtEndereco.Text = objCliente.Endereco
        txtNumero.Text = objCliente.Numero
        txtComplemento.Text = objCliente.Complemento
        txtBairro.Text = objCliente.Bairro
        ddlEstado.SelectedValue = objCliente.CodigoEstado
        txtCodMunicipio.Text = objCliente.CodigoMunicipio
        objCliente.Municipio.CodigoEstado = objCliente.CodigoEstado
        objCliente.Municipio.CodigoMunicipio = objCliente.Cidade

        If objCliente.CodigoClienteCorrespondencia.Length > 0 Then
            txtCorrespondencia.Text = objCliente.ClienteCorrespondencia.Nome & " / " & objCliente.ClienteCorrespondencia.Complemento & " / " & objCliente.ClienteCorrespondencia.Cidade & "-" & objCliente.ClienteCorrespondencia.CodigoEstado
        Else
            txtCorrespondencia.Text = String.Empty
        End If

        If objCliente.CodigoEstado = "EX" Then
            txtCidade.Enabled = True
        End If

        txtCidade.Text = objCliente.Cidade

        If objCliente.CodigoRegiao > 0 Then
            ddlRegiao.SelectedValue = objCliente.CodigoRegiao
            ddlRegiao_SelectedIndexChanged(Me.ddlRegiao, New EventArgs())
        End If

        If ddlRegiao.SelectedIndex > 0 Then
            ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, " Regiao_Id= " & ddlRegiao.SelectedValue)
            If objCliente.CodigoMicroRegiao > 0 Then ddlMicroRegiao.SelectedValue = objCliente.CodigoMicroRegiao
        End If

        txtCep.Text = objCliente.CEP
        ddlPais.SelectedValue = objCliente.CodigoPais
        ddlPais_SelectedIndexChanged(Nothing, Nothing)

        If objCliente.Codigo.Length = 11 Then
            If objCliente.Sexo = "M" Then radMasculino.Checked = True
            If objCliente.Sexo = "F" Then radFeminino.Checked = True
            ddlEstadoCivil.SelectedValue = objCliente.EstadoCivil
        End If

        txtDataCliente.Text = objCliente.ClienteDesde.ToShortDateString()
        txtDataNascimento.Text = objCliente.NascimentoConstituicao.ToShortDateString()
        txtUltimaAtualizacao.Text = IIf(objCliente.UsuarioAlteracaoData.ToString().Equals(String.Empty), objCliente.ClienteDesde.ToShortDateString(), objCliente.UsuarioAlteracaoData.ToShortDateString())
        ddlNaturalidade.SelectedValue = objCliente.NaturalidadeEstado
        txtCidadeNatural.Text = objCliente.NaturalidadeCidade

        If objCliente.InscricaoEstadual = "ISENTO" Then
            chkIsento.Checked = True
            txtInscricao.Enabled = False
        Else
            chkIsento.Checked = False
            txtInscricao.Enabled = True
        End If

        txtInscricao.Text = objCliente.InscricaoEstadual
        txtCodigoInscricao.Value = objCliente.InscricaoEstadual
        txtRG.Text = objCliente.RG
        txtTelefone.Text = objCliente.Telefone
        txtFax.Text = objCliente.Fax
        txtRNTRCTransportador.Text = objCliente.RNTRCTransportador
        txtOutrosTelefones.Text = objCliente.OutrosTelefones
        txtHabilitacao.Text = objCliente.Habilitacao
        txtEmail.Text = objCliente.Email
        txtRegistros.Text = objCliente.OrgaoRegCategoria
        txtEmailNFE.Text = objCliente.EmailNFE
        ddlCategoria.SelectedValue = objCliente.CodigoCategoria
        txtSite.Text = objCliente.Site
        ddlSituacao.SelectedValue = objCliente.CodigoSituacao

        If objCliente.CodigoSituacao = 50 Then
            ddlSituacao.Enabled = False
            imgBloqueio.Visible = True
        Else
            ddlSituacao.Enabled = True
            imgBloqueio.Visible = False
        End If

        If objCliente.Imagem.ToString.Length > 0 Then
            Dim strCaminho As String = Server.MapPath("FotosClientes/") & objCliente.Imagem
            If IO.File.Exists(strCaminho) Then
                HiddenUpload.Value = objCliente.Imagem
                imgFoto.ImageUrl = "~/FotosClientes/" & objCliente.Imagem
            Else
                HiddenUpload.Value = objCliente.Imagem
                imgFoto.ImageUrl = "~/Images/" & objCliente.Imagem
            End If
        End If

        If Not objCliente.InscricaoEstadual Is Nothing AndAlso objCliente.InscricaoEstadual.Length > 0 Then
            ValidarInscricao()
        End If

        For Each row As [Lib].Negocio.ClientexTipo In objCliente.Tipos
            For i As Integer = 0 To chkListTipoDeCliente.Items.Count - 1
                If chkListTipoDeCliente.Items(i).Value = row.CodigoTipo Then
                    chkListTipoDeCliente.Items(i).Selected = True
                    Exit For
                End If
            Next
        Next

        ddlUsuarios.Items.Clear()

        If objCliente.UsuarioAlteracao IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objCliente.UsuarioAlteracao) AndAlso objCliente.UsuarioInclusao <> objCliente.UsuarioAlteracao Then
            ddlUsuarios.Items.Add("Alt.- " & objCliente.UsuarioAlteracao)
            ddlUsuarios.Items.Add("Inc.- " & objCliente.UsuarioInclusao)
        Else
            ddlUsuarios.Items.Add("Inc.- " & objCliente.UsuarioInclusao)
        End If

        SalvarSessionCliente()

        Carregar_Contatos()
        Carregar_Dependentes()
        Carregar_ContasBancarias()
        Carregar_Safras()
        Carregar_Veiculos()
        Carregar_Equipamentos()
        Carregar_Imoveis()
        Carregar_Arrendantes()
        Carregar_Representantes()
        Carregar_Socios()
        Carregar_Matriculas()
        Carregar_Financiamentos()
        Carregar_ReceitasDespesas()
        Carregar_Documentos()
        CarregarSubstitutoTributario()
        CarregarTabelaDePrecos()

        gridCC.DataSource = objCliente.CentrosDeCustos
        gridCC.DataBind()


        TabContainer1.ActiveTabIndex = 0
    End Sub

    Private Sub CarregarCliente(ByVal objCliente As [Lib].Negocio.Cliente)
        LimparCampos()
        objCliente.IUD = "I"

        imgFicha.Parent.Visible = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        lnkDuplicar.Parent.Visible = False
        lnkConsultar.Parent.Visible = False
        txtCpfCnpj.ReadOnly = True

        txtReduzido.Text = objCliente.Reduzido
        txtCpfCnpj.Text = objCliente.Codigo
        txtCodigoEndereco.Text = objCliente.CodigoEndereco
        txtNome.Text = objCliente.Nome
        txtNomeFantasia.Text = objCliente.Fantasia
        txtEndereco.Text = objCliente.Endereco
        txtNumero.Text = objCliente.Numero
        txtComplemento.Text = objCliente.Complemento
        txtBairro.Text = objCliente.Bairro
        ddlEstado.SelectedValue = objCliente.CodigoEstado
        txtCodMunicipio.Text = objCliente.CodigoMunicipio
        objCliente.Municipio.CodigoEstado = objCliente.CodigoEstado
        objCliente.Municipio.CodigoMunicipio = objCliente.Cidade

        If objCliente.CodigoClienteCorrespondencia.Length > 0 Then
            txtCorrespondencia.Text = objCliente.ClienteCorrespondencia.Nome & " / " & objCliente.ClienteCorrespondencia.Complemento & " / " & objCliente.ClienteCorrespondencia.Cidade & "-" & objCliente.ClienteCorrespondencia.CodigoEstado
        Else
            txtCorrespondencia.Text = String.Empty
        End If

        If objCliente.CodigoEstado = "EX" Then
            txtCidade.Enabled = True
        End If

        txtCidade.Text = objCliente.Cidade

        ddlRegiao.SelectedValue = objCliente.CodigoRegiao

        If ddlRegiao.SelectedIndex > 0 Then
            ddl.Carregar(ddlMicroRegiao, CarregarDDL.Tabela.MicroRegiao, " Regiao_Id= " & ddlRegiao.SelectedValue)
            If objCliente.CodigoMicroRegiao > 0 Then ddlMicroRegiao.SelectedValue = objCliente.CodigoMicroRegiao
        End If

        txtCep.Text = objCliente.CEP
        ddlPais.SelectedValue = objCliente.CodigoPais
        ddlPais_SelectedIndexChanged(Nothing, Nothing)

        If objCliente.Codigo.Length = 11 Then
            If objCliente.Sexo = "M" Then radMasculino.Checked = True
            If objCliente.Sexo = "F" Then radFeminino.Checked = True
            ddlEstadoCivil.SelectedValue = objCliente.EstadoCivil
        End If

        txtDataCliente.Text = objCliente.ClienteDesde
        txtDataNascimento.Text = objCliente.NascimentoConstituicao
        ddlNaturalidade.SelectedValue = objCliente.NaturalidadeEstado
        txtCidadeNatural.Text = objCliente.NaturalidadeCidade
        txtInscricao.Text = objCliente.InscricaoEstadual
        txtCodigoInscricao.Value = objCliente.InscricaoEstadual
        txtRG.Text = objCliente.RG
        txtTelefone.Text = objCliente.Telefone
        txtFax.Text = objCliente.Fax
        txtRNTRCTransportador.Text = objCliente.RNTRCTransportador
        txtOutrosTelefones.Text = objCliente.OutrosTelefones
        txtHabilitacao.Text = objCliente.Habilitacao
        txtEmail.Text = objCliente.Email
        txtRegistros.Text = objCliente.OrgaoRegCategoria
        txtEmailNFE.Text = objCliente.EmailNFE
        ddlCategoria.SelectedValue = objCliente.CodigoCategoria
        txtSite.Text = objCliente.Site
        ddlSituacao.SelectedValue = objCliente.CodigoSituacao

        If objCliente.Imagem.ToString.Length > 0 Then
            Dim strCaminho As String = Server.MapPath("FotosClientes/") & objCliente.Imagem
            If IO.File.Exists(strCaminho) Then
                HiddenUpload.Value = objCliente.Imagem
                imgFoto.ImageUrl = "~/FotosClientes/" & objCliente.Imagem
            Else
                imgFoto.ImageUrl = "~/Images/foto.jpg"
            End If
        End If

        If Not objCliente.InscricaoEstadual Is Nothing AndAlso objCliente.InscricaoEstadual.Length > 0 Then
            ValidarInscricao()
        End If

        For Each row As [Lib].Negocio.ClientexTipo In objCliente.Tipos
            For i As Integer = 0 To chkListTipoDeCliente.Items.Count - 1
                If chkListTipoDeCliente.Items(i).Value = row.CodigoTipo Then
                    chkListTipoDeCliente.Items(i).Selected = True
                    Exit For
                End If
            Next
        Next

        SalvarSessionCliente()
        TabContainer1.ActiveTabIndex = 0
    End Sub

    Private Function getEndCorrespondencia(ByVal clienteCorrespondencia As String, ByVal EnderecoCorrespondencia As Integer, ByVal cliente As String, ByVal endereco As Integer) As [Lib].Negocio.Cliente
        Dim sql As String = "select top 1 " & vbCrLf &
                            "       cast(c.Endereco as varchar(200)) + ', ' + cast(c.Numero as varchar(10)) + ' - ' + cast(c.Complemento as varchar(100)) " & vbCrLf &
                            "       + ' - ' + cast(c.Bairro as varchar(100)) + ' - ' + cast(c.Cep as varchar(100)) + ' - ' + cast(c.Cidade as varchar(100)) " & vbCrLf &
                            "       + ' - ' + cast(c.Estado as varchar(100)) as EndCorrespondencia, " & vbCrLf &
                            "       c.EndClienteCorrespondencia, " & vbCrLf &
                            "       (case when len(c.Cliente_Id) = 11 then 'F' else 'J' end) as FisicaJuridica, " & vbCrLf &
                            "       c.Email, " & vbCrLf &
                            "       c.Sexo, " & vbCrLf &
                            "       c.Rg, " & vbCrLf &
                            "       c.Inscricao as InscricaoEstadual, " & vbCrLf &
                            "       c.[Site] " & vbCrLf &
                            "  from Clientes c where 1=1 " & vbCrLf &
                            "   AND  c.Cliente_id= CASE WHEN '" & clienteCorrespondencia & "' <> '' THEN  '" & clienteCorrespondencia & "' ELSE '" & cliente & "' end " & vbCrLf &
                            "   AND (c.ClienteCorrespondencia is null or c.ClienteCorrespondencia = '') " & vbCrLf

        Dim obj As [Lib].Negocio.Cliente = Nothing
        Dim ds = Banco.ConsultaDataSet(sql, "ClienteCorrespondencia")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            obj = New [Lib].Negocio.Cliente
            Dim row As DataRow = ds.Tables(0).Rows(0)
            obj.EndCorrespondencia = IIf(IsDBNull(row("EndCorrespondencia")), String.Empty, row("EndCorrespondencia").ToString)
            obj.FisicaJuridica = IIf(IsDBNull(row("FisicaJuridica")), String.Empty, row("FisicaJuridica").ToString)
            obj.Email = IIf(IsDBNull(row("Email")), String.Empty, row("Email").ToString)
            obj.Sexo = IIf(IsDBNull(row("Sexo")), String.Empty, row("Sexo").ToString)
            obj.RG = IIf(IsDBNull(row("Rg")), String.Empty, row("Rg").ToString)
            obj.InscricaoEstadual = IIf(IsDBNull(row("InscricaoEstadual")), String.Empty, row("InscricaoEstadual").ToString)
            obj.Site = IIf(IsDBNull(row("Site")), String.Empty, row("Site").ToString)
        End If

        Return obj
    End Function

    Private Sub Relatorio(ByVal Pdf As Boolean)
        Dim obj As New Cliente()

        obj.Codigo = txtCpfCnpj.Text.Trim()
        If ddlRegiao.SelectedIndex > 0 Then obj.CodigoRegiao = ddlRegiao.SelectedValue
        If ddlCategoria.SelectedIndex > 0 Then obj.CodigoCategoria = ddlCategoria.SelectedValue
        If ddlEstado.SelectedIndex > 0 Then obj.CodigoEstado = ddlEstado.SelectedValue
        If ddlPais.SelectedIndex > 0 Then obj.CodigoPais = ddlPais.SelectedValue
        obj.Nome = txtNome.Text.Trim
        obj.Fantasia = txtNomeFantasia.Text.Trim()
        obj.Endereco = txtEndereco.Text.Trim
        obj.Complemento = txtComplemento.Text.Trim
        obj.Bairro = txtBairro.Text.Trim
        obj.Cidade = txtCidade.Text.Trim
        If ddlSituacao.SelectedIndex > 0 Then obj.CodigoSituacao = ddlSituacao.SelectedValue

        If obj.Codigo.Length = 11 Then obj.Sexo = IIf(radMasculino.Checked, "M", "F")

        If ddlEstadoCivil.SelectedIndex > 0 Then obj.EstadoCivil = ddlEstadoCivil.SelectedValue


        If Pdf Then
            Dim ds As DataSet = obj.RelatorioDeClientes(obj)
            Funcoes.BindReport(Me.Page, ds, "Cr_Clientes", IIf(CkExcell.Checked, eExportType.ExcelCrystal, eExportType.PDF))
        Else
            Dim ds As DataSet = obj.RelatorioDeClientesExcel(obj)
            GerarExcel(ds)
        End If
    End Sub

    Private Sub GerarExcel(ByVal ds As DataSet)
        Try
            If ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                Exit Sub
            End If

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")
            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando planilha e título
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório De Clientes.")

                    'criando linha com o cabeçalho da planilha
                    Dim rowIndex As Integer = 1
                    Dim columnIndex As Integer = 1

                    'criando linha que informa o título do relatório
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "CLIENTES")
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    ' criando cabeçalho da planilha com os dados do dataset
                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next

                    'criando auto filtro na planilha
                    worksheet.Cells("A1:F" & rowIndex).AutoFilter = True

                    'formatando células numéricas
                    'worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                    'worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1

                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In ds.Tables(0).Rows
                        columnIndex = 1

                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células datas
                        worksheet.Cells(String.Format("AA{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("AC{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("AH{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"
                        worksheet.Cells(String.Format("AK{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            End Using
                        End If
                        rowIndex += 1
                    Next

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando quinta linha (cabeçalho)
                    worksheet.View.FreezePanes(6, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            'download do arquivo pelo browser
            'Download(fileName)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function getDatasetExcel() As DataSet
        Dim ds As New DataSet
        Dim Sql = " SELECT C.Cliente_Id AS Codigo, C.Endereco_Id AS CodigoEndereco, C.Nome," & vbCrLf &
              "   C.Fantasia, C.Endereco, C.Numero, C.Complemento, C.Bairro, isnull(C.Cep,'') AS Cep, C.Cidade, C.Estado AS CodigoEstado, C.Inscricao AS InscricaoEstadual,  " & vbCrLf &
              "       isnull(C.Telefone,'') as Telefone, isnull(C.Email,'') as Email, C.EmailNFE  " & vbCrLf &
              "          FROM Clientes C  " & vbCrLf &
              "          inner join ClientesXTipos ct" & vbCrLf &
              "       on ct.Cliente_Id   = C.Cliente_Id " & vbCrLf &
              "  and ct.Endereco_Id = C.Endereco_Id  " & vbCrLf &
              " WHERE (C.Categoria = 9 or ct.Tipo_Id = 7) " & vbCrLf

        ds = Banco.ConsultaDataSet(Sql, "Consistencia")

        Return ds
    End Function


    Protected Sub lnkHelp_Click(sender As Object, e As EventArgs) Handles lnkHelp.Click
        Try
            Funcoes.Ajuda(Me.Page, "Clientes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub


    'TIPOS DE CLIENTES ********************************************************************************************************
    '**************************************************************************************************************************
    Private Sub CarregarTiposDeCliente()
        chkListTipoDeCliente.Items.Clear()
        chkListTipoDeCliente.DataValueField = "CodigoTipo"
        chkListTipoDeCliente.DataTextField = "Descricao"
        Dim Lista As New [Lib].Negocio.ListTipoDeCliente(True)
        chkListTipoDeCliente.DataSource = Lista.ToArray()
        chkListTipoDeCliente.DataBind()
    End Sub


    'CONTATOS *****************************************************************************************************************
    '**************************************************************************************************************************
    Private Sub Carregar_Contatos()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Contatos IsNot Nothing AndAlso objCliente.Contatos.Count > 0 Then
            gridContatos.DataSource = objCliente.Contatos.ToArray
            gridContatos.DataBind()
        Else
            gridContatos.DataBind()
        End If
    End Sub

    Private Sub Limpar_Contatos()
        txtNomeContato.Enabled = True
        SetStringEmptyinControls({txtNomeContato, txtFuncaoContato, txtTelefoneContato, txtEmailContato, txtAgenciaContato, txtDGAgContato, txtCtaContato, txtDGCtaContato, txtObservacaoContato})
        SetStringEmptyinControls({ddlBancoContato})
        HgridRowIndexContato.Value = String.Empty
    End Sub


    'DEPENDENTES **************************************************************************************************************
    '**************************************************************************************************************************
    Private Sub Carregar_Dependentes()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Dependentes IsNot Nothing AndAlso objCliente.Dependentes.Count > 0 Then
            gridDependentes.DataSource = objCliente.Dependentes.ToArray
            gridDependentes.DataBind()
        Else
            gridDependentes.DataBind()
        End If
    End Sub

    Private Sub Limpar_Dependentes()
        txtNomeDependente.Enabled = True
        SetStringEmptyinControls({txtNomeDependente, txtRGDependente, txtCPFDependente, txtNascimentoDependente, txtProfissaoDependente, txtCustoDependente})
        HgridRowIndexDependente.Value = String.Empty
        ddlTipoDependente.SelectedValue = String.Empty
    End Sub


    'CONTAS BANCARIAS *********************************************************************************************************
    '**************************************************************************************************************************
    Private Sub Carregar_ContasBancarias()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.ContasBancarias IsNot Nothing AndAlso objCliente.ContasBancarias.Count > 0 Then
            gridContasBancarias.DataSource = objCliente.ContasBancarias.ToArray
            gridContasBancarias.DataBind()
        Else
            gridContasBancarias.DataSource = Nothing
        End If
    End Sub

    Private Sub Limpar_ContasBancarias()
        SetEnabledinControls(True, {ddlBancoContaBancaria, ddlBancoContaBancaria, txtAgenciaContaBancaria,
                                    txtDGAgContaBancaria, txtCtaContaBancaria, txtDGCtaContaBancaria})

        SetStringEmptyinControls({txtAgenciaContaBancaria, txtDGAgContaBancaria, txtCtaContaBancaria, txtDGCtaContaBancaria,
                                  txtCidadeContaBancaria, txtPracaContaBancaria, txtObservacaoContaBancaria})

        SetStringEmptyinControls({ddlBancoContaBancaria, ddlTipoConta, ddlEstadoContaBancaria})

        HgridRowIndexContaBancaria.Value = String.Empty
    End Sub

    Private Function ContemVinculoTitulo(ByVal row As GridViewRow) As Boolean
        Dim sql As String = String.Empty

        If FinanceiroNovo Then
            sql &= "select 1 from Titulos" & vbCrLf &
                   "	Where Situacao = 1             " & vbCrLf &
                   "	  And CliFor = '" & txtCpfCnpj.Text & "'" & vbCrLf &
                   "	  And EnderecoCliFor = " & txtCodigoEndereco.Text & vbCrLf &
                   "	  And BancoCliFor = '" & Server.HtmlDecode(row.Cells(1).Text) & "'" & vbCrLf &
                   "	  And AgenciaCliFor = '" & Server.HtmlDecode(row.Cells(3).Text).Trim() & "'" & vbCrLf &
                   "	  And DigitoAgenciaCliFor = '" & Server.HtmlDecode(row.Cells(4).Text).Trim() & "'" & vbCrLf &
                   "	  And ContaCliFor = '" & Server.HtmlDecode(row.Cells(5).Text) & "'" & vbCrLf &
                   "	  And DigitoContaCliFor = '" & Server.HtmlDecode(row.Cells(6).Text).Trim() & "'" & vbCrLf
        Else
            sql &= "SELECT 1 FROM ContasAPagar" & vbCrLf &
                   "	WHERE Situacao = 1                  " & vbCrLf &
                   "	And Cliente = '" & txtCpfCnpj.Text & "'" & vbCrLf &
                   "	And EndCliente = " & txtCodigoEndereco.Text & vbCrLf &
                   "	And BancoCliente = '" & Server.HtmlDecode(row.Cells(1).Text) & "'" & vbCrLf &
                   "	And AgenciaCliente = '" & HttpUtility.HtmlDecode(row.Cells(3).Text).Trim() & "'" & vbCrLf &
                   "	And DigitoAgenciaCliente = '" & Server.HtmlDecode(row.Cells(4).Text).Trim() & "'" & vbCrLf &
                   "	and ContaCliente = '" & Server.HtmlDecode(row.Cells(5).Text) & "'" & vbCrLf &
                   "	And DigitoContaCliente = '" & HttpUtility.HtmlDecode(row.Cells(6).Text).Trim() & "'" & vbCrLf &
                   "Union All                              " & vbCrLf &
                   "SELECT 2 FROM ContasAReceber 	    " & vbCrLf &
                   "	WHERE Situacao = 1                  " & vbCrLf &
                   "	And Cliente = '" & txtCpfCnpj.Text & "'" & vbCrLf &
                   "	And EndCliente = " & txtCodigoEndereco.Text & vbCrLf &
                   "	And BancoCliente = '" & Server.HtmlDecode(row.Cells(1).Text) & "'" & vbCrLf &
                   "	And AgenciaCliente = '" & Server.HtmlDecode(row.Cells(3).Text).Trim() & "'" & vbCrLf &
                   "	And DigitoAgenciaCliente = '" & Server.HtmlDecode(row.Cells(4).Text).Trim() & "'" & vbCrLf &
                   "	and ContaCliente = '" & Server.HtmlDecode(row.Cells(5).Text) & "'" & vbCrLf &
                   "	And DigitoContaCliente = '" & Server.HtmlDecode(row.Cells(6).Text).Trim() & "'" & vbCrLf
        End If

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Titulos")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If

        Return False
    End Function


    'CULTURA **************************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Safras()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Safras IsNot Nothing AndAlso objCliente.Safras.Count > 0 Then
            gridCultura.DataSource = objCliente.Safras.ToArray
            gridCultura.DataBind()
        Else
            gridCultura.DataBind()
        End If
    End Sub

    Private Sub Limpar_Cultura()
        ddlSafra.Enabled = True
        ddlCultura.Enabled = True
        ddlSafra.SelectedIndex = 0
        ddlCultura.SelectedIndex = 0
        SetStringEmptyinControls({txtArea, txtProdutividade, txtConsumoProprio, txtComprometido, txtEstimativaDeEntrega, txtObservacaoSafra})
        HgridRowIndexCultura.Value = String.Empty
    End Sub


    'VEICULOS *************************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Veiculos()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Veiculos IsNot Nothing AndAlso objCliente.Veiculos.Count > 0 Then
            gridVeiculos.DataSource = objCliente.Veiculos.ToArray
            gridVeiculos.DataBind()
        Else
            gridVeiculos.DataSource = Nothing
            gridVeiculos.DataBind()
        End If
    End Sub

    Private Sub Limpar_Veiculos()
        txtPlacaVeiculo.Enabled = True

        txtPlacaVeiculo.Text = String.Empty
        ddlTipoVeiculo.SelectedIndex = 0
        HgridRowIndexVeiculo.Value = String.Empty
        txtAnoVeiculo.Text = String.Empty
        txtFabricanteVeiculo.Text = String.Empty
        txtModeloVeiculo.Text = String.Empty
        txtDataAvaliacaoVeiculo.Text = String.Empty
        txtVlrOficialVeiculo.Text = String.Empty
        txtVlrMoedaVeiculo.Text = String.Empty
        chkVeiculoOnerado.Checked = False
    End Sub


    'EQUIPAMENTOS *********************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Equipamentos()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Equipamentos IsNot Nothing AndAlso objCliente.Equipamentos.Count > 0 Then
            gridEquipamentos.DataSource = objCliente.Equipamentos.ToArray
            gridEquipamentos.DataBind()
        Else
            gridEquipamentos.DataBind()
        End If
    End Sub

    Private Sub Limpar_Equipamentos()
        txtRegistroEquipamento.Enabled = True
        SetStringEmptyinControls({txtRegistroEquipamento, txtAnoEquipamento, txtFabricanteEquipamento, txtMarcaEquipamento, txtModeloEquipamento,
                                  txtDataAvaliacaoEquipamento, txtVlrOficialEquipamento, txtVlrMoedaEquipamento})
        ddlTipoEquipamento.SelectedIndex = 0
        HgridRowIndexEquipamento.Value = String.Empty
        chkEquipamentoOnerado.Checked = False
    End Sub


    'IMOVEIS **************************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Imoveis()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Imoveis IsNot Nothing AndAlso objCliente.Imoveis.Count > 0 Then
            gridImoveis.DataSource = objCliente.Imoveis.ToArray
            gridImoveis.DataBind()
        Else
            gridImoveis.DataBind()
        End If
    End Sub

    Private Sub Limpar_Imoveis()
        txtCodigoImovel.Enabled = True
        SetStringEmptyinControls({txtCodigoImovel, txtDescricaoImovel, txtCidadeImovel, txtAreaTotalImovel, txtAreaConstruidaImovel, txtRegistroImovel,
                                  txtDataAvaliacaoImovel, txtCartorioImovel, txtVlrOficialImovel, txtVlrMoedaImovel, txtObservacaoImovel})
        HgridRowIndexImovel.Value = String.Empty
        chkOneradoImovel.Checked = False
        ddlEstadoImovel.SelectedIndex = 0
        ddlUnidadeImovel.SelectedIndex = 0
    End Sub


    'ARRENDANTES **********************************************************************************************************
    '**********************************************************************************************************************
    Public Sub Carregar_Arrendantes()
        RecuperarSessionCliente()
        If objCliente IsNot Nothing AndAlso objCliente.Arrendantes IsNot Nothing AndAlso objCliente.Arrendantes.Count > 0 Then
            gridArrendantes.DataSource = objCliente.Arrendantes.ToArray()
            gridArrendantes.DataBind()
        Else
            gridArrendantes.DataSource = Nothing
            gridArrendantes.DataBind()
        End If
    End Sub


    'REPRESENTANTES *******************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Representantes()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Representantes IsNot Nothing AndAlso objCliente.Representantes.Count > 0 Then
            gridRepresentantes.DataSource = objCliente.Representantes.ToArray
            gridRepresentantes.DataBind()
        Else
            gridRepresentantes.DataSource = Nothing
            gridRepresentantes.DataBind()
        End If
    End Sub

    Private Sub Limpar_Representantes()
        btnRepresentante.Enabled = True

        HgridRowIndexRepresentante.Value = String.Empty
        txtRepresentante.Text = String.Empty
        txtCodigoRepresentante.Value = String.Empty
        chkPrincipalRepresentante.Checked = False
    End Sub


    'SOCIOS ***************************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Socios()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Socios IsNot Nothing AndAlso objCliente.Socios.Count > 0 Then
            gridSocios.DataSource = objCliente.Socios.ToArray
            gridSocios.DataBind()
        Else
            gridSocios.DataBind()
        End If
    End Sub

    Private Sub Limpar_Socios()
        btnSocio.Enabled = True

        HgridRowIndexSocio.Value = String.Empty
        txtSocio.Text = String.Empty
        txtCodigoSocio.Value = String.Empty
        txtParticipacaoSocio.Text = String.Empty
    End Sub


    'MATRICULAS ***********************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Matriculas()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Matriculas IsNot Nothing AndAlso objCliente.Matriculas.Count > 0 Then
            gridMatriculas.DataSource = objCliente.Matriculas.ToArray
            gridMatriculas.DataBind()
        Else
            gridMatriculas.DataBind()
        End If
    End Sub

    Private Sub Limpar_Matricula()
        txtNumeroMatricula.Enabled = True

        HgridRowIndexMatricula.Value = String.Empty
        txtNumeroMatricula.Text = String.Empty
        txtRegistroMatricula.Text = String.Empty
        txtLivroMatricula.Text = String.Empty
        txtFolhaMatricula.Text = String.Empty
        txtMunicipioMatricula.Text = String.Empty
        ddlEstadoMatricula.SelectedIndex = 0
        txtDataAvaliacaoMatricula.Text = String.Empty
        txtVlrOficialMatricula.Text = String.Empty
        txtAreaMatricula.Text = String.Empty
    End Sub


    'FINANCIAMENTOS *******************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Carregar_Financiamentos()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.Financiamentos IsNot Nothing AndAlso objCliente.Financiamentos.Count > 0 Then
            gridFinanciamentos.DataSource = objCliente.Financiamentos.ToArray
            gridFinanciamentos.DataBind()
        Else
            gridFinanciamentos.DataBind()
        End If
    End Sub

    Private Sub Limpar_Financiamentos()
        btnFinanciador.Enabled = True
        txtCodigoFinanciamento.Enabled = True
        txtOficialFinanciamento.Enabled = False
        txtMoedaFinanciamento.Enabled = False

        HgridRowIndexFinanciamento.Value = String.Empty
        txtFinanciador.Text = String.Empty
        txtCodigoFinanciador.Value = String.Empty
        txtCodigoFinanciamento.Text = String.Empty
        txtTipoFinanciamento.Text = String.Empty
        ddlSafraFinanciamento.SelectedIndex = 0
        ddlGrupoProdutoFinanciamento.SelectedIndex = 0
        ddlProdutoFinanciamento.Items.Clear()
        txtDataFinanciamento.Text = String.Empty
        txtVencimentoFinanciamento.Text = String.Empty
        txtParcelaFinanciamento.Text = String.Empty
        txtQuantidadeFinanciamento.Text = String.Empty
        ddlMoedaFinanciamento.SelectedIndex = 0
        txtOficialFinanciamento.Text = String.Empty
        txtMoedaFinanciamento.Text = String.Empty
        txtObservacaoFinanciamento.Text = String.Empty
    End Sub


    'RECEITAS/DESPESAS ****************************************************************************************************
    '**********************************************************************************************************************
    Private Sub Limpar_ReceitasDespesas()
        HgridIndexRecDes.Value = -1
        ddlAno.SelectedValue = Now.Year
        txtValorAnoRecDes.Text = "0"
        txtDescricaoRecDes.Text = String.Empty
    End Sub

    Private Sub Carregar_ReceitasDespesas()
        RecuperarSessionCliente()

        If objCliente IsNot Nothing AndAlso objCliente.ReceitasDespesas IsNot Nothing AndAlso objCliente.ReceitasDespesas.Count > 0 Then
            gridReceitasDespesas.DataSource = objCliente.ReceitasDespesas.ToArray
            gridReceitasDespesas.DataBind()
        Else
            gridReceitasDespesas.DataBind()
        End If
    End Sub

    Private Sub Carregar_Documentos()
        RecuperarSessionCliente()

        If objCliente.Documentos IsNot Nothing AndAlso objCliente.Documentos.Count > 0 Then
            gridDocumentos.DataSource = objCliente.Documentos
            gridDocumentos.DataBind()

            Dim x As Integer = 0
            While x < gridDocumentos.Rows.Count
                Dim pExtensao() As String = gridDocumentos.Rows(x).Cells(2).Text.Split(".")

                If pExtensao(1).ToUpper.Equals("XML") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                Else
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                End If

                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                x += 1
            End While

        End If
    End Sub

    Protected Sub btnAdicionar_Click(sender As Object, e As EventArgs) Handles btnAdicionar.Click

        If String.IsNullOrWhiteSpace(txtDescricaoDocumento.Text) Then
            MsgBox(Me.Page, "Descrição não foi informada", eTitulo.Info)
            txtDescricaoDocumento.Focus()
            Exit Sub
        End If

        If fupArquivo.HasFile Then
            Dim NomeDoArquivo As String = Path.GetFileName(fupArquivo.PostedFile.FileName)
            Dim TamanhoDoArquivo As Long = fupArquivo.PostedFile.ContentLength
            Dim extensao As String = Path.GetExtension(NomeDoArquivo)
            Dim contentType As String = String.Empty

            If Not extensao.ToLower.Equals(".pdf") AndAlso
                Not extensao.ToLower.Equals(".xls") AndAlso
                Not extensao.ToLower.Equals(".xlsx") AndAlso
                Not extensao.ToLower.Equals(".doc") AndAlso
                Not extensao.ToLower.Equals(".docx") Then
                MsgBox(Me.Page, "São permitidos apenas documentos com extensões pdf, xls, xlsx, doc e docx.")
                Exit Sub
            End If

            RecuperarSessionCliente()

            Dim cXd = New ClienteXDocumento(objCliente)

            cXd.IUD = "I"
            cXd.Codigo = [Lib].Negocio.Numerador.PegarNumero(HttpContext.Current.Session("ssNomeServidor"), eTiposNumerador.Documentos)
            cXd.Descricao = txtDescricaoDocumento.Text
            cXd.NomeDoArquivo = NomeDoArquivo

            cXd.Arquivo = fupArquivo.FileBytes
            txtNomeDoArquivo.Text = NomeDoArquivo

            If objCliente.IUD = "U" Then
                If cXd.Salvar Then
                    objCliente.Documentos = New ListClienteXDocumento(objCliente)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                End If
            Else
                objCliente.Documentos.Add(cXd)
            End If

            SalvarSessionCliente()

            gridDocumentos.DataSource = objCliente.Documentos
            gridDocumentos.DataBind()

            Dim x As Integer = 0
            While x < gridDocumentos.Rows.Count
                Dim pExtensao() As String = gridDocumentos.Rows(x).Cells(2).Text.Split(".")

                If pExtensao(1).ToUpper.Equals("XML") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                Else
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                End If

                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                x += 1
            End While

            txtDescricaoDocumento.Text = String.Empty
            txtNomeDoArquivo.Text = String.Empty
        Else
            MsgBox(Me.Page, "Selecione um arquivo.")
        End If
    End Sub

    Protected Sub imgDownload_Click(sender As Object, e As ImageClickEventArgs)
        Try
            RecuperarSessionCliente()

            Dim imgArquivo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgArquivo.NamingContainer, GridViewRow)

            Dim CaminhoNomeArquivo As String = Server.MapPath("~/Files/" & objCliente.Documentos(row.RowIndex).NomeDoArquivo)
            If Dir(CaminhoNomeArquivo).Length > 0 Then Kill(CaminhoNomeArquivo)

            Dim Fs As FileStream = New FileStream(CaminhoNomeArquivo, FileMode.Create)
            Fs.Write(objCliente.Documentos(row.RowIndex).Arquivo, 0, objCliente.Documentos(row.RowIndex).Arquivo.Length)
            Fs.Flush()
            Fs.Close()

            Funcoes.AbrirArquivo(Me.Page, "Files/" & objCliente.Documentos(row.RowIndex).NomeDoArquivo)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgExcluirDocumento_Click(sender As Object, e As ImageClickEventArgs)

        Try
            If Funcoes.VerificaPermissao("Clientes", "EXCLUIR") Then
                RecuperarSessionCliente()

                Dim img As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

                Dim Pos = objCliente.Documentos.FindIndex(Function(s) s.Codigo = row.Cells(0).Text())

                objCliente.Documentos(Pos).IUD = "D"

                If objCliente.IUD = "U" Then
                    If objCliente.Documentos(Pos).Salvar Then

                        objCliente.Documentos = New ListClienteXDocumento(objCliente)

                        SalvarSessionCliente()

                        gridDocumentos.DataSource = objCliente.Documentos
                        gridDocumentos.DataBind()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                        Exit Sub
                    End If
                Else
                    SalvarSessionCliente()

                    gridDocumentos.DataSource = objCliente.Documentos.Where(Function(s) s.IUD <> "D")
                    gridDocumentos.DataBind()
                End If

                Dim x As Integer = 0
                While x < gridDocumentos.Rows.Count
                    Dim pExtensao() As String = gridDocumentos.Rows(x).Cells(2).Text.Split(".")

                    If pExtensao(1).ToUpper.Equals("XML") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/xml16x16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("PDF") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icopdf16X16.jpg"
                    ElseIf pExtensao(1).ToUpper.Equals("XLS") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("XLSX") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/excel16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOC") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    ElseIf pExtensao(1).ToUpper.Equals("DOCX") Then
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/word16X16.png"
                    Else
                        CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ImageUrl = "~/images/icoJpg.gif"
                    End If

                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).ToolTip = String.Format("Download - {0}.{1}", pExtensao(0), pExtensao(1))
                    CType(gridDocumentos.Rows(x).FindControl("imgDownload"), ImageButton).Style.Value = "cursor: pointer; border: 0 none;"

                    x += 1
                End While

            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
#End Region

End Class