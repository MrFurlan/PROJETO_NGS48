<%@ Page Title="" Language="vb" AutoEventWireup="true" MasterPageFile="~/Principal.Master"
    CodeBehind="Clientes.aspx.vb" Inherits="NGS.Web.UI.Clientes" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript" src="Scripts/App/cliente.js"></script>
    <style type="text/css">
        #meioconteudo {
            width: 1135px !important;
        }

        .collbl {
            width: 150px;
        }

        .w113px {
            width: 113px;
        }

        input[type="image"], img {
            margin-top: 4px;
        }
    </style>
    <script type="text/javascript">
        $(document).ready(function () {
            clienteload();
            var prmcliente = Sys.WebForms.PageRequestManager.getInstance();
            prmcliente.add_endRequest(clienteload);
        });

        function Arquivo() {
            $("#btnAdicionar").click();
        }

    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngClientes" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlClientes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro de Clientes
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel runat="server" HeaderText="Cadastro" ID="TabPanelCadastro">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconConsultar" ID="lnkConsultar" Text="Consultar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                                        <ul>
                                            <li>
                                                <asp:LinkButton class="iconPdf" ID="lnkRelatorio" runat="server" Text="Pdf" />
                                            </li>
                                            <li>
                                                <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel Dados" />
                                            </li>
                                        </ul>
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconMais" ID="lnkDuplicar" Text="Duplicar" runat="server" />
                                    </li>
                                    <li runat="server" style="width: 126px;">
                                        <asp:LinkButton class="iconRelatorio" ID="lnkFichaCadastral" Text="Ficha Cadastral"
                                            runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkHelp" Text="Ajuda" runat="server" />
                                    </li>
                                    <div style="float: right;">
                                        <li runat="server">
                                            <div class="row" style="margin-top: 0;">
                                                <div class="coltxt">
                                                    <asp:Image ID="imgUsuario" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                                        data-ToolTip="default" ToolTip="Usuário Lançamento" Width="20px" />
                                                </div>
                                                <div class="coltxt">
                                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px" />
                                                </div>
                                            </div>
                                        </li>
                                    </div>
                                </ul>
                            </div>
                        </div>
                        <div class="painelleft" style="border-right-color: Black;">
                            <div class="row">
                                <div class="collbl" runat="server">
                                    Clientes:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEmpresas" runat="server" Width="705px" AutoPostBack="True" />
                                </div>
                            </div>
                            <div class="row" runat="server">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div id="Div1" class="coltxt" style="width: 275px;" runat="server">
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtReduzido" runat="server" Width="60px" CssClass="txtNumerico"
                                            placeholder="Reduzido" data-ToolTip="default" ToolTip="Identificação da empresa Matriz/Filial. Obrigatório apenas para empresa." />
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox CausesValidation="false" autofocus ID="txtCpfCnpj" CssClass="txtNumerico" placeholder="CPF/CNPJ" runat="server" Width="93px" AutoPostBack="True" data-ToolTip="default" ToolTip="Código do cliente/empresa." />
                                    </div>
                                    <div class="coltxt">
                                        <asp:TextBox ID="txtCodigoEndereco" runat="server" Width="40px" AutoPostBack="True" placeholder="End" Enabled="False"
                                            CssClass="txtNumerico" data-tooltip="default" ToolTip="Informe o endereço do cliente/empresa." />
                                    </div>

                                    <div class="coltxt">
                                        <asp:Button ID="btnConsultaCadastro" runat="server"
                                            CssClass="btn" UseSubmitBehavior="False" Text=">" help="" data-ToolTip="default" ToolTip="Consultar o Cadastro do CPF/CNPJ junto a Sefaz." />
                                    </div>


                                    <div class="coltxt" runat="server" style="margin-top: -5px;">
                                        <asp:ImageButton ID="imgFicha" OnClick="imgFicha_Click" runat="server" Width="30px"
                                            Style="border: 0px;" Height="30px" data-ToolTip="default" ToolTip="Visualizar Ficha Cadastral"
                                            ImageUrl="~/images/pdf.png" ImageAlign="AbsMiddle" CausesValidation="False" />
                                    </div>
                                </div>
                                <div class="collbl">
                                    Situação:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="220px" />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgBloqueio" runat="server" Height="25px" Style="border: 0px; margin-top: 0;"
                                        ImageAlign="AbsMiddle" ImageUrl="~/Images/liberar.png" data-ToolTip="default"
                                        ToolTip="Requer Autorizacao Processo: ALTERARSITUACAOCLIENTE Permissao: ALTERAR"
                                        Width="25px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Nome:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNome" runat="server" MaxLength="60" Width="255px" data-ToolTip="default" ToolTip="Identificação do cliente/empresa." />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Fantasia:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNomeFantasia" runat="server" MaxLength="60" Width="255px" data-ToolTip="default"
                                        ToolTip="Nome popular do cliente/fornecedor." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Categoria:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlCategoria" runat="server" Width="265px" />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Sexo:
                                </div>
                                <div class="coltxt">
                                    <div>
                                        <asp:RadioButton ID="radMasculino" runat="server" Text="Masculino" GroupName="Sexo"
                                            data-ToolTip="default" ToolTip="Gênero para as pessoas físicas." />
                                        <asp:RadioButton ID="radFeminino" runat="server" Text="Feminino" GroupName="Sexo"
                                            data-ToolTip="default" ToolTip="Gênero para as pessoas físicas." />
                                    </div>
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Correspondência:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCorrespondencia" runat="server" Width="632px" Enabled="False"
                                        data-ToolTip="default" ToolTip="Informe o endereço do cliente correspondente." />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnCorrespondencia" OnClick="btnCorrespondencia_Click" runat="server"
                                        CssClass="btn" UseSubmitBehavior="False" Text=">" help="" data-ToolTip="default" ToolTip="Pesquise um cliente/correspondencia para o preenchimento automático do campo." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgLimparEndCorrespondencia" OnClick="imgLimparEndCorrespondencia_Click"
                                        runat="server" Width="17px" Height="17px" ImageUrl="~/images/erro.jpg" help="Limpa o cliente/correspondecia já selecionado no campo." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    CEP:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCep" CssClass="txtCep" runat="server" data-ToolTip="default" ToolTip="Código de endereçamento postal." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgCep" OnClick="imgCep_Click" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" ImageAlign="AbsMiddle"
                                        CssClass="imgconsultar" data-ToolTip="default" ToolTip="Procura por endereço com base no Cep informado, caso o campo CPF/CNPJ estiver preenchido." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgConsultaCep" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_home20x20.png" ImageAlign="AbsMiddle"
                                        Style="margin-bottom: 4px;" data-ToolTip="default" ToolTip="Não sabe seu cep? clique aqui." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Endereço:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEndereco" runat="server" Width="255px" MaxLength="60" data-ToolTip="default" ToolTip="Local que a empresa está situada." />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Número:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtNumero" CssClass="txtNumerico" runat="server" data-ToolTip="default"
                                        ToolTip="Identificação do imóvel em seu endereço." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Complemento:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtComplemento" runat="server" MaxLength="49" Width="255px" data-ToolTip="default" ToolTip="Complemento do endereço." />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Bairro:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtBairro" runat="server" MaxLength="50" Width="255px" data-ToolTip="default"
                                        ToolTip="Referente ao local em que o endereço está situado na cidade." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    País:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlPais" runat="server" Width="265px" AutoPostBack="True" OnSelectedIndexChanged="ddlPais_SelectedIndexChanged" />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    UF:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEstado" runat="server" Width="265px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlEstado_SelectedIndexChanged" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCidade" runat="server" MaxLength="50" Width="255px" Enabled="False" data-ToolTip="default" ToolTip="Localização da empresa dentro do Estado." />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Cód. Município:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodMunicipio" runat="server" Enabled="False"
                                        data-ToolTip="default" ToolTip="Código de cadastro do município." />
                                </div>
                                <div class="coltxt">
                                    <asp:ImageButton ID="imgBtAlteraCidade" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/editdocument20x.png" ImageAlign="AbsMiddle"
                                        Style="margin-bottom: 4px;" data-ToolTip="default" ToolTip="Escolher/Alterar Município." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Região:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlRegiao" runat="server" Width="265px" AutoPostBack="True" />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Micro Região:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlMicroRegiao" runat="server" Width="265px" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Estado Civil:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEstadoCivil" runat="server" Width="265px">
                                        <asp:ListItem></asp:ListItem>
                                        <asp:ListItem Value="Solteiro">Solteiro(a)</asp:ListItem>
                                        <asp:ListItem Value="Casado">Casado(a)</asp:ListItem>
                                        <asp:ListItem Value="Separado">Separado(a)</asp:ListItem>
                                        <asp:ListItem Value="Divorciado">Divorciado(a) </asp:ListItem>
                                        <asp:ListItem Value="Viuvo">Vi&#250;vo(a)</asp:ListItem>
                                    </asp:DropDownList>
                                </div>

                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Constituição/Nascimento:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataNascimento" CssClass="calendario" runat="server" Width="80px" data-ToolTip="default"
                                        ToolTip="Data de nascimento para pessoas físicas e Data de criação para pessoas jurídicas." />
                                </div>
                                <div class="collbl" style="margin-left: 164px;">
                                    Cliente Desde:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataCliente" runat="server" CssClass="calendario" Width="80px"
                                        data-ToolTip="default" ToolTip="Data da primeira negociação entre a empresa e o cliente." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Naturalidade:
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlNaturalidade" runat="server" Width="265px" AutoPostBack="True"
                                        OnSelectedIndexChanged="ddlNaturalidade_SelectedIndexChanged" />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Cidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCidadeNatural" runat="server" Width="255px" Enabled="False"
                                        data-ToolTip="default" ToolTip="Estado e cidade de Nascimento." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Inscrição:
                                </div>
                                <div class="coltxt">
                                    <asp:HiddenField ID="txtCodigoInscricao" runat="server" />
                                    <asp:TextBox ID="txtInscricao" runat="server" CssClass="txtNumerico" OnTextChanged="txtInscricao_TextChanged"
                                        AutoPostBack="True" data-ToolTip="default" ToolTip="Inscrição estadual da empresa." />
                                </div>
                                <div class="coltxt">
                                    <asp:Image ID="imgInscricao" runat="server" Width="20px" Height="20px" ImageAlign="AbsMiddle" Style="border: 0px;"
                                        data-ToolTip="default" ToolTip="Verifica se a Inscrição Estadual é Válida" />
                                </div>
                                <div class="coltxt">
                                    <asp:CheckBox ID="chkIsento" runat="server" AutoPostBack="True" Text="Isento" />
                                </div>
                                <div class="collbl" style="margin-left: 79px;">
                                    RG:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRG" CssClass="txtrg" runat="server" data-ToolTip="default"
                                        ToolTip="Registro Geral para pessoas físicas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Telefone:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtTelefone" CssClass="txtFone" runat="server"
                                        data-ToolTip="default" ToolTip="Contato da empresa." />
                                </div>
                                <div class="collbl" style="margin-left: 165px;">
                                    Fax:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtFax" CssClass="txtFone" runat="server"
                                        data-ToolTip="default" ToolTip="Contato da empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Outros Telefones:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtOutrosTelefones" runat="server" Width="255px" MaxLength="50"
                                        help="Informe outros números do cliente, se existir." data-ToolTip="default" ToolTip="Outros contatos da empresa." />
                                </div>
                                <div class="collbl" style="margin-left: 10px;">
                                    Email:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEmail" runat="server" Width="255px" MaxLength="50"
                                        data-ToolTip="default" ToolTip="Endereço eletrônico da empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Habilitação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtHabilitacao" runat="server" data-ToolTip="default" CssClass="txtNumerico11"
                                        ToolTip="Informe o código de habilitação do cliente" />
                                </div>
                                <div class="collbl" style="margin-left: 165px;">
                                    Email NFE:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtEmailNFE" runat="server" Width="255px" data-ToolTip="default"
                                        ToolTip="Informe o endereço de email NFE(Nota Fiscal Eletronica)." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Crea/OAB/Outros:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRegistros" runat="server" data-ToolTip="default"
                                        ToolTip="Número de registro das profissões." />
                                </div>
                                <div class="collbl" style="margin-left: 165px;">
                                    Site:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtSite" runat="server" Width="255px" data-ToolTip="default"
                                        ToolTip="Página da empresa na internet." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Última Atualização:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtUltimaAtualizacao" Width="127px" runat="server" Enabled="False"
                                        Font-Bold="True" data-ToolTip="default" ToolTip="Data da atualização dos dados." />
                                </div>
                                <div class="collbl" style="margin-left: 138px;">
                                    RNTRC/Transportador:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtRNTRCTransportador" runat="server" MaxLength="8" data-ToolTip="default"
                                        ToolTip="Registro Nacional de Transportadores Rodoviários de Cargas " />
                                </div>
                            </div>
                        </div>
                        <div class="painelright " style="width: 222px; margin-left: 10px;">
                            <div class="row">
                                <div class="collbl" style="width: 190px;">
                                    <asp:CheckBox ID="chkApenasAVista" runat="server" Enabled="False" Text="Faturar apenas à Vista" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 190px;">
                                    <asp:CheckBox ID="CkExcell" runat="server" Text="Gerar Excel:" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl" style="width: 190px;">
                                    Anexar Foto:
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt">
                                    <asp:HiddenField ID="HiddenUpload" runat="server" />
                                    <input style="float: right; position: relative;" id="filUpload" type="file"
                                        onchange="SelecionaImagem(this);" runat="server" />
                                    &nbsp;
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt lg" style="text-align: center;">
                                    <asp:Image ID="imgFoto" runat="server" Width="200px" Style="border-radius: 5px;" />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Tipos de Clientes" ID="TabPanelTiposDeClientes">
                    <ContentTemplate>
                        <div class="bordagrid" style="height: 367px;">
                            <div class="row">
                                <div class="coltxt" style="line-height: 20px; width: 100%;">
                                    <asp:CheckBoxList ID="chkListTipoDeCliente" runat="server" RepeatColumns="4" AutoPostBack="True"
                                        OnSelectedIndexChanged="chkListTipoDeCliente_SelectedIndexChanged" data-ToolTip="default" ToolTip="selecione um tipo que represente o cliente." />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Contatos" ID="TabPanelContatos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoCONT" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparCONT" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaCONT" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Nome:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeContato" runat="server" MaxLength="50" Width="550px"
                                    data-ToolTip="default" ToolTip="Nome do funcionário." />
                                <asp:HiddenField ID="HgridRowIndexContato" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Função:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFuncaoContato" runat="server" MaxLength="50" Width="317px"
                                    data-ToolTip="default" ToolTip="Função na empresa." />
                            </div>
                            <div class="collbl w113px">
                                Telefone:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTelefoneContato" runat="server" MaxLength="50" Width="100px"
                                    data-ToolTip="default" ToolTip="Contato do funcionário." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlBancoContato" runat="server" Width="327px" />
                            </div>
                            <div class="collbl w113px">
                                Email:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEmailContato" runat="server" MaxLength="50" Width="100px"
                                    data-ToolTip="default" ToolTip="Endereço eletrônico do funcionário." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Agência/DG:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAgenciaContato" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Agência bancária." />
                                <asp:TextBox ID="txtDGAgContato" runat="server" MaxLength="1" Width="37px"
                                    data-ToolTip="default" ToolTip="Agência bancária." />
                            </div>
                            <div class="collbl w113px" style="margin-left: 166px;">
                                Conta/DG:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCtaContato" runat="server" MaxLength="20" Width="100px"
                                    data-ToolTip="default" ToolTip="Conta bancária." />
                                <asp:TextBox ID="txtDGCtaContato" runat="server" MaxLength="2" Width="37px"
                                    data-ToolTip="default" ToolTip="Conta bancária." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Observação:
                            </div>
                            <div class="coltxt " style="width: 55%;">
                                <asp:TextBox ID="txtObservacaoContato" TextMode="MultiLine" runat="server" MaxLength="200"
                                    Width="100%" data-ToolTip="default" ToolTip="Observações importantes." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridContatos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarContato" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarContato_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="NomeContato" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Funcao" HeaderText="Função">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Telefone" HeaderText="Telefone">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeBanco" HeaderText="Banco">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoAgencia" HeaderText="Agência">
                                        <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                        <ItemStyle HorizontalAlign="Left" Width="50px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoAgencia" HeaderText="DG">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoContaCorrente" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                        <ItemStyle HorizontalAlign="Left" Width="50px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoConta" HeaderText="DG">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="email" HeaderText="Email">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirContato" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirContato_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Dependentes" ID="TabPanelDependentes">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoD" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparD" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaD" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nome:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDependente" runat="server" MaxLength="50" Width="358px"
                                    data-ToolTip="default" ToolTip="Nome do dependente." />
                                <asp:HiddenField ID="HgridRowIndexDependente" runat="server" />
                            </div>
                            <div class="collbl">
                                RG:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRGDependente" runat="server" MaxLength="20" Width="100px" CssClass="txtNumerico"
                                    data-ToolTip="default" ToolTip="registro regal do dependente." />
                            </div>
                            <div class="collbl">
                                CPF:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCPFDependente" runat="server" Width="100px" CssClass="txtCpf"
                                    data-ToolTip="default" ToolTip="Número do CPF do dependente." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo Do Dependente:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoDependente" runat="server" Width="368px">
                                    <asp:ListItem></asp:ListItem>
                                    <asp:ListItem Value="Conjuge">Cônjuge</asp:ListItem>
                                    <asp:ListItem Value="Companheiro">Companheiro(a)</asp:ListItem>
                                    <asp:ListItem Value="Filho">Filho(a)/enteado(a)</asp:ListItem>
                                    <asp:ListItem Value="Irmao">Irmão/Irmã</asp:ListItem>
                                    <asp:ListItem Value="Pais">Pai, mãe, padrasto ou madrasta</asp:ListItem>
                                    <asp:ListItem Value="Outros">Pessoa sob guarda ou tutela (até 18 anos)</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                Data De Nascimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNascimentoDependente" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Data de nascimento do dependente." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Profissão:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProfissaoDependente" runat="server" MaxLength="50" Width="358px"
                                    data-ToolTip="default" ToolTip="Profissão do dependente." />
                            </div>
                            <div class="collbl">
                                Custos R$(ANO):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCustoDependente" runat="server" CssClass="txtDecimal" Width="100px"
                                    data-ToolTip="default" ToolTip="Custos com o dependente no ano." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridDependentes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarDependente" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarDependente_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoDeDependente" HeaderText="Dependente">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="RG" HeaderText="RG">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CPF" HeaderText="CPF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataNascimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Nascimento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Profissao" HeaderText="Profissão">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CustoAno" HeaderText="Custo" />
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirDependente" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirDependente_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Contas Bancárias" ID="TabPanelContasBancárias">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoCB" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparBC" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaBC" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Banco:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HgridRowIndexContaBancaria" runat="server" />
                                <asp:DropDownList ID="ddlBancoContaBancaria" runat="server" Width="306px" />
                            </div>
                            <div class="collbl w113px">
                                Agência/DG:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAgenciaContaBancaria" runat="server" Width="120px"
                                    data-ToolTip="default" ToolTip="Agência bancária." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDGAgContaBancaria" runat="server" MaxLength="1" Width="20px"
                                    data-ToolTip="default" ToolTip="Agência bancária." />
                            </div>
                            <div class="collbl w113px">
                                Conta/DG:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCtaContaBancaria" runat="server" MaxLength="20" Width="120px"
                                    data-ToolTip="default" ToolTip="Conta bancária." />
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDGCtaContaBancaria" runat="server" MaxLength="2" Width="20px"
                                    data-ToolTip="default" ToolTip="Conta bancária." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Estado:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEstadoContaBancaria" runat="server" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlEstadoContaBancaria_SelectedIndexChanged" Width="306px" />
                            </div>
                            <div class="collbl w113px">
                                Tipo Conta:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoConta" runat="server" Width="164px">
                                    <asp:ListItem Value="C">C. Corrente</asp:ListItem>
                                    <asp:ListItem Value="P">C. Popança</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                <asp:CheckBox ID="chkAtivo" runat="server" Text="Ativa" Checked="true" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Cidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCidadeContaBancaria" runat="server" Enabled="False" Width="296px"
                                    data-ToolTip="default" ToolTip="Localização do cliente dentro do Estado." />
                            </div>
                            <div class="collbl w113px">
                                Praça:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPracaContaBancaria" runat="server" MaxLength="50" Width="154px"
                                    data-ToolTip="default" ToolTip="Local da onde se encontra a agência." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 80%;">
                                <asp:TextBox ID="txtObservacaoContaBancaria" runat="server" MaxLength="200" Width="100%"
                                    data-ToolTip="default" ToolTip="Observações importantes." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridContasBancarias" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarContaBancaria" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarContaBancaria_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoBanco" HeaderText="Banco">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeBanco" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoAgencia" HeaderText="Agência">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoAgencia" HeaderText="DG">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ContaCorrente" HeaderText="Conta">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DigitoConta" HeaderText="DG">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoConta" HeaderText="Tipo" />
                                    <asp:BoundField DataField="Praca" HeaderText="Praça">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Estado" HeaderText="UF">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirContaBancaria" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirContaBancaria_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" OnClientClick="return confirm('Tem certeza que deseja excluir o registro?');"
                                                Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Cultura" ID="TabPanelCultura">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoCULT" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparCULT" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaCULT" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:HiddenField ID="HgridRowIndexCultura" runat="server" />
                                <asp:DropDownList ID="ddlSafra" runat="server" Width="232px" />
                            </div>
                            <div class="collbl">
                                Cultura:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCultura" runat="server" Width="232px" />
                            </div>
                            <div class="collbl" style="margin-left: 5px;">
                                Área(HA):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtArea" CssClass="txtDecimal" runat="server" data-ToolTip="default"
                                    ToolTip="Tamanho equivalente do equitare." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Produtividade(HA):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtProdutividade" CssClass="txtDecimal" runat="server" data-ToolTip="default" ToolTip="Redimento do equitare." />
                            </div>
                            <div class="collbl" style="margin-left: 122px;">
                                Consumo Própio(HA):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtConsumoProprio" CssClass="txtDecimal" runat="server"
                                    data-ToolTip="default" ToolTip="Despesa própria do equitare." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Comprometido(HA):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtComprometido" CssClass="txtDecimal" runat="server" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl" style="margin-left: 122px;">
                                Estimativa De Entrega(HA):
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEstimativaDeEntrega" CssClass="txtDecimal" runat="server"
                                    data-ToolTip="default" ToolTip="Suposição da quantidade de entrega do equitare." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 84%;">
                                <asp:TextBox ID="txtObservacaoSafra" runat="server" MaxLength="2000" Width="100%" TextMode="MultiLine"
                                    data-ToolTip="default" ToolTip="Observações importantes." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridCultura" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarCultura" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarCultura_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoSafra" HeaderText="Safra">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCultura" HeaderText="Cultura">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeCultura" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AreaPlantada" DataFormatString="{0:N2}" HeaderText="Área">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Produtividade" DataFormatString="{0:N2}" HeaderText="Produtividade">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ConsumoProprio" DataFormatString="{0:N2}" HeaderText="Consumo Próprio">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Comprometido" DataFormatString="{0:N2}" HeaderText="Compromedido">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EstimativaEntrega" DataFormatString="{0:N2}" HeaderText="Estimativa de Entrega">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirCultura" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirCultura_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Veículos" ID="TabPanelVeículos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoV" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparV" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaV" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Placa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPlacaVeiculo" runat="server" MaxLength="20" Width="150px"
                                    data-ToolTip="default" ToolTip="Inserir a placa do veículo." />
                                <asp:HiddenField ID="HgridRowIndexVeiculo" runat="server" />
                            </div>
                            <div class="collbl">
                                Tipo Do Veículo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoVeiculo" runat="server" Width="236px" />
                            </div>
                            <div class="collbl">
                                Ano:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAnoVeiculo" runat="server" MaxLength="4" Width="150px" CssClass="txtNumerico4"
                                    data-ToolTip="default" ToolTip="Ano de fabricação do veículo." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fabricante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFabricanteVeiculo" runat="server" MaxLength="20" Width="150px"
                                    data-ToolTip="default" ToolTip="Nome do fabricante do veículo." />
                            </div>
                            <div class="collbl">
                                Marca/Modelo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtModeloVeiculo" runat="server" MaxLength="50" Width="150px"
                                    data-ToolTip="default" ToolTip="Informar qual a marca ou o modelo do veículo." />
                            </div>
                            <div class="collbl" style="margin-left: 76px;">
                                Data De Avaliação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataAvaliacaoVeiculo" CssClass="calendario" runat="server" Width="150px"
                                    data-ToolTip="default" ToolTip="Data de avaliação." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Oficial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrOficialVeiculo" CssClass="txtDecimal" runat="server" AutoPostBack="True" OnTextChanged="txtVlrOficialVeiculo_TextChanged"
                                    Width="150px" data-ToolTip="default" ToolTip="Valor oficial do veículo." />
                            </div>
                            <div class="collbl">
                                Valor Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrMoedaVeiculo" CssClass="txtDecimal" runat="server" Enabled="False"
                                    Width="150px" data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl" style="margin-left: 76px;">
                                <asp:CheckBox ID="chkVeiculoOnerado" runat="server" Font-Bold="True" Text="Onerado:" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridVeiculos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarVeiculo" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarVeiculo_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" Width="20px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoPlaca" HeaderText="Placa">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoTipoDeVeiculo" HeaderText="Veiculo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ano" HeaderText="Ano">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Fabricante" HeaderText="Fabricante">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MarcaModelo" HeaderText="Marca/Modelo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataAvaliacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Avaliação">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Onerado">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="CheckBox1" runat="server" Checked='<%# eval("Onerado") %>' />
                                        </ItemTemplate>
                                        <FooterStyle HorizontalAlign="Center" />
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirVeiculo" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirVeiculo_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Equipamentos" ID="TabPanelEquipamentos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoE" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparE" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="LinkButton7" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroEquipamento" runat="server" MaxLength="20" Width="150px"
                                    data-ToolTip="default" ToolTip="Número de registro do equipamento." />
                                <asp:HiddenField ID="HgridRowIndexEquipamento" runat="server" />
                            </div>
                            <div class="collbl">
                                Tipo Do Equipamento:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoEquipamento" runat="server" Width="236px" help="Selecione um tipo de equipamento." />
                            </div>
                            <div class="collbl">
                                Ano:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAnoEquipamento" runat="server" MaxLength="4" Width="150px" CssClass="txtNumerico4"
                                    data-ToolTip="default" ToolTip="Ano de fabricação do equipamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fabricante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFabricanteEquipamento" runat="server" MaxLength="20" Width="150px"
                                    data-ToolTip="default" ToolTip="Fabricante do equipamento." />
                            </div>
                            <div class="collbl">
                                Marca:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMarcaEquipamento" runat="server" MaxLength="50" Width="226px"
                                    help="Informe o marca do equipamento." data-ToolTip="default" ToolTip="Marca do equipamento." />
                            </div>
                            <div class="collbl">
                                Modelo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtModeloEquipamento" runat="server" MaxLength="50" Width="150px"
                                    data-ToolTip="default" ToolTip="Modelo do equipamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data De Avaliação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataAvaliacaoEquipamento" CssClass="calendario" runat="server" Width="129px"
                                    data-ToolTip="default" ToolTip="Selecione a data de avaliação." />
                            </div>
                            <div class="collbl">
                                Valor Oficial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrOficialEquipamento" CssClass="txtDecimal" runat="server" AutoPostBack="True" OnTextChanged="txtVlrOficialEquipamento_TextChanged"
                                    Width="150px" data-ToolTip="default" ToolTip="Informe o valor oficial do equipamento." />
                            </div>
                            <div class="collbl" style="margin-left: 76px;">
                                Valor Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrMoedaEquipamento" CssClass="txtDecimal" runat="server" Enabled="False"
                                    Width="150px" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                <asp:CheckBox ID="chkEquipamentoOnerado" runat="server" Font-Bold="True" Text="Onerado:" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridEquipamentos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarEquipamento" runat="server" ImageAlign="AbsMiddle"
                                                ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" OnClick="imgConsultarEquipamento_Click"
                                                CssClass="imgconsultar" data-ToolTip="default" ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Codigo" HeaderText="Registro">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoTipoDeEquipamento" HeaderText="Equipamento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Ano" HeaderText="Ano">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Fabricante" HeaderText="Fabricante">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Marca" HeaderText="Marca">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Modelo" HeaderText="Modelo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataAvaliacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Avaliação">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Onerado">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="CheckBox2" runat="server" Checked='<%# eval("Onerado") %>' />
                                        </ItemTemplate>
                                        <FooterStyle HorizontalAlign="Center" />
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirEquipamento" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirEquipamento_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Imóveis" ID="TabPanelImóveis">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoI" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparI" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaI" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Imóvel:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoImovel" runat="server" Width="122px"
                                    data-ToolTip="default" ToolTip="Informe o imóvel." />
                                <asp:HiddenField ID="HgridRowIndexImovel" runat="server" />
                            </div>
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoImovel" runat="server" MaxLength="50" Width="300px"
                                    data-ToolTip="default" ToolTip="Descrição detalhada do imóvel." />
                            </div>
                            <div class="collbl">
                                <asp:CheckBox ID="chkOneradoImovel" runat="server" Text="Onerado:" data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Estado:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEstadoImovel" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEstadoImovel_SelectedIndexChanged"
                                    Width="132px" />
                            </div>
                            <div class="collbl">
                                Cidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCidadeImovel" runat="server" Enabled="False" Width="300px"
                                    data-ToolTip="default" ToolTip="Localização do cliente dentro do Estado." />
                            </div>
                            <div class="collbl">
                                Área Total:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAreaTotalImovel" runat="server" CssClass="txtDecimal" data-ToolTip="default" ToolTip="Área total do bem." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Área Construída:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAreaConstruidaImovel" runat="server" CssClass="txtDecimal" Width="122px"
                                    data-ToolTip="default" ToolTip="Informar a metragem da área construída." />
                            </div>
                            <div class="collbl">
                                Unidade:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlUnidadeImovel" runat="server" Width="310px" />
                            </div>
                            <div class="collbl">
                                Número Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroImovel" runat="server" MaxLength="20" CssClass="txtNumerico"
                                    data-ToolTip="default" ToolTip="Registro numérico do imóvel." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cartório:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCartorioImovel" runat="server" MaxLength="50" Width="598px"
                                    data-ToolTip="default" ToolTip="Informar o nome do cartório de registro." />
                            </div>
                            <div class="collbl">
                                Data De Avaliação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataAvaliacaoImovel" runat="server" CssClass="calendario" Width="100px"
                                    data-ToolTip="default" ToolTip="Selecione a data de avaliação do imóvel." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrOficialImovel" runat="server" AutoPostBack="True" CssClass="txtDecimal"
                                    OnTextChanged="txtVlrOficialImovel_TextChanged" Width="122px" data-ToolTip="default" ToolTip="Valor da moeda no dia." />
                            </div>
                            <div class="collbl">
                                Valor Oficial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrMoedaImovel" runat="server" CssClass="txtDecimal" Enabled="False" data-ToolTip="default" ToolTip="Informe o valor oficial do imóveis." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 81%;">
                                <asp:TextBox ID="txtObservacaoImovel" runat="server" MaxLength="2000" Width="100%" TextMode="MultiLine"
                                    data-ToolTip="default" ToolTip="Observações importantes." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridImoveis" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarImovel" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarImovel_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoImovel" HeaderText="Imóvel" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoCidade" HeaderText="Cidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEstado" HeaderText="UF">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AreaTotal" DataFormatString="{0:N2}" HeaderText="Área Total">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="AreaConstruida" DataFormatString="{0:N2}" HeaderText="Área Construída">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" DataFormatString="{0:N2}" HeaderText="Valor Oficial">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorMoeda" DataFormatString="{0:N2}" HeaderText="Valor Moeda">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirImovel" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirImovel_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Arrendantes" ID="TabPanelArrendantes">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoA" Text="Adicionar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="LinkButton8" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridArrendantes" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="100%" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="CodigoArrendante" HeaderText="Arrendante">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndArrendante" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left" Width="30px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeArrendante" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataContrato" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Contrato">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataVencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Matricula" HeaderText="Matr&#237;cula">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Area" HeaderText="Area">
                                        <FooterStyle HorizontalAlign="Right" />
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CustoArrendante" HeaderText="Custo Scs/ha">
                                        <FooterStyle HorizontalAlign="Right" />
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirArrendante" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                data-ToolTip="default" ToolTip="Excluir Registro" Width="18px" Style="cursor: pointer"
                                                OnClick="imgExcluirArrendante_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px"></HeaderStyle>
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Representantes" ID="TabPanelRepresentantes">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoR" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparR" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaR" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Representante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRepresentante" runat="server" Width="450px" Enabled="False" />
                                <asp:HiddenField ID="txtCodigoRepresentante" runat="server" />
                                <asp:HiddenField ID="HgridRowIndexRepresentante" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button CssClass="btn" ID="btnRepresentante" runat="server" Text=">" UseSubmitBehavior="False"
                                    OnClick="btnRepresentante_Click" data-ToolTip="default" ToolTip="Selecionar o representante quando houver." />
                            </div>
                            <div class="collbl w113px">
                                <asp:CheckBox ID="chkPrincipalRepresentante" runat="server" />
                                Principal:
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridRepresentantes" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarRepresentante" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarRepresentante_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoRepresentante" HeaderText="Representante">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndRepresentante" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeRepresentante" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Principal">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="chkRepresentante" runat="server" Checked='<%# eval("Principal") %>'
                                                Enabled="False" />
                                        </ItemTemplate>
                                        <HeaderStyle HorizontalAlign="Center" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:TemplateField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirRepresentante" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirRepresentante_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Sócios" ID="TabPanelSócios">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoS" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparS" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaS" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Sócio:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSocio" runat="server" Width="450px" Enabled="False" />
                                <asp:HiddenField ID="txtCodigoSocio" runat="server" />
                                <asp:HiddenField ID="HgridRowIndexSocio" runat="server" />
                            </div>
                            <div class="coltxt ">
                                <asp:Button ID="btnSocio" CssClass="btn" runat="server" Text=">" UseSubmitBehavior="False"
                                    OnClick="btnSocio_Click" data-ToolTip="default" ToolTip="Informar o nome do sócio." />
                            </div>
                            <div class="collbl w113px">
                                %Participação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtParticipacaoSocio" runat="server" data-ToolTip="default"
                                    ToolTip="Informar o percentual de participação na empresa." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridSocios" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarSocio" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarSocio_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoSocio" HeaderText="Socio">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndSocio" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeSocio" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Participacao" HeaderText="Participacao">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirSocio" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirSocio_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Matrículas" ID="TabPanelMatrículas">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoM" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparM" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaM" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Matrícula:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumeroMatricula" runat="server" MaxLength="20" data-ToolTip="default" ToolTip="Informe a matrícula." />
                                <asp:HiddenField ID="HgridRowIndexMatricula" runat="server" />
                            </div>
                            <div class="collbl">
                                Data De Avaliação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataAvaliacaoMatricula" CssClass="calendario" runat="server"
                                    Width="136px" data-ToolTip="default" ToolTip="Selecione a data da matrícula." />
                            </div>
                            <div class="collbl">
                                Área:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAreaMatricula" CssClass="txtDecimal" runat="server"
                                    data-ToolTip="default" ToolTip="Informe a área da matrícula." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Oficial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVlrOficialMatricula" CssClass="txtDecimal" runat="server" AutoPostBack="True"
                                    data-ToolTip="default" ToolTip="Informe o valor oficial da matrícula." />
                            </div>
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroMatricula" runat="server" MaxLength="20" Width="157px"
                                    data-ToolTip="default" ToolTip="Informe o registro da matrícula." />
                            </div>
                            <div class="collbl">
                                Livro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtLivroMatricula" runat="server" MaxLength="20"
                                    data-ToolTip="default" ToolTip="Informe o número do livro da matrícula." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Folha:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFolhaMatricula" runat="server" MaxLength="10" CssClass="txtNumerico"
                                    data-ToolTip="default" ToolTip="Informe o número da folha da matrícula." />
                            </div>
                            <div class="collbl">
                                Estado:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlEstadoMatricula" runat="server" AutoPostBack="True" OnSelectedIndexChanged="ddlEstadoMatricula_SelectedIndexChanged"
                                    Width="167px" />
                            </div>
                            <div class="collbl">
                                Município:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMunicipioMatricula" runat="server" AutoPostBack="True" Enabled="False"
                                    Width="300px" data-ToolTip="default" ToolTip="Informe o município da matrícula." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridMatriculas" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarMatricula" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                OnClick="imgConsultarMatricula_Click" CssClass="imgconsultar" data-ToolTip="default"
                                                ToolTip="Consultar Registro" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoMatricula" HeaderText="Matrícula">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro" />
                                    <asp:BoundField DataField="Livro" HeaderText="Livro" />
                                    <asp:BoundField DataField="Folha" HeaderText="Folha" />
                                    <asp:BoundField DataField="CodigoMunicipio" HeaderText="Municipio" />
                                    <asp:BoundField DataField="CodigoEstado" HeaderText="UF" />
                                    <asp:BoundField DataField="Area" HeaderText="Área">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataAvaliacao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Data Avaliação">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirMatricula" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirMatricula_Click" Style="cursor: pointer" data-ToolTip="default"
                                                ToolTip="Excluir Registro" Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Financiamentos" ID="TabPanelFinanciamentos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoF" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparF" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaF" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Financiador:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFinanciador" runat="server" Width="394px" Enabled="False" />
                                <asp:HiddenField ID="txtCodigoFinanciador" runat="server" />
                                <asp:HiddenField ID="HgridRowIndexFinanciamento" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="btnFinanciador" runat="server" Text=">" OnClick="btnFinanciador_Click" CssClass="btn"
                                    UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o financiador" />
                            </div>
                            <div class="collbl w113px">
                                Código:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCodigoFinanciamento" runat="server" Width="127px"
                                    data-ToolTip="default" ToolTip="Informe o código do Financiamento." />
                            </div>
                            <div class="collbl w113px">
                                Tipo Financ:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTipoFinanciamento" runat="server" Width="135px" MaxLength="20"
                                    data-ToolTip="default" ToolTip="Informe o tipo do financiamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Safra:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlSafraFinanciamento" runat="server" Width="434px" />
                            </div>
                            <div class="collbl w113px">
                                Grupo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlGrupoProdutoFinanciamento" runat="server" Width="412px" AutoPostBack="True"
                                    OnSelectedIndexChanged="ddlGrupoProdutoFinanciamento_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Produto:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlProdutoFinanciamento" runat="server" Width="434px" />
                            </div>
                            <div class="collbl w113px">
                                Núm. Parcelas:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtParcelaFinanciamento" CssClass="txtNumerico" runat="server" Width="150px"
                                    data-ToolTip="default" ToolTip="Informar o número de parcelas do financiamento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Data Do Financ:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataFinanciamento" CssClass="calendario" runat="server" Width="86px"
                                    data-ToolTip="default" ToolTip="Data que o produto foi adquirido." />
                            </div>
                            <div class="collbl w113px">
                                Vencimento:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtVencimentoFinanciamento" CssClass="calendario" runat="server"
                                    Width="86px" data-ToolTip="default" ToolTip="Vencimento do financiamento." />
                            </div>
                            <div class="collbl w113px" style="margin-left: 72px;">
                                Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlMoedaFinanciamento" runat="server" Width="412px" OnSelectedIndexChanged="ddlMoedaFinanciamento_SelectedIndexChanged"
                                    AutoPostBack="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Quantidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQuantidadeFinanciamento" CssClass="txtDecimal4" runat="server"
                                    Width="107px" data-ToolTip="default" ToolTip="Quantidade comprada." />
                            </div>
                            <div class="collbl w113px">
                                Valor Oficial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtOficialFinanciamento" CssClass="txtDecimal" runat="server" OnTextChanged="txtOficialFinanciamento_TextChanged"
                                    AutoPostBack="True" Enabled="False" Width="107px" data-ToolTip="default" ToolTip="Informe o valor oficial do financiamento." />
                            </div>
                            <div class="collbl w113px" style="margin-left: 72px;">
                                Valor Moeda:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMoedaFinanciamento" CssClass="txtDecimal" runat="server" Enabled="False"
                                    OnTextChanged="txtMoedaFinanciamento_TextChanged" AutoPostBack="True" Width="107px"
                                    data-ToolTip="default" ToolTip="Valor da moeda no dia." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl w113px">
                                Observação:
                            </div>
                            <div class="coltxt" style="width: 88%;">
                                <asp:TextBox ID="txtObservacaoFinanciamento" runat="server" Width="100%" MaxLength="2000"
                                    TextMode="MultiLine" data-ToolTip="default" ToolTip="Observações importantes." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridFinanciamentos" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarFinanciamento" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                data-ToolTip="default" ToolTip="Consultar Registro" Width="20px" OnClick="imgConsultarFinanciamento_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="CodigoFinanciador" HeaderText="Financiador">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="EndFinanciador" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeFinanciador" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoFinanciamento" HeaderText="Codigo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoFinanciamento" HeaderText="Tipo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoSafra" HeaderText="Safra">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeProduto" HeaderText="Produto">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NumeroDeParcelas" HeaderText="Parcelas">
                                        <HeaderStyle HorizontalAlign="Left" Width="50px" />
                                        <ItemStyle HorizontalAlign="Left" Width="50px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataFinanciamento" HeaderText="Financiamento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DataVencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorOficial" HeaderText="Valor">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirFinanciamento" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                data-ToolTip="default" ToolTip="Excluir Registro" Width="18px" Style="cursor: pointer"
                                                OnClick="imgExcluirFinanciamento_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Receitas/Despesas" ID="TabPanelReceitasDespesas">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoRD" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparRD" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaRD" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Ano:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlAno" runat="server" Width="80px">
                                    <asp:ListItem>2012</asp:ListItem>
                                    <asp:ListItem>2013</asp:ListItem>
                                    <asp:ListItem>2014</asp:ListItem>
                                    <asp:ListItem>2015</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                            <div class="collbl">
                                Receita/Despesa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlReceitaDespesa" runat="server" Width="119px">
                                    <asp:ListItem Value="R">RECEITA</asp:ListItem>
                                    <asp:ListItem Value="D">DESPESA</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tipo Receita/Despesa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTipoReceitaDespesa" runat="server" Width="341px">
                                    <asp:ListItem Value="ADVOGADO">Advogado</asp:ListItem>
                                    <asp:ListItem Value="ALUGUEL">Aluguel</asp:ListItem>
                                    <asp:ListItem Value="APOSENTADORIA">Aposentadoria</asp:ListItem>
                                    <asp:ListItem Value="APLICACOES">Aplicacoes</asp:ListItem>
                                    <asp:ListItem Value="CONTADOR">Contador</asp:ListItem>
                                    <asp:ListItem Value="EDUCACAO">Educacao</asp:ListItem>
                                    <asp:ListItem Value="PECUARIA">Pecuaria</asp:ListItem>
                                    <asp:ListItem Value="PENSAO">Pensao</asp:ListItem>
                                    <asp:ListItem Value="SAUDE">Saude</asp:ListItem>
                                    <asp:ListItem Value="DESPESAS">Outras Despesas</asp:ListItem>
                                    <asp:ListItem Value="RECEITAS">Outras Receitas</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Valor Ano:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtValorAnoRecDes" runat="server" Width="94px" CssClass="txtDecimal"
                                    data-ToolTip="default" ToolTip="Valor da despesa/receita no ano." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoRecDes" runat="server" MaxLength="100" Width="713px"
                                    data-ToolTip="default" ToolTip="Descrição da despesa/receita." />
                                <asp:HiddenField ID="HgridIndexRecDes" runat="server" Value="-1" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridReceitasDespesas" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" Width="98%" AutoGenerateColumns="False">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgConsultarReceitaDespesa" runat="server" ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png"
                                                data-ToolTip="default" ToolTip="Consultar Registro" CssClass="imgconsultar" OnClick="imgConsultarReceitaDespesa_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="Ano" HeaderText="Ano">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ReceitaDespesa" HeaderText="R/D">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="TipoReceitaDespesa" HeaderText="Tipo Rec. / Desp.">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorAno" HeaderText="Valor Liq. Anual">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                                        <HeaderStyle HorizontalAlign="Right" />
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirReceitaDespesa" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                data-ToolTip="default" ToolTip="Excluir Registro" Width="18px" Style="cursor: pointer"
                                                OnClick="imgExcluirReceitaDespesa_Click" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanelCentrosDeCustos" runat="server" HeaderText="Centros De Custos">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoCC" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparCC" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Centro de Custo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCentroDeCusto" runat="server" Width="417px" />
                            </div>
                            <div class="collbl">
                                % Responsavel:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPercResponsavelCC" runat="server" CssClass="txtDecimal" Width="67px"
                                    data-ToolTip="default" ToolTip="Informar o percentual do centro de custo do responsável." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridCC" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="CodigoCentroDeCusto" HeaderText="CC">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoCentroCusto" HeaderText="Descricao">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="PercentualFixo" HeaderText="% Fixo">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirCC" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirCC_Click" Style="cursor: pointer" data-ToolTip="default" ToolTip="Excluir Centro De Custo"
                                                Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabSubstitutoTributario" runat="server" HeaderText="Substituto Tributário" Visible="false">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoST" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparST" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaST" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row" runat="server" style="margin-left: 1px;">
                            <div class="collbl">
                                Substituto Tributário:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtSubTributario" runat="server" Style="width: 200px;" ToolTip="Substituto Tributário (IE)." />
                            </div>
                            <div class="collbl">
                                Estado:
                            </div>
                            <div class="coltxt">
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlEstadoST" runat="server" Width="265px" ToolTip="Estado do substituto Tributário (IE)." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft bordagrid" style="width: 100%; min-height: 225px;">
                            <asp:GridView ID="gridST" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                                GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Estado_Id" HeaderText="Estado">
                                        <HeaderStyle HorizontalAlign="Left" Width="300px" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="IESubstitutoTributario" HeaderText="Sub.Tributario">
                                        <HeaderStyle HorizontalAlign="Left" Width="300px" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirST" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirST_Click" Style="cursor: pointer" data-ToolTip="default" ToolTip="Excluir substituto Tributário (IE)."
                                                Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Documentos" ID="TabDocumentos">
                    <HeaderTemplate>
                        Documentos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Descrição:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDescricaoDocumento" runat="server" MaxLength="100" Width="500px" data-ToolTip="default"
                                    ToolTip="Descrição do Documento." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Arquivo:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDoArquivo" runat="server" Enabled="false" Style="width: 200px;" />
                                <asp:FileUpload ID="fupArquivo" OnChange="Arquivo()" runat="server" Width="120px" Font-Size="11px" ClientIDMode="Static" />
                            </div>
                        </div>
                        <div class="painelleft bordagrid" style="width: 99%; margin-left: 0.5%; height: 225px;">
                            <asp:GridView ID="gridDocumentos" runat="server" AutoGenerateColumns="False" ForeColor="#333333"
                                GridLines="None" Width="100%">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:BoundField DataField="Codigo" HeaderText="Código">
                                        <HeaderStyle HorizontalAlign="Center" Width="70px" />
                                        <ItemStyle HorizontalAlign="Center" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Descricao" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left" Width="500px" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="NomeDoArquivo" HeaderText="Nome do arquivo">
                                        <HeaderStyle HorizontalAlign="Left" Width="300px" />
                                        <ItemStyle HorizontalAlign="Left" Width="300px" />
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText="Arquivo">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgDownload" runat="server" ImageAlign="AbsMiddle" ImageUrl="~/Images/download32x32.png" Style="margin-top: 0;"
                                                Height="18px" Width="18px" OnClick="imgDownload_Click" ToolTip="Baixar Arquivo" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                    <asp:TemplateField HeaderText="Excluir">
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirDocumento" runat="server" ImageUrl="~/Images/deletar.gif" Style="border: 0;"
                                                OnClick="imgExcluirDocumento_Click" data-ToolTip="default" ToolTip="Excluir"
                                                OnClientClick="return confirm('Deseja realmente excluir o Documento?');" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                        <ItemStyle Width="25px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                    </asp:TemplateField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <asp:UpdatePanel ID="updpnlAdicionar" runat="server">
                            <ContentTemplate>
                                <asp:Button ID="btnAdicionar" runat="server" Text="Adicionar" ClientIDMode="Static"
                                    OnClick="btnAdicionar_Click" CssClass="none" />
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="btnAdicionar" />
                            </Triggers>
                        </asp:UpdatePanel>
                        <asp:Label ID="lblMsg" runat="server" />
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabTabelaPrecos" runat="server" HeaderText="Tabela de preços" Visible="false">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server">
                                        <asp:LinkButton class="iconNovo" ID="lnkNovoTP" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconLimpar" ID="lnkLimparTP" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server">
                                        <asp:LinkButton class="iconAjuda" ID="lnkAjudaTP" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row" runat="server">
                            <div class="collbl">
                                Representante:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRepresentanteTP" runat="server" MaxLength="100" Width="500px" data-ToolTip="default"
                                    ToolTip="Representante da tabela de preços." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Tabelas:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlTabelasDePrecos" runat="server" AutoPostBack="True" Width="417px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridTP" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:BoundField DataField="CodigoTabela" HeaderText="Tabela">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="DescricaoTipo" HeaderText="Descrição">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:TemplateField>
                                        <ItemTemplate>
                                            <asp:ImageButton ID="imgExcluirTP" runat="server" Height="18px" ImageUrl="~/images/erro.jpg"
                                                OnClick="imgExcluirTP_Click" Style="cursor: pointer" data-ToolTip="default" ToolTip="Excluir vínculo com tabela de preços"
                                                Width="18px" />
                                        </ItemTemplate>
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:TemplateField>
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="Consulta" ID="TabPanelConsulta">
                    <ContentTemplate>
                        <div class="bordagrid">
                            <asp:GridView ID="gridConsultaCliente" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="gridConsultaCliente_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True">
                                        <HeaderStyle Width="30px" />
                                        <ItemStyle Width="30px" />
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Codigo" HeaderText="CPF/CNPJ">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEndereco" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Endereco" HeaderText="Endereço">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Numero" HeaderText="Número">
                                        <HeaderStyle HorizontalAlign="Left" Width="40px" />
                                        <ItemStyle HorizontalAlign="Left" Width="40px" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Complemento" HeaderText="Complemento">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                                        <HeaderStyle HorizontalAlign="Left" />
                                        <ItemStyle HorizontalAlign="Left" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="CodigoEstado" HeaderText="UF">
                                        <HeaderStyle HorizontalAlign="Left" Width="30px" />
                                        <ItemStyle HorizontalAlign="Left" Width="30px" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                        <div class="row" style="text-align: center;">
                            <div class="coltxt">
                                <asp:Button ID="btnImpressaoClientes" runat="server" OnClick="btnImpressaoClientes_Click"
                                    CssClass="botao" Text="&lt; &lt; &lt;  Impressão  &gt; &gt; &gt;" UseSubmitBehavior="False"
                                    Visible="False" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:Arrendante ID="ucArrendante" runat="server" />
    <uc:ConsultaCep ID="ucConsultaCep" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
    <uc:ConsultaEstados ID="ucConsultaEstados" runat="server" />
    <uc:ConsultaCadastro ID="ucConsultaCadastro" runat="server" />
</asp:Content>
