<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Empresas.aspx.vb" Inherits="NGS.Web.UI.Empresas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1280px !important;
        }

        .collbl {
            width: 155px;
        }

        .w133px {
            width: 133px;
        }

        .painelleft {
            width: 32.9%;
            margin-right: 4px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmEmpresas" runat="server" />
    <orea:ajaxupdating id="ajaxUpdating" runat="server" text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEmpresas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro de Empresas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                        <li runat="server" style="float: right;">
                            <div class="row" style="margin-top: 0;">
                                <div class="coltxt">
                                    <asp:Image ID="Image1" runat="server" Height="20px" ImageAlign="AbsMiddle" ImageUrl="~/Images/man2.png"
                                        data-ToolTip="default" ToolTip="Usuário Lançamento" Width="18px" />
                                </div>
                                <div class="coltxt">
                                    <asp:DropDownList ID="ddlUsuarios" runat="server" Width="175px">
                                    </asp:DropDownList>
                                </div>
                            </div>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="673px" AutoPostBack="True"
                        OnSelectedIndexChanged="DdlEmpresa_SelectedIndexChanged" />
                </div>
            </div>
            <ajaxtoolkit:tabcontainer id="TabContainer1" runat="server" activetabindex="0" width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server">
                    <HeaderTemplate>
                        Cadastro
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="row">
                            <div class="collbl">
                                Parametros
                            </div>
                        </div>
                        <div class="row">
                            <div class="painelleft" style="width: 15%;">
                                <asp:CheckBox ID="chkNossaEmissao" runat="server" Text="Nossa Emissão"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chknfe" runat="server" Text="Emite NF-e"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chknfProdutor" runat="server" Text="Obriga NF do Produtor"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkObrigaNavio" runat="server" Text="Obriga Navio"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                            </div>
                            <div class="painelleft" style="width: 24%;">
                                <asp:CheckBox ID="chkFluxoDeCaixa" runat="server" Text="Participa Do Fluxo De Caixa"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkRegistroDeExportacao" runat="server" Text="Obrigar Registro de Exportação"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkBaixaFinanceiroPorLote" runat="server" Text="Baixa Financeiro por Lote"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkUsarDescricaoProduto" runat="server" Text="Acrescentar Descrição ao Nome do Produto"
                                    data-ToolTip="default" ToolTip="Acrescentar Descrição ao Nome do Produto." />
                            </div>
                            <div class="painelleft" style="width: 34%;">
                                <asp:CheckBox ID="chkSugereDeposito" runat="server" Text="Sugerir Deposito igual a Empresa"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkCertidaoNegativa" runat="server" Text="Validar com Certidão Negativa"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkObrigaChaveNf" runat="server" Text="Obriga Chave Eletrônica para NFe do Fornecedor"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkUsarRegistroMinAgr" runat="server" Text="Acrescentar Registro Ministério da Agricultura ao nome do Produto"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                            </div>
                            <div class="painelleft" style="width: 24%;">
                                <asp:CheckBox ID="chkFretePedido" runat="server" Text="Frete no Pedido"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkPedidoBloqueado" runat="server" Text="Pedido Bloqueado"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkLiberaCarregamento" runat="server" Text="Controla Carregamento"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                <br />
                                <asp:CheckBox ID="chkObrigaChaveNfg" runat="server" Text="Obriga Chave Eletrônica em Notas Fiscais Gerais para NFe e CTe"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                            </div>
                        </div>
                                 <div class="row" runat="server">
                                                                     <asp:CheckBox ID="chkEmitirCTe" runat="server" Text="Emitir CT-e"
                                    data-ToolTip="default" ToolTip="Selecionar os parâmetros para cada empresa." />
                                 </div>


                        <div class="row" runat="server">
                            <asp:CheckBox ID="chkObgEncargo" runat="server" Text="" data-ToolTip="default" ToolTip="" />
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Servidor:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlServidor" runat="server" Width="673px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Prazo cancelamento NFe:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPrazoCancelamentoNFE" runat="server" CssClass="txtNumerico" Width="40px"
                                    data-ToolTip="default" ToolTip="Informar o prazo em horas para cancelamento de nota fiscal eletrônica." />
                                <lable style="margin-left: 4px;">Horas</lable>
                            </div>
                            <div class="collbl" style="margin-left: 29px;">
                                Número de vias NFe:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumeroDeViasNFE" runat="server" Width="40px" CssClass="txtNumerico"
                                    data-ToolTip="default" ToolTip="Informar o numero de vias de nota fiscal eletrônica." />
                            </div>
                            <div class="collbl" style="margin-left: 53px;">
                                Controla Emissao Cheque:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RadControlaEmissaChequeNao" runat="server" Text="Não" GroupName="ControlaEmissaCheque"
                                    data-ToolTip="default" ToolTip="Informar se irá controlar emissão de cheques." />
                                <asp:RadioButton ID="RadControlaEmissaChequeSim" runat="server" Text="Sim" GroupName="ControlaEmissaCheque"
                                    data-ToolTip="default" ToolTip="Informar se irá controlar emissão de cheques." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Controla Data Movim. NFG:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RadControlaDataMOvimentoNFGNao" runat="server" Text="Não" GroupName="ControlaDataMovimentoNFG"
                                    data-ToolTip="default" ToolTip="Informar se irá controlar data de movimentação de NFG." />
                                <asp:RadioButton ID="RadControlaDataMOvimentoNFGSim" runat="server" Text="Sim" GroupName="ControlaDataMovimentoNFG"
                                    data-ToolTip="default" ToolTip="Informar se irá controlar data de movimentação de NFG." />
                            </div>
                            <div class="collbl" style="margin-left: 22px;">
                                Obriga Arquivo NFe:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="radArquivoNFENao" runat="server" Text="Não" GroupName="ArquivoNFE"
                                    data-ToolTip="default" ToolTip="Informar se é obrigatório o arquivo de nota fiscal eletrônica." />
                                <asp:RadioButton ID="radArquivoNFESim" runat="server" Text="Sim" GroupName="ArquivoNFE"
                                    data-ToolTip="default" ToolTip="Informar se é obrigatório o arquivo de nota fiscal eletrônica." />
                            </div>
                            <div class="collbl">
                                Obriga Conferência NFe:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="radConferenciaNFENao" runat="server" Text="Não" GroupName="ConferenciaNFE"
                                    data-ToolTip="default" ToolTip="Informar se é obrigatória a conferência de nota fiscal eletrônica." />
                                <asp:RadioButton ID="radConferenciaNFESim" runat="server" Text="Sim" GroupName="ConferenciaNFE"
                                    data-ToolTip="default" ToolTip="Informar se é obrigatória a conferência de nota fiscal eletrônica." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Observação Sefaz NFE:
                            </div>
                            <div class="coltxt" style="width: 70%;">
                                <asp:TextBox ID="txtObservacaoSefazNFE" runat="server" Style="width: 100%;" TextMode="MultiLine"
                                    data-ToolTip="default" ToolTip="Inserir informações da sefaz quando houver." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Limite de Dias NF Retroativa:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDiasNFRetroativa" class="txtNumerico" runat="server" Width="100"
                                    data-ToolTip="default" ToolTip="Informar o limite de Nota Fiscal retroativa." />
                            </div>
                            <div class="collbl">
                                Download NFe:
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblDownloadNFe" runat="server" Font-Bold="true" Width="250px" />&nbsp
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">Data Financeiro </div>
                            <asp:TextBox ID="txtDataFinanceiro" CssClass="calendario" runat="server" Width="100px"
                                data-ToolTip="default" ToolTip="" />
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Registro na Junta
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Número:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroNaJunta" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Número do registro na Junta Comercial" />
                            </div>
                            <div class="collbl">
                                Fundacao:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataDaFundacao" CssClass="calendario" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Estado:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEstadoDaJunta" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Estado da Junta Comercial." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                NIRE:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNire" runat="server" Width="100px" data-ToolTip="default"
                                    ToolTip="Numero de Inscrição do Registro de Empresas fornecido pela Junta Comercial." />
                            </div>
                            <div class="collbl">
                                Data da Nire:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataDaNire" CssClass="calendario" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Data de criação do NIRE." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Dados da Empresa
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cód.Regime Tributário:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCrt" runat="server" Width="673px">
                                    <asp:ListItem Value="1">Simples Nacional</asp:ListItem>
                                    <asp:ListItem Value="2">Simples Nacional - excesso de sublimite de receita bruta</asp:ListItem>
                                    <asp:ListItem Value="3">Regime Normal</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Relacionamento:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCodigoDeRelacionamento" runat="server" Width="673px">
                                    <asp:ListItem Value="10">10 - Vinculadas (Art 23 da Lei 9.430/96) Exceto as que se enquadrem nos tipos precedentes</asp:ListItem>
                                    <asp:ListItem Value="11">11 - Localizadas em Pa&#237;s com tributa&#231;&#227;o favborecida (Art 24 da Lei 9.430/96)</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Esc. Contábil:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEscrituracaoContabil" runat="server" Width="673px">
                                    <asp:ListItem Value="G">G - Livro Di&#225;rio (Completo, sem escritura&#231;&#227;o auxiliar)</asp:ListItem>
                                    <asp:ListItem Value="R">R - Livro Diario Com Escritura&#231;&#227;o Resumida (Com Escritura&#231;&#227;o Auxiliar)</asp:ListItem>
                                    <asp:ListItem Value="A">A - Livro Di&#225;rio Auxiliar ao Di&#225;rio Com Escritura&#231;&#227;o Resumida</asp:ListItem>
                                    <asp:ListItem Value="B">B - Livro Balancetes Di&#225;rios e Balan&#231;os</asp:ListItem>
                                    <asp:ListItem Value="Z">Z - Razao Auxiliar (Livro Contabil Auxiliar conforme leioute definido nor registros I500 a I555)</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Natureza Juridica:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlNaturezaJuridica" runat="server" Width="673px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CNAE:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlCnae" runat="server" Width="673px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Insc. Municipal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtInscricaoMunicipal" runat="server" CssClass="txtNumerico" MaxLength="50" Width="100"
                                    data-ToolTip="default" ToolTip="Informar a inscrição Municipal." />
                            </div>
                            <div class="collbl">
                                At. Ec. Estadual:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAtividadeEconomicaEstadual" runat="server" CssClass="txtNumerico" MaxLength="10" Width="100"
                                    data-ToolTip="default" ToolTip="Informar o código da Atividade Econômica Estadual" />
                            </div>
                            <div class="collbl">
                                At. Ec. Federal:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAtividadeEconomicaFederal" runat="server" CssClass="txtNumerico" MaxLength="10" Width="100"
                                    data-ToolTip="default" ToolTip="Informar o código da Atividade Econômica Federal" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Matriz:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMatriz" runat="server" Width="100" MaxLength="1"
                                    data-ToolTip="default" ToolTip="Caso seja Matriz informar S senão informar N" />
                            </div>
                            <div class="collbl">
                                Registro EC:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroEC" runat="server" Width="100"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Registro EP:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroEP" runat="server" Width="100"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Registro EI:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistroEI" runat="server" Width="100"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                            <div class="collbl">
                                Marca:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtMarca" runat="server" Width="100"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Dados do Titular
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nome:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDoTitular" runat="server" Width="662px"
                                    data-ToolTip="default" ToolTip="Inserir o nome do contador." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CPF:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCpfDoTitular" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Informar o CPF do contador." />
                            </div>
                            <div class="collbl">
                                Qualificação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQualificacaoDoTitular" runat="server" Width="381px"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cd. Qualificação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCodigoDaQualificacaoDoTitular" runat="server" Width="673px">
                                    <asp:ListItem Value=""></asp:ListItem>
                                    <asp:ListItem Value="203">203 - Diretor</asp:ListItem>
                                    <asp:ListItem Value="204">204 - Conselheiro de Administra&#231;&#227;o</asp:ListItem>
                                    <asp:ListItem Value="205">205 - Administrador</asp:ListItem>
                                    <asp:ListItem Value="206">206 - Administrador de Grupo</asp:ListItem>
                                    <asp:ListItem Value="207">207 - Administrador de Sociedade Filiada</asp:ListItem>
                                    <asp:ListItem Value="220">220 - Administrador Judicial</asp:ListItem>
                                    <asp:ListItem Value="222">222 - Administrador Judicial - Pessoa Juridica - Profissional Respons&#225;vel</asp:ListItem>
                                    <asp:ListItem Value="223">223 Administrador Judicial/Gestor</asp:ListItem>
                                    <asp:ListItem Value="226">226 Gestor Judicial</asp:ListItem>
                                    <asp:ListItem Value="309">309 - Procurador</asp:ListItem>
                                    <asp:ListItem Value="312">312 - Inventariante</asp:ListItem>
                                    <asp:ListItem Value="313">313 - Liquidante</asp:ListItem>
                                    <asp:ListItem Value="315">315 - Interventor</asp:ListItem>
                                    <asp:ListItem Value="801">801 - Empres&#225;rio</asp:ListItem>
                                    <asp:ListItem Value="900">900 - Contador</asp:ListItem>
                                    <asp:ListItem Value="999">999 - Outros</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                        <div class="row">
                            <div class="subtitulodiv">
                                Dados do Contador
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nome:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeDoContador" runat="server" Width="662px"
                                    data-ToolTip="default" ToolTip="Inserir o nome do contador." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CPF:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCpfDoContador" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Informar o CPF do contador." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CNPJ:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCnpjDoContador" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="Informar o CNPJ do contador." />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                CRC:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCrcDoContador" runat="server" Width="100px" data-ToolTip="default"
                                    ToolTip="Informar o código de registro do contador no conselho de contabilidade." />
                                <asp:DropDownList ID="ddlEstadoExpCRC" runat="server" Width="50px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Telefone Contador:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtTelefoneContador" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Email Contador:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtEmailContador" runat="server" Width="400px"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Qualificação:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtQualificacaoDoContador" runat="server" Width="100px"
                                    data-ToolTip="default" ToolTip="" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cd. Qualificação:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlCodigoDeQualificacaoDoContador" runat="server" Width="673px">
                                    <asp:ListItem Value=""></asp:ListItem>
                                    <asp:ListItem Value="203">203 - Diretor</asp:ListItem>
                                    <asp:ListItem Value="204">204 - Conselheiro de Administra&#231;&#227;o</asp:ListItem>
                                    <asp:ListItem Value="205">205 - Administrador</asp:ListItem>
                                    <asp:ListItem Value="206">206 - Administrador de Grupo</asp:ListItem>
                                    <asp:ListItem Value="207">207 - Administrador de Sociedade Filiada</asp:ListItem>
                                    <asp:ListItem Value="220">220 - Administrador Judicial</asp:ListItem>
                                    <asp:ListItem Value="222">222 - Administrador Judicial - Pessoa Juridica - Profissional Respons&#225;vel</asp:ListItem>
                                    <asp:ListItem Value="223">223 Administrador Judicial/Gestor</asp:ListItem>
                                    <asp:ListItem Value="226">226 Gestor Judicial</asp:ListItem>
                                    <asp:ListItem Value="309">309 - Procurador</asp:ListItem>
                                    <asp:ListItem Value="312">312 - Inventariante</asp:ListItem>
                                    <asp:ListItem Value="313">313 - Liquidante</asp:ListItem>
                                    <asp:ListItem Value="315">315 - Interventor</asp:ListItem>
                                    <asp:ListItem Value="801">801 - Empres&#225;rio</asp:ListItem>
                                    <asp:ListItem Value="900">900 - Contador</asp:ListItem>
                                    <asp:ListItem Value="999">999 - Outros</asp:ListItem>
                                </asp:DropDownList>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server">
                    <HeaderTemplate>
                        Contas
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Faixa das Contas de Custo
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta Inicial:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaInicial" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaInicial" runat="server" OnClick="btnContaInicial_Click" Text=" > "
                                        CssClass="btn" data-ToolTip="default" ToolTip="Conta Inicial de Custos." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta Final:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaFinal" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaFinal" runat="server" OnClick="btnContaFinal_Click" Text=" > "
                                        CssClass="btn" data-ToolTip="default" ToolTip="Conta final de Custos." />
                                </div>
                            </div>
                            <div class="subtitulodiv">
                                Patrimonio
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta Patrimonio:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaPatrimonio" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaPatrimonio" runat="server" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="" />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Contas de Variação Monetária
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Ativa:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaVariacaoAtiva" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaAtiva" runat="server" OnClick="btnContaAtiva_Click" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de variação monetária ativa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Passiva:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaVariacaoPassiva" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaPassiva" runat="server" OnClick="btnContaPassiva_Click" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de variação monetária passiva." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaVariacaoCliente" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaCliente" runat="server" OnClick="btnContaCliente_Click" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de variação monetária de clientes." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Fornecedor:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaVariacaoFornecedor" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaFornecedor" runat="server" OnClick="btnContaFornecedor_Click" CssClass="btn" Text=" > "
                                        data-ToolTip="default" ToolTip="Conta de variação monetária de fornecedores." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Caixa compensação:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCaixaCompensacao" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnCaixaCompensacao" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de variação monetária de caixa de compensação." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Conta Grupo de Contas
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Bancos:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaContabilGrupoBanco" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaBanco" runat="server" OnClick="btnContaBanco_Click" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta Bancos." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Juro Recebido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaJuroRecebido" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaJuroRecebido" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta Juros Recebidos." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Juro Pago:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaJuroPago" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaJuroPago" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta Juros Pagos." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Duplicatas Desc:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaDuplicatasDescontada" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaDuplicatasDescontada" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta de Duplicatas Descontadas." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Comissões:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaGrupoComissoes" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaGrupoComissoes" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta de comissões." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta TED/DOC:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaTEDDOC" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaTEDDOC" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta de contas TED/DOC." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Desconto Concedido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaDescontoConcedido" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaDescontoConcedido" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta Desconto Concedido." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Desconto Obtido:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaDescontoObtido" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaDescontoObtido" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Grupo de conta Desconto Obtido." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Contas de Fretes
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta Adto:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaAdiantamentoDeFrete" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaAdiantamentoDeFrete" runat="server" OnClick="btnContaAdiantamentoDeFrete_Click" CssClass="btn" Text=" > "
                                        data-ToolTip="default" ToolTip="Conta de fretes de adiantamento." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Conta Pedágio:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaPedagioDeFrete" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaPedagioDeFrete" runat="server" OnClick="btnContaPedagioDeFrete_Click" CssClass="btn" Text=" > "
                                        data-ToolTip="default" ToolTip="Conta de fretes de pedágio." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Fornecedor de Frete
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtFornecedorFrete" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnFornecedorFrete" runat="server" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de frornecedores de fretes." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Produtos de Fretes
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Produto de Frete:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProdutoDeFrete" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnProdutoDeFrete" runat="server" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de Produto de frete." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Produto de Estadia:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProdutoDeEstadia" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnProdutoDeEstadia" runat="server" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de Produto de estadia." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Produto de MF-e:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtProdutoDeMDFe" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnProdutoDeMDFe" runat="server" Text=" > " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de Produto MF-e." />
                                </div>
                            </div>
                        </div>
                        <div class="painelleft">
                            <div class="subtitulodiv">
                                Contas De Estoque
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Estoques:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaEstoque" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaEstoque" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de Estoques." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Em Nosso Poder:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaEstoqueEmNossoPoder" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaEstoqueEmNossoPoder" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de estoques em poder da empresa." />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl w133px">
                                    Em Poder De Terceiros:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtContaEstoqueEmPoderDeTerceiros" runat="server" Enabled="False" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnContaEstoqueEmPoderDeTerceiros" runat="server" Text=" &gt; " CssClass="btn"
                                        data-ToolTip="default" ToolTip="Conta de estoques em poder da terceiros." />
                                </div>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabECF" runat="server">
                    <HeaderTemplate>
                        ECF
                    </HeaderTemplate>
                    <ContentTemplate>
                        <ajaxToolkit:TabContainer ID="TabContainer2" runat="server" ActiveTabIndex="0" Width="100%">
                            <ajaxToolkit:TabPanel ID="TabECF0010" runat="server">
                                <HeaderTemplate>
                                    Bloco 0010
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        Parametros Bloco 0010
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Refis/Paes:
                                        </div>
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkOPT_REFIS" runat="server" Text="Optante Pelo Refis"
                                                data-ToolTip="default" ToolTip="" />
                                            <asp:CheckBox ID="chkOPT_PAES" runat="server" Text="Optante Pelo Paes"
                                                data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Forma de Tributação:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlFORMA_TRIB" runat="server" Width="740px" AutoPostBack="True">
                                                <asp:ListItem Value="0">Selecione uma Opção.</asp:ListItem>
                                                <asp:ListItem Value="1">1 – Lucro Real.</asp:ListItem>
                                                <asp:ListItem Value="2">2 – Lucro Real/Arbitrado.</asp:ListItem>
                                                <asp:ListItem Value="3">3 – Lucro Presumido/Real.</asp:ListItem>
                                                <asp:ListItem Value="4">4 – Lucro Presumido/Real/Arbitrado.</asp:ListItem>
                                                <asp:ListItem Value="5">5 – Lucro Presumido.</asp:ListItem>
                                                <asp:ListItem Value="6">6 – Lucro Arbitrado.</asp:ListItem>
                                                <asp:ListItem Value="7">7 – Lucro Presumido/Arbitrado.</asp:ListItem>
                                                <asp:ListItem Value="8">8 – Imune do IRPJ.</asp:ListItem>
                                                <asp:ListItem Value="9">9 – Isenta do IRPJ.</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Forma Periodo Ap. IRPJ e da CSLL:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlFORMA_APUR" runat="server" Width="740px">
                                                <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                                                <asp:ListItem Value="T">T – Trimestral</asp:ListItem>
                                                <asp:ListItem Value="A">A – Anual</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Qualificação da Pessoa Jurídica:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlCOD_QUALIF_PJ" runat="server" Width="740px">
                                                <asp:ListItem Value="0">Selecione uma Opção</asp:ListItem>
                                                <asp:ListItem Value="1">01 – PJ em Geral</asp:ListItem>
                                                <asp:ListItem Value="2">02 – PJ Componente do Sistema Financeiro</asp:ListItem>
                                                <asp:ListItem Value="3">03 – Sociedades Seguradoras, de Capitalização ou Entidade Aberta de Previdência Complementar</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Forma Tributação no Período: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlFORMA_TRIB_PER" runat="server" Width="740px">
                                                <asp:ListItem Value=" ">Selecione uma Opção.</asp:ListItem>
                                                <asp:ListItem Value="0">0 – ZERO – Não informado – trimestre não compreendido no período de apuração.</asp:ListItem>
                                                <asp:ListItem Value="R">R – Real</asp:ListItem>
                                                <asp:ListItem Value="P">P – Presumido</asp:ListItem>
                                                <asp:ListItem Value="A">A – Arbitrado</asp:ListItem>
                                                <asp:ListItem Value="E">E – Real Estimativa</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Forma Apuração Estimativa Mensal: 
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlMES_BAL_RED" runat="server" Width="740px">
                                                <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                                                <asp:ListItem Value="0">0 – Fora do Período: Fora do período de apuração/ Forma de tributação diferente de “R” ou “E”.</asp:ListItem>
                                                <asp:ListItem Value="E">E – Receita Bruta: Estimativa com base na receita bruta e acréscimos.</asp:ListItem>
                                                <asp:ListItem Value="B">B – Balanço ou Balancete: Estimativa com base no balanço ou balancete de suspensão/redução.</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            Escrituração:
                                        </div>
                                        <div class="coltxt">
                                            <asp:DropDownList ID="ddlTIP_ESC_PRE" runat="server" Width="740px">
                                                <asp:ListItem Value=" ">Selecione uma Opção.</asp:ListItem>
                                                <asp:ListItem Value="L">L – Livro Caixa ou Hipótese prevista no §1o do art. 129, Instrução Normativa no 1.515/2014 (Lucro Presumido) ou Sem Escrituração (Imunes ou Isentas) ou Não obrigadas a entregar a ECD, de acordo com a Instrução Normativa no 1.420/2014..</asp:ListItem>
                                                <asp:ListItem Value="C">C – Contábil (Lucro Presumido, Imunes ou Isentas)</asp:ListItem>
                                            </asp:DropDownList>
                                        </div>
                                    </div>
                                    <div id="DivIsentaImune" runat="server">
                                        <div class="row">
                                            <div class="collbl">
                                                Tipo PJ Imune/Isenta:
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlTIP_ENT" runat="server" Width="740px">
                                                    <asp:ListItem Value="0">Selecione uma Opção</asp:ListItem>
                                                    <asp:ListItem Value="1">01 – Assistência Social</asp:ListItem>
                                                    <asp:ListItem Value="2">02 – Educacional</asp:ListItem>
                                                    <asp:ListItem Value="3">03 – Sindicato de Trabalhadores</asp:ListItem>
                                                    <asp:ListItem Value="4">04 – Associação Civil</asp:ListItem>
                                                    <asp:ListItem Value="5">05 – Cultural</asp:ListItem>
                                                    <asp:ListItem Value="6">06 – Entidade Fechada de Previdência Complementar</asp:ListItem>
                                                    <asp:ListItem Value="7">07 – Filantrópica</asp:ListItem>
                                                    <asp:ListItem Value="8">08 – Sindicato</asp:ListItem>
                                                    <asp:ListItem Value="9">09 – Recreativa</asp:ListItem>
                                                    <asp:ListItem Value="10">10 – Científica</asp:ListItem>
                                                    <asp:ListItem Value="11">11 – Associação de Poupança e Empréstimo</asp:ListItem>
                                                    <asp:ListItem Value="12">12 – Entidade Aberta de Previdência Complementar (Sem Fins Lucrativos)</asp:ListItem>
                                                    <asp:ListItem Value="13">13 – Fifa e Entidades Relacionadas</asp:ListItem>
                                                    <asp:ListItem Value="14">14 – CIO e Entidades Relacionadas</asp:ListItem>
                                                    <asp:ListItem Value="15">15 – Partidos Políticos</asp:ListItem>
                                                    <asp:ListItem Value="99">99 – Outras.</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl">
                                                Apuração IRPJ p/ Imunes/Isentas: 
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlFORMA_APUR_I" runat="server" Width="740px">
                                                    <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                                                    <asp:ListItem Value="A">A – Anual</asp:ListItem>
                                                    <asp:ListItem Value="T">T – Trimestral</asp:ListItem>
                                                    <asp:ListItem Value="D">D – Desobrigada</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                        <div class="row">
                                            <div class="collbl">
                                                Apuração CSLL p/ Imunes/Isentas:
                                            </div>
                                            <div class="coltxt">
                                                <asp:DropDownList ID="ddlAPUR_CSLL" runat="server" Width="740px">
                                                    <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                                                    <asp:ListItem Value="A">A – Anual, se optou pela apuração da CSLL sobre a base de cálculo estimada, facultada a opção pelo levantamento de balanço ou balancete de suspensão ou redução.</asp:ListItem>
                                                    <asp:ListItem Value="T">T – Trimestral, no caso de ter adotado a apuração trimestral da CSLL.</asp:ListItem>
                                                    <asp:ListItem Value="D">D – Desobrigada, na hipótese de pessoa jurídica imune ou isenta da CSLL.</asp:ListItem>
                                                </asp:DropDownList>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="row">
                                        <div class="collbl">
                                            RTT:
                                        </div>
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkOPT_EXT_RTT" runat="server" Text="Optante pela extinção do RTT no ano-calendário de 2014"
                                                data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>

                                    <div class="row">
                                        <div class="collbl">
                                            FCont:
                                        </div>
                                        <div class="coltxt">
                                            <asp:CheckBox ID="chkDIF_FCONT" runat="server" Text="Existe diferenças entre a contabilidade societária e Fcont."
                                                data-ToolTip="default" ToolTip="" />
                                        </div>
                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                            <ajaxToolkit:TabPanel ID="TabECF0020" runat="server">
                                <HeaderTemplate>
                                    Bloco 0020
                                </HeaderTemplate>
                                <ContentTemplate>
                                    <div class="subtitulodiv">
                                        Parametros Bloco 0020
                                    </div>
                                    <div class="row">
                                        <div class="painelleft" style="width: 49%;">
                                            PJ Sujeita à Alíquota da CSLL de 9% ou 17% ou 20% em 31/12/2015:
                                            <asp:DropDownList ID="ddlIND_ALIQ_CSLL_I" runat="server" />
                                            <br>
                                            <asp:CheckBox ID="chkIND_ALIQ_CSLL" runat="server" Text="PJ Sujeita à Alíquota da CSLL de 15% - (Apenas para o ano 2014)"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            Qtde de SCP da PJ - Sócio Ostensivo de SCP :
                                            <asp:TextBox ID="txtIND_QTE_SCP" runat="server"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ADM_FUN_CLU" runat="server" Text="Administradora de Fundos e Clubes de Investimento"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PART_CONS" runat="server" Text="Participações em Consórcios de Empresas"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_OP_EXT" runat="server" Text="Operações com o Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_OP_VINC" runat="server" Text="Operações com Pessoa Vinculada/Interposta Pessoa / País com Tributação Favorecida"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PJ_ENQUAD" runat="server" Text="PJ Enquadrada nos artigos 48 ou 49 da IN RFB no 1.312/2012"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PART_EXT" runat="server" Text="Participações no Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ATIV_RURAL" runat="server" Text="Atividade Rural"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_LUC_EXP" runat="server" Text="Existência de Lucro da Exploração"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_RED_ISEN" runat="server" Text="Isenção e Redução do Imposto para Lucro Presumido"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_FIN" runat="server" Text="Indicativo da Existência de FINOR/FINAM/FUNRES"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_DOA_ELEIT" runat="server" Text="Doações a Campanhas Eleitorais"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PART_COLIG" runat="server" Text="Participação Avaliada pelo Método de Equivalência Patrimonial"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_VEND_EXP" runat="server" Text="PJ Efetuou Vendas a Empresa Comercial Exportadora com Fim Específico de Exportação"
                                                data-ToolTip="default" ToolTip="" />
                                        </div>
                                        <div class="painelleft" style="width: 49%;">
                                            <asp:CheckBox ID="chkIND_REC_EXT" runat="server" Text="Recebimentos do Exterior ou de Não Residentes"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ATIV_EXT" runat="server" Text="Ativos no Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_COM_EXP" runat="server" Text="PJ Comercial Exportadora"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PGTO_EXT" runat="server" Text="Pagamentos ao Exterior ou a Não Residentes"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_E_COM_TI" runat="server" Text="Comércio Eletrônico e Tecnologia da Informação"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ROY_REC" runat="server" Text="Royalties Recebidos do Brasil e do Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ROY_PAG" runat="server" Text="Royalties Pagos a Beneficiários do Brasil e do Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_REND_SERV" runat="server" Text="Rendimentos Relativos a Serviços, Juros e Dividendos Recebidos do Brasil e do Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PGTO_REM" runat="server" Text="Pagamentos ou Remessas a Título de Serviços, Juros e Dividendos a Beneficiários do Brasil e do Exterior"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_INOV_TEC" runat="server" Text="Inovação Tecnológica e Desenvolvimento Tecnológico"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_CAP_INF" runat="server" Text="Capacitação de Informática e Inclusão Digital"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_POLO_AM" runat="server" Text="Pólo Industrial de Manaus e Amazônia Ocidental"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_ZON_EXP" runat="server" Text="Zonas de Processamento de Exportação"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_AREA_COM" runat="server" Text="Áreas de Livre Comércio"
                                                data-ToolTip="default" ToolTip="" />
                                            <br />
                                            <asp:CheckBox ID="chkIND_PJ_HAB" runat="server" Text="PJ Habilitada no Repes, Recap, Padis, PATVD, Reidi, Repenec, Reicomp, Retaero, Recine, Resíduos Sólidos, Recopa, Copa do Mundo, Retid, REPNBL-Redes, Reif e Olimpíadas"
                                                data-ToolTip="default" ToolTip="" />
                                        </div>

                                    </div>
                                </ContentTemplate>
                            </ajaxToolkit:TabPanel>
                        </ajaxToolkit:TabContainer>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxtoolkit:tabcontainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:consultaplanodecontas id="ucConsultaPlanoDeContas" runat="server" />
    <uc:consultaproduto id="ucConsultaProduto" runat="server" />
</asp:Content>
