<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CadastroDePlacas.aspx.vb" Inherits="NGS.Web.UI.CadastroDePlacas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCadastroDePlacas" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCadastroDePlacas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro De Placas
            </div>
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
                            <asp:LinkButton class="iconAtualizar" ID="lnkLiberar" Text="Liberar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <fieldset style="border: 1px solid #BBBBBB; padding: 0px 0px 10px 10px; margin-top: 6px;">
                <legend style="font-size: 0.8em;"><strong><em>PLACA 1</em></strong></legend>
                <div class="row">
                    <div class="collbl">
                        Placa 1:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPlaca1" runat="server" Width="100px" CssClass="txtUpper" MaxLength="8"
                            data-ToolTip="default" ToolTip="Placa principal do veículo" />
                    </div>
                    <div class="collbl" style="margin-left: 133px">
                        RNTRC:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRNTRC1" runat="server" CssClass="txtNumerico8" Width="100px"
                            Enabled="false" data-ToolTip="default" ToolTip="Máximo 8 caracteres" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Proprietário:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtCodigoProprietario1" runat="server" />
                        <asp:TextBox ID="txtProprietario1" runat="server" Width="620px" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnProprietario1" runat="server" Text=" > " Enabled="false" OnClick="btnProprietario1_Click"
                            CssClass="btn" UseSubmitBehavior="False" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Estado:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEstadoPlaca1" AutoPostBack="True" runat="server" MaxLength="60"
                            Width="245px" />
                    </div>
                    <div class="collbl">
                        Município:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMunicipioPlaca1" runat="server" Width="400px" data-ToolTip="default"
                            ToolTip="Município da primeira placa" Enabled="false" />
                    </div>
                </div>
            </fieldset>
            <fieldset style="border: 1px solid #BBBBBB; padding: 0px 0px 10px 10px; margin-top: 6px;">
                <legend style="font-size: 0.8em;"><strong><em>PLACA 2</em></strong></legend>
                <div class="row">
                    <div class="collbl">
                        Placa 2:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPlaca2" runat="server" Width="100px" CssClass="txtUpper" MaxLength="8"
                            data-ToolTip="default" ToolTip="Segunda placa do veículo" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgLimparPlaca2" runat="server" ImageAlign="AbsMiddle" CssClass="btn"
                            Enabled="false" ImageUrl="~/Images/borracha.JPG" Width="22px" Style="border: 0px none;" />
                    </div>
                    <div class="collbl" style="margin-left: 104px">
                        RNTRC:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRNTRC2" runat="server" CssClass="txtNumerico8" MaxLength="8"
                            Enabled="false" Width="100px" data-ToolTip="default" ToolTip="Máximo 8 caracteres" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Proprietário:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtCodigoProprietario2" runat="server" />
                        <asp:TextBox ID="txtProprietario2" runat="server" Width="620px" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnProprietario2" runat="server" Text=" > " OnClick="btnProprietario2_Click"
                            CssClass="btn" Enabled="false" UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnRepPro2" runat="server" Text="Repetir Proprietário" UseSubmitBehavior="False"
                            Enabled="false" CssClass="botao" Width="130px" OnClick="btnRepPro2_Click" data-ToolTip="default"
                            ToolTip="Repete os dados conforme Placa Principal" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Estado:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEstadoPlaca2" AutoPostBack="True" runat="server" MaxLength="60"
                            Width="245px" />
                    </div>
                    <div class="collbl">
                        Município:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMunicipioPlaca2" runat="server" Width="400px" data-ToolTip="default"
                            ToolTip="Município da segunda placa" Enabled="false" />
                    </div>
                </div>
            </fieldset>
            <fieldset style="border: 1px solid #BBBBBB; padding: 0px 0px 10px 10px; margin-top: 6px;">
                <legend style="font-size: 0.8em;"><strong><em>PLACA 3</em></strong></legend>
                <div class="row">
                    <div class="collbl">
                        Placa 3:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPlaca3" Enabled="false" runat="server" CssClass="txtUpper" MaxLength="8"
                            Width="100px" data-ToolTip="default" ToolTip="Terceira placa do veículo" />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgLimparPlaca3" runat="server" CssClass="btn" ImageAlign="AbsMiddle"
                            ImageUrl="~/Images/borracha.JPG" Width="22px" Style="border: 0px none;" Enabled="false" />
                    </div>
                    <div class="collbl" style="margin-left: 104px">
                        RNTRC:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRNTRC3" runat="server" CssClass="txtNumerico8" MaxLength="8"
                            Enabled="false" Width="100px" data-ToolTip="default" ToolTip="Máximo 8 caracteres" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Proprietário:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtCodigoProprietario3" runat="server" />
                        <asp:TextBox ID="txtProprietario3" runat="server" Width="620px" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnProprietario3" runat="server" Text=" > " OnClick="btnProprietario3_Click"
                            CssClass="btn" UseSubmitBehavior="False" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnRepPro3" runat="server" Text="Repetir Proprietário" UseSubmitBehavior="False"
                            data-ToolTip="default" ToolTip="Repete os dados conforme Placa Principal" Width="130px"
                            OnClick="btnRepPro3_Click" CssClass="botao" Enabled="false" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Estado:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEstadoPlaca3" AutoPostBack="True" runat="server" MaxLength="60"
                            Width="245px" />
                    </div>
                    <div class="collbl">
                        Município:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMunicipioPlaca3" runat="server" Width="400px" data-ToolTip="default"
                            ToolTip="Município da terceira placa" Enabled="false" />
                    </div>
                </div>
            </fieldset>
            <fieldset style="border: 1px solid #BBBBBB; padding: 0px 0px 10px 10px; margin-top: 6px;">
                <legend style="font-size: 0.8em;"><strong><em>PLACA 4</em></strong></legend>
                <div class="row">
                    <div class="collbl">
                        Placa 4:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPlaca4" runat="server" Width="100px" CssClass="txtUpper" MaxLength="8"
                            data-ToolTip="default" ToolTip="Quarta placa do veículo" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgLimparPlaca4" runat="server" Height="22px" ImageAlign="AbsMiddle"
                            CssClass="btn" ImageUrl="~/Images/borracha.JPG" Width="22px" Style="border: 0px none;"
                            Enabled="false" />
                    </div>
                    <div class="collbl" style="margin-left: 104px">
                        RNTRC:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtRNTRC4" runat="server" CssClass="txtNumerico8" MaxLength="8"
                            Enabled="false" Width="100px" data-ToolTip="default" ToolTip="Máximo 8 caracteres" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Proprietário:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtCodigoProprietario4" runat="server" />
                        <asp:TextBox ID="txtProprietario4" runat="server" Width="620px" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnProprietario4" runat="server" Text=" > " OnClick="btnProprietario4_Click"
                            CssClass="btn" Enabled="false" UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnRepPro4" runat="server" Text="Repetir Proprietário" UseSubmitBehavior="False"
                            Enabled="false" data-ToolTip="default" ToolTip="Repete os dados conforme Placa Principal"
                            OnClick="btnRepPro4_Click" CssClass="botao" Width="130px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Estado:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEstadoPlaca4" AutoPostBack="True" runat="server" MaxLength="60"
                            Width="245px" />
                    </div>
                    <div class="collbl">
                        Município:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMunicipioPlaca4" runat="server" Width="400px" data-ToolTip="default"
                            ToolTip="Município da quarta placa" Enabled="false" />
                    </div>
                </div>
            </fieldset>
            <fieldset style="border: 1px solid #BBBBBB; padding: 0px 0px 10px 10px; margin-top: 6px;">
                <legend style="font-size: 0.8em;"><strong><em>OUTRAS INFORMAÇÕES</em></strong></legend>
                <div class="row">
                    <div class="collbl">
                        Tipo de Veículo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTipoDeVeiculo" runat="server" Width="510px" Enabled="false" />
                    </div>
                    <div class="collbl" style="margin-left: 26px">
                        Via de Transporte:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlViaDeTransporte" Enabled="false" Width="115px" runat="server" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Motorista:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMotorista" runat="server" data-ToolTip="default" ToolTip="Dados do Motorista"
                            Width="440px" Enabled="false" />
                    </div>
                    <div class="coltxt">
                        <asp:Button ID="btnMotorista" CssClass="btn" Enabled="false" runat="server" Text=" > "
                            UseSubmitBehavior="False" />
                    </div>
                    <div class="coltxt">
                        <asp:ImageButton ID="imgLimparMotorista" CssClass="btn" runat="server" Height="22px"
                            ImageAlign="AbsMiddle" Enabled="false" ImageUrl="~/Images/borracha.JPG" Width="22px"
                            Style="border: 0px none;" />
                    </div>
                    <div class="collbl" style="margin-left: 27px">
                        Habilitação:
                    </div>
                    <div class="coltxt">
                        <asp:HiddenField ID="txtCodigoMotorista" runat="server" />
                        <asp:TextBox ID="txtHabilitacao" Enabled="false" runat="server" data-ToolTip="default"
                            ToolTip="Habilitação do Motorista" Width="110px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Estado:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtEstadoMotorista" runat="server" data-ToolTip="default" ToolTip="Estado do Motorista"
                            Enabled="false" Width="35px" />
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNomeEstadoMotorista" runat="server" data-ToolTip="default" ToolTip="Nome do estado do Motorista"
                            Enabled="false" Width="130px" />
                    </div>
                    <div class="collbl">
                        Município:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtMunicipioMotorista" runat="server" data-ToolTip="default" ToolTip="Município do Motorista"
                            Width="210px" Enabled="false" />
                    </div>
                    <div class="collbl">
                        Restrição:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="radNao" Enabled="false" runat="server" GroupName="restricao" Text="Não" AutoPostBack="true" Checked="true" />
                        <asp:RadioButton ID="radSim" Enabled="false" runat="server" GroupName="restricao" Text="Sim" AutoPostBack="true" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Observação:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtObservacao" runat="server" TextMode="MultiLine" Width="751px" />
                    </div>
                </div>
            </fieldset>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
</asp:Content>
