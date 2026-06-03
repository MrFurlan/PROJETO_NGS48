<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CertidaoNegativa.aspx.vb" Inherits="NGS.Web.UI.CertidaoNegativa" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .rot {
            font-family: Tahoma,Arial,Helvetica,sans-serif;
            font-size: 11px;
            font-weight: bold;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCertidaoNegativa" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCertidaoNegativa" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="hdnControlePopup" runat="server" />
            <div class="titulodiv">
                Certidão Negativa de Débito
            </div>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="TabPanel1" runat="server" HeaderText="TabPanel1">
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" class="iconNovo">
                                        <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                                    </li>
                                    <li runat="server" class="iconAtualizar">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server"
                                            Enabled="False" />
                                    </li>
                                    <li runat="server" class="iconExcluir">
                                        <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server"
                                            OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;"
                                            Enabled="False" />
                                    </li>
                                    <li runat="server" class="iconLimpar">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server" class="iconAjuda">
                                        <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Visualizar:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdConsolidado" runat="server" AutoPostBack="True" Checked="True"
                                    GroupName="rdc" OnCheckedChanged="rdConsolidado_CheckedChanged" Text="Consolidado" />
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="rdIndividual" runat="server" AutoPostBack="True" GroupName="rdc"
                                    OnCheckedChanged="rdIndividual_CheckedChanged" Text="Individual" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Modelo:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="ddlModeloCN" runat="server" Width="615px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCliente" runat="server" Enabled="False" TabIndex="3" Width="580px" />
                                <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdClientes" CssClass="btn" runat="server" OnClick="cmdClientes_Click"
                                    TabIndex="4" Text=" > " UseSubmitBehavior="False" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Número:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNumero" runat="server" />
                            </div>
                            <div class="collbl" style="margin-left: 27px">
                                Cód Autenticidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtAutenticidade" runat="server" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Data de Emissão:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataEmissao" CssClass="calendario" runat="server" Width="107px" />
                            </div>
                            <div class="collbl">
                                Data de Validade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtDataValidade" CssClass="calendario" runat="server" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="gridHistorico" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="gridHistorico_SelectedIndexChanged"
                                Width="100%">
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" />
                                    <asp:BoundField DataField="NomeCliente" HeaderText="Nome" />
                                    <asp:BoundField DataField="Numero" HeaderText="Numero" />
                                    <asp:BoundField DataField="DataEmissao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Emissao" />
                                    <asp:BoundField DataField="DataValidade" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Validade" />
                                    <asp:BoundField DataField="CodigoAutenticidade" HeaderText="Cod. Autenticidade" />
                                </Columns>
                                <EditRowStyle BackColor="#999999" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                    <HeaderTemplate>
                        Manutenção
                    </HeaderTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="TabPanel2" runat="server" HeaderText="TabPanel2">
                    <HeaderTemplate>
                        Consulta
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="painelleft" style="width: 720px;">
                            <div class="row">
                                <div class="collbl">
                                    Cliente:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtClienteConsulta" runat="server" Enabled="False" TabIndex="3"
                                        Width="450px" />
                                    <asp:HiddenField ID="txtCodigoClienteConsulta" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="cmdConsultaClientes" CssClass="btn" runat="server" OnClick="cmdConsultaClientes_Click"
                                        TabIndex="4" Text=" > " UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Inscrição:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtIE" runat="server" />
                                </div>
                                <div class="collbl">
                                    Número:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoNumero" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    Cód. Autenticidade:
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtCodigoAutenticidade" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:CheckBox ID="chkEmissao" Text="Emissao" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataEmissaoInicial" CssClass="calendario" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataEmissaoFinal" CssClass="calendario" runat="server" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="collbl">
                                    <asp:CheckBox ID="chkValidade" Text="Validade:" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataValidadeInicial" CssClass="calendario" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDataValidadeFinal" CssClass="calendario" runat="server" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnConsultaPrincipal" runat="server" CssClass="botao" OnClick="btnConsultaPrincipal_Click"
                                        Text="Consultar" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt rot">
                                    01. Clientes com Certidao para vencer em
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDias01" runat="server" Width="32px" />
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnComCertidao" runat="server" CssClass="botao" OnClick="btnComCertidao_Click"
                                        Text="Relatório" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt rot">
                                    02. Clientes com movimentacao no(s) ultimo(s)
                                </div>
                                <div class="coltxt">
                                    <asp:TextBox ID="txtDias02" runat="server" Width="34px" />
                                </div>
                                <div class="coltxt rot" style="">
                                    dias que estão sem certidao negativa.
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnSemCertidao" runat="server" CssClass="botao" OnClick="btnSemCertidao_Click"
                                        Text="Relatório" UseSubmitBehavior="False" />
                                </div>
                            </div>
                            <div class="row">
                                <div class="coltxt rot">
                                    03. Consultar Certidoes de nossas Empresas.
                                </div>
                                <div class="coltxt">
                                    <asp:Button ID="btnNossaEmpresa" runat="server" CssClass="botao" OnClick="btnNossaEmpresa_Click"
                                        Text="Consultar" UseSubmitBehavior="False" />
                                </div>
                            </div>
                        </div>
                        <div class="painelright" style="width: 157px;">
                            <table class="borda" border="0" style="text-align: right; position: relative;">
                                <tr>
                                    <td bgcolor="#e8e8e8" align="center" class="subtitulodiv">
                                        <label>
                                            Estados
                                        </label>
                                    </td>
                                </tr>
                                <tr>
                                    <td>
                                        <asp:CheckBoxList ID="chkEstados" runat="server" RepeatColumns="4" />
                                    </td>
                                </tr>
                            </table>
                        </div>
                        <div class="row">
                            <div class="bordagrid">
                                <asp:GridView ID="gridConsulta" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" Width="100%">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="CodigoCliente" HeaderText="Cliente" />
                                        <asp:BoundField DataField="NomeCliente" HeaderText="Nome" />
                                        <asp:BoundField DataField="Numero" HeaderText="Numero" />
                                        <asp:BoundField DataField="DataEmissao" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Emissao" />
                                        <asp:BoundField DataField="DataValidade" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Validade" />
                                        <asp:BoundField DataField="CodigoAutenticidade" HeaderText="Cod. Autenticidade" />
                                    </Columns>
                                </asp:GridView>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
