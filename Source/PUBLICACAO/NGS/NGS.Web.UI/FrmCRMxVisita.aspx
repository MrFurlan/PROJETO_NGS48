<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="FrmCRMxVisita.aspx.vb" Inherits="NGS.Web.UI.frmCRMxVisita" %>

<%@ Register Assembly="AjaxControlToolkit" Namespace="AjaxControlToolkit" TagPrefix="cc1" %>
<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        input[type='text']
        {
            height: 16px;
        }
        
        #meioconteudo
        {
            width: 980px !important;
        }
        
        select
        {
            height: 20px;
        }
        
        .Hide
        {
            display: none;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCRMxVisitas" runat="server" />
    <asp:HiddenField ID="HID" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCRMxVisitas" runat="server">
        <ContentTemplate>
            <asp:Panel ID="Panel1" runat="server" Height="100%" Width="100%">
                <table width="100%">
                    <tr>
                        <td colspan="3" style="font-weight: bold; font-size: 32px; color: white; font-family: Arial;
                            background-color: #003366; text-align: center">
                            CRM x Visitas
                        </td>
                    </tr>
                    <tr>
                        <td colspan="3">
                            <!-- TabContainer Principal -->
                            <ajaxToolkit:TabContainer ID="tcCRMxVisitas" runat="server" ActiveTabIndex="2" 
                                Width="100%">
                                <!-- Tab de Consulta CRM -->
                                <ajaxToolkit:TabPanel ID="tbConsultaCRM" runat="server" HeaderText="Parâmetros">
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel8" runat="server" Height="600px" Width="100%">
                                            <table width="100%">
                                                <tr>
                                                    <td align="center" class="titulotabela" colspan="4" style="font-weight: bold; font-size: 22px;
                                                        background-color: #C0C0C0;">
                                                        CRM
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td class="primario" style="width: 100px; white-space: nowrap;" valign="top">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td>
                                                        <asp:DropDownList ID="ddlParametroAnoConsulta" runat="server" AutoPostBack="True"
                                                            Height="19px" Width="94px">
                                                        </asp:DropDownList>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="4">
                                                        <div style="overflow: auto;">
                                                            <asp:GridView ID="gridParametroConsulta" runat="server" AutoGenerateColumns="False"
                                                                CellPadding="4" ForeColor="#333333" GridLines="None">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <EditRowStyle BackColor="#999999" />
                                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                <Columns>
                                                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                                    <asp:BoundField DataField="CodigoEmpresa" HeaderText="Codigo">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="EndEmpresa" HeaderText="End" />
                                                                    <asp:BoundField DataField="NomeEmpresa" HeaderText="Nome">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                </Columns>
                                                            </asp:GridView>
                                                        </div>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="2">
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td style="width: 25%;" valign="top">
                                                        <div class="bordasimples" style="min-height: 100%; max-height: 100%; vertical-align: top">
                                                            <asp:GridView ID="grdParametroRepresentante" runat="server" AutoGenerateColumns="False"
                                                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                <Columns>
                                                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                                        <ItemStyle Width="25px"></ItemStyle>
                                                                    </asp:CommandField>
                                                                    <asp:BoundField DataField="CodigoRepresentante" HeaderText="Código">
                                                                        <HeaderStyle HorizontalAlign="Right" CssClass="Hide" />
                                                                        <ItemStyle HorizontalAlign="Right" Width="20px" CssClass="Hide" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="EndRepresentante" HeaderText="Endereço">
                                                                        <HeaderStyle HorizontalAlign="Right" CssClass="Hide" />
                                                                        <ItemStyle HorizontalAlign="Right" Width="20px" CssClass="Hide" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="NomeRepresentante" HeaderText="Representante">
                                                                        <HeaderStyle HorizontalAlign="Left" Height="30px"></HeaderStyle>
                                                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                                    </asp:BoundField>
                                                                </Columns>
                                                                <EditRowStyle BackColor="#999999" />
                                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                <PagerTemplate>
                                                                    Representantes
                                                                </PagerTemplate>
                                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            </asp:GridView>
                                                        </div>
                                                    </td>
                                                    <td style="width: 75%;" valign="top">
                                                        <asp:Panel ID="pnlParametrosVisitas" runat="server">
                                                            <table class="actions" style="width: 100%;">
                                                                <tr>
                                                                    <td class="iconNovo" style="width: 15%;">
                                                                        <asp:LinkButton ID="lnkNovo" runat="server">
                                                                                            <span>Gravar</span>
                                                                        </asp:LinkButton>
                                                                    </td>
                                                                    <td class="iconRelatorio" style="width: 15%;">
                                                                        <asp:LinkButton ID="lnkRelatorio" runat="server">
                                                                                            <span>Relatório</span>
                                                                        </asp:LinkButton>
                                                                    </td>
                                                                    <td class="iconAjuda" style="width: 15%;">
                                                                        <asp:LinkButton ID="lnkAjuda" runat="server">
                                                                                            <span>Ajuda</span>
                                                                        </asp:LinkButton>
                                                                    </td>
                                                                    <td style="width: 55%; display: block;">
                                                                    </td>
                                                                </tr>
                                                            </table>
                                                        </asp:Panel>
                                                        <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                            <asp:GridView ID="grdParametroVisita" runat="server" AutoGenerateColumns="False"
                                                                CellPadding="4" ForeColor="#333333" GridLines="None" OnSelectedIndexChanged="grdParametroCliente_SelectedIndexChanged"
                                                                Width="100%">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                <Columns>
                                                                    <asp:BoundField DataField="ID" HeaderText="ID">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="NomeCliente" HeaderText="Cliente">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Fazenda" HeaderText="Fazenda">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Classificacao" HeaderText="Classificação">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Qualitativo" HeaderText="Qualitativo">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:TemplateField HeaderText="Visitas Mês">
                                                                        <ItemTemplate>
                                                                            <asp:TextBox ID="txtNumVisita" runat="server" AutoPostBack="True" OnTextChanged="txtNumVisita_TextChanged"
                                                                                Text='<%# Eval("MinimoVisita") %>' Width="30px"/>
                                                                        </ItemTemplate>
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:TemplateField>
                                                                </Columns>
                                                                <EditRowStyle BackColor="#999999" />
                                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            </asp:GridView>
                                                        </div>
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <!-- Tab de Visitas -->
                                <ajaxToolkit:TabPanel ID="tabVisitas" runat="server" HeaderText="Visitas">
                                    <ContentTemplate>
                                        <table width="100%">
                                            <!-- Título -->
                                            <tr>
                                                <td class="titulotabela" align="center" colspan="2" style="font-weight: bold; font-size: 22px;
                                                    background-color: #C0C0C0;">
                                                    Visita aos Clientes
                                                </td>
                                            </tr>
                                            <tr>
                                                <td>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td colspan="2" >
                                                    <asp:Panel ID="pHeader" runat="server" CssClass="cpHeader, titulotabela" >
                                                        <asp:Image ID="imgCPanel" runat="server" />
                                                        <asp:Label ID="lblText" runat="server" CssClass="tituloTabela" />
                                                    </asp:Panel>
                                                    <asp:Panel ID="pBody" runat="server" CssClass="cpBody">
                                                        <table width="100%">
                                                            <tr>
                                                                <td>
                                                                    <asp:RadioButton ID="rbVisitaCRM" runat="server" Text="CRM" GroupName="gpConsultaVisita"
                                                                        AutoPostBack="True" Checked="True" />
                                                                    <asp:RadioButton ID="rbVisitaAvulsa" runat="server" Text="Avulsas" GroupName="gpConsultaVisita"
                                                                        AutoPostBack="True" />
                                                                    <asp:RadioButton ID="rbTodasVisitas" runat="server" Text="Todas" GroupName="gpConsultaVisita"
                                                                        AutoPostBack="True" />
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td style="width: 40%">
                                                                    <asp:Panel runat="server" ID="pnlConsultaAvulsaTodas">
                                                                        <table width="100%">
                                                                            <tr valign="top">
                                                                                <td class="primario" style="width: 50px; white-space: nowrap;" valign="top">
                                                                                    <div class="headerGray">
                                                                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Período:</span>
                                                                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                                    </div>
                                                                                </td>
                                                                                <td>
                                                                                    <asp:TextBox ID="txtConsultaDataInicial" CssClass="calendario" runat="server" Width="65px"/>
                                                                                    &nbsp; a
                                                                                    <asp:TextBox ID="txtConsultaDataFinal" CssClass="calendario" runat="server" Width="65px"/>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2">
                                                                                    <br />
                                                                                    <asp:Label ID="Label1" runat="server" Text="Região" />
                                                                                    <br />
                                                                                    <asp:DropDownList ID="ddlRegiao" runat="server" Width="90%"  
                                                                                        AutoPostBack="True">
                                                                                    </asp:DropDownList>
                                                                                    <br />
                                                                                    <asp:Label ID="Label2" runat="server" Text="Micro Região" />
                                                                                    <br />
                                                                                    <asp:DropDownList ID="ddlMicroRegiao" runat="server" Width="90%" 
                                                                                         AutoPostBack="True">
                                                                                    </asp:DropDownList>
                                                                                    <br />
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2" align="right">
                                                                                    <asp:LinkButton ID="lnkConsultaVisitas" runat="server" CssClass="iconConsultar">
                                                                                        <span>Consultar</span>
                                                                                    </asp:LinkButton>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                </td>
                                                            </tr>
                                                            <tr>
                                                                <td valign="top">
                                                                    <asp:Panel ID="pnlConsultaCRM" runat="server" Visible="False" Width="100%">
                                                                        <table width="100%">
                                                                            <tr>
                                                                                <td class="primario" style="width: 100px; white-space: nowrap;" valign="top">
                                                                                    <div class="headerGray">
                                                                                        <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                                                        <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                                    </div>
                                                                                </td>
                                                                                <td>
                                                                                    <asp:DropDownList ID="ddlConsultaAnoCRM" runat="server" AutoPostBack="True" Height="19px"
                                                                                        Width="94px">
                                                                                    </asp:DropDownList>
                                                                                </td>
                                                                            </tr>
                                                                            <tr>
                                                                                <td colspan="2">
                                                                                    <div style="overflow: auto;">
                                                                                        <asp:GridView ID="grdConsultaCRM" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                                                            ForeColor="#333333" GridLines="None">
                                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                            <EditRowStyle BackColor="#999999" />
                                                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                                            <Columns>
                                                                                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                                                                                <asp:BoundField DataField="CodigoEmpresa" HeaderText="Codigo">
                                                                                                    <HeaderStyle HorizontalAlign="Left" />
                                                                                                    <ItemStyle HorizontalAlign="Left" />
                                                                                                </asp:BoundField>
                                                                                                <asp:BoundField DataField="EndEmpresa" HeaderText="End" />
                                                                                                <asp:BoundField DataField="NomeEmpresa" HeaderText="Nome">
                                                                                                    <HeaderStyle HorizontalAlign="Left" />
                                                                                                    <ItemStyle HorizontalAlign="Left" />
                                                                                                </asp:BoundField>
                                                                                            </Columns>
                                                                                        </asp:GridView>
                                                                                    </div>
                                                                                </td>
                                                                            </tr>
                                                                        </table>
                                                                    </asp:Panel>
                                                                </td>
                                                            </tr>
                                                        </table>
                                                    </asp:Panel>
                                                    <cc1:CollapsiblePanelExtender ID="CollapsiblePanelExtender1" runat="server" TargetControlID="pBody"
                                                        CollapseControlID="pHeader" ExpandControlID="pHeader" Collapsed="True" TextLabelID="lblText"
                                                        CollapsedText="Consulta" ExpandedText="Consulta" CollapsedSize="0" Enabled="True"
                                                        ImageControlID="imgCPanel" CollapsedImage="~/images/duasSetasParaCima.png" ExpandedImage="~/images/duasSetasParaBaixo.png">
                                                    </cc1:CollapsiblePanelExtender>
                                                </td>
                                            </tr>
                                            <tr>
                                                <td class="titulotabela" align="center" colspan="2" style="font-weight: bold; font-size: 10px;
                                                    background-color: #C0C0C0;">
                                            </tr>
                                            <tr>
                                                <td valign="top">
                                                    <table width="100%">
                                                        <tr>
                                                            <td valign="top">
                                                                <asp:Panel ID="pnlRepresentante" runat="server" Width="230px">
                                                                    <div class="bordasimples" style="min-height: 100%; max-height: 100%; vertical-align: top">
                                                                        <asp:GridView ID="grdConsultaRepresentante" runat="server" AutoGenerateColumns="False"
                                                                            CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                            <EditRowStyle BackColor="#999999" />
                                                                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                            <Columns>
                                                                                <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                                                                    <ItemStyle Width="25px"></ItemStyle>
                                                                                </asp:CommandField>
                                                                                <asp:BoundField DataField="CodigoRepresentante" HeaderText="Código">
                                                                                    <HeaderStyle HorizontalAlign="Right" CssClass="Hide" />
                                                                                    <ItemStyle HorizontalAlign="Right" Width="20px" CssClass="Hide" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="EndRepresentante" HeaderText="Endereço">
                                                                                    <HeaderStyle HorizontalAlign="Right" CssClass="Hide" />
                                                                                    <ItemStyle HorizontalAlign="Right" Width="20px" CssClass="Hide" />
                                                                                </asp:BoundField>
                                                                                <asp:BoundField DataField="NomeRepresentante" HeaderText="Representante">
                                                                                    <HeaderStyle HorizontalAlign="Left" Height="30px"></HeaderStyle>
                                                                                    <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                                                                </asp:BoundField>
                                                                            </Columns>
                                                                        </asp:GridView>
                                                                    </div>
                                                                </asp:Panel>
                                                            </td>
                                                            <td>
                                                                <!-- TabContainer Interno Aba Visitas-->
                                                                <asp:Panel ID="pnlClientesVisitas" runat="server" Width="690px">
                                                                    <ajaxToolkit:TabContainer ID="tcVisitas" runat="server" ActiveTabIndex="1" 
                                                                        Width="100%">
                                                                        <ajaxToolkit:TabPanel ID="tbClientes" runat="server" HeaderText="Clientes">
                                                                            <ContentTemplate>
                                                                                <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                                                    <asp:GridView ID="grdConsultaCliente" runat="server" AutoGenerateColumns="False"
                                                                                        CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                                                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                                                        <Columns>
                                                                                            <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                                                                                <ItemStyle Width="25px" />
                                                                                            </asp:CommandField>
                                                                                            <asp:BoundField DataField="id" HeaderText="ID">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="NomeCliente" HeaderText="Cliente">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="Fazenda" HeaderText="Fazenda">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="CodigoRepresentante" HeaderText="Cód. Representante">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" Width="25px" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="NomeRepresentante" HeaderText="Nome Representante">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                                            </asp:BoundField>
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
                                                                        <ajaxToolkit:TabPanel ID="tbVisualizaVisita" runat="server" HeaderText="Visitas">
                                                                            <ContentTemplate>
                                                                                <table width="100%">
                                                                                    <tr>
                                                                                        <td>
                                                                                            <div class="headerGray" style="min-width: 100%; max-width: 100%; overflow: auto;">
                                                                                                <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                                                                                                <span style="margin-left: 20px; margin-top: 1px; margin-right: 0px;"></span>
                                                                                                <asp:Label ID="lblCliente" runat="server" Style="font-size: 18px;"/>
                                                                                            </div>
                                                                                        </td>
                                                                                        <td>
                                                                                            <div class="headerGray" style="min-width: 100%; max-width: 100%; overflow: auto;">
                                                                                                <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Fazenda:</span>
                                                                                                <span style="margin-left: 20px; margin-top: 1px; margin-right: 0px;"></span>
                                                                                                <asp:Label ID="lblFazenda" runat="server" Style="font-size: 18px;"/>
                                                                                            </div>
                                                                                        </td>
                                                                                    </tr>
                                                                                </table>
                                                                                <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                                                    <asp:GridView ID="grdVisualizaVisita" runat="server" CellPadding="4" ForeColor="#333333"
                                                                                        GridLines="None" Width="100%" AutoGenerateColumns="False">
                                                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                                        <Columns>
                                                                                            <asp:BoundField DataField="NumeroVisita" HeaderText="Visita">
                                                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                                                <ItemStyle HorizontalAlign="Right" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="Data" HeaderText="Data" DataFormatString="{0:dd/MM/yyyy}">
                                                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                                                <ItemStyle HorizontalAlign="Left" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="KmInicial" HeaderText="Km Inicial">
                                                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                                                <ItemStyle HorizontalAlign="Right" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="KmFinal" HeaderText="Km Final">
                                                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                                                <ItemStyle HorizontalAlign="Right" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="KmRodado" HeaderText="Km Rodado">
                                                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                                                <ItemStyle HorizontalAlign="Right" Width="40px" />
                                                                                            </asp:BoundField>
                                                                                            <asp:BoundField DataField="Finalidade" HeaderText="Finalidade">
                                                                                                <HeaderStyle HorizontalAlign="Justify" />
                                                                                                <ItemStyle HorizontalAlign="Left" Width="250px" />
                                                                                            </asp:BoundField>
                                                                                            <asp:TemplateField HeaderText="Alterar">
                                                                                                <ItemTemplate>
                                                                                                    <asp:LinkButton ID="imdAlterar" runat="server" OnClick="ImdAlterar_Click">
                                                                                                        <asp:Image ID="imgEdit" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/edit.png"
                                                                                                            data-ToolTip="default" ToolTip="Alterar Visita" />
                                                                                                    </asp:LinkButton>
                                                                                                </ItemTemplate>
                                                                                                <HeaderStyle HorizontalAlign="Center" />
                                                                                                <ItemStyle Width="30px" HorizontalAlign="Center" VerticalAlign="Middle" />
                                                                                            </asp:TemplateField>
                                                                                            <asp:TemplateField HeaderText="Visualizar">
                                                                                                <ItemTemplate>
                                                                                                    <asp:LinkButton ID="imdVisualizar" runat="server" OnClick="ImdVisualizar_Click">
                                                                                                        <asp:Image ID="imgpdf" runat="server" Width="16px" Height="16px" ImageUrl="~/Images/pdf.png"
                                                                                                            data-ToolTip="default" ToolTip="Visualizar Visita" />
                                                                                                    </asp:LinkButton>
                                                                                                </ItemTemplate>
                                                                                                <HeaderStyle HorizontalAlign="Center" />
                                                                                                <ItemStyle Width="30px" HorizontalAlign="Center" VerticalAlign="Middle" />
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
                                                                    </ajaxToolkit:TabContainer>
                                                                </asp:Panel>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                </td>
                                                <td>
                                                </td>
                                            </tr>
                                        </table>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <!-- Tab de Cadastros -->
                                <ajaxToolkit:TabPanel ID="tbCadastro" runat="server" HeaderText="TabPanel4">
                                    <HeaderTemplate>
                                        Cadastro
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <table id="tblActionsParametros" class="actions" style="width: 100%;">
                                            <tr>
                                                <td class="iconNovo" style="width: 10%;">
                                                    <asp:LinkButton ID="lnkNovoCadVisita" runat="server">
                                                         <span>Gravar</span>
                                                    </asp:LinkButton>
                                                </td>
                                                <td class="iconAtualizar" runat="server" style="width: 10%;">
                                                    <asp:LinkButton ID="lnkAtualizar" runat="server">
                                                        <span>Atualizar</span>
                                                    </asp:LinkButton>
                                                </td>
                                                <td class="iconLimpar" style="width: 10%;">
                                                    <asp:LinkButton ID="lnkLimparCadVisita" runat="server">
                                                         <span>Limpar</span>
                                                    </asp:LinkButton>
                                                </td>
                                                <td class="iconRelatorio" style="width: 10%;">
                                                    <asp:LinkButton ID="lnkFormularioCadVisita" runat="server" data-ToolTip="default" ToolTip="Espelho do Formulário">
                                                                    <span>Formulário</span>
                                                    </asp:LinkButton>
                                                </td>
                                                <td class="iconAjuda" style="width: 10%;">
                                                    <asp:LinkButton ID="lnkAjudaCadVisita" runat="server">
                                                          <span>Ajuda</span>
                                                    </asp:LinkButton>
                                                </td>
                                                <td style="width: 50%; display: block;">
                                                </td>
                                            </tr>
                                        </table>
                                        <div class="bordasimples" style="min-height: 130px; max-height: 130px; overflow: auto;">
                                            <table width="100%">
                                                <tr>
                                                    <td width="100px">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Representante:</span>
                                                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td colspan="7">
                                                        <asp:DropDownList ID="ddlRepresentante" runat="server" 
                                                            Width="510px">
                                                        </asp:DropDownList>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td width="100px">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                                                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td colspan="3">
                                                        <asp:Button ID="btnCliente" runat="server" Text=" > "></asp:Button>
                                                        <asp:Label ID="txtNomeDoCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                            />
                                                        <asp:HiddenField ID="txtCodigoCliente" runat="server"></asp:HiddenField>
                                                        <asp:HiddenField ID="txtEndCliente" runat="server"></asp:HiddenField>
                                                    </td>
                                                    <td>
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Fazenda:</span>
                                                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td width="310px" colspan="3">
                                                        <asp:Label ID="txtFazendaCliente" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                            />
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Data:</span>
                                                        </div>
                                                    </td>
                                                    <td width="120px">
                                                        <asp:TextBox ID="txtData" CssClass="calendario" runat="server" Width="80px"/>
                                                    </td>
                                                    <td width="70px">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">KmInicial:</span>
                                                        </div>
                                                    </td>
                                                    <td width="92px">
                                                        <asp:TextBox ID="txtKmInicial" runat="server" Width="90px" Style="text-align: right;"/>
                                                    </td>
                                                    <td>
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">KmFinal:</span>
                                                        </div>
                                                    </td>
                                                    <td>
                                                        <asp:TextBox ID="txtKmFinal" runat="server" Width="90px" Style="text-align: right;"/>
                                                    </td>
                                                    <td>
                                                    </td>
                                                    <td>
                                                    </td>
                                                </tr>
                                            </table>
                                        </div>
                                        <!-- TabContainer Interno - Aba Cadastro de Visitas  -->
                                        <ajaxToolkit:TabContainer ID="tcOutros" runat="server" ActiveTabIndex="0" CssClass=""
                                            Width="100%">
                                            <ajaxToolkit:TabPanel ID="tbMotivos" runat="server">
                                                <HeaderTemplate>
                                                    Motivo
                                                </HeaderTemplate>
                                                <ContentTemplate>
                                                    <table class="actions" style="width: 100%;">
                                                        <tr>
                                                            <td class="iconNovo" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkGravarMotivo" runat="server">
                                                                                            <span>Gravar</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconExcluir" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkExcluirMotivo" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                                                    <span>Excluir</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconRelatorio" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkRelatorioMotivo" runat="server">
                                                                                            <span>Relatório</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconAjuda" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkAjudaMotivo" runat="server">
                                                                                            <span>Ajuda</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td style="width: 60%; display: block;">
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table style="width: 100%;">
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Motivo:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlMotivo" runat="server" 
                                                                    Width="180px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Observação:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtObsMotivo" runat="server" Width="500px"/>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <!--Inclusão de Motivos de Visitas -->
                                                    <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                        <asp:GridView ID="grdVisitaMotivos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                                            ForeColor="#333333" GridLines="None">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <EditRowStyle BackColor="#999999" />
                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            <Columns>
                                                                <asp:CommandField ButtonType="Button" SelectText="&gt;" ShowSelectButton="True">
                                                                    <ItemStyle Width="25px" />
                                                                </asp:CommandField>
                                                                <asp:BoundField DataField="CodigoMotivo" HeaderText="Código">
                                                                    <HeaderStyle HorizontalAlign="Right" CssClass="Hide" />
                                                                    <ItemStyle HorizontalAlign="Right" Width="20px" CssClass="Hide" />
                                                                </asp:BoundField>
                                                                <asp:TemplateField HeaderText="Motivo">
                                                                    <ItemTemplate>
                                                                        <%# Eval("Motivo.Descricao") %>
                                                                    </ItemTemplate>
                                                                    <HeaderStyle HorizontalAlign="Left" />
                                                                    <ItemStyle Width="100px" />
                                                                </asp:TemplateField>
                                                                <asp:BoundField DataField="Observacao" HeaderText="Observação">
                                                                    <HeaderStyle HorizontalAlign="Left" />
                                                                    <ItemStyle Width="300px" />
                                                                </asp:BoundField>
                                                            </Columns>
                                                        </asp:GridView>
                                                    </div>
                                                </ContentTemplate>
                                            </ajaxToolkit:TabPanel>
                                            <ajaxToolkit:TabPanel ID="tbAmeaca" runat="server">
                                                <HeaderTemplate>
                                                    Ameaça
                                                </HeaderTemplate>
                                                <ContentTemplate>
                                                    <table class="actions" style="width: 100%;">
                                                        <tr>
                                                            <td class="iconNovo" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkGravarAmeaca" runat="server">
                                                                                            <span>Gravar</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconExcluir" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkExcluirAmeaca" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                                                    <span>Excluir</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconRelatorio" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkRelatorioAmeaca" runat="server">
                                                                                            <span>Relatório</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconAjuda" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkAjudaAmeaca" runat="server">
                                                                                            <span>Ajuda</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td style="width: 60%; display: block;">
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table style="width: 100%;">
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ameaça:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlAmeaca" runat="server" 
                                                                    Width="300px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Observação:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="txtObsAmeaca" runat="server" Width="350px"/>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                        <asp:GridView ID="grdVisitaAmeaca" runat="server" CellPadding="4" ForeColor="#333333"
                                                            GridLines="None" Width="100%" AutoGenerateColumns="False">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <EditRowStyle BackColor="#999999" />
                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px"
                                                                HorizontalAlign="Left" />
                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Left" />
                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            <Columns>
                                                                <asp:CommandField ButtonType="Button" ShowSelectButton="True" SelectText="&gt;">
                                                                    <ItemStyle Width="25px" />
                                                                </asp:CommandField>
                                                                <asp:BoundField DataField="CodigoAmeaca">
                                                                    <HeaderStyle CssClass="Hide" />
                                                                    <ItemStyle CssClass="Hide" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="EndAmeaca">
                                                                    <HeaderStyle CssClass="Hide" />
                                                                    <ItemStyle CssClass="Hide" />
                                                                </asp:BoundField>
                                                                <asp:TemplateField HeaderText="Ameaça" ItemStyle-Width="150px">
                                                                    <ItemTemplate>
                                                                        <%# Eval("Ameaca.Nome")%>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                                <asp:TemplateField HeaderText="Endereço" ItemStyle-Width="150px">
                                                                    <ItemTemplate>
                                                                        <%# Eval("Ameaca.Endereco")%>
                                                                    </ItemTemplate>
                                                                </asp:TemplateField>
                                                                <asp:BoundField DataField="Observacao" HeaderText="Observação" ItemStyle-Width="150px" />
                                                            </Columns>
                                                        </asp:GridView>
                                                    </div>
                                                </ContentTemplate>
                                            </ajaxToolkit:TabPanel>
                                            <ajaxToolkit:TabPanel ID="tbProdutividade" runat="server">
                                                <HeaderTemplate>
                                                    Produtividade
                                                </HeaderTemplate>
                                                <ContentTemplate>
                                                    <table class="actions" style="width: 100%;">
                                                        <tr>
                                                            <td class="iconNovo" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkGravarProdutividade" runat="server">
                                                                                            <span>Gravar</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconExcluir" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkExcluirProdutividade" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                                                                    <span>Excluir</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconRelatorio" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkRelatorioProdutividade" runat="server">
                                                                                            <span>Relatório</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td class="iconAjuda" style="width: 10%;">
                                                                <asp:LinkButton ID="lnkAjudaProdutividade" runat="server">
                                                                                            <span>Ajuda</span>
                                                                </asp:LinkButton>
                                                            </td>
                                                            <td style="width: 60%; display: block;">
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <table style="width: 100%;">
                                                        <tr>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Safra:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlSafra" runat="server" 
                                                                    Width="180px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cultura:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:DropDownList ID="ddlCultura" runat="server" 
                                                                    Width="200px">
                                                                </asp:DropDownList>
                                                            </td>
                                                            <td>
                                                                <div class="headerGray">
                                                                    <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Produtividade:</span>
                                                                    <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                                </div>
                                                            </td>
                                                            <td>
                                                                <asp:TextBox ID="TxtProdutividade" CssClass="txtDecimal" runat="server" Width="100px"
                                                                    Style="text-align: right;"/>
                                                            </td>
                                                        </tr>
                                                    </table>
                                                    <div class="bordasimples" style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                        <asp:GridView ID="grdVisitaProdutividade" runat="server" CellPadding="4" ForeColor="#333333"
                                                            GridLines="None" AutoGenerateColumns="False">
                                                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                            <EditRowStyle BackColor="#999999" />
                                                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px"
                                                                HorizontalAlign="Left" />
                                                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                            <Columns>
                                                                <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True">
                                                                    <ItemStyle Width="25px" />
                                                                </asp:CommandField>
                                                                <asp:BoundField DataField="CodigoSafra" HeaderText="Safra">
                                                                    <ItemStyle Width="200px" />
                                                                </asp:BoundField>
                                                                <asp:BoundField DataField="CodigoCultura" HeaderText="CodigoCultura">
                                                                    <HeaderStyle CssClass="Hide" />
                                                                    <ItemStyle CssClass="Hide" />
                                                                </asp:BoundField>
                                                                <asp:TemplateField HeaderText="Cultura">
                                                                    <ItemTemplate>
                                                                        <%# Eval("Cultura.Descricao")%>
                                                                    </ItemTemplate>
                                                                    <ItemStyle Width="200px" />
                                                                </asp:TemplateField>
                                                                <asp:BoundField DataField="Produtividade" HeaderText="Produtividade">
                                                                    <ItemStyle Width="100px" HorizontalAlign="Right" />
                                                                    <HeaderStyle HorizontalAlign="Right" />
                                                                </asp:BoundField>
                                                            </Columns>
                                                        </asp:GridView>
                                                    </div>
                                                </ContentTemplate>
                                            </ajaxToolkit:TabPanel>
                                        </ajaxToolkit:TabContainer>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                                <!-- Tab de Gerenciamento de Visitas -->
                                <ajaxToolkit:TabPanel ID="tbGerenciamentoVisita" runat="server" HeaderText="Gerenciamento de Visitas">
                                    <HeaderTemplate>
                                        Indicadores
                                    </HeaderTemplate>
                                    <ContentTemplate>
                                        <asp:Panel ID="Panel2" runat="server" Height="600px" Width="100%">
                                            <table width="100%">
                                                <tr>
                                                    <td width="100px">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Empresa:</span>
                                                            <span style="margin-left: 20px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td width="350px">
                                                        <asp:Label ID="lblIndiceEmpresa" runat="server" Style="font-size: 14px;"/>
                                                    </td>
                                                    <td width="100px">
                                                        <div class="headerGray">
                                                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Ano:</span>
                                                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                                                        </div>
                                                    </td>
                                                    <td width="30px">
                                                        <asp:Label ID="lblIndiceAno" runat="server" Style="font-size: 12px;"/>
                                                    </td>
                                                    <td class="iconConsultar">
                                                        <asp:LinkButton ID="lnkConsulta" runat="server">
                                                         <span>Consultar</span>
                                                        </asp:LinkButton>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td colspan="5">
                                                        <div style="min-height: 415px; max-height: 415px; overflow: auto;">
                                                            <asp:GridView ID="grdGerenciamentoVisita" runat="server" AutoGenerateColumns="False"
                                                                CellPadding="4" ForeColor="#333333" GridLines="None">
                                                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                                <EditRowStyle BackColor="#999999" />
                                                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                                <Columns>
                                                                    <asp:BoundField DataField="NomeCliente" HeaderText="Nome">
                                                                        <HeaderStyle HorizontalAlign="Left" />
                                                                        <ItemStyle HorizontalAlign="Left" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="MesAtual" HeaderText="Mês Atual">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="MinimoVisitas" HeaderText="Mínimo de Visitas">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes01" HeaderText="Mês 01">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes02" HeaderText="Mês 02">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes03" HeaderText="Mês 03">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes04" HeaderText="Mês 04">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes06" HeaderText="Mês 05">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes07" HeaderText="Mês 07">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes08" HeaderText="Mês 08">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes09" HeaderText="Mês 09">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes10" HeaderText="Mês 10">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes11" HeaderText="Mês 11">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="Mes12" HeaderText="Mês 12">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="TotalAno" HeaderText="Total Ano">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="SaldoDeVisitaObrigatorio" HeaderText="Saldo de V. Obrigatório">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="IPVMin" HeaderText="IPV Min.">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                    <asp:BoundField DataField="IPVAno" HeaderText="IPV Ano">
                                                                        <HeaderStyle HorizontalAlign="Right" />
                                                                        <ItemStyle HorizontalAlign="Right" />
                                                                    </asp:BoundField>
                                                                </Columns>
                                                            </asp:GridView>
                                                        </div>
                                                    </td>
                                                </tr>
                                                <tr>
                                                    <td>
                                                        &nbsp;&nbsp;
                                                    </td>
                                                    <td>
                                                        &nbsp;&nbsp;
                                                    </td>
                                                    <td>
                                                        &nbsp;&nbsp;
                                                    </td>
                                                </tr>
                                            </table>
                                        </asp:Panel>
                                    </ContentTemplate>
                                </ajaxToolkit:TabPanel>
                            </ajaxToolkit:TabContainer>
                        </td>
                    </tr>
                </table>
            </asp:Panel>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
