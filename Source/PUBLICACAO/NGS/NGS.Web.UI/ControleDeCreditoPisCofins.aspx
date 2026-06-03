<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ControleDeCreditoPisCofins.aspx.vb" Inherits="NGS.Web.UI.ControleDeCreditoPisCofins" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl
        {
            width: 150px;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngControleDeCreditoPisCofins" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlControleDeCreditoPisCofins" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Controle de Crédito Pis/Cofins
            </div>
            <div class="row">
                <div class="collbl">
                    <lable>Referente Mês/Ano:</lable>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtMesAnoRef" CssClass="calendario" runat="server" Width="79px" />
                </div>
                <div class="collbl">
                    <lable>Registro:</lable>
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPis1100" runat="server" GroupName="Registro" Text="Pis 1100" />
                    <asp:RadioButton ID="rdCofins1500" runat="server" GroupName="Registro" Text="Cofins 1500" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <lable>Empresa Matriz:</lable>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresaMatriz" runat="server" Width="590px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <lable>Per_Apu_Cred:</lable>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataApuracaoCredito" CssClass="txtmesano" runat="server" Width="100px"
                        data-tooltip="default" ToolTip="Período de Apuração do Crédito (Mês/Ano)" />
                </div>
                <div class="collbl">
                    <lable>ORIG_CRED:</lable>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlOrigemCredito" runat="server" Width="314px" data-tooltip="default"
                        ToolTip="Indicador da origem do crédito: 01 – Crédito decorrente de operações próprias; 
                        02 – Crédito transferido por pessoa jurídica sucedida.">
                        <asp:ListItem Value=""></asp:ListItem>
                        <asp:ListItem Value="1">01 – Crédito decorrente de operações próprias</asp:ListItem>
                        <asp:ListItem Value="2">02 – Crédito transferido por pessoa jurídica sucedida.</asp:ListItem>
                    </asp:DropDownList>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <lable>COD_CRED:</lable>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlCodigoCredito" runat="server" Width="590px" data-tooltip="default"
                        ToolTip="Código do Tipo do Crédito." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    VL_CRED_APU:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCreditoApurado" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor total do crédito apurado na Escrituração Fiscal Digital (Registro M100) ou em 
                        Demonstrativo DACON (fichas 06A e 06B) de período anterior." />
                </div>
                <div class="collbl">
                    VL_CRED_EXT_APU:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCreditoExtApurado" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor de Crédito Extemporâneo Apurado (Registro 1101), referente a Período Anterior, Informado no Campo 02 – PER_APU_CRED" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    VL_CRED_DESC_PA_ANT:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCreditoDescontoAnterior" runat="server" CssClass="txtDecimal"
                        Width="100px" data-tooltip="default" ToolTip="Valor do Crédito utilizado mediante Desconto, em Período(s) Anterior(es)." />
                </div>
                <div class="collbl">
                    VL_CRED_DCOMP_PA_ANT:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCreditoCompraAnterior" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor do Crédito utilizado mediante Declaração de Compensação Intermediária 
                        (Crédito de Exportação), em Período(s) Anterior(es)." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cr_Dec_Interm_Pa_Ant" runat="server" CssClass="txtDecimal"
                        Width="100px" />
                </div>
                <div class="collbl">
                    VL_CRED_PER_EFD
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cr_Objeto_Per_Efd" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor do Crédito objeto de Pedido de Ressarcimento (PER) neste período de escrituração." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    VL_CRED_DESC_EFD:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cred_Desc_Efd" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor do Crédito descontado neste período de escrituração." />
                </div>
                <div class="collbl">
                    VL_CRED_DCOMP_EFD:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cr_Dec_Interm_Efd" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor do Crédito utilizado mediante Declaração de Compensação Intermediária neste período de escrituração." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    VL_CRED_TRANS:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cr_Transf_Cisao_Fusao" runat="server" CssClass="txtDecimal"
                        Width="100px" data-tooltip="default" ToolTip="Valor do crédito transferido em evento de cisão, fusão ou incorporação." />
                </div>
                <div class="collbl">
                    VL_CRED_OUT:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtVl_Cr_Outras_Formas" runat="server" CssClass="txtDecimal" Width="100px"
                        data-tooltip="default" ToolTip="Valor do crédito utilizado por outras formas" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="gridApuracao" runat="server" CellPadding="4" AutoGenerateColumns="False"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:BoundField DataField="Per_Apu_Cred" HeaderText="Per_Apu_Cred" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" Width="100px" />
                            <ItemStyle HorizontalAlign="Left" Width="100px" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Orig_Cred" HeaderText="Orig_Cred" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cnpj_Suc" HeaderText="Cnpj_Suc" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cod_Cred" HeaderText="Cod_Cred" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Apu" HeaderText="Vl_Cred_Apu" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Ext_Apu" HeaderText="Vl_Cred_Ext_Apu" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Tot_Cred_Apu" HeaderText="Vl_Tot_Cred_Apu" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Desc_Pa_Ant" HeaderText="Vl_Cred_Desc_Pa_Ant"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Dec" HeaderText="Vl_Cred_Dec" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Sd_Cred_Disp_Efd" HeaderText="Sd_Cred_Disp_Efd" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Desc_Efd" HeaderText="Vl_Cred_Desc_Efd" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred" HeaderText="Vl_Cred" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Dec" HeaderText="Vl_Cred_Dec" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Transf" HeaderText="Vl_Cred_Transf" HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Vl_Cred_Outras_Formas" HeaderText="Vl_Cred_Outras_Formas"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Sld_Creditos_ Utilizar Futuro" HeaderText="Sld_Creditos_ Utilizar Futuro"
                            HtmlEncode="False">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
