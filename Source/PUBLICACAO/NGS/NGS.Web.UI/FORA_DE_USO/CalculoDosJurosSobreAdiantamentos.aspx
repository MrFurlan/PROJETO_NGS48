<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CalculoDosJurosSobreAdiantamentos.aspx.vb" Inherits="NGS.Web.UI.CalculoDosJurosSobreAdiantamentos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCalculoDosJurosSobreAdiantamentos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCalculoDosJurosSobreAdiantamentos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <table style="width: 100%; border: 0px none">
                <tr>
                    <td class="titulotabela" colspan="2">
                        <label>
                            Cálculo dos Juros Sobre Adiantamentos
                        </label>
                    </td>
                </tr>
                <tr>
                    <td colspan="2">
                        <table class="actions" style="width: 100%;">
                            <tr>
                                <td class="iconRelatorio" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkProcessar" runat="server" CssClass="lnkMenu">
                                        <span>Processar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconLimpar" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkLimpar" runat="server" CssClass="lnkMenu">
                                        <span>Limpar</span>
                                    </asp:LinkButton>
                                </td>
                                <td class="iconAjuda" runat="server" style="width: 10%;">
                                    <asp:LinkButton ID="lnkAjuda" runat="server" CssClass="lnkMenu">
                                        <span>Ajuda</span>
                                    </asp:LinkButton>
                                </td>
                                <td style="display: block;">
                                </td>
                            </tr>
                        </table>
                    </td>
                </tr>
                <tr>
                    <td style="width: 100px">
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Unidade:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="DdlUnidade" runat="server" Width="600px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                            AutoPostBack="True" Font-Names="Courier New" Font-Size="8pt">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Empresa:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:DropDownList ID="DdlEmpresa" runat="server" Width="600px" Font-Names="Courier New"
                            Font-Size="8pt">
                        </asp:DropDownList>
                    </td>
                </tr>
                <tr>
                    <td>
                        <div class="headerGray">
                            <span class="HeaderSpanFirstBlue"></span><span class="HeaderSpanSecond">Cliente:</span>
                            <span style="margin-left: 71px; margin-top: 1px; margin-right: 0px;"></span>
                        </div>
                    </td>
                    <td>
                        <asp:TextBox ID="txtCliente" runat="server" Width="560px"></asp:TextBox>
                        <asp:Button ID="cmdCliente" runat="server" Text=" > " OnClick="cmdCliente_Click">
                        </asp:Button>
                    </td>
                </tr>
                <tr>
                    <td style="text-align: right;">
                        <label>
                            Até a Data
                        </label>
                    </td>
                    <td>
                        <asp:Calendar ID="CalData" runat="server" Width="350px" Height="190px" Font-Names="Verdana"
                            Font-Size="9pt" BorderWidth="1px" BorderColor="White" BackColor="White" ForeColor="Black"
                            NextPrevFormat="FullMonth">
                            <SelectedDayStyle BackColor="#333399" ForeColor="White"></SelectedDayStyle>
                            <TodayDayStyle BackColor="#CCCCCC"></TodayDayStyle>
                            <OtherMonthDayStyle ForeColor="#999999"></OtherMonthDayStyle>
                            <NextPrevStyle VerticalAlign="Bottom" Font-Bold="True" Font-Size="8pt" ForeColor="#333333">
                            </NextPrevStyle>
                            <DayHeaderStyle Font-Bold="True" Font-Size="8pt"></DayHeaderStyle>
                            <TitleStyle BackColor="White" BorderColor="Black" BorderWidth="4px" Font-Bold="True"
                                Font-Size="12pt" ForeColor="#333399"></TitleStyle>
                        </asp:Calendar>
                    </td>
                </tr>
            </table>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
