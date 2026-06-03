<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CotacoesDeMoedas.aspx.vb" Inherits="NGS.Web.UI.CotacoesDeMoedas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngCotacoesDeMoedas" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCotacoesDeMoedas" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Cotações de Moedas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server"
                                 />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server"
                                 />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnMesAnterior" OnClick="btnMesAnterior_Click" runat="server" Text=" < "
                        CssClass="btn" UseSubmitBehavior="False"></asp:Button>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlDia" runat="server" Width="71px" OnSelectedIndexChanged="ddlDia_SelectedIndexChanged"
                        AutoPostBack="True">
                        <asp:ListItem>1</asp:ListItem>
                        <asp:ListItem>2</asp:ListItem>
                        <asp:ListItem>3</asp:ListItem>
                        <asp:ListItem>4</asp:ListItem>
                        <asp:ListItem>5</asp:ListItem>
                        <asp:ListItem>6</asp:ListItem>
                        <asp:ListItem>7</asp:ListItem>
                        <asp:ListItem>8</asp:ListItem>
                        <asp:ListItem>9</asp:ListItem>
                        <asp:ListItem>10</asp:ListItem>
                        <asp:ListItem>11</asp:ListItem>
                        <asp:ListItem>12</asp:ListItem>
                        <asp:ListItem>13</asp:ListItem>
                        <asp:ListItem>14</asp:ListItem>
                        <asp:ListItem>15</asp:ListItem>
                        <asp:ListItem>16</asp:ListItem>
                        <asp:ListItem>17</asp:ListItem>
                        <asp:ListItem>18</asp:ListItem>
                        <asp:ListItem>19</asp:ListItem>
                        <asp:ListItem>20</asp:ListItem>
                        <asp:ListItem>21</asp:ListItem>
                        <asp:ListItem>22</asp:ListItem>
                        <asp:ListItem>23</asp:ListItem>
                        <asp:ListItem>24</asp:ListItem>
                        <asp:ListItem>25</asp:ListItem>
                        <asp:ListItem>26</asp:ListItem>
                        <asp:ListItem>27</asp:ListItem>
                        <asp:ListItem>28</asp:ListItem>
                        <asp:ListItem>29</asp:ListItem>
                        <asp:ListItem>30</asp:ListItem>
                        <asp:ListItem>31</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMes" runat="server" Width="109px" OnSelectedIndexChanged="ddlMes_SelectedIndexChanged"
                        AutoPostBack="True">
                        <asp:ListItem Value="1">Janeiro</asp:ListItem>
                        <asp:ListItem Value="2">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="3">Marco</asp:ListItem>
                        <asp:ListItem Value="4">Abril</asp:ListItem>
                        <asp:ListItem Value="5">Maio</asp:ListItem>
                        <asp:ListItem Value="6">Junho</asp:ListItem>
                        <asp:ListItem Value="7">Julho</asp:ListItem>
                        <asp:ListItem Value="8">Agosto</asp:ListItem>
                        <asp:ListItem Value="9">Setembro</asp:ListItem>
                        <asp:ListItem Value="10">Outubro</asp:ListItem>
                        <asp:ListItem Value="11">Novembro</asp:ListItem>
                        <asp:ListItem Value="12">Dezembro</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAno" runat="server" Width="76px" OnSelectedIndexChanged="ddlAno_SelectedIndexChanged"
                        AutoPostBack="True">
                        <asp:ListItem>2008</asp:ListItem>
                        <asp:ListItem>2009</asp:ListItem>
                        <asp:ListItem>2010</asp:ListItem>
                        <asp:ListItem>2011</asp:ListItem>
                        <asp:ListItem>2012</asp:ListItem>
                        <asp:ListItem>2013</asp:ListItem>
                        <asp:ListItem>2014</asp:ListItem>
                        <asp:ListItem>2015</asp:ListItem>
                        <asp:ListItem>2016</asp:ListItem>
                        <asp:ListItem>2017</asp:ListItem>
                        <asp:ListItem>2018</asp:ListItem>
                        <asp:ListItem>2019</asp:ListItem>
                        <asp:ListItem>2020</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnProxMes" OnClick="btnProxMes_Click" runat="server" Text=" > "
                        CssClass="btn" UseSubmitBehavior="False"></asp:Button>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Indexador:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlIndexador" runat="server" Width="324px" OnSelectedIndexChanged="ddlIndexador_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Valor:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtIndice" CssClass="txtDecimal8" runat="server" data-ToolTip="default"
                        ToolTip="Valor da moeda no dia." />
                </div>
                <div class="coltxt">
                    <asp:ImageButton ID="imgCotacao" runat="server" Height="22px" ImageAlign="AbsMiddle"
                        ImageUrl="~/App_Themes/ngssolucoes/imagens/icon_consultar.png" Width="22px" BorderStyle="None" />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdPrevisto" runat="server" Text="Previsto" AutoPostBack="True"
                        Checked="True" GroupName="dolar" />
                    <asp:RadioButton ID="rdRealizado" runat="server" Text="Realizado" AutoPostBack="True"
                        GroupName="dolar" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt" style="width: 296px;">
                    <strong>LIBERADO POR:</strong>
                    <asp:Label ID="LblLiberado" runat="server" Font-Bold="True" ForeColor="SeaGreen" />
                </div>
                <div class="coltxt">
                    <strong>ALTERADO POR: </strong>
                    <asp:Label ID="LblAlterado" runat="server" Font-Bold="True" ForeColor="Red" />
                </div>
            </div>
            <div class="bordagrid" style="width: 946px;">
                <asp:GridView ID="gridCotacoesDeMoedas" runat="server" AutoGenerateColumns="False"
                    CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <%--<asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />--%>
                        <asp:BoundField DataField="Indexador_id" HeaderText="ID" />
                        <asp:BoundField DataField="Descricao" HeaderText="Indexador"></asp:BoundField>
                        <asp:BoundField DataField="Dia01" HeaderText="Dia 01" />
                        <asp:BoundField DataField="Dia02" HeaderText="Dia 02" />
                        <asp:BoundField DataField="Dia03" HeaderText="Dia 03" />
                        <asp:BoundField DataField="Dia04" HeaderText="Dia 04" />
                        <asp:BoundField DataField="Dia05" HeaderText="Dia 05" />
                        <asp:BoundField DataField="Dia06" HeaderText="Dia 06" />
                        <asp:BoundField DataField="Dia07" HeaderText="Dia 07" />
                        <asp:BoundField DataField="Dia08" HeaderText="Dia 08" />
                        <asp:BoundField DataField="Dia09" HeaderText="Dia 09" />
                        <asp:BoundField DataField="Dia10" HeaderText="Dia 10" />
                        <asp:BoundField DataField="Dia11" HeaderText="Dia 11" />
                        <asp:BoundField DataField="Dia12" HeaderText="Dia 12" />
                        <asp:BoundField DataField="Dia13" HeaderText="Dia 13" />
                        <asp:BoundField DataField="Dia14" HeaderText="Dia 14" />
                        <asp:BoundField DataField="Dia15" HeaderText="Dia 15" />
                        <asp:BoundField DataField="Dia16" HeaderText="Dia 16" />
                        <asp:BoundField DataField="Dia17" HeaderText="Dia 17" />
                        <asp:BoundField DataField="Dia18" HeaderText="Dia 18" />
                        <asp:BoundField DataField="Dia19" HeaderText="Dia 19" />
                        <asp:BoundField DataField="Dia20" HeaderText="Dia 20" />
                        <asp:BoundField DataField="Dia21" HeaderText="Dia 21" />
                        <asp:BoundField DataField="Dia22" HeaderText="Dia 22" />
                        <asp:BoundField DataField="Dia23" HeaderText="Dia 23" />
                        <asp:BoundField DataField="Dia24" HeaderText="Dia 24" />
                        <asp:BoundField DataField="Dia25" HeaderText="Dia 25" />
                        <asp:BoundField DataField="Dia26" HeaderText="Dia 26" />
                        <asp:BoundField DataField="Dia27" HeaderText="Dia 27" />
                        <asp:BoundField DataField="Dia28" HeaderText="Dia 28" />
                        <asp:BoundField DataField="Dia29" HeaderText="Dia 29" />
                        <asp:BoundField DataField="Dia30" HeaderText="Dia 30" />
                        <asp:BoundField DataField="Dia31" HeaderText="Dia 31" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
