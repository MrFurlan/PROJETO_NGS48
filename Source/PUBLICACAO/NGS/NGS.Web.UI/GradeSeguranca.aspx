<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="GradeSeguranca.aspx.vb" Inherits="NGS.Web.UI.GradeSeguranca" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadNotaItens() {
            $("#txtDataMovimento").datepicker({
                dateFormat: 'dd/mm/yy',
                dayNames: ['Domingo', 'Segunda', 'Terça', 'Quarta', 'Quinta', 'Sexta', 'Sábado', 'Domingo'],
                dayNamesMin: ['D', 'S', 'T', 'Q', 'Q', 'S', 'S', 'D'],
                dayNamesShort: ['Dom', 'Seg', 'Ter', 'Qua', 'Qui', 'Sex', 'Sáb', 'Dom'],
                monthNames: ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho', 'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro'],
                monthNamesShort: ['Jan', 'Fev', 'Mar', 'Abr', 'Mai', 'Jun', 'Jul', 'Ago', 'Set', 'Out', 'Nov', 'Dez'],
                nextText: 'Próximo',
                prevText: 'Anterior',
                showOn: "button",
                buttonImage: "Images/calendar.png",
                buttonImageOnly: true
            });

            $("#txtDataMovimento").setMask("date");
        }

        $(document).ready(function () {
            pageLoadNotaItens();
            var prmNotaItens = Sys.WebForms.PageRequestManager.getInstance();
            prmNotaItens.add_endRequest(pageLoadNotaItens);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngGradeSeguranca" runat="server" AsyncPostBackTimeout="1000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlGradeSeguranca" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <asp:HiddenField ID="HAno" runat="server" />
            <div class="titulodiv">
                Grade De Segurança
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkConfirmar" Text="Cadastrar Novas Datas" runat="server" Enabled="false" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" AutoPostBack="True" Width="595px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataMovimento" ClientIDMode="Static" runat="server" OnTextChanged="txtDataMovimento_TextChanged"
                        AutoPostBack="true" />
                    <asp:CheckBox ID="chkTodas" runat="server" AutoPostBack="True" Text="Todas" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnLiberar" CssClass="btn" BackColor="Green" ForeColor="White" Height="26px"
                        runat="server" OnClick="btnLiberar_Click" Text="Liberar" UseSubmitBehavior="False" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="btnBloqueado" CssClass="btn" BackColor="Red" ForeColor="White" Height="26px"
                        runat="server" OnClick="btnBloqueado_Click" Text="Bloquear" UseSubmitBehavior="False" />
                </div>
            </div>
            <div class="row">
                <div class="painelleft" style="width: 32%; margin-right: 10px;">
                    <div class="bordagrid" style="height: 600px;">
                        <asp:GridView ID="GridMovimento" runat="server" Width="100%" OnSelectedIndexChanged="GridMovimento_SelectedIndexChanged"
                            AutoGenerateColumns="False" GridLines="None" ForeColor="#333333" CellPadding="4"
                            AllowSorting="True">
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <EditRowStyle BackColor="#999999" />
                            <Columns>
                                <asp:TemplateField FooterText="seleciona" ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:Button ID="Button2" runat="server" Text=">>" CommandName="Delete" CausesValidation="False"
                                            CssClass="btn" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Movimento_Id" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movimento"
                                    HtmlEncode="False" />
                                <asp:TemplateField FooterText="Situação" HeaderText="Situação" ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:Button ID="btnSituacao" runat="server" Width="90px" Text='<%# bind("Situacao") %>'
                                            Font-Bold="True" CommandName="Select" CausesValidation="False" CssClass="btn" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
                <div class="painelleft" style="width: 32%; margin-right: 10px;">
                    <div class="bordagrid" style="height: 600px;">
                        <asp:GridView ID="GridProcessos" runat="server" Width="100%" OnSelectedIndexChanged="GridProcessos_SelectedIndexChanged"
                            AutoGenerateColumns="False" GridLines="None" ForeColor="#333333" CellPadding="4">
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <EditRowStyle BackColor="#999999" />
                            <Columns>
                                <asp:TemplateField FooterText="seleciona" ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:Button ID="Button2" runat="server" Text=">>" CausesValidation="False" CommandName="Delete"
                                            CssClass="btn" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                                <asp:BoundField DataField="Processo_Id" HeaderText="Processo" FooterText="Processo"></asp:BoundField>
                                <asp:TemplateField FooterText="Situacao" HeaderText="Situacao" ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:Button ID="btnSituacao" runat="server" Width="90px" Text='<%# bind("Situacao") %>'
                                            CausesValidation="False" CommandName="Select" Font-Bold="True" CssClass="btn" />
                                        <asp:Label ID="Movimento_Id" runat="server" Text='<%# bind("Movimento_Id") %>' Visible="False" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
                <div class="painelleft" style="width: 32%;">
                    <div class="bordagrid" style="height: 600px;">
                        <asp:GridView ID="GridUsuarios" runat="server" Width="100%" OnSelectedIndexChanged="GridUsuarios_SelectedIndexChanged"
                            AutoGenerateColumns="False" GridLines="None" ForeColor="#333333" CellPadding="4">
                            <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                            <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                            <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                            <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                            <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                            <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                            <EditRowStyle BackColor="#999999" />
                            <Columns>
                                <asp:BoundField DataField="Usuario_Id" HeaderText="Nome" FooterText="Nome" />
                                <asp:TemplateField FooterText="Situacao" HeaderText="Situacao" ShowHeader="False">
                                    <ItemTemplate>
                                        <asp:Button ID="btnSituacao" runat="server" Width="90px" Text='<%# bind("Situacao") %>'
                                            CausesValidation="False" CommandName="Select" Font-Bold="True" CssClass="btn" />
                                        <asp:Label ID="Movimento" runat="server" Text='<%# bind("Movimento") %>' Visible="False" />
                                        <asp:Label ID="Processo" runat="server" Text='<%# bind("Processo") %>' Visible="False" />
                                        <asp:Label ID="Usuario_Id" runat="server" Text='<%# bind("Usuario_Id") %>' Visible="False" />
                                    </ItemTemplate>
                                </asp:TemplateField>
                            </Columns>
                        </asp:GridView>
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:InputDate ID="ucInputDate" runat="server" />
</asp:Content>
