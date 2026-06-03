<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="PosicaoDeEncargos.aspx.vb" Inherits="NGS.Web.UI.PosicaoDeEncargos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        function pageLoadPosicaoDeEncargos() {
            verificachk();

            $("#<%=chkTodosEncargos.ClientID%>").click(function () {
                verificachk();
            });

            $("#<%=lstEmpresa.ClientID%>").multiselect({
                header: "Escolha as empresas!",
                selectedList: 1
            }).multiselectfilter();

            function verificachk() {
                var link = $("#<%=lnkEmiteTodosEncargos.ClientID%>")
                if ($("#<%=chkTodosEncargos.ClientID%>").attr('checked') == "checked")
                    link.parent().show();
                else
                    link.parent().hide();
            }
        };

        $(document).ready(function () {
            pageLoadPosicaoDeEncargos();
            var prmPosicaoDeEncargos = Sys.WebForms.PageRequestManager.getInstance();
            prmPosicaoDeEncargos.add_endRequest(pageLoadPosicaoDeEncargos);
        });
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngPosicaoDeEncargos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlPosicaoDeEncargos" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Posição de Encargos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio" runat="server" visible="true">
                            <asp:LinkButton ID="lnkEmiteTodosEncargos" Text="Emitir Encargos" runat="server" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged1"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa
                </div>
                <div class="coltxt">
                    <asp:ListBox ID="lstEmpresa" Width="634px" runat="server" CssClass="multiselect" SelectionMode="Multiple" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="595px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="BtnCliente" runat="server" OnClick="BtnCliente_Click" CssClass="btn"
                        Text=">" UseSubmitBehavior="False" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px"
                        data-ToolTip="default" ToolTip="Informar a data inicial a final da consulta." />
                </div>
                <div class="collbl">
                    à:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px"
                        data-ToolTip="default" ToolTip="Informar a data inicial a final da consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Notas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadEntradas" runat="server" Text=" Entradas " Checked="True"
                        GroupName="Notas" Style="margin-right: 5px;" data-ToolTip="default" ToolTip="Informar se é nota de entrada ou saída." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadSaidas" runat="server" Text=" Saidas " GroupName="Notas"
                        Style="margin-right: 10px;" data-ToolTip="default" ToolTip="Informar se é nota de entrada ou saída." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="ChkAgrupar" runat="server" Text="Agrupar por Clientes" Checked="True"
                        Style="margin-right: 5px;" />
                    <asp:CheckBox ID="chkResumoEncargo" runat="server" Text="Resumo Encargos" />
                    <asp:CheckBox ID="chkTodosEncargos" runat="server" Text="Emitir todos encargos no período" />
                    <%--<input type="checkbox" onchange
                    --%>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Percentual %:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtPercentual" Text="" ToolTip="Caso queira filtrar por um determinado % informe o valor aqui" CssClass="txtDecimal" runat="server"></asp:TextBox>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Relatório:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadPDF" runat="server" Checked="True" GroupName="relatorio"
                        Text="PDF" data-ToolTip="default" ToolTip="Selecionar o formato de impressão do relatório." />
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadExcel" runat="server" GroupName="relatorio" Text="Excel"
                        data-ToolTip="default" ToolTip="Selecionar o formato de impressão do relatório." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridOpcoes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridOpcoes_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                            <ItemStyle Width="25px"></ItemStyle>
                        </asp:CommandField>
                        <asp:BoundField DataField="Codigo" HeaderText="C&#243;digo">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                            <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Left"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
