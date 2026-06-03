<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ConciliacaoDeContas.aspx.vb" AsyncTimeout="5000" EnableEventValidation="false"
    Inherits="NGS.Web.UI.ConciliacaoDeContas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngConciliacaoDeContas" runat="server" AsyncPostBackTimeout="5000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updPnlConciliacaoDeContas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Conciliação de Contas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server" class="iconConsultar">
                            <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
                        </li>
                        <li runat="server" class="iconRelatorio">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server" class="iconConfirmar">
                            <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
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
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="668px" AutoPostBack="True" OnSelectedIndexChanged="ddlUnidade_SelectedIndexChanged" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlEmpresa" runat="server" Width="668px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Conta:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="hdfConta" runat="server" />
                    <asp:TextBox ID="txtConta" runat="server" Width="630px" ReadOnly="true" data-ToolTip="default"
                        ToolTip="Selecionar a conta desejada." />
                    <asp:Button ID="btnConta" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar a conta desejada." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="630px" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                    <asp:Button ID="btnCliente" runat="server" Text=">" CssClass="btn" data-ToolTip="default"
                        ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    C. Custo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlCentroDeCusto" runat="server" Width="668px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Grupo:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlGrupo" runat="server" Width="668px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Produto:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlProduto" runat="server" Width="668px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Data:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" runat="server" Width="112px" CssClass="calendario"
                        data-ToolTip="default" ToolTip="Período Inicial dos lançamentos para conciliação." />
                </div>
                <div class="collbl">
                    à
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="120px"
                        data-ToolTip="default" ToolTip="Período final dos lançamentos para conciliação." />
                </div>
                <div class="collbl">
                    Lançamentos:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdAberto" runat="server" Text="Abertos" Checked="True" GroupName="tiposel"
                        data-ToolTip="default" ToolTip="Selecionar se os lançamentos estiverem abertos ou baixados." />
                    <asp:RadioButton ID="rdBaixados" runat="server" Text="Baixados" GroupName="tiposel"
                        data-ToolTip="default" ToolTip="Selecionar se os lançamentos estiverem abertos ou baixados." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="grdConciliacao" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"></asp:CommandField>
                        <asp:BoundField DataField="Movimento_Id" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Movto"></asp:BoundField>
                        <asp:BoundField DataField="Lote_Id" HeaderText="Lote"></asp:BoundField>
                        <asp:BoundField DataField="Sequencia_Id" HeaderText="Seq."></asp:BoundField>
                        <asp:BoundField DataField="TITULO" HeaderText="Titulo"></asp:BoundField>
                        <asp:TemplateField HeaderText="C.B.">
                            <ItemTemplate>
                                <asp:CheckBox ID="ChkConciliacao" runat="server" Checked='<%# Eval("consciliacao") = "B"%>'></asp:CheckBox>
                                <asp:HiddenField ID="HiddenField1" runat="server" Value='<%# Eval("Cliente_id")%>' />
                                <asp:HiddenField ID="HiddenField2" runat="server" Value='<%# Eval("EndCliente_id") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Baixa">
                            <ItemTemplate>
                                <asp:TextBox ID="txtBaixa" CssClass="calendario" runat="server" Width="80px" Text='<%# Eval("DataDaBaixa", "{0:dd/MM/yyyy}")%>' />
                            </ItemTemplate>
                            <ItemStyle Wrap="false" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Hist&#243;rico">
                            <ItemTemplate>
                                <div style="max-width: 400px;">
                                    <asp:Label ID="lblHistorico" runat="server" Width="400px" Text='<%# Eval("HISTORICO") %>' />
                                </div>
                            </ItemTemplate>
                            <ItemStyle Wrap="true" Width="400px" />
                        </asp:TemplateField>
                        <asp:BoundField DataField="DEBITO" DataFormatString="{0:N}" HeaderText="D&#233;bito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="CREDITO" DataFormatString="{0:N}" HeaderText="Cr&#233;dito">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                        <asp:BoundField DataField="SALDO" DataFormatString="{0:N}" HeaderText="Saldo">
                            <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                            <ItemStyle HorizontalAlign="Right"></ItemStyle>
                        </asp:BoundField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaPlanoDeContas ID="ucConsultaPlanoDeContas" runat="server" />
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
