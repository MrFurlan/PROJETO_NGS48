<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucConsultaClientes.ascx.vb"
    Inherits="NGS.Web.UI.ucConsultaClientes" %>
<script type="text/javascript">
    function pageLoadConsultaCliente() {

        function consulta() {
            var btn = document.getElementById('<%=lnkConsultar.ClientID%>');
            btn.click();
        }

        $("#MainContent_ucConsultaClientes_txtNome", "#divConsultaCliente").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaClientes_txtCodigo", "#divConsultaCliente").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaClientes_txtFantasia", "#divConsultaCliente").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaClientes_txtCidade", "#divConsultaCliente").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaClientes_txtEstado", "#divConsultaCliente").keydown(function (event) {
            if (event.which == 13) {
                event.preventDefault();
                consulta();
            }
        });

        $("#MainContent_ucConsultaClientes_txtNome", "#divConsultaCliente").focus();
    }

    function verificaSituacao(lnkButton) {
        var lblSit = $(lnkButton).parent().parent().find(".lbl");
        var lblDesc = $(lnkButton).parent().parent().find(".lblDesc");
        if (lblSit.text() == "50" && lblSit.text() != undefined) {
            //return
            //msgbox('ATENÇÃO:\n\nA Situação do cliente é: ' + lblSit.text() + '-' + lblDesc.text(), "ATENÇÃO!", "Info");
            alert('ATENÇÃO:\n\nA Situação do cliente é: ' + lblSit.text() + '-' + lblDesc.text());
            return true;
        }
        else if (lblSit.text() != "1" && lblSit.text() != undefined) {
            //return
            //msgbox('ATENÇÃO:\n\nA Situação do cliente é: ' + lblSit.text() + '-' + lblDesc.text(), "ATENÇÃO!", "Info");
            alert('ATENÇÃO:\n\nA Situação do cliente é: ' + lblSit.text() + '-' + lblDesc.text());
            return false;
        }
        return true;
    };

    $(document).ready(function () {
        pageLoadConsultaCliente();
    });

    var prmConsultaCliente = Sys.WebForms.PageRequestManager.getInstance();
    prmConsultaCliente.add_endRequest(pageLoadConsultaCliente);
</script>
<div id="divConsultaCliente" class="uc" title="Consulta de Clientes" style="display: none;">
    <asp:UpdatePanel ID="updpnlConsultaClientes" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconConsultar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkConsultar" runat="server" OnClick="lnkConsultar_click">
                                <span>Consultar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconLimpar" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkLimpar" runat="server">
                                <span>Limpar</span>
                            </asp:LinkButton>
                        </li>
                        <li class="iconSair" style="width: 10%;" runat="server">
                            <asp:LinkButton ID="lnkFechar" runat="server">
                                <span>Fechar</span>
                            </asp:LinkButton>
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cpf/Cnpj:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" runat="server" CausesValidation="false" CssClass="txtNumerico" placeholder="CPF/CNPJ" data-tooltip="default" ToolTip="Código do cliente é caracterizado por cpf(quando pessoa física), ou cnpj(quando pessoa jurídica)." />
                    <asp:HiddenField ID="hdnTipoCliente" runat="server" />
                </div>
                <div class="collbluc">
                    Situação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlSituacao" runat="server" Width="338px" data-tooltip="default"
                        ToolTip="Situação do cliente no sistema." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Nome:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtNome" runat="server" Width="462px" data-tooltip="default" ToolTip="Nome de registro do cliente/empresa." />
                </div>
                <div class="coltxt">
                    <asp:CheckBox ID="chkParteNome" runat="server" Font-Bold="True" Text="Parte do nome."
                        data-tooltip="default" ToolTip="Procurar por palavra chave." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Fantasia:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtFantasia" runat="server" Width="583px" data-tooltip="default"
                        ToolTip="Nome de divulgação da cliente/empresa." />
                </div>
            </div>
            <div class="row">
                <div class="collbluc">
                    Cidade/Estado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCidade" runat="server" Width="520px" data-tooltip="default" ToolTip="Localização municipal do cliente." />
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtEstado" runat="server" MaxLength="2" Width="50px" data-tooltip="default"
                        ToolTip="Localização estadual do cliente." />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridClientes" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:TemplateField>
                            <ItemTemplate>
                                <asp:LinkButton ID="lnkSelecionar" CssClass="lnk" OnClientClick="return verificaSituacao(this);"
                                    data-tooltip="default" ToolTip="Selecionar registro." runat="server" Text=" &gt; "
                                    OnClick="lnkSelecionar_Click">
                                    <i class="fa fa-arrow-right seta"></i>
                                </asp:LinkButton>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Nome" HeaderText="Nome" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Left" Width="40%" />
                            <ItemStyle HorizontalAlign="left" Width="40%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Complemento" HeaderText="Complemento" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Left" Width="10%" />
                            <ItemStyle HorizontalAlign="left" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Cidade" HeaderText="Cidade" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Left" Width="10%" />
                            <ItemStyle HorizontalAlign="left" Width="10%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Estado" HeaderText="UF" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Center" Width="5%" />
                            <ItemStyle HorizontalAlign="Center" Width="5%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Codigo" HeaderText="CNPJ/CPF" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Right" Width="12%" />
                            <ItemStyle HorizontalAlign="Right" Width="12%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Endereco_Id" HeaderText="End" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="center" Width="3%" />
                            <ItemStyle HorizontalAlign="center" Width="3%" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Inscricao" HeaderText="Inscrição" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Right" Width="10%" />
                            <ItemStyle HorizontalAlign="Right" Width="10%" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Situação" HeaderStyle-CssClass="none" ItemStyle-CssClass="none">
                            <ItemTemplate>
                                <asp:Label ID="lblSituacao" CssClass="lbl" runat="server" Text='<%# Eval("Situacao_Id") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Situação">
                            <ItemTemplate>
                                <asp:Label ID="lblDescricao" CssClass="lblDesc" runat="server" Text='<%# Eval("DescSituacao") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
