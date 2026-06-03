<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="BaixasFinanceirasIndividuais.aspx.vb" Inherits="NGS.Web.UI.BaixasFinanceirasIndividuais" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngBaixasFinanceirasIndividuais" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <div class="titulodiv">
        Baixas Financeiras Individuais (Bradesco)
    </div>
    <asp:UpdatePanel ID="updpnlBaixasFinanceirasIndividuais" runat="server">
        <ContentTemplate>
            <ajaxToolkit:TabContainer ID="TabContainer1" runat="server" Width="100%" ActiveTabIndex="0">
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel4" ID="TabPanel4">
                    <HeaderTemplate>
                        Titulos
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultar" Text="Consultar" runat="server" />
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
                                <asp:DropDownList ID="DdlUnidadeConsultaTitulos" runat="server" Width="598px" AutoPostBack="True"
                                    OnSelectedIndexChanged="DdlUnidadeConsultaTitulos_SelectedIndexChanged" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Empresa:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlEmpresaConsultaTitulos" runat="server" Width="598px" />
                            </div>
                            <div class="coltxt" style="float: right;">
                                <asp:Button ID="cmdAgrupar" OnClick="cmdAgrupar_Click" runat="server" Width="110px"
                                    Text="Baixa Financeira" class="botao" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cliente:
                            </div>
                            <div class="coltxt">
                                <asp:DropDownList ID="DdlClienteConsultaTitulos" runat="server" Width="568px" />
                            </div>
                            <div class="coltxt">
                                <asp:Button ID="cmdConsultaClientes" OnClick="cmdConsultaClientes_Click" runat="server"
                                    CssClass="btn" Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                            </div>
                            <div class="coltxt" style="float: right;">
                                <asp:Button ID="Btn_RetaEnviar" OnClick="Btn_RetaEnviar_Click" runat="server" Width="110px"
                                    Height="24px" Text="Marca a Enviar" class="botao" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Periodo Inicial:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoInicialConsultaTitulos" CssClass="calendario" runat="server"
                                    Width="116px" data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                            </div>
                            <div class="collbl">
                                Periodo Final:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtPeriodoFinalConsultaTitulos" CssClass="calendario" runat="server"
                                    Width="108px" CausesValidation="True" data-ToolTip="default" ToolTip="Informar o data final de consulta." />
                            </div>
                            <div class="coltxt" style="float: right;">
                                <asp:Button ID="Btn_retnaoproc" OnClick="Btn_retnaoproc_Click" runat="server" Width="110px"
                                    Height="24px" Text="Ret. Não Proc" class="botao" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Opção:
                            </div>
                            <div class="coltxt">
                                <asp:RadioButton ID="RbGeral" runat="server" OnCheckedChanged="RbGeral_CheckedChanged"
                                    Checked="True" GroupName="TipoRel" Text="Não Processado" data-ToolTip="default"
                                    ToolTip="Selecionar a opção para gerar o relatório." />
                                <asp:RadioButton ID="RbAEnviar" runat="server" OnCheckedChanged="RbAEnviar_CheckedChanged"
                                    GroupName="TipoRel" Text="A Enviar" data-ToolTip="default" ToolTip="Selecionar a opção para gerar o relatório." />
                                <asp:RadioButton ID="RbEnviado" runat="server" OnCheckedChanged="RbEnviado_CheckedChanged"
                                    GroupName="TipoRel" Text="Enviado" data-ToolTip="default" ToolTip="Selecionar a opção para gerar o relatório." />
                                <asp:RadioButton ID="RbBaixado" runat="server" OnCheckedChanged="RbBaixado_CheckedChanged"
                                    GroupName="TipoRel" Text="Baixado" data-ToolTip="default" ToolTip="Selecionar a opção para gerar o relatório." />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridConsultaTitulos" runat="server" AutoGenerateColumns="False"
                                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" EnableTheming="False"
                                OnSelectedIndexChanged="GridConsultaTitulos_SelectedIndexChanged">
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                                <SelectedRowStyle BackColor="#C5BBAF" Font-Bold="True" ForeColor="#333333" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <EditRowStyle BackColor="#999999" />
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button"
                                        HeaderText="   &gt; ">
                                        <ItemStyle Width="30px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Registro" HeaderText="Registro">
                                        <ItemStyle Width="60px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Vencimento" DataFormatString="{0:dd/MM/yyyy}" HeaderText="Vencimento"
                                        HtmlEncode="False">
                                        <ItemStyle Width="80px"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cliente" HeaderText="Cliente">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Historico" HeaderText="Historico">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Valor" DataFormatString="{0:N}" HeaderText="Valor" HtmlEncode="False">
                                        <HeaderStyle HorizontalAlign="Right"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Right"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:TemplateField HeaderText=" Lb">
                                        <ItemTemplate>
                                            <asp:CheckBox ID="ChkGridTitulos" runat="server" />
                                        </ItemTemplate>
                                    </asp:TemplateField>
                                    <asp:BoundField DataField="UsuarioLiberacao" HeaderText="Situa&#231;&#227;o Bancaria">
                                        <ItemStyle Width="0px"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel5" ID="TabPanel5">
                    <HeaderTemplate>
                        Consulta Clientes
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconConsultar" runat="server">
                                        <asp:LinkButton ID="lnkConsultarCC" Text="Consultar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparCC" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudarCC" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cnpj/CPF:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCnpjConsulta" runat="server" Width="118px" />
                            </div>
                            <div class="coltxt">
                                <asp:Label ID="lblPainel" runat="server" Width="97px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Nome:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtNomeConsulta" runat="server" Width="378px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Fantasia:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtFantasiaConsulta" runat="server" Width="378px" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Cidade:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtCidadeConsulta" runat="server" Width="333px" />
                                <asp:TextBox ID="txtEstadoConsulta" runat="server" Width="31px" />
                            </div>
                        </div>
                        <div class="bordagrid">
                            <asp:GridView ID="GridClientes" runat="server" Width="100%" Height="21px" BorderStyle="None"
                                OnSelectedIndexChanged="GridClientes_SelectedIndexChanged" CellPadding="1" AutoGenerateColumns="False"
                                BackColor="White" BorderColor="#CCCCCC" BorderWidth="1px">
                                <FooterStyle BackColor="White" ForeColor="#000066"></FooterStyle>
                                <RowStyle ForeColor="#000066"></RowStyle>
                                <Columns>
                                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                                        <ItemStyle Height="19px" Width="19px"></ItemStyle>
                                    </asp:CommandField>
                                    <asp:BoundField DataField="Nome" HeaderText="Nome ">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Estado" HeaderText="UF">
                                        <HeaderStyle HorizontalAlign="Center"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Center"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Codigo" HeaderText="CNPJ/CPF">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                    <asp:BoundField DataField="Endereco_Id" HeaderText="End">
                                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                                    </asp:BoundField>
                                </Columns>
                                <PagerStyle HorizontalAlign="Left" BackColor="White" ForeColor="#000066"></PagerStyle>
                                <SelectedRowStyle BackColor="#006699" Font-Bold="True" ForeColor="White"></SelectedRowStyle>
                                <HeaderStyle BackColor="#006699" CssClass="DataGridFixedHeader" Font-Bold="True"
                                    ForeColor="White"></HeaderStyle>
                            </asp:GridView>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel runat="server" HeaderText="TabPanel1" ID="TabPanel1">
                    <HeaderTemplate>
                        Observações
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li class="iconAtualizar" runat="server">
                                        <asp:LinkButton ID="lnkAtualizar" Text="Atualizar" runat="server" />
                                    </li>
                                    <li class="iconLimpar" runat="server">
                                        <asp:LinkButton ID="lnkLimparO" Text="Limpar" runat="server" />
                                    </li>
                                    <li class="iconAjuda" runat="server">
                                        <asp:LinkButton ID="lnkAjudaO" Text="Ajuda" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="row">
                            <div class="collbl">
                                Registro:
                            </div>
                            <div class="coltxt">
                                <asp:TextBox ID="txtRegistro" TabIndex="2" runat="server" Width="87px" Font-Size="10pt"
                                    Font-Bold="True" />
                            </div>
                        </div>
                        <div class="row">
                            <div class="coltxt lg">
                                <asp:TextBox ID="txtObservacoes" runat="server" Width="99%" Height="312px" TextMode="MultiLine" />
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
