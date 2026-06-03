<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Encargos.aspx.vb" Inherits="NGS.Web.UI.Encargos" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .painel
        {
            line-height: 18px;
            float: left;
        }
        
        #meioconteudo
        {
            width: 1080px !important;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmEncargos" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlEncargos" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Encargos
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconNovo" runat="server">
                            <asp:LinkButton ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li class="iconExcluir" runat="server">
                            <asp:LinkButton ID="lnkExcluir" Text="Excluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li class="iconLimpar" runat="server">
                            <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li class="iconRelatorio" runat="server">
                            <asp:LinkButton ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li class="iconAjuda" runat="server">
                            <asp:LinkButton ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <div class="collbl">
                        Encargo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="TxtEncargo" TabIndex="1" runat="server" Width="250px" MaxLength="15"
                            data-ToolTip="default" ToolTip="Nome abreviado do encargo." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Descrição:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="TxtNomeEncargo" TabIndex="2" runat="server" Width="415px" MaxLength="80"
                            data-ToolTip="default" ToolTip="Nome completo do encargo." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        C. Débito:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlDebito" TabIndex="3" runat="server" Width="425px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        C. Crédito:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlCredito" TabIndex="4" runat="server" Width="425px" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        % da Base Cálculo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtBaseCalculo" runat="server" CssClass="txtDecimal" TabIndex="5"
                            Width="110px" MaxLength="5" data-ToolTip="default" ToolTip="Percentual da base de cálculo." />
                    </div>
                    <div class="collbl" style="width: 150px">
                        % Base Calculado Sobre:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlValorOuPeso" runat="server" Width="140px" TabIndex="8">
                            <asp:ListItem Value="-1" Selected="True">Nenhum</asp:ListItem>
                            <asp:ListItem Value="0">VALOR</asp:ListItem>
                            <asp:ListItem Value="1">PESO/QUANTIDADE</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Operador Aliquota:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlOperador" runat="server" Width="120px" TabIndex="8">
                            <asp:ListItem Value="" Selected="True">Nenhum</asp:ListItem>
                            <asp:ListItem Value="*">Multiplicado</asp:ListItem>
                            <asp:ListItem Value="%">Percentual</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="collbl" style="width: 150px">
                        Alíquota:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtAliquota" runat="server" CssClass="txtDecimal" TabIndex="6" MaxLength="5"
                            Width="110px" data-ToolTip="default" ToolTip="Informar a alíquota do encargo." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Etapa:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEtapa" runat="server" Width="120px" TabIndex="7">
                            <asp:ListItem Value="0">Normal</asp:ListItem>
                            <asp:ListItem Value="1">Adiantamento</asp:ListItem>
                            <asp:ListItem Value="2">Amortizacao</asp:ListItem>
                            <asp:ListItem Value="3">Liquidacao</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                    <div class="collbl" style="width: 150px">
                        Encargo Agrupador:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlEncargoAgrupador" runat="server" Width="120px" TabIndex="7">
                            <asp:ListItem Value="" Selected="True">Nenhum</asp:ListItem>
                            <asp:ListItem Value="PIS">PIS</asp:ListItem>
                            <asp:ListItem Value="ICMS">ICMS</asp:ListItem>
                            <asp:ListItem Value="COFINS">COFINS</asp:ListItem>
                            <asp:ListItem Value="IPI">IPI</asp:ListItem>
                            <asp:ListItem Value="ISS">ISS</asp:ListItem>
                            <asp:ListItem Value="PRODUTO">PRODUTO</asp:ListItem>
                            <asp:ListItem Value="LIQUIDO">LIQUIDO</asp:ListItem>
                            <asp:ListItem Value="DESCONTOS">DESCONTOS</asp:ListItem>
                            <asp:ListItem Value="FRETES">FRETES</asp:ListItem>
                            <asp:ListItem Value="SEGURO">SEGURO</asp:ListItem>
                            <asp:ListItem Value="DESP.ADUANEIRAS">DESP.ADUANEIRAS</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
            </div>
            <div class="painelleft">
                <div class="row">
                    <asp:CheckBox ID="chkOperacaoXEncargos" Style="margin-top: 4px;" runat="server" Text="Encargo vai ser Configurado no OperacõesXEncargos"
                        AutoPostBack="True" />
                </div>
                <div class="row">
                    <asp:Label ID="lbl" Style="margin-top: 4px;" runat="server" Text="Encargo de Pessoa:" />
                    <asp:DropDownList ID="ddlPessoa" TabIndex="3" runat="server" Width="80px" AutoPostBack="True">
                        <asp:ListItem Selected="True" Value="0">Nenhuma</asp:ListItem>
                        <asp:ListItem Value="1">Fisica</asp:ListItem>
                        <asp:ListItem Value="2">Juridica</asp:ListItem>
                        <asp:ListItem Value="3">Ambas</asp:ListItem>
                    </asp:DropDownList>
                    &nbsp; ou &nbsp;
                    <asp:CheckBox ID="chkObgEmpresa" runat="server" Text="Verificar Obrigatoriedade de Emissao(Empresa Emitente)" />
                </div>
                <div class="row">
                    <asp:CheckBox ID="chkRetencao" Style="margin-top: 4px;" runat="server" Text="Encargo Poderá Sofrer Retenção no tipo de Pessoa:"
                        AutoPostBack="True" data-ToolTip="default" ToolTip="" />
                    <asp:DropDownList ID="ddlPessoaRetencao" TabIndex="3" runat="server" Width="80px">
                        <asp:ListItem Selected="True" Value="0">Nenhuma</asp:ListItem>
                        <asp:ListItem Value="1">Fisica</asp:ListItem>
                        <asp:ListItem Value="2">Juridica</asp:ListItem>
                        <asp:ListItem Value="3">Ambas</asp:ListItem>
                    </asp:DropDownList>
                </div>
                <div class="row">
                    <asp:CheckBox ID="chkAtualizar" runat="server" Text="Os Valores Podem ser Alterados" />
                </div>
                <div class="row">
                    <asp:CheckBox ID="chkGravaCentroDeCusto" runat="server" Text="Grava Centro de Custo no Razao" />
                </div>
                <div class="row">
                    <asp:CheckBox ID="chkImprimirNFE" runat="server" Text="Imprimir NFE" />
                </div>
            </div>
            <div class="bordagrid">
                <asp:GridView ID="GridEncargos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                    ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridEncargos_SelectedIndexChanged">
                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                    <EditRowStyle BackColor="#999999" />
                    <Columns>
                        <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                        <asp:BoundField DataField="Codigo" HeaderText="Encargo" ReadOnly="True">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Descricao" HeaderText="Descricao">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="TipoPessoa" HeaderText="Tipo Pessoa" />
                        <asp:TemplateField HeaderText="Verifica Empresa">
                            <ItemTemplate>
                                <asp:Label ID="lblVerEmp" runat="server" Text='<%# Bind("VerificaEmpresa") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="ContaDebito" HeaderText="Debito">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="ContaCredito" HeaderText="Credito">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="BaseCalculo" HeaderText="Base">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Aliquota" HeaderText="Aliquota">
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="C.C.">
                            <ItemTemplate>
                                <asp:Label ID="lblCC" runat="server" Text='<%# Bind("GravaCentroDeCusto") %>'></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="OpXEnc">
                            <ItemTemplate>
                                <asp:Label ID="lblOpEnc" runat="server" Text='<%# Bind("OperacaoXEncargo") %>'></asp:Label>
                            </ItemTemplate>
                            <HeaderStyle HorizontalAlign="Right" />
                            <ItemStyle HorizontalAlign="Right" />
                        </asp:TemplateField>
                        <asp:TemplateField HeaderText="Atual.">
                            <ItemTemplate>
                                <asp:Label ID="lblAtua" runat="server" Text='<%# Bind("Atualizacao") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="Etapa" HeaderText="Etapa">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Imp. NFE">
                            <ItemTemplate>
                                <asp:Label ID="lblNfe" runat="server" Text='<%# Bind("ImprimirNFE") %>'></asp:Label>
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="EncargoAgrupador" HeaderText="Enc. Agrupador">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:BoundField DataField="Operador" HeaderText="Operador">
                            <HeaderStyle HorizontalAlign="Left" />
                            <ItemStyle HorizontalAlign="Left" />
                        </asp:BoundField>
                        <asp:TemplateField HeaderText="Sobre">
                            <ItemTemplate>
                                <asp:Label ID="ValorOuPeso" runat="server" Text='<%# Bind("ValorOuPeso") %>' />
                            </ItemTemplate>
                        </asp:TemplateField>
                        <asp:BoundField DataField="TipoPessoaRetencao" HeaderText="Pessoa Retenção" />
                    </Columns>
                </asp:GridView>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
