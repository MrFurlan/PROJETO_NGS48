<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="panteste.aspx.vb" Inherits="NGS.Web.UI.panteste" %>

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
        EnableScriptLocalization="True" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlpanteste" runat="server">
        <ContentTemplate>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li id="Li1" runat="server">
                            <asp:LinkButton class="iconNovo" Text="Processar" ID="lnkNovo" runat="server" OnClientClick=" if(!confirm('Deseja realmente Processar?')) return false;" />
                        </li>
                        <li id="Li2" runat="server">
                            <asp:LinkButton class="iconRelatorio" Text="Excel" ID="lnkExcel" runat="server" />
                        </li>
                        <li id="Li3" runat="server">
                            <asp:LinkButton class="iconLimpar" Text="Limpar" ID="lnkLimpar" runat="server" />
                        </li>
                        <li id="Li4" runat="server">
                            <asp:LinkButton class="iconLimpar" Text="API" ID="lnkAPI" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ano:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlAno" runat="server" AutoPostBack="True" />
                </div>
                <div class="collbl">
                    Mês:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlMes" runat="server">
                        <asp:ListItem Value="1">Janeiro</asp:ListItem>
                        <asp:ListItem Value="2">Fevereiro</asp:ListItem>
                        <asp:ListItem Value="3">Março</asp:ListItem>
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
            </div>
            <div class="row">
                <asp:Panel ID="Panel1" runat="server" GroupingText="Tabelas">
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkOperacoes" Style="margin-top: 4px;" runat="server" Text="Operações" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkOpeXEncargos" Style="margin-top: 4px;" runat="server" Text="OperaçõesXEncargos" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkCarteiras" Style="margin-top: 4px;" runat="server" Text="Carteiras" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkEncargos" Style="margin-top: 4px;" runat="server" Text="Encargos" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkCartXEnc" Style="margin-top: 4px;" runat="server" Text="CarteirasXEncargos" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <asp:Panel ID="Panel2" runat="server" GroupingText="Cadastros">
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkPlanoDeContas" Style="margin-top: 4px;" runat="server" Text="Plano de Contas" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkProduto" Style="margin-top: 4px;" runat="server" Text="Produto" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkClientes" Style="margin-top: 4px;" runat="server" Text="Clientes" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkClientesFaltantes" Style="margin-top: 4px;" runat="server" Text="Transfere Clientes faltantes" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <asp:Panel ID="Panel3" runat="server" GroupingText="Pedidos">
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkPedidos" Style="margin-top: 4px;" runat="server" Text="Transferir Pedidos" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <asp:Panel ID="Panel4" runat="server" GroupingText="NotasFiscais">
                    <div class="row">
                       <div class="coltxt">
                            <asp:CheckBox ID="chkPasso1" Style="margin-top: 4px;" runat="server" Text="Passo 1 - Deletar Notas/Razão" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkNotas" Style="margin-top: 4px;" runat="server" Text="Passo 2 - Transferir Notas Fiscais" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkTNotas" Style="margin-top: 4px;" runat="server" Visible="false" Text="Transferir Notas Fiscais(FURLAN)" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkLerPlanilha" Style="margin-top: 4px;" runat="server" Visible="false" Text="Ler Planilha(FURLAN)" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <asp:Panel ID="Panel5" runat="server" GroupingText="Estoque">
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkProducao" Style="margin-top: 4px;" runat="server" Text="Produção" />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row">
                <asp:Panel ID="Panel6" runat="server" GroupingText="Financeiro/Contábil">
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chkTitulos" Style="margin-top: 4px;" runat="server" Text="Títulos" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkRazao" Style="margin-top: 4px;" runat="server" Text="Razão" />
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkTitTransf" Style="margin-top: 4px;" runat="server" Text="Remove Título Transf." />
                        </div>
                    </div>
                </asp:Panel>
            </div>
            <div class="row" id="rExcel" runat="server" visible="false">
                <%--<asp:LinkButton ID="btnRodarExcel" runat="server" Text="RodarExcel" />
                <span>  Arquivo:
                    <uc:File ID="ucFile" runat="server" />
                </span  --%>
                <asp:LinkButton ID="lnkLerExcelImobilizado" runat="server" Text="Ler EXCEL Imobilizado" /> <br />           
                <asp:LinkButton ID="lnkLerExcelEncargos" runat="server" Text="Ler EXCEL Encargos" /> <br /> 
                <asp:LinkButton ID="lnkRecontabilizaTitulo" runat="server" Text="Recontabilizar Titulo" /> <br /> 
                <asp:LinkButton ID="lnkImportaClientes" runat="server" Text="Importa Clientes RTGraos" /> <br />
                <asp:LinkButton ID="lnkImportaProdutos" runat="server" Text="ImportaProdutosRTGraos" /> <br />
                <asp:LinkButton ID="lnkTransferirEncargos" runat="server" Text="Transferir Encargos de um Banco para Outro" /> <br /> 
               <%-- <asp:LinkButton ID="lnkverHorario" runat="server" Text="VerHorario" /> <br />
                <asp:Label ID="lblHorario" runat="server" Text=""></asp:Label> <br />--%>
                <asp:LinkButton ID="lnkImporta" runat="server" Text="Importa Excel BAXI" /> <br />
                <asp:LinkButton ID="lnkLerIBSCBS" runat="server" Text="Ler EXCEL IBS CBS" /> <br /> 
                <asp:LinkButton ID="lnkTransferirCotacoes" runat="server" Text="Transferir Cotacoes" /> <br /> 
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
