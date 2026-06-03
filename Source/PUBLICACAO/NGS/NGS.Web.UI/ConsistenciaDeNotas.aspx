<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ConsistenciaDeNotas.aspx.vb" Inherits="NGS.Web.UI.ConsistenciaDeNotas" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngConsistenciaDeNotas" runat="server" AsyncPostBackTimeout="900" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlConsistenciaDeNotas" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Consistência de Notas
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconPdf" ID="lnkPdf" runat="server" Text="Pdf" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel" runat="server" Text="Excel Dados" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcelPlaca" runat="server" Text="Excel/Placa" />
                                </li>
                            </ul>
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server"
                                Text="Limpar" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server"
                                Text="Ajuda" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Unidade:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkUnificarEmpresa" runat="server" Text="Empresa  :" ToolTip="Unificar os CNPJs da empresa." />
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cliente:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoCliente" runat="server" />
                    <asp:TextBox ID="txtCliente" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdCliente" CssClass="btn" OnClick="cmdCliente_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente desejado." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Local de Embarque:
                </div>
                <div class="coltxt">
                    <asp:HiddenField ID="txtCodigoEmbarque" runat="server" />
                    <asp:TextBox ID="txtClienteEmbarque" runat="server" Width="600px" Enabled="false" />
                </div>
                <div class="coltxt">
                    <asp:Button ID="cmdClienteEmbarque" CssClass="btn" OnClick="cmdClienteEmbarque_Click" runat="server"
                        Text=">" data-ToolTip="default" ToolTip="Selecionar o cliente de embarque." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Operação:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="cmbOperacao" runat="server" Width="294px" AutoPostBack="True"
                        OnSelectedIndexChanged="cmbOperacao_SelectedIndexChanged" />
                    <asp:DropDownList ID="cmbSubOperacao" runat="server" Width="295px" />
                </div>
            </div>
            <div class="row">
                <div class="coltxt lg">
                    <uc:SelecaoProduto ID="ucSelecaoProduto" runat="server" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Notas:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadEntradas" runat="server" Text="Entradas" Checked="True" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                    <asp:RadioButton ID="RadSaidas" runat="server" Text="Saidas" GroupName="Notas" data-ToolTip="default" ToolTip="Informar se a nota é de entrada ou saída." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Transitado:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="rdDentroEstado" runat="server" Text="Dentro do Estado" GroupName="Estado" data-ToolTip="default" ToolTip="Apenas notas com o mesmo Estado da Empresa." />
                    <asp:RadioButton ID="rdForaEstado" runat="server" Text="Fora do Estado" GroupName="Estado" data-ToolTip="default" ToolTip="Apenas notas com o Estado diferente da Empresa." />
                    <asp:RadioButton ID="rdTodosEstados" runat="server" Text="Todos os Estados" Checked="True" GroupName="Estado" data-ToolTip="default" ToolTip="Notas de todos os Estados." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Placa:
                </div>
                <div class="coltxt" style="width: 126px;">
                    <asp:TextBox ID="txtPlaca" runat="server" Width="96px" />
                </div>
                <div class="collbl">
                    OperaçãoXEstado:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtOpercaoXEstado" runat="server" Width="96px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Período Inicial:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data inicial de consulta." />
                </div>
                <div class="collbl">
                    Período Final:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataFinal" CssClass="calendario" runat="server" Width="96px" data-ToolTip="default" ToolTip="Informar o data final de consulta." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Ordem:
                </div>
                <div class="coltxt">
                    <asp:RadioButton ID="RadMovimento" runat="server" Text="Movimento" Checked="True" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena relatório pelo movimento." />
                    <asp:RadioButton ID="RadNome" runat="server" Text="Nome" GroupName="Ordem" data-ToolTip="default" ToolTip="Ordena o relatório pelo nome." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkNossaEmissao" Text="Nossa Emissão:" runat="server" AutoPostBack="True" data-ToolTip="default" ToolTip="Apenas Notas Fiscais Nossa Emissão." />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <asp:CheckBox ID="chkAllTipos" data-ToolTip="default" ToolTip="Seleciona todos os Tipos de Documentos."
                        Text="Tipo Documento:" runat="server" AutoPostBack="True" OnCheckedChanged="chkAllTipos_CheckedChanged" />
                </div>
                <div class="coltxt" style="line-height: 12px;">
                    <asp:CheckBoxList ID="chkTipoDeDocumento" runat="server" RepeatColumns="3" data-ToolTip="default"
                        ToolTip="Selecionar qual ou quais os tipos de documentos." />
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaClientes ID="ucConsultaClientes" runat="server" />
</asp:Content>
