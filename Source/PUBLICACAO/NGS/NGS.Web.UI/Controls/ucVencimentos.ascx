<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucVencimentos.ascx.vb"
    Inherits="NGS.Web.UI.ucVencimentos" %>
<script type="text/javascript" language="javascript">
    function pageLoadVencimentos() {
        $("#MainContent_ucVencimentos_txtDataVencimento").datepicker({
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

        $("#MainContent_ucVencimentos_txtDataVencimento").setMask('date');
        $("#MainContent_ucVencimentos_txtValorDocumento").setMask('decimal');
        $("#MainContent_ucVencimentos_txtDesconto").setMask('decimal');
        $("#MainContent_ucVencimentos_txtDeducao").setMask('decimal');
        $("#MainContent_ucVencimentos_txtMultaJuro").setMask('decimal');
        $("#MainContent_ucVencimentos_txtOutrosAcrescimos").setMask('decimal');
        $("#MainContent_ucVencimentos_txtValorLiquido").setMask('decimal');
    }

    $(document).ready(function () {
        pageLoadVencimentos();
    });

    var prmVencimentos = Sys.WebForms.PageRequestManager.getInstance();
    prmVencimentos.add_endRequest(pageLoadVencimentos);
</script>
<div id="divConsultaVencimentos" class="uc" title="Consulta de Vencimentos" style="display: none;">
    <center>
        <asp:UpdatePanel ID="updpnlVencimentos" runat="server">
            <ContentTemplate>
                <asp:HiddenField ID="HID" runat="server" />
                <asp:HiddenField ID="IndiceGrid" runat="server" Value="-1" />
                <asp:HiddenField ID="Origem" runat="server" />
                <div class="titulodiv">
                    Vencimentos
                </div>
                <div class="painelleft" style="width: 49.5%;">
                    <div class="row">
                        <div class="coltxt lg textcenter">
                            Carteira:
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt lg textcenter">
                            <asp:DropDownList ID="ddlCarteira" runat="server" Font-Names="monospace" Width="350" AutoPostBack="True" OnSelectedIndexChanged="ddlCarteira_SelectedIndexChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt lg textcenter">
                            Condição de Pagamento:
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt lg textcenter">
                            <asp:DropDownList ID="ddlCondicaoPagamento" runat="server" Width="350px" AutoPostBack="True"
                                OnSelectedIndexChanged="ddlCondicaoPagamento_SelectedIndexChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="bordagrid">
                            <asp:GridView ID="gridVencimentos" runat="server" CellPadding="4" ForeColor="#333333"
                                GridLines="None" AutoGenerateColumns="False" OnSelectedIndexChanged="gridVencimentos_SelectedIndexChanged"
                                Width="100%" >
                                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                <EditRowStyle BackColor="#999999" />
                                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                <Columns>
                                    <asp:CommandField ButtonType="Button" SelectText=" &gt; " ShowSelectButton="True" />
                                    <asp:BoundField DataField="Codigo" HeaderText="Titulo" />
                                    <asp:BoundField DataField="Vencimento" HeaderText="Vencimento" DataFormatString="{0:dd/MM/yyyy}">
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="ValorLiquido" HeaderText="Valor Parcela Oficial" DataFormatString="{0:N2}">
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                    <asp:BoundField DataField="MoedaValorLiquido" HeaderText="Valor Parcela Moeda" DataFormatString="{0:N2}">
                                        <ItemStyle HorizontalAlign="Right" />
                                    </asp:BoundField>
                                </Columns>
                            </asp:GridView>
                        </div>
                    </div>
                </div>
                <div class="painelright" style="width: 49.5%;">
                    <div class="row">
                        <div class="collbluc">
                            Título:
                        </div>
                        <div class="coltxt">
                            <asp:Label ID="txtCodigoTitulo" runat="server" Font-Bold="True" Font-Size="Medium"
                                ForeColor="Red" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Data:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDataVencimento" runat="server" AutoPostBack="True" OnTextChanged="txtDataVencimento_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Tipo Pagto:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTipoDePagto" runat="server" Width="365px" AutoPostBack="True"
                                OnSelectedIndexChanged="ddlTipoDePagto_SelectedIndexChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Cód. Barra:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtCodigoDeBarra" runat="server" Width="300px" AutoPostBack="True"
                                OnTextChanged="txtCodigoDeBarra_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            &nbsp;
                        </div>
                        <div class="coltxt">
                            <asp:CheckBox ID="chkDigitado" runat="server" Text="Digitado" AutoPostBack="True"
                                OnCheckedChanged="chkDigitado_CheckedChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Moeda:
                        </div>
                        <div class="coltxt">
                            <asp:Label ID="txtMoeda" runat="server" Font-Bold="True" Font-Size="Medium" ForeColor="Red" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Valor Doc.:
                        </div>
                        <div class="coltxt">
                            <asp:HiddenField ID="ValorDocumento" runat="server" Value="0" />
                            <asp:TextBox ID="txtValorDocumento" runat="server" AutoPostBack="True" OnTextChanged="txtValorDocumento_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Desconto:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDesconto" runat="server" AutoPostBack="True" OnTextChanged="txtDesconto_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Dedução:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDeducao" runat="server" AutoPostBack="True" OnTextChanged="txtDeducao_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Multa/Juro:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtMultaJuro" runat="server" AutoPostBack="True" OnTextChanged="txtMultaJuro_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Outros Acréscimos:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtOutrosAcrescimos" runat="server" AutoPostBack="True" OnTextChanged="txtOutrosAcrescimos_TextChanged" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Vlr Líquido
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtValorLiquido" runat="server" AutoPostBack="True" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt lg textright">
                            <asp:Button ID="BtnConta" runat="server" OnClick="BtnConta_Click" CssClass="botao"
                                Text="Carregar Conta Cadastrada do Cliente" Width="255px" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Banco:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtCodBanco" runat="server" Width="63px" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtNomeBanco" runat="server" Width="280px" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Agência:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtCodAgencia" runat="server" Width="63px" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDigitoAgencia" runat="server" Width="21px" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtPracaAgencia" runat="server" Width="250px" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbluc">
                            Conta:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtConta" runat="server" Width="96px" />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtDigitoConta" runat="server" Width="24px" />
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="coltxt right">
                        <asp:Button ID="btnSair" runat="server" CssClass="botao" Text="Sair" OnClick="btnSair_Click" />
                    </div>
                    <div class="coltxt right">
                        <asp:Button ID="btnConfirmar" runat="server" CssClass="botao" Text="Confirmar" OnClick="btnConfirmar_Click" />
                    </div>
                </div>
            </ContentTemplate>
        </asp:UpdatePanel>
    </center>
</div>
