<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="SpedReinf.aspx.vb" Inherits="NGS.Web.UI.SpedReinf" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        #meioconteudo {
            width: 1024 !important;
        }
    </style>

    <script type="text/javascript">
        function downloadArquivo() {
            alert("Processo Concluido");
            //msgbox("Processo Concluido", "SUCESSO!", "Sucess");
            $("#MainContent_imdDownload").click();
        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngSpedReinf" runat="server" AsyncPostBackTimeout="50000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updnlSpedReinf" runat="server">
        <ContentTemplate>
            <div class="titulodiv">
                Sped Reinf
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li class="iconRelatorio rel" runat="server"><a>Relatório</a>
                            <ul>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel2010" runat="server" Text="R2010" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel2040" runat="server" Text="R2040" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel2055" runat="server" Text="R2055" />
                                </li>
                                <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel4010" runat="server" Text="R4010" />
                                </li>
                                 <li>
                                    <asp:LinkButton class="iconExcel" ID="lnkExcel4020" runat="server" Text="R4020" />
                                </li>
                            </ul>
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
                    <asp:DropDownList ID="ddlUnidade" runat="server" Width="634px" OnSelectedIndexChanged="DdlUnidade_SelectedIndexChanged"
                        AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="634px" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DropDownList1" runat="server" Width="634px" />
                </div>
            </div>

            <div class="row">
                <div class="collbl">
                    Período:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDataInicial" CssClass="calendario" runat="server" Width="75px" data-ToolTip="default" ToolTip="Informar o período inicial de apuração." />
                    &nbsp;à&nbsp;
                    <asp:TextBox ID="txtDataFinal" runat="server" CssClass="calendario" Width="75px" data-ToolTip="default" ToolTip="Informar o período final de apuração." />
                </div>
            </div>
            <div class="row">
                <asp:Panel ID="Panel1" runat="server" GroupingText="Blocos">
                    <div class="row">
                        <div>
                            <asp:Button ID="btn1000" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R1000 DADOS DO CONTRIBUINTE:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo1000" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R1000 DADOS DO CONTRIBUINTE." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo1000" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo1000" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo1000" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                    </div>
                    <%--<div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk1070" Style="margin-top: 4px;" runat="server" Text="R1070 PROCESSOS REF. A EMPRESA OU FORNECEDORES REF. ISENÇÃO OU RETENÇÃO" />
                        </div>
                    </div>--%>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn2010" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R2010 TOMADOR DE SERVIÇO - INFORMAÇÃO COM RETENÇÃO INSS SERVIÇOS COM RETENÇÃO REF. MÃO DE OBRA:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo2010" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2010 TOMADOR DE SERVIÇO - INFORMAÇÃO COM RETENÇÃO INSS SERVIÇOS COM RETENÇÃO REF. MÃO DE OBRA." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo2010" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo2010" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo2010" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <div>
	                        <asp:button ID="btnReat2010" runat="server" Text="Retificação"
		                        data-ToolTip="default" ToolTip="Enviar Retificação 2010." />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtRecib2010" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2010." Enabled="False" />
                            <asp:TextBox ID="txtseq2010" runat="server"  data-ToolTip="default"
                                ToolTip="R2010." Enabled="True" />
                        </div>
                    </div>
                    <div class="row">
                        <div>
                            <asp:Button ID="btn2040" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R2040 RECURSOS REPASSADOS PARA ASSOCIAÇÃO DESPORTIVA:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo2040" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2040 RECURSOS REPASSADOS PARA ASSOCIAÇÃO DESPORTIVA." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo2040" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo2040" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo2040" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <div>
	                        <asp:Button ID="btnReat2040" runat="server" Text="Retificação"
		                        data-ToolTip="default" ToolTip="Enviar Retificação 2040." />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtRecib2040" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2040." Enabled="False" />
                            <asp:TextBox ID="txtseq2040" runat="server"  data-ToolTip="default"
                                ToolTip="R2040." Enabled="True" />
                        </div>
                    </div>
                    <div class="row">
                        <div>
                            <asp:Button ID="btn2055" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R2055 AQUISIÇÃO DE PRODUÇÃO RURAL:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo2055" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2055 TOMADOR DE SERVIÇO - INFORMAÇÃO COM RETENÇÃO INSS SERVIÇOS COM RETENÇÃO REF. MÃO DE OBRA." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo2055" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo2055" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo2055" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <div>
	                        <asp:Button ID="btnReat2055" runat="server" Text="Retificação"
                                data-ToolTip="default" ToolTip="Enviar Retificação 2055." />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtRecib2055" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2055." Enabled="False" />
                            <asp:TextBox ID="txtseq2055" runat="server"  data-ToolTip="default"
                                ToolTip="R2055." Enabled="True" />
                        </div>
                    </div>

                    <%--<div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2020" Style="margin-top: 4px;" runat="server" Text="R2020 PRESTADOR DE SERVIÇO - INFORMAÇÃO COM RETENÇÃO INSS SERVIÇOS COM RETENÇÃO REF. MÃO DE OBRA" Enabled="False" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2030" Style="margin-top: 4px;" runat="server" Text="R2030 RECURSOS RECEBIDOS DAS ASSOCIAÇOES/TIMES ESPORTIVAS" Enabled="False" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2040" Style="margin-top: 4px;" runat="server" Text="R2040 RECURSOS REPASSADOS PARA ASSOCIAÇOES/TIMES ESPORTIVAS" Enabled="False" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2050" Style="margin-top: 4px;" runat="server" Text="R2050 COMERCIALIZAÇÃO DE PRODUTOS RURAIS OU AGROINDUSTRIA VER DENTRO DO MATERIA DA REINF" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2060" Style="margin-top: 4px;" runat="server" Text="R2060 CONTRIBUIÇÃO PREVIDENCIÁRIA SOBRE A RECEITA BRUTA - CPRB DESONERAÇÃO DA FOLHA" Enabled="False" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2070" Style="margin-top: 4px;" runat="server" Text="R2070 RETENÇÕES NA FONTE, IR - CSLL, COFINS, PIS/PASESP PAGAMENTOS DIVERSOS" Enabled="False" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk2098" Style="margin-top: 4px;" runat="server" Text="R2098 REABERTURA DOS EVENTOS PERIÓDICOS SÓ É PERMITIDO CASO JÁ TENHA FEITO O PRIMEIRO ENVIO E EXISTIR O R2099" />
                        </div>
                    </div>--%>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn2099" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R2099 FECHAMENTO DOS EVENTOS PERIÓDICOS REF. O PERÍODO APURADO:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo2099" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2099 FECHAMENTO DOS EVENTOS PERIÓDICOS REF. O PERÍODO APURADO." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo2099" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo2099" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo2099" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                    </div>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn2098" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R2098 REABERTURA DOS EVENTOS PERIÓDICOS REF. O PERÍODO APURADO:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo2098" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R2098 REABERTURA DOS EVENTOS PERIÓDICOS REF. O PERÍODO APURADO." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo2098" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo2098" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo2098" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                    </div>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn4010" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R4010 PAGAMENTOS/CRÉDITOS A BENEFICIÁRIO PESSOA FÍSICA:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo4010" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R4010." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo4010" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo4010" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo4010" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                    </div>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn4020" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R4020 PAGAMENTOS/CRÉDITOS A BENEFICIÁRIO PESSOA JURÍDICA:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo4020" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R4020." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo4020" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo4020" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo4020" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <div>
	                        <asp:Button ID="btnReat4020" runat="server" Text="Retificação"
		                        data-ToolTip="default" ToolTip="Enviar Retificação 4020." />
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtRecib4020" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R4020." Enabled="False" />
                            <asp:TextBox ID="txtseq4020" runat="server"  data-ToolTip="default"
                                ToolTip="R4020." Enabled="True" />
                        </div>
                    </div>

                    <div class="row">
                        <div>
                            <asp:Button ID="btn4099" runat="server" Visible="false" UseSubmitBehavior="False" Text="&gt;" CssClass="btn" />&nbsp; R4099 FECHAMENTO/REABERTURA DOS EVENTOS DA SÉRIE R-4000:
                        </div>
                        <div class="coltxt">
                            <asp:TextBox ID="txtArquivo4099" runat="server" Width="550px" data-ToolTip="default"
                                ToolTip="R4099." Enabled="False" />
                        </div>
                        <div class="coltxt">
                            <asp:UpdatePanel ID="updpnlArquivo4099" runat="server">
                                <ContentTemplate>
                                    <asp:Button ID="btnArquivo4099" runat="server" Height="20px" Text="Download"
                                        CssClass="botao" Visible="false" />
                                </ContentTemplate>
                                <Triggers>
                                    <asp:PostBackTrigger ControlID="btnArquivo4099" />
                                </Triggers>
                            </asp:UpdatePanel>
                        </div>
                        <div>
                            <asp:RadioButton ID="radReab4099" runat="server" Text="Reabertura"
                                data-ToolTip="default" ToolTip="Reabertura sim/não." />
                        </div>
                    </div>


<%--                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk5001" Style="margin-top: 4px;" runat="server" Text="R5001 INFORMAÇÕES DAS BASES E DOS TRIBUTOS CONSOLIDADOS POR CONTRIBUINTE CONSOLIDA O QUE FOI TRANSMITIDO DO MÊS ARQUIVO DE RETORNO DA RECEITA" />
                        </div>
                    </div>
                    <div class="row">
                        <div class="coltxt">
                            <asp:CheckBox ID="chk9000" Style="margin-top: 4px;" runat="server" Text="R9000 EXCLUSÃO DO EVENTO COM BASE NO PROTOCOLO DE ENVIO DESDE QUE O PERÍODO ESTEJA ABERTO" />
                        </div>
                    </div>--%>
                </asp:Panel>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
