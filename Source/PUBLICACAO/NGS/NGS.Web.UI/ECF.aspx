<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="ECF.aspx.vb" Inherits="NGS.Web.UI.ECF" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .collbl {
            width: 170px;
        }
    </style>

    <script type="text/javascript">
        function downloadArquivo() {
            alert("Processo Concluido");
            //msgbox("Processo Concluido", "SUCESSO!", "Sucess");
            //$("#MainContent_imdDownload").click();
            $("#MainContent_cmdArquivoDeSaida").click();

        };
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scrmngECF" runat="server" AsyncPostBackTimeout="100000" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlECF" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                ECF - Escrituração Contábil Fiscal
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkProcessar" Text="Processar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Empresa:
                </div>
                <div class="coltxt">
                    <asp:DropDownList ID="DdlEmpresa" runat="server" Width="650px" AutoPostBack="True" />
                </div>
            </div>
            <div class="subtitulodiv">
                Parametros Bloco 0000
            </div>
            <div class="row">
                <div class="row">
                    <div class="collbl">
                        Inicio Do Periodo:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlIND_SIT_INI_PER" runat="server" Width="740px">
                            <asp:ListItem Value="0">0 – Regular (Início no primeiro dia do ano)</asp:ListItem>
                            <asp:ListItem Value="1">1 – Abertura (Início de atividades no ano-calendário)</asp:ListItem>
                            <asp:ListItem Value="2">2 – Resultante de cisão/fusão ou remanescente de cisão, ou realizou incorporação</asp:ListItem>
                            <asp:ListItem Value="3">3 – Resultante de Transformação</asp:ListItem>
                            <asp:ListItem Value="4">4 – Início de obrigatoriedade da entrega no curso do ano calendário. (Ex. Exclusão do Simples Nacional ou desenquadramento como imune ou isenta do IRPJ)</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Situação Especial:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlSIT_ESPECIAL" runat="server" Width="740px" AutoPostBack="True">
                            <asp:ListItem Value="0">0 – Normal (Sem ocorrência de situação especial ou evento)</asp:ListItem>
                            <asp:ListItem Value="1">1 – Extinção</asp:ListItem>
                            <asp:ListItem Value="2">2 – Fusão</asp:ListItem>
                            <asp:ListItem Value="3">3 – Incorporação \ Incorporada</asp:ListItem>
                            <asp:ListItem Value="4">4 – Incorporação \ Incorporadora</asp:ListItem>
                            <asp:ListItem Value="5">5 – Cisão Total</asp:ListItem>
                            <asp:ListItem Value="6">6 – Cisão Parcial</asp:ListItem>
                            <asp:ListItem Value="7">7 – Transformação</asp:ListItem>
                            <asp:ListItem Value="8">8 – Desenquadramento de Imune/Isenta</asp:ListItem>
                            <asp:ListItem Value="9">9 – Inclusão no Simples Nacional</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row" runat="server">
                    <div class="collbl">
                        Pat. Remanesc. Caso Cisão:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtPAT_REMAN_CIS" runat="server" data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Escrituração Retificadora:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlRETIFICADORA" runat="server" Width="740px" AutoPostBack="True">
                            <asp:ListItem Value="N">ECF original</asp:ListItem>
                            <asp:ListItem Value="S">ECF retificadora</asp:ListItem>
                            <asp:ListItem Value="F">ECF original com mudança de forma de tributação (Art. 5o da Instrução Normativa no 166/1999).</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row" runat="server">
                    <div class="collbl">
                        Número Recibo ECF Anterior:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNUM_REC" runat="server" data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Tipo ECF:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTIP_ECF" runat="server" Width="740px">
                            <asp:ListItem Value="0">0 – ECF de empresa não participante de SCP como sócio ostensivo.</asp:ListItem>
                            <asp:ListItem Value="1">1 – ECF de empresa participante de SCP como sócio ostensivo</asp:ListItem>
                            <asp:ListItem Value="2">2 – ECF da SCP.</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
            </div>
            <div class="subtitulodiv">
                Parametros Bloco 0010
            </div>
            <div class="row">
                <div class="row">
                    <div class="collbl">
                        Refis/Paes:
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkOPT_REFIS" runat="server" Text="Optante Pelo Refis"
                            data-ToolTip="default" ToolTip="Selecionar quando foir optante do Refis ou do Paes." />
                        <asp:CheckBox ID="chkOPT_PAES" runat="server" Text="Optante Pelo Paes"
                            data-ToolTip="default" ToolTip="Selecionar quando foir optante do Refis ou do Paes." />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Forma de Tributação:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlFORMA_TRIB" runat="server" Width="740px" AutoPostBack="True">
                            <asp:ListItem Value="0">Selecione uma Opção.</asp:ListItem>
                            <asp:ListItem Value="1">1 – Lucro Real.</asp:ListItem>
                            <asp:ListItem Value="2">2 – Lucro Real/Arbitrado.</asp:ListItem>
                            <asp:ListItem Value="3">3 – Lucro Presumido/Real.</asp:ListItem>
                            <asp:ListItem Value="4">4 – Lucro Presumido/Real/Arbitrado.</asp:ListItem>
                            <asp:ListItem Value="5">5 – Lucro Presumido.</asp:ListItem>
                            <asp:ListItem Value="6">6 – Lucro Arbitrado.</asp:ListItem>
                            <asp:ListItem Value="7">7 – Lucro Presumido/Arbitrado.</asp:ListItem>
                            <asp:ListItem Value="8">8 – Imune do IRPJ.</asp:ListItem>
                            <asp:ListItem Value="9">9 – Isenta do IRPJ.</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Forma Periodo Ap. IRPJ e da CSLL:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlFORMA_APUR" runat="server" Width="740px">
                            <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                            <asp:ListItem Value="T">T – Trimestral</asp:ListItem>
                            <asp:ListItem Value="A">A – Anual</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Qualificação da Pessoa Jurídica:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlCOD_QUALIF_PJ" runat="server" Width="740px">
                            <asp:ListItem Value="0">Selecione uma Opção</asp:ListItem>
                            <asp:ListItem Value="1">01 – PJ em Geral</asp:ListItem>
                            <asp:ListItem Value="2">02 – PJ Componente do Sistema Financeiro</asp:ListItem>
                            <asp:ListItem Value="3">03 – Sociedades Seguradoras, de Capitalização ou Entidade Aberta de Previdência Complementar</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Forma Tributação no Período: 
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlFORMA_TRIB_PER" runat="server" Width="740px">
                            <asp:ListItem Value=" ">Selecione uma Opção.</asp:ListItem>
                            <asp:ListItem Value="0">0 – ZERO – Não informado – trimestre não compreendido no período de apuração.</asp:ListItem>
                            <asp:ListItem Value="R">R – Real</asp:ListItem>
                            <asp:ListItem Value="P">P – Presumido</asp:ListItem>
                            <asp:ListItem Value="A">A – Arbitrado</asp:ListItem>
                            <asp:ListItem Value="E">E – Real Estimativa</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Forma Apuração Estimativa Mensal: 
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlMES_BAL_RED" runat="server" Width="740px">
                            <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                            <asp:ListItem Value="0">0 – Fora do Período: Fora do período de apuração/ Forma de tributação diferente de “R” ou “E”.</asp:ListItem>
                            <asp:ListItem Value="E">E – Receita Bruta: Estimativa com base na receita bruta e acréscimos.</asp:ListItem>
                            <asp:ListItem Value="B">B – Balanço ou Balancete: Estimativa com base no balanço ou balancete de suspensão/redução.</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Escrituração:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="ddlTIP_ESC_PRE" runat="server" Width="740px">
                            <asp:ListItem Value=" ">Selecione uma Opção.</asp:ListItem>
                            <asp:ListItem Value="L">L – Livro Caixa ou Hipótese prevista no §1o do art. 129, Instrução Normativa no 1.515/2014 (Lucro Presumido) ou Sem Escrituração (Imunes ou Isentas) ou Não obrigadas a entregar a ECD, de acordo com a Instrução Normativa no 1.420/2014..</asp:ListItem>
                            <asp:ListItem Value="C">C – Contábil (Lucro Presumido, Imunes ou Isentas)</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div id="DivIsentaImune" runat="server">
                    <div class="row">
                        <div class="collbl">
                            Tipo PJ Imune/Isenta:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlTIP_ENT" runat="server" Width="740px">
                                <asp:ListItem Value="0">Selecione uma Opção</asp:ListItem>
                                <asp:ListItem Value="01">01 – Assistência Social</asp:ListItem>
                                <asp:ListItem Value="02">02 – Educacional</asp:ListItem>
                                <asp:ListItem Value="03">03 – Sindicato de Trabalhadores</asp:ListItem>
                                <asp:ListItem Value="04">04 – Associação Civil</asp:ListItem>
                                <asp:ListItem Value="05">05 – Cultural</asp:ListItem>
                                <asp:ListItem Value="06">06 – Entidade Fechada de Previdência Complementar</asp:ListItem>
                                <asp:ListItem Value="07">07 – Filantrópica</asp:ListItem>
                                <asp:ListItem Value="08">08 – Sindicato</asp:ListItem>
                                <asp:ListItem Value="09">09 – Recreativa</asp:ListItem>
                                <asp:ListItem Value="10">10 – Científica</asp:ListItem>
                                <asp:ListItem Value="11">11 – Associação de Poupança e Empréstimo</asp:ListItem>
                                <asp:ListItem Value="12">12 – Entidade Aberta de Previdência Complementar (Sem Fins Lucrativos)</asp:ListItem>
                                <asp:ListItem Value="13">13 – Fifa e Entidades Relacionadas</asp:ListItem>
                                <asp:ListItem Value="14">14 – CIO e Entidades Relacionadas</asp:ListItem>
                                <asp:ListItem Value="15">15 – Partidos Políticos</asp:ListItem>
                                <asp:ListItem Value="99">99 – Outras.</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Apuração IRPJ p/ Imunes/Isentas: 
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlFORMA_APUR_I" runat="server" Width="740px">
                                <asp:ListItem Value="">Selecione uma Opção</asp:ListItem>
                                <asp:ListItem Value="A">A – Anual</asp:ListItem>
                                <asp:ListItem Value="T">T – Trimestral</asp:ListItem>
                                <asp:ListItem Value="D">D – Desobrigada</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Apuração CSLL p/ Imunes/Isentas:
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlAPUR_CSLL" runat="server" Width="740px">
                                <asp:ListItem Value=" ">Selecione uma Opção</asp:ListItem>
                                <asp:ListItem Value="A">A – Anual, se optou pela apuração da CSLL sobre a base de cálculo estimada, facultada a opção pelo levantamento de balanço ou balancete de suspensão ou redução.</asp:ListItem>
                                <asp:ListItem Value="T">T – Trimestral, no caso de ter adotado a apuração trimestral da CSLL.</asp:ListItem>
                                <asp:ListItem Value="D">D – Desobrigada, na hipótese de pessoa jurídica imune ou isenta da CSLL.</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                    <div class="row">
                        <div class="collbl">
                            Regime: 
                        </div>
                        <div class="coltxt">
                            <asp:DropDownList ID="ddlIND_REC_RECEITA" runat="server" Width="740px">
                                <asp:ListItem Value="">Selecione uma Opção</asp:ListItem>
                                <asp:ListItem Value="1">1 – Regime de caixa</asp:ListItem>
                                <asp:ListItem Value="2">2 – Regime de competência</asp:ListItem>
                            </asp:DropDownList>
                        </div>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        RTT:
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkOPT_EXT_RTT" runat="server" Text="Optante pela extinção do RTT no ano-calendário de 2014"
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        FCont:
                    </div>
                    <div class="coltxt">
                        <asp:CheckBox ID="chkDIF_FCONT" runat="server" Text="Existe diferenças entre a contabilidade societária e Fcont."
                            data-ToolTip="default" ToolTip="" />
                    </div>
                </div>
            </div>
            <div class="subtitulodiv">
                Parametros Bloco 0020
            </div>
            <div class="row">
                <div class="painelleft" style="width: 49%;">
                    <div id="Div_Ind_Aliq_Csll_I" runat="server">
                        PJ Sujeita à Alíquota da CSLL de 9% ou 17% ou 20% em 31/12/2015:
                    <asp:DropDownList ID="ddlIND_ALIQ_CSLL_I" runat="server" />
                        <br>
                    </div>
                    <div id="Div_Ind_Aliq_Csll" runat="server">
                        <asp:CheckBox ID="chkIND_ALIQ_CSLL" runat="server" Text="PJ Sujeita à Alíquota da CSLL de 15%" data-ToolTip="default" ToolTip="" />
                        <br />
                    </div>

                    Qtde de SCP da PJ - Sócio Ostensivo de SCP :
                    <asp:TextBox ID="txtIND_QTE_SCP" runat="server" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ADM_FUN_CLU" runat="server" Text="Administradora de Fundos e Clubes de Investimento" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PART_CONS" runat="server" Text="Participações em Consórcios de Empresas" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_OP_EXT" runat="server" Text="Operações com o Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_OP_VINC" runat="server" Text="Operações com Pessoa Vinculada/Interposta Pessoa / País com Tributação Favorecida" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PJ_ENQUAD" runat="server" Text="PJ Enquadrada nos artigos 48 ou 49 da IN RFB no 1.312/2012" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PART_EXT" runat="server" Text="Participações no Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ATIV_RURAL" runat="server" Text="Atividade Rural" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_LUC_EXP" runat="server" Text="Existência de Lucro da Exploração" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_RED_ISEN" runat="server" Text="Isenção e Redução do Imposto para Lucro Presumido" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_FIN" runat="server" Text="Indicativo da Existência de FINOR/FINAM/FUNRES" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_DOA_ELEIT" runat="server" Text="Doações a Campanhas Eleitorais" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PART_COLIG" runat="server" Text="Participação Avaliada pelo Método de Equivalência Patrimonial" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_VEND_EXP" runat="server" Text="PJ Efetuou Vendas a Empresa Comercial Exportadora com Fim Específico de Exportação - *Foi retirada a partir de 2020.*" data-ToolTip="default" ToolTip="" />
                </div>
                <div class="painelleft" style="width: 49%;">
                    <asp:CheckBox ID="chkIND_REC_EXT" runat="server" Text="Recebimentos do Exterior ou de Não Residentes" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ATIV_EXT" runat="server" Text="Ativos no Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_COM_EXP" runat="server" Text="PJ Comercial Exportadora" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PGTO_EXT" runat="server" Text="Pagamentos ao Exterior ou a Não Residentes" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_E_COM_TI" runat="server" Text="Comércio Eletrônico e Tecnologia da Informação" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ROY_REC" runat="server" Text="Royalties Recebidos do Brasil e do Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ROY_PAG" runat="server" Text="Royalties Pagos a Beneficiários do Brasil e do Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_REND_SERV" runat="server" Text="Rendimentos Relativos a Serviços, Juros e Dividendos Recebidos do Brasil e do Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PGTO_REM" runat="server" Text="Pagamentos ou Remessas a Título de Serviços, Juros e Dividendos a Beneficiários do Brasil e do Exterior" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_INOV_TEC" runat="server" Text="Inovação Tecnológica e Desenvolvimento Tecnológico" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_CAP_INF" runat="server" Text="Capacitação de Informática e Inclusão Digital" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_POLO_AM" runat="server" Text="Pólo Industrial de Manaus e Amazônia Ocidental" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_ZON_EXP" runat="server" Text="Zonas de Processamento de Exportação" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_AREA_COM" runat="server" Text="Áreas de Livre Comércio" data-ToolTip="default" ToolTip="" />
                    <br />
                    <asp:CheckBox ID="chkIND_PJ_HAB" runat="server" Text="PJ Habilitada no Repes, Recap, Padis, PATVD, Reidi, Repenec, Reicomp, Retaero, Recine, Resíduos Sólidos, Recopa, Copa do Mundo, Retid, REPNBL-Redes, Reif e Olimpíadas"
                        data-ToolTip="default" ToolTip="" />
                </div>
            </div>

            <div class="subtitulodiv">
                Outros
            </div>
            <div class="row">
                <div class="row">
                    <div class="collbl">
                        Periodo:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDT_INI" CssClass="calendario" Width="86px" runat="server"
                            data-ToolTip="default" ToolTip="Informar o período de apuração." />
                        &nbsp;a&nbsp;
                        <asp:TextBox ID="txtDT_FIN" runat="server" Width="86px" CssClass="calendario"
                            data-ToolTip="default" AutoPostBack="true" ToolTip="Informar o período de apuração." />
                    </div>
                </div>
                <div class="row" runat="server">
                    <div class="collbl">
                        Data Sit. Especial:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtDT_SIT_ESP" CssClass="calendario" Width="86px" runat="server" data-ToolTip="default" ToolTip="" />
                    </div>
                </div>

                <div class="row">
                    <div class="collbl">
                        Indicador de Movimento:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlIndicadorDeMovimento" runat="server" Width="550px">
                            <asp:ListItem Value="0">0 - Bloco Com Dados Informados</asp:ListItem>
                            <asp:ListItem Value="1">1 - Bloco sem dados informados</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Natureza do Livro Diário:
                    </div>
                    <div class="coltxt">
                        <asp:DropDownList ID="DdlNaturezaDoLivro" runat="server" Width="550px">
                            <asp:ListItem>livro Diario</asp:ListItem>
                        </asp:DropDownList>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Número de Ordem do Livro:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtNumeroDoLivro" runat="server" CssClass="txtNumerico"
                            data-ToolTip="default" ToolTip="Número de Ordem do Livro" />
                    </div>
                    <div class="collbl" style="display: none;">
                        <asp:CheckBox ID="cbTrimestral" runat="server" Text="Apuração do IRPJ Trimestral"
                            data-ToolTip="default" ToolTip="Número de Ordem do Livro" />
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Arquivo de Saida:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtArquivoDeSaida" ReadOnly="true" runat="server" Width="550px" />
                    </div>
                    <div class="coltxt">
                        <asp:UpdatePanel ID="updpnlArquivoDeSaida" runat="server">
                            <ContentTemplate>
                                <asp:Button ID="cmdArquivoDeSaida" CssClass="botao" runat="server" Height="20px"
                                    Text="Download" Visible="false" />
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="cmdArquivoDeSaida" />
                            </Triggers>
                        </asp:UpdatePanel>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Arquivo Auxiliar:
                    </div>
                    <div class="coltxt">
                        <asp:TextBox ID="txtArquivoAuxiliar" runat="server" Width="550px" data-ToolTip="default"
                            ToolTip="É o arquivo dos demonstrativos contábeis preparado pelo contador em formato (RTF) que será importado no processo de geração para ser incorporado no arquivo texto gerado pelo processo." />
                    </div>
                    <div class="coltxt">
                        <asp:UpdatePanel ID="updpnlArquivoDeSaidaAux" runat="server">
                            <ContentTemplate>
                                <asp:Button ID="cmdArquivoAuxDeSaida" runat="server" Height="20px" Text="Download"
                                    CssClass="botao" Visible="false" />
                            </ContentTemplate>
                            <Triggers>
                                <asp:PostBackTrigger ControlID="cmdArquivoAuxDeSaida" />
                            </Triggers>
                        </asp:UpdatePanel>
                    </div>
                </div>
                <div class="row">
                    <div class="collbl">
                        Plano De Contas:
                    </div>
                    <div class="coltxt">
                        <asp:RadioButton ID="rbPlanoNovo" runat="server" Checked="True" GroupName="plano" Text="Novo"
                            data-ToolTip="default" ToolTip="Selecionar o plano de contas." />
                        <asp:RadioButton ID="rbPlanoAntigo" runat="server" GroupName="plano" Text="Antigo"
                            data-ToolTip="default" ToolTip="Selecionar o plano de contas." />
                    </div>
                </div>
            </div>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>
