<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Impdados.aspx.vb" Inherits="NGS.Web.UI.Impdados" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        var cancela = false;
        var arquivoTitulo = "";

        function iniciar() {
            window.history.forward(1);
            document.cookie = "pagina=" + escape("Impdados");
        }

        function GravaBancos() {
            cancela = true;

            document.getElementById("Tempo").style.display = "block";
            document.body.style.cursor = "wait";

            if (document.forms[0].checkBancos.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Bancos.";
                Impdados.ImportaBancos(RetornaBancos);
            } else {
                GravaAgencias();
            }
        }

        function RetornaBancos(response) {
            if (response.value == "importado")
                document.getElementById("divBancos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divBancos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkBancos.checked = false;
            GravaAgencias();
        }

        function GravaAgencias() {
            if (document.forms[0].checkAgencias.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Agências.";
                Impdados.ImportaAgencias(RetornaAgencias);
            } else {
                GravaEstados();
            }
        }

        function RetornaAgencias(response) {
            if (response.value == "importado")
                document.getElementById("divAgencias").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divAgencias").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkAgencias.checked = false;
            GravaEstados();
        }

        function GravaEstados() {
            if (document.forms[0].checkEstados.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Importando Estados.";
                Impdados.ImportaEstados(RetornaEstados);
            } else {
                GravaHistoricos();
            }
        }

        function RetornaEstados(response) {
            if (response.value == "importado")
                document.getElementById("divEstados").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divEstados").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkEstados.checked = false;
            GravaHistoricos();
        }

        function GravaHistoricos() {
            if (document.forms[0].checkHistoricos.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Históricos Padrões.";
                Impdados.ImportaHistoricos(RetornaHistoricos);
            } else {
                GravaProdutos();
            }
        }

        function RetornaHistoricos(response) {
            if (response.value == "importado")
                document.getElementById("divHistoricos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divHistoricos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkHistoricos.checked = false;
            GravaProdutos();
        }

        function GravaProdutos() {
            if (document.forms[0].checkProdutos.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Produtos.";
                Impdados.ImportaProdutos(RetornaProdutos);
            } else {
                GravaGrupos();
            }
        }

        function RetornaProdutos(response) {
            if (response.value == "importado")
                document.getElementById("divProdutos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divProdutos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkProdutos.checked = false;
            GravaGrupos();
        }

        function GravaGrupos() {
            if (document.forms[0].checkGrupos.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Grupos de Estoque.";
                Impdados.ImportaGruposDeEstoques(RetornaGruposDeEstoque);
            } else {
                GravaClientes();
            }
        }

        function RetornaGruposDeEstoque(response) {
            if (response.value == "importado")
                document.getElementById("divGrupos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divGrupos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkGrupos.checked = false;
            GravaClientes();
        }

        function GravaClientes() {
            if (document.forms[0].checkClientes.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Clientes.";
                Impdados.ImportaClientes(RetornaClientes);
            } else {
                GravaTipos();
            }
        }

        function RetornaClientes(response) {
            if (response.value == "importado")
                document.getElementById("divClientes").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divClientes").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkClientes.checked = false;
            GravaTipos();
        }

        function GravaTipos() {
            if (document.forms[0].checkTipos.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Tipos de Clientes.";
                Impdados.ImportaTiposDeClientes(RetornaTiposDeClientes);
            } else {
                GravaNotas();
            }
        }

        function RetornaTiposDeClientes(response) {
            if (response.value == "importado")
                document.getElementById("divTipos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divTipos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkTipos.checked = false;
            GravaNotas();
        }

        function GravaNotas() {
            if (document.forms[0].checkNotas.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Notas Fiscais.";
                Impdados.ImportaNotas(RetornaNotas);
            } else {
                GravaTitulos();
            }
        }

        function RetornaNotas(response) {
            if (response.value == "importado")
                document.getElementById("divNotas").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divNotas").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkNotas.checked = false;
            GravaTitulos();
        }

        function GravaTitulos() {
            if (document.forms[0].checkTitulos.checked == true) {
                if (arquivoTitulo == "") {
                    alert("Arquivo para importação de Títulos não foi selecionado e não será processado.");
                    //msgbox("Arquivo para importação de Títulos não foi selecionado e não será processado.", "ATENÇÃO!", "Info");
                    document.forms[0].checkTitulos.checked = false;
                    GravaRazao();
                } else {
                    document.getElementById("divTitulo").innerHTML = "Aguarde, importando Títulos.";
                    Impdados.ImportaTitulos(arquivoTitulo, RetornaTitulos);
                }
            } else {
                GravaRazao();
            }
        }

        function RetornaTitulos(response) {
            if (response.value == "importado")
                document.getElementById("divTitulos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divTitulos").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkTitulos.checked = false;
            GravaRazao();
        }

        function GravaRazao() {
            if (document.forms[0].checkRazao.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Razão.";
                Impdados.ImportaRazao(RetornaRazao);
            } else {
                GravaRegfis();
            }
        }

        function RetornaRazao(response) {
            if (response.value == "importado")
                document.getElementById("divRazao").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divRazao").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkRazao.checked = false;
            GravaRegfis();
        }

        function GravaRegfis() {
            if (document.forms[0].checkRegfis.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Registros Fiscais.";
                Impdados.ImportaRegfis(RetornaRegfis);
            } else {
                GravaCodfis();
            }
        }

        function RetornaRegfis(response) {
            if (response.value == "importado")
                document.getElementById("divRegfis").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divRegfis").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkRegfis.checked = false;
            GravaCodfis();
        }

        function GravaCodfis() {
            if (document.forms[0].checkCodfis.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Códigos Fiscais.";
                Impdados.ImportaCodFis(RetornaCodfis);
            } else {
                GravaCademp();
            }
        }

        function RetornaCodfis(response) {
            if (response.value == "importado")
                document.getElementById("divCodfis").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divCodfis").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkCodfis.checked = false;
            GravaCademp();
        }

        function GravaCademp() {
            if (document.forms[0].checkCademp.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Cadastro de Empresas.";
                Impdados.ImportaCademp(RetornaCademp);
            } else {
                GravaCli491();
            }
        }

        function RetornaCademp(response) {
            if (response.value == "importado")
                document.getElementById("divCademp").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divCademp").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkCademp.checked = false;
            GravaCli491();
        }

        function GravaCli491() {
            if (document.forms[0].checkCli491.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Clintes IN-491.";
                Impdados.ImportaClientes491(RetornaCli491);
            } else {
                GravaProduto495();
            }
        }

        function RetornaCli491(response) {
            if (response.value == "importado")
                document.getElementById("divCli491").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divCli491").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkCli491.checked = false;
            GravaProduto495();
        }

        function GravaProduto495() {
            if (document.forms[0].checkPrd495.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Produtos IN-495.";
                Impdados.ImportaProdutos495(RetornaProduto495);
            } else {
                GravaNotas431();
            }
        }

        function RetornaProduto495(response) {
            if (response.value == "importado")
                document.getElementById("divPrd495").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divPrd495").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkPrd495.checked = false;
            GravaNotas431();
        }

        function GravaNotas431() {
            if (document.forms[0].checkNot431.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Notas Fiscais IN-431.";
                Impdados.ImportaNotas431(RetornaNotas431);
            } else {
                GravaNotas432();
            }
        }

        function RetornaNotas431(response) {
            if (response.value == "importado")
                document.getElementById("divNot431").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divNot431").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkNot431.checked = false;
            GravaNotas432();
        }

        function GravaNotas432() {
            if (document.forms[0].checkNot432.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Notas Fiscais IN-432.";
                Impdados.ImportaNotas432(RetornaNotas432);
            } else {
                GravaNotas433();
            }
        }

        function RetornaNotas432(response) {
            if (response.value == "importado")
                document.getElementById("divNot432").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divNot432").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkNot432.checked = false;
            GravaNotas433();
        }

        function GravaNotas433() {
            if (document.forms[0].checkNot433.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Notas Fiscais IN-433.";
                Impdados.ImportaNotas433(RetornaNotas433);
            } else {
                GravaNotas434();
            }
        }

        function RetornaNotas433(response) {
            if (response.value == "importado")
                document.getElementById("divNot433").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divNot433").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkNot433.checked = false;
            GravaNotas434();
        }

        function GravaNotas434() {
            if (document.forms[0].checkNot434.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Notas Fiscais IN-434.";
                Impdados.ImportaNotas434(RetornaNotas434);
            } else {
                GravaPlacon();
            }
        }

        function RetornaNotas434(response) {
            if (response.value == "importado")
                document.getElementById("divNot434").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divNot434").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkNot434.checked = false;
            GravaPlacon();
        }

        function GravaPlacon() {
            if (document.forms[0].checkPlacon.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando Plano de Contas.";
                Impdados.ImportaPlacon(RetornaPlacon);
            } else {
                GravaCompTit();
            }
        }

        function RetornaPlacon(response) {
            if (response.value == "importado")
                document.getElementById("divPlacon").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divPlacon").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkPlacon.checked = false;
            GravaCompTit();
        }

        function GravaCompTit() {
            if (document.forms[0].checkCompTit.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando ComprasXTítulos.";
                Impdados.ImportaComprasXTitulos(RetornaCompTit);
            } else {
                GravaCompPrd();
            }
        }

        function RetornaCompTit(response) {
            if (response.value == "importado")
                document.getElementById("divCompTit").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divCompTit").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkCompTit.checked = false;
            GravaCompPrd();
        }

        function GravaCompPrd() {
            if (document.forms[0].checkCompPrd.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando ComprasXProdutos.";
                Impdados.ImportaComprasXProdutos(RetornaCompPrd);
            } else {
                GravaTitulosXFamilias();
            }
        }

        function RetornaCompPrd(response) {
            if (response.value == "importado")
                document.getElementById("divCompPrd").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divCompPrd").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkCompPrd.checked = false;
            GravaTitulosXFamilias();
        }

        function GravaTitulosXFamilias() {
            if (document.forms[0].checkTitulosXFamilias.checked == true) {
                document.getElementById("divTitulo").innerHTML = "Aguarde, importando TitulosXFamilia.";
                Impdados.ImportaTitulosXFamilias(RetornaTitulosXFamilia);
            } else {
                Finaliza();
            }
        }

        function RetornaTitulosXFamilia(response) {
            if (response.value == "importado")
                document.getElementById("divTitulosXFamilias").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/certo.jpg' width='24' title='Importado com sucesso' />";
            else {
                document.getElementById("divTitulosXFamilias").innerHTML = "<img style='width: 24px; height: 24px; cursor: pointer' height='24' alt='' src='images/erro.jpg' width='24'title='" + response.value + "' />";
            }
            document.forms[0].checkTitulosXFamilias.checked = false;
            Finaliza();
        }

        function Finaliza() {
            fecharDiv('Tempo');
            document.body.style.cursor = "auto";
            alert("Processo concluído.");
            //msgbox("Processo concluído.", "SUCESSO!", "Sucess");
            document.forms[0].Finaliza.focus();
            arquivoTitulo = "";
            cancela = false;
        }

        function Sair() {
            if (cancela == false)
                parent.location.href = 'Gestao.aspx';
            else {
                alert("Aguarde o final da importação dos dados.");
                //msgbox("Aguarde o final da importação dos dados.", "ATENÇÃO!", "Info");
            }
        }

        function Cancelar() {
            if (cancela == true)
                alert("Aguarde o final da importação dos dados e click em Finaliza.");
                //msgbox("Aguarde o final da importação dos dados e click em Finaliza.", "ATENÇÃO!", "Info");
            else {
                parent.location.href = 'Imagem.aspx';
            }
        }

        function procurarArquivo() {
            if (document.forms[0].checkTitulos.checked == true) {
                document.getElementById("divArquivo").style.top = "250px";
                document.getElementById("divArquivo").style.left = "230px";
                document.getElementById("divArquivo").style.display = "block";
            }
        }

        function pegarCaminho() {
            arquivoTitulo = document.forms[0].fileImagem.value;
            fecharDiv('divArquivo');
        }
    </script>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngImpdados" runat="server" EnableScriptGlobalization="True"
        EnableScriptLocalization="True" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <table style="width: 100%; border: 0px none;">
        <tr>
            <td class="titulotabela" colspan="4">
                <label>
                    Importações de Arquivos
                </label>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <table class="actions" style="width: 100%;">
                    <tr>
                        <td class="iconConfirmar " runat="server" style="width: 10%;">
                            <asp:LinkButton ID="lnkConfirmar" runat="server">
                                <span>Confirmar</span>
                            </asp:LinkButton>
                        </td>
                        <td class="iconSair" runat="server" style="width: 10%;">
                            <asp:LinkButton ID="lnkCancelar" runat="server">
                                <span>Cancelar</span>
                            </asp:LinkButton>
                        </td>
                        <td class="iconAjuda" runat="server" style="width: 10%;">
                            <asp:LinkButton ID="lnkAjuda" runat="server">
                                <span>Ajuda</span>
                            </asp:LinkButton>
                        </td>
                        <td style="display: block;">
                        </td>
                    </tr>
                </table>
            </td>
        </tr>
        <tr>
            <td colspan="4">
                <div id="Tempo" style="display: none; position: static">
                    <img src="images/tempo.gif" alt="" />&nbsp;<div id="divTitulo" style="display: inline;
                        width: 20%">
                    </div>
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Bancos:<input id="checkBancos" type="checkbox" />
                <div id="divBancos" style="display: block; z-index: 100; position: static; width: 24px;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Agências:<input id="checkAgencias" type="checkbox" />
                <div id="divAgencias" style="display: block; z-index: 101; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Estados:<input id="checkEstados" type="checkbox" value="" />
                <div id="divEstados" style="display: block; z-index: 102; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Históricos:<input id="checkHistoricos" type="checkbox" />
                <div id="divHistoricos" style="display: block; z-index: 103; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Produtos:<input id="checkProdutos" type="checkbox" />
                <div id="divProdutos" style="display: block; z-index: 104; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Grupos de Estoque:<input id="checkGrupos" type="checkbox" />
                <div id="divGrupos" style="display: block; z-index: 106; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Clientes:<input id="checkClientes" type="checkbox" />
                <div id="divClientes" style="display: block; z-index: 107; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Tipos de Clientes:<input id="checkTipos" type="checkbox" />
                <div id="divTipos" style="display: block; z-index: 108; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Notas Fiscais:<input id="checkNotas" type="checkbox" />
                <div id="divNotas" style="display: block; z-index: 109; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Títulos:<input id="checkTitulos" type="checkbox" onclick="procurarArquivo();" />
                <div id="divTitulos" style="display: block; z-index: 110; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Razão:<input id="checkRazao" type="checkbox" />
                <div id="divRazao" style="display: block; z-index: 111; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Registros Fiscais:<input id="checkRegfis" type="checkbox" />
                <div id="divRegfis" style="display: block; z-index: 112; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Códigos Fiscais:<input id="checkCodfis" type="checkbox" />
                <div id="divCodfis" style="display: block; z-index: 113; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Cadastro de Empresas:<input id="checkCademp" type="checkbox" />
                <div id="divCademp" style="display: block; z-index: 114; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Clientes IN-491:<input id="checkCli491" type="checkbox" />
                <div id="divCli491" style="display: block; z-index: 115; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Produtos IN-495:<input id="checkPrd495" type="checkbox" />
                <div id="divPrd495" style="display: block; z-index: 116; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Notas Fiscais IN-431:<input id="checkNot431" type="checkbox" />
                <div id="divNot431" style="display: block; z-index: 117; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Notas Fiscais IN-432:<input id="checkNot432" type="checkbox" />
                <div id="divNot432" style="display: block; z-index: 118; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Notas Fiscais IN-433:<input id="checkNot433" type="checkbox" />
                <div id="divNot433" style="display: block; z-index: 119; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Notas Fiscais IN-434:<input id="checkNot434" type="checkbox" />
                <div id="divNot434" style="display: block; z-index: 120; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
        <tr>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                Plano de Contas:<input id="checkPlacon" type="checkbox" />
                <div id="divPlacon" style="display: block; z-index: 121; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                ComprasXTítulos:<input id="checkCompTit" type="checkbox" />
                <div id="divCompTit" style="display: block; z-index: 122; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                ComprasXProdutos:<input id="checkCompPrd" type="checkbox" />
                <div id="divCompPrd" style="display: block; z-index: 123; width: 24px; position: static;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
            <td style="width: 100px; text-align: center; vertical-align: text-top;">
                TitulosXFamilias:<input id="checkTitulosXFamilias" type="checkbox" />
                <div id="divTitulosXFamilias" style="display: block; z-index: 124; width: 24px; static: absolute;
                    height: 24px; background-color: #ffffff">
                </div>
            </td>
        </tr>
    </table>
    <div id="divArquivo" style="display: none; z-index: 125; left: 6px; width: 350px;
        position: absolute; top: 514px; background-color: #ffffff">
        <table class="tituloconsulta" id="table2" style="width: 100%; border: 0px none;">
            <tr>
                <td align="center">
                    <div id="div1" style="display: inline">
                        Arquivo:
                    </div>
                    &nbsp;<input id="fileImagem" style="width: 220px; height: 19px" type="file" size="15" />
                </td>
                <td style="width: 7%">
                    <img style="cursor: pointer" onclick="fecharDiv('divArquivo');" height="21" alt=""
                        src="images/Fechar.jpg" width="21" />
                </td>
            </tr>
            <tr>
                <td align="center">
                    <img style="cursor: pointer" onclick="pegarCaminho();" alt="" src="images/confirmar.gif" />
                </td>
            </tr>
        </table>
    </div>
</asp:Content>
