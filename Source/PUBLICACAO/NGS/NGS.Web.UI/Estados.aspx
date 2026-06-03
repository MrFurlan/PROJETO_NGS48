<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Estados.aspx.vb" Inherits="NGS.Web.UI.Estados" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="http://cdn.datatables.net/1.10.4/css/jquery.dataTables.css" rel="stylesheet" type="text/css"/>
    <link href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" rel="stylesheet" type="text/css"/>
    <link href="http://cdn.datatables.net/plug-ins/3cfcc339e89/integration/jqueryui/dataTables.jqueryui.css" rel="stylesheet" />
    
    <script src="http://cdn.datatables.net/1.10.4/js/jquery.dataTables.min.js" type="text/javascript"></script>
    
    <style type="text/css">
        div.dataTables_length, div.dataTables_filter { margin: 5px; }
        div.dataTables_filter { margin-right: 2px; }
        div.dataTables_wrapper { border: 1px solid #5D7B9D; border-radius: 5px; color: #5D789D; margin-top: 5px; }
        .paginate_button { border-radius: 5px; }
        .dataTables_paginate { padding: 0.25em; }
        dataTables_info { padding: 0.755em; }
    </style>

    <script type="text/javascript">
        const resources = {
            pt: { translation: {
                'page.title': 'Estados',
                'actions.new': 'Gravar',
                'actions.refresh': 'Atualizar', 
                'actions.delete': 'Excluir',
                'actions.clear': 'Limpar',
                'actions.report': 'Relatório',
                'actions.help': 'Ajuda',
                'labels.uf': 'UF',
                'labels.region': 'Região:',
                'labels.description': 'Descrição:',
                'confirm.delete': 'Tem certeza que deseja excluir este registro?',
                'datatables.length': 'MENU Registros por páginas',
                'datatables.zeroRecords': 'desculpe - nenhum registro encontrado...',
                'datatables.info': 'Mostrando PAGE a PAGES',
                'datatables.infoEmpty': 'não avaliado',
                'datatables.infoFiltered': '(filtrado MAX Total de registros)',
                'datatables.search': 'Pesquisa:',
                'datatables.next': 'Próxima',
                'datatables.previous': 'Anterior'
            }},
            en: { translation: {
                'page.title': 'States',
                'actions.new': 'Save',
                'actions.refresh': 'Refresh',
                'actions.delete': 'Delete',
                'actions.clear': 'Clear',
                'actions.report': 'Report',
                'actions.help': 'Help',
                'labels.uf': 'UF',
                'labels.region': 'Region:',
                'labels.description': 'Description:',
                'confirm.delete': 'Are you sure you want to delete this record?',
                'datatables.length': 'MENU Records per page',
                'datatables.zeroRecords': 'sorry - no records found...',
                'datatables.info': 'Showing PAGE to PAGES',
                'datatables.infoEmpty': 'not evaluated',
                'datatables.infoFiltered': '(filtered from MAX total records)',
                'datatables.search': 'Search:',
                'datatables.next': 'Next',
                'datatables.previous': 'Previous'
            }},
            es: { translation: {
                'page.title': 'Estados',
                'actions.new': 'Guardar',
                'actions.refresh': 'Actualizar',
                'actions.delete': 'Eliminar',
                'actions.clear': 'Limpiar',
                'actions.report': 'Informe',
                'actions.help': 'Ayuda',
                'labels.uf': 'UF',
                'labels.region': 'Región:',
                'labels.description': 'Descripción:',
                'confirm.delete': '¿Está seguro que desea eliminar este registro?',
                'datatables.length': 'MENU Registros por página',
                'datatables.zeroRecords': 'lo siento - no se encontraron registros...',
                'datatables.info': 'Mostrando PAGE a PAGES',
                'datatables.infoEmpty': 'no evaluado',
                'datatables.infoFiltered': '(filtrado de MAX registros totales)',
                'datatables.search': 'Buscar:',
                'datatables.next': 'Siguiente',
                'datatables.previous': 'Anterior'
            }}
        };

        i18next.init({
            lng: localStorage.getItem("lang") || "pt",
            fallbackLng: "pt",
            resources
        }, function() {
            updateContent();
        });

        function updateContent() {
            $('[data-i18n]').each(function() {
                $(this).text(i18next.t($(this).data('i18n')));
            });
            
            $('.iconNovo').attr('title', i18next.t('actions.new'));
            $('.iconAtualizar').attr('title', i18next.t('actions.refresh'));
            $('.iconExcluir').attr('title', i18next.t('actions.delete'));
            $('.iconLimpar').attr('title', i18next.t('actions.clear'));
            $('.iconRelatorio').attr('title', i18next.t('actions.report'));
            $('.iconAjuda').attr('title', i18next.t('actions.help'));
            
            $('.lang-btn').removeClass('active');
            $(`.lang-btn[data-lang="${i18next.language}"]`).addClass('active');
        }

        function changeLanguage(lang) {
            i18next.changeLanguage(lang, function() {
                localStorage.setItem("lang", lang);
                updateContent();
                if ($.fn.DataTable.isDataTable('#MainContent_GridEstados')) {
                    $('#MainContent_GridEstados').DataTable().destroy();
                }
                pageLoadEstados();
            });
        }

        function pageLoadEstados() {
            $('#MainContent_GridEstados').DataTable({
                stateSave: true,
                "language": {
                    "lengthMenu": i18next.t('datatables.length'),
                    "zeroRecords": i18next.t('datatables.zeroRecords'),
                    "info": i18next.t('datatables.info'),
                    "infoEmpty": i18next.t('datatables.infoEmpty'),
                    "infoFiltered": i18next.t('datatables.infoFiltered'),
                    "search": i18next.t('datatables.search'),
                    "paginate": {
                        "next": i18next.t('datatables.next'),
                        "previous": i18next.t('datatables.previous')
                    }
                }
            });
           
            $("#resultado").click(function () {
                $.ajax({
                    type: "POST",
                    url: "/Estados.aspx/MeuMetodo",
                    data: "{'desc': '" + $("#resultado").text() + "'}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (msg) {
                        $("#resultado").text(msg.d);
                    },
                    error: function () {
                        alert(i18next.t('confirm.delete'));
                    }
                });
            });
        }

        $(document).ready(function () {
            pageLoadEstados();

            changeLanguage('pt');

            var prmEstados = Sys.WebForms.PageRequestManager.getInstance();
            prmEstados.add_endRequest(function() {
                pageLoadEstados();
                updateContent();
            });
        });
    </script>
</asp:Content>

<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngEstados" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    
    <asp:UpdatePanel ID="updpnlEstados" runat="server">
        <ContentTemplate>
            <div id="dialog-confirm" title="ATENÇÃO?" style="display: none;">
                <p><span class="ui-icon ui-icon-alert" style="float: left; margin: 0 7px 20px 0;"></span>
                    <span data-i18n="confirm.delete"></span>
                </p>
            </div>
            
            <div class="titulodiv">
                <span data-i18n="page.title"></span>
            </div>            
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" runat="server" data-i18n="actions.new" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" runat="server" data-i18n="actions.refresh" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" runat="server" data-i18n="actions.delete"
                                OnClientClick="return msgconfirm(this);" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" runat="server" data-i18n="actions.clear" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" runat="server" data-i18n="actions.report" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" runat="server" data-i18n="actions.help" />
                        </li>
                    </ul>
                </div>
            </div>
            
            <div class="row">
                <div class="collbl">
                    <label id="resultado" data-i18n="labels.uf"></label>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCodigo" Style="text-transform: uppercase;" runat="server" Width="50px" minlength="2" required=""
                        data-tooltip="default" ToolTip="Codigo do Estado (sigla caracterizado por dois dígitos) Ex: Para o estado Paraná utiliza-se ''PR''"
                        MaxLength="2" />
                </div>
                <div class="collbl">
                    <span data-i18n="labels.region"></span>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtRegiao" Style="text-transform: uppercase;" runat="server" Width="50px"
                        MaxLength="2" data-tooltip="default" ToolTip="Região (sigla caracterizado por dois digitos) Ex: DR. " />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    <span data-i18n="labels.description"></span>
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtDescricao" data-tooltip="default" ToolTip="Nome de registro do estado. "
                        Style="text-transform: uppercase;" runat="server" Width="445px" MaxLength="50" />
                </div>
            </div>
            
            <asp:GridView ID="GridEstados" runat="server" AutoGenerateColumns="False"
                CellPadding="4" radius-border="5px" ForeColor="#333333" GridLines="None"
                Width="100%" OnSelectedIndexChanged="GridEstados_SelectedIndexChanged" OnPreRender="GridEstados_PreRender">
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <HeaderStyle CssClass="hStyleBordagrid" />
                <SelectedRowStyle BackColor="#e1e7ef" Font-Bold="True" ForeColor="#333333" />
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EditRowStyle BackColor="#999999" />
                <Columns>
                    <asp:CommandField SelectText="&gt;" ShowSelectButton="True" ButtonType="Button">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                    </asp:CommandField>
                    <asp:BoundField DataField="Codigo" HeaderText="UF">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Regiao" HeaderText="Regi&#227;o">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="70px"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                </Columns>
            </asp:GridView>
        </ContentTemplate>
    </asp:UpdatePanel>
</asp:Content>