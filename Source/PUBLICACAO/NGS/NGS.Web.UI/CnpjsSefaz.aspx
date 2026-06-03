<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="CnpjsSefaz.aspx.vb" Inherits="NGS.Web.UI.CnpjsSefaz" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <link href="http://cdn.datatables.net/1.10.4/css/jquery.dataTables.css" rel="stylesheet"
        type="text/css" />
    <link href="http://code.jquery.com/ui/1.10.3/themes/smoothness/jquery-ui.css" rel="stylesheet"
        type="text/css" />
    <link href="http://cdn.datatables.net/plug-ins/3cfcc339e89/integration/jqueryui/dataTables.jqueryui.css" rel="stylesheet" />

    <style type="text/css">
        div.dataTables_length, div.dataTables_filter {
            margin: 5px;
        }

        div.dataTables_filter {
            margin-right: 2px;
        }

        div.dataTables_wrapper {
            border: 1px solid #5D7B9D;
            border-radius: 5px;
            color: #5D7B9D;
            margin-top: 5px;
        }

        .paginate_button {
            border-radius: 5px;
        }

        .dataTables_paginate {
            padding: 0.25em;
        }

        dataTables_info {
            padding: 0.755em;
        }
    </style>

    <script src="http://cdn.datatables.net/1.10.4/js/jquery.dataTables.min.js" type="text/javascript"></script>
    <script type="text/javascript">
        function pageLoadCnpjSefaz() {
            $('#MainContent_GridCnpjSefaz').DataTable({
                stateSave: true,
                "language": {
                    "lengthMenu": "_MENU_ Registros por páginas",
                    "zeroRecords": "desculpe - nenhum registro encontrado...",
                    "info": "Mostrando _PAGE_ a _PAGES_",
                    "infoEmpty": "não avaliado",
                    "infoFiltered": "(filtrado _MAX_ Total de registros)",
                    "search": "Pesquisa:",
                    "paginate": {
                        "next": "Próxima",
                        "previous": "Anterior"
                    }
                }
            });

            $("#resultado").click(function () {
                $.ajax({
                    type: "POST",
                    url: "/CnpjsSefaz.aspx/MeuMetodo",
                    data: "{'desc': '" + $("#resultado").text() + "'}",
                    contentType: "application/json; charset=utf-8",
                    dataType: "json",
                    success: function (msg) {
                        $("#resultado").text(msg.d);
                    },
                    error: function () {
                        alert("Falha ao carregar dados");
                    }
                });
            });

        }

        $(document).ready(function () {
            pageLoadCnpjSefaz();

            var prmCnpjSefaz = Sys.WebForms.PageRequestManager.getInstance();
            prmCnpjSefaz.add_endRequest(pageLoadCnpjSefaz);
        });

    </script>

</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngCnpjsSefaz" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <asp:UpdatePanel ID="updpnlCnpjsSefaz" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <div class="titulodiv">
                Cadastro do CNPJs da Sefaz
            </div>
            <div class="menu_acoes">
                <div class="acoes">
                    <ul>
                        <li runat="server">
                            <asp:LinkButton class="iconNovo" ID="lnkNovo" Text="Gravar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAtualizar" ID="lnkAtualizar" Text="Atualizar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconExcluir" ID="lnkExcluir" Text="Excluir" runat="server"
                                OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconLimpar" ID="lnkLimpar" Text="Limpar" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconRelatorio" ID="lnkRelatorio" Text="Relatório" runat="server" />
                        </li>
                        <li runat="server">
                            <asp:LinkButton class="iconAjuda" ID="lnkAjuda" Text="Ajuda" runat="server" />
                        </li>
                    </ul>
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    CNPJ:
                </div>
                <div class="coltxt">
                    <asp:TextBox  ID="txtCnpj" CausesValidation="false" autofocus CssClass="txtNumerico" placeholder="CNPJ" runat="server" Width="93px" AutoPostBack="True" data-ToolTip="default" ToolTip="Código CNPJ da Sefaz." MaxLength="18" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Estado:
                </div>
                <div class="coltxt">
                        <asp:DropDownList ID="ddlEstado" runat="server" Width="265px" AutoPostBack="True" />
                </div>
            </div>
            <div class="row">
                <div class="collbl">
                    Cidade:
                </div>
                <div class="coltxt">
                    <asp:TextBox ID="txtCidade" runat="server" MaxLength="50" Width="255px" Enabled="False" data-ToolTip="default" ToolTip="Localização da Sefaz dentro do Estado." />
                </div>
            </div>
            <asp:GridView ID="GridCnpjSefaz" runat="server" AutoGenerateColumns="False"
                CellPadding="4" ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridCnpjSefaz_SelectedIndexChanged"
                OnPreRender="GridCnpjSefaz_PreRender">
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <HeaderStyle CssClass="hStyleBordagrid" />
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EditRowStyle BackColor="#999999" />
                <Columns>
                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                    </asp:CommandField>
                    <asp:BoundField DataField="CNPJ" HeaderText="CNPJ">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="60px"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="ESTADO" HeaderText="Estado">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Cidade" HeaderText="Cidade">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                </Columns>
            </asp:GridView>
        </ContentTemplate>
    </asp:UpdatePanel>
    <uc:ConsultaCodMunicipios ID="ucConsultaCodMunicipios" runat="server" />
</asp:Content>
