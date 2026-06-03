<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="Processos.aspx.vb" Inherits="NGS.Web.UI.Processos" ValidateRequest="false" %>

<asp:Content ID="ScriptContent" ContentPlaceHolderID="HeadContent" runat="server">
    <script type="text/javascript">
        tinymce.init({
            selector: "textarea",
            theme: "modern",
            width: 717,
            height: 380,
            plugins: [
                "advlist autolink lists link image charmap print preview hr anchor pagebreak",
                "searchreplace wordcount visualblocks visualchars code fullscreen",
                "insertdatetime media nonbreaking save table contextmenu directionality",
                "emoticons template paste textcolor colorpicker textpattern"
            ],
            toolbar1: "newdocument fullpage | bold italic underline strikethrough | alignleft aligncenter alignright alignjustify | styleselect formatselect fontselect fontsizeselect",
            toolbar2: "cut copy paste | searchreplace | bullist numlist | outdent indent blockquote | undo redo | link unlink anchor image media code | insertdatetime preview | forecolor backcolor",
            toolbar3: "table | hr removeformat | subscript superscript | charmap emoticons | print fullscreen | ltr rtl | spellchecker | visualchars visualblocks nonbreaking template pagebreak restoredraft",
            image_advtab: true
        });
        $(document).ready(function () {
            $('#<%=txtDescricao.ClientID%>').change(function () {
                $(this).val().toUpperCase();
            });
        });

    </script>
    <style type="text/css">
        #meioconteudo {
            width: 1340px !important;
        }

        .none {
            display: none;
        }

        .menu_acoes .acoes ul li {
            width: 6%;
        }
    </style>
</asp:Content>
<asp:Content ID="BodyContent" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="scmngProcessos" runat="server" />
    <orea:AjaxUpdating ID="ajaxUpdating" runat="server" Text="Aguarde..." />
    <div class="titulodiv">
        <label>
            Processos
        </label>
    </div>
    <div class="menu_acoes">
        <div class="acoes">
            <ul>
                <li class="iconNovo" runat="server">
                    <asp:LinkButton ID="lnkNovo" runat="server">
                        <span>Gravar</span>
                    </asp:LinkButton>
                </li>
                <li class="iconAtualizar" runat="server">
                    <asp:LinkButton ID="lnkAtualizar" runat="server">
                        <span>Atualizar</span>
                    </asp:LinkButton>
                </li>
                <li class="iconExcluir" runat="server">
                    <asp:LinkButton ID="lnkExcluir" runat="server" OnClientClick="if(!confirm('Deseja realmente excluir este registro?')) return false;">
                        <span>Excluir</span>
                    </asp:LinkButton>
                </li>
                <li class="iconLimpar" runat="server">
                    <asp:LinkButton ID="lnkLimpar" runat="server">
                        <span>Limpar</span>
                    </asp:LinkButton>
                </li>
                <li class="iconRelatorio" runat="server">
                    <asp:LinkButton ID="lnkRelatorio" runat="server">
                        <span>Relatório</span>
                    </asp:LinkButton>
                </li>
                <li class="iconAjuda" runat="server">
                    <asp:LinkButton ID="lnkAjuda" runat="server">
                        <span>Ajuda</span>
                    </asp:LinkButton>
                </li>
                <li class="iconAtualizar" runat="server">
                    <asp:LinkButton ID="lnkAtualizarTudo" runat="server">
                        <span>Atualizar Tudo</span>
                    </asp:LinkButton>
                </li>
            </ul>
        </div>
    </div>
    <div class="painelleft" style="width: 600px;">
        <div class="row">
            <div class="collbl">
                Processo:
            </div>
            <div class="coltxt">
                <asp:TextBox ID="txtProcesso" runat="server" Width="450px" data-ToolTip="default"
                    ToolTip="Identificação do processo." />
            </div>
        </div>
        <div class="row">
            <div class="collbl">
                Descrição:
            </div>
            <div class="coltxt">
                <asp:TextBox ID="txtDescricao" TabIndex="2" runat="server" Width="450px" data-ToolTip="default"
                    ToolTip="Descrição do processo." />
            </div>
        </div>
        <div class="bordagrid" style="height: 540px;">
            <asp:GridView ID="GridProcessos" runat="server" AutoGenerateColumns="False" CellPadding="4"
                ForeColor="#333333" GridLines="None" Width="100%" OnSelectedIndexChanged="GridProcessos_SelectedIndexChanged">
                <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" Height="26px" />
                <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                <EditRowStyle BackColor="#999999" />
                <Columns>
                    <asp:CommandField SelectText=" &gt; " ShowSelectButton="True" ButtonType="Button">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="25px"></ItemStyle>
                    </asp:CommandField>
                    <asp:BoundField DataField="Processo" HeaderText="Processo">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left" Width="225px"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Descricao" HeaderText="Descri&#231;&#227;o">
                        <HeaderStyle HorizontalAlign="Left"></HeaderStyle>
                        <ItemStyle HorizontalAlign="Left"></ItemStyle>
                    </asp:BoundField>
                    <asp:BoundField DataField="Manual" ItemStyle-CssClass="none"></asp:BoundField>
                </Columns>
            </asp:GridView>
        </div>
    </div>
    <div class="painelright" style="width: 720px; margin-top: 8px;">
        <asp:TextBox ID="txtManual" runat="server" TextMode="MultiLine" />
    </div>
</asp:Content>
