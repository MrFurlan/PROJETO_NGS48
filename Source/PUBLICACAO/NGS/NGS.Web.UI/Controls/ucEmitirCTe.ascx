<%@ Control Language="vb" AutoEventWireup="false" CodeBehind="ucEmitirCTe.ascx.vb"
    Inherits="NGS.Web.UI.ucEmitirCTe" %>

<div id="divEmitirCTe" class="uc" title="Emitir CTe" style="display: none;">
    <style type="text/css">
        .layout {
            display: flex;
            align-items: flex-start;
            gap: 10px;
            width: 100%;
        }
    
        .left-main {
            flex: 1;
            min-width: 0;
        }
        .right-main {
            width: 450px;
            flex-shrink: 0;
    }
    </style>
    <asp:UpdatePanel ID="updpnlEmitirCTe" runat="server">
        <ContentTemplate>
            <asp:HiddenField ID="HID" runat="server" />
            <ajaxToolkit:TabContainer ID="tabDados" runat="server" ActiveTabIndex="0" Width="100%">
                <ajaxToolkit:TabPanel ID="tabEmbarque" runat="server" HeaderText="Cadastro">
                    <HeaderTemplate>
                        Dados Cadastrais
                    </HeaderTemplate>
                    <ContentTemplate>
                        <div class="menu_acoes">
                            <div class="acoes">
                                <ul>
                                    <li runat="server" class="iconNovo">
                                        <asp:LinkButton ID="lnkConfirmar" Text="Confirmar" runat="server" />
                                    </li>
                                    <li runat="server" class="iconLimpar">
                                        <asp:LinkButton ID="lnkLimpar" Text="Limpar" runat="server" />
                                    </li>
                                    <li runat="server" class="iconSair">
                                        <asp:LinkButton ID="lnkFechar" Text="Fechar" runat="server" />
                                    </li>
                                </ul>
                            </div>
                        </div>
                        <div class="layout">                  
                            <div class="left-main">
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>REMETENTE</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" width="100%" valign="top">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                <asp:Label ID="txtRemetenteNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                                    Font-Names="Tahoma" Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtRemetenteInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>DESTINATÁRIO</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" width="100%">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                <asp:Label ID="txtDestinatarioNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                                    Font-Names="Tahoma" Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtDestinatarioInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>TOMADOR</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" width="100%">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                <asp:Label ID="txtTomadorNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                <asp:Button ID="btnTomador" runat="server" Text=" > " UseSubmitBehavior="False" Visible="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtTomadorInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <div class="bordagrid" style="height: 150px;">
                                    <asp:GridView ID="gridNFe" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                        ForeColor="#333333" GridLines="None" Width="100%">
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <EditRowStyle BackColor="#999999" />
                                        <Columns>
                                            <asp:BoundField DataField="Nota" HeaderText="Nota">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="TipoDocumento" HeaderText="Tipo">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CodigoOperacao" HeaderText="Operação">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="Produto" HeaderText="Produto">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="QuantidadeFisica" HeaderText="Físico" DataFormatString="{0:N0}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="QuantidadeFiscal" HeaderText="Fiscal" DataFormatString="{0:N0}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorUnitario" HeaderText="Unit&#225;rio" DataFormatString="{0:N10}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                            <asp:BoundField DataField="ValorTotal" HeaderText="Total" DataFormatString="{0:N2}">
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:BoundField>
                                        </Columns>
                                    </asp:GridView>
                                </div>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>LOCAL DE EMBARQUE/COLETA</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" width="100%">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                <asp:Label ID="txtEmbarqueNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                <asp:Button ID="btnEmbarque" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEmbarqueInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>LOCAL DE ENTREGA/DESCARGA</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" width="100%">
                                        <tr>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                                <asp:Label ID="txtEntregaNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                                    Font-Names="Tahoma" Font-Italic="False" />
                                                <asp:Button ID="btnEntrega" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                            <td>
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                                <asp:Label ID="txtEntregaInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                                    Font-Italic="False" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <%--<strong><em>REMETENTE</em></strong>
                                <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                <asp:Label ID="txtRemetenteNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                    Font-Names="Tahoma" Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtRemetenteInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <strong><em>DESTINATÁRIO</em></strong>
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                <asp:Label ID="txtDestinatarioNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                    Font-Names="Tahoma" Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtDestinatarioEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtDestinatarioComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtDestinatarioBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                     <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                 <asp:Label ID="txtDestinatarioCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtDestinatarioCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtDestinatarioInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <strong><em>TOMADOR</em></strong>
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                <asp:Label ID="txtTomadorNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                    Font-Names="Tahoma" Font-Italic="False" />
                                <asp:Button ID="btnTomador" runat="server" Text=" > " UseSubmitBehavior="False" Visible="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtTomadorInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <strong><em>NFe</em></strong>
                                <asp:GridView ID="gridNFe" runat="server" AutoGenerateColumns="False" CellPadding="4"
                                    ForeColor="#333333" GridLines="None" style="table-layout:fixed;">
                                    <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                    <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                    <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                    <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                    <EditRowStyle BackColor="#999999" />
                                    <Columns>
                                        <asp:BoundField DataField="Nota" HeaderText="Nota">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="TipoDocumento" HeaderText="Tipo">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CodigoOperacao" HeaderText="Operação">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="Produto" HeaderText="Produto">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="CFOP" HeaderText="CFOP">
                                            <HeaderStyle HorizontalAlign="Left" />
                                            <ItemStyle HorizontalAlign="Left" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="QuantidadeFisica" HeaderText="Físico" DataFormatString="{0:N0}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="QuantidadeFiscal" HeaderText="Fiscal" DataFormatString="{0:N0}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorUnitario" HeaderText="Unit&#225;rio" DataFormatString="{0:N10}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                        <asp:BoundField DataField="ValorTotal" HeaderText="Total" DataFormatString="{0:N2}">
                                            <HeaderStyle HorizontalAlign="Right" />
                                            <ItemStyle HorizontalAlign="Right" />
                                        </asp:BoundField>
                                    </Columns>
                                </asp:GridView>
                                <strong><em>LOCAL DE EMBARQUE/COLETA</em></strong>
                                <label class="titulo">
                                <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                <asp:Label ID="txtEmbarqueNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                    Font-Names="Tahoma" Font-Italic="False" />
                                <asp:Button ID="btnEmbarque" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEmbarqueInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <strong><em>LOCAL DE ENTREGA/DESCARGA</em></strong>
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">&nbsp;NOME/RAZÃO SOCIAL</span></strong></label><br />
                                <asp:Label ID="txtEntregaNome" runat="server" ForeColor="Blue" Font-Bold="False"
                                    Font-Names="Tahoma" Font-Italic="False" />
                                <asp:Button ID="btnEntrega" runat="server" Text=" > " UseSubmitBehavior="False" Enabled="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">ENDEREÇO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaEndereco" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">COMPLEMENTO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaComplemento" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">BAIRRO</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaBairro" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaCidade" runat="server" ForeColor="Blue" Font-Names="Tahoma" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">CNPJ</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaCnpj" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                <asp:Label ID="txtEntregaInscricao" runat="server" ForeColor="Blue" Font-Names="Tahoma"
                                    Font-Italic="False" />
                            --%>
                            </div> 
                            <div class="right-main">
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>BUSCAR NF-e</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" border="0">
                                        <tr>
                                            <td valign="top" align="left">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 6pt;">INFORME A CHAVE DA NF-e</span></strong>
                                                </label>
                                                <br />
                                                <asp:TextBox ID="txtNFe" runat="server" Width="305px" Enabled="False" MaxLength="50"
                                                    BorderColor="Red" />
                                                <asp:LinkButton ID="btnConsultarNFe" CssClass="lnk" Enabled="False"
                                                    data-tooltip="default" ToolTip="Consultar Chave NFE" runat="server" Text="
                                                        &lt;i class=&quot;fa fa-arrow-right seta&quot;&gt;&lt;/i&gt;" OnClick="btnConsultarNFe_Click"></asp:LinkButton>
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>INFORMAÇÕES E VALORES</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" border="0" width="97%">
                                        <tr>
                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 10pt; color: black;">PESO</span></strong>
                                                </label>
                                            </td>
                                            <td align="right">
                                                <asp:TextBox ID="txtQuantidade" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                                    CssClass="numerico9" Font-Size="14pt" BorderColor="Blue" Style="text-align: right;" Enabled="False" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 10pt; color: red;">TARIFA (TON)</span></strong>
                                                </label>
                                            </td>
                                            <td align="right">
                                                <asp:TextBox ID="txtUnitario" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                                    AutoPostBack="True" CssClass="txtDecimal4"
                                                    Font-Size="14pt" BorderColor="Blue" Style="color: red; text-align: right;" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 10pt; color: black;">TOTAL</span></strong>
                                                </label>
                                            </td>
                                            <td align="right">
                                                <asp:Label ID="txtValor" runat="server" Font-Bold="False" Font-Size="14pt" Font-Names="Tahoma"
                                                    Height="24px" BorderColor="Blue" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 10pt; color: black;">OPERAÇÃO</span></strong></label>
                                            </td>
                                            <td align="right">
                                                <asp:Label ID="txtOperacao" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                                    Height="36px" Style="padding-right: 5px; vertical-align: middle;" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 115px">&nbsp;</td>
                                            <td align="right">
                                                <asp:Label ID="txtEntSai" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma" Style="padding-right: 5px; vertical-align: middle;" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td align="left" style="width: 115px; padding-left: 10px;">
                                                <label class="titulo">
                                                    <strong><span style="font-size: 10pt; color: black;">CFOP</span></strong></label>
                                            </td>
                                            <td align="right">
                                                <asp:Label ID="txtCfop" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                                    Height="24px" Style="padding-right: 5px;" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>DADOS DO CT-e</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" border="0">
                                        <tr>
                                            <td style="font-size: 7pt; font-family: Tahoma" width: 70px; align="left">
                                                <strong>PAGAMENTO DO FRETE:</strong>
                                            </td>
                                            <td>
                                                <asp:DropDownList ID="ddlFrete" runat="server" AutoPostBack="True" Enabled="False" Width="200px" />
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="font-size: 7pt; font-family: Tahoma; width: 70px;" align="left">
                                                <strong>EMITE CONTRATO DE FRETE:</strong>
                                            </td>
                                            <td>
                                                <table cellpadding="0">
                                                    <tr>
                                                        <td style="border: 0;">
                                                            <asp:RadioButton ID="rdContratoSim" runat="server" Text="SIM" Enabled="False" GroupName="grContrato"
                                                                AutoPostBack="True" />
                                                        </td>
                                                        <td style="border: 0;">
                                                            <asp:RadioButton ID="rdoContratoNao" runat="server" Text="NÃO" Enabled="False" GroupName="grContrato"
                                                                AutoPostBack="True" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="font-size: 7pt; font-family: Tahoma; width: 70px;" align="left">
                                                <strong>TEM PEDÁGIO:</strong>
                                            </td>
                                            <td>
                                                <table cellpadding="0">
                                                    <tr>
                                                        <td style="border: 0;">
                                                            <asp:RadioButton ID="rdPedagioSim" runat="server" Text="SIM" Enabled="False" GroupName="grPedagio"
                                                                AutoPostBack="True" />
                                                        </td>
                                                        <td style="border: 0;">
                                                            <asp:RadioButton ID="rdPedagioNao" runat="server" Text="NÃO" Enabled="False" GroupName="grPedagio"
                                                                AutoPostBack="True" />
                                                        </td>
                                                    </tr>
                                                </table>
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="font-size: 7pt; font-family: Tahoma; width: 70px;" align="left">
                                                <strong>TIPO DE CONHECIMENTO:</strong>
                                            </td>
                                            <td>
                                                <asp:DropDownList ID="ddlTipoCTe" runat="server" AutoPostBack="True" Enabled="False" Width="200px" />
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <fieldset style="border: 1px solid #BBBBBB;">
                                    <legend style="font-size: 0.8em;"><strong><em>ENCARGOS</em></strong></legend>
                                    <table cellspacing="0" cellpadding="0" border="0" width="100%">
                                        <tr>
                                            <td>
                                                <asp:Panel ID="pnlEncargos" runat="server" CssClass="bordagrid" Height="100px"
                                                    ScrollBars="Vertical" Width="100%">
                                                    <asp:GridView ID="gridEncargos" runat="server" CellPadding="3" ForeColor="#333333"
                                                        GridLines="None" Width="100%" AutoGenerateColumns="False">
                                                        <EditRowStyle BackColor="#999999" />
                                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                                        <Columns>
                                                            <asp:BoundField HeaderText="Descri&#231;&#227;o" DataField="Codigo">
                                                                <HeaderStyle HorizontalAlign="Left" />
                                                                <ItemStyle HorizontalAlign="Left" />
                                                            </asp:BoundField>
                                                            <asp:TemplateField HeaderText="Valor">
                                                                <ItemTemplate>
                                                                    <asp:TextBox ID="txtValorEncargo" Style="text-align: right" CssClass="txtDecimal"
                                                                        Text='<%# Eval("Valor") %>' runat="server" BorderColor="White" Width="60px" Enabled="false" />
                                                                </ItemTemplate>
                                                                <HeaderStyle HorizontalAlign="Right" />
                                                                <ItemStyle HorizontalAlign="Right" />
                                                            </asp:TemplateField>
                                                            <asp:TemplateField>
                                                                <ItemTemplate>
                                                                    <asp:Label ID="lblSinal" runat="server" Text='<%# Eval("Sinal") %>' />
                                                                </ItemTemplate>
                                                                <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                                            </asp:TemplateField>
                                                        </Columns>
                                                    </asp:GridView>
                                                </asp:Panel>
                                            </td>
                                        </tr>
                                    </table>
                                </fieldset>
                                <%--<strong><em>BUSCAR NF-e</em></strong>
                                <label class="titulo">
                                    <strong><span style="font-size: 6pt;">INFORME A CHAVE DA NF-e</span></strong>
                                </label>
                                <br />
                                <asp:TextBox ID="txtNFe" runat="server" Enabled="False" MaxLength="50"
                                    BorderColor="Red" />
                                <asp:LinkButton ID="btnConsultarNFe" CssClass="lnk" Enabled="False"
                                    data-tooltip="default" ToolTip="Consultar Chave NFE" runat="server" Text="
                                        &lt;i class=&quot;fa fa-arrow-right seta&quot;&gt;&lt;/i&gt;" OnClick="btnConsultarNFe_Click"></asp:LinkButton>
                                <strong><em>INFORMAÇÕES E VALORES</em></strong>
                                <label class="titulo">
                                    <strong><span style="font-size: 10pt; color: black;">PESO</span></strong>
                                </label>
                                <asp:TextBox ID="txtQuantidade" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                    CssClass="numerico9" Font-Size="14pt" BorderColor="Blue" Style="text-align: right;" Enabled="False" />
                                <label class="titulo">
                                    <strong><span style="font-size: 10pt; color: red;">TARIFA (TON)</span></strong>
                                </label>
                                <asp:TextBox ID="txtUnitario" runat="server" BorderStyle="None" Width="195px" Height="24px"
                                    AutoPostBack="True" CssClass="txtDecimal4"
                                    Font-Size="14pt" BorderColor="Blue" Style="color: red; text-align: right;" />
                                <label class="titulo">
                                    <strong><span style="font-size: 10pt; color: black;">TOTAL</span></strong>
                                </label>
                                <asp:Label ID="txtValor" runat="server" Font-Bold="False" Font-Size="14pt" Font-Names="Tahoma"
                                    Height="24px" BorderColor="Blue" />
                                <label class="titulo">
                                    <strong><span style="font-size: 10pt; color: black;">OPERAÇÃO</span></strong></label>
                                <asp:Label ID="txtOperacao" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                    Height="36px" Style="padding-right: 5px; vertical-align: middle;" />
                                <asp:Label ID="txtEntSai" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma" Style="padding-right: 5px; vertical-align: middle;" />
                                <label class="titulo">
                                    <strong><span style="font-size: 10pt; color: black;">CFOP</span></strong></label>
                                <asp:Label ID="txtCfop" runat="server" Font-Bold="False" Font-Size="7pt" Font-Names="Tahoma"
                                    Height="24px" Style="padding-right: 5px;" />
                                <strong><em>DADOS DO CT-e</em></strong>
                                <strong>PAGAMENTO DO FRETE:</strong>
                                <asp:DropDownList ID="ddlFrete" runat="server" AutoPostBack="True" Enabled="False" Width="200px" />
                                <strong>EMITE CONTRATO DE FRETE:</strong>
                                <asp:RadioButton ID="rdContratoSim" runat="server" Text="SIM" Enabled="False" GroupName="grContrato"
                                    AutoPostBack="True" />
                                <asp:RadioButton ID="rdoContratoNao" runat="server" Text="NÃO" Enabled="False" GroupName="grContrato"
                                    AutoPostBack="True" />
                                <strong>TEM PEDÁGIO:</strong>
                                <asp:RadioButton ID="rdPedagioSim" runat="server" Text="SIM" Enabled="False" GroupName="grPedagio"
                                    AutoPostBack="True" />
                                <asp:RadioButton ID="rdPedagioNao" runat="server" Text="NÃO" Enabled="False" GroupName="grPedagio"
                                    AutoPostBack="True" />
                                <strong>TIPO DE CONHECIMENTO:</strong>
                                <asp:DropDownList ID="ddlTipoCTe" runat="server" AutoPostBack="True" Enabled="False" Width="200px" />
                                <strong><em>ENCARGOS</em></strong>
                                <asp:Panel ID="pnlEncargos" runat="server" CssClass="bordagrid" Height="100px"
                                    ScrollBars="Vertical" Width="100%">
                                    <asp:GridView ID="gridEncargos" runat="server" CellPadding="3" ForeColor="#333333"
                                        GridLines="None" Width="100%" AutoGenerateColumns="False">
                                        <EditRowStyle BackColor="#999999" />
                                        <RowStyle BackColor="#F7F6F3" ForeColor="#333333" />
                                        <FooterStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <PagerStyle BackColor="#284775" ForeColor="White" HorizontalAlign="Center" />
                                        <SelectedRowStyle BackColor="#E2DED6" Font-Bold="True" ForeColor="#333333" />
                                        <HeaderStyle BackColor="#5D7B9D" Font-Bold="True" ForeColor="White" />
                                        <AlternatingRowStyle BackColor="White" ForeColor="#284775" />
                                        <Columns>
                                            <asp:BoundField HeaderText="Descri&#231;&#227;o" DataField="Codigo">
                                                <HeaderStyle HorizontalAlign="Left" />
                                                <ItemStyle HorizontalAlign="Left" />
                                            </asp:BoundField>
                                            <asp:TemplateField HeaderText="Valor">
                                                <ItemTemplate>
                                                    <asp:TextBox ID="txtValorEncargo" Style="text-align: right" CssClass="txtDecimal"
                                                        Text='<%# Eval("Valor") %>' runat="server" BorderColor="White" Width="60px" Enabled="false" />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Right" />
                                                <ItemStyle HorizontalAlign="Right" />
                                            </asp:TemplateField>
                                            <asp:TemplateField>
                                                <ItemTemplate>
                                                    <asp:Label ID="lblSinal" runat="server" Text='<%# Eval("Sinal") %>' />
                                                </ItemTemplate>
                                                <HeaderStyle HorizontalAlign="Center" Width="30px" />
                                                <ItemStyle HorizontalAlign="Center" Width="30px" />
                                            </asp:TemplateField>
                                        </Columns>
                                    </asp:GridView>
                                </asp:Panel>--%>
                            </div>
                        </div>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
                <ajaxToolkit:TabPanel ID="tabTransporte" runat="server" HeaderText="Transporte">
                    <HeaderTemplate>
                        Informações de Transporte
                    </HeaderTemplate>
                    <ContentTemplate>
                        <fieldset style="border: 1px solid #BBBBBB;">
                            <legend style="font-size: 0.8em;"><strong><em>TRANSPORTADOR</em></strong></legend>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;RAZÃO SOCIAL</span></strong></label>&nbsp;<br />
                                        <asp:Button ID="btnTransportador" runat="server" Text=" > " UseSubmitBehavior="False" />
                                        <asp:Label ID="txtTransportadorNome" runat="server" Width="260px" ForeColor="Blue"
                                            Font-Bold="False" Font-Names="Tahoma" />&nbsp;&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;CÓDIGO ANTT(CIOT)</span></strong></label><span
                                                style="font-size: 5pt">&nbsp;<br />
                                            </span>&nbsp;<asp:Label ID="txtTransportadorRNTRC" runat="server" Font-Names="Tahoma"
                                                ForeColor="Blue" />
                                    </td>
                                </tr>
                            </table>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtTransportadorEndereco" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">CIDADE/UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtTransportadorCidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">CNPJ/CPF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtTransportadorCnpj" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">INSCRIÇÃO ESTADUAL</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtTransportadorInscricao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                        <fieldset style="border: 1px solid #BBBBBB;">
                            <legend style="font-size: 0.8em;"><strong><em>DADOS DO CAMINHÃO</em></strong></legend>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;TRATOR (PLACA 1)</span></strong></label>&nbsp;<br />
                                        <asp:Button ID="btnPlacaTrator" runat="server" Text=" > " OnClick="btnPlacaTrator_Click" UseSubmitBehavior="False" />
                                        <asp:Label ID="txtCPlaca" runat="server" Width="230px" ForeColor="Blue" Font-Bold="False"
                                            Font-Names="Tahoma" />&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlacaCidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlacaEstado" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlacaRNTRC" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlacaProprietario" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;CAVALO(PLACA 2)</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtCPlaca02" runat="server" Width="260px" ForeColor="Blue" Font-Bold="False"
                                            Font-Names="Tahoma" />&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca2Cidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca2Estado" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca2RNTRC" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca2Proprietario" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;CAVALO(PLACA 2)</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtCPlaca03" runat="server" Width="260px" ForeColor="Blue" Font-Bold="False"
                                            Font-Names="Tahoma" />&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca3Cidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca3Estado" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca3RNTRC" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca3Proprietario" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;CAVALO(PLACA 4)</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtCPlaca04" runat="server" Width="260px" ForeColor="Blue" Font-Bold="False"
                                            Font-Names="Tahoma" />&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca4Cidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca4Estado" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;RNTRC</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca4RNTRC" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;PROPRIETÁRIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtPlaca4Proprietario" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                        <fieldset style="border: 1px solid #BBBBBB;">
                            <legend style="font-size: 0.8em;"><strong><em>DADOS DO MOTORISTA</em></strong></legend>
                            <table cellspacing="0" cellpadding="0" width="100%">
                                <tr>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;NOME</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaNome" runat="server" Width="260px" ForeColor="Blue"
                                            Font-Bold="False" Font-Names="Tahoma" />&nbsp;
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;ENDEREÇO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaEndereco" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;MUNICÍPIO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaCidade" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;UF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaEstado" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;CPF</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaCPF" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                    <td>
                                        <label class="titulo">
                                            <strong><span style="font-size: 6pt">&nbsp;HABILITAÇÃO</span></strong></label>&nbsp;<br />
                                        <asp:Label ID="txtMotoristaHabilitacao" runat="server" ForeColor="Blue" Font-Bold="False" />
                                    </td>
                                </tr>
                            </table>
                        </fieldset>
                    </ContentTemplate>
                </ajaxToolkit:TabPanel>
            </ajaxToolkit:TabContainer>
        </ContentTemplate>
    </asp:UpdatePanel>
</div>
