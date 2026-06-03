<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="Mensagem.aspx.vb" Inherits="NGS.Web.UI.Mensagem" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head id="Head1" runat="server">
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
    <title>Atenção</title>
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <asp:ScriptManager ID="ScriptManager1" runat="server">
        </asp:ScriptManager>
    </div>
    <asp:UpdatePanel ID="UpdatePanel1" runat="server">
        <ContentTemplate>
            <asp:TextBox ID="txtMensagem" runat="server" Height="200px" TextMode="MultiLine"
                Width="480px"/>
            <asp:Button ID="cmdOk" runat="server" Text="Ok" Width="96px" OnClick="cmdOk_Click" />
        </ContentTemplate>
    </asp:UpdatePanel>
    </form>
</body>
</html>
