<%@ Page Language="vb" AutoEventWireup="false" CodeBehind="ItensDaLiberacao.aspx.vb"
    Inherits="NGS.Web.UI.ItensDaLiberacao" %>

<%@ Register Assembly="CrystalDecisions.Web, Version=13.0.2000.0, Culture=neutral, PublicKeyToken=692fbea5521e1304"
    Namespace="CrystalDecisions.Web" TagPrefix="CR" %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>NGS Soluções Ltda</title>
    <link href="Styles/menucentral.css" rel="stylesheet" type="text/css" />
</head>
<body>
    <form id="form1" runat="server">
    <div>
        <%--<CR:CrystalReportViewer ID="Crv_Liberacoes" runat="server" AutoDataBind="true" DisplayGroupTree="False" />--%>
        <CR:CrystalReportViewer ID="Crv_Liberacoes" runat="server" AutoDataBind="true" />
    </div>
    </form>
</body>
</html>
