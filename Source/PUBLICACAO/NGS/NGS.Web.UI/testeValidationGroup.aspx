<%@ Page Title="" Language="vb" AutoEventWireup="false" MasterPageFile="~/Principal.Master"
    CodeBehind="testeValidationGroup.aspx.vb" Inherits="NGS.Web.UI.testeValidationGroup" %>

<asp:Content ID="Content1" ContentPlaceHolderID="HeadContent" runat="server">
    <style type="text/css">
        .validError, .divWarning
        {
            color: Red;
            font-size: 10px;
            font-weight: bold;
        }
        .divWarning
        {
            padding: 5px;
            margin: 4px 0;
            border: 1px solid #eee0cd;
            border-left-width: 5px;
            border-radius: 3px;
            border-left-color: #aa6708;
        }
    </style>
  </asp:Content>
<asp:Content ID="Content2" ContentPlaceHolderID="MainContent" runat="server">
    <asp:ScriptManager ID="teste" runat='server' />
    <div class="titulodiv">
        Button.ValidationGroup Example
    </div>
    <asp:CustomValidator ID="CustomValidator1" runat="server" ErrorMessage="CustomValidator" ControlToValidate="txtNome" ValidationGroup="form1"></asp:CustomValidator>
    <div class="divWarning">
        <label class="validError">
            Os campos com *(asterístico) são obrigatórios.</label>
    </div>
    <div class="divWarning">
        <asp:ValidationSummary ID="ValidationSummary1" runat="server" ValidationGroup="form1"
            HeaderText="teste" />
    </div>
    <div class="row">
        <div class="collbl">
            <asp:Label ID="lblNome" Text="Informe seu nome:" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:TextBox ID="txtNome" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:RequiredFieldValidator CssClass="validError" ID="rfv1" ControlToValidate="txtNome"
                ValidationGroup="form1" SetFocussOnError="true" ErrorMessage="Informe o Nome"
                Text="*" runat="Server" />
        </div>
    </div>
    <div class="row">
        <div class="collbl">
            <asp:Label ID="Label1" Text="Informe seu nome:" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:TextBox ID="TextBox1" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:RequiredFieldValidator CssClass="validError" ID="RequiredFieldValidator1" ControlToValidate="txtNome"
                ValidationGroup="form1" SetFocussOnError="true" ErrorMessage="Informe o Nome"
                runat="Server">*</asp:RequiredFieldValidator>
        </div>
    </div>
    <div class="row">
        <div class="collbl">
            <asp:Label ID="Label2" Text="Informe seu teste:" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:TextBox ID="TextBox2" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:RequiredFieldValidator CssClass="validError" ID="RequiredFieldValidator4" ControlToValidate="txtNome"
                ValidationGroup="form1" SetFocussOnError="true" ErrorMessage="Informe o Nome"
                runat="Server">*</asp:RequiredFieldValidator>
        </div>
    </div>
    <div class="row">
        <div class="collbl">
            <asp:Label ID="AgeLabel" Text="Informe a Idade:" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:TextBox ID="txtIdade" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:RequiredFieldValidator CssClass="validError" ID="RequiredFieldValidator2" ControlToValidate="txtIdade"
                ValidationGroup="form1" ErrorMessage="Informe a Idade" runat="Server" Style="color: Red;"
                SetFocussOnError="true">*</asp:RequiredFieldValidator>
        </div>
    </div>
    <div class="row">
        <asp:Button ID="Button1" Text="Validate" CausesValidation="true" ValidationGroup="form1"
            runat="Server" />
    </div>
    <div class="row">
        <div class="collbl">
            <asp:Label ID="CityLabel" Text="INforme a Cidade" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:TextBox ID="CityTextbox" runat="Server" />
        </div>
        <div class="coltxt">
            <asp:RequiredFieldValidator ID="RequiredFieldValidator3" ControlToValidate="CityTextBox"
                ValidationGroup="LocationInfoGroup" ErrorMessage="Enter a city name." runat="Server" />
        </div>
    </div>
    <div class="row">
        <!--When Button2 is clicked, only validation
            controls that are a part of LocationInfoGroup
            are validated.-->
        <asp:Button ID="Button2" Text="Validate" CausesValidation="true" ValidationGroup="LocationInfoGroup"
            runat="Server" />
    </div>
</asp:Content>
