'------------------------------------------------------------------------------
' <gerado automaticamente>
'     Esse código foi gerado por uma ferramenta.
'
'     As alterações ao arquivo poderão causar comportamento incorreto e serão perdidas se
'     o código for recriado
' </gerado automaticamente>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Partial Public Class AjustarPedido

    '''<summary>
    '''Controle scrmngPedidosXItens.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents scrmngPedidosXItens As Global.System.Web.UI.ScriptManager

    '''<summary>
    '''Controle ajaxUpdating.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ajaxUpdating As Global.Orea.WebControls.AjaxUpdating

    '''<summary>
    '''Controle updpnlPedidosXItens.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents updpnlPedidosXItens As Global.System.Web.UI.UpdatePanel

    '''<summary>
    '''Controle HID.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents HID As Global.System.Web.UI.WebControls.HiddenField

    '''<summary>
    '''Controle lnkConfirmar.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkConfirmar As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''Controle lnkConsultar.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkConsultar As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''Controle lnkLimpar.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkLimpar As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''Controle lnkAjuda.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkAjuda As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''Controle ddlUnidadeDeNegocio.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlUnidadeDeNegocio As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle ddlEmpresa.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlEmpresa As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle txtCodigoCliente.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents txtCodigoCliente As Global.System.Web.UI.WebControls.HiddenField

    '''<summary>
    '''Controle txtCliente.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents txtCliente As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle cmdConsultarCliente.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents cmdConsultarCliente As Global.System.Web.UI.WebControls.Button

    '''<summary>
    '''Controle txtPedido.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents txtPedido As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle txtAnterior.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents txtAnterior As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle txtAtual.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents txtAtual As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle ddlOperacao.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlOperacao As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle ddlSubOperacao.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlSubOperacao As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle tcPedido.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents tcPedido As Global.AjaxControlToolkit.TabContainer

    '''<summary>
    '''Controle TabItens.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents TabItens As Global.AjaxControlToolkit.TabPanel

    '''<summary>
    '''Controle grdItens.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents grdItens As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''Controle gridEncargosGerais.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents gridEncargosGerais As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''Controle TabVencimentos.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents TabVencimentos As Global.AjaxControlToolkit.TabPanel

    '''<summary>
    '''Controle grdTitulos.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents grdTitulos As Global.System.Web.UI.WebControls.GridView

    '''<summary>
    '''Controle ucConsultaClientes.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ucConsultaClientes As Global.NGS.Web.UI.ucConsultaClientes

    '''<summary>
    '''Controle ucPedidoxEncargo.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ucPedidoxEncargo As Global.NGS.Web.UI.ucPedidoEncargo

    '''<summary>
    '''Controle ucConsultaPedidos.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ucConsultaPedidos As Global.NGS.Web.UI.ucConsultaPedidos
End Class
