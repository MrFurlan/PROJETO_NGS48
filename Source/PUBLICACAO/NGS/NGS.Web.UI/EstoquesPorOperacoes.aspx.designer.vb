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


Partial Public Class EstoquesPorOperacoes

    '''<summary>
    '''Controle scmngEstoquesPorOperacoes.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents scmngEstoquesPorOperacoes As Global.System.Web.UI.ScriptManager

    '''<summary>
    '''Controle ajaxUpdating.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ajaxUpdating As Global.Orea.WebControls.AjaxUpdating

    '''<summary>
    '''Controle updpnlEstoquesPorOperacoes.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents updpnlEstoquesPorOperacoes As Global.System.Web.UI.UpdatePanel

    '''<summary>
    '''Controle lnkRelatorio.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkRelatorio As Global.System.Web.UI.WebControls.LinkButton

    '''<summary>
    '''Controle lnkExcelDados.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents lnkExcelDados As Global.System.Web.UI.WebControls.LinkButton

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
    '''Controle ddlUnidade.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlUnidade As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle DdlEmpresa.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents DdlEmpresa As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle DdlDeposito.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents DdlDeposito As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle ucSelecaoProduto.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ucSelecaoProduto As Global.NGS.Web.UI.ucSelecaoProduto

    '''<summary>
    '''Controle rdAtivo.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents rdAtivo As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle rdInativo.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents rdInativo As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle rbPorNome.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents rbPorNome As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle rbPorDescricaoMapa.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents rbPorDescricaoMapa As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle rbTodos.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents rbTodos As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle RadFisico.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents RadFisico As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle RadFiscal.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents RadFiscal As Global.System.Web.UI.WebControls.RadioButton

    '''<summary>
    '''Controle ddlExercicio.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlExercicio As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle CkCDeposito.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents CkCDeposito As Global.System.Web.UI.WebControls.CheckBox

    '''<summary>
    '''Controle chkContraPartida.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents chkContraPartida As Global.System.Web.UI.WebControls.CheckBox

    '''<summary>
    '''Controle chkOperacaoCusto.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents chkOperacaoCusto As Global.System.Web.UI.WebControls.CheckBox

    '''<summary>
    '''Controle chkConsNotasComSerieEspecificas.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents chkConsNotasComSerieEspecificas As Global.System.Web.UI.WebControls.CheckBox

    '''<summary>
    '''Controle TxtDiaMesInicial.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents TxtDiaMesInicial As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle ddlMesInicial.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlMesInicial As Global.System.Web.UI.WebControls.DropDownList

    '''<summary>
    '''Controle TxtDiaMesFinal.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents TxtDiaMesFinal As Global.System.Web.UI.WebControls.TextBox

    '''<summary>
    '''Controle ddlMesFinal.
    '''</summary>
    '''<remarks>
    '''Campo gerado automaticamente.
    '''Para modificar, mova a declaração de campo do arquivo de designer a um arquivo code-behind.
    '''</remarks>
    Protected WithEvents ddlMesFinal As Global.System.Web.UI.WebControls.DropDownList
End Class
