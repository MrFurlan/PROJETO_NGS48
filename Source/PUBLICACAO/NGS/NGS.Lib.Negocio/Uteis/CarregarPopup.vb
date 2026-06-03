Imports Microsoft.VisualBasic
Imports System.Web
Imports System.Web.UI
Imports System.Web.UI.WebControls

Public Class CarregarPopup

    'M�TODOS PARA CHAMAR POPUP VIA JQUERY DIALOG

    Public Sub ShowMask(ByRef page As System.Web.UI.Control, ByVal mask As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('.divLoading').mask('" & mask & "');"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ShowMask(ByRef page As System.Web.UI.Control, ByVal mask As String, ByVal divID As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & divID & "').mask('" & mask & "');"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub HideMask(ByRef page As System.Web.UI.Control, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('.divLoading').unmask();"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub HideClose(ByRef page As System.Web.UI.Control, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('.ui-dialog-titlebar-close').hide();"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub MoveTo(ByRef page As System.Web.UI.Control, ByVal divID As String)
        'Dim strJavaScript As String = String.Empty
        'strJavaScript = "$(document).ready(function() { "
        Dim strJavaScript As String = "$(document).ready(function() { "
        strJavaScript &= "function scroll_to(div) { "
        strJavaScript &= "$('html, body').animate({ "
        strJavaScript &= "scrollTop: $(div).offset().top "
        strJavaScript &= "}, 100); "
        strJavaScript &= "} "
        strJavaScript &= "scroll_to('#" & divID & "'); "
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub SetTitle(ByRef page As System.Web.UI.Control, ByVal divID As String, ByVal title As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & divID & "').prop('title', '" & title & "') "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub SetFocus(ByRef page As System.Web.UI.Control, ByVal controlID As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & controlID & "').focus() "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Function ShowDialog(ByVal divID As String, ByVal width As Integer, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100) As String
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$('#" & divID & "').dialog({"
        strJavaScript &= "autoOpen: true,"
        strJavaScript &= "modal: true,"
        strJavaScript &= "resizable: false,"
        strJavaScript &= "closeOnEscape: false, "
        strJavaScript &= "width: " & width & ","
        strJavaScript &= "hide: 'explode',"
        strJavaScript &= "open: function (type, data) {"
        strJavaScript &= "$(this).parent().appendTo('form');"
        strJavaScript &= "$('.ui-dialog-titlebar-close').hide();"
        If Not (String.IsNullOrWhiteSpace(control)) Then
            strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
            strJavaScript &= "$('#" + (control) + "', '#" + (divID) + "').focus();"
            strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        End If
        strJavaScript &= "},"
        strJavaScript &= "close: function () { $(this).dialog('destroy'); }"
        strJavaScript &= "});"
        strJavaScript &= "$('#" & (divID) & "').dialog('option', 'position', 'center');"
        strJavaScript &= "$('#" & (divID) & "').css('min-height', '50px');"
        Return strJavaScript
    End Function

    Public Sub CloseDialog(ByRef page As System.Web.UI.Control, ByVal divID As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & divID & "').dialog('close');"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub DestroyDialog(ByRef page As System.Web.UI.Control, ByVal divID As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & divID & "').dialog('destroy');"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CenterDialog(ByRef page As System.Web.UI.Control, ByVal divID As String, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$('#" & (divID) & "').dialog('option', 'position', 'center');"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub SubmitForm(ByRef page As System.Web.UI.Control, Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "document.forms[0].submit();"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeAdiantamentos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaAdiantamento", 1100, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeAutorizacaoDeRetirada(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaAutorizacaoDeRetirada", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeCarteirasDeTitulo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCarteirasDeTitulo", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeCep(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCep", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeClientes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCliente", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CadastroDeProcessos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCadastrarProcesso", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CadastrarTipoDeCertidao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCadastrarTipoDeCertidao", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CadastroDeGrupos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCadastrarGrupo", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeMunicipios(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCodMunicipios", 600, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeDadosBancarios(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaDadosBancarios", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub BaixaLoteFinanceiro(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divBaixaLoteFinanceiro", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub


    Public Sub ConsultaDeEmpresas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaEmpresas", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub TrocaDeEmpresa(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divTrocaEmpresa", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeEstados(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaEstados", 450, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeFaturasDeFrete(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaFaturasDeFrete", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeFinalidades(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaFinalidades", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeItensDoPedido(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaItensDoPedido", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeMemorandoEquiparado(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaMemorandoEquiparado", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotasDeFrete(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotasDeFrete", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotasXEncargos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotasXEncargos", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotaTroca(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotaTroca", 1300, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotaVendaAOrdem(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotaVendaAOrdem", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaNFReferencialSaida(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNFReferencialSaida", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaNFProdutor(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNFProdutor", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeObservacoes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaObservacoes", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeObservacoesEmbarque(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaObservacoesEmbarque", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeOperacoes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaOperacoes", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaPedidoDeTroca(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPedidoDeTroca", 1300, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePedidos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPedidos", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePedidoxSaldo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPedidoxSaldo", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePlacas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPlacas", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePlanoDeContas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPlanoDeContas", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeProcuracao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCessaoDeCredito", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeProcuracoes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaProcuracoes", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeProduto(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaProduto", 800, control, delay, tempo) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeProdutoCupomFiscal(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaProdutoCupomFiscal", 1024, control, delay, tempo) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaOrdemDeProducao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaOrdemDeProducao", 800, control, delay, tempo) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeRomaneios(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaRomaneios", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeSubOperacoes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaProduto", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeTiposDeVeiculos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaTiposDeVeiculos", 450, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeViasDeTransportes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaViasDeTransportes", 450, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotaFiscalXClassificacao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 1000)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotaFiscalXClassificacao", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotaFiscalSemAbrirClassificacao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 1000)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePedidosXNotas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaPedidosXNotas", 1150, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeEncargosPlanoDeContas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaEncargosPlanoDeContas", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeClientesDireto(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaClienteDireto", 750, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeVencimentos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaVencimentos", 1024, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeArrendantes(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaArrendantes", 1024, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeNotaDeDevolucaoXNota(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divNotaDeDevolucaoXNota", 1024, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeRegistrosIcmsAjustaResumo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divRegistrosIcmsAjustaResumo", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeRegistrosIpiAjustaResumo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divRegistrosIpiAjustaResumo", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeInutilizacaoFiscal(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divInutilizacaoFiscal", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeInutilizacao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divInutilizacao", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePedidoXSaldoXFinanceiro(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divPedidoXSaldo", 1500, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeFazendaCPR(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaFazendaCPR", 1024, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDePesoDeChegada(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divPesoDeChegada", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeDestinoContabil(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divDestinoContabil", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CadastrarOrigemDestinoRoteiro(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divOrigemDestinoRoteiro", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CadastrarFavorecidoPamcard(ByRef page As System.Web.UI.Control, ByVal type As String, ByVal Favorecido As String, Optional ByVal PostBack As Boolean = False)
        Dim strJavaScript As String
        strJavaScript = "var x = (screen.height / 2) - 140; "
        strJavaScript &= "var y = (screen.width / 2) - 275; "
        strJavaScript &= "var newWindow = window.open(""CadastroFavorecidoPamcard.aspx?tipo=" & type & "&ref=" & PostBack.ToString & "&cmp=" & Favorecido & """, """", ""resizable=no, menubar=no, scrollbars=Yes, width=550, height=280, top="" + x + "", left="" + y + """"); setTimeout(function () { newWindow.focus() }, 100);"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), "BuscaCliente", strJavaScript, True)
    End Sub

    Public Sub ConsultaDeContaCliente(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaContaClientes", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeCancelamento(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCancelamento", 500, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeAutorizacaoDeEmbarque(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divAutorizacaoDeEmbarque", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeAutorizacaoDeCarregamento(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divAutorizacaoDeCarregamento", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaLancamentosPedido(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divPedidoLancamentoItem", 1360, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaEncargosPedido(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divPedidoEncargo", 700, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaFixacaoPedido(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divPedidoFixacao", 1360, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaFixacaoProcuracao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divFixacaoProcuracao", 1150, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDarDiferencialDeAliquota(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divDarDiferencialDeAliquota", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaNotaFiscalReferencial(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divNotaFiscalReferencial", 1200, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaRelatorio(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaRelatorio", 900, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub CancelamentoDeCheques(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divCancelamentoDeCheque", 690, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaProdutoNFG(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divProdutoNFG", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub InformarDadosProdutoNFG(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divInformarDadosProdutoNFG", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaProdutoXML(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divProdutoXML", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaProdutoCupomFiscal(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divProdutoCupomFiscal", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaNotaOrigem(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divNFOrigem", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaRateio(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divRateio", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub RecompraDeTitulos(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divRecompraDeTitulos", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaNotaFiscal(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaNotaFiscal", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeComissoesXBaixas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divComissoesXBaixas", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeEmailNFe(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divEmailNFe", 600, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub


    Public Sub ConsultaDeEmailNFePedido(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divEmailNFePedido", 600, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeMDFeXNotas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaMDFeXNotas", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeCTeXNotas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCTeXNotas", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaCadastro(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaCadastro", 540, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaMDFeXEstado(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divMDFeXEstado", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ItemRateioDePesagem(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divRateioDePesagem", 1150, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub InputOfDate(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divInputDate", 500, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub AlterarItemRateioDePesagem(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divAlterarRateioDePesagem", 1150, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaTransferencias(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divTransferencias", 1100, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaMonitorDeNotas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divMonitorDeNotas", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaMonitorCupomFiscal(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divMonitorCupomFiscal", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub Supervisor(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divSupervisor", 350, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub NFObsProduto(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divNFObsProduto", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ControleNumeroDeLote(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divControleNumeroDeLote", 600, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub NFEncargo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divNFEncargo", 1000, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeLote(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaDeLote", 500, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub LaudoManual(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divLaudoManual", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ContabilizarNotas(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divContabilizarNotas", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub EnviarXMLEmissao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divEnviarXMLEmissao", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultarNavios(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultarNavios", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultarNaviosXInvoice(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= "$(document).ready(function () {"
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultarNaviosXInvoice", 800, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        strJavaScript &= "});"
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaDeTitulo(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divConsultaTitulo", 350, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaLiberacao(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divLiberacao", 350, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub ConsultaEmitirCTe(ByRef page As System.Web.UI.Control, ByVal type As String, Optional ByVal control As String = "", Optional ByVal delay As Boolean = False, Optional ByVal tempo As Integer = 500)
        HttpContext.Current.Session("ssTipoRetorno") = type
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() {", "")
        strJavaScript &= " " & ShowDialog("divEmitirCTe", 1366, control) & " "
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

End Class