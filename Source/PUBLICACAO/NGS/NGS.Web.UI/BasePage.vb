Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BasePage
    Inherits System.Web.UI.Page
    Implements IBasePage

#Region "Atributos"
    Public Popup As New CarregarPopup
    Public ddl As New CarregarDDL
    Public _Banco As AcessaBanco
#End Region

#Region "Propriedades"
    Public ReadOnly Property Banco() As AcessaBanco
        Get
            If _Banco Is Nothing Then
                _Banco = New AcessaBanco()
            End If

            Return _Banco
        End Get
    End Property

    Public ReadOnly Property IsConnect() As Boolean
        Get
            Try
                If Banco IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Banco.GetConnectionString()) Then
                    Return True
                Else
                    Response.Redirect("~/Logout.aspx", False)
                End If
                Return False
            Catch ex As Exception
                Throw New Exception()
            End Try
        End Get
    End Property

    Public ReadOnly Property FinanceiroNovo() As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
        End Get
    End Property

    Public ReadOnly Property FinanceiroVirtual() As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("FinanceiroVirtual"))
        End Get
    End Property

    Public ReadOnly Property EmbarqueNovo() As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("embarqueNovo"))
        End Get
    End Property

    Public ReadOnly Property HomologAlfasig() As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("homologAlfasig"))
        End Get
    End Property

    Public ReadOnly Property LancarSaldoInicial() As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("lancarSaldoInicial"))
        End Get
    End Property

#End Region

    Protected Overrides Sub OnPreInit(ByVal e As System.EventArgs)
        MyBase.OnPreInit(e)
        Me.Page.Theme = ConfigurationManager.AppSettings("themeName").ToString
    End Sub

    Public Overridable Sub Carregar(ByVal str As String) Implements IBasePage.Carregar
    End Sub

    Public Overridable Sub Carregar(ByVal str As String, ByVal dec As Decimal) Implements IBasePage.Carregar
    End Sub

    Public Overridable Sub Carregar(ByVal obj As IBaseEntity) Implements IBasePage.Carregar
    End Sub

    Public Sub setMenu(ByVal modulo As eModulo)
        Dim master As Principal = CType(Me.Master, Principal)
        If master IsNot Nothing Then
            master.setMenu(modulo)
        End If
    End Sub

    Public Sub ShowCalendar(ByVal page As System.Web.UI.Page, ByVal txt As System.Web.UI.WebControls.TextBox, Optional ByVal delay As Boolean = True, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$(document).ready(function () { mostrarCalendario('#" & txt.ClientID & "', true); });"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Public Sub HideCalendar(ByVal page As System.Web.UI.Page, ByVal txt As System.Web.UI.WebControls.TextBox, Optional ByVal delay As Boolean = True, Optional ByVal tempo As Integer = 100)
        Dim strJavaScript As String = String.Empty
        strJavaScript &= IIf(delay, "window.setTimeout(function() { ", "")
        strJavaScript &= "$(document).ready(function () { mostrarCalendario('#" & txt.ClientID & "', false); });"
        strJavaScript &= IIf(delay, "}, " & tempo & ");", "")
        ScriptManager.RegisterClientScriptBlock(page, page.GetType(), Guid.NewGuid().ToString, strJavaScript, True)
    End Sub

    Protected Sub MsgBox(ByVal page As Page, ByVal message As String, Optional ByVal etitle As eTitulo = eTitulo.Info, Optional eliminarCaracteresEspeciais As Boolean = True)
        'ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("msgbox('{0}','{1}','{2}');", IIf(eliminarCaracteresEspeciais, Funcoes.EliminarCaracteresEspeciais(message), message), WebHelpers.GetEnumDescription(etitle), etitle), True)
        ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("alert('{0}');", IIf(eliminarCaracteresEspeciais, Funcoes.EliminarCaracteresEspeciais(message), message)), True)
    End Sub

    Protected Sub MsgBox(ByVal page As Page, ByVal message As String, ByVal url As String, Optional ByVal etitle As eTitulo = eTitulo.Info)
        ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("alert('{0}'); window.location.href = '{1}';", Funcoes.EliminarCaracteresEspeciais(message), ResolveUrl(url)), True)
        'ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("msgbox('{0}','{1}',{2}); window.location.href = '{3}';", Funcoes.EliminarCaracteresEspeciais(message), WebHelpers.GetEnumDescription(etitle), etitle, ResolveUrl(url)), True)
    End Sub

    Public Sub SetarData(ByVal txt As TextBox)
        txt.Text = String.Format("01/{0}/{1}", Now.Month.ToString.PadLeft(2, "0"), Now.Year)
    End Sub

    Public Sub SetarData(ByVal txt1 As TextBox, ByVal txt2 As TextBox, Optional ByVal setar01 As Boolean = True)
        txt1.Text = String.Format("{0}/{1}/{2}", IIf(setar01, "01", Now.Day.ToString.PadLeft(2, "0")), Now.Month.ToString.PadLeft(2, "0"), Now.Year)
        txt2.Text = Format(Today, "dd/MM/yyyy")
    End Sub

#Region "Setar Valores Nulos ou vazios in control"


    ''' <summary>
    ''' Remove valores de GridViews passado como parametro.
    ''' </summary>
    ''' <param name="GridViews">Array de GridView</param>
    ''' <remarks>Ex°: SetClearValue({GridTeste1, GridTeste2, ...})</remarks>
    Public Sub SetClearValue(ByVal GridViews As GridView())
        For Each grid As GridView In GridViews
            grid.DataBind()
        Next
    End Sub

    ''' <summary>
    ''' Remove valores de todos os controles TextBox passado como parametro.
    ''' </summary>
    ''' <param name="TextBoxs">Array de TextBox</param>
    ''' <remarks>Ex°: SetStringEmptyinControls({txtNome, txtEndereco, ...})</remarks>
    Public Sub SetStringEmptyinControls(ByVal TextBoxs As TextBox())
        For Each txt As TextBox In TextBoxs
            txt.Text = String.Empty
        Next
    End Sub

    ''' <summary>
    ''' Habilita ou desabilita edições ou ações dos controles.
    ''' </summary>
    ''' <param name="enabled">"True" spara Habilitar e "False" para Desabilitar</param>
    ''' <param name="controls">Array de Controles Asp.net</param>
    ''' <remarks></remarks>
    Public Sub SetEnabledinControls(ByVal enabled As Boolean, ByVal controls As Object())
        For Each control As WebControl In controls
            control.Enabled = enabled
        Next
    End Sub

    ''' <summary>
    ''' Visibiliza ou Esconde os controles.
    ''' </summary>
    ''' <param name="visible">"True" spara Habilitar e "False" para Desabilitar</param>
    ''' <param name="controls">Array de Controles Asp.net</param>
    ''' <remarks></remarks>
    Public Sub SetVisibledinControls(ByVal visible As Boolean, ByVal controls As Object())
        For Each control As WebControl In controls
            control.Visible = visible
        Next
    End Sub

    ''' <summary>
    '''  Remove valores de todos os controles HiddenField passado como parametro.
    ''' </summary>
    ''' <param name="HiddenFields">Array de HiddenField</param>
    ''' <remarks>Ex°: SetStringEmptyinControls({hdnCodigoCliente, hdnEndereco, ...})</remarks>
    Public Sub SetStringEmptyinControls(ByVal HiddenFields As HiddenField())
        For Each hdn As HiddenField In HiddenFields
            hdn.Value = String.Empty
        Next
    End Sub

    ''' <summary>
    '''  Remove valores de todos os controles DropDownList passado como parametro.
    ''' </summary>
    ''' <param name="DropDownLists">Array de DropDownList</param>
    ''' <remarks>Ex°: SetStringEmptyinControls({ddlEmpresa, ddlCliente, ...})</remarks>
    Public Sub SetStringEmptyinControls(ByVal DropDownLists As DropDownList())
        For Each ddl As DropDownList In DropDownLists
            ddl.SelectedValue = String.Empty
        Next
    End Sub

    ''' <summary>
    ''' Desmarca os CheckBox, de todos os controles CheckBox passado como parametro.
    ''' </summary>
    ''' <param name="CheckBoxs">Array de CheckBox</param>
    ''' <remarks>Ex°: SetStringEmptyinControls({chkTipo, chkParamentro, ...})</remarks>
    Public Sub SetStringEmptyinControls(ByVal CheckBoxs As CheckBox())
        For Each chk As CheckBox In CheckBoxs
            chk.Checked = False
        Next
    End Sub

    ''' <summary>
    ''' Desmarca os RadioButton, de todos os controles RadioButton passado como parametro.
    ''' </summary>
    ''' <param name="RadioButtons">Array de RadioButton</param>
    ''' <remarks>Ex°: SetStringEmptyinControls({rbTipo, rbParametro, ...})</remarks>
    Public Sub SetStringEmptyinControls(ByVal RadioButtons As RadioButton())
        For Each rb As CheckBox In RadioButtons
            rb.Checked = False
        Next
    End Sub

#End Region

End Class