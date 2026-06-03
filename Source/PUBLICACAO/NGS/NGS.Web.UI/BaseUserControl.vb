Imports System.Xml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BaseUserControl
    Inherits System.Web.UI.UserControl
    Implements IBaseUserControl

#Region "Atributos"
    Private _Banco As AcessaBanco
    Private _MainUserControl As UserControl
    Public Popup As New CarregarPopup
    Public ddl As New CarregarDDL
#End Region

#Region "Propriedades"

    Public ReadOnly Property Banco As AcessaBanco
        Get
            If _Banco Is Nothing Then _Banco = New AcessaBanco
            Return _Banco
        End Get
    End Property

    Public Property MainUserControl() As UserControl
        Get
            Return Session("_MainUserControl")
        End Get
        Set(ByVal value As UserControl)
            Session("_MainUserControl") = value
        End Set
    End Property

    Public Property ActiveUser() As String
        Get
            If Session("ssNomeUsuario") IsNot Nothing Then
                Return CType(Session("ssNomeUsuario"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            Session("ssNomeUsuario") = value
        End Set
    End Property

    Public Property ConexaoDb() As String
        Get
            If Session("Conexao") IsNot Nothing Then
                Return CType(Session("Conexao"), String)
            End If
            Return Nothing
        End Get
        Set(ByVal value As String)
            Session("Conexao") = value
        End Set
    End Property

    Public ReadOnly Property IsConnect() As Boolean
        Get
            Try
                If Banco IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Banco.GetConnectionString()) Then
                    Return True
                End If
                Return False
            Catch ex As Exception
                Return False
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

#End Region

    Public Overridable Sub SetarHID(ByVal guid As String)
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub Limpar()
        Throw New NotImplementedException()
    End Sub

    Protected Overridable Sub Selecionar()
        Throw New NotImplementedException()
    End Sub

    Protected Overridable Sub Selecionar(ByVal args As String)
        Throw New NotImplementedException()
    End Sub

    Protected Overridable Sub Selecionar(ByVal obj As IBaseEntity)
        Throw New NotImplementedException()
    End Sub

    Public Overridable Sub Carregar(ByVal str As String) Implements IBaseUserControl.Carregar
    End Sub

    Public Overridable Sub Carregar(ByVal str As String, ByVal dec As Decimal) Implements IBaseUserControl.Carregar
    End Sub

    Public Overridable Sub Carregar(ByVal obj As IBaseEntity) Implements IBaseUserControl.Carregar
    End Sub

    Public Overridable Sub Carregar(ByVal parameters As Dictionary(Of String, Object)) Implements IBaseUserControl.Carregar
    End Sub

    Protected Sub MsgBox(ByVal page As Page, ByVal message As String, Optional ByVal etitle As eTitulo = eTitulo.Info)
        ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("alert('{0}');", Funcoes.EliminarCaracteresEspeciais(message)), True)
        'ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("msgbox('{0}','{1}','{2}');", Funcoes.EliminarCaracteresEspeciais(message), WebHelpers.GetEnumDescription(etitle), etitle), True)
    End Sub

    Protected Sub MsgBox(ByVal page As Page, ByVal message As String, ByVal url As String, Optional ByVal etitle As eTitulo = eTitulo.Info)
        ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("alert('{0}'); window.location.href = '{1}';", Funcoes.EliminarCaracteresEspeciais(message), ResolveUrl(url)), True)
        'ScriptManager.RegisterClientScriptBlock(page, GetType(Page), Guid.NewGuid().ToString(), String.Format("msgbox('{0}','{1}','{2}'); window.location.href = '{3}';", Funcoes.EliminarCaracteresEspeciais(message), WebHelpers.GetEnumDescription(etitle), etitle, ResolveUrl(url)), True)
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

End Class
