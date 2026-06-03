Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaClientes
    Inherits BaseUserControl

    Private strSQL As String = String.Empty

    Private Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "Situacao_id in (1,2,3,4,5,6,50)")
            Me.Limpar()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub CarregarCNPJ(ByVal pcodigo As String)
        txtCodigo.Text = pcodigo
        Consultar()
    End Sub

    Private Sub Consultar()
        If txtCodigo.Text = "" And txtNome.Text = "" And txtFantasia.Text = "" And txtCidade.Text = "" And txtEstado.Text = "" Then
            strSQL = "SELECT top 500 C.Cliente_Id as Codigo, C.Endereco_Id , C.Nome, C.Complemento, C.Cidade, C.Estado, C.Reduzido, C.Inscricao, C.Situacao as Situacao_id, S.Descricao as DescSituacao " & vbCrLf & _
                     "  FROM Clientes C" & vbCrLf & _
                     " Inner Join Situacoes S" & vbCrLf & _
                     "    on S.Situacao_id = C.Situacao" & vbCrLf & _
                     " Where  Nome  like '%" & txtNome.Text & "%'" & vbCrLf

            If ddlSituacao.SelectedIndex > 0 Then strSQL &= "  And C.Situacao = " & ddlSituacao.SelectedValue


            If Not String.IsNullOrWhiteSpace(hdnTipoCliente.Value) Then
                strSQL &= "AND EXISTS (SELECT 1 " & _
                          "              FROM ClientesXTipos CT " & _
                          "             WHERE C.Cliente_Id = CT.Cliente_Id " & _
                          "               AND CT.Tipo_Id IN (" & hdnTipoCliente.Value & ")) "
            End If
        Else
            Dim pcodigo As String = Funcoes.FormatarCPFeCNPJ(txtCodigo.Text)

            strSQL = "SELECT C.Cliente_Id as Codigo, C.Endereco_Id, C.Nome, C.Complemento, C.Cidade, C.Estado, C.Inscricao, C.Situacao as Situacao_id, S.Descricao as DescSituacao" & vbCrLf & _
                     "  FROM Clientes C" & vbCrLf & _
                     " Inner Join Situacoes S" & vbCrLf & _
                     "    on S.Situacao_id = C.Situacao" & vbCrLf & _
                     " Where  Cliente_Id  like '" & pcodigo & "%'" & vbCrLf

            If ddlSituacao.SelectedIndex > 0 Then strSQL &= "  And C.Situacao = " & ddlSituacao.SelectedValue

            If txtNome.Text <> "" Then
                If chkParteNome.Checked = True Then
                    strSQL &= " And C.Nome like '%" & txtNome.Text & "%'"
                Else
                    strSQL &= " And C.Nome like '" & txtNome.Text & "%'"
                End If
            End If

            If txtFantasia.Text <> "" Then
                strSQL &= " And C.Fantasia like '%" & txtFantasia.Text & "%'"
            End If

            If txtCidade.Text <> "" Then
                strSQL &= " And C.Cidade like '" & txtCidade.Text & "%'"
            End If

            If txtEstado.Text <> "" Then
                strSQL &= " And C.Estado like '" & txtEstado.Text & "%'"
            End If

            If Not String.IsNullOrWhiteSpace(hdnTipoCliente.Value) Then
                strSQL &= "AND EXISTS (SELECT 1 " & _
                          "              FROM ClientesXTipos CT " & _
                          "             WHERE C.Cliente_Id = CT.Cliente_Id " & _
                          "               AND CT.Tipo_Id IN (" & hdnTipoCliente.Value & ")) "
            End If
        End If

        'strSQL &= " And Situacao in (1,50) "
        strSQL &= "ORDER BY C.Nome"

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(strSQL, "Clientes")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 Then
            For Each dra As DataRow In ds.Tables(0).Rows
                dra("Codigo") = Funcoes.FormatarCpfCnpj(dra("Codigo"))
            Next
        End If

        If ds Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Cliente(s) não encontrado.", eTitulo.Info)
            GridClientes.DataSource = New List(Of Object)()
            GridClientes.DataBind()
            txtNome.Focus()
        Else
            GridClientes.DataSource = ds
            GridClientes.DataBind()
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtCodigo.Text = ""
        txtNome.Text = ""
        txtFantasia.Text = ""
        txtCidade.Text = ""
        txtEstado.Text = ""
        hdnTipoCliente.Value = ""
        chkParteNome.Checked = False
        Session.Remove("_MainUserControl")
        GridClientes.DataSource = New List(Of Object)()
        GridClientes.DataBind()
        txtNome.Focus()
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            Dim cliente As String() = args.Split(";")
            Dim objCliente As New [Lib].Negocio.Cliente(cliente(0), Convert.ToInt32(cliente(1)))
            Session(Session("ssTipoRetorno")) = objCliente
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(objCliente)
                Else
                    CType(Me.Page, IBasePage).Carregar(objCliente)
                End If
                Popup.CloseDialog(Me.Page, "divConsultaCliente")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub SetarTituloDIV(ByVal Titulo As String)
        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "$(document).ready(function () { $('#divConsultaCliente').prop('title', '" & Titulo & "'); });", True)
    End Sub

    Public Sub SetarTipoCliente(ByVal TipoCliente As String)
        hdnTipoCliente.Value = TipoCliente
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            Consultar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Me.Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Try
            If TypeOf Me.Page Is Laudo AndAlso Session("ssCampo" & HID.Value) = "Transportador" Then CType(Me.Page, Laudo).LiberaTransportador()
            Popup.CloseDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkSelecionar_Click(sender As Object, e As EventArgs)
        Try
            Dim lnk As LinkButton = CType(sender, LinkButton)
            Dim row As GridViewRow = CType(lnk.NamingContainer, GridViewRow)

            Dim strCodigo As String = row.Cells(5).Text().Replace(".", "").Replace("/", "").Replace("-", "")
            Dim strEndereco As String = row.Cells(6).Text()
            Selecionar(strCodigo & ";" & strEndereco)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class