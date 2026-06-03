Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucMDFeXEstado
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            BindGridView()
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal NewGuid As String)
        HID.Value = NewGuid.ToString
    End Sub

    Public Sub BindGridView()
        Dim lstNotaFiscalXPercurso As [Lib].Negocio.ListNotaFiscalXPercurso = CType(Session("Estados" & HID.Value), [Lib].Negocio.ListNotaFiscalXPercurso)
        Dim lst As New [Lib].Negocio.Estados(True)

        If lstNotaFiscalXPercurso IsNot Nothing AndAlso lstNotaFiscalXPercurso.Count > 0 Then
            For Each e As [Lib].Negocio.Estado In lst
                For Each sl As [Lib].Negocio.NotaFiscalXPercurso In lstNotaFiscalXPercurso
                    If sl.Estado_Id = e.Codigo Then
                        e.Sequencia = sl.Ordem
                        e.Selecionado = True
                    End If
                Next
            Next
        End If

        grd.DataSource = lst
        grd.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            Dim lstSelecionados As List(Of String) = grd.GetSelectedValues("chkCodigo")
            Dim lstEstados As New [Lib].Negocio.Estados

            For Each sl In lstSelecionados
                If Not String.IsNullOrWhiteSpace(sl) Then
                    Dim e As New [Lib].Negocio.Estado(sl)
                    For Each row As GridViewRow In grd.Rows
                        If row.Cells(2).Text.Trim() = e.Codigo Then
                            Dim txtSequencia As TextBox = CType(row.FindControl("txtSequencia"), TextBox)
                            e.Sequencia = txtSequencia.Text
                        End If
                    Next
                    lstEstados.Add(e)
                End If
            Next

            Session(Session("ssTipoRetorno")) = lstEstados
            If Session("ssTipoRetorno") IsNot Nothing Then
                If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                    Dim ucName = MainUserControl.ClientID.Split("_")
                    Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                    CType(uc, IBaseUserControl).Carregar(lstEstados)
                Else
                    CType(Me.Page, IBasePage).Carregar(lstEstados)
                End If
                Popup.CloseDialog(Me.Page, "divMDFeXEstado")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divMDFeXEstado")
    End Sub

    Protected Sub btnSelecionar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnSelecionar.Click
        Selecionar()
    End Sub

End Class