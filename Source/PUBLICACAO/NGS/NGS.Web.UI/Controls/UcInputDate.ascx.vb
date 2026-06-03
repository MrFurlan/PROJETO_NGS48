Public Class UcInputDate
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Protected Sub btnConfirmar_Click(sender As Object, e As EventArgs) Handles btnConfirmar.Click
        If Validar() Then
            Session("objGradeSeguranca" & HID.Value) = txtAno1.Text & "-" & txtAno2.Text.Trim()
            DirectCast(Me.Page, IBasePage).Carregar(txtAno1.Text & "-" & txtAno2.Text.Trim())
            Limpar()
            Popup.CloseDialog(Me.Page, "divInputDate")
        End If
    End Sub

    Public Overrides Sub Limpar()
        txtAno1.Text = String.Empty
        txtAno2.Text = String.Empty
        txtAno1.Enabled = True
    End Sub

    Protected Sub btnFechar_Click(sender As Object, e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divInputDate")
    End Sub

    Public Sub SetarAnoInicial(ByVal AnoInicial As Integer)
        txtAno1.Text = AnoInicial
        txtAno1.Enabled = False
    End Sub

    Public Sub SetarAnoFinal(ByVal AnoFinal As Integer)
        txtAno2.Text = AnoFinal
        txtAno2.Enabled = False
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(txtAno1.Text) OrElse Not IsNumeric(txtAno1.Text) OrElse Len(txtAno1.Text) <> 4 Then
            MsgBox(Me.Page, "Informe o Ano Inicial.", [Lib].Negocio.eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtAno2.Text) OrElse Not IsNumeric(txtAno2.Text) OrElse Len(txtAno2.Text) <> 4 Then
            MsgBox(Me.Page, "Informe o Ano Final.", [Lib].Negocio.eTitulo.Info)
            Return False
        End If
        Return True
    End Function
End Class