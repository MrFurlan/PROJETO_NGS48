Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucCancelamentoDeCheque
    Inherits BaseUserControl

#Region "Property"
    Public Property MyParameters() As Dictionary(Of String, Object)
        Get
            Return CType(Session("MyParameters" & HID.Value), Dictionary(Of String, Object))
        End Get
        Set(ByVal value As Dictionary(Of String, Object))
            Session("MyParameters" & HID.Value) = value
        End Set
    End Property

    'Public Property ObjBxCxC() As BancosXContasXCancelamentoCheque
    '    Get
    '        Return CType(Session("ssObjCancelamentoCheque" & HID.Value), BancosXContasXCancelamentoCheque)
    '    End Get
    '    Set(ByVal value As BancosXContasXCancelamentoCheque)
    '        Session("ssObjCancelamentoCheque" & HID.Value) = value
    '    End Set
    'End Property

    Public Property ObjBxC As [Lib].Negocio.BancosXContas
        Get
            Return CType(Session("ssObjBXC" & HID.Value), [Lib].Negocio.BancosXContas)
        End Get
        Set(ByVal value As [Lib].Negocio.BancosXContas)
            Session("ssObjBXC" & HID.Value) = value
        End Set
    End Property

    Public Property ListBxCxC() As ListBancosXContasXCancelamentoCheque
        Get
            Return CType(Session("ssListCancelamentoCheque" & HID.Value), ListBancosXContasXCancelamentoCheque)
        End Get
        Set(ByVal value As ListBancosXContasXCancelamentoCheque)
            Session("ssListCancelamentoCheque" & HID.Value) = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub pageLoad()
        If MyParameters("tipo") = "I" Then
            txtEmpresa.Text = MyParameters("empresaNome")
            txtCondigoEmp.Value = MyParameters("empresaCodigo")
            txtBanco.Text = MyParameters("banco").ToString().PadLeft(4, "0")
            txtAgencia.Text = MyParameters("agencia")
            txtDigitoAgencia.Text = MyParameters("digitoAgencia")
            txtConta.Text = MyParameters("conta")
            txtDigitoConta.Text = MyParameters("digitoConta")
            txtChequeInicial.Text = MyParameters("ChequeAtual")
            txtChequeFinal.Text = MyParameters("ChequeFinal")
            txtBanco.ToolTip = MyParameters("nomeBanco").ToString().PadLeft(4, "0")


            CarregaGrid()
        End If
    End Sub

    Private Sub CarregaGrid()
        If ObjBxC Is Nothing Then
            ObjBxC = New [Lib].Negocio.BancosXContas()
            ObjBxC.CodigoEmpresa = txtCondigoEmp.Value.Split("-")(0)
            ObjBxC.EndEmpresa = txtCondigoEmp.Value.Split("-")(1)
            ObjBxC.CodigoBanco = txtBanco.Text
            ObjBxC.Agencia = txtAgencia.Text
            ObjBxC.DigitoAgencia = txtDigitoAgencia.Text
            ObjBxC.Conta = txtConta.Text
            ObjBxC.DigitoConta = txtDigitoConta.Text
        End If

        ListBxCxC = New ListBancosXContasXCancelamentoCheque(ObjBxC)

        gridCancelados.DataSource = ListBxCxC
        gridCancelados.DataBind()
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtChequeInicial.Text) OrElse String.IsNullOrWhiteSpace(txtChequeFinal.Text) OrElse _
        CInt(txtChequeInicial.Text) < CInt(MyParameters("ChequeAtual")) OrElse CInt(txtChequeFinal.Text) > CInt(MyParameters("ChequeFinal")) Then
            MsgBox(Me.Page, "Informe um intervalo de cheques entre: " & MyParameters("ChequeAtual") & " e " & MyParameters("ChequeFinal"))
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtObservacao.Text) Then
            MsgBox(Me.Page, "Informe o motivo da Inutilizacao no campo Observação.")
            Return False
        ElseIf txtObservacao.Text.Length < 15 Then
            MsgBox(Me.Page, "Campo Observação de conter no mínimo 15 caracteres.")
            Return False
        End If
        Return True
    End Function

    Private Sub SendMail()
        'SEND E-MAIL
        Try
            Dim body As String = ""
            Dim cheques As String = ""
            Dim subject As String = ""
            Dim lstTo As New List(Of String)

            Dim Sql = "SELECT cxu.*, u.Email " & vbCrLf & _
                  "FROM ConfiguracaoXUsuario cxu " & vbCrLf & _
                  "INNER JOIN Usuarios u ON (u.Usuario_Id = cxu.Usuario_Id) " & vbCrLf & _
                  "WHERE cxu.Etapa_Id = " & eEtapa.Cheque

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Vinculados")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    lstTo.Add(row("Email"))
                Next
            Else
                Exit Sub
            End If

            If txtChequeInicial.Text.Trim = txtChequeFinal.Text.Trim Then
                cheques = txtChequeInicial.Text
            Else
                cheques = txtChequeInicial.Text & " ao " & txtChequeFinal.Text
            End If

            If MyParameters("tipo") = "I" Then
                subject = "Cheque(s) Inutilizado(s)"
                body = " O(s) Cheque(s) de Número(s) " & cheques & " foram INUTILIZADO(S) pelo usuário: " & UsuarioServidor.NomeUsuario & " dia: " & DateTime.Now.ToString("dd/MM/yyyy") & " as " & DateTime.Now.ToString("HH:mm:ss") & "<br />" & vbCrLf & _
                       " Pelo Motivo abaixo descrito:<br /><br />" & vbCrLf & _
                       txtObservacao.Text.Trim
            Else
                subject = "Cheque(s) Cancelado(s)"
                body = " O(s) Cheque(s) de Número(s) " & cheques & " foram CANCELADO(S) pelo usuário: " & UsuarioServidor.NomeUsuario & " dia: " & DateTime.Now.ToString("dd/MM/yyyy") & " as " & DateTime.Now.ToString("HH:mm:ss") & "<br />" & vbCrLf & _
                      " Pelo Motivo abaixo descrito:<br /><br />" & vbCrLf & _
                      txtObservacao.Text.Trim
            End If

            'Dim bodyHTML = "Arquivos excluídos em teste, porém, o mesmo conteúdo esta encaminhado para ....."
            Dim errorMsg = String.Empty
            'Dim subject = String.Format("Cancelamento de cheques")
            Dim smtp = Funcoes.GetSmtpSettings()
            Dim fromMail = Funcoes.GetFromMail()

            'Dim toMail As String = "ricardo@ngssolucoes.com.br"
            If Funcoes.SendMail(fromMail, "NGS Soluções", lstTo, subject, body, smtp, errorMsg) Then
                MsgBox(Me.Page, errorMsg)
            End If
        Catch ex As Exception
            Throw New Exception()
        End Try
    End Sub

    Public Overrides Sub Limpar()
        txtEmpresa.Text = ""
        txtCondigoEmp.Value = ""
        txtBanco.Text = ""
        txtAgencia.Text = ""
        txtDigitoAgencia.Text = ""
        txtConta.Text = ""
        txtDigitoConta.Text = ""
        txtChequeInicial.Text = ""
        txtChequeFinal.Text = ""
        txtObservacao.Text = ""
    End Sub

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Limpar()
        End If
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If ValidarCampos() Then
                Dim sqls As New ArrayList
                Dim ultimoNumeroInsert As Integer
                Dim objBXCXC As New [Lib].Negocio.BancosXContasXCancelamentoCheque()

                For i As Integer = CInt(txtChequeInicial.Text) To CInt(txtChequeFinal.Text)
                    Dim numcheque As Integer = i

                    If ListBxCxC.Count = 0 OrElse ListBxCxC.Where(Function(s) s.NumCheque = numcheque).FirstOrDefault().NumCheque <> numcheque Then
                        objBXCXC.IUD = "I"
                        objBXCXC.DataCancelamento = DateTime.Now
                        objBXCXC.Pedido = 0
                        objBXCXC.Titulo = 0
                        objBXCXC.Observacao = txtObservacao.Text.Trim()
                        objBXCXC.UsuarioCancelamento = UsuarioServidor.NomeUsuario
                        objBXCXC.TipoCancelamento = MyParameters("tipo")
                        objBXCXC.NumCheque = numcheque

                        ultimoNumeroInsert = numcheque

                        objBXCXC.SalvarSql(sqls)
                    End If
                Next

                If ultimoNumeroInsert > 0 Then
                    Dim objBancoXConta As New [Lib].Negocio.BancosXContas(txtCondigoEmp.Value.Split("-")(0), txtCondigoEmp.Value.Split("-")(1), txtBanco.Text, txtAgencia.Text, txtDigitoAgencia.Text, txtConta.Text, txtDigitoConta.Text)
                    objBancoXConta.IUD = "U"
                    objBancoXConta.NumChequeAtual = ultimoNumeroInsert
                    objBancoXConta.SalvarSql(sqls)
                End If

                If Banco.GravaBanco(sqls) Then
                    MsgBox(Me.Page, "Dados Incluídos com Sucesso.", eTitulo.Sucess)
                    Popup.CloseDialog(Me.Page, "divCancelamentoDeCheque")

                    SendMail()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkSair_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkSair.Click
        Popup.CloseDialog(Me.Page, "divCancelamentoDeCheque")
    End Sub

    Protected Sub gridCancelados_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles gridCancelados.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(4).Text = "I" Then
                e.Row.Cells(12).Text = "Inutilizado"
            Else
                e.Row.Cells(12).Text = "Cancelado"
                'e.Row.Cells(12).ForeColor = Drawing.Color.Red
            End If
        End If
    End Sub

#End Region

End Class