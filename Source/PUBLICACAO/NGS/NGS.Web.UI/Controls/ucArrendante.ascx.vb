Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucArrendante
    Inherits BaseUserControl

#Region "Cliente"
    Private objCliente As Cliente

    Private Sub SessaoSalvaCliente()
        Session("objSCliente") = objCliente
    End Sub

    Private Sub SessaoRecuperaCliente()
        objCliente = CType(Session("objSCliente"), Cliente)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objArrendanteCliXArr" + HID.Value) Is Nothing Then
            Dim objArrendanteCliXArr = CType(Session("objArrendanteCliXArr" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(objArrendanteCliXArr)
            txtArrendante.Text = itemCliente.Text
            txtCodigoArrendante.Value = itemCliente.Value
            'consultar()
            Session.Remove("objArrendanteCliXArr")
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()
        txtCodigoArrendante.Value = ""
        txtArrendante.Text = ""
        txtDataContrato.Text = ""
        txtDataVencimento.Text = ""
        Session.Remove("objArrendanteCliXArr")
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        SessaoRecuperaCliente()
        Dim cli As String() = txtCodigoArrendante.Value.Split("-")
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArrendamento")
            Dim txtObs As TextBox = gridMatricula.Rows(i).FindControl("txtObservacao")
            If CDec(txt.Text) > 0 Then
                Dim cxa As New ClienteXArrendante(objCliente)
                Dim objArrendante = New Cliente(cli(0), cli(1))
                cxa.IUD = "I"
                cxa.CodigoArrendante = cli(0)
                cxa.EndArrendante = cli(1)
                cxa.NomeArrendante = objArrendante.Nome
                cxa.DataContrato = CDate(txtDataContrato.Text)
                cxa.DataVencimento = CDate(txtDataVencimento.Text)
                cxa.Matricula = gridMatricula.Rows(i).Cells(0).Text
                cxa.Area = CDec(txt.Text)
                cxa.Observacao = txtObs.Text
                If objCliente.IUD <> "I" Then
                    If cxa.Salvar() Then
                        objCliente.Arrendantes.Add(cxa)
                    Else
                        MsgBox(Me.Page, "Erro Durante o Processo de Inclusao")
                    End If
                End If
            End If
        Next
        SessaoSalvaCliente()
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objArrendanteCli")) Then
            If TypeOf Me.Page Is Clientes Then
                CType(Me.Page, Clientes).Carregar_Arrendantes()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divConsultaArrendantes")
    End Sub

    Protected Sub btnArrendante_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Session("ssCampo") = "LivreClasse"
        Dim ucConsultaClientes As ucConsultaClientes = DirectCast(Me.NamingContainer.FindControl("ucConsultaClientes"), ucConsultaClientes)
        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            Me.MainUserControl = Me
            Popup.ConsultaDeClientes(Me.Page, "objArrendanteCliXArr" & HID.Value, "txtNome")
        End If
    End Sub

    Protected Sub txtDataContrato_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
    End Sub

    Protected Sub txtDataVencimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles txtDataVencimento.TextChanged
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
        consultar()
    End Sub

    Private Sub Confirmar()
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArrendamento")
            If CDec(gridMatricula.Rows(i).Cells(3).Text) < CDec(txt.Text) Then
                MsgBox(Me.Page, "Area Arrendada Nao esta Disponivel para a Matricula: " & gridMatricula.Rows(i).Cells(0).Text)
                Exit Sub
            End If
        Next
        SessaoRecuperaCliente()
        Dim cli As String() = txtCodigoArrendante.Value.Split("-")
        For i As Integer = 0 To gridMatricula.Rows.Count - 1
            Dim txt As TextBox = gridMatricula.Rows(i).FindControl("txtArrendamento")
            Dim txtObs As TextBox = gridMatricula.Rows(i).FindControl("txtObservacao")
            If CDec(txt.Text) > 0 Then
                Dim cxa As New ClienteXArrendante(objCliente)
                Dim objArrendante = New Cliente(cli(0), cli(1))
                cxa.IUD = "I"
                cxa.CodigoArrendante = cli(0)
                cxa.EndArrendante = cli(1)
                cxa.NomeArrendante = objArrendante.Nome
                cxa.DataContrato = CDate(txtDataContrato.Text)
                cxa.DataVencimento = CDate(txtDataVencimento.Text)
                cxa.Matricula = gridMatricula.Rows(i).Cells(0).Text
                cxa.Area = CDec(txt.Text)
                cxa.Observacao = txtObs.Text
                If objCliente.IUD <> "I" Then
                    If cxa.Salvar() Then
                        objCliente.Arrendantes.Add(cxa)
                    Else
                        MsgBox(Me.Page, "Erro Durante o Processo de Inclusao")
                    End If
                End If
            End If
        Next
        SessaoSalvaCliente()
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objArrendanteCli")) Then
            If TypeOf Me.Page Is Clientes Then
                CType(Me.Page, Clientes).Carregar_Arrendantes()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divConsultaArrendantes")
    End Sub

    Protected Sub lnkFechar_Click(sender As Object, e As EventArgs) Handles lnkFechar.Click
        Session("Arrendantes") = True
        If (Session("ssTipoRetorno") IsNot Nothing AndAlso Session("ssTipoRetorno").ToString.Contains("objArrendanteCli")) Then
            If TypeOf Me.Page Is Clientes Then
                CType(Me.Page, Clientes).Carregar_Arrendantes()
            End If
        End If
        Popup.CloseDialog(Me.Page, "divConsultaArrendantes")
    End Sub

    Private Sub consultar()
        If Not IsDate(txtDataVencimento.Text) Or Not IsDate(txtDataContrato.Text) Then
            MsgBox(Me.Page, "Informe as datas de contrato e vencimento")
            Exit Sub
        End If

        If CDate(txtDataVencimento.Text) < CDate(txtDataContrato.Text) Then
            MsgBox(Me.Page, "Data de Vencimento Menor que a data de contrato")
            Exit Sub
        End If

        Dim cli As String() = txtCodigoArrendante.Value.Split("-")
        Dim sql As String
        sql = "SELECT CM.Cliente_Id, " & vbCrLf & _
              "       CM.EndCliente_Id," & vbCrLf & _
              "       CM.Matricula_Id, " & vbCrLf & _
              "       CM.Area," & vbCrLf & _
              "       isnull((select sum(isnull(area,0)) " & vbCrLf & _
              "		  from clientexarrendante CA" & vbCrLf & _
              "		 WHERE CA.Arrendante_Id    = CM.Cliente_Id " & vbCrLf & _
              "		   AND CA.EndArrendante_Id = CM.EndCliente_Id" & vbCrLf & _
              "		   AND CA.matricula_id     = CM.Matricula_Id" & vbCrLf & _
              "		   and ('" & CDate(txtDataContrato.Text).ToString("yyyy-MM-dd") & "' between CA.dataContrato_id and CA.DataVencimento" & vbCrLf & _
              "              or " & vbCrLf & _
              "                '" & CDate(txtDataVencimento.Text).ToString("yyyy-MM-dd") & "' between CA.dataContrato_id and CA.DataVencimento" & vbCrLf & _
              "              or" & vbCrLf & _
              "                CA.dataContrato_id between '" & CDate(txtDataContrato.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataVencimento.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "              or " & vbCrLf & _
              "                CA.DataVencimento  between '" & CDate(txtDataContrato.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataVencimento.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "                )" & vbCrLf & _
              "        ),0) AreaArrendada" & vbCrLf & _
              "  Into #Matriculas " & vbCrLf & _
              "  FROM Clientes AS c" & vbCrLf & _
              " INNER JOIN ClienteXMatricula AS CM" & vbCrLf & _
              "    ON c.Cliente_Id  = CM.Cliente_Id" & vbCrLf & _
              "   AND c.Endereco_Id = CM.EndCliente_Id" & vbCrLf & _
              " WHERE c.Cliente_Id  = '" & cli(0) & "'" & vbCrLf & _
              "   AND c.Endereco_Id =  " & cli(1) & vbCrLf & _
              " Select *, case when Area - AreaArrendada < 0 then 0 else Area - AreaArrendada end as Saldo from #Matriculas "

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Matriculas")

        gridMatricula.DataSource = ds.Tables(0)
        gridMatricula.DataBind()
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        consultar()
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            If Not IsDate(txtDataVencimento.Text) Or Not IsDate(txtDataContrato.Text) Then
                MsgBox(Me.Page, "Informe as datas de contrato e vencimento")
            ElseIf CDate(txtDataVencimento.Text) < CDate(txtDataContrato.Text) Then
                MsgBox(Me.Page, "Data de Vencimento Menor que a data de contrato")
            Else
                Confirmar()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class