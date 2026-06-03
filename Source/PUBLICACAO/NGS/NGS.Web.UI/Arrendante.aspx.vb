Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Arrendante
    Inherits BasePage

#Region "Cliente"
    Private objCliente As [Lib].Negocio.Cliente

    Private Sub SessaoSalvaCliente()
        Session("objSCliente") = objCliente
    End Sub

    Private Sub SessaoRecuperaCliente()
        objCliente = CType(Session("objSCliente"), [Lib].Negocio.Cliente)
    End Sub
#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            'If Funcoes.VerificaPermissao("Arrendante", "ACESSAR") Then
            HID.Value = Guid.NewGuid().ToString()
            ucConsultaClientes.SetarHID(HID.Value)
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
            'End If
        End If
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objClienteCLIXA" & HID.Value) Is Nothing Then
            Dim objArrendante = CType(Session("objClienteCLIXA" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemArrendante As ListItem = Funcoes.FormatarListItemCliente(objArrendante)
            txtArrendante.Text = itemArrendante.Text
            txtCodigoArrendante.Value = itemArrendante.Value & "-" & CType(Session("objClienteCLIXA" & HID.Value), [Lib].Negocio.Cliente).Nome
            Session.Remove("objClienteCLIXA" & HID.Value)
        End If
    End Sub

    Protected Sub gridVencimentos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub btnArrendante_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnArrendante.Click
        Session("ssCampo") = "LivreClasse"
        Popup.ConsultaDeClientes(Me.Page, "objClienteCLIXA" & HID.Value)
    End Sub

    Protected Sub txtDataContrato_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
    End Sub

    Protected Sub txtDataVencimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        gridMatricula.DataSource = Nothing
        gridMatricula.DataBind()
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
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

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
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
                Dim cxa As New [Lib].Negocio.ClienteXArrendante(objCliente)
                cxa.IUD = "I"
                cxa.CodigoArrendante = cli(0)
                cxa.EndArrendante = cli(1)
                cxa.NomeArrendante = cli(2)
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
        Dim strJS As String = "window.opener.document.forms[0].submit();" & _
                      "window.close();"
        ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), strJS, True)
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Arrendante")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class