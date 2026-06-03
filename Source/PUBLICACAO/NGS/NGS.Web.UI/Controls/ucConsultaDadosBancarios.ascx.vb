Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaDadosBancarios
    Inherits BaseUserControl

    Dim Cliente As [Lib].Negocio.Cliente
    Dim Contas As [Lib].Negocio.ListClienteXContaBancaria

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Sub SalvarSessao()
        Session("Contas" & HID.Value) = Contas
    End Sub

    Public Sub RecuperarSessao()
        Contas = Session("Contas" & HID.Value)
        Cliente = Session("ClienteConta" & HID.Value)
    End Sub

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Protected Sub ddlEstado_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlCidade, CarregarDDL.Tabela.Cidades, "Estado_Id= '" & ddlEstado.SelectedValue & "'")
    End Sub

    Protected Sub Button2_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If TxtCodAgencia.Text.Trim.Equals(String.Empty) Or TxtConta.Text.Trim.Equals(String.Empty) _
         Or TxtPraca.Text.Trim.Equals(String.Empty) Or ddlCidade.SelectedValue.ToString.Trim.Equals(String.Empty) Or ddlEstado.SelectedValue.ToString.Trim.Equals(String.Empty) Then
            MsgBox(Me.Page, "Todos os campos são obrigatórios!", eTitulo.Info)
            Exit Sub
        End If

        RecuperarSessao()
        Dim sqls As New ArrayList
        Dim ds As New DataSet

        Dim sql = "Select 1 " & vbCrLf & _
                  "  from Agencias " & vbCrLf & _
                  " where Banco_id = " & ddlBanco.SelectedValue & vbCrLf & _
                  "   and Agencia_id = " & TxtCodAgencia.Text & vbCrLf & _
                  "   And Digito_Id = '" & TxtAgenciaDigito.Text & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "DB")

        If ds.Tables("DB").Rows.Count = 0 Then
            sql = "Insert into Agencias( Banco_Id, Agencia_Id, Digito_Id, Praca, Cidade, Estado_Id)"
            sql &= " Values(" & ddlBanco.SelectedValue & "," & TxtCodAgencia.Text & ",'" & TxtAgenciaDigito.Text & "','" & TxtPraca.Text & "','" & ddlCidade.SelectedItem.Text & "','" & ddlEstado.SelectedValue.ToString & "')"
            sqls.Add(sql)
        Else
            sql = " update Agencias set "
            sql &= "  Praca     ='" & TxtPraca.Text & "'"
            sql &= " ,Cidade    ='" & ddlCidade.SelectedItem.Text & "'"
            sql &= " ,Estado_Id ='" & ddlEstado.SelectedValue.ToString & "'"
            sql &= " where Banco_Id = " & ddlBanco.SelectedValue
            sql &= "   and Agencia_Id = " & TxtCodAgencia.Text
            sql &= "   and Digito_Id  ='" & TxtAgenciaDigito.Text & "'"
            sqls.Add(sql)
        End If


        sql = "Select 1 " & vbCrLf & _
              "  from ClientesXContasBancarias " & vbCrLf & _
              " where Cliente_Id = '" & Cliente.Codigo & "'" & vbCrLf & _
              "   AND Endereco_Id =" & Cliente.CodigoEndereco & vbCrLf & _
              "   AND Banco_id = " & ddlBanco.SelectedValue & vbCrLf & _
              "   and Agencia_id = " & TxtCodAgencia.Text & vbCrLf & _
              "   And DigitoAgencia_Id = '" & TxtAgenciaDigito.Text & "'" & vbCrLf & _
              "   and ContaCorrente_Id = '" & TxtConta.Text & "'" & vbCrLf & _
              "   and DigitoConta_Id = '" & TxtContaDigito.Text & "'" & vbCrLf
        ds = Banco.ConsultaDataSet(sql, "DB")

        If ds.Tables("DB").Rows.Count = 0 Then
            sql = "Insert into ClientesXContasBancarias(Cliente_Id, Endereco_Id, Banco_Id, Agencia_Id, DigitoAgencia_Id, ContaCorrente_Id, DigitoConta_Id, TipoConta)"
            sql &= "values ('" & Cliente.Codigo & "'," & Cliente.CodigoEndereco & "," & ddlBanco.SelectedValue & "," & TxtCodAgencia.Text & ",'" & TxtAgenciaDigito.Text & "','" & TxtConta.Text.Trim & "','" & TxtContaDigito.Text & "','" & ddlTipoConta.SelectedValue & "')"
            sqls.Add(sql)
        End If

        If Banco.GravaBanco(sqls) Then
            RecuperarSessao()
            Dim NovoReg As New [Lib].Negocio.ClienteXContaBancaria(CType(Session("ClienteConta"), [Lib].Negocio.Cliente))
            NovoReg.Cliente = Cliente
            NovoReg.CodigoBanco = ddlBanco.SelectedValue
            NovoReg.CodigoAgencia = TxtCodAgencia.Text
            NovoReg.DigitoAgencia = TxtAgenciaDigito.Text
            NovoReg.ContaCorrente = TxtConta.Text.Trim
            NovoReg.DigitoConta = TxtContaDigito.Text
            NovoReg.TipoConta = ddlTipoConta.SelectedValue
            NovoReg.Praca = TxtPraca.Text
            NovoReg.Cidade = ddlCidade.SelectedItem.Text
            NovoReg.Estado = ddlEstado.SelectedValue.ToString
            CarregaGrid(Cliente.Codigo, Cliente.CodigoEndereco)
            TabContainer1.ActiveTabIndex = 0
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Info)
        End If
    End Sub

    Protected Sub DG_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Selecionar()
    End Sub

    Public Overrides Sub Limpar()
        Session.Remove("_MainUserControl")
        DG.DataSource = New List(Of Object)()
        DG.DataBind()
    End Sub

    Protected Overrides Sub Selecionar()
        Try
            If DG.Rows.Count > 0 Then
                RecuperarSessao()
                Dim obj As [Lib].Negocio.ClienteXContaBancaria = Contas(DG.SelectedIndex)
                Session(Session("ssTipoRetorno")) = obj
                If Session("ssTipoRetorno") IsNot Nothing Then
                    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
                        Dim ucName = MainUserControl.ClientID.Split("_")
                        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
                        CType(uc, IBaseUserControl).Carregar(obj)
                    Else
                        CType(Me.Page, IBasePage).Carregar(obj)
                    End If
                    Popup.CloseDialog(Me.Page, "divConsultaDadosBancarios")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "")
        ddl.Carregar(ddlBanco, CarregarDDL.Tabela.Bancos, "")
        TabContainer1.ActiveTabIndex = 1
    End Sub

    Public Sub CarregaGrid(ByVal pCodigo As String, ByVal pEndereco As String)
        Session("ClienteConta" & HID.Value) = New [Lib].Negocio.Cliente(pCodigo, pEndereco)
        Cliente = Session("ClienteConta" & HID.Value)
        Contas = New [Lib].Negocio.ListClienteXContaBancaria(Cliente, True)
        DG.DataSource = Contas
        DG.DataBind()
        SalvarSessao()
        TabContainer1.ActiveTabIndex = 0
        TabContainer1.Tabs.RemoveAt(1)
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaDadosBancarios")
    End Sub

    Protected Sub btnClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnClose.Click
        Popup.CloseDialog(Me.Page, "divConsultaDadosBancarios")
    End Sub

End Class