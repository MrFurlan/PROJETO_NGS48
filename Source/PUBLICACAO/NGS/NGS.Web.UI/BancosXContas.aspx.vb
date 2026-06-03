Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class BancosXContas
    Inherits BasePage

    Private SqlArray As New ArrayList
    Private Sql As String

#Region "Métodos"

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlClienteDadosBancarios, CarregarDDL.Tabela.ClientesXEmpresas)
        Funcoes.VerificaEmpresa(DdlClienteDadosBancarios)
    End Sub

    Private Sub CargaBancos()
        ddl.Carregar(DdlBancos, CarregarDDL.Tabela.Bancos)
    End Sub

    Private Sub CargaContasCorrentes()
        If Funcoes.VerificaPermissao("BancosXContas", "LEITURA") Then
            LimparDadosBancarios()
            GridContasCorrentes.DataSource = Nothing
            GridContasCorrentes.DataBind()

            Sql = "    SELECT BancosXContas.Banco_Id, " & vbCrLf & _
                  "           BancosXContas.Agencia_Id, " & vbCrLf & _
                  "		      BancosXContas.DigitoAgencia_Id, " & vbCrLf & _
                  "		      BancosXContas.Conta_Id, " & vbCrLf & _
                  "		      BancosXContas.DigitoConta_Id, " & vbCrLf & _
                  "           BancosXContas.Observacoes, " & vbCrLf & _
                  "		      BancosXContas.ContaContabil, " & vbCrLf & _
                  "		      BancosXContas.Empresa_Id, " & vbCrLf & _
                  "		      BancosXContas.EndEmpresa_Id, " & vbCrLf & _
                  "		      Clientes.Reduzido, " & vbCrLf & _
                  "		      Isnull(TipoConta,'C') AS TipoConta, " & vbCrLf & _
                  "		      isnull(BancosXContas.Ativo, 0) as Ativo, " & vbCrLf & _
                  "           isnull(BancosXContas.FluxoDeCaixa, 1) as FluxoDeCaixa,  " & vbCrLf & _
                  "           LimiteBancario" & vbCrLf & _
                  "      FROM BancosXContas" & vbCrLf & _
                  "     INNER JOIN Clientes " & vbCrLf & _
                  "	       ON BancosXContas.Empresa_Id = Clientes.Cliente_Id " & vbCrLf & _
                  "	      AND BancosXContas.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf & _
                  "     WHERE Empresa_Id = '" & DdlClienteDadosBancarios.SelectedValue.Split("-")(0) & "'" & vbCrLf

            If String.IsNullOrWhiteSpace(DdlBancos.SelectedValue) Then
                Sql &= ""
            Else
                Sql &= "AND Banco_Id = " & DdlBancos.SelectedValue & ""
            End If

            Sql &= "     ORDER BY Clientes.Reduzido," & vbCrLf & _
                  "           BancosXContas.Banco_Id," & vbCrLf & _
                  "			  BancosXContas.Agencia_Id" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")

            If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado.")
            Else
                GridContasCorrentes.DataSource = ds
                GridContasCorrentes.DataBind()
            End If
        End If
    End Sub

    Private Sub LimparDadosBancarios()
        DdlClienteDadosBancarios.Enabled = True
        DdlBancos.Enabled = True
        txtAgencia.Enabled = True
        txtDigitoAgencia.Enabled = True
        txtContaCorrente.Enabled = True
        txtDigitoDaConta.Enabled = True

        txtAgencia.Text = ""
        txtDigitoAgencia.Text = ""
        txtContaCorrente.Text = ""
        txtDigitoDaConta.Text = ""
        ddlTipoConta.SelectedIndex = 0
        txtObservacoesDaConta.Text = ""
        DdlGrupoDeContas.SelectedIndex = 0
        ddlAtivo.SelectedIndex = 0
    End Sub

    Private Sub HabilitaAlteracaoDeCheques()
        Try
            Dim objBXC As New [Lib].Negocio.BancosXContas(DdlClienteDadosBancarios.SelectedValue.Split("-")(0), DdlClienteDadosBancarios.SelectedValue.Split("-")(1), _
                                                               DdlBancos.SelectedValue, txtAgencia.Text, txtDigitoAgencia.Text, txtContaCorrente.Text, txtDigitoDaConta.Text)
            txtNumChequeInicial.Text = objBXC.NumChequeInicial
            txtNumChequeFinal.Text = objBXC.NumChequeFinal
            txtNumChequeAtual.Text = objBXC.NumChequeAtual

            If objBXC.NumChequeAtual = objBXC.NumChequeFinal Then
                txtNumChequeInicial.Enabled = True
                txtNumChequeFinal.Enabled = True
                btnInutilizarCheques.Visible = False
            Else
                txtNumChequeInicial.Enabled = False
                txtNumChequeFinal.Enabled = False
                btnInutilizarCheques.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaChequeDigitado() As Boolean
        Try
            If txtNumChequeInicial.Enabled Then
                Dim objBXC As New [Lib].Negocio.BancosXContas(DdlClienteDadosBancarios.SelectedValue.Split("-")(0), DdlClienteDadosBancarios.SelectedValue.Split("-")(1), _
                                                                              DdlBancos.SelectedValue, txtAgencia.Text, txtDigitoAgencia.Text, txtContaCorrente.Text, txtDigitoDaConta.Text)
                If (CInt(txtNumChequeInicial.Text)) <= objBXC.NumChequeFinal Then
                    MsgBox(Me.Page, "Número inicial do cheque não pode ser menor ou igual ao último número de cheque registrado.")
                    Return False
                End If
            End If
            Return True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Sub CarregarContas()
        Dim strSQL As String

        strSQL = " SELECT Conta_Id as Codigo,  Conta_Id + ' - '  + Titulo as Descricao"
        strSQL &= " FROM PlanoDeContas"

        strSQL &= " WHERE  (LEN(Conta_Id) > 6)"

        DdlGrupoDeContas.DataValueField = "Codigo"
        DdlGrupoDeContas.DataTextField = "Descricao"
        DdlGrupoDeContas.DataSource = Banco.ConsultaDataSet(strSQL, "PlanoDeContas")
        DdlGrupoDeContas.DataBind()

        DdlGrupoDeContas.Items.Insert(0, "")
        DdlGrupoDeContas.SelectedIndex = 0
    End Sub

    Private Function ValidaCampos(Optional ByVal insert As Boolean = False) As Boolean
        If String.IsNullOrWhiteSpace(DdlClienteDadosBancarios.SelectedValue) Then
            MsgBox(Me.Page, "Empresa é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlBancos.SelectedValue) Then
            MsgBox(Me.Page, "Banco é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtAgencia.Text) Then
            MsgBox(Me.Page, "Agencia é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtContaCorrente.Text) Then
            MsgBox(Me.Page, "Conta é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDigitoDaConta.Text) Then
            MsgBox(Me.Page, "Digito da Conta é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlGrupoDeContas.SelectedValue) Then
            MsgBox(Me.Page, "C.Contábil é obrigatório")
            Return False
        ElseIf Not insert AndAlso (String.IsNullOrWhiteSpace(txtNumChequeInicial.Text) OrElse String.IsNullOrWhiteSpace(txtNumChequeFinal.Text) _
                    OrElse CInt(txtNumChequeInicial.Text) = 0 OrElse CInt(txtNumChequeFinal.Text) = 0) Then
            MsgBox(Me.Page, "Informe os Numeros de cheque (Inicial e Final)")
            Return False
        ElseIf Not insert AndAlso CInt(txtNumChequeInicial.Text) >= CInt(txtNumChequeFinal.Text) Then
            MsgBox(Me.Page, "Numero de cheque Inicial deve ser menor que Numero de cheque Final")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        CargaEmpresas()
        CarregarContas()
        CargaBancos()
        ddlAtivo.SelectedIndex = 0
        LimparDadosBancarios()
        Funcoes.VerificaEmpresa(DdlClienteDadosBancarios)

        txtLimiteBanc.Text = ""
        txtNumChequeAtual.Text = String.Empty
        txtNumChequeInicial.Text = String.Empty
        txtNumChequeFinal.Text = String.Empty
        txtNumChequeInicial.Enabled = True
        txtNumChequeFinal.Enabled = True
        btnInutilizarCheques.Visible = False

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        CargaContasCorrentes()

    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlClienteDadosBancarios.SelectedValue) Then
            Dim obj As Cliente = New Cliente(DdlClienteDadosBancarios.SelectedValue.Split("-")(0), DdlClienteDadosBancarios.SelectedValue.Split("-")(1))
            param &= "Empresa Cliente: " & Funcoes.FormatarCpfCnpj(obj.Codigo) & " - " & obj.Nome & " - " & obj.Cidade & " - " & obj.Estado.Codigo & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlBancos.SelectedItem.Text) Then
            param &= "Banco: " & DdlBancos.SelectedItem.Text
        End If

        Return param
    End Function

    Function validarCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlClienteDadosBancarios.SelectedValue) Then
            MsgBox(Me.Page, "Empresa deve ser Informada.")
            Return False
        End If
        If String.IsNullOrWhiteSpace(DdlBancos.SelectedValue) Then
            MsgBox(Me.Page, "Banco deve ser Informado.")
            Return False
        End If

        Return True
    End Function

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("BancosXContas", "ACESSAR") Then
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlClienteDadosBancarios_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlClienteDadosBancarios.SelectedIndexChanged
        Try
            'CargaBancos()
            CargaContasCorrentes()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnInutilizarCheques_Click(sender As Object, e As EventArgs) Handles btnInutilizarCheques.Click
        Dim param As New Dictionary(Of String, Object)
        param.Add("tipo", "I")
        param.Add("empresaNome", DdlClienteDadosBancarios.SelectedItem.Text)
        param.Add("empresaCodigo", DdlClienteDadosBancarios.SelectedValue)
        param.Add("banco", DdlBancos.SelectedValue)
        param.Add("nomeBanco", DdlBancos.SelectedItem.Text)
        param.Add("agencia", txtAgencia.Text)
        param.Add("digitoAgencia", txtDigitoAgencia.Text)
        param.Add("conta", txtContaCorrente.Text)
        param.Add("digitoConta", txtDigitoDaConta.Text)
        param.Add("ChequeAtual", IIf(txtNumChequeAtual.Text = 0, txtNumChequeInicial.Text, txtNumChequeAtual.Text + 1))
        param.Add("ChequeFinal", txtNumChequeFinal.Text)
        Session("MyParameters" & HID.Value) = param
        ucCancelamentoDeCheque.pageLoad()
        Popup.CancelamentoDeCheques(Me.Page, "objBancosXContas" & HID.Value)
        Limpar()
    End Sub

    Protected Sub GridContasCorrentes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridContasCorrentes.SelectedIndexChanged
        DdlBancos.SelectedIndex = DdlBancos.Items.IndexOf(DdlBancos.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(1).Text()))
        txtAgencia.Text = GridContasCorrentes.SelectedRow.Cells(2).Text()
        txtDigitoAgencia.Text = GridContasCorrentes.SelectedRow.Cells(3).Text()

        If Left(txtDigitoAgencia.Text, 1) = "&" Then
            txtDigitoAgencia.Text = ""
        End If

        txtContaCorrente.Text = GridContasCorrentes.SelectedRow.Cells(4).Text()
        txtDigitoDaConta.Text = GridContasCorrentes.SelectedRow.Cells(5).Text()

        If Left(txtDigitoDaConta.Text, 1) = "&" Then
            txtDigitoDaConta.Text = ""
        End If

        txtObservacoesDaConta.Text = GridContasCorrentes.SelectedRow.Cells(7).Text()
        If Left(txtObservacoesDaConta.Text, 1) = "&" Then
            txtObservacoesDaConta.Text = ""
        End If

        If GridContasCorrentes.SelectedRow.Cells(12).Text = "Sim" Then
            ddlAtivo.SelectedValue = 1
        Else
            ddlAtivo.SelectedValue = 0
        End If

        DdlGrupoDeContas.SelectedIndex = DdlGrupoDeContas.Items.IndexOf(DdlGrupoDeContas.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(8).Text()))
        DdlClienteDadosBancarios.SelectedIndex = DdlClienteDadosBancarios.Items.IndexOf(DdlClienteDadosBancarios.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(9).Text() & "-" & GridContasCorrentes.SelectedRow.Cells(10).Text()))

        ddlTipoConta.SelectedIndex = ddlTipoConta.Items.IndexOf(ddlTipoConta.Items.FindByValue(GridContasCorrentes.SelectedRow.Cells(6).Text()))
        txtLimiteBanc.Text = Server.HtmlDecode(GridContasCorrentes.SelectedRow.Cells(14).Text())

        If GridContasCorrentes.SelectedRow.Cells(13).Text = "Sim" Then
            chkFluxoDeCaixa.Checked = True
        Else
            chkFluxoDeCaixa.Checked = False
        End If

        DdlClienteDadosBancarios.Enabled = False
        DdlBancos.Enabled = False
        txtAgencia.Enabled = False
        txtDigitoAgencia.Enabled = False
        txtContaCorrente.Enabled = False
        txtDigitoDaConta.Enabled = False

        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True
        lnkExcluir.Parent.Visible = True

        HabilitaAlteracaoDeCheques()

    End Sub

    Protected Sub GridContasCorrentes_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridContasCorrentes.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            If e.Row.Cells(12).Text Then
                e.Row.Cells(12).Text = "Sim"
            Else
                e.Row.Cells(12).Text = "Não"
            End If
            If e.Row.Cells(13).Text Then
                e.Row.Cells(13).Text = "Sim"
            Else
                e.Row.Cells(13).Text = "Não"
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("BancosXContas", "LEITURA") Then

                CargaContasCorrentes()

                'Sql = "  SELECT     BancosXContas.Banco_Id, BancosXContas.Agencia_Id, BancosXContas.DigitoAgencia_Id, BancosXContas.Conta_Id, BancosXContas.DigitoConta_Id, " & vbCrLf & _
                '             " BancosXContas.Observacoes, BancosXContas.ContaContabil, BancosXContas.Empresa_Id, BancosXContas.EndEmpresa_Id, Clientes.Reduzido, Isnull(TipoConta,'C') AS TipoConta, isnull(BancosXContas.Ativo, 1) as Ativo" & vbCrLf & _
                '             " FROM       BancosXContas INNER JOIN" & vbCrLf & _
                '             "            Clientes ON BancosXContas.Empresa_Id = Clientes.Cliente_Id AND BancosXContas.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf & _
                '             " Where 1=1"

                'If DdlClienteDadosBancarios.SelectedIndex <> 0 Then
                '    Sql &= "And BancosXContas.Empresa_Id = " & DdlClienteDadosBancarios.SelectedValue.Split("-")(0) & vbCrLf
                'End If

                'If DdlBancos.SelectedIndex <> 0 Then
                '    Sql &= "And BancosXContas.Banco_Id = " & DdlBancos.SelectedValue.Split("-")(0) & vbCrLf
                'End If

                'Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")

                'If ds Is Nothing OrElse ds.Tables Is Nothing OrElse ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then
                '    MsgBox(Me.Page, "Nenhum Registro encontrado")
                'Else
                '    GridContasCorrentes.DataSource = ds
                '    GridContasCorrentes.DataBind()
                'End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("BancosXContas", "GRAVAR") Then
                If ValidaCampos(True) Then
                    Dim objBXC As New [Lib].Negocio.BancosXContas()
                    objBXC.IUD = "I"
                    objBXC.CodigoEmpresa = DdlClienteDadosBancarios.SelectedValue.Split("-")(0)
                    objBXC.EndEmpresa = DdlClienteDadosBancarios.SelectedValue.Split("-")(1)
                    objBXC.CodigoBanco = DdlBancos.SelectedValue
                    objBXC.Agencia = txtAgencia.Text
                    objBXC.DigitoAgencia = txtDigitoAgencia.Text
                    objBXC.Conta = txtContaCorrente.Text
                    objBXC.DigitoConta = txtDigitoDaConta.Text
                    objBXC.TipoConta = ddlTipoConta.SelectedValue
                    objBXC.CodigoContaContabil = DdlGrupoDeContas.SelectedValue
                    objBXC.Observacoes = txtObservacoesDaConta.Text
                    objBXC.UsuarioInclusao = Session("ssNomeUsuario")
                    objBXC.DataInclusao = Format(Date.Now, "yyyy-MM-dd")
                    objBXC.NumChequeInicial = IIf(String.IsNullOrWhiteSpace(txtNumChequeInicial.Text), 0, txtNumChequeInicial.Text)
                    objBXC.NumChequeFinal = IIf(String.IsNullOrWhiteSpace(txtNumChequeFinal.Text), 0, txtNumChequeFinal.Text)
                    objBXC.NumChequeAtual = 0
                    objBXC.FluxoDeCaixa = chkFluxoDeCaixa.Checked
                    objBXC.LimiteBanc = IIf(String.IsNullOrWhiteSpace(txtLimiteBanc.Text), 0, txtLimiteBanc.Text)
                    objBXC.Ativo = ddlAtivo.SelectedValue
                    objBXC.SalvarSql(SqlArray)

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Dados incluídos com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CargaContasCorrentes()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("BancosXContas", "ALTERAR") Then

                If String.IsNullOrWhiteSpace(txtNumChequeInicial.Text) Then txtNumChequeInicial.Text = 0
                If String.IsNullOrWhiteSpace(txtNumChequeFinal.Text) Then txtNumChequeFinal.Text = 0
                If String.IsNullOrWhiteSpace(txtNumChequeAtual.Text) Then txtNumChequeAtual.Text = 0

                If ValidaCampos() AndAlso ValidaChequeDigitado() Then
                    SqlArray = New ArrayList

                    Dim objBXC As New [Lib].Negocio.BancosXContas()
                    objBXC.IUD = "U"
                    objBXC.CodigoEmpresa = DdlClienteDadosBancarios.SelectedValue.Split("-")(0)
                    objBXC.EndEmpresa = DdlClienteDadosBancarios.SelectedValue.Split("-")(1)
                    objBXC.CodigoBanco = DdlBancos.SelectedValue
                    objBXC.Agencia = txtAgencia.Text
                    objBXC.DigitoAgencia = txtDigitoAgencia.Text
                    objBXC.Conta = txtContaCorrente.Text
                    objBXC.DigitoConta = txtDigitoDaConta.Text
                    objBXC.TipoConta = ddlTipoConta.SelectedValue
                    objBXC.CodigoContaContabil = DdlGrupoDeContas.SelectedValue
                    objBXC.Observacoes = txtObservacoesDaConta.Text
                    objBXC.UsuarioAlteracao = Session("ssNomeUsuario")
                    objBXC.DataAlteracao = Format(Date.Now, "yyyy-MM-dd")
                    objBXC.NumChequeInicial = txtNumChequeInicial.Text
                    objBXC.NumChequeFinal = txtNumChequeFinal.Text
                    objBXC.NumChequeAtual = txtNumChequeAtual.Text
                    objBXC.FluxoDeCaixa = chkFluxoDeCaixa.Checked
                    objBXC.LimiteBanc = IIf(String.IsNullOrWhiteSpace(txtLimiteBanc.Text), 0, txtLimiteBanc.Text)
                    objBXC.Ativo = ddlAtivo.SelectedValue
                    objBXC.SalvarSql(SqlArray)

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Dados atualizado com Sucesso.", eTitulo.Sucess)
                        Limpar()
                        CargaContasCorrentes()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("BancosXContas", "EXCLUIR") Then
                SqlArray = New ArrayList

                Dim Empresa() As String = DdlClienteDadosBancarios.SelectedValue.ToString.Split("-")

                Sql = "DELETE FROM BancosXContas" & vbCrLf & _
                    "       WHERE   Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "           And EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                    "           And Banco_Id = " & DdlBancos.SelectedValue & vbCrLf & _
                    "           And Agencia_Id = " & txtAgencia.Text & vbCrLf & _
                    "           And DigitoAgencia_ID = '" & txtDigitoAgencia.Text & "'" & vbCrLf & _
                    "           And Conta_Id =  " & txtContaCorrente.Text & vbCrLf & _
                    "           And DigitoConta_Id = '" & txtDigitoDaConta.Text & "'" & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Dados excluído com Sucesso.", eTitulo.Sucess)
                    LimparDadosBancarios()
                    CargaContasCorrentes()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("BancosXContas", "RELATORIO") Then
                If validarCampos() Then
                    Dim ds As New DataSet
                    Sql = "SELECT (C.Nome + '- ' + C.Cidade + '- ' + C.Estado) as Empresa, BC.EndEmpresa_id as EndEmpresa,            " & vbCrLf & _
                          "       BC.Empresa_Id as CodigoEmpresa, BC.Banco_id as Banco, Bancos.Descricao as NomeBanco,                " & vbCrLf & _
                          "       BC.Agencia_id as Agencia, BC.DigitoAgencia_id as DigitoAgencia, BC.Conta_id as Conta,               " & vbCrLf & _
                          "       BC.DigitoConta_id as DigitoConta, BC.ContaContabil as ContaContabil, BC.Observacoes as Observacao,  " & vbCrLf & _
                          "       Case When BC.Ativo = 1 Then 'Sim' Else 'Não' End as Ativo,                                          " & vbCrLf & _
                          "       Case When BC.FluxoDeCaixa = 1 Then 'Sim' Else 'Não' End as FluxoDeCaixa,                            " & vbCrLf & _
                          "       BC.LimiteBancario as LimiteBancario                                                                 " & vbCrLf & _
                          "    From BancosXContas as BC                                                                               " & vbCrLf & _
                          "    	INNER JOIN Clientes as C                                                                            " & vbCrLf & _
                          "    		ON BC.Empresa_id = C.Cliente_id                                                                 " & vbCrLf & _
                          "    	INNER JOIN Bancos                                                                                   " & vbCrLf & _
                          "    		ON BC.Banco_id = Bancos.Banco_id                                                                " & vbCrLf & _
                          "    where 1=1 " & vbCrLf

                    If Not String.IsNullOrWhiteSpace(DdlClienteDadosBancarios.SelectedValue) Then
                        Sql &= " AND BC.Empresa_id = " & DdlClienteDadosBancarios.SelectedValue.Split("-")(0) & vbCrLf & _
                               "    	and	BC.EndEmpresa_id = " & DdlClienteDadosBancarios.SelectedValue.Split("-")(1) & vbCrLf
                    End If
                    If Not String.IsNullOrWhiteSpace(DdlBancos.SelectedValue) Then
                        Sql &= "    	and BC.Banco_id = " & DdlBancos.SelectedValue & vbCrLf
                    End If

                    Sql &= "    ORDER BY BC.Empresa_id                                                                                 " & vbCrLf
                    ds = Banco.ConsultaDataSet(Sql, "BancosXContas")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("ConsultaParametros", getParam())

                    Funcoes.BindReport(Me.Page, ds, "Cr_BancosXContas", eExportType.PDF, parameters)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "BancosXContas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub


#End Region

End Class