Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class ContratoBancario
    Inherits BasePage

    Dim Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not Page.IsPostBack AndAlso IsConnect Then
                If Funcoes.VerificaPermissao("ParcelamentoDeContratoBancario", "ACESSAR") Then
                    ddl.Carregar(DdlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocio, ddlEmpresa)
                    ddl.Carregar(ddlFinalidade, CarregarDDL.Tabela.CarteiraFinanceira, "classificacao = 'P'")
                    txtMovimento.Text = Format(Today, "dd/MM/yyyy")
                    CargaBanco()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas)
    End Sub

    Private Sub CargaBanco()
        ddl.Carregar(ddlBanco, CarregarDDL.Tabela.Bancos)
    End Sub

    Private Sub limpar()
        txtContratoBanco.Text = String.Empty
        txtMovimento.Text = Format(Today, "dd/MM/yyyy")
        txtParcelas.Text = String.Empty
        txtValor.Text = String.Empty
        ddlBanco.SelectedValue = String.Empty

        ddlFinalidade.SelectedValue = String.Empty
        ddlEncargo.Items.Clear()

        GridParcBancario.DataBind()

        lnkConfirmar.Parent.Visible = True
        lnkNovo.Parent.Visible = False
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function validaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            MsgBox(Me.Page, "Informe o banco.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtContratoBanco.Text) Then
            MsgBox(Me.Page, "Informe o contrato.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlFinalidade.SelectedValue) Then
            MsgBox(Me.Page, "Informe a finalidade financeira.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtParcelas.Text) Then
            MsgBox(Me.Page, "Informe as parcelas.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtValor.Text) OrElse CDec(txtValor.Text) = 0 Then
            MsgBox(Me.Page, "Sem valor a Parcelar.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            MsgBox(Me.Page, "Informe o movimento.")
            Return False
        End If
        Return True
    End Function

    Protected Sub DdlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(sender As Object, e As EventArgs) Handles lnkConfirmar.Click
        Try
            If Funcoes.VerificaPermissao("ParcelamentoDeContratoBancario", "CONSULTAR") Then
                If validaCampos() AndAlso verificaContratosPagos() Then
                    Dim valorParcela As Decimal = Math.Round((txtValor.Text / txtParcelas.Text), 2)
                    Dim totalrest As Decimal = Math.Round(CDec(txtValor.Text), 2)
                    Dim Vencimento As DateTime = txtMovimento.Text

                    Dim coltitulo As New DataTable()
                    coltitulo.Columns.Add("Vencimento", GetType(System.DateTime))
                    coltitulo.Columns.Add("Titulo", Type.GetType("System.String"))
                    coltitulo.Columns.Add("ValorDoDocumento", Type.GetType("System.Double"))
                    coltitulo.Columns.Add("NrParcela", Type.GetType("System.String"))

                    For i = 1 To CInt(txtParcelas.Text)
                        Vencimento = Vencimento.AddMonths(1)

                        Dim rowtitulo As DataRow = coltitulo.NewRow()
                        rowtitulo("Vencimento") = Funcoes.ValidaDataUtil(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1), Vencimento)
                        rowtitulo("ValorDoDocumento") = valorParcela
                        rowtitulo("NrParcela") = "Parcela(s): " & i & " de " & txtParcelas.Text

                        If i = CInt(txtParcelas.Text) Then
                            rowtitulo("valorDoDocumento") = totalrest
                        End If

                        coltitulo.Rows.Add(rowtitulo)
                        totalrest -= valorParcela
                    Next

                    GridParcBancario.DataSource = coltitulo
                    GridParcBancario.DataBind()
                    lnkConfirmar.Parent.Visible = False
                    lnkNovo.Parent.Visible = True
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão parar gerar parcelas.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function verificaContratosPagos() As Boolean
        Dim sql As String = "select count(contratobancario) from Contasapagar where contratobancario = '" & txtContratoBanco.Text & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Contratos")

        If ds.Tables(0).Rows(0)(0) > 0 Then
            MsgBox(Me.Page, "Contrato ja parcelado.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("ParcelamentoDeContratoBancario", "GRAVAR") Then
                Dim Sqls As New ArrayList
                Dim contrato As New Titulo(getContrato())

                For Each row As GridViewRow In GridParcBancario.Rows
                    Dim Tit As New Titulo
                    Dim txtVencimento As TextBox = CType(row.FindControl("txtVencimento"), TextBox)

                    Tit.IUD = "I"
                    Tit.Sequencia = 0
                    Tit.CodigoProvisao = 2
                    Tit.ReceberPagar = "P"

                    Dim objCarteira As New [Lib].Negocio.CarteiraFinanceira(ddlFinalidade.SelectedValue)

                    Tit.CodigoCarteira = objCarteira.CodigoCarteira
                    Tit.ContaContabilCliente = objCarteira.CodigoContaCliente

                    Tit.CodigoCliente = ddlBanco.SelectedValue
                    Tit.ContratoBancario = txtContratoBanco.Text

                    Tit.Tributo = ddlEncargo.SelectedValue
                    Tit.CodigoUnidadeDeNegocio = DdlUnidadeDeNegocio.SelectedValue
                    Tit.CodigoEmpresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    Tit.EnderecoEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    Tit.CodigoCliente = contrato.CodigoCliente
                    Tit.EndCliente = contrato.EndCliente
                    Tit.CodigoEmpresaPedido = contrato.CodigoEmpresaPedido
                    Tit.EndEmpresaPedido = contrato.EndEmpresaPedido
                    Tit.CodigoIndexador = 3
                    Tit.CodigoMoeda = 1
                    Tit.CodigoTipoPgto = 1
                    Tit.CodigoSituacao = 1
                    Tit.Lote = 70
                    Tit.Movimento = txtMovimento.Text
                    Tit.Vencimento = txtVencimento.Text
                    Tit.Prorrogacao = txtVencimento.Text
                    Tit.DataMoeda = txtMovimento.Text
                    Tit.Baixa = txtVencimento.Text
                    Tit.CodigoBancoCliente = 0
                    Tit.CodigoAgenciaCliente = ""
                    Tit.DigitoAgenciaCliente = ""
                    Tit.ContaCliente = ""
                    Tit.DigitoContaCliente = ""
                    Tit.Cheque = False
                    Tit.Slips = False
                    Tit.Recibo = False
                    Tit.Aviso = False
                    Tit.ReciboDeposito = False
                    Tit.ValorDoDocumento = row.Cells(2).Text

                    Tit.Historico = "Parcela " & row.RowIndex + 1 & "/" & txtParcelas.Text & " referente ao contrato (" & txtContratoBanco.Text & ")"
                    Tit.CodigoDeBarras = ""
                    Tit.CodigoDigitado = False
                    Tit.CodigoDeBarrasPreImpresso = False

                    Tit.CodigoDestinatario = contrato.CodigoDestinatario
                    Tit.EndDestinatario = contrato.EndDestinatario
                    Tit.Destinacao = ""
                    Tit.Solicitacao = 0

                    Tit.Agrupado = eAgrupamentoFinanceiro.NaoAgrupado
                    Tit.RegistroMestre = 0
                    Tit.SituacaoBancaria = 0
                    Tit.NumeroDoCheque = 0

                    Tit.UsuarioInclusao = HttpContext.Current.Session("ssNomeUsuario")
                    Tit.UsuarioInclusaoData = Now()

                    Tit.Salvar()
                Next
                MsgBox(Me.Page, "Parcelamento inserido com Sucesso.", eTitulo.Sucess)

                lnkRelatorio_Click(sender, e)

                limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão parar inserir.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = "   select t.ContratoBancario, b.Descricao as Banco, t.Registro_Id as Titulo, t.Movimento, t.Vencimento,                                                  " & vbCrLf & _
                            "          t.Prorrogacao, t.Empresa as CodEmpresa, t.EndEmpresa, Emp.Nome + ' - ' + Emp.Endereco as DescricaoEmpresa,                                     " & vbCrLf & _
                            "   	   t.Cliente as CodCliente, t.EndCliente, cli.Nome + ' - ' + Cli.Endereco as DescricaoCliente, t.ValorDoDocumento, t.ValorLiquido,                " & vbCrLf & _
                            "   	   t.Juros + Acrescimos as JurosAcrescimos, Descontos + t.Deducoes as DescontosDeducoes, Historico                                                 " & vbCrLf & _
                            "     from ContasAreceber t                                                                                                                               " & vbCrLf & _
                            "    Inner Join BancosXContas bc                                                                                                                          " & vbCrLf & _
                            "       on bc.Banco_Id = t.BancoPagador                                                                                                                   " & vbCrLf & _
                            "      And bc.Agencia_Id = t.AgenciaPagadora                                                                                                              " & vbCrLf & _
                            "      And bc.DigitoAgencia_Id = t.DigitoAgenciaPagadora                                                                                                  " & vbCrLf & _
                            "      And bc.Conta_Id = t.ContaPagadora                                                                                                                  " & vbCrLf & _
                            "    Inner Join Bancos b                                                                                                                                  " & vbCrLf & _
                            "       on b.Banco_Id = bc.Banco_Id                                                                                                                       " & vbCrLf & _
                            "    Inner Join Clientes Emp                                                                                                                              " & vbCrLf & _
                            "       on Emp.Cliente_Id  = t.Empresa                                                                                                                    " & vbCrLf & _
                            "      and Emp.Endereco_id = t.EndEmpresa                                                                                                                 " & vbCrLf & _
                            "    Inner Join Clientes Cli                                                                                                                              " & vbCrLf & _
                            "       on Cli.Cliente_Id  = t.Cliente                                                                                                                    " & vbCrLf & _
                            "      and Cli.Endereco_id = t.EndCliente                                                                                                                 " & vbCrLf & _
                            "    where t.Situacao = 1                                                                                                                                 " & vbCrLf & _
                            "      and isnull(t.ContratoBancario, '') <> ''                                                                                                           " & vbCrLf & _
                            "      and t.ContratoBancario = '" & txtContratoBanco.Text & "'                                                                                           " & vbCrLf & _
                            "      and t.Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                            "                                                                                                                                                         " & vbCrLf & _
                            "   select t.ContratoBancario, t.Registro_Id as Titulo, case when t.provisao = 1 then 'Pago' else 'Pendente' end as Situacao, t.Movimento, t.Vencimento,  " & vbCrLf & _
                            "          t.Prorrogacao, t.Empresa as CodEmpresa, t.EndEmpresa, Emp.Nome + ' - ' + Emp.Endereco as DescricaoEmpresa,                                     " & vbCrLf & _
                            "   	   t.Cliente as CodCliente, t.EndCliente, cli.Nome + ' - ' + Cli.Endereco as DescricaoCliente, t.ValorDoDocumento, t.ValorLiquido,                " & vbCrLf & _
                            "   	   t.Juros + Acrescimos as JurosAcrescimos, Descontos + t.Deducoes as DescontosDeducoes, Historico                                                 " & vbCrLf & _
                            "     from ContasAPagar t                                                                                                                                 " & vbCrLf & _
                            "    Inner Join Clientes Emp                                                                                                                              " & vbCrLf & _
                            "       on Emp.Cliente_Id  = t.Empresa                                                                                                                    " & vbCrLf & _
                            "      and Emp.Endereco_id = t.EndEmpresa                                                                                                                 " & vbCrLf & _
                            "    Inner Join Clientes Cli                                                                                                                              " & vbCrLf & _
                            "       on Cli.Cliente_Id  = t.Cliente                                                                                                                    " & vbCrLf & _
                            "      and Cli.Endereco_id = t.EndCliente                                                                                                                 " & vbCrLf & _
                            "    where t.Situacao = 1                                                                                                                                 " & vbCrLf & _
                            "      and isnull(t.ContratoBancario, '') <> ''                                                                                                           " & vbCrLf & _
                            "      and t.ContratoBancario = '" & txtContratoBanco.Text & "'                                                                                           " & vbCrLf & _
                            "      and t.Empresa = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'"
        Return Banco.ConsultaDataSet(sql, "TitulosReceber")
    End Function

    Private Function getParam()
        Dim param As String = ""

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            param &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            param &= "Banco: " & ddlBanco.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtContratoBanco.Text) Then
            param &= "Contrato: " & txtContratoBanco.Text & "-"
        End If
        If Not String.IsNullOrWhiteSpace(txtMovimento.Text) Then
            param &= "Movimento: " & txtMovimento.Text
        End If

        Return param
    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ParcelamentoDeContratoBancario")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub txtContratoBanco_TextChanged(sender As Object, e As EventArgs) Handles txtContratoBanco.TextChanged
        Try
            If Not String.IsNullOrWhiteSpace(txtContratoBanco.Text) Then
                vereficaContrato()
            Else : txtValor.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub vereficaContrato()
        Dim sql As String = "select isnull(sum(valordodocumento), 0) as valor from ContasAReceber where isnull(ContratoBancario, '') = '" & txtContratoBanco.Text & "'" & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ContasAReceber")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows(0)(0) > 0 Then
            txtValor.Text = ds.Tables(0).Rows(0)(0)
        Else
            MsgBox(Me.Page, "Contrato inexistente.")
            txtValor.Text = String.Empty
        End If
    End Sub

    Private Function getContrato() As Integer
        Dim sql As String = "select top 1 registro_id from ContasAReceber where isnull(ContratoBancario, '') = '" & txtContratoBanco.Text & "'" & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ContasAReceber")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count Then
            Return ds.Tables(0).Rows(0)(0)
        Else
            Return 0
        End If
    End Function

    Private Function ValidarEmissao() As Boolean
        If String.IsNullOrWhiteSpace(ddlBanco.SelectedValue) Then
            MsgBox(Me.Page, "Informe o banco.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtContratoBanco.Text) Then
            MsgBox(Me.Page, "Informe o contrato.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If ValidarEmissao() Then
                Dim ds As DataSet = getDataSet()
                ds.Tables(1).TableName = "TitulosPagar"

                Dim parametros As New Dictionary(Of String, Object)()
                parametros.Add("ParametrosConsulta", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_ContratoBancario", eExportType.PDF, parametros)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlFinalidade_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFinalidade.SelectedIndexChanged
        Try
            If ddlFinalidade.SelectedIndex > 0 Then
                ddl.Carregar(ddlEncargo, CarregarDDL.Tabela.CarteirasXTributos, "Carteira_Id = '" & ddlFinalidade.SelectedValue & "'", True)
            Else : ddlEncargo.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class