Imports System.Data
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PrecoDeMercado
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Custos)
                If Funcoes.VerificaPermissao("PrecoDeMercado", "ACESSAR") Then
                    CargaEmpresas()
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    CargaDepositos()
                    CargaGrupo()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Custos.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CargaDepositos()
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.Depositos, "", True)
    End Sub

    Private Sub CargaGrupo()
        Dim lstGruposProduto As New [Lib].Negocio.ListGrupoProduto()

        If lstGruposProduto.SelecionarComProdutos() Then
            ddlGrupo.DataSource = lstGruposProduto.ToArray()
            ddlGrupo.DataTextField = "Descricao"
            ddlGrupo.DataValueField = "Codigo"
            ddlGrupo.DataBind()

            ddlGrupo.Items.Insert(0, "")
            ddlGrupo.SelectedIndex = 0
        End If
    End Sub

    Private Sub CargaProduto()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo= '" & ddlGrupo.SelectedValue & "'", True)
    End Sub

    Private Sub Limpar()
        ddlDeposito.SelectedIndex = 0
        ddlGrupo.SelectedIndex = 0
        ddlProduto.SelectedIndex = 0
        txtData.Text = Format(Today, "dd/MM/yyyy")
        txtValorOficial.Text = ""
        txtValorMoeda.Text = ""
        txtBaseDeCalculo.Text = ""
        txtDatade.Text = Format(Today, "01/MM/yyyy")
        txtDataAte.Text = Format(Today, "dd/MM/yyyy")
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        HabilitarCampos(True)
        gridPrecoDeMercado.DataBind()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCamposPrecoDeMercado() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlDeposito.SelectedValue) Then
            MsgBox(Me.Page, "Informe o depósito.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            MsgBox(Me.Page, "Informe o Produto.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) OrElse Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Informe uma data valida.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtValorOficial.Text) OrElse Not IsNumeric(txtValorOficial.Text) Then
            MsgBox(Me.Page, "Informe o valor Oficial.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtBaseDeCalculo.Text) OrElse Not IsNumeric(txtBaseDeCalculo.Text) Then
            MsgBox(Me.Page, "Informe a base de cálculo.")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub HabilitarCampos(ByVal enable As Boolean)
        ddlEmpresa.Enabled = enable
        ddlDeposito.Enabled = enable
        ddlProduto.Enabled = enable
        ddlGrupo.Enabled = enable
        txtData.Enabled = enable
    End Sub

    Protected Sub TxtValorMoeda_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim cot As New [Lib].Negocio.Cotacao(3, CDate(txtData.Text))
            txtValorOficial.Text = Math.Round(CDec(txtValorMoeda.Text) * cot.Indice, 6)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub TxtValorOficial_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim cot As New [Lib].Negocio.Cotacao(3, CDate(txtData.Text))
            txtValorMoeda.Text = Math.Round(CDec(txtValorOficial.Text) / cot.Indice, 6)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridPrecoDeMercado_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Empresa As String = Server.HtmlDecode(gridPrecoDeMercado.SelectedRow.Cells(1).Text())
            Dim Deposito As String = Server.HtmlDecode(gridPrecoDeMercado.SelectedRow.Cells(2).Text())
            Dim Produto As String = Server.HtmlDecode(gridPrecoDeMercado.SelectedRow.Cells(3).Text())
            Dim Data As String = Server.HtmlDecode(gridPrecoDeMercado.SelectedRow.Cells(4).Text())

            Dim objPrecoMercado As New [Lib].Negocio.TabelaPrecoDeMercado(Empresa, Deposito, Produto.Split("-")(0), Data)

            ddlEmpresa.SelectedValue = objPrecoMercado.Empresa & "-" & objPrecoMercado.EndEmpresa
            ddlDeposito.SelectedValue = objPrecoMercado.Deposito & "-" & objPrecoMercado.EndDeposito
            ddlGrupo.SelectedValue = objPrecoMercado.Produto.CodigoGrupo
            CargaProduto()
            ddlProduto.SelectedValue = objPrecoMercado.CodigoProduto
            txtData.Text = objPrecoMercado.Data_Id

            txtValorOficial.Text = objPrecoMercado.ValorOficial
            txtValorMoeda.Text = objPrecoMercado.ValorMoeda
            txtBaseDeCalculo.Text = objPrecoMercado.BaseDeCalculo

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            HabilitarCampos(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDeMercado", "GRAVAR") Then
                If ValidaCamposPrecoDeMercado() Then
                    Dim objPrecoMercado As New [Lib].Negocio.TabelaPrecoDeMercado
                    Dim Sqls As New ArrayList

                    objPrecoMercado.IUD = "I"
                    objPrecoMercado.Empresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objPrecoMercado.EndEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objPrecoMercado.Deposito = ddlDeposito.SelectedValue.Split("-")(0)
                    objPrecoMercado.EndDeposito = ddlDeposito.SelectedValue.Split("-")(1)
                    objPrecoMercado.CodigoProduto = ddlProduto.SelectedValue
                    objPrecoMercado.Data_Id = Format(CDate(txtData.Text), "yyyy-MM-dd")
                    objPrecoMercado.ValorOficial = CDec(txtValorOficial.Text)
                    objPrecoMercado.ValorMoeda = CDec(txtValorMoeda.Text)
                    objPrecoMercado.BaseDeCalculo = txtBaseDeCalculo.Text
                    If Not objPrecoMercado.JaExiste Then
                        If objPrecoMercado.Salvar Then
                            MsgBox(Me.Page, "sucesso na inclusão.", eTitulo.Sucess)
                            Limpar()
                        End If
                    Else
                        MsgBox(Me.Page, "Este registro já existe.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDeMercado", "ALTERAR") Then
                If ValidaCamposPrecoDeMercado() Then
                    Dim objPrecoMercado As New [Lib].Negocio.TabelaPrecoDeMercado
                    Dim Sqls As New ArrayList

                    objPrecoMercado.IUD = "U"
                    objPrecoMercado.Empresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    objPrecoMercado.EndEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    objPrecoMercado.Deposito = ddlDeposito.SelectedValue.Split("-")(0)
                    objPrecoMercado.EndDeposito = ddlDeposito.SelectedValue.Split("-")(1)
                    objPrecoMercado.CodigoProduto = ddlProduto.SelectedValue
                    objPrecoMercado.Data_Id = Format(CDate(txtData.Text), "yyyy-MM-dd")
                    objPrecoMercado.ValorOficial = CDec(txtValorOficial.Text)
                    objPrecoMercado.ValorMoeda = CDec(txtValorMoeda.Text)
                    objPrecoMercado.BaseDeCalculo = txtBaseDeCalculo.Text
                    If objPrecoMercado.JaExiste Then
                        If objPrecoMercado.Salvar Then
                            MsgBox(Me.Page, "Dados Alterados com Sucesso.", eTitulo.Sucess)
                            Limpar()
                        Else
                            MsgBox(Me.Page, "Este registro não existe.")
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para atulizar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDeMercado", "EXCLUIR") Then
                If ValidaCamposPrecoDeMercado() Then
                    Dim SalvarTPM As New [Lib].Negocio.TabelaPrecoDeMercado
                    Dim Sqls As New ArrayList

                    SalvarTPM.IUD = "D"
                    SalvarTPM.Empresa = ddlEmpresa.SelectedValue.Split("-")(0)
                    SalvarTPM.EndEmpresa = ddlEmpresa.SelectedValue.Split("-")(1)
                    SalvarTPM.Deposito = ddlDeposito.SelectedValue.Split("-")(0)
                    SalvarTPM.EndDeposito = ddlDeposito.SelectedValue.Split("-")(1)
                    SalvarTPM.CodigoProduto = ddlProduto.SelectedValue
                    SalvarTPM.Data_Id = Format(CDate(txtData.Text), "yyyy-MM-dd")
                    SalvarTPM.ValorOficial = CDec(txtValorOficial.Text)
                    SalvarTPM.ValorMoeda = CDec(txtValorMoeda.Text)
                    SalvarTPM.BaseDeCalculo = txtBaseDeCalculo.Text
                    If SalvarTPM.JaExiste Then
                        SalvarTPM.SalvarSql(Sqls)
                        If Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, "Dados excluídos com Sucesso.", eTitulo.Sucess)
                        End If
                    Else
                        MsgBox(Me.Page, "Este registro não existe.")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("PrecoDeMercado", "LEITURA") Then
                If ValidaCamposConsulta() Then
                    Dim ListPrecosDeMercado As New [Lib].Negocio.ListTabelaPrecoDeMercado(ddlEmpresa.SelectedValue, ddlDeposito.SelectedValue, ddlProduto.SelectedValue, txtDatade.Text, txtDataAte.Text)
                    If ListPrecosDeMercado.Count > 0 Then
                        gridPrecoDeMercado.DataSource = ListPrecosDeMercado
                        gridPrecoDeMercado.DataBind()
                    Else
                        MsgBox(Me.Page, "Registro(s) não encontrado(s) para esta seleção/período.")
                    End If
                End If

            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        Try
            If ValidaCamposConsulta() Then
                Dim ds As DataSet = getDataSet()
                Funcoes.BindReport(Me.Page, ds, "Cr_PrecoDeMercado", eExportType.PDF)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If ValidaCamposConsulta() Then
                Dim colunas As New Dictionary(Of String, eTipoCampo)
                colunas.Add("Data", eTipoCampo.Data)
                Dim ds As DataSet = getDataSet()
                Funcoes.BindExcelOffice(Me.Page, ds, "PrecoDeMercado", colunas)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCamposConsulta() As Boolean
        If String.IsNullOrWhiteSpace(txtDatade.Text) OrElse String.IsNullOrWhiteSpace(txtDataAte.Text) Then
            MsgBox(Me.Page, "Informe um período para realizar a consulta.")
            Return False
        ElseIf Not IsDate(txtDatade.Text) OrElse Not IsDate(txtDataAte.Text) Then
            MsgBox(Me.Page, "Informe um período válido para realizar a consulta.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataSet() As DataSet
        Dim sql As String = "  SELECT TPM.Empresa_Id as CodigoEmpresa,                 " & vbCrLf &
                            "         TPM.EndEmpresa_Id as EndEmpresa,                 " & vbCrLf &
                            "         Empresa.Nome as DescricaoEmpresa,                " & vbCrLf &
                            "         Empresa.Cidade as CidadeEmpresa," & vbCrLf &
                            "         Empresa.Estado as EstadoEmpresa," & vbCrLf &
                            "         TPM.Deposito_Id as CodigoDeposito,               " & vbCrLf &
                            "         TPM.EndDeposito_Id as EndDeposito,               " & vbCrLf &
                            "         Deposito.Nome AS DescricaoDeposito,              " & vbCrLf &
                            "         Deposito.Cidade as CidadeDeposito," & vbCrLf &
                            "         Deposito.Estado as EstadoDeposito," & vbCrLf &
                            "         TPM.Produto_Id as CodigoProduto,                 " & vbCrLf &
                            "         Produtos.Nome AS DescricaoProduto,               " & vbCrLf &
                            "         TPM.Data_Id as Data,                             " & vbCrLf &
                            "         TPM.ValorOficial,                                " & vbCrLf &
                            "         TPM.ValorMoeda,                                  " & vbCrLf &
                            "         TPM.BaseDeCalculo                                " & vbCrLf &
                            "    FROM TabelaDePrecosDeMercado TPM                      " & vbCrLf &
                            "   INNER JOIN Clientes AS Empresa                         " & vbCrLf &
                            "      ON TPM.Empresa_Id    = Empresa.Cliente_Id           " & vbCrLf &
                            "     AND TPM.EndEmpresa_Id = Empresa.Endereco_Id          " & vbCrLf &
                            "   INNER JOIN Clientes AS Deposito                        " & vbCrLf &
                            "      ON TPM.Deposito_Id    = Deposito.Cliente_Id         " & vbCrLf &
                            "     AND TPM.EndDeposito_Id = Deposito.Endereco_Id        " & vbCrLf &
                            "   INNER JOIN Produtos                                    " & vbCrLf &
                            "      ON TPM.Produto_Id = Produtos.Produto_Id             " & vbCrLf &
                            "   Where TPM.Data_Id Between '" & CDate(txtDatade.Text).ToString("yyyy/MM/dd") & "' and '" & CDate(txtDataAte.Text).ToString("yyyy/MM/dd") & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= "     and TPM.Empresa_id     ='" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   "     and TPM.EndEmpresa_id  = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlDeposito.SelectedValue) Then
            sql &= "     and TPM.Deposito_Id    = '" & ddlDeposito.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                   "     and TPM.EndDeposito_id = " & ddlDeposito.SelectedValue.Split("-")(1) & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then sql &= "     and TPM.Produto_Id     = '" & ddlProduto.SelectedValue & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PrecoDeMercado")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                row("CodigoEmpresa") = Funcoes.FormatarCpfCnpj(row("CodigoEmpresa"))
                row("CodigoDeposito") = Funcoes.FormatarCpfCnpj(row("CodigoDeposito"))
            Next
        End If

        Return ds
    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PrecoDeMercado")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class