Imports NGS.Lib.Negocio
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class RelatorioNotasNaoContabilizadas
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioNotasNaoContabilizadas", "ACESSAR") Then
                    Unidade()
                    Limpar()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParametros() As String
        Dim param As String = "Parametros:" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param = "Unidade: " & ddlUnidade.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Data de: " & txtDataInicial.Text & " "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "à : " & txtDataFinal.Text
        End If

        Return param
    End Function

    Private Function getDataSet() As DataSet
        Dim ds As New DataSet()
        Dim sql As String = "           SELECT	nXe.Empresa_Id + '-' + CAST(nxe.EndEmpresa_Id AS VARCHAR) + '-' + emp.Nome AS Empresa,                   " & vbCrLf & _
        "   		nXe.Cliente_Id + '-' + CAST(nXe.EndCliente_Id AS VARCHAR) + '-' + Cli.Nome AS Cliente,                           " & vbCrLf & _
        "   		nxe.EntradaSaida_Id as EntradaSaida, cast(nxe.Nota_Id as varchar) + '-' + nxe.Serie_Id as Nota, n.Movimento,     " & vbCrLf & _
        "   		cast(nxi.Operacao as varchar) + '-' + cast(nxi.SubOperacao as varchar) + '-' + sO.Descricao as Operacao,         " & vbCrLf & _
        "           nxe.Produto_Id + '-' + Prd.Nome as Produto , nxe.Encargo_Id as Encargo, nXe.Valor,                               " & vbCrLf & _
        "   		isnull(r.DebitoOficial,0) AS DebitoOficial, isnull(r.CreditoOficial,0) AS CreditoOficial,                        " & vbCrLf & _
        "   		isnull(r.Historico,'NAO CONTABILIZOU') AS Historico                                                              " & vbCrLf & _
        "   FROM NotasFiscaisXEncargos nXe                                                                                           " & vbCrLf & _
        "     inner join NotasFiscaisXItens nXi                                                                                      " & vbCrLf & _
        "       on nXi.Empresa_Id = nXe.Empresa_Id                                                                                   " & vbCrLf & _
        "       and nXi.EndEmpresa_Id = nxe.EndEmpresa_Id                                                                            " & vbCrLf & _
        "       and nxi.Cliente_Id = nXe.Cliente_Id                                                                                  " & vbCrLf & _
        "       and nXi.EndCliente_Id = nXe.EndCliente_Id                                                                            " & vbCrLf & _
        "       and nxi.EntradaSaida_Id = nxe.EntradaSaida_Id                                                                        " & vbCrLf & _
        "       and nxi.Serie_Id = nxe.Serie_Id                                                                                      " & vbCrLf & _
        "       and nxi.Nota_Id = nxe.Nota_Id                                                                                        " & vbCrLf & _
        "       and nxi.Produto_Id = nxe.Produto_Id                                                                                  " & vbCrLf & _
        "       and nxi.CFOP_Id = nxe.CFOP_Id                                                                                        " & vbCrLf & _
        "     inner join NotasFiscais n                                                                                              " & vbCrLf & _
        "       on n.Empresa_Id = nxi.Empresa_Id                                                                                     " & vbCrLf & _
        "       and n.EndEmpresa_Id = nxi.EndEmpresa_Id                                                                              " & vbCrLf & _
        "       and n.Cliente_Id = nxi.Cliente_Id                                                                                    " & vbCrLf & _
        "       and n.EndCliente_Id = nxi.EndCliente_Id                                                                              " & vbCrLf & _
        "       and n.EntradaSaida_Id = nxi.EntradaSaida_Id                                                                          " & vbCrLf & _
        "       and n.Serie_Id = nxi.Serie_Id                                                                                        " & vbCrLf & _
        "       and n.Nota_Id = nxi.Nota_Id                                                                                          " & vbCrLf & _
        "     inner join SubOperacoes sO                                                                                             " & vbCrLf & _
        "   	on sO.Operacao_Id = nxi.Operacao                                                                                     " & vbCrLf & _
        "   	and sO.SubOperacoes_Id = nxi.SubOperacao                                                                             " & vbCrLf & _
        "     inner join Produtos Prd                                                                                                " & vbCrLf & _
        "   	on Prd.Produto_Id = nxe.Produto_Id                                                                                   " & vbCrLf & _
        "     Inner Join Clientes Emp                                                                                                " & vbCrLf & _
        "   	on emp.Cliente_Id = nXe.Empresa_Id                                                                                   " & vbCrLf & _
        "   	And emp.Endereco_Id = nxe.EndEmpresa_Id                                                                              " & vbCrLf & _
        "     Inner Join Clientes Cli                                                                                                " & vbCrLf & _
        "   	on Cli.Cliente_Id = nXe.Cliente_Id                                                                                   " & vbCrLf & _
        "   	And Cli.Endereco_Id = nxe.EndCliente_Id                                                                              " & vbCrLf & _
        "     left join Razao r                                                                                                      " & vbCrLf & _
        "   	on r.Empresa_Id = nXe.Empresa_Id                                                                                     " & vbCrLf & _
        "   	and r.EndEmpresa_Id = nxe.EndEmpresa_Id                                                                              " & vbCrLf & _
        "   	and r.Cliente_NF = nxe.Cliente_Id                                                                                    " & vbCrLf & _
        "   	and r.EndCliente_Nf = nxe.EndCliente_Id                                                                              " & vbCrLf & _
        "   	and r.EntradaSaida_Nf = nxe.EntradaSaida_Id                                                                          " & vbCrLf & _
        "   	and r.Serie_Nf = nxe.Serie_Id                                                                                        " & vbCrLf & _
        "   	and r.Numero_Nf = nxe.Nota_Id                                                                                        " & vbCrLf & _
        "   	and r.Encargo_NF = nxe.Encargo_Id                                                                                    " & vbCrLf & _
        "   where n.Situacao = 1                                                                                                     " & vbCrLf & _
        "   and r.Historico is null                                                                                                  " & vbCrLf & _
        "   and nXe.Valor > 0                                                                                                        " & vbCrLf & _
        "   and sO.Contabil = 'S'                                                                                                    " & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= "   and n.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'                                            " & vbCrLf & _
        "   and n.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & "                                                     " & vbCrLf
        End If

        sql &= "   and n.Movimento between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
        "   order by nxe.Empresa_Id, n.Movimento, n.Nota_Id   "

        ds = Banco.ConsultaDataSet(sql, "NotasNaoContabilizadas")
        Return ds

    End Function

    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe um intervalo de datas para pesquisa.")
            Return False
        ElseIf DateTime.Compare(txtDataInicial.Text, txtDataFinal.Text) > 0 Then
            MsgBox(Me.Page, "Intervalo de datas inválido.")
            Return False
        End If
        Return True
    End Function

    Private Sub Limpar()
        txtDataInicial.Text = Now().ToString("01/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        'ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        'Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub Unidade()
        'ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)

        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue)
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioNotasNaoContabilizadas", "RELATORIO") Then
                If ValidaCampos() Then
                    Dim ds As DataSet
                    ds = getDataSet()

                    Dim param As New Dictionary(Of String, Object)()
                    param.Add("Parametros", getParametros())
                    param.Add("Titulo", "RELATÓRIO DE NOTAS NÃO CONTABILIZADAS")

                    Funcoes.BindReport(Me.Page, ds, "Cr_NotasNaoContabilizadas", eExportType.PDF, param)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioNotasNaoContabilizadas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class