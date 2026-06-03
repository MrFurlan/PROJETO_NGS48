Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ConsistenciaDeProducao
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ConsistenciaDeProducao", "ACESSAR") Then
                CargaUnidade()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id" & vbCrLf & _
              " WHERE ClientesXTipos.Tipo_Id = 050 Order by Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Nome"), Dr("Codigo")))
        Next

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub CargaEmpresa()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next
        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        ucSelecaoProduto.Limpar()

        txtDataInicial.Text = Format(Today, "01/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        chkFisico.Checked = True
        chkFiscal.Checked = True

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlUnidade.SelectedValue) Then
            param &= "Unidade: " & DdlUnidade.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            param &= "Empresa: " & DdlEmpresa.SelectedValue & " - "
        End If
        If chkFisico.Checked AndAlso chkFiscal.Checked Then
            param &= "Estoque: Físico - Fiscal" & vbCrLf
        ElseIf chkFiscal.Checked Then
            param &= "Estoque: Fiscal" & vbCrLf
        ElseIf chkFisico.Checked Then
            param &= "Estoque: Físico" & vbCrLf
        End If
        If CkCDeposito.Checked Then
            param &= "Consolidar Depósito: Consolidado " & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Período Inicial: " & txtDataInicial.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "Período Final: " & txtDataFinal.Text & vbCrLf
        End If

        Return param
    End Function

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Validar()
        Mensagem = ""

        If DdlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório."
            Return Mensagem
        End If
        If DdlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório."
            Return Mensagem
        End If

        Return Mensagem
    End Function

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("ConsistenciaDeProducao", "RELATORIO") Then
                Validar()
                If Mensagem = "" Then

                    Cliente = DdlEmpresa.SelectedValue
                    campo = Cliente.Split("-")
                    Sql = "SELECT Clientes.Reduzido, Clientes.Nome AS NomeDoDeposito, Clientes.Cidade, Clientes.Estado, Producao.Produto_Id AS Produto, Produtos.Nome AS NomeDoProduto, " & vbCrLf & _
                          "       Producao.Operacao_Id AS Operacao, Producao.SubOperacao_Id AS SubOperacao, SubOperacoes.Descricao AS NomeDaOperacao, Producao.Movimento_Id AS Movimento," & vbCrLf & _
                          "       sum(case when Producao.FisicoFiscal_Id = 1 then Producao.Entradas else 0 end) as EntradaFisico, sum(case when Producao.FisicoFiscal_Id = 1 then 0 else Producao.Entradas end)  as EntradaFiscal," & vbCrLf & _
                          "       sum(case when Producao.FisicoFiscal_Id = 1 then Producao.Saidas else 0 end) as SaidaFisico, sum(case when Producao.FisicoFiscal_Id = 1 then 0 else Producao.Saidas end) as SaidaFiscal," & vbCrLf & _
                          "       ISNULL(ProducaoXAnalises.Analise_Id, 0) AS Analise, ISNULL(Analises.Descricao, N'') AS NomeDaAnalise, ISNULL(ProducaoXAnalises.Indice, 0) AS Indice, " & vbCrLf & _
                          "       Producao.Lote_Id + case when len(Producao.Lote_Id)>0 then '-' else '' end +Producao.Classificacao_Id +'('+ Embalagens.EmbalagemIndea +'-'+ Producao.TipoDeEmbalagem_Id+' / '+ Convert(nvarchar(10),isnull(Producao.CapacidadeEmbalagem_Id,''))+')' as LoteEmbalagem, ProdutoDerivado.Produto_Id + '-' + ProdutoDerivado.Nome as ProdutoDerivado" & vbCrLf & _
                          "  FROM Clientes" & vbCrLf & _
                          " INNER JOIN Producao" & vbCrLf

                    If CkCDeposito.Checked = False Then
                        Sql &= "    ON Clientes.Cliente_Id  = Producao.Deposito_Id " & vbCrLf & _
                               "   AND Clientes.Endereco_Id = Producao.EndDeposito_Id " & vbCrLf
                    Else
                        Sql &= "    ON Clientes.Cliente_Id  = Producao.Empresa_Id " & vbCrLf & _
                               "   AND Clientes.Endereco_Id = Producao.EndEmpresa_Id " & vbCrLf
                    End If
                    Sql &= " INNER JOIN Produtos " & vbCrLf & _
                          "    ON Producao.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
                          " INNER JOIN SubOperacoes " & vbCrLf & _
                          "    ON Producao.Operacao_Id    = SubOperacoes.Operacao_Id " & vbCrLf & _
                          "   AND Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                          " LEFT JOIN Produtos ProdutoDerivado                                 " & vbCrLf & _
                          "     On ProdutoDerivado.Produto_Id = Producao.ProdutoDerivado_Id    " & vbCrLf & _
                          "  LEFT OUTER JOIN ProducaoXAnalises " & vbCrLf & _
                          "    ON Producao.Empresa_Id             = ProducaoXAnalises.Empresa_Id " & vbCrLf & _
                          "   AND Producao.EndEmpresa_Id          = ProducaoXAnalises.EndEmpresa_Id " & vbCrLf & _
                          "   AND Producao.Deposito_Id            = ProducaoXAnalises.Deposito_Id " & vbCrLf & _
                          "   AND Producao.EndDeposito_Id         = ProducaoXAnalises.EndDeposito_Id " & vbCrLf & _
                          "   AND Producao.Produto_Id             = ProducaoXAnalises.Produto_Id " & vbCrLf & _
                          "   AND Producao.Operacao_Id            = ProducaoXAnalises.Operacao_Id " & vbCrLf & _
                          "   AND Producao.SubOperacao_Id         = ProducaoXAnalises.SubOperacao_Id " & vbCrLf & _
                          "   AND Producao.Movimento_Id           = ProducaoXAnalises.Movimento_Id " & vbCrLf & _
                          "   AND Producao.FisicoFiscal_Id        = ProducaoXAnalises.FisicoFiscal_Id " & vbCrLf & _
                          "   AND Producao.ProdutoDerivado_Id     = ProducaoXAnalises.ProdutoDerivado_Id" & vbCrLf & _
                          "   And Producao.Lote_Id                = ProducaoXAnalises.Lote_Id" & vbCrLf & _
                          "   And Producao.Classificacao_Id       = ProducaoXAnalises.Classificacao_Id" & vbCrLf & _
                          "   And Producao.Embalagem_Id           = ProducaoXAnalises.Embalagem_Id" & vbCrLf & _
                          "   And Producao.TipoDeEmbalagem_Id     = ProducaoXAnalises.TipoDeEmbalagem_Id" & vbCrLf & _
                          "   And Producao.CapacidadeEmbalagem_Id = ProducaoXAnalises.CapacidadeEmbalagem_Id" & vbCrLf & _
                          "  LEFT OUTER JOIN Analises " & vbCrLf & _
                          "    ON Analises.Analise_Id = ProducaoXAnalises.Analise_Id " & vbCrLf & _
                          "  Left Join Embalagens" & vbCrLf & _
                          "    on Embalagens.Embalagem_id = Producao.Embalagem_Id" & vbCrLf

                    Sql &= " WHERE Producao.Empresa_Id = '" & campo(0) & "' And Producao.EndEmpresa_Id = " & campo(1)
                    Sql &= " And Producao.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'"

                    If rdAtivo.Checked Then
                        Sql &= " AND Produtos.Situacao in(1) " & vbCrLf
                    Else
                        Sql &= " AND Produtos.Situacao not in(1) " & vbCrLf
                    End If

                    If ucSelecaoProduto.TemSelecionado Then
                        Dim RetornoProdutos As ArrayList
                        RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Producao.Produto_Id", "")
                        Sql &= " AND " & RetornoProdutos(0)
                    End If

                    If chkFisico.Checked = True And chkFiscal.Checked = False Then
                        Sql &= " And (Producao.FisicoFiscal_Id = 1)"
                    End If

                    If chkFiscal.Checked = True Then
                        If chkFisico.Checked = True Then
                            Sql &= " And (Producao.FisicoFiscal_Id = 1 Or Producao.FisicoFiscal_Id = 2)"
                        Else
                            Sql &= " And (Producao.FisicoFiscal_Id = 2)"
                        End If
                    End If

                    Sql &= "group by Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Producao.Produto_Id, Produtos.Nome, " & vbCrLf & _
                           "         Producao.Operacao_Id, Producao.SubOperacao_Id, SubOperacoes.Descricao, Producao.Movimento_Id," & vbCrLf & _
                           "         ISNULL(ProducaoXAnalises.Analise_Id, 0), ISNULL(Analises.Descricao, N''), ISNULL(ProducaoXAnalises.Indice, 0), " & vbCrLf & _
                           "         Producao.Lote_Id + case when len(Producao.Lote_Id) > 0 then '-' else '' end + Producao.Classificacao_Id +'('+ Embalagens.EmbalagemIndea +'-'+ Producao.TipoDeEmbalagem_Id+' / '+ Convert(nvarchar(10), " & vbCrLf & _
                           "         isnull(Producao.CapacidadeEmbalagem_Id,''))+')', ProdutoDerivado.Produto_Id + '-' + ProdutoDerivado.Nome" & vbCrLf


                    DS = Banco.ConsultaDataSet(Sql, "ConsistenciaDeProducao")

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("Titulo", "Relatório De Consistência De Produção.")
                    parameters.Add("ConsultaParametros", getParam())

                    Funcoes.BindReport(Me.Page, DS, "Cr_ConsistenciaDeProducao", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            Funcoes.Ajuda(Me.Page, "ConsistenciaDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class