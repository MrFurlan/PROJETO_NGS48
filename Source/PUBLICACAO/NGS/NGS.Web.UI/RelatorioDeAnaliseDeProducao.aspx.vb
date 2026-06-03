Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioDeAnaliseDeProducao
    Inherits BasePage

    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeAnaliseDeProducao", "ACESSAR") Then
                    CargaUnidadeDeNegocio()
                    BuscarGrupoDeProdutos()
                    LimparCampos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddl.Carregar(DdlUnidadeDeNegocioOrigem, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocioOrigem, DdlEmpresaClienteOrigem)
    End Sub

    Private Sub BuscarGrupoDeProdutos()
        Dim ds As New DataSet

        sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao" & vbCrLf & _
              " FROM         GruposDeEstoques INNER JOIN" & vbCrLf & _
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf & _
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao" & vbCrLf
        ds = Banco.ConsultaDataSet(sql, "GruposDeEstoques")

        For Each drGrupo As DataRow In ds.Tables(0).Rows
            DdlGrupoDeProdutos.Items.Add(New ListItem(drGrupo("Descricao"), drGrupo("Grupo_Id")))
        Next

        DdlGrupoDeProdutos.Items.Insert(0, "")
        DdlGrupoDeProdutos.SelectedIndex = 0
    End Sub

    Private Sub BuscarProdutos()
        Dim ds As New DataSet

        sql = "SELECT Produto_Id, Descricao FROM Produtos WHERE Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "'"
        ds = Banco.ConsultaDataSet(sql, "Produtos")

        For Each drProduto As DataRow In ds.Tables(0).Rows
            DdlProdutos.Items.Add(New ListItem(Funcoes.AlinharEsquerda(drProduto("Descricao"), 50, ".") & " - " & drProduto("Produto_Id"), _
                                                   drProduto("Produto_Id")))
        Next

        DdlProdutos.Items.Insert(0, "")
        DdlProdutos.SelectedIndex = 0
    End Sub

    Private Sub LimparCampos()
        txtPeriodoInicialConsultaTitulos.Text = Format(Today, "01/01/" & "yyyy")
        txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
        DdlFabrica.SelectedValue = ""
        DdlGrupoDeProdutos.SelectedValue = ""
        DdlProdutos.SelectedValue = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocioOrigem.Enabled = False
            DdlEmpresaClienteOrigem.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidadeDeNegocioOrigem_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresaClienteOrigem, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocioOrigem.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupoDeProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            DdlProdutos.Items.Clear()
            BuscarProdutos()
            DdlProdutos.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioDeAnaliseDeProducao", "RELATORIO") Then
                Dim linha As String
                Dim sql As String
                Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
                Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
                Dim ds As New DataSet
                Dim dr As DataRow

                Dim strm As StreamWriter = Nothing
                If Dir(arquivo).Length > 0 Then Kill(arquivo)

                Dim Empresa As String
                Dim Fabrica As String

                Dim Grupo As String
                Dim Produto As String

                Dim cliente As String

                Dim campo() As String

                If DdlFabrica.Text <> "" Then
                    Fabrica = DdlFabrica.SelectedValue
                Else
                    Fabrica = ""          'Fabrica/Deposito
                End If

                If DdlGrupoDeProdutos.Text <> "" Then
                    Grupo = DdlGrupoDeProdutos.SelectedValue
                Else
                    Grupo = ""                   'Grupo de Produto
                End If

                If DdlProdutos.Text <> "" Then
                    Produto = DdlProdutos.SelectedValue
                Else
                    Produto = ""                   'UnidadeDeNegocio
                End If


                If DdlEmpresaClienteOrigem.Text <> "" Then
                    cliente = DdlEmpresaClienteOrigem.SelectedValue
                    campo = cliente.Split("-")
                    Empresa = campo(0)                      'Empresa
                Else
                    Empresa = ""                            'Empresa
                End If

                If DdlFabrica.Text <> "" Then
                    cliente = DdlFabrica.SelectedValue
                    campo = cliente.Split("-")
                    Fabrica = campo(0)                      'Fabrica
                Else
                    Fabrica = ""                            'Fabrica
                End If

                Dim DataInicial As String = Format(CDate(txtPeriodoInicialConsultaTitulos.Text), "yyyy/MM/dd")
                Dim DataFinal As String = Format(CDate(txtPeriodoFinalConsultaTitulos.Text), "yyyy/MM/dd")

                sql = "   SELECT     Producao.Empresa_Id AS Empresa, Clientes.Cidade, Producao.Deposito_Id AS Deposito, Producao.Produto_Id AS Produto, Produtos.Nome AS NomeDoProduto, Producao.Operacao_Id as Operacao, Producao.SubOperacao_Id as SubOPeracao, SubOperacoes.Descricao,  " & vbCrLf & _
                      "             Producao.Movimento_Id AS Movimento, Producao.Entradas AS Producao, Producao.Saidas AS Consumo," & vbCrLf & _
                      "             ISNULL ((SELECT     Indice" & vbCrLf & _
                      "              FROM       ProducaoXAnalises" & vbCrLf & _
                      "              WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "             (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "             (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "             (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 1)), 0 ) " & vbCrLf & _
                      "              AS PercentualDeOleo," & vbCrLf & _
                      "             ISNULL ((SELECT Indice" & vbCrLf & _
                      "              FROM ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "              WHERE (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "             (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "             (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "             (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 2)), 0) AS Umidade," & vbCrLf & _
                      "             ISNULL ((SELECT Indice" & vbCrLf & _
                      "              FROM   ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "              WHERE (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "              (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "              (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "              (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 4)), 0) " & vbCrLf & _
                      "              AS ProteinaTotal," & vbCrLf & _
                      "              ISNULL ((SELECT Indice" & vbCrLf & _
                      "               FROM   ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "               WHERE  (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                             (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                             (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                             (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 7)), 0) " & vbCrLf & _
                      "                AS Impurezas," & vbCrLf & _
                      "                       ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                       FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                           WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                                  (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                                  (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                                 (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 11)), 0) AS Acidez," & vbCrLf & _
                      "                      ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                        FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                         WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                              (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                               (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                               (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 41)), 0) " & vbCrLf & _
                      "                   AS OleoLaboratorio," & vbCrLf & _
                      "                    ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                       FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                       WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                              (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                             (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                            (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 13)), 0) AS Saboes," & vbCrLf & _
                      "                   ISNULL ( (SELECT     Indice" & vbCrLf & _
                      "                      FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                      WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                              (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                               (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                               (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 34)), 0) " & vbCrLf & _
                      "                   AS Fosfatideos," & vbCrLf & _
                      "                     ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                        FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                       WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                            (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                          (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                             (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 35)), 0) AS Fosforo," & vbCrLf & _
                      "                     ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                        FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                         WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                               (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                               (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                             (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 37)), 0) " & vbCrLf & _
                      "                AS Clorofila," & vbCrLf & _
                      "                  ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                     FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                     WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                              (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                               (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                               (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 3)), 0) " & vbCrLf & _
                      "                   AS MateriaGraxa," & vbCrLf & _
                      "                    ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                         FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                         WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                           (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                            (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                              (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 5)), 0) " & vbCrLf & _
                      "                AS MateriaMineral," & vbCrLf & _
                      "                      ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                         FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                         WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                                 (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                              (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                               (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 10)), 0) " & vbCrLf & _
                      "                  AS AtividadeUreatica," & vbCrLf & _
                      "                     ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                        FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                        WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                             (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                              (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                                (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 28)), 0) " & vbCrLf & _
                      "                  AS PercentualDeFarelo," & vbCrLf & _
                      "                    ISNULL ((SELECT     Indice" & vbCrLf & _
                      "                       FROM          ProducaoXAnalises AS ProducaoXAnalises_1" & vbCrLf & _
                      "                      WHERE      (Producao.Empresa_Id = Empresa_Id) AND (Producao.EndEmpresa_Id = EndEmpresa_Id) AND (Producao.Deposito_Id = Deposito_Id) AND " & vbCrLf & _
                      "                                                 (Producao.EndDeposito_Id = EndDeposito_Id) AND (Producao.Produto_Id = Produto_Id) AND (Producao.Operacao_Id = Operacao_Id) AND " & vbCrLf & _
                      "                                                 (Producao.SubOperacao_Id = SubOperacao_Id) AND (Producao.Movimento_Id = Movimento_Id) AND " & vbCrLf & _
                      "                                                  (Producao.FisicoFiscal_Id = FisicoFiscal_Id) AND (Producao.ProdutoDerivado_Id = ProdutoDerivado_Id) AND (Analise_Id = 32)), 0) AS Profat " & vbCrLf & _
                      " FROM  Producao INNER JOIN" & vbCrLf & _
                      "       SubOperacoes ON Producao.Operacao_Id = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                      "       Producao.SubOperacao_Id = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
                      "       Produtos ON Producao.Produto_Id = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
                      "       Clientes ON Producao.Empresa_Id = Clientes.Cliente_Id AND Producao.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf

                sql &= " WHERE   (Producao.Empresa_Id = '" & Empresa & "') " & vbCrLf & _
                       " AND     (Producao.Operacao_Id in (40, 90)) " & vbCrLf & _
                       " AND     (Producao.SubOperacao_Id = 50) " & vbCrLf

                If Grupo <> "" Then
                    sql &= " AND     (Produtos.Grupo = '" & Grupo & "') "
                End If

                If Produto <> "" Then
                    sql &= " AND     (Producao.Produto_Id = '" & Produto & "') "
                End If

                sql &= " AND     (Producao.Deposito_Id = '" & Fabrica & "') AND (Producao.FisicoFiscal_Id = 2) " & vbCrLf & _
                       " AND     ((Producao.Entradas <> 0) or (Producao.Saidas <> 0))" & vbCrLf

                ds = Banco.ConsultaDataSet(sql, "Producao")

                linha = "<HTML>" & vbCrLf
                '<HEAD>
                linha &= "<HEAD>" & vbCrLf & _
                         "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf & _
                         "<TITLE>Analises de Producao</TITLE>" & vbCrLf & _
                         "</HEAD>" & vbCrLf

                '<BODY>
                linha &= "<BODY>" & vbCrLf

                '-----------------
                'Cabeçalho Padrao
                '-----------------
                linha &= "<table width= '3000' cellpadding='0' cellspacing='0' Border=1>"

                linha &= "<TR>" & vbCrLf & _
                         "<TD>Empresa</TD>" & vbCrLf & _
                         "<TD>Local</TD>" & vbCrLf & _
                         "<TD>Deposito</TD>" & vbCrLf & _
                         "<TD>Produto</TD>" & vbCrLf & _
                         "<TD>NomeDoProduto</TD>" & vbCrLf

                linha &= "<TD >Operacao</TD>"

                linha &= "<TD>SubOperacao</TD>" & vbCrLf & _
                         "<TD>NomeDaOperacao</TD>" & vbCrLf & _
                         "<TD>Data</TD>" & vbCrLf & _
                         "<TD>QuantidadeProducao</TD>" & vbCrLf

                linha &= "<TD>QuantidadeConsumo</TD>" & vbCrLf & _
                         "<TD>PercentualDeOleo</TD>" & vbCrLf & _
                         "<TD>Umidade</TD>" & vbCrLf & _
                         "<TD>ProteinaTotal</TD>" & vbCrLf & _
                         "<TD>Impurezas</TD>" & vbCrLf

                linha &= "<TD >Acidez</TD>" & vbCrLf & _
                         "<TD >OleoLaboratorio</TD>" & vbCrLf & _
                         "<TD >Saboes</TD>" & vbCrLf & _
                         "<TD >Fosfatideos</TD>" & vbCrLf & _
                         "<TD >Fosforo</TD>" & vbCrLf & _
                         "<TD >Clorofila</TD>" & vbCrLf & _
                         "<TD>MateriaGraxa</TD>" & vbCrLf & _
                         "<TD>MateriaMineral</TD>" & vbCrLf

                linha &= "<TD>AtividadeUreatica</TD>" & vbCrLf & _
                         "<TD>PercentualDeFarelo</TD>" & vbCrLf & _
                         "<TD>Profat</TD>" & vbCrLf & _
                         "</TR>" & vbCrLf

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        linha &= "<TR><TD>" & dr("Empresa") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Cidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Deposito") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Produto") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("NomeDoProduto") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Operacao") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("SubOperacao") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Descricao") & "</TD>" & vbCrLf

                        linha &= "<TD>" & dr("Movimento").ToStrDate() & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Producao") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Consumo") & "</TD>" & vbCrLf

                        linha &= "<TD>" & dr("PercentualDeOleo") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Umidade") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("ProteinaTotal") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Impurezas") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Acidez") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("OleoLaboratorio") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Saboes") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Fosfatideos") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Fosforo") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Clorofila") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("MateriaGraxa") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("MateriaMineral") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("AtividadeUreatica") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("PercentualDeFarelo") & "</TD>" & vbCrLf & _
                                 "<TD>" & dr("Profat") & "</TD>" & vbCrLf & _
                                 "</TR>" & vbCrLf

                    Next
                End If

                Try
                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                    ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
                Finally
                    strm.Close()
                End Try
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeAnaliseDeProducao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class