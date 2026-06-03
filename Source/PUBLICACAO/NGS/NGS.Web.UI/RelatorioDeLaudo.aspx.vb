Imports NGS.Lib.Negocio
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class RelatorioDeLaudo
    Inherits BasePage

#Region "Methods"

    Private Sub BuscaUnidade()
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa, True)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
    End Sub

    Private Sub BuscarClientes()
        ddl.Carregar(ddlCliente, CarregarDDL.Tabela.Clientes, "", True)
    End Sub

    Private Function getDataSet() As DataSet
        Dim sql As String = ""

        sql = " SELECT Empresa.Reduzido + '-' + Left(Empresa.Nome, 5) + '-' + Empresa.Cidade + '-' + Empresa.Estado AS Empresa," & vbCrLf & _
              "   Clientes.Nome + '-' + Clientes.Cidade + '-' + Clientes.Estado + ' (' + Pesagem.Cliente  + '-' + convert(varchar, Pesagem.EndCliente) + ')' AS Cliente," & vbCrLf & _
              "   Depositante.Nome + '-' + Depositante.Cidade + '-' + Depositante.Estado + ' (' + Pesagem.Depositante  + '-' + convert(varchar, Pesagem.EndDepositante) + ')' AS Depositante," & vbCrLf & _
              "   Transportador.Nome + '-' + Transportador.Cidade + '-' + Transportador.Estado + ' (' + Pesagem.Transportador  + '-' + convert(varchar, Pesagem.EndTransportador) + ')' AS Transportador," & vbCrLf & _
              "   Deposito.Nome + '-' + Deposito.Cidade + '-' + Deposito.Estado + ' (' + Pesagem.Deposito  + '-' + convert(varchar, Pesagem.EndDeposito) + ')' AS Deposito," & vbCrLf & _
              "   Pesagem.Produto  + '-' + Produtos.Nome AS Produto," & vbCrLf & _
              "   convert(varchar, Pesagem.Operacao)  + '-' +  convert(varchar, Pesagem.SubOperacao)  + '-' +  SubOperacoes.Descricao AS Operacao," & vbCrLf & _
              "                       Pesagem.Pesagem_Id as Laudo, Pesagem.Pedido, Pesagem.Movimento, " & vbCrLf & _
              "                       Pesagem.Placa, Pesagem.EntradaSaida as ES, Pesagem.PrimeiraPesagem, Pesagem.SegundaPesagem, Pesagem.BrutoBalanca, Pesagem.Liquido," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN PesagemXAnalises.Percentual ELSE 0 END) AS UmidadePercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN PesagemXAnalises.Indice ELSE 0 END) AS UmidadeIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN PesagemXAnalises.Desconto ELSE 0 END) AS UmidadeDesconto," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Percentual ELSE 0 END) AS ImpurezaPercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Indice ELSE 0 END) AS ImpurezaIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Desconto ELSE 0 END) AS ImpurezaDesconto," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Percentual ELSE 0 END) AS AvariadoPercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Indice ELSE 0 END) AS AvariadoIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Desconto ELSE 0 END) AS AvariadoDesconto," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Percentual ELSE 0 END) AS EsverdeadoPercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Indice ELSE 0 END) AS EsverdeadoIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Desconto ELSE 0 END) AS EsverdeadoDesconto," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Percentual ELSE 0 END) AS QuebradosPercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Indice ELSE 0 END) AS QuebradosIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Desconto ELSE 0 END) AS QuebradosDesconto," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 6 THEN PesagemXAnalises.Percentual ELSE 0 END) AS GmoPercentual," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 6 THEN PesagemXAnalises.Indice ELSE 0 END) AS GmoIndice," & vbCrLf & _
              "                       SUM(CASE WHEN PesagemXAnalises.Analise_ID = 6 THEN PesagemXAnalises.Desconto ELSE 0 END) AS GmoDesconto, Pesagem.UsuarioInclusao as Usuario" & vbCrLf & _
              "  FROM         Pesagem INNER JOIN" & vbCrLf & _
              "                    Clientes ON Pesagem.Cliente = Clientes.Cliente_Id AND Pesagem.EndCliente = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    Clientes AS Empresa ON Pesagem.Empresa_Id = Empresa.Cliente_Id AND Pesagem.EndEmpresa_Id = Empresa.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    Clientes AS Transportador ON Pesagem.Transportador = Transportador.Cliente_Id AND" & vbCrLf & _
              "                    Pesagem.EndTransportador = Transportador.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    PesagemXAnalises ON Pesagem.Empresa_Id = PesagemXAnalises.Empresa_Id AND" & vbCrLf & _
              "                    Pesagem.EndEmpresa_Id = PesagemXAnalises.EndEmpresa_Id AND Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
              "                    Clientes AS Depositante ON Pesagem.Depositante = Depositante.Cliente_Id AND Pesagem.EndDepositante = Depositante.Endereco_Id" & vbCrLf & _
              " WHERE             (Pesagem.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "')" & vbCrLf & _
              "                   And (Pesagem.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf & _
              "                   And (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue) Then
            sql &= " And Pesagem.Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
        End If

        sql &= " Group By          Empresa.Reduzido, Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id," & vbCrLf & _
               "                   Empresa.Nome, Empresa.Cidade, Empresa.Estado, Pesagem.Cliente, Pesagem.EndCliente," & vbCrLf & _
               "                   Clientes.Nome, Clientes.Cidade, Clientes.Estado, Pesagem.Depositante, Pesagem.EndDepositante," & vbCrLf & _
               "                   Depositante.Nome, Depositante.Cidade, Depositante.Estado, Pesagem.Transportador," & vbCrLf & _
               "                   Pesagem.EndTransportador, Transportador.Nome, Transportador.Cidade," & vbCrLf & _
               "                   Transportador.Estado, Pesagem.Deposito, Pesagem.EndDeposito, Deposito.Nome," & vbCrLf & _
               "                   Deposito.Cidade, Deposito.Estado, Pesagem.Produto, Produtos.Nome, Pesagem.Operacao," & vbCrLf & _
               "                   Pesagem.SubOperacao, SubOperacoes.Descricao, Pesagem.Pesagem_Id, Pesagem.Pedido,  " & vbCrLf & _
               "                   Pesagem.Placa, Pesagem.EntradaSaida, Pesagem.PrimeiraPesagem, Pesagem.SegundaPesagem, Pesagem.BrutoBalanca, Pesagem.Liquido," & vbCrLf & _
               "                   Pesagem.Movimento, Pesagem.UsuarioInclusao" & vbCrLf & _
               "" & vbCrLf

        If Not chkConsolidarCliente.Checked Then

            sql &= " SELECT Produto + '-' + NomeProduto as Produto, NomeCliente as Cliente, convert(varchar, Operacao) + '-' + convert(varchar, SubOperacao) + '-' + NomeOperacao as Operacao, SUM(BrutoBalanca) AS BrutoBalanca, SUM(Liquido) AS Liquido," & vbCrLf & _
                   "           SUM(UmidadePonderada) AS UmidadePonderada, SUM(UmidadeDesconto) AS UmidadeDesconto, SUM(ImpurezaPonderada) AS ImpurezaPonderada," & vbCrLf & _
                   "           SUM(ImpurezaDesconto) AS ImpurezaDesconto, SUM(AvariadoPonderada) AS AvariadoPonderada, SUM(AvariadoDesconto) AS AvariadoDesconto," & vbCrLf & _
                   "           SUM(EsverdeadoPonderada) AS EsverdeadoPonderada, SUM(EsverdeadoDesconto) AS EsverdeadoDesconto, SUM(QuebradosPonderada)" & vbCrLf & _
                   "           AS QuebradoPonderada, SUM(QuebradosDesconto) AS QuebradoDesconto" & vbCrLf & _
                   " FROM  (SELECT Clientes.Nome AS NomeCliente, Pesagem.Produto, Produtos.Nome AS NomeProduto, Pesagem.Operacao, Pesagem.SubOperacao," & vbCrLf & _
                   "                          SubOperacoes.Descricao AS NomeOperacao, Pesagem.BrutoBalanca, Pesagem.Liquido," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN PesagemXAnalises.Desconto ELSE 0 END) AS UmidadeDesconto," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS UmidadePonderada," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Desconto ELSE 0 END) AS ImpurezaDesconto," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS ImpurezaPonderada," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Desconto ELSE 0 END) AS AvariadoDesconto," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS AvariadoPonderada," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Desconto ELSE 0 END) AS EsverdeadoDesconto," & vbCrLf & _
                   "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                   "                          AS EsverdeadoPonderada, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                   "                          AS QuebradosDesconto, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                   "                          AS QuebradosPonderada" & vbCrLf & _
                   "           FROM   Pesagem INNER JOIN" & vbCrLf & _
                   "                          Clientes ON Pesagem.Cliente = Clientes.Cliente_Id AND Pesagem.EndCliente = Clientes.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          Clientes AS Empresa ON Pesagem.Empresa_Id = Empresa.Cliente_Id AND Pesagem.EndEmpresa_Id = Empresa.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          PesagemXAnalises ON Pesagem.Empresa_Id = PesagemXAnalises.Empresa_Id AND Pesagem.EndEmpresa_Id = PesagemXAnalises.EndEmpresa_Id AND" & vbCrLf & _
                   "                          Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                   "                          Clientes AS Depositante ON Pesagem.Depositante = Depositante.Cliente_Id AND Pesagem.EndDepositante = Depositante.Endereco_Id" & vbCrLf & _
                   " WHERE             (Pesagem.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "')" & vbCrLf & _
                   "                   And (Pesagem.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf & _
                   "                   And (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0 AND (Pesagem.EntradaSaida = 'E')" & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue) Then
                sql &= " And Pesagem.Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'"
            End If

            sql &= " GROUP BY Clientes.Nome, Pesagem.Produto, Produtos.Nome, Pesagem.Operacao, Pesagem.SubOperacao, SubOperacoes.Descricao, Pesagem.BrutoBalanca," & vbCrLf & _
                   "                           Pesagem.Liquido) AS Consulta" & vbCrLf & _
                   " WHERE (UmidadePonderada <> 0)" & vbCrLf & _
                   " GROUP BY NomeCliente, Produto, NomeProduto, Operacao, SubOperacao, NomeOperacao" & vbCrLf & _
                   " ORDER BY Produto, NomeCliente, Operacao, SubOperacao" & vbCrLf
        Else
            sql &= " SELECT Consulta.Produto + '-' + Consulta.NomeProduto AS Produto, Clientes.Nome as Cliente , CONVERT(varchar, Consulta.Operacao) + '-' + CONVERT(varchar," & vbCrLf & _
                  "            Consulta.SubOperacao) + '-' + Consulta.NomeOperacao AS Operacao, SUM(Consulta.BrutoBalanca) AS BrutoBalanca, SUM(Consulta.Liquido) AS Liquido," & vbCrLf & _
                  "            SUM(Consulta.UmidadePonderada) AS UmidadePonderada, SUM(Consulta.UmidadeDesconto) AS UmidadeDesconto, SUM(Consulta.ImpurezaPonderada)" & vbCrLf & _
                  "           AS ImpurezaPonderada, SUM(Consulta.ImpurezaDesconto) AS ImpurezaDesconto, SUM(Consulta.AvariadoPonderada) AS AvariadoPonderada," & vbCrLf & _
                  "            SUM(Consulta.AvariadoDesconto) AS AvariadoDesconto, SUM(Consulta.EsverdeadoPonderada) AS EsverdeadoPonderada, SUM(Consulta.EsverdeadoDesconto)" & vbCrLf & _
                  "           AS EsverdeadoDesconto, SUM(Consulta.QuebradosPonderada) AS QuebradoPonderada, SUM(Consulta.QuebradosDesconto) AS QuebradoDesconto," & vbCrLf & _
                  "          Clientes.Nome" & vbCrLf & _
                  " FROM  (SELECT Pesagem.Produto, Produtos.Nome AS NomeProduto," & vbCrLf & _
                  "                           CASE WHEN Pesagem.Operacao = 55 THEN Pesagem.Cliente ELSE Pesagem.Empresa_Id END AS Cliente," & vbCrLf & _
                  "                           CASE WHEN Pesagem.Operacao = 55 THEN Pesagem.EndCliente ELSE Pesagem.EndEmpresa_Id END AS EndCliente, Pesagem.Operacao," & vbCrLf & _
                  "                           Pesagem.SubOperacao, SubOperacoes.Descricao AS NomeOperacao, Pesagem.BrutoBalanca, Pesagem.Liquido," & vbCrLf & _
                  "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN PesagemXAnalises.Desconto ELSE 0 END) AS UmidadeDesconto," & vbCrLf & _
                  "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 1 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS UmidadePonderada," & vbCrLf & _
                  "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Desconto ELSE 0 END) AS ImpurezaDesconto," & vbCrLf & _
                  "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS ImpurezaPonderada," & vbCrLf & _
                  "                          SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Desconto ELSE 0 END) AS AvariadoDesconto," & vbCrLf & _
                  "                         SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS AvariadoPonderada," & vbCrLf & _
                  "                         SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Desconto ELSE 0 END) AS EsverdeadoDesconto," & vbCrLf & _
                  "                        SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                  "                         AS EsverdeadoPonderada, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                  "                         AS QuebradosDesconto, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                  "                           AS QuebradosPonderada" & vbCrLf & _
                  "            FROM   Pesagem LEFT OUTER JOIN" & vbCrLf & _
                  "                          PesagemXAnalises ON Pesagem.Empresa_Id = PesagemXAnalises.Empresa_Id AND Pesagem.EndEmpresa_Id = PesagemXAnalises.EndEmpresa_Id AND" & vbCrLf & _
                  "                          Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id LEFT OUTER JOIN" & vbCrLf & _
                  "                           SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
                  "                           Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
                  "                          Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id" & vbCrLf & _
                  " WHERE             (Pesagem.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "')" & vbCrLf & _
                  "                   And (Pesagem.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf & _
                  "                   And (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0 AND (Pesagem.EntradaSaida = 'E')" & vbCrLf

            If Not String.IsNullOrWhiteSpace(ddlCliente.SelectedValue) Then
                sql &= " And Pesagem.Cliente = '" & ddlCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf
            End If

            sql &= "           GROUP BY Pesagem.Produto, Produtos.Nome, Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Pesagem.Cliente, Pesagem.EndCliente, Pesagem.Operacao," & vbCrLf & _
                   "                          Pesagem.SubOperacao, SubOperacoes.Descricao, Pesagem.BrutoBalanca, Pesagem.Liquido) AS Consulta INNER JOIN" & vbCrLf & _
                   "          Clientes ON Consulta.Cliente = Clientes.Cliente_Id AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                   " WHERE (Consulta.UmidadePonderada <> 0)" & vbCrLf & _
                   " GROUP BY Consulta.Produto, Consulta.NomeProduto, Clientes.Nome, Consulta.Operacao, Consulta.SubOperacao, Consulta.NomeOperacao" & vbCrLf & _
                   " ORDER BY Produto, Clientes.Nome, Operacao, Consulta.SubOperacao" & vbCrLf
        End If

        Return Banco.ConsultaDataSet(sql, "Laudo")
    End Function

    Private Sub Limpar()
        ddlCliente.SelectedValue = ""
        txtDataInicial.Text = New DateTime(Now.Year, Now.Month, 1)
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gerencial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioDeLaudo", "ACESSAR") Then
                    BuscaUnidade()
                    BuscarClientes()
                    txtDataInicial.Text = New DateTime(Now.Year, Now.Month, 1)
                    txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gerencial.aspx")
                    Exit Sub
                End If
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

    Protected Sub DdlUnidade_SelectedIndexChanged1(sender As Object, e As EventArgs) Handles DdlUnidade.SelectedIndexChanged
        Try
            BuscaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim ds As DataSet = getDataSet()

            BindExcel(Me.Page, ds, "Laudo de pesagem.")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BindExcel(ByVal page As Page, ByVal ds As DataSet, ByVal TituloAba As String, Optional ByVal colunas As Dictionary(Of String, eTipoCampo) = Nothing)
        Try
            If ds Is Nothing OrElse ds.Tables(0) Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                Throw New Exception("Nenhum resultado encontrado.")
            Else
                Try
                    Dim rowIndex As Integer = 1
                    Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then File.Delete(fileName)

                    'emitir excel.xsls do office / relatório padrão em lista
                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)
                            'criando aba da planilha

                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add(TituloAba)
                            Dim columnIndex As Integer = 1

                            'worksheet.Cells(rowIndex, columnIndex).Value = ds.Tables(0).TableName
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            worksheet.Cells(rowIndex, columnIndex).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells(rowIndex, columnIndex).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.White)
                            rowIndex += 1

                            'inserindo o cabeçalho
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            Next

                            'aplicando formatação nas células do cabeçalho
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            rowIndex += 1

                            'exportando conteúdo da planilha com os dados da tabela
                            For Each row As DataRow In ds.Tables(0).Rows
                                columnIndex = 1
                                For Each col As DataColumn In ds.Tables(0).Columns
                                    If IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Trim.Contains(",") Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                        worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If

                                    'formatações de valores
                                    If colunas IsNot Nothing Then
                                        For Each coluna In colunas
                                            If coluna.Key = col.ColumnName Then
                                                If coluna.Value = eTipoCampo.Numerico Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.ValorComTotalizador OrElse coluna.Value = eTipoCampo.ValorSemTotalizador Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                    Exit For
                                                ElseIf coluna.Value = eTipoCampo.Data Then
                                                    worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                                    Exit For
                                                End If
                                            End If
                                        Next
                                    Else
                                        If IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                        ElseIf IsDate(row(col.ColumnName)) Then
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "dd/MM/yyyy"
                                        End If
                                    End If
                                    columnIndex += 1
                                Next

                                'formatações de celulas
                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                rowIndex += 1
                            Next

                            'aplicando formatação nas células do rodapé
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            columnIndex = 1

                            'soma dos campos de valores
                            If colunas IsNot Nothing Then

                                For Each col In colunas
                                    If col.Value = eTipoCampo.ValorComTotalizador Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = "TOTAL"
                                        Exit For
                                    End If
                                Next

                                For Each col As DataColumn In ds.Tables(0).Columns
                                    For Each coluna In colunas
                                        If coluna.Key = col.ColumnName Then
                                            If coluna.Value = eTipoCampo.ValorComTotalizador Then
                                                worksheet.Cells(rowIndex, columnIndex).Formula = "SUM(" & worksheet.Cells(1, columnIndex).Address & ":" & worksheet.Cells(rowIndex - 1, columnIndex).Address & ")"
                                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                                Exit For
                                            End If
                                        End If
                                    Next
                                    columnIndex += 1
                                Next
                            End If

                            '_____________Rodapé custumer__________________________________________________________________________________________________________________________________________
                            'Dim Empresa As String
                            'Dim Deposito As String
                            Dim Produto As String = ""
                            Dim Cliente As String = ""

                            Dim BrutoBalanca As Double
                            Dim Liquido As Double
                            Dim UmidadePonderada As Double
                            Dim UmidadeDesconto As Double
                            Dim ImpurezaPonderada As Double
                            Dim ImpurezaDesconto As Double
                            Dim AvariadoPonderada As Double
                            Dim AvariadoDesconto As Double
                            Dim EsverdeadoPonderada As Double
                            Dim EsverdeadoDesconto As Double
                            Dim QuebradoPonderada As Double
                            Dim QuebradoDesconto As Double

                            Dim TPBrutoBalanca As Double
                            Dim TPLiquido As Double
                            Dim TPUmidadePonderada As Double
                            Dim TPUmidadeDesconto As Double
                            Dim TPImpurezaPonderada As Double
                            Dim TPImpurezaDesconto As Double
                            Dim TPAvariadoPonderada As Double
                            Dim TPAvariadoDesconto As Double
                            Dim TPEsverdeadoPonderada As Double
                            Dim TPEsverdeadoDesconto As Double
                            Dim TPQuebradoPonderada As Double
                            Dim TPQuebradoDesconto As Double

                            Dim TGBrutoBalanca As Double
                            Dim TGLiquido As Double
                            Dim TGUmidadePonderada As Double
                            Dim TGUmidadeDesconto As Double
                            Dim TGImpurezaPonderada As Double
                            Dim TGImpurezaDesconto As Double
                            Dim TGAvariadoPonderada As Double
                            Dim TGAvariadoDesconto As Double
                            Dim TGEsverdeadoPonderada As Double
                            Dim TGEsverdeadoDesconto As Double
                            Dim TGQuebradoPonderada As Double
                            Dim TGQuebradoDesconto As Double

                            rowIndex += 2
                            columnIndex = 1

                            'For Each col As DataColumn In ds.Tables(1).Columns
                            '    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            '    columnIndex += 1
                            'Next

                            worksheet.Cells("A" & rowIndex).Value = "Produto"
                            worksheet.Cells("B" & rowIndex).Value = "Cliente"
                            worksheet.Cells("C" & rowIndex).Value = "Operacao"
                            worksheet.Cells("D" & rowIndex).Value = "PesoBruto"
                            worksheet.Cells("E" & rowIndex).Value = "PerUmidade"
                            worksheet.Cells("F" & rowIndex).Value = "DesUmidade"
                            worksheet.Cells("G" & rowIndex).Value = "PerImpureza"
                            worksheet.Cells("H" & rowIndex).Value = "DesImpureza"
                            worksheet.Cells("I" & rowIndex).Value = "PerAvariado"
                            worksheet.Cells("J" & rowIndex).Value = "DesAvariado"
                            worksheet.Cells("K" & rowIndex).Value = "PerEsverdeado"
                            worksheet.Cells("L" & rowIndex).Value = "DesEsveredeado"
                            worksheet.Cells("M" & rowIndex).Value = "PerQuebrado"
                            worksheet.Cells("N" & rowIndex).Value = "DesQuebrado"
                            worksheet.Cells("O" & rowIndex).Value = "Liquido"

                            worksheet.Cells("A" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True

                            rowIndex += 1

                            For Each row As DataRow In ds.Tables(1).Rows

                                If row("Produto") <> Produto And Produto = "" Then
                                    worksheet.Cells("A" & rowIndex).Value = row("Produto")
                                    worksheet.Cells("A" & rowIndex).Style.Font.Bold = True
                                    Produto = row("Produto")
                                    rowIndex += 1
                                End If

                                If row("Produto") <> Produto Then
                                    worksheet.Cells("C" & rowIndex).Value = "Sub Total do Cliente.."
                                    worksheet.Cells("D" & rowIndex).Value = BrutoBalanca

                                    If UmidadePonderada > 0 Then
                                        worksheet.Cells("E" & rowIndex).Value = UmidadePonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("F" & rowIndex).Value = UmidadeDesconto

                                    If ImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & rowIndex).Value = ImpurezaPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("H" & rowIndex).Value = ImpurezaDesconto

                                    If AvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & rowIndex).Value = AvariadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("J" & rowIndex).Value = AvariadoDesconto

                                    If EsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & rowIndex).Value = EsverdeadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("L" & rowIndex).Value = EsverdeadoDesconto

                                    If QuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & rowIndex).Value = QuebradoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("N" & rowIndex).Value = QuebradoDesconto

                                    worksheet.Cells("O" & rowIndex).Value = Liquido

                                    worksheet.Cells("" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                                    Cliente = row("Cliente")

                                    '------Totaliza Produto

                                    rowIndex += 1

                                    worksheet.Cells("C" & rowIndex).Value = "Sub Total do Produto.."
                                    worksheet.Cells("D" & rowIndex).Value = TPBrutoBalanca
                                    If TPUmidadePonderada > 0 Then
                                        worksheet.Cells("E" & rowIndex).Value = TPUmidadePonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("F" & rowIndex).Value = TPUmidadeDesconto

                                    If TPImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & rowIndex).Value = TPImpurezaPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("H" & rowIndex).Value = TPImpurezaDesconto

                                    If TPAvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & rowIndex).Value = TPAvariadoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("J" & rowIndex).Value = TPAvariadoDesconto

                                    If TPEsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & rowIndex).Value = TPEsverdeadoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("L" & rowIndex).Value = TPEsverdeadoDesconto

                                    If TPQuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & rowIndex).Value = TPQuebradoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("N" & rowIndex).Value = TPQuebradoDesconto

                                    worksheet.Cells("O" & rowIndex).Value = TPLiquido

                                    worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True

                                    BrutoBalanca = 0
                                    UmidadePonderada = 0
                                    UmidadeDesconto = 0
                                    ImpurezaPonderada = 0
                                    ImpurezaDesconto = 0
                                    AvariadoPonderada = 0
                                    AvariadoDesconto = 0
                                    EsverdeadoPonderada = 0
                                    EsverdeadoDesconto = 0
                                    EsverdeadoPonderada = 0
                                    EsverdeadoDesconto = 0
                                    Liquido = 0

                                    TPBrutoBalanca = 0
                                    TPUmidadePonderada = 0
                                    TPUmidadeDesconto = 0
                                    TPImpurezaPonderada = 0
                                    TPImpurezaDesconto = 0
                                    TPAvariadoPonderada = 0
                                    TPAvariadoDesconto = 0
                                    TPEsverdeadoPonderada = 0
                                    TPEsverdeadoDesconto = 0
                                    TPEsverdeadoPonderada = 0
                                    TPEsverdeadoDesconto = 0
                                    TPLiquido = 0

                                    rowIndex += 2

                                    worksheet.Cells("A" & rowIndex).Value = row("Produto")
                                    Produto = row("Produto")
                                    worksheet.Cells("A" & rowIndex).Style.Font.Bold = True
                                    rowIndex += 1

                                End If

                                If Cliente = "" Then
                                    Cliente = row("Cliente")
                                End If

                                If row("Cliente") <> Cliente Then
                                    worksheet.Cells("C" & rowIndex).Value = "Sub Total do Cliente.."
                                    worksheet.Cells("D" & rowIndex).Value = BrutoBalanca
                                    If UmidadePonderada > 0 Then
                                        worksheet.Cells("E" & rowIndex).Value = UmidadePonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("F" & rowIndex).Value = UmidadeDesconto

                                    If ImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & rowIndex).Value = ImpurezaPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("H" & rowIndex).Value = ImpurezaDesconto

                                    If AvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & rowIndex).Value = AvariadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("J" & rowIndex).Value = AvariadoDesconto

                                    If EsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & rowIndex).Value = EsverdeadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("L" & rowIndex).Value = EsverdeadoDesconto

                                    If QuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & rowIndex).Value = QuebradoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("N" & rowIndex).Value = QuebradoDesconto

                                    worksheet.Cells("O" & rowIndex).Value = Liquido

                                    worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                    worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                    worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True

                                    BrutoBalanca = 0
                                    UmidadePonderada = 0
                                    UmidadeDesconto = 0
                                    ImpurezaPonderada = 0
                                    ImpurezaDesconto = 0
                                    AvariadoPonderada = 0
                                    AvariadoDesconto = 0
                                    EsverdeadoPonderada = 0
                                    EsverdeadoDesconto = 0
                                    EsverdeadoPonderada = 0
                                    EsverdeadoDesconto = 0
                                    Liquido = 0
                                    Cliente = row("Cliente")
                                    rowIndex += 2
                                End If

                                worksheet.Cells("B" & rowIndex).Value = row("Cliente")
                                worksheet.Cells("B" & rowIndex).Style.Font.Bold = True

                                worksheet.Cells("C" & rowIndex).Value = row("Operacao")
                                worksheet.Cells("D" & rowIndex).Value = row("BrutoBalanca")

                                If row("UmidadePonderada") > 0 Then
                                    worksheet.Cells("E" & rowIndex).Value = row("UmidadePonderada") / row("BrutoBalanca")
                                End If
                                worksheet.Cells("F" & rowIndex).Value = row("UmidadeDesconto")

                                If row("ImpurezaPonderada") > 0 Then
                                    worksheet.Cells("G" & rowIndex).Value = row("ImpurezaPonderada") / row("BrutoBalanca")
                                End If
                                worksheet.Cells("H" & rowIndex).Value = row("ImpurezaDesconto")

                                If row("AvariadoPonderada") Then
                                    worksheet.Cells("I" & rowIndex).Value = row("AvariadoPonderada") / row("BrutoBalanca")
                                End If
                                worksheet.Cells("J" & rowIndex).Value = row("AvariadoDesconto")

                                If row("EsverdeadoPonderada") > 0 Then
                                    worksheet.Cells("K" & rowIndex).Value = row("EsverdeadoPonderada") / row("BrutoBalanca")
                                End If
                                worksheet.Cells("L" & rowIndex).Value = row("EsverdeadoDesconto")

                                If row("QuebradoPonderada") > 0 Then
                                    worksheet.Cells("M" & rowIndex).Value = row("QuebradoPonderada") / row("BrutoBalanca")
                                End If

                                worksheet.Cells("N" & rowIndex).Value = row("QuebradoDesconto")
                                worksheet.Cells("O" & rowIndex).Value = row("Liquido")

                                worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                                worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"

                                BrutoBalanca = BrutoBalanca + row("BrutoBalanca")
                                UmidadePonderada = UmidadePonderada + row("UmidadePonderada")
                                UmidadeDesconto = UmidadeDesconto + row("UmidadeDesconto")
                                ImpurezaPonderada = ImpurezaPonderada + row("ImpurezaPonderada")
                                ImpurezaDesconto = ImpurezaDesconto + row("ImpurezaDesconto")
                                AvariadoPonderada = AvariadoPonderada + row("AvariadoPonderada")
                                AvariadoDesconto = AvariadoDesconto + row("AvariadoDesconto")
                                EsverdeadoPonderada = EsverdeadoPonderada + row("EsverdeadoPonderada")
                                EsverdeadoDesconto = EsverdeadoDesconto + row("EsverdeadoDesconto")
                                EsverdeadoPonderada = EsverdeadoPonderada + row("EsverdeadoPonderada")
                                EsverdeadoDesconto = EsverdeadoDesconto + row("EsverdeadoDesconto")
                                Liquido = Liquido + row("Liquido")

                                TPBrutoBalanca = TPBrutoBalanca + row("BrutoBalanca")
                                TPUmidadePonderada = TPUmidadePonderada + row("UmidadePonderada")
                                TPUmidadeDesconto = TPUmidadeDesconto + row("UmidadeDesconto")
                                TPImpurezaPonderada = TPImpurezaPonderada + row("ImpurezaPonderada")
                                TPImpurezaDesconto = TPImpurezaDesconto + row("ImpurezaDesconto")
                                TPAvariadoPonderada = TPAvariadoPonderada + row("AvariadoPonderada")
                                TPAvariadoDesconto = TPAvariadoDesconto + row("AvariadoDesconto")
                                TPEsverdeadoPonderada = TPEsverdeadoPonderada + row("EsverdeadoPonderada")
                                TPEsverdeadoDesconto = TPEsverdeadoDesconto + row("EsverdeadoDesconto")
                                TPEsverdeadoPonderada = TPEsverdeadoPonderada + row("EsverdeadoPonderada")
                                TPEsverdeadoDesconto = TPEsverdeadoDesconto + row("EsverdeadoDesconto")
                                TPLiquido = TPLiquido + row("Liquido")

                                TGBrutoBalanca = TGBrutoBalanca + row("BrutoBalanca")
                                TGUmidadePonderada = TGUmidadePonderada + row("UmidadePonderada")
                                TGUmidadeDesconto = TGUmidadeDesconto + row("UmidadeDesconto")
                                TGImpurezaPonderada = TGImpurezaPonderada + row("ImpurezaPonderada")
                                TGImpurezaDesconto = TGImpurezaDesconto + row("ImpurezaDesconto")
                                TGAvariadoPonderada = TGAvariadoPonderada + row("AvariadoPonderada")
                                TGAvariadoDesconto = TGAvariadoDesconto + row("AvariadoDesconto")
                                TGEsverdeadoPonderada = TGEsverdeadoPonderada + row("EsverdeadoPonderada")
                                TGEsverdeadoDesconto = TGEsverdeadoDesconto + row("EsverdeadoDesconto")
                                TGEsverdeadoPonderada = TGEsverdeadoPonderada + row("EsverdeadoPonderada")
                                TGEsverdeadoDesconto = TGEsverdeadoDesconto + row("EsverdeadoDesconto")
                                TGLiquido = TGLiquido + row("Liquido")

                                rowIndex += 1
                            Next

                            worksheet.Cells("C" & rowIndex).Value = "Sub Total do Cliente.."
                            worksheet.Cells("D" & rowIndex).Value = BrutoBalanca
                            If UmidadePonderada > 0 Then
                                worksheet.Cells("E" & rowIndex).Value = UmidadePonderada / BrutoBalanca
                            End If
                            worksheet.Cells("F" & rowIndex).Value = UmidadeDesconto

                            If ImpurezaPonderada > 0 Then
                                worksheet.Cells("G" & rowIndex).Value = ImpurezaPonderada / BrutoBalanca
                            End If
                            worksheet.Cells("H" & rowIndex).Value = ImpurezaDesconto

                            If AvariadoPonderada > 0 Then
                                worksheet.Cells("I" & rowIndex).Value = AvariadoPonderada / BrutoBalanca
                            End If
                            worksheet.Cells("J" & rowIndex).Value = AvariadoDesconto

                            If EsverdeadoPonderada > 0 Then
                                worksheet.Cells("K" & rowIndex).Value = EsverdeadoPonderada / BrutoBalanca
                            End If
                            worksheet.Cells("L" & rowIndex).Value = EsverdeadoDesconto

                            If QuebradoPonderada > 0 Then
                                worksheet.Cells("M" & rowIndex).Value = QuebradoPonderada / BrutoBalanca
                            End If
                            worksheet.Cells("N" & rowIndex).Value = QuebradoDesconto
                            worksheet.Cells("O" & rowIndex).Value = Liquido

                            worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                            rowIndex += 1

                            '------Totaliza Produto

                            worksheet.Cells("C" & rowIndex).Value = "Sub Total do Produto.."
                            worksheet.Cells("D" & rowIndex).Value = TPBrutoBalanca
                            If TPUmidadePonderada > 0 Then
                                worksheet.Cells("E" & rowIndex).Value = TPUmidadePonderada / TPBrutoBalanca
                            End If
                            worksheet.Cells("F" & rowIndex).Value = TPUmidadeDesconto

                            If TPImpurezaPonderada > 0 Then
                                worksheet.Cells("G" & rowIndex).Value = TPImpurezaPonderada / TPBrutoBalanca
                            End If
                            worksheet.Cells("H" & rowIndex).Value = TPImpurezaDesconto

                            If TPAvariadoPonderada > 0 Then
                                worksheet.Cells("I" & rowIndex).Value = TPAvariadoPonderada / TPBrutoBalanca
                            End If
                            worksheet.Cells("J" & rowIndex).Value = TPAvariadoDesconto

                            If TPEsverdeadoPonderada > 0 Then
                                worksheet.Cells("K" & rowIndex).Value = TPEsverdeadoPonderada / TPBrutoBalanca
                            End If
                            worksheet.Cells("L" & rowIndex).Value = TPEsverdeadoDesconto

                            If TPQuebradoPonderada > 0 Then
                                worksheet.Cells("M" & rowIndex).Value = TPQuebradoPonderada / TPBrutoBalanca
                            End If
                            worksheet.Cells("N" & rowIndex).Value = TPQuebradoDesconto
                            worksheet.Cells("O" & rowIndex).Value = TPLiquido

                            worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True
                            rowIndex += 1

                            '-----Total Geral------------------------------------------------

                            worksheet.Cells("C" & rowIndex).Value = "Total Geral..."
                            worksheet.Cells("D" & rowIndex).Value = TGBrutoBalanca

                            If TGUmidadePonderada > 0 Then
                                worksheet.Cells("E" & rowIndex).Value = TGUmidadePonderada / TGBrutoBalanca
                            End If
                            worksheet.Cells("F" & rowIndex).Value = TGUmidadeDesconto

                            If TGImpurezaPonderada > 0 Then
                                worksheet.Cells("G" & rowIndex).Value = TGImpurezaPonderada / TGBrutoBalanca
                            End If
                            worksheet.Cells("H" & rowIndex).Value = TGImpurezaDesconto

                            If TGAvariadoPonderada > 0 Then
                                worksheet.Cells("I" & rowIndex).Value = TGAvariadoPonderada / TGBrutoBalanca
                            End If
                            worksheet.Cells("J" & rowIndex).Value = TGAvariadoDesconto

                            If TGEsverdeadoPonderada > 0 Then
                                worksheet.Cells("K" & rowIndex).Value = TGEsverdeadoPonderada / TGBrutoBalanca
                            End If
                            worksheet.Cells("L" & rowIndex).Value = TGEsverdeadoDesconto

                            If TGQuebradoPonderada > 0 Then
                                worksheet.Cells("M" & rowIndex).Value = TGQuebradoPonderada / TGBrutoBalanca
                            End If
                            worksheet.Cells("N" & rowIndex).Value = TGQuebradoDesconto

                            worksheet.Cells("O" & rowIndex).Value = TGLiquido

                            worksheet.Cells("D" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("E" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("F" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("G" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("H" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("I" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("J" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("K" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("L" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("M" & rowIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells("N" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("O" & rowIndex).Style.Numberformat.Format = "#,##0_ ;[Red]-#,##0"
                            worksheet.Cells("C" & rowIndex & ":O" & rowIndex).Style.Font.Bold = True

                            'criando auto filtro na planilha

                            worksheet.Cells(1, 1, 1, ds.Tables(0).Columns.Count).AutoFilter = True

                            'setando autofit nas células da planilha
                            worksheet.Cells.AutoFitColumns(0)

                            worksheet.Column(2).Width = 5
                            worksheet.Column(3).Width = 60
                            worksheet.Column(4).Width = 14
                            worksheet.Column(5).Width = 11
                            worksheet.Column(6).Width = 27
                            worksheet.Column(7).Width = 26

                            'congelando primeira linham
                            worksheet.View.FreezePanes(2, 1)

                            'salvando planilha do excel
                            package.Save()
                        End Using
                    End Using

                    'download do arquivo pelo browser
                    Funcoes.AbrirExcel(page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                Catch ex As Exception
                    Throw New Exception(ex.Message)
                End Try
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioDeLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class