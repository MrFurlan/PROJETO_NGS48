Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Drawing

Public Class Cardex
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Cardex", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaDepositos()
                carrega_oper()
                GruposdeEstoques()

                LimparCampos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, ddlEmpresa, True)
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue.ToString, True)
    End Sub

    Private Sub CargaDepositos()
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.Depositos, "", True)
    End Sub

    Private Sub GruposdeEstoques()
        ddl.Carregar(ddlGrupo, CarregarDDL.Tabela.GrupoProduto, "", True)
    End Sub

    Private Sub Produtos()
        ddl.Carregar(ddlProduto, CarregarDDL.Tabela.Produto, "Grupo = '" & ddlGrupo.SelectedValue & "'", True)
    End Sub

    Protected Sub lnkBuscaProduto_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkBuscaProduto.Click
        Try
            ucConsultaProduto.Limpar()
            Dim txtNome As TextBox = CType(ucConsultaProduto.FindControlRecursive("txtNome"), TextBox)
            Popup.ConsultaDeProduto(Me.Page, "objProdutoxCDX" & HID.Value, txtNome.ClientID, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub carrega_oper()
        ddl.Carregar(DdlOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub Carrega_sub()
        ddl.Carregar(DdlSuboperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "So.Operacao_Id = " & DdlOperacao.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Session("objProdutoxCDX" & HID.Value) IsNot Nothing Then
            Dim objProduto As [Lib].Negocio.Produto = CType(obj, [Lib].Negocio.Produto)

            ddlGrupo.SelectedValue = objProduto.CodigoGrupo

            Produtos()

            ddlProduto.SelectedValue = objProduto.Codigo

            Session.Remove("objProdutoxCDX" & HID.Value)
        End If
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            param &= "Unidade: " & ddlUnidade.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            param &= "Empresa: " & ddlEmpresa.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlDeposito.SelectedValue) Then
            param &= "Depósito: " & ddlDeposito.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlGrupo.SelectedValue) Then
            param &= "Grupo: " & ddlGrupo.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            param &= "Produto: " & ddlProduto.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlOperacao.SelectedValue) Then
            param &= "Operação: " & DdlOperacao.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(DdlSuboperacao.Text) Then
            param &= "S.Operacao: " & DdlSuboperacao.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
            param &= "Data: " & txtDataInicial.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            param &= "" & txtDataFinal.Text & " - "
        End If

        Return param
    End Function

    Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Informe a data inicial e a data final")
            Return False
        ElseIf CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a Data Final.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then
            MsgBox(Me.Page, "Informe o produto.")
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Carrega_sub()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Produtos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSet() As DataSet
        Dim strSQL As String = ""
        Dim strSQLEstoques As String = ""
        Dim strSQLNotas As String = ""

        strSQL = "SELECT convert(datetime,'" & txtDataInicial.Text.ToSqlDate() & "') - 1  AS Data, 0 AS Reduzido, '' AS Cliente_Id, 0 AS EndCliente, '' AS Nome, 0 AS Operacao, " & vbCrLf &
                 "0 AS SubOperacao, 'Saldo Anterior' AS Descricao, 0 AS Ordem, 0 AS Entradas, 0 AS Saidas, SUM(SaldoAnterior) AS Saldo " & vbCrLf &
                 "FROM (SELECT COALESCE(SUM(E.Entradas - E.Saidas), 0) AS SaldoAnterior " & vbCrLf

        strSQLEstoques = "FROM Producao E " & vbCrLf &
                         "INNER JOIN SubOperacoes SO " & vbCrLf &
                         "ON SO.Operacao_Id = E.Operacao_Id " & vbCrLf &
                         "AND SO.SubOperacoes_Id = E.SubOperacao_Id " & vbCrLf &
                         "INNER JOIN Produtos Prd " & vbCrLf &
                         "ON Prd.Produto_Id = E.Produto_Id " & vbCrLf &
                         "AND Prd.Situacao  = 1 " & vbCrLf &
                         "WHERE E.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                         "AND E.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            strSQLEstoques &= "AND E.Deposito_Id = '" & ddlDeposito.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                              "AND E.EndDeposito_Id = " & ddlDeposito.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        End If

        If ddlProduto.SelectedIndex > 0 Then
            strSQLEstoques &= "AND E.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
        End If

        If RadFiscal.Checked = True Then
            strSQLEstoques &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf &
                              "AND E.FisicoFiscal_Id = 2 " & vbCrLf
        End If

        If RadFisico.Checked = True Then
            strSQLEstoques &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            strSQLEstoques &= "AND E.FisicoFiscal_Id = 1 " & vbCrLf
        End If

        If DdlOperacao.Text <> Nothing Then
            strSQLEstoques &= "AND SO.Operacao_Id = '" & DdlOperacao.SelectedValue & "' " & vbCrLf
        End If

        If DdlSuboperacao.Text <> Nothing Then
            strSQLEstoques &= "AND SO.SubOperacoes_Id = '" & DdlSuboperacao.SelectedValue & "' " & vbCrLf
        End If

        If ckCusto.Checked Then
            strSQLEstoques &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) " & vbCrLf
        End If

        strSQL &= strSQLEstoques & "AND E.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        If chkOrdemDeProducaoInterna.Checked Then
            strSQL &= "SELECT" & vbCrLf &
                             "	SUM(CASE" & vbCrLf &
                             "		WHEN SO.EntradaSaida = 'E'" & vbCrLf &
                             "			THEN Quantidade" & vbCrLf &
                             "			ELSE Quantidade * -1" & vbCrLf &
                             "		END) AS SaldoAnterior" & vbCrLf &
                             "FROM OrdemDeProducaoInterna OPI " & vbCrLf &
                             "INNER JOIN SubOperacoes SO " & vbCrLf &
                             "ON SO.Operacao_Id      = OPI.Operacao " & vbCrLf &
                             "AND SO.SubOperacoes_Id = OPI.SubOperacao " & vbCrLf &
                             "INNER JOIN Produtos Prd " & vbCrLf &
                             "ON Prd.Produto_Id = OPI.Produto_Id " & vbCrLf &
                             "AND Prd.Situacao  = 1 " & vbCrLf &
                             "WHERE OPI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                             "AND OPI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf &
                             "AND OPI.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

            If ddlProduto.SelectedIndex > 0 Then
                strSQL &= "AND OPI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
            End If

            If RadFiscal.Checked = True Then
                strSQL &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
            End If

            If RadFisico.Checked = True Then
                strSQL &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            End If

            strSQL &= "UNION ALL " & vbCrLf
        End If

        strSQL &= "SELECT COALESCE(SUM(NFI.QuantidadeFiscal), 0) AS SaldoAnterior " & vbCrLf

        strSQLNotas = "FROM NotasFiscaisXItens NFI " & vbCrLf

        strSQLNotas &= "INNER JOIN NotasFiscais NF " & vbCrLf &
                       "ON  NFI.Empresa_Id = NF.Empresa_Id " & vbCrLf &
                       "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf &
                       "AND NFI.Cliente_Id = NF.Cliente_Id " & vbCrLf &
                       "AND NFI.EndCliente_Id = NF.EndCliente_Id " & vbCrLf &
                       "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                       "AND NFI.Serie_Id = NF.Serie_Id " & vbCrLf &
                       "AND NFI.Nota_Id = NF.Nota_Id " & vbCrLf

        strSQLNotas &= "INNER JOIN SubOperacoes SO " & vbCrLf &
                       "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf &
                       "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf

        strSQLNotas &= "INNER JOIN Produtos Prd " & vbCrLf &
                       "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf &
                       "AND Prd.Situacao  = 1 " & vbCrLf

        strSQLNotas &= "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf


        strSQLNotas &= "AND NFI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                       "AND NFI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            strSQLNotas &= "AND NFI.Deposito = '" & ddlDeposito.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                           "AND NFI.EndDeposito = " & ddlDeposito.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        End If

        If ddlProduto.SelectedIndex > 0 Then
            strSQLNotas &= " And NFI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
        End If

        If RadFiscal.Checked = True Then
            strSQLNotas &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
        End If

        If RadFisico.Checked = True Then
            strSQLNotas &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
        End If

        If DdlOperacao.Text <> Nothing Then
            strSQLNotas &= "AND SO.Operacao_Id = '" & DdlOperacao.SelectedValue & "' " & vbCrLf
        End If

        If DdlSuboperacao.Text <> Nothing Then
            strSQLNotas &= "AND SO.SubOperacoes_Id = '" & DdlSuboperacao.SelectedValue & "' " & vbCrLf
        End If

        If ckCusto.Checked Then
            strSQLNotas &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) " & vbCrLf
        End If

        strSQL &= strSQLNotas & "AND NFI.EntradaSaida_Id = 'E' AND NF.Situacao=1 " & vbCrLf &
                                "AND NF.Serie_id <> '101' AND NF.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        strSQL &= "SELECT COALESCE(SUM(NFI.QuantidadeFiscal) * (-1), 0) AS SaldoAnterior " & vbCrLf
        strSQL &= strSQLNotas & " AND NF.Serie_id <> '101' AND NF.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "AND NFI.EntradaSaida_Id = 'S' AND NF.Situacao=1 ) SA " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        strSQL &= "SELECT MOV.Data, C.Reduzido, MOV.Cliente_Id, MOV.EndCliente, C.Nome, MOV.Operacao, MOV.SubOperacao, SO.Descricao, " & vbCrLf &
                  "MOV.Ordem, " & vbCrLf &
                  "" & IIf(Not CkCDeposito.Checked, "MOV.Entradas ", "Sum(MOV.Entradas) as Entradas") & "," & IIf(Not CkCDeposito.Checked, "MOV.Saidas ", "Sum(MOV.Saidas) as Saidas") & ", 0 AS Saldo " & vbCrLf

        strSQL &= "FROM (SELECT E.Movimento_Id AS Data, E.Deposito_Id AS Cliente_Id , E.EndDeposito_Id AS EndCliente , 1 AS Ordem, " & vbCrLf &
                  "E.Operacao_Id AS Operacao, E.SubOperacao_Id AS SubOperacao, COALESCE(E.Entradas, 0) AS Entradas, COALESCE(E.Saidas, 0) AS Saidas " & vbCrLf

        strSQL &= strSQLEstoques & "AND E.Movimento_Id BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                                   "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102) " & vbCrLf

        '	ADICIONADO
        If DdlOperacao.Text.Length > 0 Then
            strSQL &= "AND E.Operacao_Id = " & DdlOperacao.SelectedValue & " " & vbCrLf
        End If
        If DdlSuboperacao.Text.Length > 0 Then
            strSQL &= "AND E.SubOperacao_Id = " & DdlSuboperacao.SelectedValue & "" & vbCrLf
        End If

        strSQL &= "UNION ALL " & vbCrLf

        If chkOrdemDeProducaoInterna.Checked Then
            strSQL &= "SELECT OPI.Movimento AS Data, OPI.Empresa_Id AS Cliente_Id , OPI.EndEmpresa_Id AS EndCliente, 1 AS Ordem, OPI.Operacao, OPI.SubOperacao, " & vbCrLf &
                        "CASE WHEN OPI.EntradaSaida = 'E' THEN COALESCE(OPI.Quantidade, 0) " & vbCrLf &
                        "ELSE 0 END AS Entradas, " & vbCrLf &
                        "CASE WHEN OPI.EntradaSaida = 'S' THEN COALESCE(OPI.Quantidade, 0) " & vbCrLf &
                        "ELSE 0 END AS Saidas " & vbCrLf &
                        "FROM OrdemDeProducaoInterna OPI " & vbCrLf &
                        "INNER JOIN SubOperacoes SO " & vbCrLf &
                        "ON SO.Operacao_Id      = OPI.Operacao " & vbCrLf &
                        "AND SO.SubOperacoes_Id = OPI.SubOperacao " & vbCrLf &
                        "INNER JOIN Produtos Prd " & vbCrLf &
                        "ON Prd.Produto_Id = OPI.Produto_Id " & vbCrLf &
                        "AND Prd.Situacao  = 1 " & vbCrLf &
                        "WHERE OPI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                        "AND OPI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf &
                        "AND OPI.Movimento BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                        "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102) " & vbCrLf

            If ddlProduto.SelectedIndex > 0 Then
                strSQL &= "AND OPI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
            End If

            If RadFiscal.Checked = True Then
                strSQL &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
            End If

            If RadFisico.Checked = True Then
                strSQL &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            End If

            strSQL &= "UNION ALL " & vbCrLf
        End If

        strSQL &= "SELECT NF.Movimento AS Data, NFI.Cliente_Id AS Cliente_Id ,NFI.EndCliente_Id AS EndCliente, 1 AS Ordem, NF.Operacao, NF.SubOperacao, " & vbCrLf &
                  "CASE WHEN NFI.EntradaSaida_Id = 'E' THEN COALESCE(NFI.QuantidadeFiscal, 0) " & vbCrLf &
                  "ELSE 0 END AS Entradas, " & vbCrLf &
                  "CASE WHEN NFI.EntradaSaida_Id = 'S' THEN COALESCE(NFI.QuantidadeFiscal, 0) " & vbCrLf &
                  "ELSE 0 END AS Saidas " & vbCrLf

        strSQL &= strSQLNotas & "AND NF.Movimento BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                                "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102)  AND NF.Situacao=1 " & vbCrLf

        strSQL &= " ) MOV " & vbCrLf

        strSQL &= "INNER JOIN Clientes C " & vbCrLf &
                  "ON MOV.Cliente_Id = C.Cliente_Id " & vbCrLf &
                  "AND MOV.EndCliente = C.Endereco_Id " & vbCrLf

        strSQL &= "INNER JOIN SubOperacoes SO " & vbCrLf &
                  "ON MOV.Operacao = SO.Operacao_Id " & vbCrLf &
                  "AND MOV.SubOperacao = SO.SubOperacoes_Id " & vbCrLf
        If CkCDeposito.Checked Then strSQL &= " group by  MOV.Data, C.Reduzido, MOV.Cliente_Id, MOV.EndCliente, C.Nome, MOV.Operacao, MOV.SubOperacao, SO.Descricao, MOV.Ordem " & vbCrLf

        strSQL &= "ORDER BY  Data, Saidas" & vbCrLf

        Return Banco.ConsultaDataSet(strSQL, "Cardex")
    End Function


    Private Function getDataSetXLS() As DataSet
        Dim strSQL As String = ""
        Dim strSQLEstoques As String = ""
        Dim strSQLNotasAnterior As String = ""
        Dim strSQLNotas As String = ""

        strSQL = "SELECT convert(datetime,'" & txtDataInicial.Text.ToSqlDate() & "') - 1  AS Data, 0 AS Reduzido, '' AS Cliente_Id, 0 AS EndCliente, '' AS Nome, 0 AS Operacao, " & vbCrLf &
                 "0 AS SubOperacao, 'Saldo Anterior' AS Descricao, 0 AS Ordem, 0 AS Entradas, 0 AS Saidas, SUM(SaldoAnterior) AS Saldo, '' AS Lote, 0 AS OrdemDeProducao " & vbCrLf &
                 "FROM (SELECT COALESCE(SUM(E.Entradas - E.Saidas), 0) AS SaldoAnterior " & vbCrLf

        strSQLEstoques = "FROM Producao E " & vbCrLf &
                         "INNER JOIN SubOperacoes SO " & vbCrLf &
                         "ON SO.Operacao_Id = E.Operacao_Id " & vbCrLf &
                         "AND SO.SubOperacoes_Id = E.SubOperacao_Id " & vbCrLf &
                         "INNER JOIN Produtos Prd " & vbCrLf &
                         "ON Prd.Produto_Id = E.Produto_Id " & vbCrLf &
                         "AND Prd.Situacao  = 1 " & vbCrLf &
                         "WHERE E.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                         "AND E.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            strSQLEstoques &= "AND E.Deposito_Id = '" & ddlDeposito.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                              "AND E.EndDeposito_Id = " & ddlDeposito.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        End If

        If ddlProduto.SelectedIndex > 0 Then
            strSQLEstoques &= "AND E.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
        End If

        If RadFiscal.Checked = True Then
            strSQLEstoques &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf &
                              "AND E.FisicoFiscal_Id = 2 " & vbCrLf
        End If

        If RadFisico.Checked = True Then
            strSQLEstoques &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            strSQLEstoques &= "AND E.FisicoFiscal_Id = 1 " & vbCrLf
        End If

        If DdlOperacao.Text <> Nothing Then
            strSQLEstoques &= "AND SO.Operacao_Id = '" & DdlOperacao.SelectedValue & "' " & vbCrLf
        End If

        If DdlSuboperacao.Text <> Nothing Then
            strSQLEstoques &= "AND SO.SubOperacoes_Id = '" & DdlSuboperacao.SelectedValue & "' " & vbCrLf
        End If

        If ckCusto.Checked Then
            strSQLEstoques &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) " & vbCrLf
        End If

        strSQL &= strSQLEstoques & "AND E.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        If chkOrdemDeProducaoInterna.Checked Then
            strSQL &= "SELECT" & vbCrLf &
                             "	SUM(CASE" & vbCrLf &
                             "		WHEN SO.EntradaSaida = 'E'" & vbCrLf &
                             "			THEN Quantidade" & vbCrLf &
                             "			ELSE Quantidade * -1" & vbCrLf &
                             "		END) AS SaldoAnterior" & vbCrLf &
                             "FROM OrdemDeProducaoInterna OPI " & vbCrLf &
                             "INNER JOIN SubOperacoes SO " & vbCrLf &
                             "ON SO.Operacao_Id      = OPI.Operacao " & vbCrLf &
                             "AND SO.SubOperacoes_Id = OPI.SubOperacao " & vbCrLf &
                             "INNER JOIN Produtos Prd " & vbCrLf &
                             "ON Prd.Produto_Id = OPI.Produto_Id " & vbCrLf &
                             "AND Prd.Situacao  = 1 " & vbCrLf &
                             "WHERE OPI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                             "AND OPI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf &
                             "AND OPI.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

            If ddlProduto.SelectedIndex > 0 Then
                strSQL &= "AND OPI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
            End If

            If RadFiscal.Checked = True Then
                strSQL &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
            End If

            If RadFisico.Checked = True Then
                strSQL &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            End If

            strSQL &= "UNION ALL " & vbCrLf
        End If

        strSQL &= "SELECT COALESCE(SUM(NFI.QuantidadeFiscal), 0) AS SaldoAnterior " & vbCrLf

        strSQLNotasAnterior = "FROM NotasFiscaisXItens NFI " & vbCrLf

        strSQLNotasAnterior &= "INNER JOIN NotasFiscais NF " & vbCrLf &
                               "ON  NFI.Empresa_Id = NF.Empresa_Id " & vbCrLf &
                               "AND NFI.EndEmpresa_Id = NF.EndEmpresa_Id " & vbCrLf &
                               "AND NFI.Cliente_Id = NF.Cliente_Id " & vbCrLf &
                               "AND NFI.EndCliente_Id = NF.EndCliente_Id " & vbCrLf &
                               "AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                               "AND NFI.Serie_Id = NF.Serie_Id " & vbCrLf &
                               "AND NFI.Nota_Id = NF.Nota_Id " & vbCrLf

        strSQLNotas = strSQLNotasAnterior
        strSQLNotas &= "LEFT JOIN NotaFiscalXLote NFL" & vbCrLf &
                       "ON  NFL.Empresa_Id = NFI.Empresa_Id " & vbCrLf &
                       "AND NFL.EndEmpresa_Id = NFI.EndEmpresa_Id " & vbCrLf &
                       "AND NFL.Cliente_Id = NFI.Cliente_Id " & vbCrLf &
                       "AND NFL.EndCliente_Id = NFI.EndCliente_Id " & vbCrLf &
                       "AND NFL.EntradaSaida_Id = NFI.EntradaSaida_Id " & vbCrLf &
                       "AND NFL.Serie_Id = NFI.Serie_Id " & vbCrLf &
                       "AND NFL.Nota_Id = NFI.Nota_Id " & vbCrLf &
                       "AND NFL.Produto_Id = NFI.Produto_Id " & vbCrLf &
                       "AND NFL.CFOP_Id = NFI.CFOP_Id " & vbCrLf &
                       "AND NFL.Sequencia_Id = NFI.Sequencia_Id " & vbCrLf

        strSQLNotasAnterior &= "INNER JOIN SubOperacoes SO " & vbCrLf &
                               "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf &
                               "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf
        strSQLNotas &= "INNER JOIN SubOperacoes SO " & vbCrLf &
                        "ON  SO.Operacao_Id = NFI.Operacao " & vbCrLf &
                        "AND SO.SubOperacoes_Id = NFI.SubOperacao " & vbCrLf


        strSQLNotasAnterior &= "INNER JOIN Produtos Prd " & vbCrLf &
                               "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf &
                               "AND Prd.Situacao  = 1 " & vbCrLf
        strSQLNotas &= "INNER JOIN Produtos Prd " & vbCrLf &
                        "ON Prd.Produto_Id = NFI.Produto_Id " & vbCrLf &
                        "AND Prd.Situacao  = 1 " & vbCrLf

        strSQLNotasAnterior &= "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf
        strSQLNotas &= "WHERE NF.Situacao=1  AND NFI.QuantidadeFiscal <> 0 " & vbCrLf

        strSQLNotasAnterior &= "AND NFI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                                "AND NFI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        strSQLNotas &= "AND NFI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                       "AND NFI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf

        If ddlDeposito.SelectedIndex > 0 Then
            strSQLNotasAnterior &= "AND NFI.Deposito = '" & ddlDeposito.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                                    "AND NFI.EndDeposito = " & ddlDeposito.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
            strSQLNotas &= "AND NFI.Deposito = '" & ddlDeposito.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                           "AND NFI.EndDeposito = " & ddlDeposito.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf
        End If

        If ddlProduto.SelectedIndex > 0 Then
            strSQLNotasAnterior &= " And NFI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
            strSQLNotas &= " And NFI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
        End If

        If RadFiscal.Checked = True Then
            strSQLNotasAnterior &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
            strSQLNotas &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
        End If

        If RadFisico.Checked = True Then
            strSQLNotasAnterior &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            strSQLNotas &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
        End If

        If DdlOperacao.Text <> Nothing Then
            strSQLNotasAnterior &= "AND SO.Operacao_Id = '" & DdlOperacao.SelectedValue & "' " & vbCrLf
            strSQLNotas &= strSQLNotasAnterior
        End If

        If DdlSuboperacao.Text <> Nothing Then
            strSQLNotasAnterior &= "AND SO.SubOperacoes_Id = '" & DdlSuboperacao.SelectedValue & "' " & vbCrLf
            strSQLNotas &= "AND SO.SubOperacoes_Id = '" & DdlSuboperacao.SelectedValue & "' " & vbCrLf
        End If

        If ckCusto.Checked Then
            strSQLNotasAnterior &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) " & vbCrLf
            strSQLNotas &= " AND (SO.ApuracaoDeCustos <> 0 And SO.ApuracaoDeCustos is not null) " & vbCrLf
        End If

        strSQL &= strSQLNotasAnterior & "AND NFI.EntradaSaida_Id = 'E' AND NF.Situacao=1 " & vbCrLf &
                                        "AND NF.Serie_id <> '101' AND NF.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        strSQL &= "SELECT COALESCE(SUM(NFI.QuantidadeFiscal) * (-1), 0) AS SaldoAnterior " & vbCrLf
        strSQL &= strSQLNotasAnterior & " AND NF.Serie_id <> '101' AND NF.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

        strSQL &= "AND NFI.EntradaSaida_Id = 'S' AND NF.Situacao=1 ) SA " & vbCrLf

        strSQL &= "UNION ALL " & vbCrLf

        strSQL &= "SELECT MOV.Data, C.Reduzido, MOV.Cliente_Id, MOV.EndCliente, C.Nome, MOV.Operacao, MOV.SubOperacao, SO.Descricao, " & vbCrLf &
                  "MOV.Ordem, " & vbCrLf &
                  "" & IIf(Not CkCDeposito.Checked, "MOV.Entradas ", "Sum(MOV.Entradas) as Entradas") & "," & IIf(Not CkCDeposito.Checked, "MOV.Saidas ", "Sum(MOV.Saidas) as Saidas") & ", 0 AS Saldo, Lote, OrdemDeProducao  " & vbCrLf

        strSQL &= "FROM (SELECT E.Movimento_Id AS Data, E.Deposito_Id AS Cliente_Id , E.EndDeposito_Id AS EndCliente , 1 AS Ordem, " & vbCrLf &
                  "E.Operacao_Id AS Operacao, E.SubOperacao_Id AS SubOperacao, COALESCE(E.Entradas, 0) AS Entradas, COALESCE(E.Saidas, 0) AS Saidas, isnull(Lote_Id, '') as Lote, isnull(OrdemDeProducao,0) AS OrdemDeProducao " & vbCrLf

        strSQL &= strSQLEstoques & "AND E.Movimento_Id BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                                   "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102) " & vbCrLf

        '	ADICIONADO
        If DdlOperacao.Text.Length > 0 Then
            strSQL &= "AND E.Operacao_Id = " & DdlOperacao.SelectedValue & " " & vbCrLf
        End If
        If DdlSuboperacao.Text.Length > 0 Then
            strSQL &= "AND E.SubOperacao_Id = " & DdlSuboperacao.SelectedValue & "" & vbCrLf
        End If

        strSQL &= "UNION ALL " & vbCrLf

        If chkOrdemDeProducaoInterna.Checked Then
            strSQL &= "SELECT OPI.Movimento AS Data, OPI.Empresa_Id AS Cliente_Id , OPI.EndEmpresa_Id AS EndCliente, 1 AS Ordem, OPI.Operacao, OPI.SubOperacao, " & vbCrLf &
                        "CASE WHEN OPI.EntradaSaida = 'E' THEN COALESCE(OPI.Quantidade, 0) " & vbCrLf &
                        "ELSE 0 END AS Entradas, " & vbCrLf &
                        "CASE WHEN OPI.EntradaSaida = 'S' THEN COALESCE(OPI.Quantidade, 0) " & vbCrLf &
                        "ELSE 0 END AS Saidas, OPI.Lote_Id AS Lote, 0 AS OrdemDeProducao " & vbCrLf &
                        "FROM OrdemDeProducaoInterna OPI " & vbCrLf &
                        "INNER JOIN SubOperacoes SO " & vbCrLf &
                        "ON SO.Operacao_Id      = OPI.Operacao " & vbCrLf &
                        "AND SO.SubOperacoes_Id = OPI.SubOperacao " & vbCrLf &
                        "INNER JOIN Produtos Prd " & vbCrLf &
                        "ON Prd.Produto_Id = OPI.Produto_Id " & vbCrLf &
                        "AND Prd.Situacao  = 1 " & vbCrLf &
                        "WHERE OPI.Empresa_Id = '" & ddlEmpresa.SelectedValue.ToString.Split("-")(0) & "' " & vbCrLf &
                        "AND OPI.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.ToString.Split("-")(1) & " " & vbCrLf &
                        "AND OPI.Movimento BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                        "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102) " & vbCrLf

            If ddlProduto.SelectedIndex > 0 Then
                strSQL &= "AND OPI.Produto_Id = '" & ddlProduto.SelectedValue & "' " & vbCrLf
            End If

            If RadFiscal.Checked = True Then
                strSQL &= "AND SO.EstoqueFiscal = 'S' " & vbCrLf
            End If

            If RadFisico.Checked = True Then
                strSQL &= "AND SO.EstoqueFisico = 'S' " & vbCrLf
            End If

            strSQL &= "UNION ALL " & vbCrLf
        End If

        strSQL &= "SELECT NF.Movimento AS Data, NFI.Cliente_Id AS Cliente_Id ,NFI.EndCliente_Id AS EndCliente, 1 AS Ordem, NF.Operacao, NF.SubOperacao, " & vbCrLf &
                  "CASE WHEN NFI.EntradaSaida_Id = 'E' THEN COALESCE(ISNULL(NFL.QUANTIDADE, NFI.QuantidadeFiscal), 0) " & vbCrLf &
                  "ELSE 0 END AS Entradas, " & vbCrLf &
                  "CASE WHEN NFI.EntradaSaida_Id = 'S' THEN COALESCE(ISNULL(NFL.QUANTIDADE, NFI.QuantidadeFiscal), 0) " & vbCrLf &
                  "ELSE 0 END AS Saidas, " & vbCrLf &
                  "ISNULL(NFL.Lote_Id,'') AS Lote, 0 AS OrdemDeProducao " & vbCrLf

        strSQL &= strSQLNotas & "AND NF.Movimento BETWEEN CONVERT(DATETIME, '" & txtDataInicial.Text.ToSqlDate() & "', 102) " & vbCrLf &
                                "AND CONVERT(DATETIME, '" & txtDataFinal.Text.ToSqlDate() & "', 102)  AND NF.Situacao=1 "

        strSQL &= " ) MOV " & vbCrLf

        strSQL &= "INNER JOIN Clientes C " & vbCrLf &
                  "ON MOV.Cliente_Id = C.Cliente_Id " & vbCrLf &
                  "AND MOV.EndCliente = C.Endereco_Id " & vbCrLf

        strSQL &= "INNER JOIN SubOperacoes SO " & vbCrLf &
                  "ON MOV.Operacao = SO.Operacao_Id " & vbCrLf &
                  "AND MOV.SubOperacao = SO.SubOperacoes_Id " & vbCrLf
        If CkCDeposito.Checked Then strSQL &= " group by  MOV.Data, C.Reduzido, MOV.Cliente_Id, MOV.EndCliente, C.Nome, MOV.Operacao, MOV.SubOperacao, SO.Descricao, MOV.Ordem, MOV.Entradas, MOV.Saidas " & vbCrLf

        strSQL &= "ORDER BY  Data, Saidas" & vbCrLf

        Return Banco.ConsultaDataSet(strSQL, "Cardex")
    End Function

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("Cardex", "RELATORIO") Then
                If ValidaCampos() Then

                    Dim ds As DataSet

                    If Pdf Then
                        ds = getDataSet()
                    Else
                        ds = getDataSetXLS()
                    End If

                    If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Período sem movimento.")
                    Else
                        Dim Saldo As Decimal = 0

                        For Each row As DataRow In ds.Tables(0).Rows
                            Dim entradas As Decimal = 0
                            Dim Saidas As Decimal = 0
                            If row("Entradas") = 0 And row("Saidas") = 0 Then
                                Saldo = row("Saldo")
                            End If
                            If Not row.IsNull("Entradas") Then
                                entradas = row("Entradas")
                            End If
                            If Not row.IsNull("Saidas") Then
                                Saidas = row("Saidas")
                            End If
                            Saldo += entradas - Saidas
                            row("Saldo") = Saldo
                        Next

                        Dim objEmpresa As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                        Dim objProduto = New Produto()
                        If Not String.IsNullOrWhiteSpace(ddlProduto.SelectedValue) Then objProduto = New Produto(ddlProduto.SelectedValue)

                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Produto", IIf(Not String.IsNullOrWhiteSpace(objProduto.Nome), objProduto.Codigo & " - " & objProduto.Nome, ""))
                        parameters.Add("ConsultaParametros", getParam())

                        If Pdf Then
                            Funcoes.BindReport(Me.Page, ds, "Cr_Cardex", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                        Else
                            gerarExcel(objEmpresa, objProduto, ds)
                        End If
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub gerarExcel(ByVal objEmpresa As Cliente, ByVal objProduto As Produto, ByVal ds As DataSet)

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
            Using package As New ExcelPackage(arquivo)

                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("CARDEX")

                'criando linha com o cabeçalho da planilha
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                'criando linha que informa o nome da empresa e o cnpj
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", objEmpresa.Nome, Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa a cidade e o estado da empresa
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", objEmpresa.Cidade, objEmpresa.CodigoEstado)
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o título do relatório
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "RELATÓRIO DE CARDEX DE ESTOQUES")
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataInicial.Text)) & " a " & String.Format("{0:dd/MM/yyyy}", CDate(txtDataFinal.Text)))
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Grupo: " & objProduto.Grupo.Codigo & "-" & objProduto.Grupo.Descricao & " - Produto: " & objProduto.Codigo & "-" & objProduto.Nome)
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha com o cabeçalho da planilha
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

                Dim Movimento As Date
                Dim Cliente As String = String.Empty
                Dim Operacao As Integer = 0
                Dim SubOpercao As Integer = 0
                Dim saldo As Decimal = 0
                Dim primeiro As Boolean = True

                ' criando conteúdo da planilha com os dados do dataset
                For Each row As DataRow In ds.Tables(0).Rows

                    Movimento = CDate(row("Data"))
                    Cliente = row("Cliente_Id")
                    Operacao = row("Operacao")
                    SubOpercao = row("SubOperacao")

                    If primeiro Then
                        saldo += row("Saldo")
                        primeiro = False
                    Else
                        saldo += row("Entradas") - row("Saidas")
                    End If

                    row("Saldo") = saldo

                    columnIndex = 1
                    For Each col As DataColumn In ds.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                        columnIndex += 1
                    Next

                    'formatando células datas
                    worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                    worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"

                    worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"

                    worksheet.Cells(String.Format("L{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"

                    'aplicando formatação nas células do conteúdo
                    If rowIndex Mod 2 = 0 Then
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        End Using
                    End If
                    rowIndex += 1
                Next

                'criando colunas de totalizadores
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Bold = True
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(String.Format("H{0}", rowIndex)).Value = String.Format("{0}", "TOTAL")

                'criando colunas de totalizadores
                worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Bold = True
                worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                worksheet.Cells(String.Format("J{0}", rowIndex)).Formula = String.Format("=SUM(J6:J{0})", rowIndex - 1)
                worksheet.Cells(String.Format("J{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"

                'criando colunas de totalizadores
                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Bold = True
                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.PatternType = ExcelFillStyle.Solid
                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Font.Color.SetColor(Color.White)
                worksheet.Cells(String.Format("K{0}", rowIndex)).Formula = String.Format("=SUM(K6:K{0})", rowIndex - 1)
                worksheet.Cells(String.Format("K{0}", rowIndex)).Style.Numberformat.Format = "##0.0000_ ;[Red]-##0.0000"


                'criando auto filtro na planilha
                worksheet.Cells("C6:L" & rowIndex).AutoFilter = True
                'worksheet.Cells("M6:N" & rowIndex).AutoFilter = True

                'setando autofit nas células da planilha
                worksheet.Cells.AutoFitColumns(0)

                'congelando quinta linha (cabeçalho)
                worksheet.View.FreezePanes(7, 1)

                'salvando planilha do excel
                package.Save()

            End Using
        End Using

        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))

    End Sub

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

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Funcoes.VerificaPermissao("Cardex", "LEITURA") Then
                Dim ds As DataSet = getDataSet()
                If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Período sem movimento.")
                Else
                    Dim dra As DataRow
                    Dim Saldo As Decimal = 0

                    For Each dra In ds.Tables(0).Rows
                        Dim entradas As Decimal = 0
                        Dim Saidas As Decimal = 0
                        If dra("Entradas") = 0 And dra("Saidas") = 0 Then
                            Saldo = dra("Saldo")
                        End If
                        If Not dra.IsNull("Entradas") Then
                            entradas = dra("Entradas")
                        End If
                        If Not dra.IsNull("Saidas") Then
                            Saidas = dra("Saidas")
                        End If
                        Saldo += entradas - Saidas
                        dra("Saldo") = Saldo
                    Next

                    gridCardex.DataSource = ds
                    gridCardex.DataBind()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para consultar registro.")
            End If
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

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Cardex")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Private Sub LimparCampos()
        DdlOperacao.SelectedIndex = 0
        DdlSuboperacao.Items.Clear()
        ddlGrupo.SelectedIndex = 0
        ddlProduto.Items.Clear()

        chkOrdemDeProducaoInterna.Checked = False

        If Left(Session("ssEmpresa"), 8) = "05272759" Then
            divOrdemDeProducaoInterna.Visible = True
        Else
            divOrdemDeProducaoInterna.Visible = False
        End If

        HID.Value = Guid.NewGuid().ToString

        txtDataInicial.Text = Format(Today, "01/01/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

        gridCardex.DataSource = Nothing
        gridCardex.DataBind()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

End Class