Imports System.Data
Imports System.IO
Imports System.Data.SqlClient
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports CrystalDecisions.Shared
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioEntradaSaidaLaudo
    Inherits BasePage

#Region "Procedimentos"

    Dim Sql As String

#End Region

#Region "Procedimentos"

    Private Sub BuscaUnidadeNegocio()
        ddl.Carregar(cmbUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(cmbUnidadeNegocio, cmbEmpresa)
    End Sub

    Private Sub BuscaEmpresa()
        ddl.Carregar(cmbEmpresa, CarregarDDL.Tabela.Empresas, cmbUnidadeNegocio.SelectedValue, True)
    End Sub

    Private Sub BuscaDeposito()
        ddl.Carregar(cmbDeposito, CarregarDDL.Tabela.Depositos, "", True)
    End Sub

    Private Sub BuscarProdutos()
        ddl.Carregar(cmbProdutos, CarregarDDL.Tabela.Produto, "Grupo = " & DdlGrupoDeProdutos.SelectedValue, True)
    End Sub

    Private Sub BuscarOperacoes()
        ddl.Carregar(cmbOperacao, CarregarDDL.Tabela.Operacao, "", True)
    End Sub

    Private Sub BuscarSubOperacoes(ByVal valor As String)
        ddl.Carregar(cmbSubOperacao, CarregarDDL.Tabela.OperacaoSubOperacao, "SO.Operacao_Id = " & cmbOperacao.SelectedValue, True)
    End Sub

    Private Sub BuscarSituacoes()
        Dim strSQL As String = "SELECT Situacao_Id, Descricao FROM Situacoes "
        Dim dsSituacoes As DataSet = Banco.ConsultaDataSet(strSQL, "Situacoes")

        For Each drSituacao As DataRow In dsSituacoes.Tables(0).Rows
            cmbSituacao.Items.Add(New ListItem(drSituacao("Descricao"), drSituacao("Situacao_Id")))
            If drSituacao("Descricao").ToString().ToUpper = "NORMAL" Then cmbSituacao.SelectedIndex = cmbSituacao.Items.Count - 1
        Next

        cmbSituacao.Items.Insert(0, "")
        If cmbSituacao.SelectedIndex = -1 Then cmbSituacao.SelectedIndex = 0
    End Sub

    Private Sub BuscarTransportes()
        Dim strSQL As String = "SELECT Codigo_Id, Descricao FROM TiposDeVeiculos "
        Dim dsTiposDeVeiculos As DataSet = Banco.ConsultaDataSet(strSQL, "TiposDeVeiculos")

        cmbTransporte.Items.Insert(0, "TODOS")
        cmbTransporte.Items.Insert(1, "TODOS MENOS CAÇAMBA")

        For Each drTiposDeVeiculos As DataRow In dsTiposDeVeiculos.Tables(0).Rows
            cmbTransporte.Items.Add(New ListItem(drTiposDeVeiculos("Descricao"), drTiposDeVeiculos("Codigo_Id")))
        Next

        cmbTransporte.SelectedIndex = 0
    End Sub

    Private Sub BuscarGrupoDeProdutos()
        Dim ds As New DataSet

        Sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao"
        Sql &= " FROM         GruposDeEstoques INNER JOIN"
        Sql &= "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo"
        Sql &= " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao"

        ds = Banco.ConsultaDataSet(Sql, "GruposDeEstoques")

        For Each drGrupo As DataRow In ds.Tables(0).Rows
            DdlGrupoDeProdutos.Items.Add(New ListItem(drGrupo("Descricao"), drGrupo("Grupo_Id")))
        Next

        DdlGrupoDeProdutos.Items.Insert(0, "")
        DdlGrupoDeProdutos.SelectedIndex = 0
    End Sub

    Public Function ValidarCampos() As Boolean
        If txtDataInicial.Text.Length = 0 Or IsDate(txtDataInicial.Text) = False Then
            MsgBox(Me.Page, "Informe uma data inicial válida!", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        If txtDataFinal.Text.Length = 0 Or IsDate(txtDataFinal.Text) = False Then
            MsgBox(Me.Page, "Informe uma data final válida!", eTitulo.Info)
            txtDataFinal.Focus()
            Return False
        End If

        If CDate(txtDataInicial.Text) > CDate(txtDataFinal.Text) Then
            MsgBox(Me.Page, "Data inicial não pode ser maior que a data final!", eTitulo.Info)
            txtDataInicial.Focus()
            Return False
        End If

        If cmbSituacao.SelectedIndex > -1 Then
            If cmbSituacao.SelectedValue.ToString.Trim = "" Then
                MsgBox(Me.Page, "Situação Obrigatoria!", eTitulo.Info)
                cmbSituacao.Focus()
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub ImprimirRelatorio(ByVal pdf As Boolean)
        Try
            If Funcoes.VerificaPermissao("RelatorioEntradaSaidaLaudo", "RELATORIO") Then
                Dim dsRelatorio As New DataSet
                Dim Texto As String = ""

                Dim ParamAnalise As String = ""
                If ddlAnalise.SelectedIndex > 0 And txtParametrosAnalise.Text.Trim.Length > 0 Then
                    If txtParametrosAnalise.Text.Contains(";") Then
                        ParamAnalise = " in (" & txtParametrosAnalise.Text.Replace(";", ",") & ")"
                    ElseIf txtParametrosAnalise.Text.Contains("-") Then
                        ParamAnalise = " Between " & txtParametrosAnalise.Text.Split("-")(0) & " and " & txtParametrosAnalise.Text.Split("-")(1)
                    Else
                        ParamAnalise = " = " & txtParametrosAnalise.Text
                    End If
                End If


                Dim strSQL As String
                strSQL = " SELECT LP.Movimento," & vbCrLf

                If optAberto.Checked Then
                    strSQL &= "        LP.Pedido AS Pedido_Id, " & vbCrLf
                Else
                    strSQL &= "        Pedidos.Pedido_Id, " & vbCrLf
                End If

                strSQL &= "        P.Produto_Id," & vbCrLf &
                         "        P.Nome," & vbCrLf &
                         " 	     CASE len(isnull(CR.Reduzido,''))   WHEN 0 THEN C.Reduzido    ELSE CR.Reduzido    END AS Reduzido," & vbCrLf &
                         " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Cliente_Id  ELSE CR.Cliente_Id  END AS Cliente_Id," & vbCrLf &
                         " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Endereco_Id ELSE CR.Endereco_Id END AS Endereco_Id," & vbCrLf &
                         " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Nome        ELSE CR.Nome        END AS NomeCliente," & vbCrLf &
                         " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Cidade      ELSE CR.Cidade      END AS ClienteCidade," & vbCrLf &
                         " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Estado      ELSE CR.Estado      END AS ClienteEstado," & vbCrLf &
                         "       case" & vbCrLf &
                         "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                         "             then LP.Deposito" & vbCrLf &
                         "             else Romaneios.Deposito" & vbCrLf &
                         "          end AS Deposito_Id," & vbCrLf &
                         "       case" & vbCrLf &
                         "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                         "             then LP.EndDeposito" & vbCrLf &
                         "             else Romaneios.EndDeposito" & vbCrLf &
                         "          end AS EndDeposito_Id," & vbCrLf &
                         "       case" & vbCrLf &
                         "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                         "             then (select Deposito.Nome from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "             else (select Deposito.Nome from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "          end AS NomeDeposito," & vbCrLf &
                         "       case" & vbCrLf &
                         "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                         "             then (select Deposito.Cidade from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "             else (select Deposito.Cidade from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "          end AS CidadeDeposito," & vbCrLf &
                         "       case" & vbCrLf &
                         "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                         "             then (select Deposito.Estado from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "             else (select Deposito.Estado from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                         "          end AS EstadoDeposito," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.PUmidade, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.PUmidade, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.PUmidade, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.PUmidade, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercUmidade," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DUmidade, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DUmidade, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DUmidade, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DUmidade, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescUmidade," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.PImpureza, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.PImpureza, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.PImpureza, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.PImpureza, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercImpureza," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DImpureza, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DImpureza, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DImpureza, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DImpureza, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescImpureza," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.PAvariado, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.PAvariado, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.PAvariado, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.PAvariado, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercAvariados," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DAvariado, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DAvariado, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DAvariado, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DAvariado, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescAvariados," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.PPH, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.PPH, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.PPH, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.PPH, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercPH," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DPH, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DPH, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DPH, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DPH, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescPH," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.PVerdoso, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.PVerdoso, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.PVerdoso, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.PVerdoso, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercVerde," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DVerdoso, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DVerdoso, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DVerdoso, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DVerdoso, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescVerde," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.Intacta, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.Intacta, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.Intacta, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.Intacta, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercTrans," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DTransgenico, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DTransgenico, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DTransgenico, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DTransgenico, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescTrans," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.POutros, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.POutros, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.POutros, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.POutros, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS PercOutros," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN ISNULL(AP.DOutros, 0)" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN ISNULL(AR.DOutros, 0)" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN ISNULL(AP.DOutros, 0)" & vbCrLf &
                         "                  ELSE ISNULL(AR.DOutros, 0)" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS DescOutros," & vbCrLf &
                         "        '' AS Intacta," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN LP.Liquido" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN Romaneios.PesoLiquido" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE len(isnull(cr.Cliente_Id,''))" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN LP.Liquido" & vbCrLf &
                         "                  ELSE Romaneios.PesoLiquido" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS Liquido," & vbCrLf &
                         "         LP.Placa," & vbCrLf &
                         "         LP.Pesagem_Id," & vbCrLf &
                         "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                         "            THEN LP.BrutoBalanca" & vbCrLf &
                         "            ELSE" & vbCrLf &
                         "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                         "            WHEN 'Rateio'" & vbCrLf &
                         "               THEN Romaneios.PesoBruto" & vbCrLf &
                         "               ELSE" & vbCrLf &
                         "             CASE len(isnull(cr.Cliente_Id,''))" & vbCrLf &
                         "               WHEN 0" & vbCrLf &
                         "                  THEN LP.BrutoBalanca" & vbCrLf &
                         "                  ELSE Romaneios.PesoBruto" & vbCrLf &
                         "               END" & vbCrLf &
                         "            END" & vbCrLf &
                         "        END AS BrutoBalanca," & vbCrLf &
                         "         LP.Empresa_Id," & vbCrLf &
                         "         LP.EndEmpresa_Id," & vbCrLf &
                         "         E.Nome AS EmpresaNome," & vbCrLf &
                         "         E.Reduzido AS EmpresaReduzido," & vbCrLf &
                         "         E.Cidade AS EmpresaCidade," & vbCrLf &
                         "         E.Estado AS EmpresaEstado," & vbCrLf &
                         "         LP.EntradaSaida," & vbCrLf &
                         " 	       CASE len(isnull(cr.Cliente_Id,'')) WHEN 0 THEN REPLICATE('0',2 - LEN(CAST(LP.Operacao AS varchar))) + CAST(LP.Operacao AS varchar) ELSE REPLICATE('0',2 - LEN(CAST(Romaneios.Operacao AS varchar))) + CAST(Romaneios.Operacao AS varchar) END AS Operacao," & vbCrLf &
                         " 	       CASE len(isnull(cr.Cliente_Id,'')) WHEN 0 THEN REPLICATE('0',2 - LEN(CAST(LP.SubOperacao AS varchar))) + CAST(LP.SubOperacao AS varchar) ELSE REPLICATE('0',2 - LEN(CAST(Romaneios.SubOperacao AS varchar))) + CAST(Romaneios.SubOperacao AS varchar) END AS SubOperacao," & vbCrLf &
                         "         LP.Depositante, " & vbCrLf &
                         "         LP.EndDepositante, " & vbCrLf

                If pdf Then
                    strSQL &= "         D.Nome AS DepositanteNome" & vbCrLf
                Else
                    strSQL &= "         D.Nome AS DepositanteNome," & vbCrLf & _
                              "         case" & vbCrLf & _
                              "            when len(isnull(LP.UsuarioAlteracao,'')) = 0 and len(isnull(LP.UsuarioCancelamento,'')) = 0" & vbCrLf & _
                              "               then 'Inclusão: ' + (convert(varchar, LP.Movimento, 103)) + '  ' + LP.UsuarioInclusao" & vbCrLf & _
                              "               else" & vbCrLf & _
                              "                  case" & vbCrLf & _
                              "                     when len(isnull(LP.UsuarioAlteracao,'')) > 0" & vbCrLf & _
                              "                        then" & vbCrLf & _
                              "                           case" & vbCrLf & _
                              "                              when len(isnull(LP.UsuarioCancelamento,'')) = 0" & vbCrLf & _
                              "                                 then 'Inclusão: ' + (convert(varchar, LP.Movimento, 103)) + '  ' + LP.UsuarioInclusao + ' - ' + 'Alteração: ' + (convert(varchar, LP.UsuarioAlteracaoData, 103)) + '  ' + LP.UsuarioAlteracao" & vbCrLf & _
                              "                                 else 'Inclusão: ' + (convert(varchar, LP.Movimento, 103)) + '  ' + LP.UsuarioInclusao + ' - ' + 'Alteração: ' + (convert(varchar, LP.UsuarioAlteracaoData, 103)) + '  ' + LP.UsuarioAlteracao + ' - ' + 'Cancelamento: ' + (convert(varchar, LP.UsuarioCancelamentoData, 103)) + '  ' + LP.UsuarioCancelamento" & vbCrLf & _
                              "                              end" & vbCrLf & _
                              "                         else  'Inclusão: ' + (convert(varchar, LP.Movimento, 103)) + '  ' + LP.UsuarioInclusao + ' - ' + 'Cancelamento: ' + (convert(varchar, LP.UsuarioCancelamentoData, 103)) + '  ' + LP.UsuarioCancelamento" & vbCrLf & _
                              "                     end" & vbCrLf & _
                              "            end AS Historico" & vbCrLf
                End If

                strSQL &= "    into #Pesagens " & vbCrLf & _
                         "    FROM Pesagem AS LP" & vbCrLf & _
                         "    LEFT JOIN RomaneiosXPesagens" & vbCrLf & _
                         "      ON LP.Empresa_Id    = RomaneiosXPesagens.Empresa_Id" & vbCrLf & _
                         "     AND LP.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id" & vbCrLf & _
                         "     AND LP.Pesagem_Id    = RomaneiosXPesagens.Pesagem_Id" & vbCrLf & _
                         "    LEFT JOIN Romaneios" & vbCrLf & _
                         "      ON RomaneiosXPesagens.Empresa_Id    = Romaneios.Empresa_Id" & vbCrLf & _
                         "     AND RomaneiosXPesagens.EndEmpresa_Id = Romaneios.EndEmpresa_Id" & vbCrLf & _
                         "     AND RomaneiosXPesagens.Romaneio_Id   = Romaneios.Romaneio_Id" & vbCrLf & _
                         "    LEFT JOIN Pedidos" & vbCrLf & _
                         "      ON Romaneios.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf & _
                         "     AND Romaneios.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf & _
                         "     AND Romaneios.Pedido        = Pedidos.Pedido_Id" & vbCrLf & _
                         "    Left Join Operacoes OP" & vbCrLf & _
                         "      on Op.Operacao_id = Pedidos.Operacao" & vbCrLf & _
                         "    LEFT JOIN Clientes CR" & vbCrLf & _
                         "      ON Pedidos.Cliente    = CR.Cliente_Id" & vbCrLf & _
                         "     AND Pedidos.EndCliente = CR.Endereco_Id" & vbCrLf & _
                         "   INNER JOIN Produtos AS P" & vbCrLf & _
                         "      ON LP.Produto = P.Produto_Id" & vbCrLf & _
                         "   INNER JOIN Clientes AS C" & vbCrLf & _
                         "      ON LP.Cliente    = C.Cliente_Id" & vbCrLf & _
                         "     AND LP.EndCliente = C.Endereco_Id" & vbCrLf & _
                         "   INNER JOIN Clientes AS E" & vbCrLf & _
                         "      ON LP.Empresa_Id    = E.Cliente_Id" & vbCrLf & _
                         "     AND LP.EndEmpresa_Id = E.Endereco_Id" & vbCrLf & _
                         "   INNER JOIN Clientes AS D" & vbCrLf & _
                         "      ON LP.Depositante    = D.Cliente_Id" & vbCrLf & _
                         "     AND LP.EndDepositante = D.Endereco_Id" & vbCrLf & _
                         "    LEFT JOIN Placas" & vbCrLf & _
                         "      ON LP.Placa = Placas.Placa_Id" & vbCrLf & _
                         "    LEFT JOIN Clientes Deposito " & vbCrLf & _
                         "      ON LP.Deposito    = Deposito.Cliente_Id " & vbCrLf & _
                         "     AND LP.EndDeposito = Deposito.Endereco_Id " & vbCrLf & _
                         "    LEFT JOIN (SELECT Empresa_Id," & vbCrLf & _
                         "                      EndEmpresa_Id," & vbCrLf & _
                         "                      Pesagem_Id," & vbCrLf & _
                         "                      Sequencia_Id," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 1  then Percentual else 0 end) as PUmidade," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 1  then Desconto   else 0 end) as DUmidade," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 2  then Percentual else 0 end) as PImpureza," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 2  then Desconto   else 0 end) as DImpureza," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 3  then Percentual else 0 end) as PAvariado," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 3  then Desconto   else 0 end) as DAvariado," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 4  then Percentual else 0 end) as PPH," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 4  then Desconto   else 0 end) as DPH," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 5  then Percentual else 0 end) as PVerdoso," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 5  then Desconto   else 0 end) as DVerdoso," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 6  then Percentual else 0 end) as PTransgenico," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 6  then Desconto   else 0 end) as DTransgenico," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 7  then Percentual else 0 end) as POutros," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 7  then Desconto   else 0 end) as DOutros," & vbCrLf & _
                         "                      sum(Case when Analise_Id = 12 then Percentual else 0 end) as Intacta" & vbCrLf

                If ParamAnalise.Length > 0 Then
                    strSQL &= "                      ,sum(Case when Analise_Id = " & ddlAnalise.SelectedValue & " then Percentual else 0 end) as AnPar" & vbCrLf
                End If

                strSQL &= "                FROM PesagemXAnalises" & vbCrLf & _
                          "               Where Analise_Id in (1,2,3,4,5,6,7,12)" & vbCrLf & _
                          "               group by Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id" & vbCrLf & _
                          "              ) AS AP" & vbCrLf & _
                          "      ON AP.Empresa_Id    = LP.Empresa_Id" & vbCrLf & _
                          "     AND AP.EndEmpresa_Id = LP.EndEmpresa_Id" & vbCrLf & _
                          "     AND AP.Pesagem_Id    = LP.Pesagem_Id" & vbCrLf & _
                          "     AND AP.Sequencia_Id  = LP.Sequencia_Id" & vbCrLf & _
                          "    LEFT JOIN (SELECT Empresa_Id," & vbCrLf & _
                          "                      EndEmpresa_Id," & vbCrLf & _
                          "                      Romaneio_Id," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 1  then Percentual else 0 end) as PUmidade," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 1  then Desconto   else 0 end) as DUmidade," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 2  then Percentual else 0 end) as PImpureza," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 2  then Desconto   else 0 end) as DImpureza," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 3  then Percentual else 0 end) as PAvariado," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 3  then Desconto   else 0 end) as DAvariado," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 4  then Percentual else 0 end) as PPH," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 4  then Desconto   else 0 end) as DPH," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 5  then Percentual else 0 end) as PVerdoso," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 5  then Desconto   else 0 end) as DVerdoso," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 6  then Percentual else 0 end) as PTransgenico," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 6  then Desconto   else 0 end) as DTransgenico," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 7  then Percentual else 0 end) as POutros," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 7  then Desconto   else 0 end) as DOutros," & vbCrLf & _
                          "                      sum(Case when Analise_Id = 12 then Percentual else 0 end) as Intacta" & vbCrLf

                If ParamAnalise.Length > 0 Then
                    strSQL &= "                      ,sum(Case when Analise_Id = " & ddlAnalise.SelectedValue & " then Percentual else 0 end) as AnPar" & vbCrLf
                End If

                strSQL &= "                 FROM RomaneiosXDescontos" & vbCrLf & _
                          "                WHERE Analise_Id in (1,2,3,4,5,6,7,12)" & vbCrLf & _
                          "                Group by Empresa_Id, EndEmpresa_Id, Romaneio_Id" & vbCrLf & _
                          "               ) AS AR" & vbCrLf & _
                          "      ON AR.Romaneio_Id   = Romaneios.Romaneio_Id" & vbCrLf & _
                          "     AND AR.Empresa_Id    = Romaneios.Empresa_Id" & vbCrLf & _
                          "     AND AR.EndEmpresa_Id = Romaneios.EndEmpresa_Id" & vbCrLf & _
                          "   WHERE LP.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd 00:00:00") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd 23:59:59") & "'" & vbCrLf & _
                          "     AND LP.Sequencia_Id  = 0" & vbCrLf & _
                          "     AND LP.Liquido      <> 0" & vbCrLf & _
                          "     AND LP.BrutoBalanca <> 0" & vbCrLf

                Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
                If op.Length > 0 Then
                    Sql &= "   AND op.classe " & IIf(rdComClasse.Checked, "", "not") & " in ('" & op & "')"
                End If

                If cmbEmpresa.SelectedIndex > 0 Then
                    Dim strEmpresa As String()
                    strEmpresa = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

                    strSQL &= " AND LP.Empresa_Id    = '" & strEmpresa(0) & "' " & vbCrLf & _
                              " AND LP.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                End If

                If cmbDeposito.SelectedIndex > 0 Then
                    Dim strDeposito As String()
                    strDeposito = cmbDeposito.SelectedValue.Split(New Char() {"-"})

                    strSQL &= " AND LP.Deposito    = '" & strDeposito(0) & "' " & vbCrLf & _
                              " AND LP.EndDeposito = " & strDeposito(1) & " " & vbCrLf
                End If

                If txtCodigoCliente.Value.Length > 0 Then
                    Dim strCliente As String()
                    strCliente = txtCodigoCliente.Value.Split(New Char() {"-"})
                    If chkClientes.Checked = True Then
                        strSQL &= " AND left(isnull(CR.Cliente_Id,C.Cliente_Id),8) = '" & strCliente(0).Substring(0, 8) & "'" & vbCrLf
                    Else
                        strSQL &= " AND isnull(CR.Cliente_Id,C.Cliente_Id)   = '" & strCliente(0) & "'" & vbCrLf & _
                                  " AND isnull(CR.Endereco_Id,C.Endereco_Id) = " & strCliente(1) & vbCrLf
                    End If
                End If

                If txtCodigoDepositante.Value.Length > 0 Then
                    Dim strDepositante As String()
                    strDepositante = txtCodigoDepositante.Value.Split(New Char() {"-"})
                    strSQL &= "AND LP.Depositante = '" & strDepositante(0) & "' " & vbCrLf & _
                              "AND LP.EndDepositante = " & strDepositante(1) & " " & vbCrLf
                End If

                If cmbOperacao.SelectedIndex > 0 Then
                    strSQL &= " AND (LP.Operacao = " & cmbOperacao.SelectedValue & ") " & vbCrLf
                End If

                If cmbSubOperacao.SelectedIndex > 0 Then
                    strSQL &= " AND (LP.SubOperacao = " & cmbSubOperacao.SelectedValue & ") " & vbCrLf
                End If

                If DdlGrupoDeProdutos.SelectedIndex > 0 Then
                    strSQL &= " AND (P.Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "') " & vbCrLf
                End If

                If cmbProdutos.SelectedIndex > 0 Then
                    strSQL &= " AND (P.Produto_Id = '" & cmbProdutos.SelectedValue & "') " & vbCrLf
                End If

                If optFechado.Checked = True And (Trim(txtPedido.Text) = "" Or Trim(txtPedido.Text) = "0") Then
                    If optEntrada.Checked = True Then
                        strSQL &= "AND LP.EntradaSaida = 'E' " & vbCrLf
                    ElseIf optSaida.Checked = True Then
                        strSQL &= "AND LP.EntradaSaida = 'S' " & vbCrLf
                    End If
                End If

                If optSim.Checked = True Then
                    strSQL &= "AND LP.Fisico = 'S' " & vbCrLf
                ElseIf optNao.Checked = True Then
                    strSQL &= "AND LP.Fisico = 'N' " & vbCrLf
                End If

                If optAberto.Checked = True Then
                    strSQL &= "AND LP.SegundaPesagem = 0 " & vbCrLf
                ElseIf optFechado.Checked = True Then
                    strSQL &= "AND LP.SegundaPesagem > 0 " & vbCrLf
                End If

                If cmbTransporte.SelectedIndex > 0 Then
                    If cmbTransporte.SelectedIndex = 1 Then
                        strSQL &= "AND Placas.TipoVeiculo_Id <> 2 " & vbCrLf
                    Else
                        strSQL &= "AND Placas.TipoVeiculo_Id = " & cmbTransporte.SelectedValue & " " & vbCrLf
                    End If
                End If

                If cmbSituacao.SelectedIndex > -1 Then strSQL &= "AND LP.Situacao = " & cmbSituacao.SelectedValue & " " & vbCrLf

                If optAberto.Checked Then
                    If Trim(txtPedido.Text) <> "" And Trim(txtPedido.Text) <> "0" Then strSQL &= "AND LP.Pedido = " & txtPedido.Text & " " & vbCrLf
                Else
                    If Trim(txtPedido.Text) <> "" And Trim(txtPedido.Text) <> "0" Then strSQL &= "AND Pedidos.Pedido_Id = " & txtPedido.Text & " " & vbCrLf
                End If

                If rdListaSemAgrupar.Checked Then strSQL &= "AND LP.Situacao = " & cmbSituacao.SelectedValue & " " & vbCrLf

                If ParamAnalise.Length > 0 Then
                    strSQL &= " AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                              "        WHEN 'AGRUPAMENTO' " & vbCrLf & _
                              "          THEN ISNULL(AP.AnPar, 0) " & vbCrLf & _
                              "          ELSE CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                              "                 WHEN 'Rateio' " & vbCrLf & _
                              "                   THEN ISNULL(AR.AnPar, 0) " & vbCrLf & _
                              "                   ELSE CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                              "                          WHEN 0 " & vbCrLf & _
                              "                            THEN ISNULL(AP.AnPar, 0) " & vbCrLf & _
                              "                            ELSE ISNULL(AR.AnPar, 0) " & vbCrLf & _
                              "                        END " & vbCrLf & _
                              "               END " & vbCrLf & _
                              "      END) " & ParamAnalise & vbCrLf
                End If

                'If rdISim.Checked Then
                '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "            ELSE " & vbCrLf & _
                '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "            WHEN 'Rateio' " & vbCrLf & _
                '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               ELSE " & vbCrLf & _
                '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                '              "               WHEN 0 " & vbCrLf & _
                '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               END " & vbCrLf & _
                '              "            END " & vbCrLf & _
                '              "         END) = 1 " & vbCrLf
                'ElseIf rdPositivo.Checked Then
                '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "            ELSE " & vbCrLf & _
                '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "            WHEN 'Rateio' " & vbCrLf & _
                '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               ELSE " & vbCrLf & _
                '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                '              "               WHEN 0 " & vbCrLf & _
                '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               END " & vbCrLf & _
                '              "            END " & vbCrLf & _
                '              "         END) = 2 " & vbCrLf
                'ElseIf rdPSim.Checked Then
                '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "            ELSE " & vbCrLf & _
                '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "            WHEN 'Rateio' " & vbCrLf & _
                '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               ELSE " & vbCrLf & _
                '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                '              "               WHEN 0 " & vbCrLf & _
                '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               END " & vbCrLf & _
                '              "            END " & vbCrLf & _
                '              "         END) IN(1,2) " & vbCrLf
                'ElseIf rdINao.Checked Then
                '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "            ELSE " & vbCrLf & _
                '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                '              "            WHEN 'Rateio' " & vbCrLf & _
                '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               ELSE " & vbCrLf & _
                '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                '              "               WHEN 0 " & vbCrLf & _
                '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                '              "               END " & vbCrLf & _
                '              "            END " & vbCrLf & _
                '              "         END) = 0 " & vbCrLf
                'End If

                If rdListaSemAgrupar.Checked Then
                    strSQL &= " AND NOT ISNULL(Romaneios.Processo,'') = 'AGRUPAMENTO' " & vbCrLf & _
                             vbCrLf & " INSERT INTO #Pesagens " & vbCrLf & _
                             " SELECT LP.Movimento," & vbCrLf & _
                             "        P.Produto_Id," & vbCrLf

                    If optAberto.Checked Then
                        strSQL &= "        LP.Pedido AS Pedido_Id, " & vbCrLf
                    Else
                        strSQL &= "        Pedidos.Pedido_Id, " & vbCrLf
                    End If

                    strSQL &= "      P.Nome," & vbCrLf &
                             " 	     CASE len(isnull(CR.Reduzido,''))   WHEN 0 THEN C.Reduzido    ELSE CR.Reduzido    END AS Reduzido," & vbCrLf &
                             " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Cliente_Id  ELSE CR.Cliente_Id  END AS Cliente_Id," & vbCrLf &
                             " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Endereco_Id ELSE CR.Endereco_Id END AS Endereco_Id," & vbCrLf &
                             " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Nome        ELSE CR.Nome        END AS NomeCliente," & vbCrLf &
                             " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Cidade      ELSE CR.Cidade      END AS ClienteCidade," & vbCrLf &
                             " 	     CASE len(isnull(CR.Cliente_Id,'')) WHEN 0 THEN C.Estado      ELSE CR.Estado      END AS ClienteEstado," & vbCrLf &
                             "       case" & vbCrLf &
                             "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                             "             then LP.Deposito" & vbCrLf &
                             "             else Romaneios.Deposito" & vbCrLf &
                             "          end AS Deposito_Id," & vbCrLf &
                             "       case" & vbCrLf &
                             "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                             "             then LP.EndDeposito" & vbCrLf &
                             "             else Romaneios.EndDeposito" & vbCrLf &
                             "          end AS EndDeposito_Id," & vbCrLf &
                             "       case" & vbCrLf &
                             "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                             "             then (select Deposito.Nome from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "             else (select Deposito.Nome from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "          end AS NomeDeposito," & vbCrLf &
                             "       case" & vbCrLf &
                             "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                             "             then (select Deposito.Cidade from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "             else (select Deposito.Cidade from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "          end AS CidadeDeposito," & vbCrLf &
                             "       case" & vbCrLf &
                             "          when isnull(Romaneios.Romaneio_Id,0) = 0" & vbCrLf &
                             "             then (select Deposito.Estado from Clientes AS Deposito where LP.Deposito = Deposito.Cliente_Id AND LP.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "             else (select Deposito.Estado from Clientes AS Deposito where Romaneios.Deposito = Deposito.Cliente_Id AND Romaneios.EndDeposito = Deposito.Endereco_Id)" & vbCrLf &
                             "          end AS EstadoDeposito," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.PUmidade, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.PUmidade, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.PUmidade, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.PUmidade, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercUmidade," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DUmidade, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DUmidade, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DUmidade, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DUmidade, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescUmidade," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.PImpureza, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.PImpureza, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.PImpureza, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.PImpureza, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercImpureza," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DImpureza, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DImpureza, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DImpureza, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DImpureza, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescImpureza," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.PAvariado, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.PAvariado, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.PAvariado, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.PAvariado, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercAvariados," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DAvariado, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DAvariado, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DAvariado, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DAvariado, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescAvariados," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.PPH, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.PPH, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.PPH, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.PPH, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercPH," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DPH, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DPH, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DPH, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DPH, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescPH," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.PVerdoso, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.PVerdoso, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.PVerdoso, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.PVerdoso, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercVerde," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DVerdoso, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DVerdoso, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DVerdoso, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DVerdoso, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescVerde," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.Intacta, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.Intacta, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.Intacta, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.Intacta, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercTrans," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DTransgenico, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DTransgenico, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DTransgenico, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DTransgenico, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescTrans," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.POutros, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.POutros, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.POutros, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.POutros, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS PercOutros," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN ISNULL(AP.DOutros, 0)" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN ISNULL(AR.DOutros, 0)" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE isnull(Romaneios.Romaneio_Id, 0)" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN ISNULL(AP.DOutros, 0)" & vbCrLf &
                             "                  ELSE ISNULL(AR.DOutros, 0)" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS DescOutros," & vbCrLf &
                             "        '' AS Intacta," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN LP.Liquido" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN Romaneios.PesoLiquido" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE len(isnull(cr.Cliente_Id,''))" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN LP.Liquido" & vbCrLf &
                             "                  ELSE Romaneios.PesoLiquido" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS Liquido," & vbCrLf &
                             "         LP.Placa," & vbCrLf &
                             "         LP.Pesagem_Id," & vbCrLf &
                             "       CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "         WHEN 'AGRUPAMENTO'" & vbCrLf &
                             "            THEN LP.BrutoBalanca" & vbCrLf &
                             "            ELSE" & vbCrLf &
                             "          CASE isnull(Romaneios.Processo,'')" & vbCrLf &
                             "            WHEN 'Rateio'" & vbCrLf &
                             "               THEN Romaneios.PesoBruto" & vbCrLf &
                             "               ELSE" & vbCrLf &
                             "             CASE len(isnull(cr.Cliente_Id,''))" & vbCrLf &
                             "               WHEN 0" & vbCrLf &
                             "                  THEN LP.BrutoBalanca" & vbCrLf &
                             "                  ELSE Romaneios.PesoBruto" & vbCrLf &
                             "               END" & vbCrLf &
                             "            END" & vbCrLf &
                             "        END AS BrutoBalanca," & vbCrLf &
                             "         LP.Empresa_Id," & vbCrLf &
                             "         LP.EndEmpresa_Id," & vbCrLf &
                             "         E.Nome AS EmpresaNome," & vbCrLf &
                             "         E.Reduzido AS EmpresaReduzido," & vbCrLf &
                             "         E.Cidade AS EmpresaCidade," & vbCrLf &
                             "         E.Estado AS EmpresaEstado," & vbCrLf &
                             "         LP.EntradaSaida," & vbCrLf &
                             " 	       CASE len(isnull(cr.Cliente_Id,'')) WHEN 0 THEN REPLICATE('0',2 - LEN(CAST(LP.Operacao AS varchar))) + CAST(LP.Operacao AS varchar) ELSE REPLICATE('0',2 - LEN(CAST(Romaneios.Operacao AS varchar))) + CAST(Romaneios.Operacao AS varchar) END AS Operacao," & vbCrLf &
                             " 	       CASE len(isnull(cr.Cliente_Id,'')) WHEN 0 THEN REPLICATE('0',2 - LEN(CAST(LP.SubOperacao AS varchar))) + CAST(LP.SubOperacao AS varchar) ELSE REPLICATE('0',2 - LEN(CAST(Romaneios.SubOperacao AS varchar))) + CAST(Romaneios.SubOperacao AS varchar) END AS SubOperacao," & vbCrLf &
                             "         LP.Depositante, LP.EndDepositante, D.Nome AS DepositanteNome" & vbCrLf &
                             "    FROM Pesagem AS LP" & vbCrLf &
                             "    LEFT JOIN RomaneiosXPesagens" & vbCrLf &
                             "      ON LP.Empresa_Id    = RomaneiosXPesagens.Empresa_Id" & vbCrLf &
                             "     AND LP.EndEmpresa_Id = RomaneiosXPesagens.EndEmpresa_Id" & vbCrLf &
                             "     AND LP.Pesagem_Id    = RomaneiosXPesagens.Pesagem_Id" & vbCrLf &
                             "    LEFT JOIN Romaneios" & vbCrLf &
                             "      ON RomaneiosXPesagens.Empresa_Id    = Romaneios.Empresa_Id" & vbCrLf &
                             "     AND RomaneiosXPesagens.EndEmpresa_Id = Romaneios.EndEmpresa_Id" & vbCrLf &
                             "     AND RomaneiosXPesagens.Romaneio_Id   = Romaneios.Romaneio_Id" & vbCrLf &
                             "    LEFT JOIN Pedidos" & vbCrLf &
                             "      ON Romaneios.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf &
                             "     AND Romaneios.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
                             "     AND Romaneios.Pedido        = Pedidos.Pedido_Id" & vbCrLf &
                             "    LEFT JOIN Clientes CR" & vbCrLf &
                             "      ON Pedidos.Cliente    = CR.Cliente_Id" & vbCrLf &
                             "     AND Pedidos.EndCliente = CR.Endereco_Id" & vbCrLf &
                             "   INNER JOIN Produtos AS P" & vbCrLf &
                             "      ON LP.Produto = P.Produto_Id" & vbCrLf &
                             "   INNER JOIN Clientes AS C" & vbCrLf &
                             "      ON LP.Cliente    = C.Cliente_Id" & vbCrLf &
                             "     AND LP.EndCliente = C.Endereco_Id" & vbCrLf &
                             "   INNER JOIN Clientes AS E" & vbCrLf &
                             "      ON LP.Empresa_Id    = E.Cliente_Id" & vbCrLf &
                             "     AND LP.EndEmpresa_Id = E.Endereco_Id" & vbCrLf &
                             "   INNER JOIN Clientes AS D" & vbCrLf &
                             "      ON LP.Depositante    = D.Cliente_Id" & vbCrLf &
                             "     AND LP.EndDepositante = D.Endereco_Id" & vbCrLf &
                             "    LEFT JOIN Placas" & vbCrLf &
                             "      ON LP.Placa = Placas.Placa_Id" & vbCrLf &
                             "    LEFT JOIN (SELECT Empresa_Id," & vbCrLf &
                             "                      EndEmpresa_Id," & vbCrLf &
                             "                      Pesagem_Id," & vbCrLf &
                             "                      Sequencia_Id," & vbCrLf &
                             "                      sum(Case when Analise_Id = 1 then Percentual else 0 end) as PUmidade," & vbCrLf &
                             "                      sum(Case when Analise_Id = 1 then Desconto   else 0 end) as DUmidade," & vbCrLf &
                             "                      sum(Case when Analise_Id = 2 then Percentual else 0 end) as PImpureza," & vbCrLf &
                             "                      sum(Case when Analise_Id = 2 then Desconto   else 0 end) as DImpureza," & vbCrLf &
                             "                      sum(Case when Analise_Id = 3 then Percentual else 0 end) as PAvariado," & vbCrLf &
                             "                      sum(Case when Analise_Id = 3 then Desconto   else 0 end) as DAvariado," & vbCrLf &
                             "                      sum(Case when Analise_Id = 4 then Percentual else 0 end) as PPH," & vbCrLf &
                             "                      sum(Case when Analise_Id = 4 then Desconto   else 0 end) as DPH," & vbCrLf &
                             "                      sum(Case when Analise_Id = 5 then Percentual else 0 end) as PVerdoso," & vbCrLf &
                             "                      sum(Case when Analise_Id = 5 then Desconto   else 0 end) as DVerdoso," & vbCrLf &
                             "                      sum(Case when Analise_Id = 6 then Percentual else 0 end) as PTransgenico," & vbCrLf &
                             "                      sum(Case when Analise_Id = 6 then Desconto   else 0 end) as DTransgenico," & vbCrLf &
                             "                      sum(Case when Analise_Id = 7 then Percentual else 0 end) as POutros," & vbCrLf &
                             "                      sum(Case when Analise_Id = 7 then Desconto   else 0 end) as DOutros," & vbCrLf &
                             "                      sum(Case when Analise_Id = 12 then Percentual else 0 end) as Intacta" & vbCrLf

                    If ParamAnalise.Length > 0 Then
                        strSQL &= "                      ,sum(Case when Analise_Id = " & ddlAnalise.SelectedValue & " then Percentual else 0 end) as AnPar" & vbCrLf
                    End If

                    strSQL &= "                FROM PesagemXAnalises" & vbCrLf & _
                              "               Where Analise_Id in (1,2,3,4,5,6,7,12)" & vbCrLf & _
                              "               group by Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id" & vbCrLf & _
                              "              ) AS AP" & vbCrLf & _
                              "      ON AP.Empresa_Id    = LP.Empresa_Id" & vbCrLf & _
                              "     AND AP.EndEmpresa_Id = LP.EndEmpresa_Id" & vbCrLf & _
                              "     AND AP.Pesagem_Id    = LP.Pesagem_Id" & vbCrLf & _
                              "     AND AP.Sequencia_Id  = LP.Sequencia_Id" & vbCrLf & _
                              "    LEFT JOIN (SELECT Empresa_Id," & vbCrLf & _
                              "                      EndEmpresa_Id," & vbCrLf & _
                              "                      Romaneio_Id," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 1 then Percentual else 0 end) as PUmidade," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 1 then Desconto   else 0 end) as DUmidade," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 2 then Percentual else 0 end) as PImpureza," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 2 then Desconto   else 0 end) as DImpureza," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 3 then Percentual else 0 end) as PAvariado," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 3 then Desconto   else 0 end) as DAvariado," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 4 then Percentual else 0 end) as PPH," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 4 then Desconto   else 0 end) as DPH," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 5 then Percentual else 0 end) as PVerdoso," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 5 then Desconto   else 0 end) as DVerdoso," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 6 then Percentual else 0 end) as PTransgenico," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 6 then Desconto   else 0 end) as DTransgenico," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 7 then Percentual else 0 end) as POutros," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 7 then Desconto   else 0 end) as DOutros," & vbCrLf & _
                              "                      sum(Case when Analise_Id = 12 then Percentual else 0 end) as Intacta" & vbCrLf

                    If ParamAnalise.Length > 0 Then
                        strSQL &= "                      ,sum(Case when Analise_Id = " & ddlAnalise.SelectedValue & " then Percentual else 0 end) as AnPar" & vbCrLf
                    End If


                    strSQL &= "                 FROM RomaneiosXDescontos" & vbCrLf & _
                              "                WHERE Analise_Id in (1,2,3,4,5,6,7,12)" & vbCrLf & _
                              "                Group by Empresa_Id, EndEmpresa_Id, Romaneio_Id" & vbCrLf & _
                              "               ) AS AR" & vbCrLf & _
                              "      ON AR.Romaneio_Id   = Romaneios.Romaneio_Id" & vbCrLf & _
                              "     AND AR.Empresa_Id    = Romaneios.Empresa_Id" & vbCrLf & _
                              "     AND AR.EndEmpresa_Id = Romaneios.EndEmpresa_Id" & vbCrLf & _
                              "   WHERE LP.Movimento BETWEEN '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd 00:00:00") & "' AND '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd 23:59:59") & "'" & vbCrLf & _
                              "     AND LP.Sequencia_Id  = 0" & vbCrLf & _
                              "     AND LP.Liquido      <> 0" & vbCrLf & _
                              "     AND LP.BrutoBalanca <> 0" & vbCrLf & _
                              "     AND isnull(LP.RegistroMestre,0) > 0 "

                    If cmbEmpresa.SelectedIndex > 0 Then
                        Dim strEmpresa As String()
                        strEmpresa = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

                        strSQL &= " AND LP.Empresa_Id    = '" & strEmpresa(0) & "' " & vbCrLf & _
                                  " AND LP.EndEmpresa_Id = " & strEmpresa(1) & " " & vbCrLf
                    End If

                    If cmbDeposito.SelectedIndex > 0 Then
                        Dim strDeposito As String()
                        strDeposito = cmbDeposito.SelectedValue.Split(New Char() {"-"})

                        strSQL &= " AND LP.Deposito    = '" & strDeposito(0) & "' " & vbCrLf & _
                                  " AND LP.EndDeposito = " & strDeposito(1) & " " & vbCrLf
                    End If

                    If txtCodigoCliente.Value.Length > 0 Then
                        Dim strCliente As String()
                        strCliente = txtCodigoCliente.Value.Split(New Char() {"-"})
                        If chkClientes.Checked = True Then
                            strSQL &= " AND Left(isnull(CR.Cliente_Id,C.Cliente_Id),8) = '" & strCliente(0).Substring(0, 8) & "'" & vbCrLf
                        Else
                            strSQL &= " AND isnull(CR.Cliente_Id,C.Cliente_Id)   = '" & strCliente(0) & "'" & vbCrLf & _
                                      " AND isnull(CR.Endereco_Id,C.Endereco_Id) = " & strCliente(1) & vbCrLf
                        End If
                    End If

                    If txtCodigoDepositante.Value.Length > 0 Then
                        Dim strDepositante As String()
                        strDepositante = txtCodigoDepositante.Value.Split(New Char() {";"})
                        strSQL &= "AND LP.Depositante = '" & strDepositante(0) & "' " & vbCrLf & _
                                  "AND LP.EndDepositante = " & strDepositante(1) & " " & vbCrLf
                    End If

                    If cmbOperacao.SelectedIndex > 0 Then
                        strSQL &= " AND (LP.Operacao = " & cmbOperacao.SelectedValue & ") " & vbCrLf
                    End If

                    If cmbSubOperacao.SelectedIndex > 0 Then
                        strSQL &= " AND (LP.SubOperacao = " & cmbSubOperacao.SelectedValue & ") " & vbCrLf
                    End If

                    If DdlGrupoDeProdutos.SelectedIndex > 0 Then
                        strSQL &= " AND (P.Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "') " & vbCrLf
                    End If

                    If cmbProdutos.SelectedIndex > 0 Then
                        strSQL &= " AND (P.Produto_Id = '" & cmbProdutos.SelectedValue & "') " & vbCrLf
                    End If

                    If optFechado.Checked = True And (Trim(txtPedido.Text) = "" Or Trim(txtPedido.Text) = "0") Then
                        If optEntrada.Checked = True Then
                            strSQL &= "AND LP.EntradaSaida = 'E' " & vbCrLf
                        ElseIf optSaida.Checked = True Then
                            strSQL &= "AND LP.EntradaSaida = 'S' " & vbCrLf
                        End If
                    End If

                    If optSim.Checked = True Then
                        strSQL &= "AND LP.Fisico = 'S' " & vbCrLf
                    ElseIf optNao.Checked = True Then
                        strSQL &= "AND LP.Fisico = 'N' " & vbCrLf
                    End If

                    If optAberto.Checked = True Then
                        strSQL &= "AND LP.SegundaPesagem = 0 " & vbCrLf
                    ElseIf optFechado.Checked = True Then
                        strSQL &= "AND LP.SegundaPesagem > 0 " & vbCrLf
                    End If

                    If cmbTransporte.SelectedIndex > 0 Then
                        If cmbTransporte.SelectedIndex = 1 Then
                            strSQL &= "AND Placas.TipoVeiculo_Id <> 2 " & vbCrLf
                        Else
                            strSQL &= "AND Placas.TipoVeiculo_Id = " & cmbTransporte.SelectedValue & " " & vbCrLf
                        End If
                    End If

                    If optAberto.Checked Then
                        If Trim(txtPedido.Text) <> "" And Trim(txtPedido.Text) <> "0" Then strSQL &= "AND LP.Pedido = " & txtPedido.Text & " " & vbCrLf
                    Else
                        If Trim(txtPedido.Text) <> "" And Trim(txtPedido.Text) <> "0" Then strSQL &= "AND Pedidos.Pedido_Id = " & txtPedido.Text & " " & vbCrLf
                    End If

                    If ParamAnalise.Length > 0 Then
                        strSQL &= " AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                                  "        WHEN 'AGRUPAMENTO' " & vbCrLf & _
                                  "          THEN ISNULL(AP.AnPar, 0) " & vbCrLf & _
                                  "          ELSE CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                                  "                 WHEN 'Rateio' " & vbCrLf & _
                                  "                   THEN ISNULL(AR.AnPar, 0) " & vbCrLf & _
                                  "                   ELSE CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                                  "                          WHEN 0 " & vbCrLf & _
                                  "                            THEN ISNULL(AP.AnPar, 0) " & vbCrLf & _
                                  "                            ELSE ISNULL(AR.AnPar, 0) " & vbCrLf & _
                                  "                        END " & vbCrLf & _
                                  "               END " & vbCrLf & _
                                  "      END) " & ParamAnalise & vbCrLf
                    End If

                    'If rdISim.Checked Then
                    '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                    '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                    '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                    '              "            ELSE " & vbCrLf & _
                    '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                    '              "            WHEN 'Rateio' " & vbCrLf & _
                    '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                    '              "               ELSE " & vbCrLf & _
                    '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                    '              "               WHEN 0 " & vbCrLf & _
                    '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                    '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                    '              "               END " & vbCrLf & _
                    '              "            END " & vbCrLf & _
                    '              "         END) = 1 " & vbCrLf
                    'ElseIf rdINao.Checked Then
                    '    strSQL &= "AND (CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                    '              "         WHEN 'AGRUPAMENTO' " & vbCrLf & _
                    '              "            THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                    '              "            ELSE " & vbCrLf & _
                    '              "          CASE isnull(Romaneios.Processo,'') " & vbCrLf & _
                    '              "            WHEN 'Rateio' " & vbCrLf & _
                    '              "               THEN ISNULL(AR.Intacta, 0) " & vbCrLf & _
                    '              "               ELSE " & vbCrLf & _
                    '              "             CASE isnull(Romaneios.Romaneio_Id, 0) " & vbCrLf & _
                    '              "               WHEN 0 " & vbCrLf & _
                    '              "                  THEN ISNULL(AP.Intacta, 0) " & vbCrLf & _
                    '              "                  ELSE ISNULL(AR.Intacta, 0) " & vbCrLf & _
                    '              "               END " & vbCrLf & _
                    '              "            END " & vbCrLf & _
                    '              "         END) = 0 " & vbCrLf
                    'End If
                End If

                strSQL &= " Select Movimento, Pedido_Id, Produto_Id, Nome, Reduzido, Cliente_Id, Endereco_Id, NomeCliente, ClienteCidade, ClienteEstado, Deposito_Id, EndDeposito_Id, " & vbCrLf &
                          "NomeDeposito, CidadeDeposito, EstadoDeposito, PercUmidade, DescUmidade, PercImpureza, DescImpureza, PercAvariados, DescAvariados, " & vbCrLf &
                          "PercPH, DescPH, PercVerde, DescVerde, PercTrans, DescTrans, PercOutros, DescOutros, Intacta, Liquido, Placa, Pesagem_Id, BrutoBalanca, " & vbCrLf &
                          "Empresa_Id, EndEmpresa_Id, EmpresaNome, EmpresaReduzido, EmpresaCidade, EmpresaEstado, EntradaSaida, Operacao, SubOperacao, " & vbCrLf &
                          "Depositante, EndDepositante, DepositanteNome  " & IIf(pdf, "", ",Historico") & vbCrLf &
                          "From #Pesagens " & vbCrLf

                dsRelatorio = Banco.ConsultaDataSet(strSQL, "Pesagem")

                If pdf Then
                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("NomeEmpresa", HttpContext.Current.Session("ssNomeEmpresa"))
                    parameters.Add("Periodo", "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text)
                    parameters.Add("Media", IIf(optAritmetica.Checked, "Aritmetica", "Ponderada"))
                    parameters.Add("EntradaSaida", IIf(ChkDepositante.Checked, "Relatório De Entrada E Saida De Laudo Por Depositante.", IIf(ckAgruparDeposito.Checked, "Relatório De Entrada E Saida De Laudo Por Deposito.", "Relatório De Entrada E Saida De Laudo.")))

                    Funcoes.BindReport(Me.Page, dsRelatorio, IIf(ChkDepositante.Checked, "Cr_EntradasSaidasLaudoDepositante", IIf(ckAgruparDeposito.Checked, "Cr_EntradasSaidasLaudoPorDeposito", "Cr_EntradasSaidasLaudo")), eExportType.PDF, parameters)
                Else
                    Dim Periodo As String = "Mês ou Período/Ano : " & txtDataInicial.Text & " a " & txtDataFinal.Text
                    Dim cabecalho As String = ""
                    cabecalho = "Pesagem" & " - Empresa : " & cmbEmpresa.SelectedValue.Split("-")(0)

                    Dim colunas As New Dictionary(Of String, eTipoCampo)
                    colunas.Add("Movimento", eTipoCampo.Data)
                    'colunas.Add("Pedido_Id", eTipoCampo.Numerico)
                    'colunas.Add("Produto_Id", eTipoCampo.Numerico)
                    'colunas.Add("Cliente_Id", eTipoCampo.Numerico)
                    'colunas.Add("Deposito_Id", eTipoCampo.Numerico)
                    colunas.Add("PercUmidade", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescUmidade", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercImpureza", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescImpureza", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercAvariados", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescAvariados", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercPH", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescPH", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercVerde", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescVerde", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercTrans", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescTrans", eTipoCampo.ValorComTotalizador)
                    colunas.Add("PercOutros", eTipoCampo.ValorComTotalizador)
                    colunas.Add("DescOutros", eTipoCampo.ValorComTotalizador)
                    colunas.Add("Intacta", eTipoCampo.ValorComTotalizador)
                    colunas.Add("Liquido", eTipoCampo.Numerico)
                    colunas.Add("Pesagem", eTipoCampo.Numerico)
                    colunas.Add("BrutoBalanca", eTipoCampo.Numerico)
                    'colunas.Add("Empresa_Id", eTipoCampo.Numerico)
                    'colunas.Add("Depositante", eTipoCampo.Numerico)

                    For Each dr In dsRelatorio.Tables(0).Rows
                        If dr("PercTrans") = 0 Then
                            dr("Intacta") = "NÃO"
                        ElseIf dr("PercTrans") = 1 Then
                            dr("Intacta") = "NEGATIVO"
                        ElseIf dr("PercTrans") = 2 Then
                            dr("Intacta") = "POSITIVO"
                        ElseIf dr("PercTrans") = 3 Then
                            dr("Intacta") = "DECLARADO"
                        ElseIf dr("PercTrans") = 4 Then
                            dr("Intacta") = "PARTICIPANTE"
                        ElseIf dr("PercTrans") = 5 Then
                            dr("Intacta") = "ORIGEM PARTICIPANTE"
                        End If
                    Next

                    Funcoes.BindExcelOffice(Me.Page, dsRelatorio, "Posição de Pedidos", colunas, True)
                End If

                'Texto = "PERÍODO DE " & txtDataInicial.Text & " À " & txtDataFinal.Text

                'If optSim.Checked = True Then
                '    Texto = Texto & "  -  FÍSICO = S"
                'ElseIf optNao.Checked = True Then
                '    Texto = Texto & "  -  FÍSICO = N"
                'End If

                'If optAberto.Checked = True Then
                '    Texto = Texto & "  -  LAUDOS ABERTOS"
                'ElseIf optFechado.Checked = True Then
                '    Texto = Texto & "  -  LAUDOS ENCERRADOS"
                'End If

                'Texto = Texto & "  -  TRANSPORTE = " & cmbTransporte.SelectedItem.Text
                'Texto = Texto & "  -  SITUAÇÃO = " & cmbSituacao.SelectedItem.Text

                'If optEntrada.Checked = True Then
                '    If ChkDepositante.Checked = True Then
                '        Texto = Texto & "Posição das Entradas de Cargas e Descargas por Depositante"
                '    ElseIf ckAgruparDeposito.Checked Then
                '        Texto = Texto & "Posição das Entradas de Cargas e Descargas por Deposito"
                '    Else
                '        If txtCodigoDepositante.Value.Length > 0 Then
                '            Dim strDepositante As String()
                '            strDepositante = txtDepositante.Text.Split(New Char() {"-"})
                '            Texto = Texto & "Posição das Entradas depositadas por " & strDepositante(0)
                '        Else
                '            Texto = Texto & "Posição das Entradas de Cargas e Descargas"
                '        End If
                '    End If
                'ElseIf optSaida.Checked = True Then
                '    If ChkDepositante.Checked = True Then
                '        Texto = Texto & "Posição das Saídas de Cargas e Descargas por Depositante"
                '    Else
                '        If txtCodigoDepositante.Value.Length > 0 Then
                '            Dim strDepositante As String()
                '            strDepositante = txtDepositante.Text.Split(New Char() {"-"})
                '            Texto = Texto & "Posição das Saidas depositadas por " & strDepositante(0)
                '        Else
                '            Texto = Texto & "Posição das Saídas de Cargas e Descargas"
                '        End If
                '    End If
                'End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        If Not Session("objClienteDESL" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteDESL" & HID.Value), [Lib].Negocio.Cliente))
            txtDepositante.Text = itemCliente.Text
            txtCodigoDepositante.Value = itemCliente.Value
            Session.Remove("objClienteDESL" & HID.Value)
        ElseIf Not Session("objClienteESL" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteESL" & HID.Value), [Lib].Negocio.Cliente))
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteESL" & HID.Value)
        ElseIf Not Session("objPedidosRESL" & HID.Value) Is Nothing Then
            Dim p As Pedido = CType(Session("objPedidosRESL" & HID.Value), [Lib].Negocio.Pedido)
            txtPedido.Text = p.Codigo
            Session.Remove("objPedidosRESL" & HID.Value)
        End If
    End Sub

    Private Function GetSql() As String
        Dim sql As String = ""
        Dim Empresa() As String
        Dim Condicao As String = ""

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
            "                    Pesagem.EndEmpresa_Id = PesagemXAnalises.EndEmpresa_Id AND Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id AND PesagemXAnalises.Sequencia_Id = 0 LEFT OUTER JOIN" & vbCrLf & _
            "                    SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
            "                    Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
            "                    Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
            "                    Clientes AS Depositante ON Pesagem.Depositante = Depositante.Cliente_Id AND Pesagem.EndDepositante = Depositante.Endereco_Id" & vbCrLf & _
            "   WHERE       (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0" & vbCrLf

        If cmbEmpresa.SelectedIndex > 0 Then
            Empresa = cmbEmpresa.SelectedValue.Split("-")
            Condicao = "        AND (Pesagem.Empresa_Id = '" & Empresa(0) & "')" & vbCrLf
        End If

        If IsDate(txtDataInicial.Text) AndAlso IsDate(txtDataFinal.Text) Then
            Condicao &= "        AND (Pesagem.Movimento  between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "') " & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(txtClientes.Text) Then
            Condicao &= "and Pesagem.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
                        "and Pesagem.EndCliente = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
        End If

        sql &= Condicao

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
         "                                                             " & vbCrLf

        If Not optPonderada.Checked Then
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
                "                          Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id AND PesagemXAnalises.Sequencia_Id = 0 LEFT OUTER JOIN" & vbCrLf & _
                "                          SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
                "                          Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
                "                          Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id LEFT OUTER JOIN" & vbCrLf & _
                "                          Clientes AS Depositante ON Pesagem.Depositante = Depositante.Cliente_Id AND Pesagem.EndDepositante = Depositante.Endereco_Id" & vbCrLf & _
                "           WHERE (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0 AND (Pesagem.EntradaSaida = 'E')" & vbCrLf

            sql &= Condicao

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
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN PesagemXAnalises.Desconto ELSE 0 END) AS ImpurezaDesconto," & vbCrLf & _
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 2 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS ImpurezaPonderada," & vbCrLf & _
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN PesagemXAnalises.Desconto ELSE 0 END) AS AvariadoDesconto," & vbCrLf & _
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 3 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END) AS AvariadoPonderada," & vbCrLf & _
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN PesagemXAnalises.Desconto ELSE 0 END) AS EsverdeadoDesconto," & vbCrLf & _
                "                           SUM(CASE WHEN PesagemXAnalises.Analise_ID = 4 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                "                           AS EsverdeadoPonderada, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                "                           AS QuebradosDesconto, SUM(CASE WHEN PesagemXAnalises.Analise_ID = 5 THEN Pesagem.BrutoBalanca * PesagemXAnalises.Percentual ELSE 0 END)" & vbCrLf & _
                "                           AS QuebradosPonderada" & vbCrLf & _
                "            FROM   Pesagem LEFT OUTER JOIN" & vbCrLf & _
                "                          PesagemXAnalises ON Pesagem.Empresa_Id = PesagemXAnalises.Empresa_Id AND Pesagem.EndEmpresa_Id = PesagemXAnalises.EndEmpresa_Id AND" & vbCrLf & _
                "                          Pesagem.Pesagem_Id = PesagemXAnalises.Pesagem_Id AND PesagemXAnalises.Sequencia_Id = 0 LEFT OUTER JOIN" & vbCrLf & _
                "                           SubOperacoes ON Pesagem.Operacao = SubOperacoes.Operacao_Id AND Pesagem.SubOperacao = SubOperacoes.SubOperacoes_Id LEFT OUTER JOIN" & vbCrLf & _
                "                           Produtos ON Pesagem.Produto = Produtos.Produto_Id LEFT OUTER JOIN" & vbCrLf & _
                "                          Clientes AS Deposito ON Pesagem.Deposito = Deposito.Cliente_Id AND Pesagem.EndDeposito = Deposito.Endereco_Id" & vbCrLf & _
                "           WHERE (Pesagem.Situacao = 1) And Pesagem.SegundaPesagem <> 0 And Pesagem.Sequencia_Id = 0 AND (Pesagem.EntradaSaida = 'E')" & vbCrLf

            sql &= Condicao
            sql &= "           GROUP BY Pesagem.Produto, Produtos.Nome, Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Pesagem.Cliente, Pesagem.EndCliente, Pesagem.Operacao," & vbCrLf & _
                "                          Pesagem.SubOperacao, SubOperacoes.Descricao, Pesagem.BrutoBalanca, Pesagem.Liquido) AS Consulta INNER JOIN" & vbCrLf & _
                "          Clientes ON Consulta.Cliente = Clientes.Cliente_Id AND Consulta.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
                " WHERE (Consulta.UmidadePonderada <> 0)" & vbCrLf & _
                " GROUP BY Consulta.Produto, Consulta.NomeProduto, Clientes.Nome, Consulta.Operacao, Consulta.SubOperacao, Consulta.NomeOperacao" & vbCrLf & _
                " ORDER BY Produto, Clientes.Nome, Operacao, Consulta.SubOperacao" & vbCrLf
        End If

        Return sql
    End Function

#End Region

#Region "Eventos"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioEntradaSaidaLaudo", "ACESSAR") Then
                BuscaUnidadeNegocio()
                BuscaDeposito()
                BuscarSituacoes()
                BuscarTransportes()
                BuscarGrupoDeProdutos()
                BuscarOperacoes()
                LiberaEmpresa()
                ddl.Carregar(ddlAnalise, CarregarDDL.Tabela.Analises)
                ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)

                txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                txtDataFinal.Text = Format(Today, "dd/MM/yyyy")

                HID.Value = Guid.NewGuid.ToString()
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaPedidos.SetarHID(HID.Value)

                If Now > "2015-12-31" Then INTACTA.Visible = True
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If

        If Not Session("txtCnpjLaudo" & HID.Value) Is Nothing Then
            Dim Cliente As New [Lib].Negocio.Cliente(Session("txtCnpjLaudo" & HID.Value), Session("txtEndLaudo" & HID.Value))
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(Cliente)

            txtClientes.Text = itemCliente.Text 'Session("txtNomeLaudo" & HID.Value) & " - " & Session("txtCidadeLaudo") & " - " & Session("txtEstadoLaudo") & " - " & Funcoes.FormatarCpfCnpj(Session("txtCnpjLaudo"))
            txtCodigoCliente.Value = itemCliente.Value 'Session("txtCnpjLaudo") & ";" & Session("txtEndLaudo")

            Session.Remove("txtCnpjLaudo" & HID.Value)
            Session.Remove("txtEndLaudo" & HID.Value)
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            cmbUnidadeNegocio.Enabled = False
            cmbEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub cmbUnidadeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdAjuda_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim NomeArquivo As String = "Manual/RelatorioEntradaSaidaLaudo.mht"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnHtml_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            Dim strm As StreamWriter = Nothing
            If Dir(arquivo).Length > 0 Then Kill(arquivo)
            If ValidarCampos() Then
                Dim dsRelatorio As New DataSet
                Dim dr As DataRow
                Dim linha As String
                Dim strSQL As String = "SELECT LP.Movimento, P.Produto_Id, P.Nome, C.Reduzido, C.Cliente_Id, " & _
                                       "C.Endereco_Id, C.Nome AS NomeCliente, C.Cidade AS ClienteCidade, " & _
                                       "C.Estado AS ClienteEstado, " & _
                                       "COALESCE(Umidade.Percentual, 0) AS PercUmidade, " & _
                                       "COALESCE(Umidade.Desconto, 0) AS DescUmidade, " & _
                                       "COALESCE(Impureza.Percentual, 0) AS PercImpureza, " & _
                                       "COALESCE(Impureza.Desconto, 0) AS DescImpureza, " & _
                                       "COALESCE(Avariados.Percentual, 0) AS PercAvariados, " & _
                                       "COALESCE(Avariados.Desconto, 0) AS DescAvariados, " & _
                                       "COALESCE(PH.Percentual, 0) AS PercPH, " & _
                                       "COALESCE(PH.Desconto, 0) AS DescPH, " & _
                                       "COALESCE(Verde.Percentual, 0) AS PercVerde, " & _
                                       "COALESCE(Verde.Desconto, 0) AS DescVerde, " & _
                                       "COALESCE(Transgenicos.Percentual, 0) AS PercTrans, " & _
                                       "COALESCE(Transgenicos.Desconto, 0) AS DescTrans, " & _
                                       "COALESCE(Outros.Percentual, 0) AS PercOutros, " & _
                                       "COALESCE(Outros.Desconto, 0) AS DescOutros, " & _
                                       "LP.Liquido, LP.Placa, LP.Pesagem_Id, LP.BrutoBalanca, " & _
                                       "LP.Empresa_Id, LP.EndEmpresa_Id, E.Nome as EmpresaNome, " & _
                                       "E.Reduzido as EmpresaReduzido, E.Cidade AS EmpresaCidade, " & _
                                       "E.Estado AS EmpresaEstado " & _
                                       "FROM Pesagem LP " & _
                                       "INNER JOIN Produtos P " & _
                                       "ON LP.Produto = P.Produto_Id " & _
                                       "INNER JOIN Clientes C " & _
                                       "ON LP.Cliente = C.Cliente_Id " & _
                                       "AND LP.EndCliente = C.Endereco_Id " & _
                                       "INNER JOIN Clientes E " & _
                                       "ON LP.Empresa_Id = E.Cliente_Id " & _
                                       "AND LP.EndEmpresa_Id = E.Endereco_Id INNER JOIN " & _
                                       "Clientes AS D ON LP.Depositante = D.Cliente_Id AND " & _
                                       "LP.EndDepositante = D.Endereco_Id " & _
                                       "LEFT JOIN vw_pesagemxanalises Umidade " & _
                                       "ON Umidade.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Umidade.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Umidade.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND Umidade.Analise_Id = 1 " & _
                                       "LEFT JOIN vw_pesagemxanalises Impureza " & _
                                       "ON Impureza.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Impureza.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Impureza.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND Impureza.Analise_Id = 2 " & _
                                       "LEFT JOIN vw_pesagemxanalises Avariados " & _
                                       "ON Avariados.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Avariados.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Avariados.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND Avariados.Analise_Id = 3 " & _
                                       "LEFT JOIN vw_pesagemxanalises PH " & _
                                       "ON PH.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND PH.Empresa_Id = LP.Empresa_Id " & _
                                       "AND PH.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND PH.Analise_Id = 4 " & _
                                       "LEFT JOIN vw_pesagemxanalises Verde " & _
                                       "ON Verde.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Verde.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Verde.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND Verde.Analise_Id = 5 " & _
                                       "LEFT JOIN vw_pesagemxanalises Transgenicos " & _
                                       "ON Transgenicos.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Transgenicos.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Transgenicos.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "AND Transgenicos.Analise_Id = 6 " & _
                                       "LEFT JOIN (SELECT 0.00 as Percentual, SUM(Desconto) as Desconto, " & _
                                       "Pesagem_Id, Empresa_Id, EndEmpresa_Id " & _
                                       "FROM vw_pesagemxanalises " & _
                                       "WHERE Analise_Id > 6 " & _
                                       "GROUP BY Pesagem_Id, Empresa_Id, EndEmpresa_Id) Outros " & _
                                       "ON Outros.Pesagem_Id = LP.Pesagem_Id " & _
                                       "AND Outros.Empresa_Id = LP.Empresa_Id " & _
                                       "AND Outros.EndEmpresa_Id = LP.EndEmpresa_Id " & _
                                       "WHERE LP.Movimento BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' " & _
                                       "AND '" & txtDataFinal.Text.ToSqlDate() & "' " & _
                                       "AND LP.Sequencia_Id = 0 " & _
                                       "AND (LP.Liquido <> 0 AND LP.BrutoBalanca<>0) "

                If cmbEmpresa.SelectedIndex > 0 Then
                    Dim strEmpresa As String()
                    strEmpresa = cmbEmpresa.SelectedValue.Split(New Char() {"-"})

                    strSQL &= "AND LP.Empresa_Id = '" & strEmpresa(0) & "' " & _
                              "AND LP.EndEmpresa_Id = " & strEmpresa(1) & " "
                End If

                If txtCodigoCliente.Value.Length > 0 Then
                    Dim strCliente As String()
                    strCliente = txtCodigoCliente.Value.Split(New Char() {"-"})
                    strSQL &= "AND LP.Cliente = '" & strCliente(0) & "' " & _
                              "AND LP.EndCliente = " & strCliente(1) & " "
                End If

                If txtCodigoDepositante.Value.Length > 0 Then
                    Dim strDepositante As String()
                    strDepositante = txtCodigoDepositante.Value.Split(New Char() {";"})
                    strSQL &= "AND LP.Depositante = '" & strDepositante(0) & "' " & _
                              "AND LP.EndDepositante = " & strDepositante(1) & " "
                End If

                If DdlGrupoDeProdutos.SelectedIndex > 0 Then
                    strSQL &= " AND (P.Grupo = '" & DdlGrupoDeProdutos.SelectedValue & "') "
                End If
                If cmbProdutos.SelectedIndex > 0 Then
                    strSQL &= " AND (P.Produto_Id = '" & cmbProdutos.SelectedValue & "') "
                End If

                If optEntrada.Checked = True Then
                    strSQL &= "AND LP.EntradaSaida = 'E' "
                ElseIf optSaida.Checked = True Then
                    strSQL &= "AND LP.EntradaSaida = 'S' "
                End If

                If optSim.Checked = True Then
                    strSQL &= "AND LP.Fisico = 'S' "
                ElseIf optNao.Checked = True Then
                    strSQL &= "AND LP.Fisico = 'N' "
                End If

                If optAberto.Checked = True Then
                    strSQL &= "AND LP.SegundaPesagem = 0 "
                ElseIf optFechado.Checked = True Then
                    strSQL &= "AND LP.SegundaPesagem > 0 "
                End If

                If cmbSituacao.SelectedIndex > -1 Then strSQL &= "AND LP.Situacao = " & cmbSituacao.SelectedValue & " "

                dsRelatorio = Banco.ConsultaDataSet(strSQL, "Pesagem")

                linha = "<HTML>" & vbCrLf
                '<HEAD>
                linha &= "<HEAD>" & vbCrLf
                linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
                linha &= "<TITLE>Posiçao de Cargas e Descargas</TITLE>" & vbCrLf
                linha &= "</HEAD>" & vbCrLf

                '<BODY>
                linha &= "<BODY>" & vbCrLf

                '-----------------
                'Cabeçalho Padrao
                '-----------------
                linha &= "<table width= '3700' cellpadding='0' cellspacing='0' Border=1>"

                linha &= "<TR>"
                linha &= "<TD >EmpresaReduzido</TD>"
                linha &= "<TD >Empresa</TD>"
                linha &= "<TD >EmpresaNome</TD>"
                linha &= "<TD >EmpresaCidade</TD>"
                linha &= "<TD >EmpresaEstado</TD>"

                linha &= "<TD>Movimento</TD>"
                linha &= "<TD >Laudo</TD>"

                linha &= "<TD>Produto_Id</TD>"
                linha &= "<TD>NomeProduto</TD>"

                linha &= "<TD >Cliente</TD>"
                linha &= "<TD >NomeCliente</TD>"
                linha &= "<TD>ClienteCidade</TD>"
                linha &= "<TD>ClienteEstado</TD>"

                linha &= "<TD >BrutoBalanca</TD>"
                linha &= "<TD>PercUmidade</TD>"
                linha &= "<TD>DescUmidade</TD>"

                linha &= "<TD>PercImpureza</TD>"
                linha &= "<TD>DescImpureza</TD>"

                linha &= "<TD>PercAvariados</TD>"
                linha &= "<TD>DescAvariados</TD>"

                linha &= "<TD>PercPH</TD>"
                linha &= "<TD >DescPH</TD>"

                linha &= "<TD >PercVerde</TD>"
                linha &= "<TD >DescVerde</TD>"

                linha &= "<TD >PercTrans</TD>"
                linha &= "<TD >DescTrans</TD>"

                linha &= "<TD >PercOutros</TD>"
                linha &= "<TD >DescOutros</TD>"

                linha &= "<TD >LiquidoDoLaudo</TD>"

                linha &= "</TR>"

                If dsRelatorio.Tables(0).Rows.Count > 0 Then
                    For Each dr In dsRelatorio.Tables(0).Rows

                        linha &= "<TR>"
                        linha &= "<TD>" & dr("EmpresaReduzido") & "</TD>"
                        linha &= "<TD>" & dr("Empresa_Id") & "</TD>"
                        linha &= "<TD>" & dr("EmpresaNome") & "</TD>"
                        linha &= "<TD>" & dr("EmpresaCidade") & "</TD>"
                        linha &= "<TD>" & dr("EmpresaEstado") & "</TD>"

                        linha &= "<TD>" & dr("Movimento") & "</TD>"
                        linha &= "<TD>" & dr("Pesagem_Id") & "</TD>"

                        linha &= "<TD>" & dr("Produto_Id") & "</TD>"
                        linha &= "<TD>" & dr("Nome") & "</TD>"

                        linha &= "<TD>" & dr("Cliente_Id") & "</TD>"
                        linha &= "<TD>" & dr("NomeCliente") & "</TD>"
                        linha &= "<TD>" & dr("ClienteCidade") & "</TD>"
                        linha &= "<TD>" & dr("ClienteEstado") & "</TD>"

                        linha &= "<TD>" & dr("BrutoBalanca") & "</TD>"

                        linha &= "<TD>" & dr("PercUmidade") & "</TD>"
                        linha &= "<TD>" & dr("DescUmidade") & "</TD>"

                        linha &= "<TD>" & dr("PercImpureza") & "</TD>"
                        linha &= "<TD>" & dr("DescImpureza") & "</TD>"

                        linha &= "<TD>" & dr("PercAvariados") & "</TD>"
                        linha &= "<TD>" & dr("DescAvariados") & "</TD>"

                        linha &= "<TD>" & dr("PercPH") & "</TD>"
                        linha &= "<TD>" & dr("DescPH") & "</TD>"

                        linha &= "<TD>" & dr("PercVerde") & "</TD>"
                        linha &= "<TD>" & dr("DescVerde") & "</TD>"
                        linha &= "<TD>" & dr("PercTrans") & "</TD>"
                        linha &= "<TD>" & dr("DescTrans") & "</TD>"

                        linha &= "<TD>" & dr("PercOutros") & "</TD>"
                        linha &= "<TD>" & dr("DescOutros") & "</TD>"

                        linha &= "<TD>" & (dr("BrutoBalanca") - dr("DescUmidade") - dr("DescImpureza") - dr("DescAvariados") - dr("DescPH") - dr("DescVerde") - dr("DescOutros")) & "</TD>"

                        linha &= "</TR>"

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
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupoDeProdutos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            cmbProdutos.Items.Clear()
            BuscarProdutos()
            cmbProdutos.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmbOperacao_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BuscarSubOperacoes(cmbOperacao.SelectedValue)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaCliente.Click
        Try
            txtDepositante.Text = ""
            txtCodigoDepositante.Value = ""

            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteESL" & HID.Value.ToString(), "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdConsultaDepositante_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdConsultaDepositante.Click
        Try
            ChkDepositante.Checked = False

            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteDESL" & HID.Value, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdBuscaPedido.Click
        Try
            If cmbEmpresa.SelectedIndex = -1 Or txtCodigoCliente.Value.Length = 0 Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido")
            Else
                HttpContext.Current.Session("ssCampo" & HID.Value) = "ApenasNumeroPedido"
                Dim strCodigo As String() = cmbEmpresa.SelectedValue.Split("-")
                Dim strCodCliente As String() = txtCodigoCliente.Value.Split("-")
                Dim strEmpresa() As String = cmbEmpresa.SelectedValue.ToString.Split("-")
                Dim strCliente() As String = txtCodigoCliente.Value.Split("-")
                Dim parameters As New Dictionary(Of String, Object)
                parameters("unidade") = cmbUnidadeNegocio.SelectedValue
                parameters("empresa") = strEmpresa(0)
                parameters("enderecoEmpresa") = strEmpresa(1)
                parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
                parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
                Popup.ConsultaDePedidos(Me.Page, "objPedidosRESL" & HID.Value)
                ucConsultaPedidos.BindGridView(parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ChkDepositante_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ChkDepositante.Checked Then
                ckAgruparDeposito.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ckAgruparDeposito_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ckAgruparDeposito.Checked Then
                ChkDepositante.Checked = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
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

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean)
        Try
            If ValidarCampos() Then
                'If Pdf Then
                ImprimirRelatorio(Pdf)
                'Else
                '    EmiteExcel()
                'End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub EmiteExcel()
        Try
            Dim Empresa As String = ""
            Dim Deposito As String = ""
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

            Dim ds = New DataSet()
            ds = Banco.ConsultaDataSet(GetSql(), "Pesagem")
            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Dim ColumName() As String = {"Produto", "Cliente", "Operacao", "PesoBruto", "PerUmidade", "DesUmidade", "PerImpureza", "DesImpureza", "PerAvariado", "DesAvariado", "PerEsverdeado", "DesEsveredeado", "PerQuebrado", "DesQuebrado", "Liquido"}
            Dim i As Integer = 1
            Dim ii As Integer = 1
            Dim nomeArquivo As String = ""

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha do excel
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("POSIÇÃO DE PEDIDOS")

                    With worksheet
                        If ds.Tables(0).Columns.Count > 0 Then
                            For Each dd As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(1, i).Value = dd.ColumnName
                                i += 1
                            Next
                            worksheet.Cells("A1:AI1").Style.Font.Bold = True
                            worksheet.Cells("A1:AI1").Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells("A1:AI1").Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            worksheet.Cells("A1:AI1").Style.Font.Color.SetColor(Color.White)
                        End If
                        i = 2

                        If ds.Tables(0).Rows.Count > 0 Then
                            For Each dr As DataRow In ds.Tables(0).Rows
                                ii = 1
                                For Each dd As DataColumn In ds.Tables(0).Columns
                                    worksheet.Cells(i, ii).Value = dr(ii - 1)
                                    ii += 1
                                Next

                                'FORMATAÇÕES
                                worksheet.Cells("J" & i).Style.Numberformat.Format = "m/d/yyyy"
                                worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("P" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("Q" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("R" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("S" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("T" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("U" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("V" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("W" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("X" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("Y" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("Z" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AA" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AB" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("AC" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AD" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AE" & i).Style.Numberformat.Format = "#,##0"

                                worksheet.Cells("AF" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AG" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("AH" & i).Style.Numberformat.Format = "#,##0"

                                If i Mod 2 = 0 Then
                                    worksheet.Cells("A" & i & ":AI" & i).Style.Fill.PatternType = ExcelFillStyle.Solid
                                    worksheet.Cells("A" & i & ":AI" & i).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End If
                                i += 1
                            Next
                        End If

                        i = i + 2
                        ii = 1

                        If ds.Tables(1).Columns.Count > 0 Then
                            'COLUNAS
                            Dim c As Integer = ColumName.Length
                            While ii <= c
                                worksheet.Cells(i, ii).Value = ColumName(ii - 1)
                                ii += 1
                            End While
                            worksheet.Cells("A" & i & ":O" & i).Style.Font.Bold = True
                            worksheet.Cells("A" & i & ":O" & i).Style.Fill.PatternType = ExcelFillStyle.Solid
                            worksheet.Cells("A" & i & ":O" & i).Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                            i += 1
                        End If

                        If ds.Tables(1).Rows.Count > 0 Then
                            For Each rs As DataRow In ds.Tables(1).Rows
                                If rs("Produto") <> Produto And Produto = "" Then
                                    worksheet.Cells("A" & i).Value = rs("Produto")
                                    worksheet.Cells("A" & i).Style.Font.Bold = True
                                    Produto = rs("Produto")
                                    i = i + 1
                                End If

                                If rs("Produto") <> Produto Then
                                    worksheet.Cells("C" & i).Value = "SUB TOTAL DO CLIENTE.."
                                    worksheet.Cells("D" & i).Value = BrutoBalanca

                                    If UmidadePonderada > 0 Then
                                        worksheet.Cells("E" & i).Value = UmidadePonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("F" & i).Value = UmidadeDesconto

                                    If ImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & i).Value = ImpurezaPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("H" & i).Value = ImpurezaDesconto

                                    If AvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & i).Value = AvariadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("J" & i).Value = AvariadoDesconto

                                    If EsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & i).Value = EsverdeadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("L" & i).Value = EsverdeadoDesconto

                                    If QuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & i).Value = QuebradoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("N" & i).Value = QuebradoDesconto

                                    worksheet.Cells("O" & i).Value = Liquido
                                    worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True
                                    Cliente = rs("Cliente")

                                    'TOTALIZA PRODUTO
                                    i = i + 1
                                    worksheet.Cells("C" & i).Value = "SUB TOTAL DO PRODUTO..."
                                    worksheet.Cells("D" & i).Value = TPBrutoBalanca
                                    If TPUmidadePonderada > 0 Then
                                        worksheet.Cells("E" & i).Value = TPUmidadePonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("F" & i).Value = TPUmidadeDesconto

                                    If TPImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & i).Value = TPImpurezaPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("H" & i).Value = TPImpurezaDesconto

                                    If TPAvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & i).Value = TPAvariadoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("J" & i).Value = TPAvariadoDesconto

                                    If TPEsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & i).Value = TPEsverdeadoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("L" & i).Value = TPEsverdeadoDesconto

                                    If TPQuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & i).Value = TPQuebradoPonderada / TPBrutoBalanca
                                    End If
                                    worksheet.Cells("N" & i).Value = TPQuebradoDesconto

                                    worksheet.Cells("O" & i).Value = TPLiquido
                                    worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True

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

                                    i = i + 2
                                    worksheet.Cells("A" & i).Value = rs("Produto")
                                    Produto = rs("Produto")
                                    worksheet.Cells("A" & i).Style.Font.Bold = True
                                    i = i + 1
                                End If

                                If Cliente = "" Then
                                    Cliente = rs("Cliente")
                                End If

                                If rs("Cliente") <> Cliente Then
                                    worksheet.Cells("C" & i).Value = "SUB TOTAL DO CLIENTE..."
                                    worksheet.Cells("D" & i).Value = BrutoBalanca

                                    If UmidadePonderada > 0 Then
                                        worksheet.Cells("E" & i).Value = UmidadePonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("F" & i).Value = UmidadeDesconto

                                    If ImpurezaPonderada > 0 Then
                                        worksheet.Cells("G" & i).Value = ImpurezaPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("H" & i).Value = ImpurezaDesconto

                                    If AvariadoPonderada > 0 Then
                                        worksheet.Cells("I" & i).Value = AvariadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("J" & i).Value = AvariadoDesconto

                                    If EsverdeadoPonderada > 0 Then
                                        worksheet.Cells("K" & i).Value = EsverdeadoPonderada / BrutoBalanca
                                    End If
                                    worksheet.Cells("L" & i).Value = EsverdeadoDesconto

                                    If QuebradoPonderada > 0 Then
                                        worksheet.Cells("M" & i).Value = QuebradoPonderada / BrutoBalanca
                                    End If

                                    worksheet.Cells("N" & i).Value = QuebradoDesconto
                                    worksheet.Cells("O" & i).Value = Liquido

                                    worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                                    worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                                    worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True

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
                                    Cliente = rs("Cliente")
                                    i = i + 2
                                End If

                                worksheet.Cells("B" & i).Value = rs("Cliente")
                                worksheet.Cells("B" & i).Style.Font.Bold = True
                                worksheet.Cells("C" & i).Value = rs("Operacao")
                                worksheet.Cells("D" & i).Value = rs("BrutoBalanca")

                                If rs("UmidadePonderada") > 0 Then
                                    worksheet.Cells("E" & i).Value = rs("UmidadePonderada") / rs("BrutoBalanca")
                                End If
                                worksheet.Cells("F" & i).Value = rs("UmidadeDesconto")

                                If rs("ImpurezaPonderada") > 0 Then
                                    worksheet.Cells("G" & i).Value = rs("ImpurezaPonderada") / rs("BrutoBalanca")
                                End If
                                worksheet.Cells("H" & i).Value = rs("ImpurezaDesconto")

                                If rs("AvariadoPonderada") Then
                                    worksheet.Cells("I" & i).Value = rs("AvariadoPonderada") / rs("BrutoBalanca")
                                End If
                                worksheet.Cells("J" & i).Value = rs("AvariadoDesconto")

                                If rs("EsverdeadoPonderada") > 0 Then
                                    worksheet.Cells("K" & i).Value = rs("EsverdeadoPonderada") / rs("BrutoBalanca")
                                End If
                                worksheet.Cells("L" & i).Value = rs("EsverdeadoDesconto")

                                If rs("QuebradoPonderada") > 0 Then
                                    worksheet.Cells("M" & i).Value = rs("QuebradoPonderada") / rs("BrutoBalanca")
                                End If

                                worksheet.Cells("N" & i).Value = rs("QuebradoDesconto")
                                worksheet.Cells("O" & i).Value = rs("Liquido")

                                worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                                worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                                worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"

                                BrutoBalanca = BrutoBalanca + rs("BrutoBalanca")
                                UmidadePonderada = UmidadePonderada + rs("UmidadePonderada")
                                UmidadeDesconto = UmidadeDesconto + rs("UmidadeDesconto")
                                ImpurezaPonderada = ImpurezaPonderada + rs("ImpurezaPonderada")
                                ImpurezaDesconto = ImpurezaDesconto + rs("ImpurezaDesconto")
                                AvariadoPonderada = AvariadoPonderada + rs("AvariadoPonderada")
                                AvariadoDesconto = AvariadoDesconto + rs("AvariadoDesconto")
                                EsverdeadoPonderada = EsverdeadoPonderada + rs("EsverdeadoPonderada")
                                EsverdeadoDesconto = EsverdeadoDesconto + rs("EsverdeadoDesconto")
                                EsverdeadoPonderada = EsverdeadoPonderada + rs("EsverdeadoPonderada")
                                EsverdeadoDesconto = EsverdeadoDesconto + rs("EsverdeadoDesconto")
                                Liquido = Liquido + rs("Liquido")

                                TPBrutoBalanca = TPBrutoBalanca + rs("BrutoBalanca")
                                TPUmidadePonderada = TPUmidadePonderada + rs("UmidadePonderada")
                                TPUmidadeDesconto = TPUmidadeDesconto + rs("UmidadeDesconto")
                                TPImpurezaPonderada = TPImpurezaPonderada + rs("ImpurezaPonderada")
                                TPImpurezaDesconto = TPImpurezaDesconto + rs("ImpurezaDesconto")
                                TPAvariadoPonderada = TPAvariadoPonderada + rs("AvariadoPonderada")
                                TPAvariadoDesconto = TPAvariadoDesconto + rs("AvariadoDesconto")
                                TPEsverdeadoPonderada = TPEsverdeadoPonderada + rs("EsverdeadoPonderada")
                                TPEsverdeadoDesconto = TPEsverdeadoDesconto + rs("EsverdeadoDesconto")
                                TPEsverdeadoPonderada = TPEsverdeadoPonderada + rs("EsverdeadoPonderada")
                                TPEsverdeadoDesconto = TPEsverdeadoDesconto + rs("EsverdeadoDesconto")
                                TPLiquido = TPLiquido + rs("Liquido")

                                TGBrutoBalanca = TGBrutoBalanca + rs("BrutoBalanca")
                                TGUmidadePonderada = TGUmidadePonderada + rs("UmidadePonderada")
                                TGUmidadeDesconto = TGUmidadeDesconto + rs("UmidadeDesconto")
                                TGImpurezaPonderada = TGImpurezaPonderada + rs("ImpurezaPonderada")
                                TGImpurezaDesconto = TGImpurezaDesconto + rs("ImpurezaDesconto")
                                TGAvariadoPonderada = TGAvariadoPonderada + rs("AvariadoPonderada")
                                TGAvariadoDesconto = TGAvariadoDesconto + rs("AvariadoDesconto")
                                TGEsverdeadoPonderada = TGEsverdeadoPonderada + rs("EsverdeadoPonderada")
                                TGEsverdeadoDesconto = TGEsverdeadoDesconto + rs("EsverdeadoDesconto")
                                TGEsverdeadoPonderada = TGEsverdeadoPonderada + rs("EsverdeadoPonderada")
                                TGEsverdeadoDesconto = TGEsverdeadoDesconto + rs("EsverdeadoDesconto")
                                TGLiquido = TGLiquido + rs("Liquido")

                                i = i + 1
                                ii = ii + 1
                            Next
                        End If

                        worksheet.Cells("C" & i).Value = "SUB TOTAL DO CLIENTE..."
                        worksheet.Cells("D" & i).Value = BrutoBalanca

                        If UmidadePonderada > 0 Then
                            worksheet.Cells("E" & i).Value = UmidadePonderada / BrutoBalanca
                        End If
                        worksheet.Cells("F" & i).Value = UmidadeDesconto

                        If ImpurezaPonderada > 0 Then
                            worksheet.Cells("G" & i).Value = ImpurezaPonderada / BrutoBalanca
                        End If
                        worksheet.Cells("H" & i).Value = ImpurezaDesconto

                        If AvariadoPonderada > 0 Then
                            worksheet.Cells("I" & i).Value = AvariadoPonderada / BrutoBalanca
                        End If
                        worksheet.Cells("J" & i).Value = AvariadoDesconto

                        If EsverdeadoPonderada > 0 Then
                            worksheet.Cells("K" & i).Value = EsverdeadoPonderada / BrutoBalanca
                        End If
                        worksheet.Cells("L" & i).Value = EsverdeadoDesconto

                        If QuebradoPonderada > 0 Then
                            worksheet.Cells("M" & i).Value = QuebradoPonderada / BrutoBalanca
                        End If

                        worksheet.Cells("N" & i).Value = QuebradoDesconto
                        worksheet.Cells("O" & i).Value = Liquido

                        worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True
                        i = i + 1

                        'TOTALIZA PRODUTO
                        worksheet.Cells("C" & i).Value = "SUB TOTAL DO PRODUTO..."
                        worksheet.Cells("D" & i).Value = TPBrutoBalanca

                        If TPUmidadePonderada > 0 Then
                            worksheet.Cells("E" & i).Value = TPUmidadePonderada / TPBrutoBalanca
                        End If
                        worksheet.Cells("F" & i).Value = TPUmidadeDesconto

                        If TPImpurezaPonderada > 0 Then
                            worksheet.Cells("G" & i).Value = TPImpurezaPonderada / TPBrutoBalanca
                        End If
                        worksheet.Cells("H" & i).Value = TPImpurezaDesconto

                        If TPAvariadoPonderada > 0 Then
                            worksheet.Cells("I" & i).Value = TPAvariadoPonderada / TPBrutoBalanca
                        End If
                        worksheet.Cells("J" & i).Value = TPAvariadoDesconto

                        If TPEsverdeadoPonderada > 0 Then
                            worksheet.Cells("K" & i).Value = TPEsverdeadoPonderada / TPBrutoBalanca
                        End If
                        worksheet.Cells("L" & i).Value = TPEsverdeadoDesconto

                        If TPQuebradoPonderada > 0 Then
                            worksheet.Cells("M" & i).Value = TPQuebradoPonderada / TPBrutoBalanca
                        End If

                        worksheet.Cells("N" & i).Value = TPQuebradoDesconto
                        worksheet.Cells("O" & i).Value = TPLiquido

                        worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True
                        i = i + 1

                        'TOTAL GERAL
                        worksheet.Cells("C" & i).Value = "TOTAL GERAL..."
                        worksheet.Cells("D" & i).Value = TGBrutoBalanca

                        If TGUmidadePonderada > 0 Then
                            worksheet.Cells("E" & i).Value = TGUmidadePonderada / TGBrutoBalanca
                        End If
                        worksheet.Cells("F" & i).Value = TGUmidadeDesconto

                        If TGImpurezaPonderada > 0 Then
                            worksheet.Cells("G" & i).Value = TGImpurezaPonderada / TGBrutoBalanca
                        End If
                        worksheet.Cells("H" & i).Value = TGImpurezaDesconto

                        If TGAvariadoPonderada > 0 Then
                            worksheet.Cells("I" & i).Value = TGAvariadoPonderada / TGBrutoBalanca
                        End If
                        worksheet.Cells("J" & i).Value = TGAvariadoDesconto

                        If TGEsverdeadoPonderada > 0 Then
                            worksheet.Cells("K" & i).Value = TGEsverdeadoPonderada / TGBrutoBalanca
                        End If
                        worksheet.Cells("L" & i).Value = TGEsverdeadoDesconto

                        If TGQuebradoPonderada > 0 Then
                            worksheet.Cells("M" & i).Value = TGQuebradoPonderada / TGBrutoBalanca
                        End If

                        worksheet.Cells("N" & i).Value = TGQuebradoDesconto

                        worksheet.Cells("O" & i).Value = TGLiquido

                        worksheet.Cells("D" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("E" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("F" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("G" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("H" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("I" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("J" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("K" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("L" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("M" & i).Style.Numberformat.Format = "#,##0.00"
                        worksheet.Cells("N" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("O" & i).Style.Numberformat.Format = "#,##0"
                        worksheet.Cells("C" & i & ":O" & i).Style.Font.Bold = True
                    End With

                    'criando auto filtro na planilha
                    worksheet.Cells("A1:AI" & i).AutoFilter = True

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando primeira linha
                    worksheet.View.FreezePanes(2, 1)

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        cmbUnidadeNegocio.SelectedIndex = 0
        cmbDeposito.SelectedIndex = 0
        cmbEmpresa.Items.Clear()
        txtClientes.Text = ""
        txtCodigoCliente.Value = ""
        txtDepositante.Text = ""
        txtCodigoDepositante.Value = ""
        cmbOperacao.SelectedIndex = 0
        cmbSubOperacao.Items.Clear()
        DdlGrupoDeProdutos.SelectedIndex = 0
        cmbProdutos.Items.Clear()
        txtDataInicial.Text = DateTime.Now.ToString("dd/MM/yyyy")
        txtDataFinal.Text = DateTime.Now.ToString("dd/MM/yyyy")
        cmbSituacao.SelectedIndex = 0
        optPonderada.Checked = True
        optFechado.Checked = True
        optGeral.Checked = True
        ChkDepositante.Checked = False
        chkClientes.Checked = False
        txtPedido.Text = ""
        HID.Value = Guid.NewGuid.ToString()
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)
    End Sub

#End Region

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioEntradaSaidaLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class