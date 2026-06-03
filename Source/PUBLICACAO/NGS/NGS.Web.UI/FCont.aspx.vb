Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Web
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class FCont
    Inherits BasePage

    Dim sqlAux As String
    Dim ds As New DataSet
    Dim dtInicialOriginal, dtFinalOriginal As Date
    Dim Conv, NatOp, ArqMag As String
    Dim Empresa As Array
    Dim Ano As Integer = 0
    Dim Dia As Integer = 0
    Dim Mes As Integer = 0
    Dim Ciclo As Integer = 0
    Dim dias As String = ""
    Dim DataInicial As String = ""
    Dim DataVirada As String = ""
    Dim DataFinal As String = ""
    Dim Saldo As Decimal = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("FCont", "ACESSAR") Then

                If Not Directory.Exists("C:\Sped\FCont\" & Today().Year() - 1) Then
                    Directory.CreateDirectory("C:\Sped\FCont\" & Today().Year() - 1)
                End If

                CargaEmpresas()
                Funcoes.VerificaEmpresa(ddlEmpresa)

                Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

                ddlTipoDeEscrituracao.SelectedIndex = 1
                ddlSituacaoSdoDaEscrituracao.SelectedIndex = 1
                ddlSituacaoSdoDaEscrituracao.SelectedIndex = 1
                ddlFormaDeApuracao.SelectedIndex = 1
                ddlFormaDeTributacao.SelectedIndex = 1

                ddlTrimestre_01.SelectedIndex = 1
                ddlTrimestre_02.SelectedIndex = 1
                ddlTrimestre_03.SelectedIndex = 1
                ddlTrimestre_04.SelectedIndex = 1

                ddlApuracaoDoTrimestre_01.SelectedIndex = 2
                ddlApuracaoDoTrimestre_02.SelectedIndex = 2
                ddlApuracaoDoTrimestre_03.SelectedIndex = 2
                ddlApuracaoDoTrimestre_04.SelectedIndex = 2

                Limpar()

            Else
                MsgBox(Me.Page, "Usuário sem permissão acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If

        SalvarSessaoCarregaObjetos()
    End Sub

    Private Sub SalvarSessaoCarregaObjetos()
        '************************************************************************************
        '****************************   XXXXX   ****************************************
        '************************************************************************************

    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    'Private Sub CargaSituacao()
    '    Dim objSituacao As New NGS.[Lib].Negocio.ListSituacao(True)

    '    ddlSituacao.Items.Clear()

    '    For Each row As NGS.Negocio.Situacao In objSituacao
    '        ddlSituacao.Items.Add(New ListItem(Funcoes.AlinharDireita(row.Codigo, 3, "0") & " - " & row.Descricao, row.Codigo))
    '    Next

    '    Funcoes.InserirLinhaEmBranco(ddlSituacao)
    'End Sub


    'Private Sub CargaCarteiras()
    '    Dim objCarteira As New NGS.Negocio.ListCarteiraFinanceira(True, "P")

    '    ddlCarteira.Items.Clear()

    '    For Each row As NGS.[Lib].Negocio.CarteiraFinanceira In objCarteira
    '        ddlCarteira.Items.Add(New ListItem(Funcoes.AlinharDireita(row.CodigoCarteira, 3, "0") & " - " & row.Descricao, row.CodigoCarteira))
    '    Next

    '    Funcoes.InserirLinhaEmBranco(ddlCarteira)
    'End Sub

    Private Sub Limpar()
        txtDataInicial.Text = "01-" & Format(Date.Now, "MM-yyyy")
        txtDataFinal.Text = Format(Date.Now, "dd-MM-yyyy")
        txtContaDeResultado.Text = "204050103"
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        If ddlEmpresa.SelectedIndex > 0 Then
            Empresa = ddlEmpresa.SelectedValue.Split("-")
            txtArquivodeSaida.Text = "C:\Sped\FCont\" & Today().Year() - 1 & "\" & Microsoft.VisualBasic.Left(Empresa(0), 8) & ".txt"
        End If
    End Sub

    Protected Sub cmdArquivoDeSaida_Click(ByVal sender As Object, ByVal e As EventArgs) Handles cmdArquivoDeSaida.Click
        Download(txtArquivodeSaida.Text)
    End Sub

    Private Sub Download(ByVal arquivo As String)
        Try
            Response.ContentType = "text/plain"
            Response.AppendHeader("Content-Disposition", "attachment; filename=" & arquivo.Split("\").ToArray(4))
            Const bufferLength As Integer = 10000
            Dim buffer As Byte() = New Byte(bufferLength - 1) {}
            Dim length As Integer = 0
            Dim download As Stream = Nothing
            Try
                download = New FileStream(txtArquivodeSaida.Text, FileMode.Open, FileAccess.Read)
                Do
                    If Response.IsClientConnected Then
                        length = download.Read(buffer, 0, bufferLength)
                        Response.OutputStream.Write(buffer, 0, length)
                        buffer = New Byte(bufferLength - 1) {}
                    Else
                        length = -1
                    End If
                Loop While length > 0
                Response.Flush()
                Response.End()
            Finally
                If download IsNot Nothing Then
                    download.Close()
                End If
            End Try
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Dim strm As StreamWriter = Nothing
        Try
            If Funcoes.VerificaPermissao("FCont", "RELATORIO") Then
                Dim sql As String
                Dim arquivo As String
                Dim linha As String
                'Dim strm As StreamWriter
                'Dim IntNota As Integer
                'Dim IntSequencia As Integer

                Dim Registro0000 As Integer = 0     'Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária

                Dim RegistroI001 As Integer = 0     'Abertura do Broco I
                Dim RegistroI050 As Integer = 0     'Plano de Contas
                Dim RegistroI051 As Integer = 0     'Plano de Contas Referencial
                Dim RegistroI150 As Integer = 0     'Saldos Periódicos - Identificação do Periodo
                Dim RegistroI155 As Integer = 0     'Detalhe dos Saldos Periódicos
                Dim RegistroI200 As Integer = 0     'Lancamento Contábil
                Dim RegistroI250 As Integer = 0     'Partidas do Lançamento
                Dim RegistroI350 As Integer = 0     'Saldo das Contas de Resultado Antes do Encerramento - Identificação da Data
                Dim RegistroI355 As Integer = 0     'Detalhes dos Saldos das Contas de Resultado Antes do Encerramento
                Dim RegistroI990 As Integer = 0     'Encerramento do Bloco I

                Dim RegistroJ001 As Integer = 0     'Abertura do Bloco J
                Dim RegistroJ930 As Integer = 0     'Identificação dos Signatários da Escrituração
                Dim RegistroJ990 As Integer = 0     'Encerramento do Bloco J

                Dim RegistroM001 As Integer = 0     'Abertura do Bloco M
                Dim RegistroM020 As Integer = 0     'Qualificação da Pessoa Juridica e Retificação
                Dim RegistroM025 As Integer = 0     'Saldos Iniciais das Contas Patrimoniais Recuperados/Preenchidos
                Dim RegistroM030 As Integer = 0     'Identificação do Periodo de Apuração do Lucro Real
                Dim RegistroM155 As Integer = 0     'Detalhes dos Saldos Referenciais das Contas Patrimoniais
                Dim RegistroM990 As Integer = 0     'Encerramento do Bloco M


                Dim Registro9001 As Integer = 0     'Abertura do Bloco 9
                Dim Registro9900 As Integer = 0     'Registros do Arquivo
                Dim Registro9990 As Integer = 0     'Encerramento do Bloco 9
                Dim Registro9999 As Integer = 0     'Encerramento do Arquivo Digital

                Dim RegistroGeral As Integer = 0    'Registro Geral

                Dim CGC As String = ""
                Dim Inscricao As String = ""

                Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")

                'Me.Cursor = Cursors.WaitCursor

                Dim PreArquivo As String = "C:\Sped\FCont\" & Today().Year() - 1 & "\" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & ".txt"
                arquivo = PreArquivo
                'Server.MapPath(PreArquivo)

                If Dir(arquivo).Length > 0 Then Kill(arquivo)

                Dim ds As New DataSet
                Dim Array As New ArrayList

                strm = New StreamWriter(arquivo, True)

                'Registro 0000 - Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
                '----------------------------------------------------------------------------------------------------------

                sql = "  SELECT    Clientes.*, ClientesXEmpresas.*"
                sql &= " FROM      Clientes INNER JOIN"
                sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
                sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
                sql &= " WHERE     Clientes.Cliente_Id = '" & strEmpresa(0) & "' And Clientes.Endereco_Id = 0"

                ds = Banco.ConsultaDataSet(sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|0000"                                                                                                'Tipo
                            linha &= "|LALU"
                            linha &= "|" & txtDataInicial.Text.ToStrDate()                                                     'Data Inicial
                            linha &= "|" & txtDataFinal.Text.ToStrDate()                                                     'Data Final
                            linha &= "|" & RTrim(.Item("Nome"))
                            linha &= "|" & Microsoft.VisualBasic.Left(.Item("Cliente_Id"), 14)
                            linha &= "|" & .Item("Estado")                                                                                 'Estado
                            linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
                            linha &= "|" & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 7, "0")
                            linha &= "|" & .Item("InscricaoMunicipal")
                            linha &= "|" '& ddlTipoDeEscrituracao.SelectedIndex
                            linha &= "|0"                                                              'Indicador de Inicio de Periodo 
                            linha &= "|"                                                               ' 0 - Inicio no Primeiro Dia Do Ano
                        End With
                        ' 1 - Abertura
                        ' 2 - Resultante de cisão/fusão ou remanencente de cisão ou realizou incorporação
                        ' 3 - Inicio da Obrigatoriedade da entrega da FCONT no curso do ano-calendário

                        strm.WriteLine(linha)
                        Registro0000 += 1
                        RegistroGeral += 1
                    Next
                End If

                ' Registro I001  - Abertura do Bloco I
                '----------------------------------

                linha = "|I001"
                linha &= "|0" 'Indicador de Movimento
                linha &= "|"

                strm.WriteLine(linha)
                RegistroI001 += 1
                RegistroGeral += 1
                Ano = Year(txtDataInicial.Text)
                Mes = Month(txtDataFinal.Text)

                DataInicial = "01/01" & " / " & Ano
                DataFinal = "31/12/" & Ano

                'Registro I050 - Plano de Contas
                '----------------------------------------------------------------------------------------------------------

                sql = "  SELECT  Empresa,   Consulta.Conta as Conta_ID , isnull ((select top 1 Titulo from planodecontas where planodecontas.conta_Id = consulta.conta), 'Conta Nao Encontrada') as Titulo, Sum(Consulta.Inicial) as Inicial, Sum(Consulta.Debitos) As Debitos, Sum(Consulta.Creditos) as Creditos "
                sql &= "   into  #PL1  FROM         (SELECT  Empresa,   Consulta_4.Conta, Consulta_4.Inicial, Consulta_4.Debitos, Consulta_4.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                sql &= "               FROM          (SELECT 99999999 as Empresa, LEFT(Conta_Id, 7) AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 AS Debitos, 0 AS Creditos"
                sql &= " FROM Razao"
                sql &= "                               WHERE       Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') AND (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                sql &= "                       GROUP BY LEFT(Conta_Id, 7)) AS Consulta_4 LEFT OUTER JOIN"
                sql &= "              PlanoDeContasXPlanoReferencial ON LEFT(Consulta_4.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id  "
                sql &= " WHERE (PlanoDeContasXPlanoReferencial.Referencial Is Not NULL)"
                sql &= " UNION"
                sql &= "       SELECT  Empresa,   Consulta_3.Conta, Consulta_3.Inicial, Consulta_3.Debitos, Consulta_3.Creditos, PlanoDeContasXPlanoReferencial_3.Referencial"
                sql &= "              FROM         (SELECT  99999999 as Empresa,   Conta_Id AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 AS Debitos, 0 AS Creditos"
                sql &= "                             FROM          Razao AS Razao_3"
                sql &= "                     WHERE       Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') AND (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                sql &= "             GROUP BY Conta_Id) AS Consulta_3 LEFT OUTER JOIN"
                sql &= "                                    PlanoDeContasXPlanoReferencial AS PlanoDeContasXPlanoReferencial_3 ON LEFT(Consulta_3.Conta, 9) "
                sql &= "                            = PlanoDeContasXPlanoReferencial_3.Conta_Id"
                sql &= " WHERE  (PlanoDeContasXPlanoReferencial_3.Referencial Is Not NULL)"
                sql &= " UNION"
                sql &= "       SELECT Empresa,    Consulta_2.Conta, Consulta_2.Inicial, Consulta_2.Debitos, Consulta_2.Creditos, PlanoDeContasXPlanoReferencial_2.Referencial"
                sql &= "       FROM         (SELECT 99999999 as Empresa,    LEFT(Conta_Id, 7) AS Conta, 0 AS Inicial, SUM(DebitoOficial) AS Debitos, SUM(CreditoOficial) AS Creditos"
                sql &= "                    FROM          Razao AS Razao_2"
                sql &= "       WHERE       Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') AND (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') And "
                sql &= "                           (Lote_Id < 7500)"
                sql &= "                              GROUP BY  LEFT(Conta_Id, 7)) AS Consulta_2 LEFT OUTER JOIN"
                sql &= "                        PlanoDeContasXPlanoReferencial AS PlanoDeContasXPlanoReferencial_2 ON LEFT(Consulta_2.Conta, 7) "
                sql &= "                                 = PlanoDeContasXPlanoReferencial_2.Conta_Id "
                sql &= "   WHERE (PlanoDeContasXPlanoReferencial_2.Referencial Is Not NULL)"
                sql &= "   UNION"
                sql &= "   SELECT  Empresa,   Consulta_1.Conta, Consulta_1.Inicial, Consulta_1.Debitos, Consulta_1.Creditos, PlanoDeContasXPlanoReferencial_1.Referencial"
                sql &= "             FROM   (SELECT 99999999 as Empresa,    Conta_Id AS Conta, 0 AS Inicial, SUM(DebitoOficial) AS Debitos, SUM(CreditoOficial) AS Creditos"
                sql &= "                                    FROM          Razao AS Razao_1"
                sql &= "                                    WHERE       Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') AND (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') And "
                sql &= "                                                   (Lote_Id < 7500)"
                sql &= "                                     GROUP BY  Conta_Id) AS Consulta_1 LEFT OUTER JOIN"
                sql &= "                                     PlanoDeContasXPlanoReferencial AS PlanoDeContasXPlanoReferencial_1 ON LEFT(Consulta_1.Conta, 9) "
                sql &= "                                     = PlanoDeContasXPlanoReferencial_1.Conta_Id"
                sql &= "               WHERE     (PlanoDeContasXPlanoReferencial_1.Referencial IS NOT NULL)) AS Consulta "
                sql &= " GROUP BY Empresa, Conta"

                sql &= " Union"

                sql &= "  SELECT Empresa, Conta_Id, (SELECT     TOP (1) Titulo "
                sql &= "         FROM PlanoDeContas"
                sql &= "         WHERE      (Conta_Id = Consulta.Conta_Id)"
                sql &= "  ) AS Titulo, 0 AS Inicial, 0 AS Debitos, 0 AS Creditos"
                sql &= "   FROM "
                sql &= " ( SELECT 99999999 as Empresa,   Conta_Id"
                sql &= "  FROM PlanoDeContas"
                sql &= "         WHERE   (LEN(Conta_Id) < 9) AND (LEN(Conta_Id) <> 0)"
                sql &= "              GROUP BY  Conta_Id) AS Consulta"
                sql &= " ORDER BY Conta_Id"

                sql &= " select * into #PL2 from #PL1"
                sql &= " select *,case when len(conta_Id) = 7 then  isnull((select top 1 Conta_Id from #pl2 where Conta_Id like ''+ #pl1.conta_Id + '%' and len(Conta_Id) = 9),'') else 'S' end as Superior into #PL3 from #pl1"
                sql &= " select *,case when len(conta_Id) > 6 then  isnull((select top 1 Referencial from  PlanoDeContasXPlanoReferencial where Conta_Id = #pl3.conta_Id),'') else 'S' end as Referencial Into #PL4 from #PL3"

                sql &= " Delete #PL4 Where  Len(Conta_Id) = 7 and Superior = '' and Referencial = '' "
                sql &= " Delete #PL4 Where  Left(Conta_ID, 3) in ('105', '205')"
                sql &= " DELETE #PL4 WHERE Len(Conta_Id) = 7 and Inicial = 0 and Debitos = 0 and Creditos = 0 and Referencial <> '' And Superior = ''"

                sql &= " Select * From #PL4"



                ds = Banco.ConsultaDataSet(sql, "PlanoDeContas")


                Dim Nivel1 As String = ""
                Dim Nivel2 As String = ""
                Dim Analitica As String = ""
                Dim Conta As String = ""
                Dim ContaAnterior As String
                Dim Referencial As String = "N"
                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|I050"
                            linha &= "|01012013"
                            Nivel1 = Microsoft.VisualBasic.Left(.Item("Conta_Id"), 1)
                            Nivel2 = Microsoft.VisualBasic.Left(.Item("Conta_Id"), 3)
                            If Nivel2 = "204" Then
                                linha &= "|03"
                            ElseIf Nivel2 = "203" Then
                                If Microsoft.VisualBasic.Left(strEmpresa(0), 8) = "31890630" Then
                                    linha &= "|02"
                                Else
                                    linha &= "|03"
                                End If

                            ElseIf Nivel2 = "105" Then
                                linha &= "|05"
                            ElseIf Nivel2 = "205" Then
                                linha &= "|05"
                            Else
                                If Nivel1 = "1" Then
                                    linha &= "|01"
                                ElseIf Nivel1 = "2" Then
                                    linha &= "|02"
                                ElseIf Nivel1 = "3" Then
                                    linha &= "|04"
                                ElseIf Nivel1 = "4" Then
                                    linha &= "|04"
                                ElseIf Nivel1 = "5" Then
                                    linha &= "|04"
                                End If
                            End If
                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|A"
                                Analitica = "A"
                            Else
                                If Len(.Item("Superior")) > 7 Then
                                    linha &= "|S"
                                    Analitica = "S"
                                Else
                                    If Len(.Item("Conta_Id")) = 7 Then
                                        linha &= "|A"
                                        Analitica = "A"
                                    Else
                                        linha &= "|S"
                                        Analitica = "S"
                                    End If
                                End If
                            End If
                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|5"
                            ElseIf Len(.Item("Conta_Id")) = 7 Then
                                linha &= "|4"
                            ElseIf Len(.Item("Conta_Id")) = 5 Then
                                linha &= "|3"
                            ElseIf Len(.Item("Conta_Id")) = 3 Then
                                linha &= "|2"
                            ElseIf Len(.Item("Conta_Id")) = 1 Then
                                linha &= "|1"
                            End If

                            linha &= "|" & .Item("Conta_Id")

                            If Len(.Item("Conta_Id")) > 7 Then
                                linha &= "|" & Microsoft.VisualBasic.Left(.Item("Conta_Id"), 7)
                            ElseIf Len(.Item("Conta_Id")) = 7 Then
                                linha &= "|" & Microsoft.VisualBasic.Left(.Item("Conta_Id"), 5)
                            ElseIf Len(.Item("Conta_Id")) = 5 Then
                                linha &= "|" & Microsoft.VisualBasic.Left(.Item("Conta_Id"), 3)
                            ElseIf Len(.Item("Conta_Id")) = 3 Then
                                linha &= "|" & Microsoft.VisualBasic.Left(.Item("Conta_Id"), 1)
                            ElseIf Len(.Item("Conta_Id")) = 1 Then
                                linha &= "|"
                            End If
                            Conta = .Item("Conta_Id")
                            ContaAnterior = Conta
                            linha &= "|" & RTrim(.Item("Titulo"))
                            linha &= "|"
                        End With
                        strm.WriteLine(linha)
                        RegistroI050 += 1
                        RegistroGeral += 1

                        'Registro I051 - Plano de Contas Referencial - RFB - 10 
                        '----------------------------------------------------------------------------------------------------------
                        Referencial = "N"

                        If Analitica = "A" Then
                            sql = " SELECT top 1 'I051' AS REG, '10' AS Cod_Ent_Ref, '' AS Cod_CCus, Referencial"
                            sql &= " FROM PlanoDeContasXPlanoReferencial"
                            sql &= " WHERE Conta_Id = '" & Conta & "'"

                            ds = Banco.ConsultaDataSet(sql, "PlanoReferencial")
                            If ds.Tables(0).Rows.Count > 0 Then
                                For Each dra In ds.Tables(0).Rows
                                    With dra
                                        linha = "|I051"
                                        linha &= "|" & .Item("Cod_Ent_Ref")
                                        linha &= "|" & .Item("Cod_CCus")
                                        linha &= "|" & .Item("Referencial")
                                        linha &= "|"
                                        Referencial = "S"
                                    End With
                                    strm.WriteLine(linha)
                                    RegistroI051 += 1
                                    RegistroGeral += 1
                                Next
                            End If
                            If Referencial = "N" Then
                                sql = " SELECT top 1 'I051' AS REG, '10' AS Cod_Ent_Ref, '' AS Cod_CCus, Referencial"
                                sql &= " FROM PlanoDeContasXPlanoReferencial"
                                sql &= " WHERE  Left(Conta_Id, 7) = '" & Microsoft.VisualBasic.Left(Conta, 7) & "'"
                                ds = Banco.ConsultaDataSet(sql, "PlanoReferencial")
                                If ds.Tables(0).Rows.Count > 0 Then
                                    For Each dra In ds.Tables(0).Rows
                                        With dra
                                            linha = "|I051"
                                            linha &= "|" & .Item("Cod_Ent_Ref")
                                            linha &= "|" & .Item("Cod_CCus")
                                            linha &= "|" & .Item("Referencial")
                                            linha &= "|"
                                            Referencial = "S"
                                        End With
                                        strm.WriteLine(linha)
                                        RegistroI051 += 1
                                        RegistroGeral += 1
                                    Next
                                End If
                            End If
                        End If

                        '----------------------------------------------------------------------
                    Next
                End If




                'Registro I150 - Saldos Periódicos - Identificação do Período
                '----------------------------------------------------------------------------------------------------------

                If Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) = "A" Then
                    Ciclo = 1
                Else
                    Ciclo = 4
                End If

                For i As Integer = 1 To Ciclo

                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_01.Text, 1) = "4" And Ciclo = 4 And i = 1 Then
                        i = 2
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_02.Text, 1) = "4" And Ciclo = 4 And i = 2 Then
                        i = 3
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) = "4" And Ciclo = 4 And i = 3 Then
                        i = 4
                    End If

                    If i = 1 And Ciclo = 1 Then
                        linha = "|I150"
                        linha &= "|0101" & Ano
                        linha &= "|3112" & Ano
                        linha &= "|"
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                    ElseIf Ciclo = 4 And i = 1 Then
                        linha = "|I150"
                        linha &= "|0101" & Ano
                        linha &= "|3103" & Ano
                        linha &= "|"
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/03/" & Ano
                    ElseIf Ciclo = 4 And i = 2 Then
                        linha = "|I150"
                        linha &= "|0104" & Ano
                        linha &= "|3006" & Ano
                        linha &= "|"
                        DataInicial = "01/04" & " / " & Ano
                        DataFinal = "30/06/" & Ano
                    ElseIf Ciclo = 4 And i = 3 Then
                        linha = "|I150"
                        linha &= "|0107" & Ano
                        linha &= "|3009" & Ano
                        linha &= "|"
                        DataInicial = "01/07" & " / " & Ano
                        DataFinal = "30/09/" & Ano
                    ElseIf Ciclo = 4 And i = 4 Then
                        linha = "|I150"
                        linha &= "|0110" & Ano
                        linha &= "|3112" & Ano
                        linha &= "|"
                        DataInicial = "01/10" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                    End If

                    strm.WriteLine(linha)
                    RegistroI150 += 1
                    RegistroGeral += 1
                    DataVirada = CDate(DataInicial).AddDays(1)


                    'Registro I155 - Detalhe dos Saldos Periodicos
                    '----------------------------------------------------------------------------------------------------------

                    sql = "  Select Conta, Sum(Inicial) as Inicial, Sum(Debitos) as Debitos, Sum(Creditos) as Creditos, Referencial from ("
                    sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                    sql &= " FROM         ("

                    sql &= "  SELECT  Left(Empresa_Id, 8) as Empresa,  Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                    sql &= "      FROM Razao"
                    sql &= "            WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')  "
                    sql &= "            And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                    sql &= "           GROUP BY Left(Empresa_Id, 8), Left(Conta_Id, 7)"
                    sql &= "           ) AS Consulta LEFT OUTER JOIN"
                    sql &= "           PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= "  Where (Referencial Is Not null)"

                    sql &= "      Union"

                    sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                    sql &= " FROM         (SELECT Left(Empresa_Id, 8) as Empresa,    Conta_Id  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                    sql &= "    FROM Razao"
                    sql &= "           WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')   "
                    sql &= "         And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                    sql &= "         GROUP BY Left(Empresa_Id, 8), Conta_Id"
                    sql &= "        ) AS Consulta LEFT OUTER JOIN"
                    sql &= "         PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= "  Where (Referencial Is Not null)"

                    sql &= "      Union"

                    sql &= " SELECT  Consulta.Conta,    Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                    sql &= " FROM         (SELECT  Left(Empresa_Id, 8) as Empresa,   Left(Conta_Id, 7)  AS Conta, 0 as Inicial,  SUM(DebitoOficial) as Debitos, Sum(CreditoOficial) AS Creditos"
                    sql &= "   FROM Razao"
                    sql &= "            WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')  "
                    sql &= "        And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000)"
                    sql &= "          GROUP BY Left(Empresa_Id, 8), Left(Conta_Id, 7)"
                    sql &= "          ) AS Consulta LEFT OUTER JOIN"
                    sql &= "          PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= "    Where (Referencial Is Not null)"

                    sql &= "    Union"

                    sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                    sql &= " FROM         (SELECT Left(Empresa_Id, 8) as Empresa, Conta_Id  AS Conta, 0 as Inicial,  SUM(DebitoOficial) as Debitos, Sum(CreditoOficial) AS Creditos"
                    sql &= "       FROM Razao"
                    sql &= "            WHERE   (LEFT(Conta_Id, 1) IN ('1', '2')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')"
                    sql &= "        And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 9000)"
                    sql &= "           GROUP BY Left(Empresa_Id, 8), Conta_Id"
                    sql &= "           ) AS Consulta LEFT OUTER JOIN"
                    sql &= "          PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= "   Where (Referencial Is Not null)"


                    sql &= " ) as Consulta"
                    sql &= " Group by Conta, Referencial Order By Conta"


                    ds = Banco.ConsultaDataSet(sql, "SaldosPeriodicos")

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr In ds.Tables(0).Rows
                            If dr("Inicial") <> 0 Or dr("Debitos") <> 0 Or dr("Creditos") <> 0 Then
                                With dr
                                    linha = "|I155"
                                    'If .Item("Cliente") <> "" Then
                                    'linha &= "|" & .Item("Conta") & .Item("Cliente") & .Item("EndCliente")
                                    'Else
                                    linha &= "|" & .Item("Conta")
                                    'End If
                                    'If .Item("CentroDeCustos") <> 0 Then
                                    ' linha &= "|" & .Item("CentroDeCustos")
                                    'Else
                                    linha &= "|"
                                    'End If
                                    If .Item("Inicial") < 0 Then
                                        linha &= "|" & .Item("Inicial") * -1
                                        linha &= "|C"
                                    Else
                                        linha &= "|" & .Item("Inicial")
                                        linha &= "|D"
                                    End If
                                    linha &= "|" & .Item("Debitos")
                                    linha &= "|" & .Item("Creditos")
                                    Saldo = .Item("Inicial") + (.Item("Debitos") - .Item("Creditos"))
                                    If Saldo < 0 Then
                                        linha &= "|" & Saldo * -1
                                        linha &= "|C"
                                    Else
                                        linha &= "|" & Saldo
                                        linha &= "|D"
                                    End If
                                    linha &= "|"
                                End With
                                strm.WriteLine(linha)
                                RegistroI155 += 1
                                RegistroGeral += 1
                            End If
                        Next
                    End If
                Next i


                'Registro I200 - Lançamento Contábil
                '----------------------------------------------------------------------------------------------------------
                Dim Ordem As Integer = 0
                Dim TemExpurgo As String = "N"

                If TemExpurgo = "S" Then

                    sql = "  SELECT    Movimento_Id, Lote_Id, 0 as Sequencia_Id, '' as Conta_Id, '' as Cliente_Id, 0 as EndCliente_Id,  "
                    sql &= "           0 as CentroDeCustos, Sum(DebitoOficial) AS DebitoOficial, Sum(CreditoOficial) as CreditoOficial, '' as Historico "
                    sql &= " FROM Razao"
                    sql &= " WHERE Razao.Lote_id <> 7500 And Left(Conta_ID, 1) in ('1','2') And Empresa_Id like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') And Movimento_ID between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "'"
                    sql &= " Group by Movimento_Id, Lote_Id"
                    sql &= " Union"
                    sql &= " SELECT    Movimento_Id, Lote_Id, Sequencia_Id, Conta_Id, ltrim(rtrim(Cliente_Id)), EndCliente_Id,  "
                    sql &= " Custo AS CentroDeCustos, DebitoOficial, CreditoOficial, Historico "
                    sql &= " FROM Razao"
                    sql &= " WHERE Left(Conta_ID, 1) in ('1','2') And Razao.Lote_id <> 7500 and Empresa_Id  like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%') And Movimento_ID between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "'"
                    sql &= " ORDER BY Movimento_Id, Lote_Id, Sequencia_Id"

                    ds = Banco.ConsultaDataSet(sql, "LancamentoContabil")

                    Dim Lote As Integer = 0

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr In ds.Tables(0).Rows
                            With dr
                                If .Item("Sequencia_Id") = 0 Then
                                    linha = "|I200"
                                    Lote += 1
                                    linha &= "|" & CDate(.Item("Movimento_Id")).ToSqlDate().Replace("/", "") & "-" & .Item("Lote_Id")
                                    linha &= "|" & CDate(.Item("Movimento_Id")).ToSqlDate()                                                     'Data Inicial
                                    linha &= "|" & .Item("DebitoOficial")
                                    linha &= "|F"
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroI200 += 1
                                    RegistroGeral += 1

                                Else
                                    'Registro I250 - Partidas Do Lancamento
                                    '----------------------------------------------------
                                    linha = "|I250"

                                    If .Item("Cliente_Id") <> "" Then
                                        linha &= "|" & .Item("Conta_Id") & .Item("Cliente_Id") & .Item("EndCliente_Id")
                                    Else
                                        linha &= "|" & .Item("Conta_Id")
                                    End If

                                    'If .Item("CentroDeCustos") <> 0 Then
                                    '    linha &= "|" & .Item("CentroDeCustos")
                                    'Else
                                    linha &= "|"
                                    'End If

                                    If .Item("DebitoOficial") > 0 Then
                                        linha &= "|" & .Item("DebitoOficial")
                                        linha &= "|D"
                                    Else
                                        linha &= "|" & .Item("CreditoOficial")
                                        linha &= "|C"
                                    End If

                                    linha &= "|"

                                    'linha &= "|" & Format(CDate(.Item("Movimento_Id")), "MMdd") & String.Format("{0:D4}", .Item("Lote_Id")) & String.Format("{0:D5}", .Item("Sequencia_Id"))
                                    linha &= "|"
                                    linha &= "|" & Funcoes.RemoveAllEnterKey(RTrim(.Item("Historico")))
                                    linha &= "|"
                                    linha &= "|"

                                    strm.WriteLine(linha)
                                    RegistroI250 += 1
                                    RegistroGeral += 1
                                End If
                            End With
                        Next
                    End If
                End If

                'Resultados


                'Registro I350 - Saldos das Contas de resultado Antes do Encerramento - Identificação da Data
                '---------------------------------------------------------------------------------------------

                If Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) = "A" Then
                    Ciclo = 1
                Else
                    Ciclo = 4
                End If

                For i As Integer = 1 To Ciclo

                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_01.Text, 1) = "4" And Ciclo = 4 And i = 1 Then
                        i = 2
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_02.Text, 1) = "4" And Ciclo = 4 And i = 2 Then
                        i = 3
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) = "4" And Ciclo = 4 And i = 3 Then
                        i = 4
                    End If

                    If i = 1 And Ciclo = 1 Then
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                    ElseIf Ciclo = 4 And i = 1 Then
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/03/" & Ano
                    ElseIf Ciclo = 4 And i = 2 Then
                        DataInicial = "01/04" & " / " & Ano
                        DataFinal = "30/06/" & Ano
                    ElseIf Ciclo = 4 And i = 3 Then
                        DataInicial = "01/07" & " / " & Ano
                        DataFinal = "30/09/" & Ano
                    ElseIf Ciclo = 4 And i = 4 Then
                        DataInicial = "01/10" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                    End If

                    linha = "|I350"
                    linha &= "|" & DataFinal.ToSqlDate()
                    linha &= "|"
                    strm.WriteLine(linha)
                    RegistroI350 += 1
                    RegistroGeral += 1

                    'Registro I355 - Detalhe dos Saldos Periodicos
                    '----------------------------------------------------------------------------------------------------------

                    sql = " Select Conta, Sum(Saldo) as Saldo, Referencial from "
                    sql &= " ( "
                    sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, Consulta.Referencial "
                    sql &= " FROM"
                    sql &= " ("
                    sql &= " SELECT  Consulta.Conta,    Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM "
                    sql &= " (SELECT LEFT(Empresa_Id, 8) AS Empresa,    Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo   FROM Razao    "
                    sql &= " WHERE  (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')   "
                    sql &= " And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500)     "
                    sql &= " GROUP BY LEFT(Empresa_Id, 8), Left(Conta_Id, 7)          ) AS Consulta LEFT OUTER JOIN   "
                    sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= " Where (Referencial Is Not null)"
                    sql &= " Union"
                    sql &= " SELECT  Consulta.Conta,   Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM        "
                    sql &= " (SELECT  LEFT(Empresa_Id, 8) AS Empresa,   Conta_Id  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo FROM Razao          "
                    sql &= " WHERE   (LEFT(Conta_Id, 1) IN ('3', '4')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')   "
                    sql &= " And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500) "
                    sql &= " Group by LEFT(Empresa_Id, 8), Conta_Id) AS Consulta LEFT OUTER JOIN  "
                    sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                    sql &= " Where (Referencial Is Not null) ) as Consulta "
                    sql &= " Group by Conta, Saldo, Referencial) as Consulta"
                    sql &= " Group by Conta, Saldo, Referencial"


                    ds = Banco.ConsultaDataSet(sql, "SaldosPeriodicos")

                    If ds.Tables(0).Rows.Count > 0 Then
                        For Each dr In ds.Tables(0).Rows
                            If dr("Saldo") <> 0 Then
                                With dr
                                    linha = "|I355"
                                    linha &= "|" & .Item("Conta")
                                    linha &= "|"

                                    If .Item("Saldo") < 0 Then
                                        linha &= "|" & .Item("Saldo") * -1
                                        linha &= "|C"
                                    Else
                                        linha &= "|" & .Item("Saldo")
                                        linha &= "|D"
                                    End If
                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                RegistroI355 += 1
                                RegistroGeral += 1
                            End If
                        Next
                    End If

                Next i

                ' Registro I990  - Encerramento do Bloco I
                '----------------------------------
                RegistroI990 += 1

                linha = "|I990"
                linha &= "|" & RegistroI001 + RegistroI050 + RegistroI051 + RegistroI150 + RegistroI155 + RegistroI200 + RegistroI250 + RegistroI350 + RegistroI355 + RegistroI990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                '-------------------------------------------------
                ' Registro J001  - Abertura do Bloco J
                '-------------------------------------------------

                linha = "|J001"
                linha &= "|0"
                linha &= "|"

                strm.WriteLine(linha)
                RegistroJ001 += 1
                RegistroGeral += 1

                '------------------------------------------------
                ' Registro J930  - Identificação dos Signatários 
                '------------------------------------------------

                sql = "  SELECT    Clientes.*, ClientesXEmpresas.*"
                sql &= " FROM      Clientes INNER JOIN"
                sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
                sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
                sql &= " WHERE     Clientes.Cliente_Id = '" & strEmpresa(0) & "' And Clientes.Endereco_Id = 0"

                ds = Banco.ConsultaDataSet(sql, "Clientes")

                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|J930"
                        linha &= "|" & RTrim(.Item("NomeDoContador"))
                        linha &= "|" & .Item("CPFContador")
                        linha &= "|" & .Item("QualificacaoContador")
                        linha &= "|" & Microsoft.VisualBasic.Left(.Item("CodigoQualificacaoContador"), 3)
                        linha &= "|" & Microsoft.VisualBasic.Left(.Item("CRCContador"), 11)
                        linha &= "|"
                        strm.WriteLine(linha)
                        RegistroJ930 += 1
                        RegistroGeral += 1

                        linha = "|J930"
                        linha &= "|" & RTrim(.Item("NomeDoTitular"))
                        linha &= "|" & .Item("CPFTitular")
                        linha &= "|" & .Item("QualificacaoTitular")
                        linha &= "|" & Microsoft.VisualBasic.Left(.Item("CodigoQualificacaoTitular"), 3)
                        linha &= "|"
                        linha &= "|"

                        strm.WriteLine(linha)
                        RegistroJ930 += 1
                        RegistroGeral += 1
                    End With
                Next

                '-------------------------------------------
                ' Registro J990  - Encerramento do Bloco J
                '-------------------------------------------
                RegistroJ990 += 1

                linha = "|J990"
                linha &= "|" & RegistroJ001 + RegistroJ930 + RegistroJ990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                '-------------------------------------------------
                ' Registro M001 - Abertura do Bloco M
                '-------------------------------------------------

                linha = "|M001"
                linha &= "|0"
                linha &= "|"

                strm.WriteLine(linha)
                RegistroM001 += 1
                RegistroGeral += 1


                '----------------------------------------------------------------
                ' Registro M020  - Qualificação da Pessoa Juridica e Retificação 
                '----------------------------------------------------------------

                sql = "  SELECT    Clientes.*, ClientesXEmpresas.*"
                sql &= " FROM      Clientes INNER JOIN"
                sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
                sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
                sql &= " WHERE     Clientes.Cliente_Id = '" & strEmpresa(0) & "' And Clientes.Endereco_Id = 0"

                ds = Banco.ConsultaDataSet(sql, "Clientes")

                For Each dr In ds.Tables(0).Rows
                    With dr
                        linha = "|M020"
                        linha &= "|10"  'Qualificação de PJ
                        linha &= "|" & Microsoft.VisualBasic.Left(ddlTipoDeEscrituracao.Text, 1)    'Tipo De Escrituracao
                        linha &= "|" & txtNumeroDoReciboAnterior.Text   'Numero do Recibo Anterior
                        linha &= "|"  'Identificação da Escrituracao do Periuodo Anterior
                        linha &= "|" & Microsoft.VisualBasic.Left(ddlSituacaoSdoDaEscrituracao.Text, 1) 'Situação do Saldo da Escrituração Anterior
                        linha &= "|1" ' Indicativo de Permissão de Lançamentos do Tipo Inicialização dos saldos Iniciais.
                        linha &= "|" & Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) 'Forma De Apuracção
                        linha &= "|" & Microsoft.VisualBasic.Left(ddlFormaDeTributacao.Text, 1) 'Forma De Tributacao

                        linha &= "|" & IIf(Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) = "A" And Microsoft.VisualBasic.Left(ddlFormaDeTributacao.Text, 1) = "2", Microsoft.VisualBasic.Left(ddlTrimestre_01.Text, 1) & Microsoft.VisualBasic.Left(ddlTrimestre_02.Text, 1) & Microsoft.VisualBasic.Left(ddlTrimestre_03.Text, 1) & Microsoft.VisualBasic.Left(ddlTrimestre_04.Text, 1), "")  'Trimestre de Lucro Arbitrado
                        linha &= "|" & IIf(Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) <> "A", Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_01.Text, 1) & Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_02.Text, 1) & Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) & Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_04.Text, 1), "") 'Apuracao do Trimestre
                        linha &= "|"

                        strm.WriteLine(linha)
                        RegistroM020 += 1
                        RegistroGeral += 1

                    End With
                Next

                If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_01.Text, 1) = "1" Then
                    DataInicial = "01/01" & " / " & Ano
                ElseIf Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_02.Text, 1) = "1" Then
                    DataInicial = "01/04" & " / " & Ano
                ElseIf Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) = "1" Then
                    DataInicial = "01/07" & " / " & Ano
                ElseIf Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) = "1" Then
                    DataInicial = "01/10" & " / " & Ano
                End If

                DataFinal = "31/12/" & Ano

                '--------------------------------------------------------------------------------------------------
                ' Registro M025 - Saldos Iniciais das Contas Patrimoniais Recuperados/Preenchidos Ativos e Passivos
                '--------------------------------------------------------------------------------------------------
                sql = "  Select Conta, Sum(Inicial) as Inicial, Referencial from ("
                sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                sql &= " FROM         ("

                sql &= "  SELECT  Left(Empresa_Id, 8) as Empresa,  Left(Conta_Id, 7)  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                sql &= "      FROM Razao"
                sql &= "            WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')  And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                sql &= "           GROUP BY Left(Empresa_Id, 8), Left(Conta_Id, 7)"
                sql &= "           ) AS Consulta LEFT OUTER JOIN"
                sql &= "           PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                sql &= "  Where (Referencial Is Not null)"

                sql &= "      Union"

                sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                sql &= " FROM         (SELECT Left(Empresa_Id, 8) as Empresa,    Conta_Id  AS Conta, SUM(DebitoOficial - CreditoOficial) AS Inicial, 0 as Debitos, 0 AS Creditos"
                sql &= "    FROM Razao"
                sql &= "           WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) AND Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')   And (Movimento_Id < '" & DataInicial.ToSqlDate() & "') AND (Lote_Id < 9000)"
                sql &= "         GROUP BY Left(Empresa_Id, 8), Conta_Id"
                sql &= "        ) AS Consulta LEFT OUTER JOIN"
                sql &= "         PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                sql &= "  Where (Referencial Is Not null)"

                sql &= "      Union"

                sql &= " SELECT  Consulta.Conta,    Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                sql &= " FROM         (SELECT  Left(Empresa_Id, 8) as Empresa,   Left(Conta_Id, 7)  AS Conta, 0 as Inicial,  SUM(DebitoOficial) as Debitos, Sum(CreditoOficial) AS Creditos"
                sql &= "   FROM Razao"
                sql &= "            WHERE  (LEFT(Conta_Id, 1) IN ('1', '2')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')   And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500)"
                sql &= "          GROUP BY Left(Empresa_Id, 8), Left(Conta_Id, 7)"
                sql &= "          ) AS Consulta LEFT OUTER JOIN"
                sql &= "          PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id"
                sql &= "    Where (Referencial Is Not null)"

                sql &= "    Union"

                sql &= " SELECT  Consulta.Conta,   Consulta.Inicial, Consulta.Debitos, Consulta.Creditos, PlanoDeContasXPlanoReferencial.Referencial"
                sql &= " FROM         (SELECT Left(Empresa_Id, 8) as Empresa, Conta_Id  AS Conta, 0 as Inicial,  SUM(DebitoOficial) as Debitos, Sum(CreditoOficial) AS Creditos"
                sql &= "       FROM Razao"
                sql &= "            WHERE   (LEFT(Conta_Id, 1) IN ('1', '2')) And  Empresa_Id  Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')  And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500)"
                sql &= "           GROUP BY Left(Empresa_Id, 8), Conta_Id"
                sql &= "           ) AS Consulta LEFT OUTER JOIN"
                sql &= "          PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id"
                sql &= "   Where (Referencial Is Not null)"


                sql &= " ) as Consulta"
                sql &= " Group by Conta, Referencial"

                sql &= " Having(SUM(Inicial) <> 0)"
                sql &= " Order By Conta"

                ds = Banco.ConsultaDataSet(sql, "Clientes")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|M025"
                            linha &= "|" & .Item("Conta")
                            linha &= "|" 'Centro de Custo
                            linha &= "|" & .Item("Referencial")
                            If .Item("Inicial") < 0 Then
                                linha &= "|" & .Item("Inicial") * -1
                                linha &= "|C"
                            Else
                                linha &= "|" & .Item("Inicial")
                                linha &= "|D"
                            End If
                            If .Item("Inicial") < 0 Then
                                linha &= "|" & .Item("Inicial") * -1
                                linha &= "|C"
                            Else
                                linha &= "|" & .Item("Inicial")
                                linha &= "|D"
                            End If
                            linha &= "|"
                            strm.WriteLine(linha)
                            RegistroM025 += 1
                            RegistroGeral += 1
                        End With
                    Next
                Else
                    linha = "|M025"
                    linha &= "|101010101"
                    linha &= "|" 'Centro de Custo
                    linha &= "|1.01.01.01.00"
                    linha &= "|0|D"
                    linha &= "|0|D"
                    linha &= "|"
                    strm.WriteLine(linha)
                    RegistroM025 += 1
                    RegistroGeral += 1
                End If


                '--------------------------------------------------------------------
                ' Registro M030 - Identificao do Periodo de Apurao do Lucro Real 
                '--------------------------------------------------------------------
                Dim Forma As String = ""

                If Microsoft.VisualBasic.Left(ddlFormaDeApuracao.Text, 1) = "A" Then
                    Ciclo = 1
                Else
                    Ciclo = 4
                End If

                For i As Integer = 1 To Ciclo

                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_01.Text, 1) = "4" And Ciclo = 4 And i = 1 Then
                        i = 2
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_02.Text, 1) = "4" And Ciclo = 4 And i = 2 Then
                        i = 3
                    End If
                    If Microsoft.VisualBasic.Left(ddlApuracaoDoTrimestre_03.Text, 1) = "4" And Ciclo = 4 And i = 3 Then
                        i = 4
                    End If

                    If i = 1 And Ciclo = 1 Then
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                        Forma = "A00"
                    ElseIf Ciclo = 4 And i = 1 Then
                        DataInicial = "01/01" & " / " & Ano
                        DataFinal = "31/03/" & Ano
                        Forma = "T01"
                    ElseIf Ciclo = 4 And i = 2 Then
                        DataInicial = "01/04" & " / " & Ano
                        DataFinal = "30/06/" & Ano
                        Forma = "T02"
                    ElseIf Ciclo = 4 And i = 3 Then
                        DataInicial = "01/07" & " / " & Ano
                        DataFinal = "30/09/" & Ano
                        Forma = "T03"
                    ElseIf Ciclo = 4 And i = 4 Then
                        DataInicial = "01/10" & " / " & Ano
                        DataFinal = "31/12/" & Ano
                        Forma = "T04"
                    End If

                    If Forma = "A00" Then
                        sql = " Select Sum(Saldo) as Resultado from "
                        sql &= " ( "
                        sql &= " SELECT Consulta.Conta, Consulta.Saldo, Consulta.Referencial "
                        sql &= " FROM"
                        sql &= " ("
                        sql &= " SELECT Consulta.Conta, Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM "
                        sql &= " (SELECT Left(Conta_Id, 7) AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo FROM Razao "
                        sql &= " WHERE (LEFT(Conta_Id, 1) IN ('3', '4')) And Empresa_Id Like ('" & Microsoft.VisualBasic.Left(strEmpresa(0), 8) & "%')  "
                        sql &= " And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500) "
                        sql &= " GROUP BY Left(Conta_Id, 7) ) AS Consulta LEFT OUTER JOIN "
                        sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 7) = PlanoDeContasXPlanoReferencial.Conta_Id "
                        sql &= " Where (Referencial Is Not null)"
                        sql &= " Union"
                        sql &= " SELECT Consulta.Conta, Consulta.Saldo, PlanoDeContasXPlanoReferencial.Referencial FROM "
                        sql &= " (SELECT Conta_Id AS Conta, SUM(DebitoOficial - CreditoOficial) AS Saldo FROM Razao "
                        sql &= " WHERE (LEFT(Conta_Id, 1) IN ('3', '4')) And Empresa_Id Like ('" & Left(strEmpresa(0), 8) & "%')  "
                        sql &= " And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500) "
                        sql &= " Group by Conta_Id) AS Consulta LEFT OUTER JOIN "
                        sql &= " PlanoDeContasXPlanoReferencial ON Left(Consulta.Conta, 9) = PlanoDeContasXPlanoReferencial.Conta_Id "
                        sql &= " Where (Referencial Is Not null) ) as Consulta "
                        sql &= " Group by Conta, Saldo, Referencial) as Consulta"

                    Else
                        sql = "  SELECT  isnull(SUM(isnull(CreditoOficial,0) - isnull(DebitoOficial,0)),0) AS Resultado"
                        sql &= " FROM Razao"
                        sql &= " WHERE Conta_Id IN ('307010101') "
                        sql &= "    And Empresa_Id Like ('" & Left(strEmpresa(0), 8) & "%')  "
                        sql &= "    And (Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DataFinal.ToSqlDate() & "') AND (Lote_Id < 7500) "
                    End If


                    ds = Banco.ConsultaDataSet(sql, "Clientes")
                    For Each dr In ds.Tables(0).Rows
                        With dr
                            linha = "|M030"
                            linha &= "|" & Forma
                            If .Item("Resultado") < 0 Then
                                linha &= "|" & .Item("Resultado") * -1
                                linha &= "|C"
                            Else
                                linha &= "|" & .Item("Resultado")
                                linha &= "|D"
                            End If
                            linha &= "|"
                            strm.WriteLine(linha)
                            RegistroM030 += 1
                            RegistroGeral += 1
                        End With
                    Next
                Next i

                '-------------------------------------------
                ' Registro M990  - Encerramento do Bloco M
                '-------------------------------------------
                RegistroM990 += 1

                linha = "|M990"
                linha &= "|" & RegistroM001 + RegistroM020 + RegistroM990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1


                '-------------------------------------------
                ' Registro 9001  - Abertura do Bloco 9
                '-------------------------------------------
                linha = "|9001"
                linha &= "|0"
                linha &= "|"

                strm.WriteLine(linha)
                Registro9001 += 1
                RegistroGeral += 1

                '-------------------------------------------
                ' Registro 9900  - Registros Do Arquivo
                '-------------------------------------------

                For i As Integer = 1 To 25

                    linha = "|9900"

                    Select Case i
                        Case 1
                            linha &= "|0000"
                            linha &= "|" & Registro0000
                        Case 8
                            linha &= "|I001"
                            linha &= "|" & RegistroI001
                        Case 11
                            linha &= "|I050"
                            linha &= "|" & RegistroI050
                        Case 13
                            linha &= "|I150"
                            linha &= "|" & RegistroI150
                        Case 14
                            linha &= "|I155"
                            linha &= "|" & RegistroI155
                        Case 15
                            linha &= "|I200"
                            linha &= "|" & RegistroI200
                        Case 16
                            linha &= "|I250"
                            linha &= "|" & RegistroI250
                        Case 17
                            linha &= "|I990"
                            linha &= "|" & RegistroI990


                        Case 18
                            linha &= "|J001"
                            linha &= "|" & RegistroJ001
                        Case 20
                            linha &= "|J930"
                            linha &= "|" & RegistroJ930
                        Case 21
                            linha &= "|J990"
                            linha &= "|" & RegistroJ990

                        Case 18
                            linha &= "|M001"
                            linha &= "|" & RegistroM001
                        Case 20
                            linha &= "|M020"
                            linha &= "|" & RegistroM020
                        Case 21
                            linha &= "|M990"
                            linha &= "|" & RegistroM990


                        Case 22
                            linha &= "|9001"
                            linha &= "|" & Registro9001
                        Case 23
                            linha &= "|9900"
                            linha &= "|" & Registro9900 + 3
                        Case 24
                            linha &= "|9990"
                            linha &= "|1"
                        Case 25
                            linha &= "|9999"
                            linha &= "|1"
                    End Select

                    linha &= "|"

                    strm.WriteLine(linha)
                    Registro9900 += 1
                    RegistroGeral += 1
                Next

                '-------------------------------------------
                ' Registro 9990  - Encerramento do Bloco 9
                '-------------------------------------------
                Registro9990 += 1

                linha = "|9990"
                linha &= "|" & Registro9001 + Registro9900 + Registro9990 + 1
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1

                '-------------------------------------------------
                ' Registro 9999  - Encerramento do Arquivo Digital
                '-------------------------------------------------
                Registro9999 += 1
                RegistroGeral += 1

                linha = "|9999"
                linha &= "|" & RegistroGeral
                linha &= "|"

                strm.WriteLine(linha)
                strm.Close()


                If File.Exists(arquivo) Then
                    Dim strm1 = New StreamReader(arquivo, True)

                    linha = strm1.ReadToEnd
                    strm1.Close()

                    linha = linha.Replace("QTD-LIN", RegistroGeral)
                    linha = linha.Replace("QTD-J900", RegistroGeral)

                    linha = Microsoft.VisualBasic.Left(linha, Len(linha) - 2)

                    If Dir(arquivo).Length > 0 Then Kill(arquivo)

                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(linha)
                    strm.Close()
                End If


                '---------------------------------------------------------------------------
                '  Fim do Bloco I
                '---------------------------------------------------------------------------

                'Me.Cursor = Cursors.Default

                '----------------------------------------------------------------------------------------
                'Shell("C:\windows\notepad.exe '" & txtArquivodeSaida.Text & "'")

                '----------------------------------------------------------------------------------------
                cmdArquivoDeSaida.Visible = True
                MsgBox(Me.Page, "Processo concluido com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            strm.Close()
            MsgBox(Me.Page, "Problema na execução do processo." & ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "FCont")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class