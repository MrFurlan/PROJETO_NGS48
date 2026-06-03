Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PisCofins
    Inherits BasePage

    Dim Sql As String
    Dim Sqla As String
    Dim Row As DataRow

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Descricao As String = ""
    Dim Crc As String

    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim IcmsARecolher As Decimal

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Opcao As String = ""

    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            Me.setMenu(eModulo.Fiscal) 'PisCofins
            'If Funcoes.VerificaPermissao("PisCofins", "ACESSAR") Then
            If Not Directory.Exists(Server.MapPath("PisCofins")) Then
                Directory.CreateDirectory(Server.MapPath("PisCofins"))
            End If
            CargaUnidade()
            VerificaUnidade()
            Limpar()
            DdlPerfil.SelectedIndex = 1
            'Else
            '    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página,", "~/Fiscal.aspx")
            '    Exit Sub
            'End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & _
              "FROM Clientes C " & _
              "INNER JOIN ClientesXTipos CT " & _
              "ON C.Cliente_Id = CT.Cliente_Id " & _
              "WHERE CT.Tipo_Id = 050 " & _
              "ORDER BY Nome"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade,"
        Sql &= "        isnull(AcessoEmpresa, '') as AcessoEmpresa,"
        Sql &= "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa"
        Sql &= " from Usuarios"
        Sql &= " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            CarregarProcesso()
            CarregarProcessoIPI()
        Next
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
        Sql &= " FROM   GruposXEmpresas INNER JOIN"
        Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
        Sql &= " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' "

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

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresas()
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        CarregarProcesso()
        CarregarProcessoIPI()
    End Sub

    Protected Sub DdlProcesso_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("ProcessoIcms") = DdlProcesso.SelectedValue

        Sql = "SELECT * " & _
                "FROM ProcessoRAICMS " & _
                "WHERE Empresa_Id = '" & Empresa(0) & "' And Processo_Id = " & DdlProcesso.SelectedValue

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            txtDataInicial.Text = Dr("PeriodoInicial")
            txtDataFinal.Text = Dr("PeriodoFinal")
            txtLivro.Text = Dr("Livro")
            txtFolha.Text = Dr("PaginaInicial")
        Next

        Empresa = DdlEmpresa.SelectedValue.Split("-")
        txtArquivoDeSaida.Text = "Mes-" & Format(DateValue(txtDataInicial.Text), "MM-yyyy") & "-" & Empresa(0) & ".Txt"
    End Sub

    Private Sub CarregarProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIcms") = Empresa(0)
        DdlProcesso.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & _
                "FROM ProcessoRAICMS " & _
                "WHERE Empresa_Id = '" & Empresa(0) & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  "
            Descricao &= CDate(Dr("PeriodoInicial")).ToSqlDate() & "  à  "
            Descricao &= CDate(Dr("PeriodoFinal")).ToSqlDate() & "  Livro  "
            Descricao &= Format(Dr("Livro"), "000") & "  Folha  "
            Descricao &= Format(Dr("PaginaInicial"), "000")

            DdlProcesso.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcesso.Items.Insert(0, "")
        DdlProcesso.SelectedIndex = 0
    End Sub

    Private Sub CarregarProcessoIPI()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIcms") = Empresa(0)
        DdlProcessoIPI.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & _
                "FROM ProcessoRAIPI " & _
                "WHERE Empresa_Id = '" & Empresa(0) & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  "
            Descricao &= CDate(Dr("PeriodoInicial")).ToSqlDate() & "  à  "
            Descricao &= CDate(Dr("PeriodoFinal")).ToSqlDate() & "  Livro  "
            Descricao &= Format(Dr("Livro"), "000") & "  Folha  "
            Descricao &= Format(Dr("PaginaInicial"), "000")

            DdlProcessoIPI.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcessoIPI.Items.Insert(0, "")
        DdlProcessoIPI.SelectedIndex = 0
    End Sub

#Region "Funções"

    Sub GravaResumoIcms()
        Empresa = DdlEmpresa.SelectedValue.Split("-")

        Sql = "INSERT INTO ResumoRAICMS " & _
         "(Empresa_Id, Processo_Id, Codigo_Id, Valor) " & _
         "SELECT '" & Empresa(0) & "', " & Processo & ", " & _
         "Codigo_Id, 0 " & _
         "FROM DescricaoRAICMS"

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) Then MsgBox(Me.Page, "Sucesso na Inclusão.", eTitulo.Sucess)

    End Sub

    Function Validar()
        Dim ok As Boolean = True

        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a Unidade de Negocio.", eTitulo.Info)
            Return False
        ElseIf DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a Empresa.", eTitulo.Info)
            Return False
        ElseIf txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o Periodo Inicial.", eTitulo.Info)
            Return False
        ElseIf txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o Período Final.", eTitulo.Info)
            Return False
        ElseIf txtLivro.Text = "" Then
            MsgBox(Me.Page, "Informe o Número do Livro.", eTitulo.Info)
            Return False
        ElseIf txtFolha.Text = "" Then
            MsgBox(Me.Page, "Informe o Número da Folha.", eTitulo.Info)
            Return False
        End If

        Return True

    End Function

    Sub PegarNumeroDoProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        Processo = 1

        Sql = "SELECT COALESCE(max(Processo_Id), 0) AS Processo  FROM ProcessoRAICMS WHERE Empresa_Id = '" & Empresa(0) & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Processo = Dr("Processo") + 1
        Next
    End Sub

    Sub Limpar()
        DdlEmpresa.Items.Clear()
        DdlProcesso.Items.Clear()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        txtLivro.Text = ""
        txtFolha.Text = ""
    End Sub

#End Region

    Protected Sub cmdNovo_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'If Validar() Then
        '    Empresa = DdlEmpresa.SelectedValue.Split("-")
        '    PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        '    PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
        '    Livro = txtLivro.Text
        '    Folha = txtFolha.Text

        '    If DdlProcesso.Text = "" Then
        '        PegarNumeroDoProcesso()
        '    Else
        '        Processo = DdlProcesso.SelectedValue
        '    End If

        '    If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
        '        GravaResumoIcms()
        '        CarregarProcesso()
        '        DdlProcesso.SelectedValue = Processo
        '    End If
        'End If

    End Sub

    Protected Sub cmdExcluir_Click(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")

        Dim NomeArquivo As String
        Dim linha As String
        Dim strm As StreamWriter
        Dim IntNota As Integer = 0
        Dim IntSequencia As Integer = 0
        Dim BlocoC As Integer = 1
        Dim BlocoD As Integer = 1
        Dim BlocoE As Integer = 1

        Dim Nota As Integer = 0
        Dim Serie As String = ""
        Dim EntradaSaida As String = ""
        Dim Cliente As String = ""
        Dim EndCliente As Integer = 0
        Dim Emissao As String = ""
        Dim Modelo As String = ""
        Dim Cancelada As String = ""
        Dim NossaEmissao As Boolean = False

        Dim Registro0000 As Integer = 0     'Abertura do Arquivo Digital e Identificação e Identificação da Entidade
        Dim Registro0001 As Integer = 0     'Abertura do Bloco 0
        'Dim Registro0005 As Integer = 0     'Dados Complementares da Entidade
        Dim Registro0100 As Integer = 0     'Dados do Contabilista
        Dim Registro0110 As Integer = 0     'Regime de Apuracao da contribuicao Social e de Apropriacao de Credito
        Dim Registro0140 As Integer = 0     'Registro 0140 - Tabela de Cadastro de Estabelecimento
        Dim Registro0150 As Integer = 0     'Tabela de Cadastro do Participante
        Dim Registro0190 As Integer = 0     'Tabela de Unidades de Medidas
        Dim Registro0200 As Integer = 0     'Tabela de Identificação do Item (Produto e Serviços)
        Dim Registro0400 As Integer = 0     'Tabela de Natureza da Operação/Prestação
        Dim Registro0450 As Integer = 0     'Tabela de Informações Complementares do Documento Fiscal
        Dim Registro0460 As Integer = 0     'Tabela de Observações do Lancamento Fiscal
        Dim Registro0990 As Integer = 0     'Encerramento do Bloco 0

        Dim RegistroC001 As Integer = 0     'Abertura do Broco C
        Dim RegistroC100 As Integer = 0     'Nota Fiscal
        Dim RegistroC170 As Integer = 0     'Nota Fiscal X Itens
        Dim RegistroC190 As Integer = 0     'Registro Analitico do Documento

        Dim RegistroC500 As Integer = 0     'Nota Fiscal
        Dim RegistroC510 As Integer = 0     'Nota Fiscal X Itens
        Dim RegistroC590 As Integer = 0     'Registro Analitico do Documento

        Dim RegistroC990 As Integer = 0     'Encerramento do Bloco C

        Dim RegistroD001 As Integer = 0     'Abertura do Bloco D
        Dim RegistroD100 As Integer = 0     'Nota Fiscal de Serviço de Transporte
        Dim RegistroD190 As Integer = 0     'Registro Analitico dos Documentos
        Dim RegistroD500 As Integer = 0     'Nota Fiscal de Serviço de Comunicação
        Dim RegistroD590 As Integer = 0     'Registro Analitico dos Documentos
        Dim RegistroD990 As Integer = 0     'Encerramento do Bloco D

        Dim RegistroE001 As Integer = 0     'Abertura do Bloco D
        Dim RegistroE100 As Integer = 0     'Periodo da Apuração
        Dim RegistroE110 As Integer = 0     'Registro Analitico dos Documentos
        Dim RegistroE111 As Integer = 0
        Dim RegistroE116 As Integer = 0     'Obrigações do Icms A Recvolher - Operações Próprias
        Dim RegistroE500 As Integer = 0     'Nota Fiscal de Serviço de Transporte
        Dim RegistroE510 As Integer = 0     'Apuração do IPI - Operações Próprias
        Dim RegistroE520 As Integer = 0     'Registro Analitico dos Documentos
        Dim RegistroE990 As Integer = 0     'Registro Analitico dos Documentos

        Dim RegistroH001 As Integer = 0     'Abertura do Bloco H
        Dim RegistroH005 As Integer = 0     'Total do Inventário
        Dim RegistroH010 As Integer = 0     'Inventário 
        Dim RegistroH990 As Integer = 0     'Encerramento do Bloco H

        Dim RegistroG001 As Integer = 0     'Abertura do Bloco G
        Dim RegistroG110 As Integer = 0     'Total do Inventário
        Dim RegistroG125 As Integer = 0     'Inventário 
        Dim RegistroG126 As Integer = 0     'Encerramento do Bloco G
        Dim RegistroG130 As Integer = 0     'Encerramento do Bloco G
        Dim RegistroG140 As Integer = 0     'Encerramento do Bloco G
        Dim RegistroG990 As Integer = 0     'Encerramento do Bloco G


        Dim Registro1001 As Integer = 0     'Abertura do Bloco 1
        Dim Registro1100 As Integer = 0     'Registro de Informação sobre exportação
        Dim Registro1105 As Integer = 0     'Documentos Fiscais de Exportação 
        Dim Registro1990 As Integer = 0     'Encerramento do Bloco 1


        Dim Registro9001 As Integer = 0     'Abertura do Bloco 9
        Dim Registro9900 As Integer = 0     'Registros do Arquivo
        Dim Registro9990 As Integer = 0     'Encerramento do Bloco 9
        Dim Registro9999 As Integer = 0     'Encerramento do Arquivo Digital


        Dim RegistroGeral As Integer = 0    'Registro Geral

        Dim CGC As String = ""
        Dim Inscricao As String = ""

        Dim NomeArquivo2 As String = "PisCofins/" & txtArquivoDeSaida.Text.ToUpper
        NomeArquivo = Server.MapPath(NomeArquivo2)

        If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

        Dim Array As New ArrayList

        Dim ds As DataSet
        Dim ds1 As DataSet
        Dim ds2 As DataSet

        strm = New StreamWriter(NomeArquivo, True)

        Try
            ' Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
            ' Teste de Movimento do Periodo
            '---------------------------------------------------------------------------------------------
            ' Bloco C

            Sql = "   SELECT top 1 EntradaSaida_Id, Cfop From ("
            Sql &= "  SELECT    NotasFiscais.EntradaSaida_Id, "

            Sql &= "		isnull      ((SELECT  Top 1   Cfop_Id AS Valor"
            Sql &= "        FROM        NotasFiscaisXEncargos"
            Sql &= "        WHERE      (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "                    AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "                    (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "                    (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'PRODUTO')), 0) AS Cfop"

            Sql &= " FROM           NotasFiscais LEFT OUTER JOIN"
            Sql &= "                NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "                NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "                NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "                NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "') as Consulta"

            Sql &= "  Where (CFOP Not Between 1350 and 1360) and  (CFOP Not Between 2350 and 2360) and "
            Sql &= "        (CFOP Not Between 5350 and 5360) and  (CFOP Not Between 6350 and 6360) and"
            Sql &= "        (CFOP Not Between 1252 and 1253) and  (CFOP Not Between 1302 and 1303)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr0 As DataRow In ds.Tables(0).Rows
                    BlocoC = 0
                Next
            End If

            '------------------------------------------------------------
            ' Bloco C Segunda Parte
            '------------------------------------------------------------
            Sql = " Select top 1 EntradaSaida_Id, CFOP From (SELECT NotasFiscais.EntradaSaida_Id,"
            Sql &= "   ISNULL ((SELECT  Top 1 CFOP_Id AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'Produto')), 0) AS CFOP"

            Sql &= "   FROM NotasFiscais LEFT OUTER JOIN"
            Sql &= "        NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "        NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "        NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "        NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F','REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "') as Consulta "

            Sql &= "  Where (CFOP Between 1252 and 1253)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                BlocoC = 0
            End If

            '--------------------------------------------------------------------
            ' Bloco D

            Sql = "   SELECT top 1  EntradaSaida_Id, Cfop From ("
            Sql &= "  SELECT NotasFiscais.EntradaSaida_Id,"
            Sql &= "		isnull      ((SELECT  Top 1   Cfop_Id AS Valor"
            Sql &= "        FROM        NotasFiscaisXEncargos"
            Sql &= "        WHERE      (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "                    AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "                    (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "                    (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'PRODUTO')), 0) AS Cfop"

            Sql &= " FROM           NotasFiscais LEFT OUTER JOIN"
            Sql &= "                NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "                NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "                NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "                NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F','REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "') as Consulta"

            Sql &= "  Where (CFOP Between 1350 and 1360) Or  (CFOP Between 2350 and 2360) Or "
            Sql &= "        (CFOP Between 5350 and 5360) Or  (CFOP Between 6350 and 6360)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                BlocoD = 0
            End If

            '------------------------------------------------------------
            ' Bloco D Segunda Parte
            '------------------------------------------------------------
            Sql = " Select top 1 EntradaSaida_Id, CFOP From (SELECT NotasFiscais.EntradaSaida_Id,"
            Sql &= "   ISNULL ((SELECT  Top 1 CFOP_Id AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'Produto')), 0) AS CFOP"

            Sql &= "   FROM NotasFiscais LEFT OUTER JOIN"
            Sql &= "        NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "        NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "        NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "        NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F','REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "') as Consulta "

            Sql &= "  Where (CFOP Between 1300 and 1350) Or  (CFOP Between 2300 and 2350) Or "
            Sql &= "        (CFOP Between 5300 and 5350) Or  (CFOP Between 6300 and 6350)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                BlocoD = 0
            End If

            '-------------------------------------------------------------------------------------------
            'Registro 0000 - Abertura do Arquivo Digital e Identificação do Empresário ou da Sociedade Empresária
            '----------------------------------------------------------------------------------------------------------

            Sql = "  SELECT  distinct  Clientes.Cliente_Id, Clientes.Nome, Clientes.CodigoDoMunicipio,Clientes.Estado, Clientes.Inscricao, Municipios.EstadoIbge, ClientesXEmpresas.InscricaoMunicipal"
            Sql &= " FROM      Clientes INNER JOIN"
            Sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            Sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN"
            Sql &= "           Estados ON Clientes.Estado = Estados.Estado_Id INNER JOIN"
            Sql &= "           Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id"
            Sql &= " WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0"

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0000"                                                                                                   'Fixo bloco
                        linha &= IIf(PeriodoFinal.Year < 2011, "|100", "|101")                                                            'Cod. Versao Leiaute
                        linha &= "|0"                                                                                                     'Tipo Escrituracao
                        linha &= "|" & Left(" ", 41)                                                                                      'Indicador de situacao especial
                        linha &= "|" & txtDataInicial.Text.ToSqlDate()                                                    'Data Inicial
                        linha &= "|" & txtDataFinal.Text.ToSqlDate()                                                       'Data Final
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                        linha &= "|" & Microsoft.VisualBasic.Left(Trim(.Item("Cliente_Id")), 14)
                        linha &= "|" & .Item("Estado")                                                                                    'Estado
                        linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 7, "0")                   'Cod. Municipio
                        linha &= "|" & Trim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual
                        linha &= "|" & Trim(.Item("InscricaoMunicipal"))
                        linha &= "|" & Left("0", 9)                             'Insc. Suframa
                        linha &= "|" & Left("0", 2)                             'ind. natureza pessoa juridica
                        linha &= "|9"                             'ind. tipo de atividade preponderante
                        'linha &= "|" & DdlPerfil.SelectedValue
                        'linha &= "|" & DdlAtividade.SelectedValue
                        linha &= "|"

                        HttpContext.Current.Session("Estado") = .Item("Estado")

                    End With

                    strm.WriteLine(linha)
                    Registro0000 += 1
                    RegistroGeral += 1
                Next
            End If

            ' Registro 0001  - Abertura do Bloco 0
            '----------------------------------

            linha = "|0001"
            linha &= "|0"
            linha &= "|"

            strm.WriteLine(linha)
            Registro0001 += 1
            RegistroGeral += 1

            ''Registro 0005 - Dados Complementares da Entidade
            ''----------------------------------------------------------------------------------------------------------
            'Sql = "  SELECT Clientes.Fantasia, Clientes.Endereco, Clientes.Numero, Clientes.Complemento, Clientes.Bairro, Clientes.Cep, Clientes.Telefone, isnull(Clientes.Fax,'') as Fax, isnull(Clientes.Email,'') as Email"
            'Sql &= " FROM         Clientes INNER JOIN"
            'Sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            'Sql &= "            Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN"
            'Sql &= "             Estados ON Clientes.Estado = Estados.Estado_Id"

            'Sql &= " WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0"

            'ds = Banco.ConsultaDataSet(Sql, "Clientes")

            'If ds.Tables(0).Rows.Count > 0 Then
            '    For Each dr As DataRow In ds.Tables(0).Rows
            '        With dr
            '            linha = "|0005"
            '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Fantasia"))
            '            linha &= "|" & .Item("Cep").ToString.Trim.Replace("-", "").Replace(".", "")
            '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
            '            linha &= "|" & Trim(.Item("Numero"))
            '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
            '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
            '            linha &= "|" & Funcoes.AlinharDireita(.Item("Telefone").ToString.Trim.Replace("(", "").Replace(")", "").Replace("-", ""), 10, "0")
            '            linha &= "|" & Funcoes.AlinharDireita(.Item("Fax").ToString.Trim.Replace("(", "").Replace(")", "").Replace("-", ""), 10, "0")
            '            linha &= "|" & Trim(.Item("Email"))
            '            linha &= "|"
            '        End With

            '        strm.WriteLine(linha)
            '        Registro0005 += 1
            '        RegistroGeral += 1
            '    Next
            'End If

            'Registro 0100 - Dados do Contabilista
            '----------------------------------------------------------------------------------------------------------

            Sql = "  SELECT ClientesXEmpresas.CPFContador, ClientesXEmpresas.CRCContador"
            Sql &= " FROM      Clientes INNER JOIN"
            Sql &= "           ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND "
            Sql &= "           Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
            Sql &= " WHERE     Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0"

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    Codigo = dr("CpfContador")
                    Crc = dr("CRCContador").Replace(".", "").Replace("-", "")
                Next
            End If

            Sql = " SELECT DISTINCT Clientes.Nome, Clientes.Cliente_Id, Clientes.Cep, Clientes.Endereco, Clientes.Numero, Clientes.Complemento, Clientes.Bairro, isnull(Clientes.Telefone,'')as Telefone, isnull(Clientes.Fax,'')as Fax, isnull(Clientes.Email,'') as Email, Municipios.Estadoibge, Clientes.CodigoDoMunicipio"
            Sql &= "  FROM Clientes "
            Sql &= " INNER JOIN Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id"
            Sql &= " WHERE Clientes.Cliente_Id = '" & Codigo & "' And Clientes.Endereco_Id = 0"

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0100"
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                        linha &= "|" & Trim(.Item("Cliente_Id"))
                        linha &= "|" & Crc
                        linha &= "|"
                        linha &= "|" & .Item("CEP").ToString.Trim.Replace("-", "").Replace(".", "")
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
                        linha &= "|" & Trim(.Item("Numero"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
                        linha &= "|" & Trim(.Item("Telefone").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""))
                        linha &= "|" & Trim(.Item("Fax").Replace(" ", "").Replace("(", "").Replace(")", "").Replace("-", ""))
                        linha &= "|" & Trim(.Item("Email"))
                        linha &= "|" & .Item("Estadoibge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0100 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0110 - Regimes de Apuração da Contribuição Social e de Apropriação de Credito

            linha = "|0110"
            linha &= "|1"
            linha &= "|0"
            linha &= "|0"
            linha &= "|"

            strm.WriteLine(linha)
            Registro0110 += 1
            RegistroGeral += 1


            'Registro 0140 - Tabela de Cadastro de Estabelecimento

            Sql = " SELECT DISTINCT Clientes.Nome, Clientes.Cliente_Id, Clientes.Cep, Clientes.Endereco, Clientes.Numero, Clientes.Complemento, Clientes.Bairro, isnull(Clientes.Telefone,'')as Telefone, isnull(Clientes.Fax,'')as Fax, isnull(Clientes.Email,'') as Email, Municipios.Estadoibge, Clientes.CodigoDoMunicipio"
            Sql &= "  FROM Clientes "
            Sql &= " INNER JOIN Municipios ON Clientes.Estado = Municipios.Estado_id AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id"
            Sql &= " WHERE Clientes.Cliente_Id = '" & Empresa(0) & "' And Clientes.Endereco_Id = 0"

            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0140"
                        linha &= "|" & Left(" ", 60)
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                        linha &= "|" & Trim(.Item("Cliente_Id"))
                        linha &= "|" & .Item("Estado")
                        linha &= "|" & Trim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", ""), 14, " "))    'Inscricao Estadual'Estado
                        linha &= "|" & .Item("EstadoIbge") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 7, "0")                   'Cod. Municipio
                        linha &= "|" & Trim(.Item("InscricaoMunicipal"))
                        linha &= "|" & Left("0", 9)                             'Insc. Suframa
                        linha &= "|"

                    End With

                    strm.WriteLine(linha)
                    Registro0140 += 1
                    RegistroGeral += 1
                Next
            End If



            'Registro 0150 - Tabela de Cadastro Do Participante
            '----------------------------------------------------------------------------------------------------------
            Sql = " SELECT distinct Cliente_Id, Endereco_Id, Regiao, Categoria, Estado, Pais, Nome, " & vbCrLf & _
                  "        Fantasia, Endereco, Numero, Complemento, Bairro, Cep, " & vbCrLf & _
                  "        Cidade, Inscricao, Telefone, Fax, Email, Imagem, Reduzido, " & vbCrLf & _
                  "        CodigoDoMunicipio, Situacao, Estado_Id, Descricao, Estadoibge " & vbCrLf & _
                  "   From (" & vbCrLf & _
                  "         SELECT distinct Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Regiao, Clientes.Categoria, Clientes.Estado, isnull(Clientes.Pais,1058) as Pais, Clientes.Nome, " & vbCrLf & _
                  "                Clientes.Fantasia, Clientes.Endereco, Clientes.Numero, ISNULL(Clientes.Complemento, N'') AS Complemento, Clientes.Bairro, Clientes.Cep, " & vbCrLf & _
                  "                Clientes.Cidade, Clientes.Inscricao, Clientes.Telefone, Clientes.Fax, Clientes.Email, Clientes.Imagem, Clientes.Reduzido, " & vbCrLf & _
                  "                ISNULL(Clientes.CodigoDoMunicipio, 0) AS CodigoDoMunicipio, Clientes.Situacao, Estados.Estado_Id, Estados.Descricao, Municipios.Estadoibge" & vbCrLf & _
                  "           FROM NotasFiscais  " & vbCrLf & _
                  "          INNER JOIN Clientes" & vbCrLf & _
                  "             ON Clientes.Cliente_Id  = NotasFiscais.Cliente_Id" & vbCrLf & _
                  "            AND Clientes.Endereco_Id = NotasFiscais.EndCliente_Id " & vbCrLf & _
                  "          INNER JOIN Estados" & vbCrLf & _
                  "             ON Clientes.Estado = Estados.Estado_Id" & vbCrLf & _
                  "          INNER JOIN Municipios" & vbCrLf & _
                  "             ON Estados.Estado_Id                     = Municipios.Estado_id" & vbCrLf & _
                  "            AND ISNULL(Clientes.CodigoDoMunicipio, 0) = Municipios.Codigo_id" & vbCrLf & _
                  "          WHERE NotasFiscais.Situacao not in (9,2) " & vbCrLf & _
                  "            And NotasFiscais.Operacao > 0 " & vbCrLf & _
                  "            And NotasFiscais.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                  "            And NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                  "            And NotasFiscais.Serie_Id not in('D','F','REC') " & vbCrLf

            If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
                Sql &= "          Union" & vbCrLf & _
                       "          SELECT Clientes.Cliente_Id, Clientes.Endereco_Id, Clientes.Regiao, Clientes.Categoria, Clientes.Estado, Clientes.Pais, Clientes.Nome, Clientes.Fantasia," & vbCrLf & _
                       "                 Clientes.Endereco, Clientes.Numero, ISNULL(Clientes.Complemento, '') AS Complemento, Clientes.Bairro, Clientes.Cep, Clientes.Cidade, " & vbCrLf & _
                       "                 Clientes.Inscricao, Clientes.Telefone, Clientes.Fax, Clientes.Email, Clientes.Imagem, Clientes.Reduzido," & vbCrLf & _
                       "                 ISNULL(Clientes.CodigoDoMunicipio, 0) AS CodigoDoMunicipio, Clientes.Situacao, Estados.Estado_Id, Estados.Descricao, Municipios.Estadoibge" & vbCrLf & _
                       "            FROM Clientes " & vbCrLf & _
                       "           INNER JOIN Estados" & vbCrLf & _
                       "              ON Clientes.Estado = Estados.Estado_Id " & vbCrLf & _
                       "           INNER JOIN ApuracaoDeCustos" & vbCrLf & _
                       "              ON Clientes.Cliente_Id  = ApuracaoDeCustos.Deposito_Id " & vbCrLf & _
                       "             AND Clientes.Endereco_Id = ApuracaoDeCustos.EndDeposito_Id " & vbCrLf & _
                       "           INNER JOIN Municipios" & vbCrLf & _
                       "              ON Clientes.Estado            = Municipios.Estado_id" & vbCrLf & _
                       "             AND Clientes.CodigoDoMunicipio = Municipios.Codigo_id" & vbCrLf & _
                       "           WHERE ApuracaoDeCustos.CodigoDeCusto_Id IN (109)" & vbCrLf & _
                       "             And ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                       "             And ApuracaoDeCustos.Ano_id =" & PeriodoInicial.Year - 1 & vbCrLf & _
                       "             And ApuracaoDeCustos.Mes_id = 12" & vbCrLf
            End If
            Sql &= " ) as Consulta" & vbCrLf & _
                   "  ORDER BY Nome, Estado, Cidade "


            ds = Banco.ConsultaDataSet(Sql, "Clientes")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0150"
                        linha &= "|" & Trim(.Item("Cliente_Id")) & Trim(.Item("Endereco_Id"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                        linha &= "|" & .Item("Pais")
                        If Microsoft.VisualBasic.Len(.Item("Cliente_Id")) = 11 And .Item("Pais") = 1058 Then
                            linha &= "|"
                            linha &= "|" & Trim(.Item("Cliente_Id"))
                        End If
                        If Microsoft.VisualBasic.Len(.Item("Cliente_Id")) = 14 And .Item("Pais") = 1058 Then
                            linha &= "|" & Trim(.Item("Cliente_Id"))
                            linha &= "|"
                        End If
                        If .Item("Pais") <> 1058 Then
                            linha &= "|"
                            linha &= "|"
                        End If

                        'linha &= "|" & .Item("Inscricao")
                        linha &= "|" & RTrim(Funcoes.AlinharEsquerda(.Item("Inscricao").Replace(".", "").Replace(",", "").Replace("-", "").Replace("ISENTO", "").Replace("ISENTA", ""), 14, " "))    'Inscricao Estadual

                        If .Item("Pais") = 1058 Then
                            linha &= "|" & .Item("EstadoIBGE") & Funcoes.AlinharDireita(.Item("CodigoDoMunicipio"), 5, "0")
                        Else
                            linha &= "|"
                        End If

                        linha &= "|"   'Suframa
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Endereco"))
                        linha &= "|" & Trim(.Item("Numero"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Complemento"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Bairro"))
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0150 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0190 - Identificação das Unidades de Medidas
            '----------------------------------------------------------------------------------------------------------


            'Sql = "SELECT distinct P.Unidade, UN.Descricao" & vbCrLf & _
            '      "  FROM NotasFiscais NF" & vbCrLf & _
            '      " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
            '      "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
            '      "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
            '      "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
            '      "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
            '      "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
            '      "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
            '      "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
            '      " INNER JOIN Produtos P" & vbCrLf & _
            '      "    on P.Produto_Id = NFxI.Produto_Id" & vbCrLf & _
            '      " INNER JOIN UnidadeDeMedida UN" & vbCrLf & _
            '      "        ON P.Unidade = UN.Unidade_Id" & vbCrLf & _
            '      " WHERE NF.Serie_Id not IN ('D', 'F','REC')" & vbCrLf & _
            '      "   AND NF.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
            '      "   AND NF.Movimento BETWEEN '" & Format(PeriodoInicial, "yyyy/MM/dd") & "' AND '" & Format(PeriodoFinal, "yyyy/MM/dd") & "'" & vbCrLf & _
            '      "   AND NF.Situacao <> 9" & vbCrLf & _
            '      "   AND NFxI.Operacao > 0" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 1350 AND 1360" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 2350 AND 2360" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 5350 AND 5360" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 6350 AND 6360" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 1252 AND 1253" & vbCrLf & _
            '      "   AND NFxI.CFOP_ID NOT BETWEEN 1302 AND 1303" & vbCrLf


            Sql = "SELECT distinct P.Unidade, UN.Descricao" & vbCrLf & _
                  "  FROM NotasFiscais NF" & vbCrLf & _
                  " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
                  "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
                  "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
                  "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
                  "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
                  "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
                  "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
                  "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
                  " INNER JOIN Produtos P" & vbCrLf & _
                  "    on P.Produto_Id = NFxI.Produto_Id" & vbCrLf & _
                  " INNER JOIN UnidadeDeMedida UN" & vbCrLf & _
                  "        ON P.Unidade = UN.Unidade_Id" & vbCrLf & _
                  " Left JOIN NFERealizadas " & vbCrLf & _
                  "   ON NF.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                  "  AND NF.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                  "  AND NF.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                  "  AND NF.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                  "  AND NF.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                  "  AND NF.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                  "  AND NF.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                  " WHERE NF.Serie_Id not IN ('D', 'F','REC')" & vbCrLf & _
                  "   AND NF.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                  "   AND NF.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                  "   AND NF.Situacao <> 9" & vbCrLf & _
                  "   AND NFxI.Operacao > 0" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 1350 AND 1360" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 2350 AND 2360" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 5350 AND 5360" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 6350 AND 6360" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 1252 AND 1253" & vbCrLf & _
                  "   AND NFxI.CFOP_ID NOT BETWEEN 1302 AND 1303" & vbCrLf & _
                  "   AND LEN(isnull(NFERealizadas.chavenfe,0)) < 10"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0190"
                        linha &= "|" & Trim(.Item("Unidade"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0190 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0200 - Tabela de Identificação do Item (Produtos e Serviços)
            '----------------------------------------------------------------------------------------------------------
            Sql = "SELECT Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS" & vbCrLf & _
                  "  FROM (" & vbCrLf & _
                  "               SELECT     Produtos.Produto_Id, Produtos.Nome, Produtos.Unidade, ISNULL(Produtos.TipoDoItem, 0) AS TipoDoItem, ISNULL(Produtos.NCM, '') AS NCM," & vbCrLf & _
                  "                                              ISNULL(Produtos.CodigoEX, '') AS CodigoEX, ISNULL(Produtos.CodigoDoGenero, 0) AS CodigoDoGenero, ISNULL(Produtos.CodigoDoServico, 0)" & vbCrLf & _
                  "                                              AS CodigoDoServico, ISNULL(Produtos.ICMS, 0) AS ICMS, NotasFiscaisXItens.CFOP_Id AS CFOP" & vbCrLf & _
                  "                 FROM NotasFiscais" & vbCrLf & _
                  "                INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                  "                   ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                  "                INNER JOIN Produtos " & vbCrLf & _
                  "                   ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                  "                 LEFT JOIN NFERealizadas" & vbCrLf & _
                  "                   ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                  "                  AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                  "                  AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                  "                WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                  "                  AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & vbCrLf & _
                  "                  AND NotasFiscais.Movimento BETWEEN  '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                  "                  AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "                  AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                  "                  AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null))" & vbCrLf & _
                  "   group by Produtos.Produto_Id, Produtos.Nome, Produtos.Unidade, ISNULL(Produtos.TipoDoItem, 0)," & vbCrLf & _
                  "            ISNULL(Produtos.NCM, ''),           ISNULL(Produtos.CodigoEX, '')," & vbCrLf & _
                  "            ISNULL(Produtos.CodigoDoGenero, 0), ISNULL(Produtos.CodigoDoServico, 0)," & vbCrLf & _
                  "            ISNULL(Produtos.ICMS, 0),           NotasFiscaisXItens.CFOP_Id" & vbCrLf & _
                  "   having sum(NotasFiscaisxitens.quantidadefiscal) > 0" & vbCrLf

            If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
                Sql &= " Union"
                Sql &= " SELECT Produtos.Produto_Id, Produtos.Nome, Produtos.Unidade, ISNULL(Produtos.TipoDoItem, 0) AS TipoDoItem, ISNULL(Produtos.NCM, '') AS NCM, " & vbCrLf & _
                       "        ISNULL(Produtos.CodigoEx, '') AS CodigoEX, ISNULL(Produtos.CodigoDoGenero, 0) AS CodigoDoGenero," & vbCrLf & _
                       "        ISNULL(Produtos.CodigoDoServico, 0) AS CodigoDoServico, ISNULL(Produtos.ICMS, 0) AS ICMS , 0 as CFOP" & vbCrLf & _
                       "   FROM Produtos " & vbCrLf & _
                       "  INNER JOIN ApuracaoDeCustos " & vbCrLf & _
                       "     ON Produtos.Produto_Id = ApuracaoDeCustos.Produto_Id" & vbCrLf & _
                       "  Where ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920)" & vbCrLf & _
                       "    And ApuracaoDeCustos.Empresa_Id ='" & Empresa(0) & "'" & vbCrLf & _
                       "    And ApuracaoDeCustos.Ano_id     = " & PeriodoInicial.Year - 1 & vbCrLf & _
                       "    And ApuracaoDeCustos.Mes_id     = 12"
            End If

            Sql &= " ) AS Consulta"
            Sql &= " GROUP BY Produto_Id, Nome, Unidade, TipoDoItem, NCM, CodigoEX, CodigoDoGenero, CodigoDoServico, ICMS"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0200"
                        linha &= "|" & Trim(.Item("Produto_Id"))
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                        linha &= "|"                    'Código de Barra
                        linha &= "|"                    'Código Anterior
                        linha &= "|" & Trim(.Item("Unidade"))
                        linha &= "|" & Funcoes.AlinharDireita(.Item("TipoDoItem"), 2, "0")
                        linha &= "|" & Left(Trim(.Item("NCM").Replace(".", "")), 8)
                        linha &= "|" & Trim(.Item("CodigoEX"))
                        If .Item("CodigoDoGenero") = 0 Then
                            linha &= "|"
                        Else
                            linha &= "|" & CInt(.Item("CodigoDoGenero")).ToString("00")
                        End If
                        linha &= "|" & Trim(.Item("CodigoDoServico"))
                        linha &= "|" & Trim(.Item("ICMS"))
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0200 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0400 - Tabela de natureza da operação/prestação
            '----------------------------------------------------------------------------------------------------------
            Sql = "  SELECT distinct SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf & _
                  "    FROM NotasFiscais" & vbCrLf & _
                  "   Inner Join NotasFiscaisXItens" & vbCrLf & _
                  "      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                  "     AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                  "     AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                  "     AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                  "     AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                  "     AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                  "     AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                  "    LEFT JOIN NFERealizadas" & vbCrLf & _
                  "      ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                  "     AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                  "     AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                  "     AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                  "     AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                  "     AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                  "     AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id     " & vbCrLf & _
                  "   INNER JOIN SubOperacoes " & vbCrLf & _
                  "      ON isnull(NotasFiscaisXItens.Operacao,NotasFiscais.Operacao) = SubOperacoes.Operacao_Id " & vbCrLf & _
                  "     AND isnull(NotasFiscaisXItens.SubOperacao,NotasFiscais.SubOperacao) = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                  "   WHERE NotasFiscais.Serie_Id Not IN ('D', 'F','REC')" & vbCrLf & _
                  "     AND (NotasFiscais.Empresa_Id =  '" & Empresa(0) & "')" & vbCrLf & _
                  "     AND (NotasFiscais.Operacao > 0) " & vbCrLf & _
                  "     AND (NotasFiscais.Situacao <> 9) " & vbCrLf & _
                  "     AND (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf & _
                  "     AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1300 AND 1360)" & vbCrLf & _
                  "     AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf & _
                  "     AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 5300 AND 5360)" & vbCrLf & _
                  "     AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 6300 AND 6360)" & vbCrLf & _
                  "     AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                  "     AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null))" & vbCrLf & _
                  "   Group by SubOperacoes.Operacao_Id, SubOperacoes.SubOperacoes_Id, SubOperacoes.Descricao" & vbCrLf & _
                  "  having(SUM(NotasFiscaisXItens.QuantidadeFiscal) > 0)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0400"
                        linha &= "|" & .Item("Operacao_Id") & .Item("SubOperacoes_Id")
                        linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0400 += 1
                    RegistroGeral += 1
                Next
            End If

            'Registro 0450 - Tabela de informações complementares do documento fiscal
            '--------------------------------------------------------------------------------------------------
            'Sql = " SELECT   codigo_id, Descricao"
            'Sql &= " FROM    ObservacoesTributarias"
            'Sql &= " WHERE   (Estado = '" & HttpContext.Current.Session("Estado") & "')"


            Sql = "SELECT codigo_id, Descricao" & vbCrLf & _
                  "  FROM ObservacoesTributarias" & vbCrLf & _
                  " WHERE codigo_id IN      (" & vbCrLf & _
                  "                          SELECT OT.codigo_id" & vbCrLf & _
                  "                            FROM NotasFiscais " & vbCrLf & _
                  "						   INNER JOIN NotasFiscaisXItens" & vbCrLf & _
                  "						      ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                  "						   INNER JOIN ObservacoesTributarias AS OT " & vbCrLf & _
                  "                              ON CONVERT(varchar(400), OT.Descricao) = CONVERT(varchar(400), NotasFiscais.ObservacoesDeEmbarque)" & vbCrLf & _
                  "							LEFT JOIN NFERealizadas" & vbCrLf & _
                  "							  ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                  "							 AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id " & vbCrLf & _
                  "							 AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                  "						   WHERE NOT NotasFiscais.Serie_Id IN ('D', 'F','REC')" & vbCrLf & _
                  "							 AND NotasFiscais.Empresa_Id =  '" & Empresa(0) & "'" & _
                  "							 AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & _
                  "							 AND NotasFiscaisXItens.Operacao > 0" & vbCrLf & _
                  "                          AND (NotasFiscais.Situacao <> 9) " & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1302 AND 1303)" & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 2350 AND 2360)" & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "							 AND (NotasFiscaisXItens.CFOP_Id NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                  "							 AND ((NFERealizadas.Chavenfe is null) or (NotasFiscais.NossaEmissao = 'N' and NFERealizadas.Chavenfe is not null))" & vbCrLf & _
                  "                           Group by OT.codigo_id" & vbCrLf & _
                  "                          having sum(NotasFiscaisxitens.quantidadefiscal) > 0)" & vbCrLf

            'Sql = "select Codigo_id, Descricao from ObservacoesTributarias where codigo_id in (" & _
            '      "select distinct OT.Codigo_id" & _
            '      "  from notasfiscais" & _
            '      " inner Join ObservacoesTributarias OT" & _
            '      "    on convert(varchar,OT.Descricao) = convert(varchar,notasfiscais.observacoesdeembarque)" & _
            '      "  WHERE NotasFiscais.Operacao > 0 And NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND " & _
            '      "        NotasFiscais.Movimento BETWEEN '" & Format(PeriodoInicial, "yyyy/MM/dd") & "' AND '" & Format(PeriodoFinal, "yyyy/MM/dd") & "'" & _
            '      "    And NotasFiscais.Serie_Id not in('D','F') )"


            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|0450"
                        linha &= "|" & .Item("Codigo_Id")
                        linha &= "|" & Left(Funcoes.EliminarCaracteresEspeciais(.Item("Descricao")), 200)
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    Registro0450 += 1
                    RegistroGeral += 1
                Next
            End If


            'Registro 0460 - Tabela de informações complementares do documento fiscal
            '--------------------------------------------------------------------------------------------------
            'Sql = " SELECT  Codigo_Id, Descricao"
            'Sql &= " FROM   ObservacoesFiscais"

            'Dim ds13 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

            'If ds13.Tables(0).Rows.Count > 0 Then
            '    For Each dr13 As DataRow In ds13.Tables(0).Rows
            '        With dr13
            '            linha = "|0460"
            '            linha &= "|" & .Item("Codigo_Id")
            '            linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Descricao"))
            '            linha &= "|"
            '        End With

            '        strm.WriteLine(linha)
            '        Registro0460 += 1
            '        RegistroGeral += 1
            '    Next
            'End If


            '' Registro 0990  - Encerramento do Bloco 0
            ''----------------------------------
            Registro0990 += 1

            linha = "|0990"
            linha &= "|" & Registro0000 + Registro0001 + Registro0100 + Registro0110 + Registro0140 + Registro0150 + Registro0190 + Registro0200 + Registro0400 + Registro0450 + Registro0460 + Registro0990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1

            '---------------------------------------------------------------------------
            '  Fim do Bloco 0
            '---------------------------------------------------------------------------

            '-------------------------------------------
            ' Abertura do Bloco C
            '-------------------------------------------

            'Registro C100 - Nota Fiscal (Código 01), Nota Fiscal Avulsa (Código 1B), 
            '                Nota Fiscal de Produtor (Código 04) e Nfe (código 55)
            '----------------------------------------------------------------------------------------------------------
            Sql = "SELECT EntradaSaida_Id, Cliente_Id, EndCliente_Id, Serie_Id, Nota_Id, NossaEmissao, ChaveNfe, DataDaNota, Movimento, ValorTotal, BaseIcms, Aliquota, ValorIcms, ValorIPI, ValorPIS, ValorCofins, Cfop, situacao, QuantidadeFiscal" & vbCrLf & _
                  "  FROM " & vbCrLf & _
                  "        (SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, " & vbCrLf & _
                  "                NotasFiscais.NossaEmissao, ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento," & vbCrLf & _
                  "                Enc.ValorTotal, Enc.BaseIcms, Enc.ValorIcms, Enc.ValorIPI, Enc.ValorPIS, Enc.ValorCOFINS, Enc.CFOP, Enc.Aliquota,  NotasFiscais.situacao," & vbCrLf & _
                  "                (select sum(notasfiscaisxitens.Quantidadefiscal) " & vbCrLf & _
                  "                   from notasfiscaisxitens" & vbCrLf & _
                  "                  Where NotasFiscais.Empresa_Id      = notasfiscaisxitens.Empresa_Id " & vbCrLf & _
                  "				    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                  "				    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                  "				    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                  "				    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                  "				    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                  "				    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                  "                 ) as quantidadefiscal" & vbCrLf & _
                  "          FROM NotasFiscais " & vbCrLf & _
                  "         Inner Join (" & vbCrLf & _
                  "					SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                  "										   then Valor" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as ValorTotal," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                  "										   then Base" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as BaseIcms," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                  "										   then Percentual" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as Aliquota," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                  "										   then Valor" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as ValorIcms," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                  "										   then Valor" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as ValorIPI," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                  "										   then Valor" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as ValorPIS," & vbCrLf & _
                  "						   isnull(sum(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                  "										   then Valor" & vbCrLf & _
                  "										   else 0" & vbCrLf & _
                  "									   end),0) as ValorCOFINS," & vbCrLf & _
                  "						   isnull(MAX(case" & vbCrLf & _
                  "										 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                  "										   then CFOP_Id" & vbCrLf & _
                  "									   end),0) as CFOP" & vbCrLf & _
                  "					  FROM NotasFiscaisXEncargos" & vbCrLf & _
                  "                     Group by Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id" & vbCrLf & _
                  "                    ) Enc" & vbCrLf & _
                  "            ON NotasFiscais.Empresa_Id      = Enc.Empresa_Id" & vbCrLf & _
                  "           AND NotasFiscais.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf & _
                  "           AND NotasFiscais.Cliente_Id      = Enc.Cliente_Id" & vbCrLf & _
                  "           AND NotasFiscais.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf & _
                  "           AND NotasFiscais.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf & _
                  "           AND NotasFiscais.Serie_Id        = Enc.Serie_Id" & vbCrLf & _
                  "           AND NotasFiscais.Nota_Id         = Enc.Nota_Id" & vbCrLf & _
                  "          LEFT JOIN NFERealizadas" & vbCrLf & _
                  "            ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                  "           AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id" & vbCrLf & _
                  "           AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id" & vbCrLf & _
                  "           AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id" & vbCrLf & _
                  "           AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id" & vbCrLf & _
                  "           AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                  "           AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                  "         WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC')))" & vbCrLf & _
                  "           AND (NotasFiscais.Operacao  > 0)" & vbCrLf & _
                  "           AND (NotasFiscais.Situacao <> 9) " & vbCrLf & _
                  "           AND (NotasFiscais.Empresa_Id =  '" & Empresa(0) & "')" & vbCrLf & _
                  "           AND (NotasFiscais.Movimento  BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "')" & vbCrLf & _
                  "     ) AS Consulta" & vbCrLf & _
                  " WHERE (Cfop NOT BETWEEN 1350 AND 1360)" & vbCrLf & _
                  "   AND (Cfop NOT BETWEEN 2300 AND 2360)" & vbCrLf & _
                  "   AND (Cfop NOT BETWEEN 5350 AND 5360)" & vbCrLf & _
                  "   AND (Cfop NOT BETWEEN 6350 AND 6360)" & vbCrLf & _
                  "   AND (Cfop NOT BETWEEN 1252 AND 1253)" & vbCrLf & _
                  "   AND (Cfop NOT BETWEEN 1302 AND 1303)" & vbCrLf

            '"   AND (Cfop NOT BETWEEN 1550 AND 1560)" & vbCrLf & _
            '"   AND (Cfop NOT BETWEEN 2550 AND 2560)" & vbCrLf & _


            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            '-------------------------------------------
            ' Registro C001  
            '-------------------------------------------
            RegistroC001 += 1

            linha = "|C001"
            linha &= "|" & IIf(ds.Tables(0).Rows.Count > 0, "0", "1") '0 - Com Dados  /  1 - Sem Dados
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1
            '-------------------------------------------

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        Nota = .Item("Nota_Id")
                        Serie = .Item("Serie_Id")
                        EntradaSaida = .Item("EntradaSaida_Id")
                        Cliente = .Item("Cliente_Id")
                        EndCliente = .Item("EndCliente_Id")
                        Emissao = .Item("NossaEmissao")


                        linha = "|C100"
                        If .Item("EntradaSaida_Id") = "E" Then
                            linha &= "|0"
                        Else
                            linha &= "|1"
                        End If

                        If .Item("NossaEmissao") = "S" Then
                            NossaEmissao = True
                            linha &= "|0"
                        Else
                            NossaEmissao = False
                            linha &= "|1"
                        End If

                        'If .Item("ValorTotal") <> 0 Then
                        If .Item("Situacao") = 1 Then    'Situação do Documento Fiscal
                            linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                        Else
                            linha &= "|"
                        End If
                        'Else
                        '    linha &= "|"
                        'End If


                        If Len(.Item("ChaveNfe")) < 10 Then
                            linha &= "|01"
                            Modelo = "01"
                        Else
                            linha &= "|55"
                            Modelo = "55"
                        End If

                        If .Item("Situacao") = 1 Then    'Situação do Documento Fiscal
                            If .Item("QuantidadeFiscal") = 0 Then
                                linha &= "|06"
                            Else
                                linha &= "|00"
                            End If
                            Cancelada = "N"
                        Else
                            linha &= "|02"
                            Cancelada = "S"
                        End If

                        linha &= "|" & RTrim(.Item("Serie_Id"))
                        linha &= "|" & .Item("Nota_Id")
                        linha &= "|" & .Item("ChaveNfe")

                        If Cancelada = "N" Then
                            linha &= "|" & CDate(.Item("DataDaNota")).ToSqlDate()
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|" & CDate(.Item("Movimento")).ToSqlDate()
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorTotal")
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|1"
                        Else
                            linha &= "|"
                        End If

                        linha &= "|"                        'Valor Total do Desconto
                        linha &= "|"                        'Valor Total do Abatimento

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorTotal")
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|1"                       'Identificador do Tipo de Frete
                        Else
                            linha &= "|"
                        End If

                        linha &= "|"                        'Valor do frete indicado no documento fiscal
                        linha &= "|"                        'Valor do seguro indicado no documento fiscal
                        linha &= "|"                        'Valor de outras despesas acessórias

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("BaseIcms")
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorIcms")
                        Else
                            linha &= "|"
                        End If

                        linha &= "|"                        'Valor da base de cálculo do ICMS substituição tributária
                        linha &= "|"                        'Valor do ICMS retido por substituição tributária

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorIPI")
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorPIS")
                        Else
                            linha &= "|"
                        End If

                        If Cancelada = "N" Then
                            linha &= "|" & .Item("ValorCofins")
                        Else
                            linha &= "|"
                        End If

                        linha &= "|"                        'Valor total do PIS retido por substituição tributária
                        linha &= "|"                        'Valor total da COFINS retido por sustituição tributária
                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    RegistroC100 += 1
                    RegistroGeral += 1

                    'Registro C170 - Notas Fiscais X Itens
                    '--------------------------------------------
                    'If Emissao <> "A" And Modelo <> "55" Then
                    If (Modelo <> "55" And Cancelada = "N") Or (Modelo = "55" And Not NossaEmissao) Then
                        Sql = "SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, 'C170' AS Registro, NotasFiscaisXItens.Produto_Id AS Item, " & vbCrLf & _
                              "       Produtos.Nome, NotasFiscais.NossaEmissao, '' AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento, NotasFiscaisXItens.QuantidadeFiscal, Produtos.Unidade," & vbCrLf & _
                              "       NotasFiscaisXItens.CFOP_Id AS CFOP, NotasFiscais.Operacao, NotasFiscais.SubOperacao," & vbCrLf & _
                              "       Enc.SituacaoTributaria, Enc.ValorTotal, Enc.BaseICMS, Enc.Aliquota, Enc.ValorICMS, Enc.ValorIPI, Enc.ValorPIS, Enc.ValorCOFINS, Enc.CST_IPI, Enc.COD_ENQ, Enc.VL_BC_IPI," & vbCrLf & _
                              "       Enc.ALI_IPI, Enc.VL_IPI, Enc.CST_PIS, Enc.VL_BC_PIS, Enc.ALI_PIS, Enc.VL_PIS, Enc.CST_COFINS, Enc.VL_BC_COFINS, Enc.ALI_COFINS, Enc.VL_COFINS," & vbCrLf & _
                              "       '' AS COD_CTA, SubOperacoes.EstoqueFiscal,0.00 AS QUANT_BC_COFINS,0.00 AS ALI_COFINSR, 0.00 AS QUANT_BC_PIS, 0 AS ALI_PISR" & vbCrLf & _
                              "  FROM NotasFiscais " & vbCrLf & _
                              " INNER JOIN NotasFiscaisXItens " & vbCrLf & _
                              "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id " & vbCrLf & _
                              "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf & _
                              "   AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id " & vbCrLf & _
                              "   AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id " & vbCrLf & _
                              "   AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf & _
                              "   AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id " & vbCrLf & _
                              "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id " & vbCrLf & _
                              " INNER JOIN(" & vbCrLf & _
                              "			SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id, Produto_Id, CFOP_ID, Sequencia_Id," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                              "								   then SituacaoTributaria" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as SituacaoTributaria," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as ValorTotal," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                              "								   then isnull(Base,0)" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as BaseIcms," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                              "								   then Percentual" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as Aliquota," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as ValorIcms," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as ValorIPI," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as ValorPIS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as ValorCOFINS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                              "								   then SituacaoTributaria	" & vbCrLf & _
                              "								   else 0" & vbCrLf & _
                              "							   end),0) as CST_IPI," & vbCrLf & _
                              "				   0 AS COD_ENQ," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                              "								   then Base" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_BC_IPI," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                              "								   then Percentual" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as ALI_IPI," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_IPI," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                              "								   then SituacaoTributaria	" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as CST_PIS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                              "								   then Base" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_BC_PIS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                              "								   then Percentual" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as ALI_PIS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_PIS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                              "								   then SituacaoTributaria" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as CST_COFINS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                              "								   then Base" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_BC_COFINS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                              "								   then Percentual	" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as ALI_COFINS," & vbCrLf & _
                              "				   isnull(sum(case" & vbCrLf & _
                              "								 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                              "								   then Valor" & vbCrLf & _
                              "								   else 0 " & vbCrLf & _
                              "							   end),0) as VL_COFINS" & vbCrLf & _
                              "			  FROM NotasFiscaisXEncargos" & vbCrLf & _
                              "			 Group By Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id, Produto_Id, CFOP_ID, Sequencia_Id" & vbCrLf & _
                              "          ) Enc" & vbCrLf & _
                              "    ON NotasFiscaisXItens.Empresa_Id      = Enc.Empresa_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.Cliente_Id      = Enc.Cliente_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.Serie_Id        = Enc.Serie_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.Nota_Id         = Enc.Nota_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.Produto_Id      = Enc.Produto_Id" & vbCrLf & _
                              "   AND NotasFiscaisXItens.CFOP_ID         = Enc.CFOP_ID" & vbCrLf & _
                              "   AND NotasFiscaisXItens.Sequencia_id    = Enc.Sequencia_id" & vbCrLf & _
                              " INNER JOIN Produtos " & vbCrLf & _
                              "    ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
                              " INNER JOIN SubOperacoes " & vbCrLf & _
                              "    ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id " & vbCrLf & _
                              "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf & _
                              " WHERE NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND " & vbCrLf & _
                              "       NotasFiscais.Cliente_Id = '" & Cliente & "' AND NotasFiscais.EndCliente_Id = " & EndCliente & " And " & vbCrLf & _
                              "       NotasFiscais.Nota_Id = " & Nota & " AND NotasFiscais.Serie_Id = '" & Serie & "' And " & vbCrLf & _
                              "       NotasFiscais.EntradaSaida_Id = '" & EntradaSaida & "'" & vbCrLf & _
                              "     AND (NotasFiscais.Operacao > 0) " & vbCrLf & _
                              "     AND (NotasFiscais.Situacao <> 9) " & vbCrLf

                        '"   AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 1550 AND 1560)" & vbCrLf & _
                        '"   AND (NotasFiscaisXItens.Cfop_Id NOT BETWEEN 2550 AND 2560)" & vbCrLf



                        ds1 = Banco.ConsultaDataSet(Sql, "Consulta")
                        IntSequencia = 0

                        If ds1.Tables(0).Rows.Count > 0 Then
                            For Each dr1 As DataRow In ds1.Tables(0).Rows
                                With dr1
                                    linha = "|C170"
                                    IntSequencia += 1
                                    linha &= "|" & IntSequencia
                                    linha &= "|" & Trim(.Item("Item"))
                                    linha &= "|" & Funcoes.EliminarCaracteresEspeciais(.Item("Nome"))
                                    linha &= "|" & .Item("QuantidadeFiscal")
                                    linha &= "|" & Trim(.Item("Unidade"))
                                    linha &= "|" & .Item("ValorTotal")
                                    linha &= "|"                            'Descontos
                                    If .Item("EstoqueFiscal") = "S" Then
                                        linha &= "|0"
                                    Else
                                        linha &= "|1"
                                    End If
                                    linha &= "|" & Funcoes.AlinharDireita(.Item("SituacaoTributaria"), 3, "0")
                                    linha &= "|" & .Item("CFOP")
                                    linha &= "|" & .Item("Operacao") & .Item("SubOperacao")
                                    linha &= "|" & .Item("BaseIcms")
                                    linha &= "|" & .Item("Aliquota")
                                    linha &= "|" & .Item("ValorIcms")
                                    linha &= "|"                            'Valor Base de Icms Substituição Tributária
                                    linha &= "|"                            'Aliquota de Icms por Substituição Tributária
                                    linha &= "|"                            'Valor de Icms por Substituição Tributária
                                    linha &= "|0"                           'Indicador de período de apuração do IPI
                                    linha &= "|"                            'Valor de Icms por Substituição Tributária
                                    linha &= "|"                            'Código de enquadramento legal do IPI, conforme tabela indicada no item 4.5.3
                                    linha &= "|" & .Item("VL_BC_IPI")
                                    linha &= "|" & .Item("ALI_IPI")
                                    linha &= "|" & .Item("VL_IPI")

                                    linha &= "|" ' & ' .Item("CST_PIS")          'Código Da situação tributaria referente ao PIS, conforme tabela indicada no item 4.3.4                            
                                    linha &= "|" & .Item("VL_BC_PIS")
                                    linha &= "|" & .Item("ALI_PIS")
                                    linha &= "|" & .Item("QUANT_BC_PIS")
                                    linha &= "|" & .Item("ALI_PISR")
                                    linha &= "|" & .Item("VL_PIS")

                                    linha &= "|" ' & .Item("CST_COFINS")       'Código Da situação tributaria referente ao COFINS, conforme tabela indicada no item 4.3.5
                                    linha &= "|" & .Item("VL_BC_COFINS")
                                    linha &= "|" & .Item("ALI_COFINS")
                                    linha &= "|" & .Item("QUANT_BC_COFINS")
                                    linha &= "|" & .Item("ALI_COFINSR")
                                    linha &= "|" & .Item("VL_COFINS")
                                    linha &= "|"
                                    linha &= "|"

                                    If .Item("QuantidadeFiscal") > 0 Then
                                        strm.WriteLine(linha)
                                        RegistroC170 += 1
                                        RegistroGeral += 1
                                    End If
                                End With
                            Next
                        End If
                    End If

                    '--------------------------------------------
                    'Registro C190 - Notas Fiscais X Itens
                    '--------------------------------------------
                    If Cancelada = "N" Then
                        Sql = "  Select SituacaoTributaria AS CST_ICMS, Cfop_Id AS CFOP, Percentual AS ALIQ_ICMS, Sum(Produto) as VL_OPR, Sum(Base) as VL_BC_ICMS, Sum(Icms) as VL_ICMS, Sum(Ipi) as VL_IPI "
                        Sql &= " FROM"
                        Sql &= " (Select Produto_Id, SituacaoTributaria, Cfop_Id, Sum(Produto) as Produto, Sum(Base) as Base, sum(Percentual) as Percentual, Sum(Icms) as Icms, Sum(Ipi) as Ipi "
                        Sql &= " FROM"

                        Sql &= " (select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, sum(valor) as Produto, 0 as Base,0 as Percentual,"
                        Sql &= " 0.00 as Icms, 0.0 as IPI"
                        Sql &= "       FROM NotasFiscaisXEncargos"
                        Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                        Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                        Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                        Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'PRODUTO'"
                        Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual"

                        Sql &= " UNION"
                        Sql &= " select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0.00 as Produto, 0 as  Base,0 as Percentual,"
                        Sql &= " 0.00 as Icms, sum(valor) as IPI  "
                        Sql &= "        from NotasFiscaisXEncargos"
                        Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                        Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                        Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                        Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'IPI'"
                        Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, Percentual"

                        Sql &= "  UNION"
                        Sql &= " select Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0 as Produto, Sum(Base) as Base,Percentual,"
                        Sql &= " sum(valor) as Icms, 0.0 as IPI  "
                        Sql &= "       from NotasFiscaisXEncargos"

                        Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                        Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                        Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                        Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id like '%ICMS%'"
                        Sql &= " group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual) as Consulta"
                        Sql &= " Group by Produto_Id, SituacaoTributaria, Cfop_Id) as Resultado"
                        Sql &= " Group by SituacaoTributaria, Cfop_Id, Percentual"

                        ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                        If ds2.Tables(0).Rows.Count > 0 Then
                            For Each dr2 As DataRow In ds2.Tables(0).Rows
                                With dr2
                                    linha = "|C190"
                                    linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                    linha &= "|" & .Item("CFOP")
                                    linha &= "|" & .Item("ALIQ_ICMS")
                                    linha &= "|" & .Item("VL_OPR")
                                    linha &= "|" & .Item("VL_BC_ICMS")
                                    linha &= "|" & .Item("VL_ICMS")
                                    linha &= "|0"
                                    linha &= "|0"

                                    'If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                                    '    linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                    'Else
                                    '    linha &= "|0"
                                    'End If

                                    If CInt(.Item("CST_ICMS")) = 20 Or CInt(.Item("CST_ICMS")) = 70 Then
                                        linha &= "|" & .Item("VL_BC_ICMS")
                                    Else
                                        linha &= "|0"
                                    End If


                                    linha &= "|0" & .Item("VL_IPI")
                                    linha &= "|"
                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                RegistroC190 += 1
                                RegistroGeral += 1
                            Next
                        End If
                    End If
                    '--------------------------------------------
                Next
            End If

            '---------------------------------------------------------------------------------
            'Registro C500 - Nota Fiscal (Código 06), Nota Fiscal de Energia Eletrica), 
            '----------------------------------------------------------------------------------------------------------
            Sql = "SELECT EntradaSaida_Id, Cliente_Id, EndCliente_Id, Serie_Id, Nota_Id, NossaEmissao, ChaveNfe, DataDaNota, Movimento, " & vbCrLf & _
                    "       ValorTotal, BaseIcms, Aliquota, ValorIcms, ValorIPI, ValorPIS, ValorCofins, Cfop" & vbCrLf & _
                    "  FROM (SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, " & vbCrLf & _
                    "               NotasFiscais.NossaEmissao, ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento, " & vbCrLf & _
                    "               Enc.ValorTotal, Enc.BaseIcms, Enc.Aliquota, Enc.ValorIcms, Enc.ValorIPI, Enc.ValorPIS, Enc.ValorCofins, Enc.Cfop " & vbCrLf & _
                    "          FROM NotasFiscais " & vbCrLf & _
                    "         INNER JOIN (" & vbCrLf & _
                    "						SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id, " & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                    "											   then Valor" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as ValorTotal," & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                    "											   then isnull(Base,0)" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as BaseIcms," & vbCrLf & _
                    "							   0 AS Aliquota," & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) like '%ICMS%'" & vbCrLf & _
                    "											   then Valor" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as ValorIcms," & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) = 'IPI'" & vbCrLf & _
                    "											   then Valor" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as ValorIPI," & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) = 'PIS'" & vbCrLf & _
                    "											   then Valor" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as ValorPIS," & vbCrLf & _
                    "							   isnull(sum(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) = 'COFINS'" & vbCrLf & _
                    "											   then Valor" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as ValorCOFINS," & vbCrLf & _
                    "							   isnull(Max(case" & vbCrLf & _
                    "											 when UPPER(Encargo_Id) = 'PRODUTO'" & vbCrLf & _
                    "											   then CFOP_ID	" & vbCrLf & _
                    "											   else 0" & vbCrLf & _
                    "										   end),0) as CFOP" & vbCrLf & _
                    "						  FROM NotasFiscaisXEncargos" & vbCrLf & _
                    "						 Group By Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Nota_Id, Serie_Id" & vbCrLf & _
                    "                     ) Enc" & vbCrLf & _
                    "          ON NotasFiscais.Empresa_Id      = Enc.Empresa_Id" & vbCrLf & _
                    "		  AND NotasFiscais.EndEmpresa_Id   = Enc.EndEmpresa_Id" & vbCrLf & _
                    "		  AND NotasFiscais.Cliente_Id      = Enc.Cliente_Id" & vbCrLf & _
                    "		  AND NotasFiscais.EndCliente_Id   = Enc.EndCliente_Id" & vbCrLf & _
                    "		  AND NotasFiscais.EntradaSaida_Id = Enc.EntradaSaida_Id" & vbCrLf & _
                    "		  AND NotasFiscais.Serie_Id        = Enc.Serie_Id" & vbCrLf & _
                    "		  AND NotasFiscais.Nota_Id         = Enc.Nota_Id" & vbCrLf & _
                    "        LEFT JOIN NFERealizadas " & vbCrLf & _
                    "          ON NotasFiscais.Empresa_Id      = NFERealizadas.Empresa_Id" & vbCrLf & _
                    "         AND NotasFiscais.EndEmpresa_Id   = NFERealizadas.EndEmpresa_Id " & vbCrLf & _
                    "         AND NotasFiscais.Cliente_Id      = NFERealizadas.Cliente_Id " & vbCrLf & _
                    "         AND NotasFiscais.EndCliente_Id   = NFERealizadas.EndCliente_Id " & vbCrLf & _
                    "         AND NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id " & vbCrLf & _
                    "         AND NotasFiscais.Serie_Id        = NFERealizadas.Serie_Id" & vbCrLf & _
                    "         AND NotasFiscais.Nota_Id         = NFERealizadas.Nota_Id" & vbCrLf & _
                    "       WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) " & vbCrLf & _
                    "         AND NotasFiscais.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                    "         AND NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'" & vbCrLf & _
                    "         AND (NotasFiscais.Operacao > 0) " & vbCrLf & _
                    "         AND (NotasFiscais.Situacao <> 9) " & vbCrLf & _
                    "      ) as Consulta" & vbCrLf & _
                    "  Where (CFOP Between 1252 and 1253)" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        Nota = .Item("Nota_Id")
                        Serie = .Item("Serie_Id")
                        EntradaSaida = .Item("EntradaSaida_Id")
                        Cliente = .Item("Cliente_Id")
                        EndCliente = .Item("EndCliente_Id")
                        Emissao = .Item("NossaEmissao")

                        linha = "|C500"
                        linha &= "|0"
                        linha &= "|1"   ' Emissao por terceiros

                        linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                        linha &= "|06"
                        linha &= "|00"  'Situacao do Documento Fiscal

                        linha &= "|" & Trim(.Item("Serie_Id"))
                        linha &= "|"   'Subserie
                        linha &= "|04"  'Classe de Consumo

                        linha &= "|" & .Item("Nota_Id")
                        linha &= "|" & CDate(.Item("DataDaNota")).ToSqlDate()
                        linha &= "|" & CDate(.Item("Movimento")).ToSqlDate()
                        linha &= "|" & .Item("ValorTotal")
                        linha &= "|"                        'Valor Total do Desconto
                        linha &= "|" & .Item("ValorTotal")

                        linha &= "|"                        'Valor total dos servicos nao tributados pelo icms
                        linha &= "|"                        'Valor total cobrado em nome de terceiros
                        linha &= "|"                        'Valor total de outras despesas acessórias

                        linha &= "|" & .Item("BaseIcms")
                        linha &= "|" & .Item("ValorIcms")

                        linha &= "|"                        'Valor da base de cálculo do ICMS substituição tributária
                        linha &= "|"                        'Valor do ICMS retido por substituição tributária

                        linha &= "|"                        'Codigo da Informacao Complementar
                        linha &= "|"                        'Valor do Pis
                        linha &= "|"                        'Valor da Cofins
                        linha &= "|3"                        'Código de Tipo de Ligação
                        linha &= "|12"                        'Código de Grupo de Tensão

                        linha &= "|"
                    End With

                    strm.WriteLine(linha)
                    RegistroC500 += 1
                    RegistroGeral += 1

                    'Registro C510 - Notas Fiscais X Itens
                    '--------------------------------------------
                    Emissao = "A"
                    If Emissao <> "A" Then
                        Sql = "   SELECT  NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, 'C170' AS Registro, "
                        Sql &= "          NotasFiscaisXItens.Produto_Id AS Item, Produtos.Nome, NotasFiscais.NossaEmissao, '' AS ChaveNfe, NotasFiscais.DataDaNota, NotasFiscais.Movimento, NotasFiscaisXItens.QuantidadeFiscal, Produtos.Unidade,"

                        Sql &= "          ISNULL ((SELECT    SituacaoTributaria AS Valor"
                        Sql &= "                   FROM      NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                   WHERE     (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                             AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                             (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                             (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PRODUTO')), 0) AS SituacaoTributaria, NotasFiscaisXItens.CFOP_Id as CFOP, NotasFiscais.Operacao, NotasFiscais.SubOperacao,  "

                        Sql &= "          ISNULL ((SELECT    SUM(Valor) AS Valor"
                        Sql &= "                   FROM      NotasFiscaisXEncargos"
                        Sql &= "                   WHERE     (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                             AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                             (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                             (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'Produto')), 0) AS ValorTotal, "

                        Sql &= "           ISNULL  ((SELECT  SUM(Base) AS Valor"
                        Sql &= "                     FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                     WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                            AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                            (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                            (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id like '%ICMS%')), 0) AS BaseIcms, "

                        Sql &= "         ISNULL   ((SELECT  Percentual AS Valor"
                        Sql &= "                    FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                    WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id like '%ICMS%')), 0) AS Aliquota, "

                        Sql &= "          ISNULL ((SELECT  SUM(Valor) AS Valor"
                        Sql &= "                   FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_4"
                        Sql &= "              WHERE        (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id like '%ICMS%')), 0) AS ValorIcms, "

                        Sql &= "          ISNULL ((SELECT  SUM(Valor) AS Valor"
                        Sql &= "                FROM       NotasFiscaisXEncargos AS NotasFiscaisXEncargos_3"
                        Sql &= "                WHERE      (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'IPI')), 0) AS ValorIPI, "

                        Sql &= "          ISNULL ((SELECT  SUM(Valor) AS Valor"
                        Sql &= "                   FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_2"
                        Sql &= "                   WHERE   (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PIS')), 0) AS ValorPIS, "

                        Sql &= "          ISNULL ((SELECT  SUM(Valor) AS Valor"
                        Sql &= "                   FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_1"
                        Sql &= "                   WHERE   (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'COFINS')), 0) AS ValorCofins,"

                        Sql &= "         ISNULL ((SELECT   SituacaoTributaria AS Valor"
                        Sql &= "                  FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                  WHERE   (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'IPI')), 0) AS CST_IPI, 0 as COD_ENQ,"

                        Sql &= "         ISNULL ((SELECT  Base AS Valor"
                        Sql &= "                  FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                  WHERE   (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'IPI')), 0) AS VL_BC_IPI,"

                        Sql &= "         ISNULL  ((SELECT Percentual AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                WHERE     (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'IPI')), 0) AS ALI_IPI,"

                        Sql &= "         ISNULL ((SELECT  Valor AS Valor"
                        Sql &= "                  FROM    NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                  WHERE   (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'IPI')), 0) AS VL_IPI,"

                        Sql &= "         ISNULL ((SELECT  SituacaoTributaria AS Valor"
                        Sql &= "                   FROM   NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                   WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PIS')), 0) AS CST_PIS,"

                        Sql &= "        ISNULL ((SELECT   Base AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE    (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PIS')), 0) AS VL_BC_PIS,"

                        Sql &= "        ISNULL ((SELECT   Percentual AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE    (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PIS')), 0) AS ALI_PIS, 0 AS QUANT_BC_PIS, 0 AS ALI_PISR,"

                        Sql &= "        ISNULL ((SELECT   Valor AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "               WHERE      (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'PIS')), 0) AS VL_PIS,"

                        Sql &= "        ISNULL ((SELECT   SituacaoTributaria AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE    (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'COFINS')), 0) AS CST_COFINS,"

                        Sql &= "        ISNULL ((SELECT   Base AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE    (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'COFINS')), 0) AS VL_BC_COFINS,"

                        Sql &= "        ISNULL ((SELECT   Percentual AS Valor"
                        Sql &= "                 FROM     NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE    (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                          AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                          (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                          (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'COFINS')), 0) AS ALI_COFINS, 0 AS QUANT_BC_COFINS, 0 AS ALI_COFINSR,"

                        Sql &= "        ISNULL ((SELECT    Valor AS Valor"
                        Sql &= "                 FROM      NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
                        Sql &= "                 WHERE     (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
                        Sql &= "                           AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
                        Sql &= "                           (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
                        Sql &= "                           (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscaisXItens.Produto_Id = Produto_Id) AND (Encargo_Id = 'COFINS')), 0) AS VL_COFINS, '' AS COD_CTA, SubOperacoes.EstoqueFiscal"


                        Sql &= "        FROM     NotasFiscais INNER JOIN"
                        Sql &= "                 NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND "
                        Sql &= "                 NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND "
                        Sql &= "                 NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND "
                        Sql &= "                 NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN"
                        Sql &= "                 Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id INNER JOIN"
                        Sql &= "                 SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id"

                        Sql &= " WHERE NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
                        Sql &= "       NotasFiscais.Cliente_Id = '" & Cliente & "' AND NotasFiscais.EndCliente_Id = " & EndCliente & " And "
                        Sql &= "       NotasFiscais.Nota_Id = " & Nota & " AND NotasFiscais.Serie_Id = '" & Serie & "' And "
                        Sql &= "       NotasFiscais.EntradaSaida_Id = '" & EntradaSaida & "'"
                        Sql &= "   AND (NotasFiscais.Operacao > 0) "
                        Sql &= "   AND (NotasFiscais.Situacao <> 9) "

                        ds1 = Banco.ConsultaDataSet(Sql, "Consulta")
                        IntSequencia = 0

                        If ds1.Tables(0).Rows.Count > 0 Then
                            For Each dr1 As DataRow In ds1.Tables(0).Rows
                                With dr1
                                    linha = "|C510"
                                    IntSequencia += 1

                                    linha &= "|" & IntSequencia
                                    linha &= "|" & Trim(.Item("Item"))
                                    linha &= "|00"      'Código de Classificação do Item

                                    linha &= "|" & .Item("QuantidadeFiscal")
                                    linha &= "|" & Trim(.Item("Unidade"))
                                    linha &= "|" & .Item("ValorTotal")
                                    linha &= "|"                            'Descontos
                                    linha &= "|" & Funcoes.AlinharDireita(.Item("SituacaoTributaria"), 3, "0")
                                    linha &= "|" & .Item("CFOP")
                                    linha &= "|" & .Item("BaseIcms")
                                    linha &= "|" & .Item("Aliquota")
                                    linha &= "|" & .Item("ValorIcms")
                                    linha &= "|"                            'Valor Base de Icms Substituição Tributária
                                    linha &= "|"                            'Aliquota de Icms por Substituição Tributária
                                    linha &= "|"                            'Valor de Icms por Substituição Tributária

                                    linha &= "|0"                           'Indicador do tipo de Receita
                                    linha &= "|"                            'Codigo Do participante Receptor da Receita de Terceiros

                                    linha &= "|" & .Item("VL_PIS")
                                    linha &= "|" & .Item("VL_COFINS")
                                    linha &= "|111010001"
                                    linha &= "|"
                                End With

                                strm.WriteLine(linha)
                                RegistroC510 += 1
                                RegistroGeral += 1
                            Next
                        End If
                    End If

                    '--------------------------------------------
                    'Registro C590 - Notas Fiscais X Itens
                    '--------------------------------------------
                    Sql = "  Select SituacaoTributaria AS CST_ICMS, Cfop_Id AS CFOP, Percentual AS ALIQ_ICMS, Sum(Produto) as VL_OPR, Sum(Base) as VL_BC_ICMS, Sum(Icms) as VL_ICMS, Sum(Ipi) as VL_IPI "
                    Sql &= " FROM"
                    Sql &= " (Select Produto_Id, SituacaoTributaria, Cfop_Id, Sum(Produto) as Produto, Sum(Base) as Base, sum(Percentual) as Percentual, Sum(Icms) as Icms, Sum(Ipi) as Ipi "
                    Sql &= " FROM"

                    Sql &= " (select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, sum(valor) as Produto, 0 as Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, 0.0 as IPI"
                    Sql &= "       FROM NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'PRODUTO'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual"

                    Sql &= " UNION"
                    Sql &= " select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0.00 as Produto, 0 as  Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, sum(valor) as IPI  "
                    Sql &= "        from NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'IPI'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, Percentual"

                    Sql &= "  UNION"
                    Sql &= " select Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0 as Produto, Sum(Base) as Base,Percentual,"
                    Sql &= " sum(valor) as Icms, 0.0 as IPI  "
                    Sql &= "       from NotasFiscaisXEncargos"

                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id like '%ICMS%'"
                    Sql &= " group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual) as Consulta"
                    Sql &= " Group by Produto_Id, SituacaoTributaria, Cfop_Id) as Resultado"
                    Sql &= " Group by SituacaoTributaria, Cfop_Id, Percentual"

                    ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds2.Tables(0).Rows.Count > 0 Then
                        For Each dr2 As DataRow In ds2.Tables(0).Rows
                            With dr2
                                linha = "|C590"
                                linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                linha &= "|" & .Item("CFOP")
                                linha &= "|" & .Item("ALIQ_ICMS")
                                linha &= "|" & .Item("VL_OPR")
                                linha &= "|" & .Item("VL_BC_ICMS")
                                linha &= "|" & .Item("VL_ICMS")
                                linha &= "|0" '& .Item("VL_BC_ICMS_ST")
                                linha &= "|0" '& .Item("VL_ICMS_ST")

                                If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                                    linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                Else
                                    linha &= "|0" '& .Item("VL_RED_BC")
                                End If

                                linha &= "|" ' Código de Observacao
                                linha &= "|"
                            End With

                            strm.WriteLine(linha)
                            RegistroC590 += 1
                            RegistroGeral += 1
                        Next
                    End If
                    '--------------------------------------------
                Next
            End If


            '-------------------------------------------
            ' Registro C990  - Encerramento do Bloco C
            '-------------------------------------------

            RegistroC990 += 1

            linha = "|C990"
            linha &= "|" & RegistroC001 + RegistroC100 + RegistroC170 + RegistroC190 + RegistroC500 + RegistroC510 + RegistroC590 + RegistroC990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1


            '-------------------------------------------
            ' Registro D001  - Abertura do Bloco D
            '-------------------------------------------
            linha = "|D001"
            linha &= "|" & BlocoD       'Indicador de Movimento - 0 - Bloco com Dados Informados
            '                                                     1 - Bloco sem dados informados
            linha &= "|"

            strm.WriteLine(linha)
            RegistroD001 += 1
            RegistroGeral += 1

            '--------------------------------------------
            'Registro D100 - Nota Fiscal de Serviço de Transporte
            '--------------------------------------------

            Sql = "   SELECT * From ("
            Sql &= "  SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe, NotasFiscais.DataDaNota, "
            Sql &= "        NotasFiscais.Movimento, NotasFiscais.Situacao, NotasFiscais.NossaEmissao, NotasFiscais.Eletronica, "

            Sql &= "   ISNULL ((SELECT  SUM(Valor) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'Produto')), 0) AS ValorTotal, "

            Sql &= "   ISNULL ((SELECT SUM(Base) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id like '%ICMS%')), 0) AS BaseIcms, 0 AS Aliquota, "

            Sql &= "   ISNULL ((SELECT SUM(Valor) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos AS NotasFiscaisXEncargos_4"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id like '%ICMS%')), 0) AS ValorIcms, "

            Sql &= "		isnull      ((SELECT  Top 1   Cfop_Id AS Valor"
            Sql &= "        FROM        NotasFiscaisXEncargos"
            Sql &= "        WHERE      (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "                    AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "                    (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "                    (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'PRODUTO')), 0) AS Cfop"

            Sql &= " FROM           NotasFiscais LEFT OUTER JOIN"
            Sql &= "                NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "                NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "                NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "                NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'"
            Sql &= "       And Notasfiscais.Situacao not in (2,9) and Notasfiscais.operacao > 0 ) as Consulta"

            Sql &= "  Where (CFOP Between 1350 and 1360) Or  (CFOP Between 2350 and 2360) Or "
            Sql &= "        (CFOP Between 5350 and 5360) Or  (CFOP Between 6350 and 6360)"

            ds = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        Nota = .Item("Nota_Id")
                        Serie = .Item("Serie_Id")
                        EntradaSaida = .Item("EntradaSaida_Id")
                        Cliente = .Item("Cliente_Id")
                        EndCliente = .Item("EndCliente_Id")

                        linha = "|D100"
                        If .Item("Cfop") < 5000 Then
                            linha &= "|0" 'Aquisicao
                        Else
                            linha &= "|1" 'Prestacao
                        End If

                        If .Item("NossaEmissao") = "S" Then
                            linha &= "|0" 'Emissao Propria
                        Else
                            linha &= "|1" 'Terceiros
                        End If

                        linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                        If .Item("Cfop") < 5000 Then
                            linha &= "|08"
                        Else
                            linha &= "|57"                      'Código do modelo do documento fiscal, conforme a Tabela 4.1.1
                        End If
                        If .Item("Situacao") = 9 Then       '05 NF-e ou CT-e - Numeração inutilizada, 04 NF-e ou CT-e - denegado
                            linha &= "|05"                  '06 Documento Fiscal Complementar, 07 Escrituração extemporânea de documento complementar
                        ElseIf .Item("Situacao") = 2 Then   '02 Documento cancelado, 03 Escrituração extemporânea de documento cancelado
                            linha &= "|02"                  '08 Documento Fiscal emitido com base em Regime Especial ou Norma Específica
                        Else                                '00 Documento regular, 01 Escrituração extemporânea de documento regular
                            linha &= "|00"
                        End If

                        linha &= "|" & Trim(.Item("Serie_Id"))
                        linha &= "|"                        'SubSérie do Documento Fiscal
                        linha &= "|" & .Item("Nota_Id")
                        If .Item("NossaEmissao") = "S" And .Item("Eletronica") = "S" Then
                            linha &= "|" & .Item("ChaveNfe")    'Chave do Conhecimento de Transporte eletronico
                        Else
                            linha &= "|"
                        End If
                        linha &= "|" & CDate(.Item("DataDaNota")).ToSqlDate()
                        linha &= "|" & CDate(.Item("Movimento")).ToSqlDate()

                        linha &= "|1"                       'Tipo de CT-e quando o modelo do documento for 57
                        linha &= "|"                        'Chave do CT-e Informar Campo vazio

                        linha &= "|" & .Item("ValorTotal")
                        linha &= "|"                        'Valor Total do Desconto
                        linha &= "|1"                       'Indicador do Tipo de Frete 0 - Por conta de Terceiros
                        '                                                               1 - Por Conta do Emitente
                        '                                                               2 - Por Conta do Destinatário
                        '                                                               9 - Sem Frete
                        linha &= "|" & .Item("ValorTotal")
                        linha &= "|" & .Item("BaseIcms")
                        linha &= "|" & .Item("ValorIcms")
                        linha &= "|"                        'Valor não Tributado
                        linha &= "|"                        'Código da informação complementar do documento fiscal - campo02 da tabela 450
                        linha &= "|111010001"               'Codigo da Conta Analitica Contabil Debitada/Creditada
                        linha &= "|"

                    End With
                    strm.WriteLine(linha)
                    RegistroD100 += 1
                    RegistroGeral += 1

                    '-------------------------------------------------
                    'Registro D190 - Registro Analitico dos Documentos
                    '-------------------------------------------------

                    Sql = "  Select SituacaoTributaria AS CST_ICMS, Cfop_Id AS CFOP, Percentual AS ALIQ_ICMS, Sum(Produto) as VL_OPR, Sum(Base) as VL_BC_ICMS, Sum(Icms) as VL_ICMS, Sum(Ipi) as VL_IPI "
                    Sql &= " FROM"
                    Sql &= " (Select Produto_Id, SituacaoTributaria, Cfop_Id, Sum(Produto) as Produto, Sum(Base) as Base, sum(Percentual) as Percentual, Sum(Icms) as Icms, Sum(Ipi) as Ipi "
                    Sql &= " FROM"

                    Sql &= " (select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, sum(valor) as Produto, 0 as Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, 0.0 as IPI"
                    Sql &= "       FROM NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'PRODUTO'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual"

                    Sql &= " UNION"
                    Sql &= " select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0.00 as Produto, 0 as  Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, sum(valor) as IPI  "
                    Sql &= "        from NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'IPI'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, Percentual"

                    Sql &= "  UNION"
                    Sql &= " select Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0 as Produto, Sum(Base) as Base,Percentual,"
                    Sql &= " sum(valor) as Icms, 0.0 as IPI  "
                    Sql &= "       from NotasFiscaisXEncargos"

                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id like '%ICMS%'"
                    Sql &= " group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual) as Consulta"
                    Sql &= " Group by Produto_Id, SituacaoTributaria, Cfop_Id) as Resultado"
                    Sql &= " Group by SituacaoTributaria, Cfop_Id, Percentual"

                    ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds1.Tables(0).Rows.Count > 0 Then
                        For Each dr1 As DataRow In ds1.Tables(0).Rows
                            With dr1
                                linha = "|D190"
                                linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                linha &= "|" & .Item("CFOP")
                                linha &= "|" & .Item("ALIQ_ICMS")
                                linha &= "|" & .Item("VL_OPR")
                                linha &= "|" & .Item("VL_BC_ICMS")
                                linha &= "|" & .Item("VL_ICMS")

                                'If .Item("VL_BC_ICMS") <> 0 And (.Item("VL_OPR") > .Item("VL_BC_ICMS")) Then
                                '    linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                'Else
                                '   linha &= "|0" '& .Item("VL_RED_BC")
                                'End If

                                If CInt(.Item("CST_ICMS")) = 20 Or CInt(.Item("CST_ICMS")) = 70 Then
                                    linha &= "|" & .Item("VL_BC_ICMS")
                                Else
                                    linha &= "|0" '& .Item("VL_RED_BC")
                                End If

                                linha &= "|"                        ' Códifo da observação do lançamento fiscal
                                linha &= "|"
                            End With
                            strm.WriteLine(linha)
                            RegistroD190 += 1
                            RegistroGeral += 1
                        Next
                    End If
                    '--------------------------------------------
                Next
            End If

            '-----------------------------------------------------
            'Registro D500 - Nota Fiscal de Serviço de Comunicação
            '-----------------------------------------------------

            Sql = " Select * From (SELECT NotasFiscais.EntradaSaida_Id, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id, ISNULL(NFERealizadas.ChaveNfe, '') AS ChaveNfe, NotasFiscais.DataDaNota, "
            Sql &= "        NotasFiscais.Movimento, "

            Sql &= "   ISNULL ((SELECT  Top 1 CFOP_Id AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'Produto')), 0) AS CFOP, "

            Sql &= "   ISNULL ((SELECT  SUM(Valor) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id = 'Produto')), 0) AS ValorTotal, "

            Sql &= "   ISNULL ((SELECT SUM(Base) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos AS NotasFiscaisXEncargos_5"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id like '%ICMS%')), 0) AS BaseIcms, 0 AS Aliquota, "

            Sql &= "   ISNULL ((SELECT SUM(Valor) AS Valor"
            Sql &= "        FROM   NotasFiscaisXEncargos AS NotasFiscaisXEncargos_4"
            Sql &= "        WHERE  (NotasFiscais.Nota_Id = Nota_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Empresa_Id = Empresa_Id) "
            Sql &= "               AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND (NotasFiscais.Cliente_Id = Cliente_Id) AND "
            Sql &= "               (NotasFiscais.EndCliente_Id = EndCliente_Id) AND (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND "
            Sql &= "               (NotasFiscais.Serie_Id = Serie_Id) AND (Encargo_Id like '%ICMS%')), 0) AS ValorIcms"

            Sql &= "   FROM NotasFiscais LEFT OUTER JOIN"
            Sql &= "        NFERealizadas ON NotasFiscais.Empresa_Id = NFERealizadas.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NFERealizadas.EndEmpresa_Id AND "
            Sql &= "        NotasFiscais.Cliente_Id = NFERealizadas.Cliente_Id AND NotasFiscais.EndCliente_Id = NFERealizadas.EndCliente_Id AND "
            Sql &= "        NotasFiscais.EntradaSaida_Id = NFERealizadas.EntradaSaida_Id AND NotasFiscais.Serie_Id = NFERealizadas.Serie_Id AND "
            Sql &= "        NotasFiscais.Nota_Id = NFERealizadas.Nota_Id"

            Sql &= " WHERE (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) and   NotasFiscais.Empresa_Id = '" & Empresa(0) & "' AND "
            Sql &= "       NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "'"
            Sql &= "       And Notasfiscais.Situacao <> 9 and Notasfiscais.operacao > 0) as Consulta "

            Sql &= "  Where (CFOP Between 1300 and 1350) Or  (CFOP Between 2300 and 2350) Or "
            Sql &= "        (CFOP Between 5300 and 5350) Or  (CFOP Between 6300 and 6350)"

            ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds1.Tables(0).Rows.Count > 0 Then
                For Each dr1 As DataRow In ds1.Tables(0).Rows
                    With dr1
                        Nota = .Item("Nota_Id")
                        Serie = .Item("Serie_Id")
                        EntradaSaida = .Item("EntradaSaida_Id")
                        Cliente = .Item("Cliente_Id")
                        EndCliente = .Item("EndCliente_Id")

                        linha = "|D500"
                        linha &= "|0"
                        linha &= "|1"
                        linha &= "|" & .Item("Cliente_Id") & .Item("EndCliente_Id")
                        linha &= "|22"
                        linha &= "|00"                      '00 - Documento Regular 02 - Cancelado
                        linha &= "|" & Trim(.Item("Serie_Id"))
                        linha &= "|"                        'SubSérie do Documento Fiscal
                        linha &= "|" & .Item("Nota_Id")
                        linha &= "|" & CDate(.Item("DataDaNota")).ToSqlDate()
                        linha &= "|" & CDate(.Item("Movimento")).ToSqlDate()

                        linha &= "|" & .Item("ValorTotal")
                        linha &= "|"                        'Valor Total do Desconto
                        linha &= "|" & .Item("ValorTotal")

                        linha &= "|"                        'Valor Total dos Servicos nao Tributados Pelo ICMS
                        linha &= "|"                        'Valor Cobrado em Nome de Terceiros
                        linha &= "|"                        'Valor de Outras Despesas Indicadas no Documento Fiscal

                        linha &= "|" & .Item("BaseIcms")
                        linha &= "|" & .Item("ValorIcms")
                        linha &= "|"                        'Codigo da Informacao Complementar
                        linha &= "|"                        'Valor do Pis
                        linha &= "|"                        'Valor da Cofins
                        linha &= "|111010001"               'Codigo da Conta Analitica Contabil Debitada/Creditada
                        linha &= "|1"                       'Tipo de Assinante 1 - Comercial/Industrial
                        linha &= "|"

                    End With
                    strm.WriteLine(linha)
                    RegistroD500 += 1
                    RegistroGeral += 1

                    '-------------------------------------------------
                    'Registro D590 - Registro Analitico dos Documentos
                    '-------------------------------------------------

                    Sql = "  Select SituacaoTributaria AS CST_ICMS, Cfop_Id AS CFOP, Percentual AS ALIQ_ICMS, Sum(Produto) as VL_OPR, Sum(Base) as VL_BC_ICMS, Sum(Icms) as VL_ICMS, Sum(Ipi) as VL_IPI "
                    Sql &= " FROM"
                    Sql &= " (Select Produto_Id, SituacaoTributaria, Cfop_Id, Sum(Produto) as Produto, Sum(Base) as Base, sum(Percentual) as Percentual, Sum(Icms) as Icms, Sum(Ipi) as Ipi "
                    Sql &= " FROM"

                    Sql &= " (select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, sum(valor) as Produto, 0 as Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, 0.0 as IPI"
                    Sql &= "       FROM NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'PRODUTO'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual"

                    Sql &= " UNION"
                    Sql &= " select 'ICMS' AS Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0.00 as Produto, 0 as  Base,0 as Percentual,"
                    Sql &= " 0.00 as Icms, sum(valor) as IPI  "
                    Sql &= "        from NotasFiscaisXEncargos"
                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id = 'IPI'"
                    Sql &= " Group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id, Percentual"

                    Sql &= "  UNION"
                    Sql &= " select Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,0 as Produto, Sum(Base) as Base,Percentual,"
                    Sql &= " sum(valor) as Icms, 0.0 as IPI  "
                    Sql &= "       from NotasFiscaisXEncargos"

                    Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND "
                    Sql &= "       Cliente_Id = '" & Cliente & "' AND EndCliente_Id = " & EndCliente & " And "
                    Sql &= "       Nota_Id = " & Nota & " AND Serie_Id = '" & Serie & "' And "
                    Sql &= "       EntradaSaida_Id = '" & EntradaSaida & "' and Encargo_Id like '%ICMS%'"
                    Sql &= " group by Encargo_Id,Produto_Id,Situacaotributaria,Cfop_Id,Percentual) as Consulta"
                    Sql &= " Group by Produto_Id, SituacaoTributaria, Cfop_Id) as Resultado"
                    Sql &= " Group by SituacaoTributaria, Cfop_Id, Percentual"

                    ds2 = Banco.ConsultaDataSet(Sql, "Consulta")

                    If ds2.Tables(0).Rows.Count > 0 Then
                        For Each dr2 As DataRow In ds2.Tables(0).Rows
                            With dr2
                                linha = "|D590"
                                linha &= "|" & Funcoes.AlinharDireita(.Item("CST_ICMS"), 3, "0")
                                linha &= "|" & .Item("CFOP")
                                linha &= "|" & .Item("ALIQ_ICMS")
                                linha &= "|" & .Item("VL_OPR")
                                linha &= "|" & .Item("VL_BC_ICMS")
                                linha &= "|" & .Item("VL_ICMS")
                                linha &= "|0"                        ' Base de Calculo ICms Substituição Tributária
                                linha &= "|0"                        ' Valor Icms Substituição Tributária

                                If .Item("VL_BC_ICMS") <> 0 Then
                                    linha &= "|" & .Item("VL_OPR") - .Item("VL_BC_ICMS")
                                Else
                                    linha &= "|0"        '& .Item("VL_RED_BC")
                                End If

                                linha &= "|"                        ' Códifo da observação do lançamento fiscal
                                linha &= "|"
                            End With
                            strm.WriteLine(linha)
                            RegistroD590 += 1
                            RegistroGeral += 1
                        Next
                    End If
                    '--------------------------------------------
                Next
            End If

            '-------------------------------------------
            ' Registro D990  - Encerramento do Bloco D
            '-------------------------------------------
            RegistroD990 += 1

            linha = "|D990"
            linha &= "|" & RegistroD001 + RegistroD100 + RegistroD190 + RegistroD500 + RegistroD590 + RegistroD990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1

            '-------------------------------------------
            ' Registro E001  - Abertura do Bloco E
            '-------------------------------------------
            linha = "|E001"
            linha &= "|0"       'Indicador de Movimento - 0 - Bloco com Dados Informados
            '                                             1 - Bloco sem dados informados
            linha &= "|"

            strm.WriteLine(linha)
            RegistroE001 += 1
            RegistroGeral += 1

            '---------------------------------------------
            ' Registro E100  - Periodo da Apuração do ICMS
            '-------------------------------------------
            linha = "|E100"
            linha &= "|" & CDate(txtDataInicial.Text).ToSqlDate()                                                     'Data Inicial
            linha &= "|" & CDate(txtDataFinal.Text).ToSqlDate()                                                     'Data Final
            linha &= "|"

            strm.WriteLine(linha)
            RegistroE100 += 1
            RegistroGeral += 1

            '-------------------------------------------------------
            ' Registro E110  - Apuração do ICms - Operações Próprias
            '-------------------------------------------------------

            Sql = "  SELECT 'E110' AS Reg,     "
            Sql &= " SUM(CASE WHEN Codigo_Id = 1 THEN Valor ELSE 0 END) AS VL_TOT_DEBITOS,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 2 THEN Valor ELSE 0 END) AS VL_AJ_DEBITOS,"
            Sql &= " SUM(CASE WHEN Codigo_Id = 2 THEN Valor ELSE 0 END) AS VL_TOT_AJ_DEBITOS,    "

            Sql &= " SUM(CASE WHEN Codigo_Id = 3 THEN Valor ELSE 0 END) AS VL_ESTORNOS_CRED,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 5 THEN Valor ELSE 0 END) AS VL_TOT_CREDITOS,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 6 THEN Valor ELSE 0 END) AS VL_AJ_CREDITOS,    "

            Sql &= " SUM(CASE WHEN Codigo_Id = 6 THEN Valor ELSE 0 END) AS VL_TOT_AJ_CREDITOS,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 7 THEN Valor ELSE 0 END) AS VL_ESTORNO_DEB,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 9 THEN Valor ELSE 0 END) AS VL_SLD_CREDOR_ANT,    "

            Sql &= " SUM(CASE WHEN Codigo_Id = 11 THEN Valor ELSE 0 END) AS VL_SLD_APURADO,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 12 THEN Valor ELSE 0 END) AS VL_TOT_DEB,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 13 THEN Valor ELSE 0 END) AS VL_ICMS_RECOLHER,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 14 THEN Valor ELSE 0 END) AS VL_SLD_CREDOR_TRANSPORTAR,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 14 THEN 0 ELSE 0 END) AS DEB_ESP    "
            Sql &= "  FROM ResumoRaIcms"
            Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcesso.SelectedValue & ")"
            Sql &= " GROUP BY Processo_Id"

            Dim ds24 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds24.Tables(0).Rows.Count > 0 Then
                For Each dr24 As DataRow In ds24.Tables(0).Rows
                    With dr24
                        'linha = "|E110"
                        'linha &= "|" & .Item("VL_TOT_DEBITOS")
                        'linha &= "|" & .Item("VL_AJ_DEBITOS")
                        'linha &= "|0" '& .Item("VL_TOT_AJ_DEBITOS")

                        'linha &= "|" & .Item("VL_ESTORNOS_CRED")
                        'linha &= "|" & .Item("VL_TOT_CREDITOS")
                        'linha &= "|" & .Item("VL_AJ_CREDITOS")

                        'linha &= "|" & .Item("VL_TOT_AJ_CREDITOS")
                        'linha &= "|" & .Item("VL_ESTORNO_DEB")
                        'linha &= "|" & .Item("VL_SLD_CREDOR_ANT")

                        'linha &= "|" & .Item("VL_SLD_APURADO")
                        'linha &= "|0" '& .Item("VL_TOT_DEBITOS")  '& .Item("VL_TOT_DEB")
                        'linha &= "|" & .Item("VL_ICMS_RECOLHER")

                        'IcmsARecolher = .Item("VL_ICMS_RECOLHER")

                        'linha &= "|" & .Item("VL_SLD_CREDOR_TRANSPORTAR")
                        'linha &= "|" & .Item("DEB_ESP")

                        'linha &= "|"

                        linha = "|E110"
                        linha &= "|" & .Item("VL_TOT_DEBITOS")
                        linha &= "|" & .Item("VL_AJ_DEBITOS")
                        linha &= "|" & .Item("VL_TOT_AJ_DEBITOS")

                        linha &= "|" & .Item("VL_ESTORNOS_CRED")
                        linha &= "|" & .Item("VL_TOT_CREDITOS")
                        linha &= "|" & .Item("VL_AJ_CREDITOS")

                        linha &= "|" & .Item("VL_TOT_AJ_CREDITOS")
                        linha &= "|" & .Item("VL_ESTORNO_DEB")
                        linha &= "|" & .Item("VL_SLD_CREDOR_ANT")

                        linha &= "|" & .Item("VL_SLD_APURADO")
                        linha &= "|" & .Item("VL_TOT_DEB")
                        linha &= "|" & .Item("VL_ICMS_RECOLHER")

                        IcmsARecolher = .Item("VL_ICMS_RECOLHER")

                        linha &= "|" & .Item("VL_SLD_CREDOR_TRANSPORTAR")
                        linha &= "|" & .Item("DEB_ESP")

                        linha &= "|"

                        strm.WriteLine(linha)
                        RegistroE110 += 1
                        RegistroGeral += 1
                    End With
                Next
            End If

            '-----------------------------------------------------------------
            ' Registro E111  - Ajustes/Beneficio/incentivo - Apuracao do Icms
            '-----------------------------------------------------------------
            Sql = " SELECT    "
            Sql &= " 'E111' AS Reg,  "
            Sql &= " ROW_NUMBER() OVER (PARTITION BY ResumoItensRAICMS.Processo_Id,ResumoItensRAICMS.Codigo_Id,ResumoItensRAICMS.Descricao order by ResumoItensRAICMS.Codigo_Id) AS Sequencia_Id, "
            Sql &= " CASE  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 2 THEN AjustesApuracaoIcms_Id     "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 3 THEN AjustesApuracaoIcms_Id     "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 6 THEN AjustesApuracaoIcms_Id     "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 7 THEN AjustesApuracaoIcms_Id     "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN AjustesApuracaoIcms_Id    "
            Sql &= " END AS COD_AJ_APUR,  "
            Sql &= " CASE "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 2 THEN AjustesDaApuracaoIcms.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 3 THEN AjustesDaApuracaoIcms.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 6 THEN AjustesDaApuracaoIcms.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 7 THEN AjustesDaApuracaoIcms.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN AjustesDaApuracaoIcms.Descricao "
            Sql &= " END AS DESCR_AJ_APUR,    "
            Sql &= " CASE "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 2 THEN ResumoItensRAICMS.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 3 THEN ResumoItensRAICMS.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 6 THEN ResumoItensRAICMS.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 7 THEN ResumoItensRAICMS.Descricao  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN ResumoItensRAICMS.Descricao "
            Sql &= " END AS DESCR_COMPL_AJ,    "
            Sql &= " SUM(CASE  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 2 THEN Valor  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 3 THEN Valor  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 6 THEN Valor  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 7 THEN Valor  "
            Sql &= " WHEN ResumoItensRAICMS.Codigo_Id = 12 THEN Valor "
            Sql &= " ELSE 0 END) AS VL_AJ_APUR    "
            Sql &= " FROM         ResumoItensRAICMS "
            Sql &= " LEFT JOIN  AjustesDaApuracaoIcms ON ResumoItensRAICMS.AjustesApuracaoIcms_Id = AjustesDaApuracaoIcms.Codigo_Id "
            Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND Processo_Id = " & DdlProcesso.SelectedValue & " "
            Sql &= " GROUP BY "
            Sql &= " ResumoItensRAICMS.Processo_Id,ResumoItensRAICMS.Codigo_Id,ResumoItensRAICMS.Descricao, "
            Sql &= " ResumoItensRAICMS.AjustesApuracaoIcms_Id,AjustesDaApuracaoIcms.descricao "

            ds = Banco.ConsultaDataSet(Sql, "Consulta")
            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    With dr
                        linha = "|E111"
                        'linha &= "|" & Estado & "0" & .Item("COD_AJ_APUR") & funcoes.AlinharDireita(.Item("Sequencia_Id"), 4, "0")
                        linha &= "|" & .Item("COD_AJ_APUR")
                        linha &= "|" & LTrim(RTrim(.Item("DESCR_AJ_APUR")))
                        linha &= "|" & .Item("VL_AJ_APUR")
                        linha &= "|"
                    End With
                    strm.WriteLine(linha)
                    RegistroE111 += 1
                    RegistroGeral += 1
                Next
            End If


            '--------------------------------------------------------------------
            ' Registro E116  - Obrigações do Icms A REcolher - Operações Próprias
            '--------------------------------------------------------------------

            Sql = "  SELECT * "
            Sql &= " FROM ProcessoRaIcms"
            Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcesso.SelectedValue & ")"

            ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds1.Tables(0).Rows.Count > 0 Then
                For Each dr1 As DataRow In ds1.Tables(0).Rows
                    With dr1
                        linha = "|E116"
                        linha &= "|000" 'Código da Obrigação a Recolher, conforme a Tabela 5.4
                        linha &= "|" & IcmsARecolher
                        linha &= "|" & CDate(.Item("VencimentoDaObrigacao")).ToSqlDate()
                        linha &= "|" & .Item("CodigoDaReceita")
                        linha &= "|"                'Numero do Processo
                        linha &= "|"                'Indicador da Origem do Processo
                        linha &= "|"                'Descricao Resumida do Processo
                        linha &= "|"                'Descricao complementar das obrigações a recolher
                        linha &= IIf(CDate(txtDataFinal.Text).Year < 2011, "|", "|" & Format(CDate(txtDataFinal.Text), "MMyyyy") & "|")

                        strm.WriteLine(linha)
                        RegistroE116 += 1
                        RegistroGeral += 1
                    End With
                Next
            End If

            '---------------------------------------------
            ' Registro E500  - Periodo da Apuração do IPI
            '---------------------------------------------
            linha = "|E500"
            linha &= "|0"
            linha &= "|" & CDate(txtDataInicial.Text).ToSqlDate()                                                     'Data Inicial
            linha &= "|" & CDate(txtDataFinal.Text).ToSqlDate()                                                     'Data Final
            linha &= "|"

            strm.WriteLine(linha)
            RegistroE500 += 1
            RegistroGeral += 1

            '-------------------------------------------------------
            ' Registro E510  - Apuração do IPI - Operações Próprias
            '-------------------------------------------------------

            Sql = " Select  "
            Sql &= " NotasFiscaisXEncargos.CFOP_Id as CFOP, "
            'Sql &= " IsNull(case "
            'Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI' then "
            'Sql &= " NotasFiscaisXEncargos.Situacaotributaria "
            'Sql &= " end,99) AS CST_IPI, "
            Sql &= "  IsNull(case  "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI' then  NotasFiscaisXEncargos.Situacaotributaria  end, "
            Sql &= " case when NotasFiscais.EntradaSaida_Id ='E' then 49 else 99 end) AS CST_IPI, "
            Sql &= " Sum(IsNull(case "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO' then "
            Sql &= " NotasFiscaisXEncargos.valor "
            Sql &= " end,0)) as VL_CONT_IPI, "
            Sql &= " Sum(IsNull(case "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI'  then "
            Sql &= " NotasFiscaisXEncargos.Base "
            Sql &= " end,0)) as VL_BC_IPI, "
            Sql &= " Sum(IsNull(case "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI'  then "
            Sql &= " NotasFiscaisXEncargos.valor "
            Sql &= " end,0)) as VL_IPI "

            Sql &= " From NotasFiscaisXEncargos  "
            Sql &= " INNER JOIN NotasFiscais ON  NotasFiscaisXEncargos.Empresa_Id = NotasFiscais.Empresa_Id "
            Sql &= "  AND NotasFiscaisXEncargos.EndEmpresa_Id = NotasFiscais.EndEmpresa_Id "
            Sql &= "  AND NotasFiscaisXEncargos.Cliente_Id = NotasFiscais.Cliente_Id "
            Sql &= "  AND NotasFiscaisXEncargos.EndCliente_Id = NotasFiscais.EndCliente_Id "
            Sql &= "  AND NotasFiscaisXEncargos.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  "
            Sql &= " AND NotasFiscaisXEncargos.Serie_Id = NotasFiscais.Serie_Id  "
            Sql &= " AND NotasFiscaisXEncargos.Nota_Id = NotasFiscais.Nota_Id   "
            Sql &= " WHERE NotasFiscaisXEncargos.Empresa_Id = '" & Empresa(0) & "'  "
            Sql &= " AND (NotasFiscaisXEncargos.Encargo_Id = 'IPI'  Or NotasFiscaisXEncargos.Encargo_Id = 'PRODUTO') and isnull(NotasFiscais.situacao,1)= 1 "
            Sql &= " AND  NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "' "
            Sql &= " Group by  NotasFiscaisXEncargos.CFOP_Id, "
            Sql &= "  IsNull(case  "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI' then  NotasFiscaisXEncargos.Situacaotributaria  end, "
            Sql &= " case when NotasFiscais.EntradaSaida_Id ='E' then 49 else 99 end) "
            Sql &= " ORDER BY   NotasFiscaisXEncargos.CFOP_Id, "
            Sql &= "  IsNull(case  "
            Sql &= " when NotasFiscaisXEncargos.Encargo_Id = 'IPI' then  NotasFiscaisXEncargos.Situacaotributaria  end, "
            Sql &= " case when NotasFiscais.EntradaSaida_Id ='E' then 49 else 99 end) "

            ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds1.Tables(0).Rows.Count > 0 Then
                For Each dr1 As DataRow In ds1.Tables(0).Rows
                    With dr1
                        linha = "|E510"
                        linha &= "|" & .Item("CFOP")
                        linha &= "|" & Funcoes.AlinharDireita(.Item("CST_IPI"), 2, "0")
                        linha &= "|" & .Item("VL_CONT_IPI")
                        linha &= "|" & .Item("VL_BC_IPI")
                        linha &= "|" & .Item("VL_IPI")
                        linha &= "|"

                    End With
                    strm.WriteLine(linha)
                    RegistroE510 += 1
                    RegistroGeral += 1
                Next
            End If


            '-------------------------------------------------------
            ' Registro E520  - Apuração do IPI - Operações Próprias
            '-------------------------------------------------------

            Sql = "  SELECT 'E520' AS Reg,     "
            Sql &= " SUM(CASE WHEN Codigo_Id = 9 THEN Valor ELSE 0 END) AS VL_SD_ANT_IPI,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 1 THEN Valor ELSE 0 END) AS VL_DEB_IPI,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 5 THEN Valor ELSE 0 END) AS VL_CRED_IPI,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 2 THEN Valor ELSE 0 END) AS VL_OD_IPI,"
            Sql &= " SUM(CASE WHEN Codigo_Id = 6 THEN Valor ELSE 0 END) AS VL_OC_IPI,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 14 THEN Valor ELSE 0 END) AS VL_SC_IPI,    "
            Sql &= " SUM(CASE WHEN Codigo_Id = 13 THEN Valor ELSE 0 END) AS VL_SD_IPI    "
            Sql &= " FROM ResumoRaIPI"
            Sql &= " WHERE Empresa_Id = '" & Empresa(0) & "' AND (Processo_Id = " & DdlProcessoIPI.SelectedValue & ")"
            Sql &= " GROUP BY Processo_Id"

            ds1 = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds1.Tables(0).Rows.Count > 0 Then
                For Each dr1 As DataRow In ds1.Tables(0).Rows
                    With dr1
                        linha = "|E520"
                        linha &= "|" & .Item("VL_SD_ANT_IPI")
                        linha &= "|" & .Item("VL_DEB_IPI")
                        linha &= "|" & .Item("VL_CRED_IPI")
                        linha &= "|" & .Item("VL_OD_IPI")
                        linha &= "|" & .Item("VL_OC_IPI")
                        linha &= "|" & .Item("VL_SC_IPI")
                        linha &= "|" & .Item("VL_SD_IPI")
                        linha &= "|"

                    End With
                    strm.WriteLine(linha)
                    RegistroE520 += 1
                    RegistroGeral += 1
                Next
            End If

            '-------------------------------------------
            ' Registro E990  - Encerramento do Bloco E
            '-------------------------------------------
            RegistroE990 += 1

            linha = "|E990"
            linha &= "|" & RegistroE001 + RegistroE100 + RegistroE110 + RegistroE111 + RegistroE116 + RegistroE500 + RegistroE510 + RegistroE520 + RegistroE990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1

            '-------------------------------------------
            ' Registro G001  - Abertura do Bloco G
            '-------------------------------------------

            If PeriodoFinal.Year > 2010 Then
                linha = "|G001"
                linha &= "|1"       'Indicador de Movimento - 0 - Bloco com Dados Informados
                '                                             1 - Bloco sem dados informados
                linha &= "|"

                strm.WriteLine(linha)
                RegistroG001 += 1
                RegistroGeral += 1
            End If

            '-------------------------------------------
            ' Registro G990  - Encerramento do Bloco G
            '-------------------------------------------
            If PeriodoFinal.Year > 2010 Then
                RegistroG990 += 1

                linha = "|G990"
                linha &= "|" & RegistroG001 + RegistroG110 + RegistroG125 + RegistroG126 + RegistroG130 + RegistroG140 + RegistroG990
                linha &= "|"

                strm.WriteLine(linha)
                RegistroGeral += 1
            End If


            '-------------------------------------------
            ' Registro H001  - Abertura do Bloco H
            '-------------------------------------------

            linha = "|H001"
            linha &= "|" & IIf(Format(CDate(txtDataFinal.Text), "MM") = 2, 0, 1)      'Indicador de Movimento - 0 - Bloco com Dados Informados
            '                                                                                                   1 - Bloco sem dados informados
            linha &= "|"

            strm.WriteLine(linha)
            RegistroH001 += 1
            RegistroGeral += 1

            '-------------------------------------------
            ' Registro H005 - Totais do Inventário H
            '-------------------------------------------
            If Format(CDate(txtDataFinal.Text), "MM") = 2 Then
                Sql = " SELECT  isnull(Sum(isnull(ValorDoProduto,0)),0) as Valor"
                Sql &= " FROM   ApuracaoDeCustos "
                Sql &= " WHERE  Empresa_Id = '" & Empresa(0) & "'"
                Sql &= "        And CodigoDeCusto_Id in (109, 920) "
                Sql &= "        And Mes_Id = 12"
                Sql &= "        And Ano_Id = " & Format(CDate(txtDataFinal.Text), "yyyy") - 1

                Dim ds27 As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds27.Tables(0).Rows.Count > 0 Then
                    For Each dr27 As DataRow In ds27.Tables(0).Rows
                        With dr27
                            linha = "|H005"
                            linha &= "|3112" & Format(CDate(txtDataFinal.Text), "yyyy") - 1
                            linha &= "|" & .Item("Valor")
                            linha &= "|"
                        End With

                        strm.WriteLine(linha)
                        RegistroH005 += 1
                        RegistroGeral += 1
                    Next
                End If

                '-------------------------------------------
                ' Registro H010 - Inventário
                '-------------------------------------------
                Sql = " SELECT  'H010' AS Tipo, ApuracaoDeCustos.Mes_Id, ApuracaoDeCustos.Ano_Id, ApuracaoDeCustos.Produto_Id AS Produto, ApuracaoDeCustos.Quantidade, " & vbCrLf & _
                      "         CONVERT(money, ApuracaoDeCustos.ValorDoProduto / ApuracaoDeCustos.Quantidade) AS Unitario, ApuracaoDeCustos.ValorDoProduto, " & vbCrLf & _
                      "         CASE " & vbCrLf & _
                      "           WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 109 THEN 2" & vbCrLf & _
                      "           WHEN ApuracaoDeCustos.CodigoDeCusto_Id = 920 and ApuracaoDeCustos.empresa_id <> Clientes.Cliente_Id then 1 " & vbCrLf & _
                      "           ELSE 0" & vbCrLf & _
                      "         END AS Propriedade," & vbCrLf & _
                      "         Clientes.Cliente_Id, Clientes.Endereco_Id, " & vbCrLf & _
                      "         Clientes.Inscricao, Clientes.Estado, PlanoDeCustos.DebitoMercadoria, PlanoDeCustos.CreditoMercadoria, Produtos.Unidade" & vbCrLf & _
                      " FROM    ApuracaoDeCustos INNER JOIN" & vbCrLf & _
                      "         PlanoDeCustos ON ApuracaoDeCustos.CodigoDeCusto_Id = PlanoDeCustos.Codigo_Id INNER JOIN" & vbCrLf & _
                      "         Produtos ON ApuracaoDeCustos.Produto_Id = Produtos.Produto_Id INNER JOIN" & vbCrLf & _
                      "         Clientes ON ApuracaoDeCustos.Deposito_Id = Clientes.Cliente_Id AND ApuracaoDeCustos.EndDeposito_Id = Clientes.Endereco_Id" & vbCrLf & _
                      " WHERE   ApuracaoDeCustos.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                      "         And ApuracaoDeCustos.CodigoDeCusto_Id In (109, 920) " & vbCrLf & _
                      "         And ApuracaoDeCustos.Quantidade <> 0 " & vbCrLf & _
                      "         And ApuracaoDeCustos.Mes_Id =  12" & vbCrLf & _
                      "         And ApuracaoDeCustos.Ano_Id = " & Format(CDate(txtDataFinal.Text), "yyyy") - 1

                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds.Tables(0).Rows.Count > 0 Then
                    For Each dr As DataRow In ds.Tables(0).Rows
                        With dr
                            linha = "|H010"
                            linha &= "|" & .Item("Produto")
                            linha &= "|" & RTrim(.Item("Unidade"))
                            linha &= "|" & Math.Round(CDec(.Item("Quantidade")), 3).ToString
                            linha &= "|" & .Item("Unitario")
                            linha &= "|" & .Item("ValorDoProduto")
                            linha &= "|" & .Item("Propriedade")

                            If .Item("Propriedade") > 0 Then
                                linha &= "|" & .Item("Cliente_Id") & .Item("Endereco_Id")
                            Else
                                linha &= "|" ' & .Item("Cliente_Id") & .Item("Endereco_Id")  '& .Item("COD_PART")
                            End If

                            linha &= "|" ' & .Item("TXT_COMPL")
                            linha &= "|" & IIf(IsDBNull(.Item("DebitoMercadoria")) OrElse Trim(.Item("DebitoMercadoria")) = "", .Item("CreditoMercadoria"), .Item("DebitoMercadoria"))
                            linha &= "|"

                        End With
                        strm.WriteLine(linha)
                        RegistroH010 += 1
                        RegistroGeral += 1
                    Next
                End If

            End If
            '-------------------------------------------
            ' Registro H990  - Encerramento do Bloco H
            '-------------------------------------------
            RegistroH990 += 1

            linha = "|H990"
            linha &= "|" & RegistroH001 + RegistroH005 + RegistroH010 + RegistroH990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1


            '-------------------------------------------
            ' Registro 1001  - Abertura do Bloco 1
            '-------------------------------------------
            linha = "|1001"
            linha &= "|1"       'Indicador de Movimento - 0 - Bloco com Dados Informados
            '                                             1 - Bloco sem dados informados
            linha &= "|"

            strm.WriteLine(linha)
            Registro1001 += 1
            RegistroGeral += 1

            '----------------------------------------------------------
            ' Registro 1100 - Registro de Informações sobre exportações
            '----------------------------------------------------------

            '-------------------------------------------------
            ' Registro 1105 - Documentos Fiscais de Exportação
            '-------------------------------------------------

            '-------------------------------------------
            ' Registro 1990  - Encerramento do Bloco 1
            '-------------------------------------------
            Registro1990 += 1

            linha = "|1990"
            linha &= "|" & Registro1001 + Registro1100 + Registro1105 + Registro1990
            linha &= "|"

            strm.WriteLine(linha)
            RegistroGeral += 1

            '-------------------------------------------
            ' Registro 9001  - Abertura do Bloco 9
            '-------------------------------------------
            linha = "|9001"
            linha &= "|0"       'Indicador de Movimento - 0 - Bloco com Dados Informados
            '                                             1 - Bloco sem dados informados
            linha &= "|"

            strm.WriteLine(linha)
            Registro9001 += 1
            RegistroGeral += 1

            '-------------------------------------------
            ' Registro 9900  - Registros Do Arquivo
            '-------------------------------------------
            linha = "|9900|0000|" & Registro0000 & "|" & vbCrLf
            linha &= "|9900|0001|" & Registro0001 & "|" & vbCrLf
            'linha &= "|9900|0005|" & Registro0005 & "|" & vbCrLf
            linha &= "|9900|0110|" & Registro0110 & "|" & vbCrLf
            linha &= "|9900|0140|" & Registro0140 & "|" & vbCrLf
            linha &= "|9900|0100|" & Registro0100 & "|" & vbCrLf
            linha &= "|9900|0150|" & Registro0150 & "|" & vbCrLf
            linha &= "|9900|0190|" & Registro0190 & "|" & vbCrLf
            linha &= "|9900|0200|" & Registro0200 & "|" & vbCrLf
            linha &= "|9900|0400|" & Registro0400 & "|" & vbCrLf
            linha &= "|9900|0450|" & Registro0450 & "|" & vbCrLf
            linha &= "|9900|0460|" & Registro0460 & "|" & vbCrLf
            linha &= "|9900|0990|" & Registro0990 & "|" & vbCrLf
            linha &= "|9900|C001|" & RegistroC001 & "|" & vbCrLf
            linha &= "|9900|C100|" & RegistroC100 & "|" & vbCrLf
            linha &= "|9900|C170|" & RegistroC170 & "|" & vbCrLf
            linha &= "|9900|C190|" & RegistroC190 & "|" & vbCrLf
            linha &= "|9900|C500|" & RegistroC500 & "|" & vbCrLf
            linha &= "|9900|C510|" & RegistroC510 & "|" & vbCrLf
            linha &= "|9900|C590|" & RegistroC590 & "|" & vbCrLf
            linha &= "|9900|C990|" & RegistroC990 & "|" & vbCrLf
            linha &= "|9900|D001|" & RegistroD001 & "|" & vbCrLf
            linha &= "|9900|D100|" & RegistroD100 & "|" & vbCrLf
            linha &= "|9900|D190|" & RegistroD190 & "|" & vbCrLf
            linha &= "|9900|D500|" & RegistroD500 & "|" & vbCrLf
            linha &= "|9900|D590|" & RegistroD590 & "|" & vbCrLf
            linha &= "|9900|D990|" & RegistroD990 & "|" & vbCrLf
            linha &= "|9900|E001|" & RegistroE001 & "|" & vbCrLf
            linha &= "|9900|E100|" & RegistroE100 & "|" & vbCrLf
            linha &= "|9900|E110|" & RegistroE110 & "|" & vbCrLf
            linha &= "|9900|E111|" & RegistroE111 & "|" & vbCrLf
            linha &= "|9900|E116|" & RegistroE116 & "|" & vbCrLf
            linha &= "|9900|E500|" & RegistroE500 & "|" & vbCrLf
            linha &= "|9900|E510|" & RegistroE510 & "|" & vbCrLf
            linha &= "|9900|E520|" & RegistroE520 & "|" & vbCrLf
            linha &= "|9900|E990|" & RegistroE990 & "|" & vbCrLf
            Registro9900 += 34
            RegistroGeral += 34

            If PeriodoFinal.Year > 2010 Then
                linha &= "|9900|G001|" & RegistroG001 & "|" & vbCrLf
                linha &= "|9900|G990|" & RegistroG990 & "|" & vbCrLf
                Registro9900 += 2
                RegistroGeral += 2
            End If

            linha &= "|9900|H001|" & RegistroH001 & "|" & vbCrLf
            linha &= "|9900|H005|" & RegistroH005 & "|" & vbCrLf
            linha &= "|9900|H010|" & RegistroH010 & "|" & vbCrLf
            linha &= "|9900|H990|" & RegistroH990 & "|" & vbCrLf
            linha &= "|9900|1001|" & Registro1001 & "|" & vbCrLf
            linha &= "|9900|1100|" & Registro1100 & "|" & vbCrLf
            linha &= "|9900|1105|" & Registro1105 & "|" & vbCrLf
            linha &= "|9900|1990|" & Registro1990 & "|" & vbCrLf
            linha &= "|9900|9001|" & Registro9001 & "|" & vbCrLf
            linha &= "|9900|9900|" & Registro9900 + 12 & "|" & vbCrLf
            linha &= "|9900|9990|1|" & vbCrLf
            linha &= "|9900|9999|1|"
            Registro9900 += 12
            RegistroGeral += 12

            strm.WriteLine(linha)


            '-------------------------------------------
            ' Registro 9990  - Encerramento do Bloco 9
            '-------------------------------------------
            Registro9990 += 1
            RegistroGeral += 1

            linha = "|9990"
            linha &= "|" & Registro9001 + Registro9900 + Registro9990 + 1
            linha &= "|"

            strm.WriteLine(linha)

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

            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.location = '" & NomeArquivo2 & "';", True)

        Catch ex As Exception
            strm.Close()
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

        'Me.Cursor = Cursors.Default

        '----------------------------------------------------------------------------------------
        'Shell("C:\windows\notepad.exe '" & txtArquivoDeSaida.Text & "'")

        '----------------------------------------------------------------------------------------

        'MsgBox("Processo Concluido com Sucesso", MsgBoxStyle.Information, "Sped Contabil", eTitulo.Sucess)
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
        DdlUnidade.SelectedIndex = 0
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PisCofins")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class