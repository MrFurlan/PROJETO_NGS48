Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.Net
Imports System.Net.Mail
Imports System.Web
Imports System.IO
Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class GeracaoBradesco
    Inherits BasePage

    Dim Sql As String
    Dim Codigo As String
    Dim Descricao As String
    Dim Inconsist As String
    Dim sql2 As String
    Dim dataproc As Date
    Dim datatu As Date = Today
    Dim SqlArray As New ArrayList
    Dim Mensagem As String = "Msg base"
    Dim j As Integer

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'TabContainer1.Tabs(1).Enabled = False
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("GeracaoBradesco", "ACESSAR") Then
                CargaUnidadeDeNegocio()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If

    End Sub

    '' Carga de unidade de negocio deve ser o pribmeiro a ser chamado . 
    Private Sub CargaUnidadeDeNegocio()
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  "
        Sql &= " FROM Clientes INNER JOIN"
        Sql &= " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id"
        Sql &= " WHERE ClientesXTipos.Tipo_Id = 050"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            Codigo = Dr("Codigo")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 20, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " " & Cidade & " " & Dr("Estado") & " " & Dr("Codigo")
            ddlUnidadeGeracaoBradesco.Items.Add(New ListItem(Descricao, Codigo))
        Next

        ddlUnidadeGeracaoBradesco.Items.Insert(0, "")
        ddlUnidadeGeracaoBradesco.SelectedIndex = 0

        ddl.Carregar(ddlUnidadeGeracaoBradesco, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeGeracaoBradesco, DdlEmpresaConsultaTitulos, True)

    End Sub

    '' Carga da empresa so podera ser chamado apos carga da unidade de negocio. 
    Private Sub CargaEmpresaGeracaoBradesco()
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaConsultaTitulos.Items.Clear()

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
        Sql &= " FROM   GruposXEmpresas INNER JOIN"
        Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
        Sql &= " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidadeGeracaoBradesco.SelectedValue & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaConsultaTitulos.Items.Insert(0, "")
        DdlEmpresaConsultaTitulos.SelectedIndex = 0

    End Sub

    '' Se for muado o indice da tabela de unidade devera automaticamente carregar a tabela de empresa
    Protected Sub ddlUnidadeGeracaoBradesco_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresaGeracaoBradesco()
    End Sub

    Private Sub Limpar()
        DdlContaConsulta.SelectedIndex = 0

        txtPeriodoInicialConsultaTitulos.Text = ""
        txtPeriodoInicialConsultaTitulos.Enabled = True

        txtPeriodoFinalConsultaTitulos.Text = ""
        txtPeriodoFinalConsultaTitulos.Enabled = True
        RbGeral.Checked = True
        RbEnviado.Checked = False
        RbBaixado.Checked = False
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeGeracaoBradesco.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Protected Sub DdlEmpresaConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        ''DdlContaPagadora.Items.Clear()

        If Campo(0) <> "" Then
            Sql = "SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes "
            Sql &= " FROM BancosXContas"
            Sql &= " Where Empresa_Id = '" & Campo(0) & "'"                 'Empresa
            Sql &= " and EndEmpresa_Id  = " & Campo(1)                      'Endereco da Empresa
            Sql &= " and Banco_Id  = 0237"      'Endereco da Empresa

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
                DdlContaConsulta.Items.Add(New ListItem(Descricao, Conta))
            Next
            DdlContaConsulta.Items.Insert(0, "")
            DdlContaConsulta.SelectedIndex = 0
        End If
    End Sub

    Private Sub GERFOR()
        '' ARQUIVO QUE SERA GERADO. 
        '' DEFINICAO DE VARIAVEIS
        Dim NomeArquivo As String
        NomeArquivo = "C:\RESUMO\FOR.TXT"
        Dim arquivo As String = NomeArquivo
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim strm As StreamWriter
        Dim strLinha As String
        Dim sql As String
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        'Dim linha As String
        Dim dataproc As Date
        dataproc = Date.Today
        Dim ds As New DataSet
        '' DEFINICAO DE VARIAVEIS - FIM 
        ''  SQL DE PESQUISA 
        sql = " DECLARE @sql varchar(8000), @empresa varchar(18), @endempresa int "
        sql &= " SET @empresa = ''"
        sql &= " SET @endempresa = -1"
        sql &= " CREATE TABLE #Destinatario"
        sql &= "("
        sql &= "	Destinatario_DES varchar(18), "
        sql &= "	Nome			varchar(60),"
        sql &= "	Inscricao		varchar(20),"
        sql &= "	Bairro			varchar(50),"
        sql &= "	Cidade			varchar(50),"
        sql &= "	Estado			varchar(2), "
        sql &= "    Cep             varchar(9)"
        sql &= ")"
        sql &= "SET @sql = 'INSERT INTO #Destinatario ' + "
        sql &= "        'SELECT CP.Destinatario, D.Nome, D.Inscricao, D.Bairro, D.Cidade, D.Estado, D.Cep ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'INNER JOIN Clientes D ' + "
        sql &= "         'ON D.Cliente_Id = CP.Destinatario ' + "
        sql &= "         'AND D.Endereco_Id = CP.EndDestinatario ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado.  
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " 	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Contas "
        sql &= "("
        sql &= " 	Destinatario_Cont	varchar(18), "
        sql &= "	Banco_CONT			int, "
        sql &= "	Agencia_CONT		varchar(5), "
        sql &= " 	DvAgencia_CONT		varchar(2), "
        sql &= " 	Conta_CONT			varchar(20), "
        sql &= " 	DvConta_CONT		varchar(2), "
        sql &= "    TipoIns_CONT        varchar(9) "
        sql &= ") "
        sql &= "SET @sql = 'INSERT INTO #Contas ' + "
        sql &= "'SELECT CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, ' + "
        sql &= " 'CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END ' + "
        sql &= " 'FROM ContasAPagar CP ' + "
        sql &= " 'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= " 'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Pagamentos "
        sql &= "("
        sql &= "	TipoIns_pag			varchar(9),"
        sql &= "	Destinatario_pag	varchar(18), "
        sql &= "	Banco_pag			int,"
        sql &= "	Agencia_pag			varchar(5),"
        sql &= "	DvAgencia_pag		varchar(2),"
        sql &= "	Conta_pag			varchar(20),"
        sql &= "	DvConta_pag			varchar(2),"
        sql &= "	Vencimento_pag		datetime,"
        sql &= " 	Movimento_pag		datetime,"
        sql &= " 	ValorLiquido_pag	numeric(18, 9),"
        sql &= "	Finalidade_pag		bit,"
        sql &= "	Doc_pag				char(1),"
        sql &= "	CodigoDigitado_pag	char(1), "
        sql &= "	CodigoDeBarras_pag	varchar(50),"
        sql &= "	Modalidade_pag		varchar(2),"
        sql &= "	TipoDoc_pag			varchar(2),"
        sql &= "	TipoConta_pag		bit,"
        sql &= "	Duplicata_pag		int,"
        sql &= "    TipoPagto_pag       Int"
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #Pagamentos ' +"
        sql &= "        'SELECT CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END, CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, ' + "
        sql &= "        'CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, CP.Vencimento, CP.Movimento, CP.ValorLiquido, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) > 0 THEN 1 ELSE 3 END, ' + "
        sql &= "        'CASE WHEN CP.BancoCliente = 237 THEN ''C'' ELSE ''D'' END, CP.CodigoDigitado, CP.CodigoDeBarras, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) = 0 THEN ''01'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente = 237 THEN ''05'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente <> 237 THEN ''08'' ' + "
        sql &= "        'ELSE ''31'' END, ''05'', 1, CP.Registro_Id, CP.TipoPagto ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' +"
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado.  
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= "	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' '"
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #GERAL "
        sql &= " ("
        sql &= "	Destinatario_DES_GER      varchar(18), "
        sql &= "	Nome_DES_GER              varchar(60), "
        sql &= "	Inscricao_DES_GER         varchar(20), "
        sql &= "	Bairro_DES_GER            varchar(50), "
        sql &= "	Cidade_DES_GER            varchar(50), "
        sql &= "	Estado_DES_GER            varchar(2),  "
        sql &= " 	Cep_DES_GER               varchar(9),  "
        sql &= "	Destinatario_CONT_GER     varchar(18), "
        sql &= "	Banco_CONT_GER            int,"
        sql &= "	Agencia_CONT_GER          varchar(5),  "
        sql &= "	DvAgencia_CONT_GER        varchar(2),  "
        sql &= "	Conta_CONT_GER            varchar(20), "
        sql &= "	DvConta_CONT_GER          varchar(2),  "
        sql &= "	TipoIns_CONT_GER          varchar(9),  "
        sql &= " 	TipoIns_pag_GER           varchar(9),  "
        sql &= "	Destinatario_pag_GER      varchar(18), "
        sql &= "	Banco_pag_GER             int, "
        sql &= "	Agencia_pag_GER           varchar(5),  "
        sql &= "	DvAgencia_pag_GER         varchar(2),  "
        sql &= "	Conta_pag_GER             varchar(20), "
        sql &= " 	DvConta_pag_GER		  varchar(2), "
        sql &= "	Vencimento_pag_GER        datetime, "
        sql &= " 	Movimento_pag_GER         datetime,"
        sql &= "	ValorLiquido_pag_GER    numeric(18, 2),"
        sql &= " 	Finalidade_pag_GER        bit,"
        sql &= " 	Doc_pag_GER				char(1),"
        sql &= "	CodigoDigitado_pag_GER	char(1),"
        sql &= " 	CodigoDeBarras_pag_GER	varchar(50), "
        sql &= "	Modalidade_pag_GER	varchar(2), "
        sql &= "	TipoDoc_pag_GER		varchar(2), "
        sql &= "	TipoConta_pag_GER	bit, "
        sql &= "	Duplicata_pag_GER	int, "
        sql &= "    TipoPagto_pag_GER   Int "
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #GERAL ' + "
        sql &= "'sELECT De.Destinatario_des, DE.nome, De.Inscricao, De.bairro, De.cidade, De.Estado, De.Cep, ' + "
        sql &= " 'ct.Destinatario_cont, Ct.Banco_cont, Ct.Agencia_cont, Ct.Dvagencia_cont, Ct.Conta_cont, Ct.Dvconta_cont, ' + "
        sql &= " 'ct.TipoIns_cont, pg.Tipoins_pag, pg.Destinatario_pag, Pg.Banco_pag, Pg.Agencia_pag, pg.Dvagencia_pag,  '  +"
        sql &= " 'pg.Conta_pag, pg.DvConta_pag, pg.VencimentO_pag, Pg.Movimento_pag, Pg.ValorLiquido_pag, Pg.Finalidade_pag, ' +"
        sql &= " 'Pg.Doc_pag, Pg.CodigoDigitado_pag, Pg.Codigodebarras_pag, pg.Modalidade_pag, Pg.TipoDoc_pag, ' + "
        sql &= " 'Pg.TipoConta_pag, Pg.Duplicata_pag, Pg.TipoPagto_pag '+"
        sql &= " 'FROM #PAGAMENTOS PG ' +"
        sql &= " 'INNER JOIN #DESTINATARIO DE ON DE.DESTINATARIO_DES = PG.DESTINATARIO_PAG ' +"
        sql &= " 'INNER JOIN #CONTAS CT ON CT.DESTINATARIO_CONT = PG.DESTINATARIO_PAG AND CT.BANCO_CONT = PG.BANCO_PAG AND CT.AGENCIA_CONT = PG.AGENCIA_PAG AND CT.CONTA_CONT = PG.CONTA_PAG '"
        sql &= " EXEC (@sql) "
        sql &= " SELECT * FROM #GERAL"
        sql &= " DROP TABLE #Destinatario"
        sql &= " DROP TABLE #Pagamentos"
        sql &= " DROP TABLE #Contas"
        ''  SQL DE PESQUISA - FIM 
        '' Geracao do arquivo texto.
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " ."
                Else
                    strLinha = "0"
                    strLinha &= Funcoes.AlinharDireita(dr("Destinatario_DES_GER"), 14, "0")
                    strLinha &= Funcoes.AlinharEsquerda(dr("Nome_DES_GER"), 30, " ")
                    '' Se tiver 11 posicoes ou menos sera pessoa fisica portanto 1 senao sera 2 portanto empresa
                    If Len(dr("Destinatario_DES_GER")) < 12 Then
                        strLinha &= "1"
                    Else
                        strLinha &= "2"
                    End If
                    strLinha &= "0"
                    strLinha &= Funcoes.AlinharDireita(dr("Destinatario_DES_GER"), 14, "0")
                    strLinha &= "                                                                        "
                    strLinha &= "00000000"
                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(strLinha)
                    strm.Close()
                End If
            Next
        End If
        ''  ARQUIVO 01 - FIM 
    End Sub

    Private Sub gercon()
        '' ARQUIVO QUE SERA GERADO. 
        '' DEFINICAO DE VARIAVEIS
        Dim NomeArquivo As String
        NomeArquivo = "C:\RESUMO\CONTAS.TXT"
        Dim arquivo As String = NomeArquivo
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim strm As StreamWriter
        Dim strLinha As String
        Dim sql As String
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        'Dim linha As String
        Dim dataproc As Date
        dataproc = Date.Today
        Dim ds As New DataSet
        '' DEFINICAO DE VARIAVEIS - FIM 
        ''  SQL DE PESQUISA 
        sql = " DECLARE @sql varchar(8000), @empresa varchar(18), @endempresa int "
        sql &= " SET @empresa = ''"
        sql &= " SET @endempresa = -1"
        sql &= " CREATE TABLE #Destinatario"
        sql &= "("
        sql &= "	Destinatario_DES varchar(18), "
        sql &= "	Nome			varchar(60),"
        sql &= "	Inscricao		varchar(20),"
        sql &= "	Bairro			varchar(50),"
        sql &= "	Cidade			varchar(50),"
        sql &= "	Estado			varchar(2), "
        sql &= "    Cep             varchar(9)"
        sql &= ")"
        sql &= "SET @sql = 'INSERT INTO #Destinatario ' + "
        sql &= "        'SELECT CP.Destinatario, D.Nome, D.Inscricao, D.Bairro, D.Cidade, D.Estado, D.Cep ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'INNER JOIN Clientes D ' + "
        sql &= "         'ON D.Cliente_Id = CP.Destinatario ' + "
        sql &= "         'AND D.Endereco_Id = CP.EndDestinatario ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado.  
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " 	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Contas "
        sql &= "("
        sql &= " 	Destinatario_Cont	varchar(18), "
        sql &= "	Banco_CONT			int, "
        sql &= "	Agencia_CONT		varchar(5), "
        sql &= " 	DvAgencia_CONT		varchar(2), "
        sql &= " 	Conta_CONT			varchar(20), "
        sql &= " 	DvConta_CONT		varchar(2), "
        sql &= "    TipoIns_CONT        varchar(9) "
        sql &= ") "
        sql &= "SET @sql = 'INSERT INTO #Contas ' + "
        sql &= "'SELECT CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, ' + "
        sql &= " 'CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END ' + "
        sql &= " 'FROM ContasAPagar CP ' + "
        sql &= " 'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= " 'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Pagamentos "
        sql &= "("
        sql &= "	TipoIns_pag			varchar(9),"
        sql &= "	Destinatario_pag	varchar(18), "
        sql &= "	Banco_pag			int,"
        sql &= "	Agencia_pag			varchar(5),"
        sql &= "	DvAgencia_pag		varchar(2),"
        sql &= "	Conta_pag			varchar(20),"
        sql &= "	DvConta_pag			varchar(2),"
        sql &= "	Vencimento_pag		datetime,"
        sql &= " 	Movimento_pag		datetime,"
        sql &= " 	ValorLiquido_pag	numeric(18, 9),"
        sql &= "	Finalidade_pag		bit,"
        sql &= "	Doc_pag				char(1),"
        sql &= "	CodigoDigitado_pag	char(1), "
        sql &= "	CodigoDeBarras_pag	varchar(50),"
        sql &= "	Modalidade_pag		varchar(2),"
        sql &= "	TipoDoc_pag			varchar(2),"
        sql &= "	TipoConta_pag		bit,"
        sql &= "	Duplicata_pag		int,"
        sql &= "    TipoPagto_pag       Int"
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #Pagamentos ' +"
        sql &= "        'SELECT CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END, CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, ' + "
        sql &= "        'CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, CP.Vencimento, CP.Movimento, CP.ValorLiquido, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) > 0 THEN 1 ELSE 3 END, ' + "
        sql &= "        'CASE WHEN CP.BancoCliente = 237 THEN ''C'' ELSE ''D'' END, CP.CodigoDigitado, CP.CodigoDeBarras, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) = 0 THEN ''01'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente = 237 THEN ''05'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente <> 237 THEN ''08'' ' + "
        sql &= "        'ELSE ''31'' END, ''05'', 1, CP.Registro_Id, CP.TipoPagto ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' +"
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= "	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' '"
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #GERAL "
        sql &= " ("
        sql &= "	Destinatario_DES_GER      varchar(18), "
        sql &= "	Nome_DES_GER              varchar(60), "
        sql &= "	Inscricao_DES_GER         varchar(20), "
        sql &= "	Bairro_DES_GER            varchar(50), "
        sql &= "	Cidade_DES_GER            varchar(50), "
        sql &= "	Estado_DES_GER            varchar(2),  "
        sql &= " 	Cep_DES_GER               varchar(9),  "
        sql &= "	Destinatario_CONT_GER     varchar(18), "
        sql &= "	Banco_CONT_GER            int,"
        sql &= "	Agencia_CONT_GER          varchar(5),  "
        sql &= "	DvAgencia_CONT_GER        varchar(2),  "
        sql &= "	Conta_CONT_GER            varchar(20), "
        sql &= "	DvConta_CONT_GER          varchar(2),  "
        sql &= "	TipoIns_CONT_GER          varchar(9),  "
        sql &= " 	TipoIns_pag_GER           varchar(9),  "
        sql &= "	Destinatario_pag_GER      varchar(18), "
        sql &= "	Banco_pag_GER             int, "
        sql &= "	Agencia_pag_GER           varchar(5),  "
        sql &= "	DvAgencia_pag_GER         varchar(2),  "
        sql &= "	Conta_pag_GER             varchar(20), "
        sql &= " 	DvConta_pag_GER		  varchar(2), "
        sql &= "	Vencimento_pag_GER        datetime, "
        sql &= " 	Movimento_pag_GER         datetime,"
        sql &= "	ValorLiquido_pag_GER    numeric(18, 2),"
        sql &= " 	Finalidade_pag_GER        bit,"
        sql &= " 	Doc_pag_GER				char(1),"
        sql &= "	CodigoDigitado_pag_GER	char(1),"
        sql &= " 	CodigoDeBarras_pag_GER	varchar(50), "
        sql &= "	Modalidade_pag_GER	varchar(2), "
        sql &= "	TipoDoc_pag_GER		varchar(2), "
        sql &= "	TipoConta_pag_GER	bit, "
        sql &= "	Duplicata_pag_GER	int, "
        sql &= "    TipoPagto_pag_GER   Int "
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #GERAL ' + "
        sql &= "'sELECT De.Destinatario_des, DE.nome, De.Inscricao, De.bairro, De.cidade, De.Estado, De.Cep, ' + "
        sql &= " 'ct.Destinatario_cont, Ct.Banco_cont, Ct.Agencia_cont, Ct.Dvagencia_cont, Ct.Conta_cont, Ct.Dvconta_cont, ' + "
        sql &= " 'ct.TipoIns_cont, pg.Tipoins_pag, pg.Destinatario_pag, Pg.Banco_pag, Pg.Agencia_pag, pg.Dvagencia_pag,  '  +"
        sql &= " 'pg.Conta_pag, pg.DvConta_pag, pg.VencimentO_pag, Pg.Movimento_pag, Pg.ValorLiquido_pag, Pg.Finalidade_pag, ' +"
        sql &= " 'Pg.Doc_pag, Pg.CodigoDigitado_pag, Pg.Codigodebarras_pag, pg.Modalidade_pag, Pg.TipoDoc_pag, ' + "
        sql &= " 'Pg.TipoConta_pag, Pg.Duplicata_pag, Pg.TipoPagto_pag '+"
        sql &= " 'FROM #PAGAMENTOS PG ' +"
        sql &= " 'INNER JOIN #DESTINATARIO DE ON DE.DESTINATARIO_DES = PG.DESTINATARIO_PAG ' +"
        sql &= " 'INNER JOIN #CONTAS CT ON CT.DESTINATARIO_CONT = PG.DESTINATARIO_PAG AND CT.BANCO_CONT = PG.BANCO_PAG AND CT.AGENCIA_CONT = PG.AGENCIA_PAG AND CT.CONTA_CONT = PG.CONTA_PAG '"
        sql &= " EXEC (@sql) "
        sql &= " SELECT * FROM #GERAL"
        sql &= " DROP TABLE #Destinatario"
        sql &= " DROP TABLE #Pagamentos"
        sql &= " DROP TABLE #Contas"
        ''  SQL DE PESQUISA - FIM 
        '' Geracao do arquivo texto.
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " ."
                Else
                    If Len(dr("Destinatario_DES_GER")) < 12 Then
                        strLinha = "1"
                    Else
                        strLinha = "2"
                    End If
                    strLinha &= Funcoes.AlinharDireita(dr("Destinatario_DES_GER"), 15, "0")
                    strLinha &= Funcoes.AlinharDireita(dr("Banco_CONT_GER"), 3, "0")
                    strLinha &= Funcoes.AlinharDireita(dr("Agencia_pag_GER"), 5, "0")
                    strLinha &= Funcoes.AlinharEsquerda(dr("DvAgencia_pag_GER"), 1, " ")
                    strLinha &= Funcoes.AlinharDireita(dr("Conta_pag_GER"), 13, "0")
                    strLinha &= Funcoes.AlinharEsquerda(dr("DvConta_pag_GER"), 2, " ")
                    strLinha &= "XXXXXXXXXXXXXXX"
                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(strLinha)
                    strm.Close()
                End If
            Next
        End If
        ''  ARQUIVO 02 - FIM 
    End Sub

    Private Sub gerpag()
        '' ARQUIVO QUE SERA GERADO. 
        '' DEFINICAO DE VARIAVEIS - INICIO 
        Dim NomeArquivo As String
        NomeArquivo = "C:\RESUMO\pag.TXT"
        Dim arquivo As String = NomeArquivo
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim strm As StreamWriter
        Dim strLinha As String
        Dim sql As String
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        'Dim linha As String
        Dim dataproc As Date
        dataproc = Date.Today
        Dim ds As New DataSet
        '' DEFINICAO DE VARIAVEIS - FIM 
        ''  SQL DE PESQUISA 
        sql = " DECLARE @sql varchar(8000), @empresa varchar(18), @endempresa int "
        sql &= " SET @empresa = ''"
        sql &= " SET @endempresa = -1"
        sql &= " CREATE TABLE #Destinatario"
        sql &= "("
        sql &= "	Destinatario_DES varchar(18), "
        sql &= "	Nome			varchar(60),"
        sql &= "	Inscricao		varchar(20),"
        sql &= "	Bairro			varchar(50),"
        sql &= "	Cidade			varchar(50),"
        sql &= "	Estado			varchar(2), "
        sql &= "    Cep             varchar(9)"
        sql &= ")"
        sql &= "SET @sql = 'INSERT INTO #Destinatario ' + "
        sql &= "        'SELECT CP.Destinatario, D.Nome, D.Inscricao, D.Bairro, D.Cidade, D.Estado, D.Cep ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'INNER JOIN Clientes D ' + "
        sql &= "         'ON D.Cliente_Id = CP.Destinatario ' + "
        sql &= "         'AND D.Endereco_Id = CP.EndDestinatario ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " 	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Contas "
        sql &= "("
        sql &= " 	Destinatario_Cont	varchar(18), "
        sql &= "	Banco_CONT			int, "
        sql &= "	Agencia_CONT		varchar(5), "
        sql &= " 	DvAgencia_CONT		varchar(2), "
        sql &= " 	Conta_CONT			varchar(20), "
        sql &= " 	DvConta_CONT		varchar(2), "
        sql &= "    TipoIns_CONT        varchar(9) "
        sql &= ") "
        sql &= "SET @sql = 'INSERT INTO #Contas ' + "
        sql &= "'SELECT CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, ' + "
        sql &= " 'CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END ' + "
        sql &= " 'FROM ContasAPagar CP ' + "
        sql &= " 'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= " 'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Pagamentos "
        sql &= "("
        sql &= "	TipoIns_pag			varchar(9),"
        sql &= "	Destinatario_pag	varchar(18), "
        sql &= "	Banco_pag			int,"
        sql &= "	Agencia_pag			varchar(5),"
        sql &= "	DvAgencia_pag		varchar(2),"
        sql &= "	Conta_pag			varchar(20),"
        sql &= "	DvConta_pag			varchar(2),"
        sql &= "	Vencimento_pag		datetime,"
        sql &= " 	Movimento_pag		datetime,"
        sql &= " 	ValorLiquido_pag	numeric(18, 9),"
        sql &= "	Finalidade_pag		bit,"
        sql &= "	Doc_pag				char(1),"
        sql &= "	CodigoDigitado_pag	char(1), "
        sql &= "	CodigoDeBarras_pag	varchar(50),"
        sql &= "	Modalidade_pag		varchar(2),"
        sql &= "	TipoDoc_pag			varchar(2),"
        sql &= "	TipoConta_pag		bit,"
        sql &= "	Duplicata_pag		int,"
        sql &= "    TipoPagto_pag       Int"
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #Pagamentos ' +"
        sql &= "        'SELECT CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END, CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, ' + "
        sql &= "        'CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, CP.Vencimento, CP.Movimento, CP.ValorLiquido, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) > 0 THEN 1 ELSE 3 END, ' + "
        sql &= "        'CASE WHEN CP.BancoCliente = 237 THEN ''C'' ELSE ''D'' END, CP.CodigoDigitado, CP.CodigoDeBarras, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) = 0 THEN ''01'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente = 237 THEN ''05'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente <> 237 THEN ''08'' ' + "
        sql &= "        'ELSE ''31'' END, ''05'', 1, CP.Registro_Id, CP.TipoPagto ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' +"
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= "	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' '"
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #GERAL "
        sql &= " ("
        sql &= "	Destinatario_DES_GER      varchar(18), "
        sql &= "	Nome_DES_GER              varchar(60), "
        sql &= "	Inscricao_DES_GER         varchar(20), "
        sql &= "	Bairro_DES_GER            varchar(50), "
        sql &= "	Cidade_DES_GER            varchar(50), "
        sql &= "	Estado_DES_GER            varchar(2),  "
        sql &= " 	Cep_DES_GER               varchar(9),  "
        sql &= "	Destinatario_CONT_GER     varchar(18), "
        sql &= "	Banco_CONT_GER            int,"
        sql &= "	Agencia_CONT_GER          varchar(5),  "
        sql &= "	DvAgencia_CONT_GER        varchar(2),  "
        sql &= "	Conta_CONT_GER            varchar(20), "
        sql &= "	DvConta_CONT_GER          varchar(2),  "
        sql &= "	TipoIns_CONT_GER          varchar(9),  "
        sql &= " 	TipoIns_pag_GER           varchar(9),  "
        sql &= "	Destinatario_pag_GER      varchar(18), "
        sql &= "	Banco_pag_GER             int, "
        sql &= "	Agencia_pag_GER           varchar(5),  "
        sql &= "	DvAgencia_pag_GER         varchar(2),  "
        sql &= "	Conta_pag_GER             varchar(20), "
        sql &= " 	DvConta_pag_GER		  varchar(2), "
        sql &= "	Vencimento_pag_GER        datetime, "
        sql &= " 	Movimento_pag_GER         datetime,"
        sql &= "	ValorLiquido_pag_GER    numeric(18, 2),"
        sql &= " 	Finalidade_pag_GER        bit,"
        sql &= " 	Doc_pag_GER				char(1),"
        sql &= "	CodigoDigitado_pag_GER	char(1),"
        sql &= " 	CodigoDeBarras_pag_GER	varchar(50), "
        sql &= "	Modalidade_pag_GER	varchar(2), "
        sql &= "	TipoDoc_pag_GER		varchar(2), "
        sql &= "	TipoConta_pag_GER	bit, "
        sql &= "	Duplicata_pag_GER	int, "
        sql &= "    TipoPagto_pag_GER   Int "
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #GERAL ' + "
        sql &= "'sELECT De.Destinatario_des, DE.nome, De.Inscricao, De.bairro, De.cidade, De.Estado, De.Cep, ' + "
        sql &= " 'ct.Destinatario_cont, Ct.Banco_cont, Ct.Agencia_cont, Ct.Dvagencia_cont, Ct.Conta_cont, Ct.Dvconta_cont, ' + "
        sql &= " 'ct.TipoIns_cont, pg.Tipoins_pag, pg.Destinatario_pag, Pg.Banco_pag, Pg.Agencia_pag, pg.Dvagencia_pag,  '  +"
        sql &= " 'pg.Conta_pag, pg.DvConta_pag, pg.VencimentO_pag, Pg.Movimento_pag, Pg.ValorLiquido_pag, Pg.Finalidade_pag, ' +"
        sql &= " 'Pg.Doc_pag, Pg.CodigoDigitado_pag, Pg.Codigodebarras_pag, pg.Modalidade_pag, Pg.TipoDoc_pag, ' + "
        sql &= " 'Pg.TipoConta_pag, Pg.Duplicata_pag, Pg.TipoPagto_pag '+"
        sql &= " 'FROM #PAGAMENTOS PG ' +"
        sql &= " 'INNER JOIN #DESTINATARIO DE ON DE.DESTINATARIO_DES = PG.DESTINATARIO_PAG ' +"
        sql &= " 'INNER JOIN #CONTAS CT ON CT.DESTINATARIO_CONT = PG.DESTINATARIO_PAG AND CT.BANCO_CONT = PG.BANCO_PAG AND CT.AGENCIA_CONT = PG.AGENCIA_PAG AND CT.CONTA_CONT = PG.CONTA_PAG '"
        sql &= " EXEC (@sql) "
        sql &= " SELECT * FROM #GERAL"
        sql &= " DROP TABLE #Destinatario"
        sql &= " DROP TABLE #Pagamentos"
        sql &= " DROP TABLE #Contas"
        ''  SQL DE PESQUISA - FIM 
        '' Geracao do arquivo texto.
        ''  ARQUIVO 03 - inicio. 
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " ."
                Else
                    'ex01 - fixo 35 espacos 
                    strLinha = "                                   "
                    'ex02 - Tipo de insricao 1 - fisica 2 - juridica
                    If Len(dr("Destinatario_DES_GER")) < 12 Then
                        strLinha &= "1"
                    Else
                        strLinha &= "2"
                    End If
                    'ex03 - Inscricao cnpj ou cpf
                    strLinha &= Funcoes.AlinharDireita(dr("Destinatario_DES_GER"), 15, "0")
                    'todos os campos de 05 a 08 quando boleto devem ser colocado como brancos somente banco que recebe bol.
                    'ex04 - Banco 
                    'ex05 - Agencia 
                    'ex06 - Digito da agencia 
                    'ex07 - Conta 
                    'ex08 - Digito da conta. 
                    If dr("Conta_pag_GER") = " " Then
                        strLinha &= "BOL"
                        strLinha &= "     "
                        strLinha &= " "
                        strLinha &= "             "
                        strLinha &= "  "
                    Else
                        strLinha &= Funcoes.AlinharDireita(dr("Banco_CONT_GER"), 3, "0")
                        strLinha &= Funcoes.AlinharDireita(dr("Agencia_CONT_GER"), 5, "0")
                        strLinha &= Funcoes.AlinharDireita(dr("DvAgencia_CONT_GER"), 1, " ")
                        strLinha &= Funcoes.AlinharDireita(dr("Conta_CONT_GER"), 13, "0")
                        strLinha &= Funcoes.AlinharDireita(dr("DvConta_CONT_GER"), 2, " ")
                    End If
                    'ex09 - Data de vencimento do titulo 
                    strLinha &= CDate(dr("Vencimento_pag_GER")).ToSqlDate()
                    'ex10 - Data de pagamento do titulo
                    strLinha &= CDate(dr("Movimento_pag_GER")).ToSqlDate()
                    'ex11 - Valor do pagamento
                    strLinha &= Funcoes.AlinharDireita(CStr(dr("ValorLiquido_pag_GER")).Replace(",", ""), 15, "0")
                    'ex12 - data do desconto - fixo 8 zeros 
                    strLinha &= "00000000"
                    'ex13 - valor do desconto - fixo 15 zeros 
                    strLinha &= "000000000000000"
                    'ex14 - valor do acrescimo - fixo 15 zeros 
                    strLinha &= "000000000000000"
                    'ex15 - Finaliade se tem conta corrente é 01 se nao é 02 
                    If dr("Conta_pag_GER") = " " Then
                        strLinha &= "02"
                    Else
                        strLinha &= "01"
                    End If
                    'ex16 - Doc Fixo C Se movimentacao dentro da Insol devera ser "D"
                    strLinha &= "C"
                    'ex17 - Linha digitavel se conta corrente diferente de brancos 47 espacos
                    'ex18 - Codigo de barras se conta corrente diferente de brancos 44 espacos
                    If dr("Conta_pag_GER") = " " Then
                        If dr("CodigoDigitado_pag_GER") = "S" Then
                            strLinha &= Funcoes.AlinharDireita(dr("CodigoDeBarras_pag_GER"), 47, " ")
                            strLinha &= "                                            "
                        Else
                            strLinha &= "                                               "
                            strLinha &= Funcoes.AlinharDireita(dr("CodigoDeBarras_pag_GER"), 44, " ")
                        End If
                    Else
                        strLinha &= "                                               "
                        strLinha &= "                                            "
                    End If
                    'ex19 - Modalidade se conta corrente diferente de zeros = 01 , se ted = 05 , se banco 
                    'diferente de bradesco = 03 , se banco diferente de bradesco e ted 08 , se conta corrente = zeros 
                    'e banco bradesco modalidade = 30 , se conta corrente = zeros e banco diferente de bradesco = 31
                    strLinha &= Funcoes.AlinharDireita(dr("Modalidade_pag_GER"), 2, "0")
                    'Ex20 - Tipo do documento - 05
                    strLinha &= "05"
                    'Ex21 - Tipo de conta corrente  se conta corrente diferente de zeros fixo 01 se nao espacos
                    If dr("Conta_pag_GER") = " " Then
                        strLinha &= "  "
                    Else
                        strLinha &= "01"
                    End If
                    'Ex22 - Numero do titulo se deposito 10 espacos se nao numero do titulo a esquerda
                    strLinha &= Funcoes.AlinharEsquerda(dr("Duplicata_pag_GER"), 10, " ")
                    'Ex23 - INSTRUCAO FIXO 40 Y
                    strLinha &= "YYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYYY"
                    'Ex24 - CONTROLE BANCO SE BRADESCO 01 SE NAO XX
                    If dr("Banco_CONT_GER") = "237" Then
                        strLinha &= "01"
                    Else
                        strLinha &= "XX"
                    End If
                    'Ex25 - Se ted 1600 se nao ZEROS CONTROLE BANCO SE BRADESCO 01 SE NAO 0000
                    strLinha &= "0000"
                    'Ex26 - FIXO 0001
                    strLinha &= "0001"
                    strm = New StreamWriter(arquivo, True)
                    strm.WriteLine(strLinha)
                    strm.Close()
                End If
            Next
        End If
        ''  ARQUIVO 03 - FIM 
        '' Mudanca de status do registro de 0 para 1 - enviado ao banco - inicio . 
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        sql2 = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " ."
                Else
                    Inconsist = " ."
                    sql2 = " UPDATE ContasAPagar"
                    sql2 &= " SET SituacaoBancaria = '" & "2" & "'"
                    sql2 &= ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"                     'Usuario que Alterou
                    sql2 &= ", UsuarioAlteracaoData = '" & CDate(datatu).ToSqlDate() & "'"   'Data da Alteracao
                    sql2 &= " WHERE Registro_ID = " & CInt(dr("Duplicata_pag_GER"))
                    SqlArray.Add(sql2)
                    j = j + 1
                End If
            Next
            If j > 0 Then
                If Banco.GravaBanco(SqlArray) Then
                    MsgBox(Me.Page, "Registros Atualizados com sucesso.Arquivos Gerados.", eTitulo.Sucess)
                End If
            End If
        End If
        If j > 0 Then
            MsgBox(Me.Page, "Registros atualizados com Sucesso.", eTitulo.Sucess)
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            MsgBox(Me.Page, "Não Existiam registros a Atualizar.", eTitulo.Sucess)
            txtPeriodoFinalConsultaTitulos.Focus()
        End If
    End Sub

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        RbGeral.Checked = True
        RbEnviado.Checked = False
        RbBaixado.Checked = False

    End Sub

    Protected Sub RbEnviado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        RbGeral.Checked = False
        RbEnviado.Checked = True
        RbBaixado.Checked = False

    End Sub

    Protected Sub RbBaixado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        RbGeral.Checked = False
        RbEnviado.Checked = False
        RbBaixado.Checked = True
    End Sub

    Protected Sub Rb_NaoProc_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim sql As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        '' Monta a Sql ..
        Dim ds As New DataSet
        Dim Inconsist As String
        Dim lista As String
        Dim Sitban As Integer
        Dim Valtot As Decimal
        Dim Valind As Decimal
        lista = "N"
        If RbGeral.Checked = True Then
            lista = "G"
        End If
        If RbAEnviar.Checked = True Then
            lista = "A"
        End If
        If RbEnviado.Checked = True Then
            lista = "E"
        End If
        If RbBaixado.Checked = True Then
            lista = "B"
        End If
        If Rb_NaoProc.Checked = True Then
            lista = "P"
        End If
        '' sQL PARA CARREGAR REGISTROS DE PAGAMENTOS
        sql = " DECLARE @sql varchar(8000), @empresa varchar(18), @endempresa int, @datinicial varchar(8), @datfinal varchar(8) "
        sql &= " SET @empresa = ''"
        sql &= " SET @endempresa = -1"
        sql &= " CREATE TABLE #Destinatario"
        sql &= "("
        sql &= "	Destinatario_DES varchar(18), "
        sql &= "	Nome			varchar(60),"
        sql &= "	Inscricao		varchar(20),"
        sql &= "	Bairro			varchar(50),"
        sql &= "	Cidade			varchar(50),"
        sql &= "	Estado			varchar(2), "
        sql &= "    Cep             varchar(9)"
        sql &= ")"
        sql &= "SET @sql = 'INSERT INTO #Destinatario ' + "
        sql &= "        'SELECT CP.Destinatario, D.Nome, D.Inscricao, D.Bairro, D.Cidade, D.Estado, D.Cep ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'INNER JOIN Clientes D ' + "
        sql &= "         'ON D.Cliente_Id = CP.Destinatario ' + "
        sql &= "         'AND D.Endereco_Id = CP.EndDestinatario ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado. 
        ''sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
        lista = "N"
        If RbGeral.Checked = True Then
            '' NAO E INCLUIDA SQL POIS NAO CONSISTE NADA . 
            lista = "G"
        End If
        If Rb_NaoProc.Checked = True Then
            '' Consiste somente com situacao = 0
            sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
            lista = "P"
        End If
        If RbAEnviar.Checked = True Then
            '' Consiste somente com situacao = 
            sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
            lista = "A"
        End If
        If RbEnviado.Checked = True Then
            '' Consiste somente com situacao = 2 
            sql &= "       +  'AND (CP.SituacaoBancaria = 2)' "
            lista = "E"
        End If
        If RbBaixado.Checked = True Then
            '' Consiste somente com situacao = 3 
            sql &= "       +  'AND (CP.SituacaoBancaria = 3)' "
            lista = "B"
        End If
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " 	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        ''sql &= " SELECT * FROM #Destinatario "
        sql &= " CREATE TABLE #Contas "
        sql &= "("
        sql &= " 	Destinatario_Cont	varchar(18), "
        sql &= "	Banco_CONT			int, "
        sql &= "	Agencia_CONT		varchar(5), "
        sql &= " 	DvAgencia_CONT		varchar(2), "
        sql &= " 	Conta_CONT			varchar(20), "
        sql &= " 	DvConta_CONT		varchar(2), "
        sql &= "    TipoIns_CONT        varchar(9) "
        sql &= ") "
        sql &= "SET @sql = 'INSERT INTO #Contas ' + "
        sql &= "'SELECT CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, ' + "
        sql &= " 'CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END ' + "
        sql &= " 'FROM ContasAPagar CP ' + "
        sql &= " 'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao enviada , 1-Enviada, 2-baixada. 
        ''sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
        lista = "N"
        If RbGeral.Checked = True Then
            '' NAO E INCLUIDA SQL POIS NAO CONSISTE NADA . 
            lista = "G"
        End If
        If Rb_NaoProc.Checked = True Then
            '' Consiste somente com situacao = 0
            sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
            lista = "P"
        End If
        If RbAEnviar.Checked = True Then
            '' Consiste somente com situacao = 
            sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
            lista = "A"
        End If
        If RbEnviado.Checked = True Then
            '' Consiste somente com situacao = 2 
            sql &= "       +  'AND (CP.SituacaoBancaria = 2)' "
            lista = "E"
        End If
        If RbBaixado.Checked = True Then
            '' Consiste somente com situacao = 3 
            sql &= "       +  'AND (CP.SituacaoBancaria = 3)' "
            lista = "B"
        End If
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= " 'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        ''sql &= " SELECT * FROM #Contas "
        sql &= " CREATE TABLE #Pagamentos "
        sql &= "("
        sql &= "	TipoIns_pag			varchar(9),"
        sql &= "	Destinatario_pag	varchar(18), "
        sql &= "	Banco_pag			int,"
        sql &= "	Agencia_pag			varchar(5),"
        sql &= "	DvAgencia_pag		varchar(2),"
        sql &= "	Conta_pag			varchar(20),"
        sql &= "	DvConta_pag			varchar(2),"
        sql &= "	Vencimento_pag		datetime,"
        sql &= " 	Movimento_pag		datetime,"
        sql &= " 	ValorLiquido_pag	numeric(18, 9),"
        sql &= "	Finalidade_pag		bit,"
        sql &= "	Doc_pag				char(1),"
        sql &= "	CodigoDigitado_pag	char(1), "
        sql &= "	CodigoDeBarras_pag	varchar(50),"
        sql &= "	Modalidade_pag		varchar(2),"
        sql &= "	TipoDoc_pag			varchar(2),"
        sql &= "	TipoConta_pag		bit,"
        sql &= "	Duplicata_pag		int,"
        sql &= "    TipoPagto_pag       Int,"
        sql &= "    SituacaoBancaria_pag   Int"
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #Pagamentos ' +"
        sql &= "        'SELECT CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END, CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, ' + "
        sql &= "        'CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, CP.Vencimento, CP.Movimento, CP.ValorLiquido, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) > 0 THEN 1 ELSE 3 END, ' + "
        sql &= "        'CASE WHEN CP.BancoCliente = 237 THEN ''C'' ELSE ''D'' END, CP.CodigoDigitado, CP.CodigoDeBarras, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) = 0 THEN ''01'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente = 237 THEN ''05'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente <> 237 THEN ''08'' ' + "
        sql &= "        'ELSE ''31'' END, ''05'', 1, CP.Registro_Id, CP.TipoPagto, Cp.SituacaoBancaria ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' +"
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 ) ' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao enviada , 1-Enviada, 2-baixada. 
        ''sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
        lista = "N"
        If RbGeral.Checked = True Then
            '' NAO E INCLUIDA SQL POIS NAO CONSISTE NADA . 
            lista = "G"
        End If
        If Rb_NaoProc.Checked = True Then
            '' Consiste somente com situacao = 0
            sql &= "       +  'AND (CP.SituacaoBancaria = 0)' "
            lista = "P"
        End If
        If RbAEnviar.Checked = True Then
            '' Consiste somente com situacao = 
            sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
            lista = "A"
        End If
        If RbEnviado.Checked = True Then
            '' Consiste somente com situacao = 2 
            sql &= "       +  'AND (CP.SituacaoBancaria = 2)' "
            lista = "E"
        End If
        If RbBaixado.Checked = True Then
            '' Consiste somente com situacao = 3 
            sql &= "       +  'AND (CP.SituacaoBancaria = 3)' "
            lista = "B"
        End If
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= "	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' '"
        sql &= " EXEC (@sql) "
        ''sql &= " SELECT * FROM #Pagamentos "
        sql &= " CREATE TABLE #GERAL "
        sql &= " ("
        sql &= "	Destinatario_DES_GER      varchar(18), "
        sql &= "	Nome_DES_GER              varchar(60), "
        sql &= "	Inscricao_DES_GER         varchar(20), "
        sql &= "	Bairro_DES_GER            varchar(50), "
        sql &= "	Cidade_DES_GER            varchar(50), "
        sql &= "	Estado_DES_GER            varchar(2),  "
        sql &= " 	Cep_DES_GER               varchar(9),  "
        sql &= "	Destinatario_CONT_GER     varchar(18), "
        sql &= "	Banco_CONT_GER            int,"
        sql &= "	Agencia_CONT_GER          varchar(5),  "
        sql &= "	DvAgencia_CONT_GER        varchar(2),  "
        sql &= "	Conta_CONT_GER            varchar(20), "
        sql &= "	DvConta_CONT_GER          varchar(2),  "
        sql &= "	TipoIns_CONT_GER          varchar(9),  "
        sql &= " 	TipoIns_pag_GER           varchar(9),  "
        sql &= "	Destinatario_pag_GER      varchar(18), "
        sql &= "	Banco_pag_GER             int, "
        sql &= "	Agencia_pag_GER           varchar(5),  "
        sql &= "	DvAgencia_pag_GER         varchar(2),  "
        sql &= "	Conta_pag_GER             varchar(20), "
        sql &= " 	DvConta_pag_GER		  varchar(2), "
        sql &= "	Vencimento_pag_GER        datetime, "
        sql &= " 	Movimento_pag_GER         datetime,"
        sql &= "	ValorLiquido_pag_GER    numeric(18, 2),"
        sql &= " 	Finalidade_pag_GER        bit,"
        sql &= " 	Doc_pag_GER				char(1),"
        sql &= "	CodigoDigitado_pag_GER	char(1),"
        sql &= " 	CodigoDeBarras_pag_GER	varchar(50), "
        sql &= "	Modalidade_pag_GER	varchar(2), "
        sql &= "	TipoDoc_pag_GER		varchar(2), "
        sql &= "	TipoConta_pag_GER	bit, "
        sql &= "	Duplicata_pag_GER	int, "
        sql &= "    TipoPagto_pag_GER   Int, "
        sql &= "    SituacaoBancaria_pag_ger  Int "
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #GERAL ' + "
        sql &= "'sELECT De.Destinatario_des, DE.nome, De.Inscricao, De.bairro, De.cidade, De.Estado, De.Cep, ' + "
        sql &= " 'ct.Destinatario_cont, Ct.Banco_cont, Ct.Agencia_cont, Ct.Dvagencia_cont, Ct.Conta_cont, Ct.Dvconta_cont, ' + "
        sql &= " 'ct.TipoIns_cont, pg.Tipoins_pag, pg.Destinatario_pag, Pg.Banco_pag, Pg.Agencia_pag, pg.Dvagencia_pag,  '  +"
        sql &= " 'pg.Conta_pag, pg.DvConta_pag, pg.VencimentO_pag, Pg.Movimento_pag, Pg.ValorLiquido_pag, Pg.Finalidade_pag, ' +"
        sql &= " 'Pg.Doc_pag, Pg.CodigoDigitado_pag, Pg.Codigodebarras_pag, pg.Modalidade_pag, Pg.TipoDoc_pag, ' + "
        sql &= " 'Pg.TipoConta_pag, Pg.Duplicata_pag, Pg.TipoPagto_pag, Pg.SituacaoBancaria_pag '+"
        sql &= " 'FROM #PAGAMENTOS PG ' +"
        sql &= " 'INNER JOIN #DESTINATARIO DE ON DE.DESTINATARIO_DES = PG.DESTINATARIO_PAG ' +"
        sql &= " 'INNER JOIN #CONTAS CT ON CT.DESTINATARIO_CONT = PG.DESTINATARIO_PAG AND CT.BANCO_CONT = PG.BANCO_PAG AND CT.AGENCIA_CONT = PG.AGENCIA_PAG AND CT.CONTA_CONT = PG.CONTA_PAG '"
        sql &= " EXEC (@sql) "
        ''sql &= " SELECT * FROM #Pagamentos"
        sql &= " SELECT * FROM #GERAL"
        sql &= " DROP TABLE #Destinatario"
        sql &= " DROP TABLE #Pagamentos"
        sql &= " DROP TABLE #Contas"
        sql &= " DROP TABLE #GERAL"
        '' SQL PARA CARREGAR REGISTROS DE PAGAMENTOS.
        '' rotina utilizada nas procuracoes inicio - html 
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf
        linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
        linha &= "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf
        linha &= "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        '-----------------
        'Cabeçalho Padrao
        '-----------------
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "<TD >" & " Unidade : " & ddlUnidadeGeracaoBradesco.SelectedItem.Text & " </TD>"
        linha &= "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>"
        linha &= "<TD >" & " Conta Bancaria : " & DdlContaConsulta.SelectedItem.Text & " </TD>"
        linha &= "<TD >" & " Banco Bradesco  " & " </TD>"
        '' opcao escolhida
        If RbGeral.Checked = True Then
            '' NAO E INCLUIDA SQL POIS NAO CONSISTE NADA . 
            linha &= "<TD >" & " Opcao de relatorio selecionado : (Geral)" & " </TD>"
        End If
        If Rb_NaoProc.Checked = True Then
            '' Consiste somente com situacao = 0
            linha &= "<TD >" & " Opcao de relatorio selecionado : (Nao Processado) " & " </TD>"
        End If
        If RbAEnviar.Checked = True Then
            '' Consiste somente com situacao = 
            linha &= "<TD >" & " Opcao de relatorio selecionado : (A Enviar) " & " </TD>"
        End If
        If RbEnviado.Checked = True Then
            '' Consiste somente com situacao = 2 
            linha &= "<TD >" & " Opcao de relatorio selecionado : (Enviado ao Banco) " & " </TD>"
        End If
        If RbBaixado.Checked = True Then
            '' Consiste somente com situacao = 3 
            linha &= "<TD >" & " Opcao de relatorio selecionado : (Baixado Banco) " & " </TD>"
        End If
        '' opcao escolhida
        linha &= "</TR>"
        linha &= "<table width= '370' cellpadding='0' cellspacing='0' Border=1>"
        linha &= "<TR>"
        linha &= "<TD >Destinatario_DES_GER</TD>"
        linha &= "<TD >Nome_DES_GER</TD>"
        linha &= "<TD >Duplicata_pag_GER</TD>"
        linha &= "<TD >ValorLiquido_pag_GER</TD>"
        linha &= "<TD >Movimento_pag_GER</TD>"
        linha &= "<TD >Vencimento_pag_GER</TD>"
        linha &= "<TD >Dados complementares</TD>"
        linha &= "<TD >Inconsistencia</TD>"
        linha &= "<TD >Banco_CONT_GER</TD>"
        linha &= "<TD >Agencia_CONT_GER</TD>"
        linha &= "<TD >DvAgencia_CONT_GER</TD>"
        linha &= "<TD >Conta_CONT_GERTD<>"
        linha &= "<TD >DvConta_CONT_GER</TD>"
        linha &= "<TD >Inscricao_DES_GER</TD>"
        linha &= "<TD >Bairro_DES_GER</TD>"
        linha &= "<TD >Cidade_DES_GER</TD>"
        linha &= "<TD >Estado_DES_GER</TD>"
        linha &= "<TD >Cep_DES_GER</TD>"
        linha &= "<TD >TipoIns_CONT_GER</TD>"
        linha &= "<TD >TipoIns_pag_GER</TD>"
        linha &= "<TD >Finalidade_pag_GER</TD>"
        linha &= "<TD >Doc_pag_GER</TD>"
        linha &= "<TD >CodigoDigitado_pag_GER</TD>"
        linha &= "<TD >CodigoDeBarras_pag_GER</TD>"
        linha &= "<TD >Modalidade_pag_GER</TD>"
        linha &= "<TD >TipoDoc_pag_GER</TD>"
        linha &= "<TD >TipoConta_pag_GER</TD>"
        linha &= "<TD >TipoPagto_pag_GER</TD>"
        linha &= "<TD >SituacaoBancaria_pag_GER</TD>"
        linha &= "</TR>"
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Dim registro As String

        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                linha &= "<TR>"
                registro = dr("Duplicata_pag_GER")
                linha &= "<TD>" & dr("Destinatario_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Nome_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Duplicata_pag_GER") & "</TD>"
                Valind = dr("ValorLiquido_pag_GER")
                Valtot = Valtot + Valind
                linha &= "<TD>" & dr("ValorLiquido_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("Movimento_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("Vencimento_pag_GER") & "</TD>"
                linha &= "<TD>" & "Dados complementares" & "</TD>"
                Inconsist = "."
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " Ted com valor menor que 5000 nao sera enviado ao banco !!! Favor Corrigir !!!"
                    conterro = conterro + 1
                End If
                If Inconsist = "." Then
                    contreg = contreg + 1
                End If
                linha &= "<TD>" & Inconsist & "</TD>"
                linha &= "<TD>" & dr("Banco_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("Agencia_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("DvAgencia_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("Conta_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("DvConta_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("Inscricao_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Bairro_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Cidade_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Estado_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("Cep_DES_GER") & "</TD>"
                linha &= "<TD>" & dr("TipoIns_CONT_GER") & "</TD>"
                linha &= "<TD>" & dr("Destinatario_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("Finalidade_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("Doc_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("CodigoDigitado_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("CodigoDeBarras_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("Modalidade_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("TipoDoc_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("TipoConta_pag_GER") & "</TD>"
                linha &= "<TD>" & dr("TipoPagto_pag_GER") & "</TD>"
                Sitban = dr("SituacaoBancaria_pag_ger")
                If Sitban = 0 Then
                    linha &= "<TD>" & "0-NAO PROCESSADO!!" & "</TD>"
                End If
                If Sitban = 1 Then
                    linha &= "<TD>" & "1-A ENVIAR!!" & "</TD>"
                End If
                If Sitban = 2 Then
                    linha &= "<TD>" & "2-ENVIADO AO BANCO!!" & "</TD>"
                End If
                If Sitban = 3 Then
                    linha &= "<TD>" & "3-BAIXADO!!" & "</TD>"
                End If
                ''linha &= "<TD>" & dr("SituacaoBancaria_pag_ger") & "</TD>"
                linha &= "</TR>"
                ''
                ''linha &= "</TR>"
                ''
            Next
        End If
        linha &= "<TR>"
        linha &= "<TD>" & "." & "</TD>"
        linha &= "<TD>" & "." & "</TD>"
        linha &= "<TD>" & "Valor total neste processamento:  " & "</TD>"
        linha &= "<TD >" & Valtot & "</TD>"
        linha &= "</TR>"
        If contreg = 0 Then
            MsgBox(Me.Page, "Nao existem registros corretos para o periodo.", eTitulo.Info)
            txtPeriodoFinalConsultaTitulos.Focus()
        Else
            If conterro > 0 Then
                MsgBox(Me.Page, "'Existem registros processados , mas tambem existem inconsistencias no Movimento favor verificar.", eTitulo.Info)
                txtPeriodoFinalConsultaTitulos.Focus()
            Else
                MsgBox(Me.Page, "Movimento com registros processados e sem erros basicos.", eTitulo.Info)
                txtPeriodoFinalConsultaTitulos.Focus()
            End If
        End If

        strm = New StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            Funcoes.AbrirArquivo(Me.Page, NomeArquivo)
        Finally
            strm.Close()
        End Try
    End Sub

    Protected Sub lnkGerar_Click(sender As Object, e As EventArgs) Handles lnkGerar.Click
        Dim Mensagem As String = "Gera arquivos"
        GERFOR()
        gercon()
        gerpag()
    End Sub

    Protected Sub lnkBaixa_Click(sender As Object, e As EventArgs) Handles lnkBaixa.Click
        Dim sql As String
        Dim sql2 As String
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim dataproc As Date
        dataproc = Date.Today
        Dim ds As New DataSet
        Dim datatu As Date = Today
        Dim SqlArray As New ArrayList
        Dim Mensagem As String = "Baixa por lotes"
        Dim j As Integer
        '' DEFINICAO DE VARIAVEIS - FIM 
        ''  SQL DE PESQUISA - inicio 
        sql = " DECLARE @sql varchar(8000), @empresa varchar(18), @endempresa int "
        sql &= " SET @empresa = ''"
        sql &= " SET @endempresa = -1"
        sql &= " CREATE TABLE #Destinatario"
        sql &= "("
        sql &= "	Destinatario_DES varchar(18), "
        sql &= "	Nome			varchar(60),"
        sql &= "	Inscricao		varchar(20),"
        sql &= "	Bairro			varchar(50),"
        sql &= "	Cidade			varchar(50),"
        sql &= "	Estado			varchar(2), "
        sql &= "    Cep             varchar(9)"
        sql &= ")"
        sql &= "SET @sql = 'INSERT INTO #Destinatario ' + "
        sql &= "        'SELECT CP.Destinatario, D.Nome, D.Inscricao, D.Bairro, D.Cidade, D.Estado, D.Cep ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'INNER JOIN Clientes D ' + "
        sql &= "         'ON D.Cliente_Id = CP.Destinatario ' + "
        sql &= "         'AND D.Endereco_Id = CP.EndDestinatario ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao enviada , 1-Enviada, 2-baixada. 
        sql &= "       +  'AND (CP.SituacaoBancaria = 1)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " 	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Contas "
        sql &= "("
        sql &= " 	Destinatario_Cont	varchar(18), "
        sql &= "	Banco_CONT			int, "
        sql &= "	Agencia_CONT		varchar(5), "
        sql &= " 	DvAgencia_CONT		varchar(2), "
        sql &= " 	Conta_CONT			varchar(20), "
        sql &= " 	DvConta_CONT		varchar(2), "
        sql &= "    TipoIns_CONT        varchar(9) "
        sql &= ") "
        sql &= "SET @sql = 'INSERT INTO #Contas ' + "
        sql &= "'SELECT CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, ' + "
        sql &= " 'CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END ' + "
        sql &= " 'FROM ContasAPagar CP ' + "
        sql &= " 'WHERE LEN(CP.UsuarioLiberacao) > 0 ' + "
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado.  
        sql &= "       +  'AND (CP.SituacaoBancaria = 2)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= " SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= " 'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' ' "
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #Pagamentos "
        sql &= "("
        sql &= "	TipoIns_pag			varchar(9),"
        sql &= "	Destinatario_pag	varchar(18), "
        sql &= "	Banco_pag			int,"
        sql &= "	Agencia_pag			varchar(5),"
        sql &= "	DvAgencia_pag		varchar(2),"
        sql &= "	Conta_pag			varchar(20),"
        sql &= "	DvConta_pag			varchar(2),"
        sql &= "	Vencimento_pag		datetime,"
        sql &= " 	Movimento_pag		datetime,"
        sql &= " 	ValorLiquido_pag	numeric(18, 9),"
        sql &= "	Finalidade_pag		bit,"
        sql &= "	Doc_pag				char(1),"
        sql &= "	CodigoDigitado_pag	char(1), "
        sql &= "	CodigoDeBarras_pag	varchar(50),"
        sql &= "	Modalidade_pag		varchar(2),"
        sql &= "	TipoDoc_pag			varchar(2),"
        sql &= "	TipoConta_pag		bit,"
        sql &= "	Duplicata_pag		int,"
        sql &= "    TipoPagto_pag       Int"
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #Pagamentos ' +"
        sql &= "        'SELECT CASE WHEN LEN(CP.Destinatario) > 13 THEN 1 ELSE 0 END, CP.Destinatario, CP.BancoCliente, CP.AgenciaCliente, ' + "
        sql &= "        'CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente, CP.Vencimento, CP.Movimento, CP.ValorLiquido, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) > 0 THEN 1 ELSE 3 END, ' + "
        sql &= "        'CASE WHEN CP.BancoCliente = 237 THEN ''C'' ELSE ''D'' END, CP.CodigoDigitado, CP.CodigoDeBarras, ' + "
        sql &= "        'CASE WHEN LEN(CP.ContaCliente) = 0 THEN ''01'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente = 237 THEN ''05'' ' + "
        sql &= "        'WHEN CP.TipoPagto = 6 AND CP.BancoCliente <> 237 THEN ''08'' ' + "
        sql &= "        'ELSE ''31'' END, ''05'', 1, CP.Registro_Id, CP.TipoPagto ' + "
        sql &= "        'FROM ContasAPagar CP ' + "
        sql &= "         'WHERE LEN(CP.UsuarioLiberacao) > 0 ' +"
        '' ROTINA original 
        ''sql &= "         'AND CP.TipoPagto <> 2 ' "
        '' rotina nova.
        sql &= "         'AND (CP.TipoPagto = 3 or CP.TipoPagto = 4 or CP.TipoPagto = 5 or CP.TipoPagto = 6 )' +  "
        sql &= "         'AND (CP.BancoPagador = 237)' "
        '' rotina nova - fim
        '' rotina nova - consistencia bancaria Situacao bancaria 0-Nao processado , 1-a Enviar, 2-Enviado, 3-baixado.  
        sql &= "       +  'AND (CP.SituacaoBancaria = 2)' "
        '' rotina nova - consistencia bancaria . 
        '' consistencia de data.
        sql &= "         + 'AND CP.Baixa between  ''" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "'' "
        sql &= "         and ''" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'''"
        '' consistencia de data fim 
        sql &= " IF @empresa <> '' "
        sql &= "	SET @sql = @sql + 'AND CP.Empresa = ''' + @empresa + ''' ' + "
        sql &= "         'AND CP.EndEmpresa = ' + CAST(@endempresa AS varchar) + ' '"
        sql &= " EXEC (@sql) "
        sql &= " CREATE TABLE #GERAL "
        sql &= " ("
        sql &= "	Destinatario_DES_GER      varchar(18), "
        sql &= "	Nome_DES_GER              varchar(60), "
        sql &= "	Inscricao_DES_GER         varchar(20), "
        sql &= "	Bairro_DES_GER            varchar(50), "
        sql &= "	Cidade_DES_GER            varchar(50), "
        sql &= "	Estado_DES_GER            varchar(2),  "
        sql &= " 	Cep_DES_GER               varchar(9),  "
        sql &= "	Destinatario_CONT_GER     varchar(18), "
        sql &= "	Banco_CONT_GER            int,"
        sql &= "	Agencia_CONT_GER          varchar(5),  "
        sql &= "	DvAgencia_CONT_GER        varchar(2),  "
        sql &= "	Conta_CONT_GER            varchar(20), "
        sql &= "	DvConta_CONT_GER          varchar(2),  "
        sql &= "	TipoIns_CONT_GER          varchar(9),  "
        sql &= " 	TipoIns_pag_GER           varchar(9),  "
        sql &= "	Destinatario_pag_GER      varchar(18), "
        sql &= "	Banco_pag_GER             int, "
        sql &= "	Agencia_pag_GER           varchar(5),  "
        sql &= "	DvAgencia_pag_GER         varchar(2),  "
        sql &= "	Conta_pag_GER             varchar(20), "
        sql &= " 	DvConta_pag_GER		  varchar(2), "
        sql &= "	Vencimento_pag_GER        datetime, "
        sql &= " 	Movimento_pag_GER         datetime,"
        sql &= "	ValorLiquido_pag_GER    numeric(18, 2),"
        sql &= " 	Finalidade_pag_GER        bit,"
        sql &= " 	Doc_pag_GER				char(1),"
        sql &= "	CodigoDigitado_pag_GER	char(1),"
        sql &= " 	CodigoDeBarras_pag_GER	varchar(50), "
        sql &= "	Modalidade_pag_GER	varchar(2), "
        sql &= "	TipoDoc_pag_GER		varchar(2), "
        sql &= "	TipoConta_pag_GER	bit, "
        sql &= "	Duplicata_pag_GER	int, "
        sql &= "    TipoPagto_pag_GER   Int "
        sql &= ")"
        sql &= " SET @sql = 'INSERT INTO #GERAL ' + "
        sql &= "'sELECT De.Destinatario_des, DE.nome, De.Inscricao, De.bairro, De.cidade, De.Estado, De.Cep, ' + "
        sql &= " 'ct.Destinatario_cont, Ct.Banco_cont, Ct.Agencia_cont, Ct.Dvagencia_cont, Ct.Conta_cont, Ct.Dvconta_cont, ' + "
        sql &= " 'ct.TipoIns_cont, pg.Tipoins_pag, pg.Destinatario_pag, Pg.Banco_pag, Pg.Agencia_pag, pg.Dvagencia_pag,  '  +"
        sql &= " 'pg.Conta_pag, pg.DvConta_pag, pg.VencimentO_pag, Pg.Movimento_pag, Pg.ValorLiquido_pag, Pg.Finalidade_pag, ' +"
        sql &= " 'Pg.Doc_pag, Pg.CodigoDigitado_pag, Pg.Codigodebarras_pag, pg.Modalidade_pag, Pg.TipoDoc_pag, ' + "
        sql &= " 'Pg.TipoConta_pag, Pg.Duplicata_pag, Pg.TipoPagto_pag '+"
        sql &= " 'FROM #PAGAMENTOS PG ' +"
        sql &= " 'INNER JOIN #DESTINATARIO DE ON DE.DESTINATARIO_DES = PG.DESTINATARIO_PAG ' +"
        sql &= " 'INNER JOIN #CONTAS CT ON CT.DESTINATARIO_CONT = PG.DESTINATARIO_PAG AND CT.BANCO_CONT = PG.BANCO_PAG AND CT.AGENCIA_CONT = PG.AGENCIA_PAG AND CT.CONTA_CONT = PG.CONTA_PAG '"
        sql &= " EXEC (@sql) "
        sql &= " SELECT * FROM #GERAL"
        sql &= " DROP TABLE #Destinatario"
        sql &= " DROP TABLE #Pagamentos"
        sql &= " DROP TABLE #Contas"
        ''  SQL DE PESQUISA - FIM 
        '' Baixa por lote Grava registro como enviado baixa efetiva Situacao bancaria = 2 - inicio 
        dsRelatorio = Banco.ConsultaDataSet(sql, "Geral")
        sql2 = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                If dr("TipoPagto_pag_GER") = 6 And dr("ValorLiquido_pag_GER") < 5000 Then
                    Inconsist = " ."
                Else
                    Inconsist = " ."
                    sql2 = " UPDATE ContasAPagar"
                    sql2 &= " SET SituacaoBancaria = '" & "3" & "'"
                    sql2 &= ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"                     'Usuario que Alterou
                    sql2 &= ", UsuarioAlteracaoData = '" & CDate(datatu).ToSqlDate() & "'"   'Data da Alteracao
                    sql2 &= " WHERE Registro_ID = " & CInt(dr("Duplicata_pag_GER"))
                    SqlArray.Add(sql2)
                    j = j + 1
                End If
            Next
            If j > 0 Then
                If Banco.GravaBanco(SqlArray) = False Then
                    txtMensagem.Text = HttpContext.Current.Session("ssMessage")
                Else
                    Mensagem = "Registros Atualizados com sucesso."
                    txtMensagem.Text = Mensagem
                End If
            End If
            If j > 0 Then
                MsgBox(Me.Page, "Registros atualizados com Sucesso.", eTitulo.Sucess)
                txtPeriodoFinalConsultaTitulos.Focus()
            Else
                MsgBox(Me.Page, "Não Existiam registros a Atualizar.", eTitulo.Info)
                txtPeriodoFinalConsultaTitulos.Focus()
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "GeracaoBradesco")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class
