Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RegistrosDeISS
    Inherits BasePage

    Dim Sql As String
    Dim Sqla As String
    Dim Row As DataRow

    Dim ds As New DataSet
    Dim dss As New DataSet

    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim Cliente() As String
    Dim SequenciaLote As Integer = 0

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Descricao As String = ""

    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date

    Dim Processo As Integer = 0
    Dim Livro As Integer = 0
    Dim Folha As Integer = 0
    Dim Opcao As String = ""

    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RegistrosDeISS", "ACESSAR") Then
                    CarregarUnidade()
                    VerificaUnidade()
                    CriaOpcoes()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf & _
              "  FROM Clientes C " & vbCrLf & _
              " INNER JOIN ClientesXTipos CT " & vbCrLf & _
              "    ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf & _
              " WHERE CT.Tipo_Id = 050 " & vbCrLf & _
              " ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            CarregarProcesso()
        Next
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
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

    Protected Sub DdlEmpresa_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProcesso()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlProcesso_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            HttpContext.Current.Session("ProcessoIpi") = DdlProcesso.SelectedValue

            Sql = "SELECT * " & vbCrLf & _
                  "  FROM ProcessoRAISS " & vbCrLf & _
                  " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
                  "   And Processo_Id = " & DdlProcesso.SelectedValue & vbCrLf & _
                  " order by PeriodoInicial DESC"

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                txtDataInicial.Text = Dr("PeriodoInicial")
                txtDataFinal.Text = Dr("PeriodoFinal")
                txtLivro.Text = Dr("Livro")
                txtFolha.Text = Dr("PaginaInicial")
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        HttpContext.Current.Session("EmpresaIpi") = Empresa(0)
        DdlProcesso.Items.Clear()

        Sql = "SELECT Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado " & vbCrLf & _
              "  FROM ProcessoRAISS " & vbCrLf & _
              " WHERE Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              " order by PeriodoInicial DESC" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Descricao = "Processo " & Format(Dr("Processo_Id"), "0000") & "      Período  " & vbCrLf & _
                        Format(Dr("PeriodoInicial"), "dd/MM/yyyy") & "  à  " & vbCrLf & _
                        Format(Dr("PeriodoFinal"), "dd/MM/yyyy") & "  Livro  " & vbCrLf & _
                        Format(Dr("Livro"), "000") & "  Folha  " & vbCrLf & _
            Format(Dr("PaginaInicial"), "000") & vbCrLf

            DdlProcesso.Items.Add(New ListItem(Descricao, Dr("Processo_ID")))
        Next

        DdlProcesso.Items.Insert(0, "")
        DdlProcesso.SelectedIndex = 0

    End Sub

    Private Sub CriaOpcoes()
        Dim i As Integer = 0
        Dim dsOpcoes As New DataSet
        Dim dtOpcoes As DataTable
        Dim Opcoes As Integer = 2

        dtOpcoes = New DataTable("Opcoes")
        dtOpcoes.Columns.Add("Codigo", Type.GetType("System.String")).DefaultValue = ""
        dtOpcoes.Columns.Add("Descricao", Type.GetType("System.String")).DefaultValue = ""

        While i < Opcoes
            Row = dtOpcoes.NewRow()
            Select Case i
                Case 0
                    Row("Codigo") = "01"
                    Row("Descricao") = "Livro Registro de ISS"
                Case 1
                    Row("Codigo") = "02"
                    Row("Descricao") = "Termos de Abertura e Encerramento"
            End Select

            dtOpcoes.Rows.Add(Row)
            i += 1
        End While

        GridOpcoes.DataSource = dtOpcoes
        GridOpcoes.DataBind()

    End Sub

    Protected Sub GridOpcoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Validar() Then
                Select Case GridOpcoes.SelectedRow.Cells(1).Text()
                    Case "01"
                        RegistrosDeISS()
                    Case "02"
                        Opcao = "20"
                        Termos()
                End Select
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#Region "Registro Fiscais ISS"

    Private Sub RegistrosDeISS()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
            PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text
            Dim EmpresaInscricaoMunicipal As String = ""
            Dim EmpresaEndereco As String = ""

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
                Dim SqlArray As New ArrayList
                Sql = " SELECT Clientes.Cliente_Id AS Cliente, Clientes.Nome, Clientes.Endereco, Clientes.Inscricao, isnull(ClientesXEmpresas.InscricaoMunicipal, '-') as InscricaoMunicipal"
                Sql &= " FROM   Clientes LEFT OUTER JOIN"
                Sql &= " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id"
                Sql &= " WHERE  Clientes.Cliente_Id = '" & Empresa(0) & "'"

                ds = Banco.ConsultaDataSet(Sql, "Clientes")

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    EmpresaNome = Dr("Nome")
                    EmpresaCNPJ = Funcoes.FormatarCpfCnpj(Dr("Cliente"))
                    EmpresaInscricao = Dr("Inscricao")
                    EmpresaEndereco = Dr("Endereco")
                    EmpresaInscricaoMunicipal = Dr("InscricaoMunicipal")
                    Exit For
                Next
            End If

            Dim Periodo As String = "Mês ou Período/Ano : " & Format(PeriodoInicial, "dd/MM/yyyy") & " a " & Format(PeriodoFinal, "dd/MM/yyyy")

            'Sql = "  SELECT Cliente_Id AS Cliente, EntradaSaida_Id AS EntradaSaida, Serie_Id AS Serie,  Nota_Id AS Nota, Movimento, DataNota, CodigoFiscal as CFOP, Valor, "
            'Sql &= "  COALESCE(BaseISS, 0) AS BaseISS, COALESCE(AliquotaISS, 0) AS AliquotaISS, COALESCE(ValorISS, 0) AS ValorISS"
            'Sql &= " FROM RegistrosFiscais"
            'Sql &= " WHERE RegistrosFiscais.Empresa_Id = '" & Empresa(0) & "' AND RegistrosFiscais.EntradaSaida_Id = 'S' And "
            'Sql &= " RegistrosFiscais.Movimento BETWEEN '" & Format(PeriodoInicial, "yyyy/MM/dd") & "' AND '" & Format(PeriodoFinal, "yyyy/MM/dd") & "' and"
            'Sql &= " RegistrosFiscais.Operacao = 51 AND (RegistrosFiscais.SubOperacao = 91 OR RegistrosFiscais.SubOperacao = 92) "
            'Sql &= " Order by RegistrosFiscais.Movimento, RegistrosFiscais.Serie_Id, RegistrosFiscais.Nota_Id"

            Sql = " SELECT	Cliente, EntradaSaida, Serie, NumeroDaNota as Nota," & vbCrLf & _
                  "        Movimento, DataDaNota, CFOP,  " & vbCrLf & _
                  " 	    ValorContabil as Valor, " & vbCrLf & _
                  " 	    CFOP, " & vbCrLf & _
                  " 	    BaseISS, Conta, " & vbCrLf & _
                  " 	    AliquotaISS AS Aliquota," & vbCrLf & _
                  "        ValorISS" & vbCrLf & _
                  " FROM" & vbCrLf & _
                  "    (SELECT NotasFiscais.Nota_Id AS NumeroDaNota, NotasFiscais.EntradaSaida_Id AS EntradaSaida, NotasFiscais.Serie_Id AS Serie, NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.Movimento, " & vbCrLf & _
                  "            NotasFiscais.DataDaNota, NotasFiscaisXItens.CFOP_Id AS CFOP, NotasFiscais.EstadoDoCliente AS Estado, SubOperacoes.GrupoDeContas as Conta," & vbCrLf & _
                  "            Case when NotasFiscais.Situacao = 1 then SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'PRODUTO' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) else 0 end AS ValorContabil, " & vbCrLf & _
                  "            Case when NotasFiscais.Situacao = 1 then SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'ISS' THEN NotasFiscaisXEncargos.Base ELSE 0 END) else 0 end AS BaseISS, " & vbCrLf & _
                  "            Case when NotasFiscais.Situacao = 1 then SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'ISS' THEN NotasFiscaisXEncargos.Valor ELSE 0 END) else 0 end AS ValorISS, " & vbCrLf & _
                  "            Case when NotasFiscais.Situacao = 1 then SUM(CASE WHEN NotasFiscaisXEncargos.Encargo_id = 'ISS' THEN NotasFiscaisXEncargos.Percentual ELSE 0 END) else 0 end AS AliquotaISS " & vbCrLf & _
                  " FROM       NotasFiscais INNER JOIN" & vbCrLf & _
                  "            NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf & _
                  "            NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND " & vbCrLf & _
                  "            NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND " & vbCrLf & _
                  "            NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND " & vbCrLf & _
                  "            NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
                  "            NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND " & vbCrLf & _
                  "            NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND" & vbCrLf & _
                  "            NotasFiscaisXItens.Sequencia_id = NotasFiscaisXEncargos.Sequencia_id AND NotasFiscaisXEncargos.Encargo_Id IN ('ISS', 'PRODUTO') " & vbCrLf & _
                  " INNER Join" & vbCrLf & _
                  "            SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf & _
                  "            NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf

            Sql &= " WHERE      (NOT (NotasFiscais.Serie_Id IN ('D', 'F', 'REC'))) And (NotasFiscais.Empresa_Id = '" & Empresa(0) & "') AND (NotasFiscais.EntradaSaida_Id = 'S') AND " & vbCrLf & _
                   "            (NotasFiscais.Movimento BETWEEN '" & PeriodoInicial.ToSqlDate() & "' AND '" & PeriodoFinal.ToSqlDate() & "') And" & vbCrLf & _
                   "            NotasFiscaisXItens.CFOP_ID IN (5933, 6933) " & vbCrLf & _
                   "            GROUP BY NotasFiscais.Nota_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Cliente_Id, NotasFiscais.Movimento, NotasFiscais.DataDaNota, NotasFiscais.Situacao, " & vbCrLf & _
                   "            NotasFiscaisXItens.CFOP_Id, SubOperacoes.GrupoDeContas, NotasFiscais.EstadoDoCliente) AS Consulta" & vbCrLf & _
                   " ORDER BY    Movimento, Serie, NumeroDaNota" & vbCrLf
            ds.Merge(Banco.ConsultaDataSet(Sql, "RegistroDeApuracaoDoISS"))


            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("EmpresaNome", EmpresaNome)
            parameters.Add("EmpresaCNPJ", EmpresaCNPJ)
            parameters.Add("EmpresaInscricao", EmpresaInscricao)
            parameters.Add("EmpresaInscricaoMunicipal", EmpresaInscricaoMunicipal)
            parameters.Add("EmpresaEndereco", EmpresaEndereco)
            parameters.Add("Periodo", Periodo)
            parameters.Add("Livro", "Livro.: " & Format(CInt(Livro), "000"))
            parameters.Add("Pagina", Folha)

            Funcoes.BindReport(Me.Page, ds, "Cr_RegistroDeApuracaoDoISS", eExportType.PDF, parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Termos"

    Private Sub Termos()
        Try
            Empresa = DdlEmpresa.SelectedValue.Split("-")
            PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
            PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
            Processo = DdlProcesso.SelectedValue
            Livro = txtLivro.Text
            Folha = txtFolha.Text

            If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
                Dim Ds As New DataSet

                Dim DIAINI As String
                Dim MESINI As String
                Dim ANOINI As String

                Dim DIAFIM As String
                Dim MESFIM As String
                Dim ANOFIM As String

                Dim LIVRO01 As String

                MESINI = LTrim(RTrim(Str(Month(PeriodoInicial))))
                If Val(MESINI) < 10 Then
                    MESINI = "0" + MESINI
                End If
                MESINI = MonthName(MESINI)
                DIAINI = LTrim(RTrim(Str(Day(PeriodoInicial))))
                If Val(DIAINI) < 10 Then
                    DIAINI = "0" + DIAINI
                End If
                ANOINI = Str(Year(PeriodoInicial))

                MESFIM = LTrim(RTrim(Str(Month(PeriodoFinal))))
                If Val(MESFIM) < 10 Then
                    MESFIM = "0" + MESFIM
                End If
                MESFIM = MonthName(MESFIM)
                DIAFIM = Str(Day(PeriodoFinal))
                If Val(DIAFIM) < 10 Then
                    DIAFIM = "0" + DIAFIM
                End If
                ANOFIM = Str(Year(PeriodoFinal))
                LIVRO01 = Format(Val(Livro), "000")

                Dim mystr As String
                mystr = Format(Val(Folha), "000")
                Folha = mystr

                Sql = "  SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                      " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                      " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                      " Estados.Descricao as NomeDoEstado, 1 as Ordem" & vbCrLf & _
                      " FROM Clientes INNER JOIN" & vbCrLf & _
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                      " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "'" & vbCrLf & _
                      " Union" & vbCrLf & _
                      " SELECT Clientes.Cliente_Id as CNPJ, Clientes.Nome as RazaoSocial, Clientes.Endereco, Clientes.Cidade, Clientes.Estado, Clientes.Inscricao, ClientesXEmpresas.RegistroNaJunta, " & vbCrLf & _
                      " ClientesXEmpresas.EstadodaJunta, ClientesXEmpresas.DatadeFundacao, ClientesXEmpresas.NomeDoContador, ClientesXEmpresas.CPFContador, " & vbCrLf & _
                      " ClientesXEmpresas.CRCContador, ClientesXEmpresas.NomeDoTitular, ClientesXEmpresas.CPFTitular, ClientesXEmpresas.CondicaoTitular, " & vbCrLf & _
                      " Estados.Descricao as NomeDoEstado, 2 as Ordem" & vbCrLf & _
                      " FROM Clientes INNER JOIN" & vbCrLf & _
                      " ClientesXEmpresas ON Clientes.Cliente_Id = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                      " Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id INNER JOIN" & vbCrLf & _
                      " Estados ON ClientesXEmpresas.EstadodaJunta = Estados.Estado_Id" & vbCrLf & _
                      " WHERE Clientes.Cliente_Id Like '" & Empresa(0) & "'" & vbCrLf
                Ds.Merge(Banco.ConsultaDataSet(Sql, "Termos"))

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Folha", Folha)
                parameters.Add("DIAINI", DIAINI)
                parameters.Add("MESINI", MESINI)
                parameters.Add("ANOINI", ANOINI)
                parameters.Add("DIAFIM", DIAFIM)
                parameters.Add("MESFIM", MESFIM)
                parameters.Add("ANOFIM", ANOFIM)
                parameters.Add("LIVRO", LIVRO01)
                parameters.Add("Opcao", Opcao)

                Funcoes.BindReport(Me.Page, Ds, "Cr_TermosRegistrosFiscaisICMS", eExportType.PDF, parameters)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Funcoes"

    Private Function GravarProcesso(ByVal Processo As String, ByVal Livro As String, ByVal Folha As String, ByVal PeriodoIni As String, ByVal PeriodoFim As String, ByVal Cliente As String) As Boolean
        Dim strSQL As String

        strSQL = "SELECT * FROM ProcessoRAISS WHERE Empresa_Id = '" & Cliente & "' And Processo_Id = " & Processo

        Dim dsProcesso As DataSet = Banco.ConsultaDataSet(strSQL, "ProcessoRAISS")

        If dsProcesso.Tables(0).Rows.Count = 0 Then
            strSQL = "INSERT INTO ProcessoRAISS " & vbCrLf & _
                     "(Empresa_Id, Processo_Id, PeriodoInicial, PeriodoFinal, Livro, PaginaInicial, Fechado) " & vbCrLf & _
                     "VALUES ('" & Cliente & "', " & Processo & ", '" & PeriodoIni.ToSqlDate() & "', " & vbCrLf & _
                     "'" & PeriodoFim.ToSqlDate() & "', " & Livro & ", " & Folha & ", 'N')" & vbCrLf
        Else
            strSQL = "UPDATE ProcessoRAISS " & vbCrLf & _
                     "SET PeriodoInicial = '" & PeriodoIni.ToSqlDate() & "', " & vbCrLf & _
                     "PeriodoFinal = '" & PeriodoFim.ToSqlDate() & "', " & vbCrLf & _
                     "Livro = " & Livro & ", " & vbCrLf & _
                     "PaginaInicial = " & Folha & " " & vbCrLf & _
                     "WHERE Empresa_Id = '" & Cliente & "' " & vbCrLf & _
                     "AND Processo_Id = " & Processo & vbCrLf
        End If

        Dim alSQL As New ArrayList
        alSQL.Add(strSQL)

        Return Banco.GravaBanco(alSQL)
    End Function

    'Function Mes(ByVal strMes As String) As String

    '    Select Case strMes
    '        Case "01"
    '            Return "janeiro"
    '        Case "02"
    '            Return "fevereiro"
    '        Case "03"
    '            Return "marco"
    '        Case "04"
    '            Return "abril"
    '        Case "05"
    '            Return "maio"
    '        Case "06"
    '            Return "junho"
    '        Case "07"
    '            Return "julho"
    '        Case "08"
    '            Return "agosto"
    '        Case "09"
    '            Return "setembro"
    '        Case "10"
    '            Return "outubro"
    '        Case "11"
    '            Return "novembro"
    '        Case "12"
    '            Return "dezembro"
    '    End Select

    '    Return ""

    'End Function

    Function Validar()
        Dim ok As Boolean = True

        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de Negócio.")
            ok = False
        End If
        If DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            ok = False
        End If
        If txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o período inicial.")
            ok = False
        End If
        If txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o período final.")
            ok = False
        End If
        If txtLivro.Text = "" Then
            MsgBox(Me.Page, "Informe o número do livro.")
            ok = False
        End If
        If txtFolha.Text = "" Then
            MsgBox(Me.Page, "Informe o número da folha.")
            ok = False
        End If

        Return ok
    End Function

    Private Sub PegarNumeroDoProcesso()
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        Processo = 1

        Sql = "SELECT COALESCE(max(Processo_Id), 0) AS Processo  FROM ProcessoRAISS WHERE Empresa_Id = '" & Empresa(0) & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Processo = Dr("Processo") + 1
        Next
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        txtLivro.Text = ""
        txtFolha.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

#End Region

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("RegistrosDeISS", "GRAVAR") Then
                If Validar() Then
                    Empresa = DdlEmpresa.SelectedValue.Split("-")
                    PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy/MM/dd")
                    PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy/MM/dd")
                    Livro = txtLivro.Text
                    Folha = txtFolha.Text

                    If DdlProcesso.Text = "" Then
                        PegarNumeroDoProcesso()
                    Else
                        Processo = DdlProcesso.SelectedValue
                    End If

                    If GravarProcesso(Processo, Livro, Folha, PeriodoInicial, PeriodoFinal, Empresa(0)) Then
                        CarregarProcesso()
                        DdlProcesso.SelectedValue = Processo
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
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
            Funcoes.Ajuda(Me.Page, "RegistrosDeISS")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class