Imports System.Data
Imports System.Drawing
Imports System.IO
Imports System.Windows.Forms
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Partial Class EstoquesPorOperacoes
    Inherits BasePage

    Dim Sql As String
    Dim Sql2 As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.ProducaoEstoque)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("EstoquesPorOperacoes", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaDepositos()
                ddl.Carregar(ddlExercicio, CarregarDDL.Tabela.Ano, "2015;15;C", False)
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
            End If
        End If
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "FROM Clientes C " & vbCrLf &
              "INNER JOIN ClientesXTipos CT " & vbCrLf &
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              "WHERE CT.Tipo_Id = 050 " & vbCrLf &
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
              " from Usuarios" & vbCrLf &
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Sub CargaDepositos()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlDeposito.Items.Clear()

        Sql = " SELECT  Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, " & vbCrLf &
              "         Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf &
              " FROM    Clientes INNER JOIN" & vbCrLf &
              "         ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id " & vbCrLf &
              "     AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id" & vbCrLf &
              " where   ClientesXTipos.Tipo_Id = 3" & vbCrLf &
              " Order by  Clientes.Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlDeposito.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlDeposito.Items.Insert(0, "")
        DdlDeposito.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        TxtDiaMesInicial.Text = Format(DateSerial(Year(Today), Month(Today), 1), "dd")
        ddlMesInicial.SelectedValue = Format(Today, "MM")
        TxtDiaMesFinal.Text = Format(Today, "dd")
        ddlMesFinal.SelectedValue = Format(Today, "MM")
        ddlExercicio.SelectedValue = Format(Today, "yyyy")

        ucSelecaoProduto.Limpar()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
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

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf &
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf &
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf &
              " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

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

    Function Validar()
        Try
            'If DdlUnidade.Text = "" Then
            '    Mensagem = "Unidade de negócio é obrigatório..."
            '    Return Mensagem
            'End If

            'If DdlEmpresa.Text = "" Then
            '    Mensagem = "Empresa é obrigatório..."
            '    Return Mensagem
            'End If

            If CInt(IIf(TxtDiaMesInicial.Text = "", 1, TxtDiaMesInicial.Text)) > CInt(DateTime.DaysInMonth(CInt(ddlExercicio.Text), CInt(ddlMesInicial.Text))) Then
                Mensagem = "O mês inicial selecionado não tem " & TxtDiaMesInicial.Text & " dias."
                Return Mensagem
            End If

            If CInt(IIf(TxtDiaMesFinal.Text = "", DateTime.DaysInMonth(CInt(ddlExercicio.Text), CInt(ddlMesFinal.Text)), TxtDiaMesFinal.Text)) > CInt(DateTime.DaysInMonth(CInt(ddlExercicio.Text), CInt(ddlMesFinal.Text))) Then
                Mensagem = "O mês final selecionado não tem " & TxtDiaMesFinal.Text & " dias."
                Return Mensagem
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Mensagem = ""

        Return Mensagem
    End Function

    Public Function Bissexto(ByVal intAno As Integer) As Boolean
        Try
            'Verifica se um ano é bissexto
            Bissexto = False

            If intAno Mod 4 = 0 Then
                If intAno Mod 100 = 0 Then
                    If intAno Mod 400 = 0 Then
                        Bissexto = True
                    End If
                Else
                    Bissexto = True
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Function getCampoProduto(Optional ByVal ext As String = "") As String
        Dim campo As String = String.Empty

        If rbPorNome.Checked Then
            campo = String.Format("{0}Nome ", IIf(String.IsNullOrWhiteSpace(ext), "", ext & "."))
        ElseIf rbPorDescricaoMapa.Checked Then
            campo = String.Format("{0}DescricaoMapa ", IIf(String.IsNullOrWhiteSpace(ext), "", ext & "."))
        Else
            campo = String.Format("{0}Nome + ' - ' + {0}DescricaoMapa ", IIf(String.IsNullOrWhiteSpace(ext), "", ext & "."))
        End If

        Return campo
    End Function

    Public Function Fisico(ByVal ExcelDados As Boolean) As DataSet
        Try
            Dim Dia As String
            Dim DiaInicial As String
            Dia = CInt(IIf(TxtDiaMesFinal.Text = "", DateTime.DaysInMonth(CInt(ddlExercicio.Text), CInt(ddlMesFinal.Text)), TxtDiaMesFinal.Text))
            DiaInicial = CInt(IIf(TxtDiaMesInicial.Text = "", 1, TxtDiaMesInicial.Text))
            Dim Mes As String = ddlMesInicial.SelectedValue
            Dim Ano As String = ddlExercicio.SelectedValue
            Dim DataInicial = DiaInicial & " / " & ddlMesInicial.SelectedValue & " / " & ddlExercicio.SelectedValue
            Dim DataFinal = Dia & "/" & ddlMesFinal.SelectedValue & "/" & ddlExercicio.SelectedValue

            Sql = "SELECT sb.Empresa_Id, sb.EndEmpresa_Id," & IIf(CkCDeposito.Checked = False, "sb.Deposito_Id ", "Sb.Empresa_Id as Deposito_Id ") & "," & IIf(CkCDeposito.Checked = False, "sb.EndDeposito_Id ", "Sb.EndEmpresa_Id as EndDeposito_Id ") & ", sb.Produto_Id, sb.Operacao ,sb.SubOperacao,sb.Descricao,sb.EntradaSaida, " & vbCrLf &
                  "       Sum(sb.ValEntradaExercicio) as ValEntradaExercicio, Sum(sb.ValEntradaPeriodo) as ValEntradaPeriodo, Sum(sb.ValSaidaExercicio) as ValSaidaExercicio, Sum(sb.ValSaidaPeriodo) as ValSaidaPeriodo, " & vbCrLf &
                  "       sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa, " & IIf(CkCDeposito.Checked = False, "sb.NomeDeposito ", "Sb.NomeEmpresa as NomeDeposito ") & " ,   " & IIf(CkCDeposito.Checked = False, "sb.CidadeDeposito ", "Sb.CidadeEmpresa as CidadeDeposito ") & " , " & IIf(CkCDeposito.Checked = False, "sb.EstadoDeposito ", "Sb.EstadoEmpresa as EstadoDeposito ") & " , " & IIf(CkCDeposito.Checked = False, "sb.RedDeposito ", "Sb.RedEmpresa as RedDeposito ") &
                  "       ,convert(decimal(18,4),0) as EstoqueAnteriorPeriodo,convert(decimal(18,4),0) as EstoqueAnteriorExercicio, sb.NomeProduto " & vbCrLf &
                  "	 Into #Temp " & vbCrLf &
                  "	 FROM " & vbCrLf &
                  "	      (SELECT E.Empresa_Id,  E.EndEmpresa_Id, " & vbCrLf &
                  "			      E.Deposito_Id, E.EndDeposito_Id, " & vbCrLf &
                  "               E.Produto_Id,  " & vbCrLf &
                  "			      E.Operacao_Id as Operacao, E.Suboperacao_Id as SubOperacao, SO.Descricao, " & vbCrLf &
                  "               SO.EntradaSaida, " & vbCrLf &
                  "               Isnull(SUM(E.Entradas),0) as ValEntradaExercicio, " & vbCrLf &
                  "               isnull(SUM(E.Saidas),0) as ValSaidaExercicio, " & vbCrLf &
                  "			      isnull(SUM(CASE WHEN (E.Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "')  AND SO.EntradaSaida='E' THEN E.Entradas ELSE 0 END),0) AS ValEntradaPeriodo, " & vbCrLf &
                  "               isnull(SUM(CASE WHEN (E.Movimento_Id BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "')  AND SO.EntradaSaida='S' THEN E.saidas   ELSE 0 END),0) AS ValSaidaPeriodo, " & vbCrLf &
                  "			      C.Nome AS NomeEmpresa, " & vbCrLf &
                  "			      C.Cidade AS CidadeEmpresa, " & vbCrLf &
                  "			      C.Estado AS EstadoEmpresa, " & vbCrLf &
                  "			      C.Reduzido AS RedEmpresa, " & vbCrLf &
                  "			      D.Nome AS NomeDeposito, " & vbCrLf &
                  "			      D.Cidade AS CidadeDeposito, " & vbCrLf &
                  "			      D.Estado AS EstadoDeposito, " & vbCrLf &
                  "			      D.Reduzido AS RedDeposito, " & vbCrLf &
                  "               " & getCampoProduto("Produtos") & " AS NomeProduto      " & vbCrLf &
                  "	         FROM Producao E  " & vbCrLf &
                  "	        INNER JOIN SubOperacoes SO " & vbCrLf &
                  "		       ON E.Operacao_Id     = SO.Operacao_Id " & vbCrLf &
                  "	          AND E.SubOperacao_Id  = SO.Suboperacoes_Id " & vbCrLf &
                  "	          AND E.FisicoFiscal_Id = 1 " & vbCrLf &
                  "	          AND SO.EstoqueFisico  = 'S' " & vbCrLf &
                  "	        INNER JOIN Clientes AS C  " & vbCrLf &
                  "		       ON E.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                  "		      AND E.EndEmpresa_Id = C.Endereco_Id  " & vbCrLf &
                  "	        INNER JOIN Clientes AS D " & vbCrLf &
                  "		       ON E.Deposito_Id    = D.Cliente_Id " & vbCrLf &
                  "		      AND E.EndDeposito_Id = D.Endereco_Id " & vbCrLf &
                  "         INNER JOIN Produtos " & vbCrLf &
                  "            ON Produtos.Produto_Id = E.Produto_Id " & vbCrLf &
                  "         INNER JOIN GruposDeEstoques " & vbCrLf &
                  "            ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                  "		      AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            '"		      AND Produtos.Situacao   = 1 " & vbCrLf & _
            Sql &= "         WHERE year(E.Movimento_Id) = " & CDate(DataInicial).Year & vbCrLf &
                  "           AND E.Movimento_Id       <='" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If chkOperacaoCusto.Checked Then
                Sql &= "           and isnull(SO.ApuracaoDeCustos,0) > 0 " & vbCrLf
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "           AND E.Empresa_Id    ='" & campo(0) & "'" & vbCrLf &
                       "	       AND E.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
            End If

            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "      And E.Deposito_Id   ='" & campo(0) & "' " & vbCrLf &
                       "	  AND E.EndDeposito_Id = " & campo(1) & " " & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "E.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            Sql &= "	GROUP BY  E.Empresa_Id, E.EndEmpresa_Id,E.Deposito_Id,EndDeposito_Id, E.Produto_Id, E.Operacao_Id, E.Suboperacao_Id,SO.Descricao,SO.EntradaSaida,  " & vbCrLf &
                   "              C.Nome,C.Cidade, C.Estado, C.Reduzido, D.Nome , D.Cidade, D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & " " & vbCrLf &
                   "	UNION ALL " & vbCrLf &
                   "   SELECT NotasFiscais.Empresa_Id, " & vbCrLf &
                   "          NotasFiscais.EndEmpresa_Id, " & vbCrLf &
                   "          NotasFiscais.Deposito AS Deposito_Id,  " & vbCrLf &
                   "          NotasFiscais.EndDeposito AS EndDeposito_Id, " & vbCrLf &
                   "          nfxi.Produto_Id,  " & vbCrLf &
                   "          NotasFiscais.Operacao, " & vbCrLf &
                   "          NotasFiscais.SubOperacao, " & vbCrLf &
                   "          SO.Descricao, " & vbCrLf &
                   "          SO.EntradaSaida, " & vbCrLf &
                   "          ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValEntradaExercicio,  " & vbCrLf &
                   "          ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'S' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValSaidaExercicio,  " & vbCrLf &
                   "          ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'E' AND (CASE WHEN Produtos.ControlarRomaneio = 'false' THEN NotasFiscais.Movimento ELSE Romaneios.Movimento END BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValEntradaPeriodo,  " & vbCrLf &
                   "          ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'S' AND (CASE WHEN Produtos.ControlarRomaneio = 'false' THEN NotasFiscais.Movimento ELSE Romaneios.Movimento END BETWEEN '" & DataInicial.ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValSaidaPeriodo,  " & vbCrLf &
                   "          C.Nome AS NomeEmpresa,  " & vbCrLf &
                   "          C.Cidade AS CidadeEmpresa,  " & vbCrLf &
                   "          C.Estado AS EstadoEmpresa,  " & vbCrLf &
                   "          C.Reduzido AS RedEmpresa,  " & vbCrLf &
                   "          D.Nome AS NomeDeposito,  " & vbCrLf &
                   "          D.Cidade AS CidadeDeposito,  " & vbCrLf &
                   "          D.Estado AS EstadoDeposito,  " & vbCrLf &
                   "          D.Reduzido AS RedDeposito,  " & vbCrLf &
                   "          " & getCampoProduto("Produtos") & " AS NomeProduto " & vbCrLf &
                   "     From NotasFiscaisXItens AS nfxi " & vbCrLf &
                   "    INNER JOIN NotasFiscais " & vbCrLf &
                   "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                   "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                   "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                   "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                   "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                   "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                   "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                   "    Left JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
                   "       ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id  " & vbCrLf &
                   "      AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id  " & vbCrLf &
                   "      AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf &
                   "      AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf &
                   "      AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id  " & vbCrLf &
                   "      AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf &
                   "      AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
                   "    Left JOIN Romaneios " & vbCrLf &
                   "       ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                   "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                   "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
                   "    INNER JOIN  SubOperacoes AS SO " & vbCrLf &
                   "       ON NotasFiscais.Operacao    = SO.Operacao_Id  " & vbCrLf &
                   "      AND NotasFiscais.SubOperacao = SO.SubOperacoes_Id  " & vbCrLf &
                   "      AND SO.EstoqueFisico      = 'S' " & vbCrLf &
                   "    INNER JOIN Clientes AS C " & vbCrLf &
                   "       ON NotasFiscais.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                   "      AND NotasFiscais.EndEmpresa_Id = C.Endereco_Id  " & vbCrLf &
                   "    Inner JOIN Clientes AS D " & vbCrLf &
                   "       ON NotasFiscais.Deposito    = D.Cliente_Id  " & vbCrLf &
                   "      AND NotasFiscais.EndDeposito = D.Endereco_Id  " & vbCrLf &
                   "    INNER JOIN Produtos " & vbCrLf &
                   "       ON Produtos.Produto_Id = nfxi.Produto_Id  " & vbCrLf &
                   "    INNER JOIN GruposDeEstoques " & vbCrLf &
                   "       ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                   "	  AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            '"      AND Produtos.Situacao   = 1  " & vbCrLf & _
            Sql &= "    WHERE year(NotasFiscais.Movimento) =  " & CDate(DataInicial).Year & vbCrLf &
                   "      AND NotasFiscais.Movimento       <= '" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If Not chkConsNotasComSerieEspecificas.Checked Then
                Sql &= " AND NotasFiscais.Serie_id not in('101','102','103','104') " & vbCrLf
            End If
            'If chkConsNotasComSerieEspecificas.Checked Then
            '    Sql &= " AND NotasFiscais.Serie_id not in('501','502') " & vbCrLf
            'Else
            '    Sql &= " AND NotasFiscais.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
            'End If
            If chkOperacaoCusto.Checked Then
                Sql &= "           and isnull(SO.ApuracaoDeCustos,0) > 0 " & vbCrLf
            End If
            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "      AND NotasFiscais.Deposito    = '" & campo(0) & "' "
                Sql &= "	  AND NotasFiscais.EndDeposito =  " & campo(1) & " "
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "nfxi.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	  AND NotasFiscais.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                       "	  AND NotasFiscais.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
            End If

            Sql &= "      AND NotasFiscais.Situacao        = 1 " & vbCrLf &
                   "      AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf &
                   "    GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito, nfxi.Produto_Id, " & vbCrLf &
                   "             NotasFiscais.Operacao, NotasFiscais.SubOperacao, SO.Descricao, SO.EntradaSaida, C.Nome, C.Cidade, C.Estado, C.Reduzido,D.Nome, D.Cidade , D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & "   " & vbCrLf

            If chkContraPartida.Checked Then
                Sql &= "    UNION ALL " & vbCrLf &
                       "	   SELECT NotasFiscais.Empresa_Id,  " & vbCrLf &
                       "			  NotasFiscais.EndEmpresa_Id," & vbCrLf &
                       "			  NotasFiscais.Destino AS Deposito_Id," & vbCrLf &
                       "			  NotasFiscais.EndDestino AS EndDeposito_Id," & vbCrLf &
                       "			  nfxi.Produto_Id," & vbCrLf &
                       "			  SOD.Operacao_Id," & vbCrLf &
                       "			  SOD.SubOperacoes_Id," & vbCrLf &
                       "			  SOD.Descricao," & vbCrLf &
                       "			  SOD.EntradaSaida," & vbCrLf &
                       "			  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValEntradaExercicio," & vbCrLf &
                       "			  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'S' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValSaidaExercicio," & vbCrLf &
                       "			  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' AND (NotasFiscais.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValEntradaPeriodo," & vbCrLf &
                       "			  ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'S' AND (NotasFiscais.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS ValSaidaPeriodo," & vbCrLf &
                       "			  C.Nome AS NomeEmpresa," & vbCrLf &
                       "			  C.Cidade AS CidadeEmpresa," & vbCrLf &
                       "			  C.Estado AS EstadoEmpresa," & vbCrLf &
                       "			  C.Reduzido AS RedEmpresa," & vbCrLf &
                       "			  D.Nome AS NomeDeposito," & vbCrLf &
                       "			  D.Cidade AS CidadeDeposito," & vbCrLf &
                       "			  D.Estado AS EstadoDeposito," & vbCrLf &
                       "			  D.Reduzido AS RedDeposito," & vbCrLf &
                       "			  " & getCampoProduto("Produtos") & " AS NomeProduto" & vbCrLf &
                       "     From NotasFiscaisXItens AS nfxi " & vbCrLf &
                       "    INNER JOIN NotasFiscais " & vbCrLf &
                       "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                       "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                       "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                       "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                       "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                       "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                       "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                       "    Left JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
                       "       ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id  " & vbCrLf &
                       "      AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf &
                       "      AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
                       "    Left JOIN Romaneios " & vbCrLf &
                       "       ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                       "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                       "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
                       "		INNER JOIN  SubOperacoes AS SO " & vbCrLf &
                       "		   ON NotasFiscais.Operacao    = SO.Operacao_Id  " & vbCrLf &
                       "		  AND NotasFiscais.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                       "		  AND SO.EstoqueFisico      = 'S' " & vbCrLf &
                       "		INNER JOIN  SubOperacoes AS SOD " & vbCrLf &
                       "		   ON SOD.Operacao_Id     = SO.OperacaoDestino" & vbCrLf &
                       "		  AND SOD.SubOperacoes_Id = SO.SubOperacaoDestino" & vbCrLf &
                       "		  AND SOD.EstoqueFisico      = 'S' " & vbCrLf &
                       "		INNER JOIN Clientes AS C " & vbCrLf &
                       "		   ON NotasFiscais.Empresa_Id    = C.Cliente_Id" & vbCrLf &
                       "		  AND NotasFiscais.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                       "		Inner JOIN Clientes AS D " & vbCrLf &
                       "		   ON NotasFiscais.Destino    = D.Cliente_Id  " & vbCrLf &
                       "		  AND NotasFiscais.EndDestino = D.Endereco_Id  " & vbCrLf &
                       "		INNER JOIN Produtos " & vbCrLf &
                       "		   ON Produtos.Produto_Id = nfxi.Produto_Id  " & vbCrLf &
                       "         INNER JOIN GruposDeEstoques " & vbCrLf &
                       "           ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                       "	      AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
                '"		  AND Produtos.Situacao   = 1  " & vbCrLf & _
                Sql &= "		WHERE year(NotasFiscais.Movimento) =   " & CDate(DataInicial).Year & vbCrLf &
                         "         AND NotasFiscais.Movimento       <= '" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf

                If rdAtivo.Checked Then
                    Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
                Else
                    Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
                End If

                If Not chkConsNotasComSerieEspecificas.Checked Then
                    Sql &= " AND NotasFiscais.Serie_id not in('101','102','103','104') " & vbCrLf
                End If
                'If chkConsNotasComSerieEspecificas.Checked Then
                '    Sql &= " AND NotasFiscais.Serie_id not in('501','502') " & vbCrLf
                'Else
                '    Sql &= " AND NotasFiscais.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
                'End If
                If chkOperacaoCusto.Checked Then
                    Sql &= "           and isnull(SOD.ApuracaoDeCustos,0) > 0 " & vbCrLf
                End If
                If DdlDeposito.Text <> "" Then
                    Cliente = DdlDeposito.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "      AND NotasFiscais.Destino    = '" & campo(0) & "'" & vbCrLf &
                           "	  AND NotasFiscais.EndDestino =  " & campo(1) & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "nfxi.Produto_Id", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If DdlEmpresa.Text <> "" Then
                    Cliente = DdlEmpresa.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "	     AND NotasFiscais.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                           "	     AND NotasFiscais.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
                End If

                Sql &= "         AND NotasFiscais.Situacao        = 1 " & vbCrLf &
                       "         AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf &
                       "         AND SO.Deposito = 'S'" & vbCrLf &
                       "		GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Destino,	NotasFiscais.EndDestino, nfxi.Produto_Id," & vbCrLf &
                       "				 SOD.Operacao_id, SOD.SubOperacoes_id, SOD.Descricao, SOD.EntradaSaida, C.Nome, C.Cidade, C.Estado, C.Reduzido,D.Nome, D.Cidade , D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & vbCrLf
            End If

            Sql &= "   ) sb " & vbCrLf

            If CkCDeposito.Checked Then
                Sql &= "	  GROUP BY sb.Empresa_Id, sb.EndEmpresa_Id, sb.Produto_Id, sb.Operacao ,sb.SubOperacao,sb.Descricao,sb.EntradaSaida, " & vbCrLf &
                       "               sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa, sb.NomeProduto;" & vbCrLf
            Else
                Sql &= "      GROUP BY sb.Empresa_Id, sb.EndEmpresa_Id,sb.Deposito_Id,sb.EndDeposito_Id, sb.Produto_Id," & vbCrLf &
                       "               sb.Operacao, sb.SubOperacao, sb.Descricao, sb.EntradaSaida, " & vbCrLf &
                       "               sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa," & vbCrLf &
                       "               sb.NomeDeposito,  sb.CidadeDeposito,  sb.EstadoDeposito,  sb.RedDeposito," & vbCrLf &
                       "               sb.NomeProduto;" & vbCrLf
            End If


            Sql &= "           SELECT sb_saldo.Empresa_Id,sb_saldo.EndEmpresa_Id," & IIf(CkCDeposito.Checked, "sb_saldo.Empresa_Id as Deposito, sb_saldo.EndEmpresa_Id as EndDeposito,", "sb_saldo.Deposito, sb_saldo.EndDeposito,") & vbCrLf &
                    "                  sb_saldo.Produto_Id, Sum(sb_saldo.EstoqueAnteriorPeriodo) as EstoqueAnteriorPeriodo, Sum(sb_saldo.EstoqueAnteriorExercicio) as EstoqueAnteriorExercicio " & vbCrLf &
                    "            Into #SA" & vbCrLf &
                    "             From  " & vbCrLf &
                    "    	          (SELECT ESub.Empresa_Id," & vbCrLf &
                    "                          ESub.EndEmpresa_Id," & vbCrLf &
                    "                          ESub.Deposito_Id as Deposito ," & vbCrLf &
                    "                          ESub.EndDeposito_Id as EndDeposito," & vbCrLf &
                    "                          ESub.Produto_Id, " & vbCrLf &
                    "                          sum(case when year(ESub.Movimento_Id) < " & CDate(DataInicial).Year & " then isnull(ESub.Entradas,0) - isnull(ESub.saidas,0) else 0 end) as EstoqueAnteriorExercicio," & vbCrLf &
                    "                          SUM(isnull(ESub.Entradas,0) - isnull(ESub.saidas,0)) AS  EstoqueAnteriorPeriodo " & vbCrLf &
                    "	                 FROM Producao ESub  " & vbCrLf &
                    "                   INNER JOIN Produtos " & vbCrLf &
                    "                      ON Produtos.Produto_Id = ESub.Produto_Id " & vbCrLf &
                    "	                INNER JOIN SubOperacoes SO " & vbCrLf &
                    "	        	       ON ESub.Operacao_Id     = SO.Operacao_Id " & vbCrLf &
                    "	                  AND ESub.SubOperacao_Id  = SO.Suboperacoes_Id " & vbCrLf &
                    "	                  AND ESub.FisicoFiscal_Id = 1  " & vbCrLf &
                    "	                  AND SO.EstoqueFisico  = 'S' " & vbCrLf &
                    "	                INNER JOIN Clientes AS C  " & vbCrLf &
                    "	        	       ON  ESub.Empresa_Id = C.Cliente_Id " & vbCrLf &
                    "	        	      AND ESub.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                    "	                INNER JOIN Clientes AS D " & vbCrLf &
                    "		               ON ESub.Deposito_Id = D.Cliente_Id " & vbCrLf &
                    "		              AND ESub.EndDeposito_Id = D.Endereco_Id " & vbCrLf &
                    "	                WHERE ESub.Movimento_Id < '" & DateValue(DataInicial).ToSqlDate() & "' " & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "	                  AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "	                  AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	                  AND ESub.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                       "	                  AND ESub.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
            End If
            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	                  And ESub.Deposito_Id   ='" & campo(0) & "' "
                Sql &= "	                  AND ESub.EndDeposito_Id = " & campo(1) & " "
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "ESub.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            Sql &= "	                GROUP BY  ESub.Empresa_Id, ESub.EndEmpresa_Id,ESub.Deposito_Id,ESub.EndDeposito_Id,ESub.Produto_Id       " & vbCrLf &
                   "                    Union All " & vbCrLf &
                   "                   SELECT NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito,    " & vbCrLf &
                   "                          nfxi.Produto_Id,  " & vbCrLf &
                   "                          ISNULL(SUM(CASE When year(NotasFiscais.Movimento) < " & CDate(DataInicial).Year & vbCrLf &
                   "                                            then case when SO.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END Else CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END * -1 END" & vbCrLf &
                   "                                            else 0" & vbCrLf &
                   "                                     end), 0) AS EstoqueAnteriorExercicio, " & vbCrLf &
                   "                          ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END Else CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END * -1 END), 0) AS EstoqueAnteriorPeriodo " & vbCrLf &
                       "     From NotasFiscaisXItens AS nfxi " & vbCrLf &
                       "    INNER JOIN NotasFiscais " & vbCrLf &
                       "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                       "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                       "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                       "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                       "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                       "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                       "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                       "    Left JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
                       "       ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id  " & vbCrLf &
                       "      AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf &
                       "      AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
                       "    Left JOIN Romaneios " & vbCrLf &
                       "       ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                       "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                       "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
                   "                    INNER JOIN SubOperacoes AS SO " & vbCrLf &
                   "                       ON NotasFiscais.Operacao    = SO.Operacao_Id  " & vbCrLf &
                   "                      AND NotasFiscais.SubOperacao = SO.SubOperacoes_Id " & vbCrLf &
                   "                      AND SO.EstoqueFisico         = 'S'" & vbCrLf &
                   "                    INNER JOIN Clientes AS C " & vbCrLf &
                   "                       ON NotasFiscais.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                   "                      AND NotasFiscais.EndEmpresa_Id = C.Endereco_Id  " & vbCrLf &
                   "                     LEFT JOIN Clientes AS D  " & vbCrLf &
                   "                       ON nfxi.Deposito    = D.Cliente_Id  " & vbCrLf &
                   "                      AND nfxi.EndDeposito = D.Endereco_Id " & vbCrLf &
                   "		            INNER JOIN Produtos " & vbCrLf &
                   "		               ON Produtos.Produto_Id = nfxi.Produto_Id  " & vbCrLf &
                   "                    INNER JOIN GruposDeEstoques " & vbCrLf &
                   "                       ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                   "		              AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            '"                      AND Produtos.Situacao   = 1 " & vbCrLf & _
            Sql &= "                    WHERE NotasFiscais.Movimento < '" & DateValue(DataInicial).ToSqlDate() & "'" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If Not chkConsNotasComSerieEspecificas.Checked Then
                Sql &= " AND NotasFiscais.Serie_id not in('101','102','103','104') " & vbCrLf
            End If
            'If chkConsNotasComSerieEspecificas.Checked Then
            '    Sql &= " AND NotasFiscais.Serie_id not in('501','502') " & vbCrLf
            'Else
            '    Sql &= " AND NotasFiscais.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
            'End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "nfxi.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "                      AND NotasFiscais.Deposito    = '" & campo(0) & "' "
                Sql &= "                      AND NotasFiscais.EndDeposito =  " & campo(1) & " "
            End If
            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "                      AND NotasFiscais.Empresa_Id    = '" & campo(0) & "' " & vbCrLf &
                       "	                  AND NotasFiscais.EndEmpresa_Id =  " & campo(1) & " " & vbCrLf
            End If

            Sql &= "                      AND NotasFiscais.Situacao        = 1" & vbCrLf &
                   "                      AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf &
                   "                    GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito,nfxi.Produto_Id  " & vbCrLf

            If chkContraPartida.Checked Then
                Sql &= "                    Union All " & vbCrLf &
                       "                   SELECT NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.Destino, NotasFiscais.EndDestino,    " & vbCrLf &
                       "                          nfxi.Produto_Id,  " & vbCrLf &
                       "                          ISNULL(SUM(CASE When year(NotasFiscais.Movimento) < " & CDate(DataInicial).Year & vbCrLf &
                       "                                            then case when SOD.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END Else CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END * -1 END" & vbCrLf &
                       "                                            else 0" & vbCrLf &
                       "                                     end), 0) AS EstoqueAnteriorExercicio, " & vbCrLf &
                       "                          ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END Else CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END * -1 END), 0) AS EstoqueAnteriorPeriodo " & vbCrLf &
                       "     From NotasFiscaisXItens AS nfxi " & vbCrLf &
                       "    INNER JOIN NotasFiscais " & vbCrLf &
                       "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                       "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                       "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                       "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                       "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                       "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                       "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                       "    Left JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
                       "       ON NotasFiscais.Empresa_Id      = nfxr.Empresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndEmpresa_Id   = nfxr.EndEmpresa_Id  " & vbCrLf &
                       "      AND NotasFiscais.Cliente_Id      = nfxr.Cliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EndCliente_Id   = nfxr.EndCliente_Id  " & vbCrLf &
                       "      AND NotasFiscais.EntradaSaida_Id = nfxr.EntradaSaida_Id  " & vbCrLf &
                       "      AND NotasFiscais.Serie_Id        = nfxr.Serie_Id  " & vbCrLf &
                       "      AND NotasFiscais.Nota_Id         = nfxr.Nota_Id  " & vbCrLf &
                       "    Left JOIN Romaneios " & vbCrLf &
                       "       ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                       "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                       "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id " & vbCrLf &
                       "                    INNER JOIN SubOperacoes AS SO " & vbCrLf &
                       "                       ON NotasFiscais.Operacao    = SO.Operacao_Id  " & vbCrLf &
                       "                      AND NotasFiscais.SubOperacao = SO.SubOperacoes_Id " & vbCrLf &
                       "                      AND SO.EstoqueFisico      = 'S'" & vbCrLf &
                       "                      AND SO.Deposito           = 'S'" & vbCrLf &
                       "                    INNER JOIN SubOperacoes AS SOD " & vbCrLf &
                       "                       ON SOD.Operacao_Id     = SO.OperacaoDestino  " & vbCrLf &
                       "                      AND SOD.SubOperacoes_Id = SO.SubOperacaoDestino " & vbCrLf &
                       "                      AND SO.EstoqueFisico      = 'S'" & vbCrLf &
                       "                    INNER JOIN Clientes AS C " & vbCrLf &
                       "                       ON NotasFiscais.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                       "                      AND NotasFiscais.EndEmpresa_Id = C.Endereco_Id  " & vbCrLf &
                       "                     LEFT JOIN Clientes AS D  " & vbCrLf &
                       "                       ON NotasFiscais.Destino    = D.Cliente_Id  " & vbCrLf &
                       "                      AND NotasFiscais.EndDestino = D.Endereco_Id " & vbCrLf &
                       "		            INNER JOIN Produtos " & vbCrLf &
                       "		               ON Produtos.Produto_Id = nfxi.Produto_Id " & vbCrLf &
                       "                    INNER JOIN GruposDeEstoques " & vbCrLf &
                       "                       ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                       "		              AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
                '"                      AND Produtos.Situacao   = 1 " & vbCrLf & _
                Sql &= "                    WHERE NotasFiscais.Situacao        = 1" & vbCrLf &
                       "                      AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf

                If rdAtivo.Checked Then
                    Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
                Else
                    Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
                End If

                If Not chkConsNotasComSerieEspecificas.Checked Then
                    Sql &= "  NotasFiscais.Serie_id not in('101','102','103','104') " & vbCrLf
                End If
                'If chkConsNotasComSerieEspecificas.Checked Then
                '    Sql &= "  NotasFiscais.Serie_id not in('501','502') " & vbCrLf
                'Else
                '    Sql &= "  NotasFiscais.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
                'End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "nfxi.Produto_Id", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If DdlDeposito.Text <> "" Then
                    Cliente = DdlDeposito.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "                      AND NotasFiscais.Destino    = '" & campo(0) & "' "
                    Sql &= "                      AND NotasFiscais.EndDestino =  " & campo(1) & " "
                End If
                If DdlEmpresa.Text <> "" Then
                    Cliente = DdlEmpresa.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "                      AND NotasFiscais.Empresa_Id    = '" & campo(0) & "' " & vbCrLf &
                           "	                  AND NotasFiscais.EndEmpresa_Id =  " & campo(1) & " " & vbCrLf
                End If

                Sql &= "                      AND NotasFiscais.Movimento     < '" & DateValue(DataInicial).ToSqlDate() & "'" & vbCrLf &
                      "                      AND SO.Deposito             = 'S'" & vbCrLf &
                      "                      AND year(NotasFiscais.Movimento) >= 2012" & vbCrLf &
                      "                    GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Destino, NotasFiscais.EndDestino, nfxi.Produto_Id" & vbCrLf
            End If

            Sql &= "                  ) sb_saldo " & vbCrLf

            If CkCDeposito.Checked Then
                Sql &= "       	group by sb_saldo.Empresa_Id,sb_saldo.EndEmpresa_Id,sb_saldo.Produto_Id " & vbCrLf
            Else
                Sql &= "       	group by sb_saldo.Empresa_Id, sb_saldo.EndEmpresa_Id ,sb_saldo.Deposito, sb_saldo.EndDeposito, sb_saldo.Produto_Id " & vbCrLf
            End If

            Sql &= " Update #temp set" & vbCrLf &
                   "     EstoqueAnteriorPeriodo   = #SA.EstoqueAnteriorPeriodo" & vbCrLf &
                   "    ,EstoqueAnteriorExercicio = #SA.EstoqueAnteriorExercicio" & vbCrLf &
                   "  from #temp " & vbCrLf &
                   " Inner Join #Sa" & vbCrLf &
                   "    ON #temp.Produto_Id     = #Sa.Produto_Id " & vbCrLf &
                   "   AND #temp.Empresa_Id     = #Sa.Empresa_Id " & vbCrLf &
                   "   AND #temp.EndEmpresa_Id  = #Sa.EndEmpresa_Id" & vbCrLf &
                   "   AND #temp.Deposito_Id    = #Sa.Deposito " & vbCrLf &
                   "   AND #temp.EndDeposito_Id = #Sa.EndDeposito " & vbCrLf

            Sql &= " Insert Into #Temp(Empresa_Id,  EndEmpresa_Id, " & vbCrLf &
                   " 				   Deposito_Id, EndDeposito_Id," & vbCrLf &
                   " 				   Produto_Id, NomeProduto," & vbCrLf &
                   "                   Operacao , SubOperacao, Descricao," & vbCrLf &
                   " 				   EntradaSaida, ValEntradaExercicio, ValEntradaPeriodo, ValSaidaExercicio, ValSaidaPeriodo," & vbCrLf &
                   " 				   NomeEmpresa,  CidadeEmpresa, EstadoEmpresa, RedEmpresa," & vbCrLf &
                   " 				   NomeDeposito,	CidadeDeposito, EstadoDeposito, RedDeposito," & vbCrLf &
                   " 				   EstoqueAnteriorPeriodo, EstoqueAnteriorExercicio)" & vbCrLf &
                   " (" & vbCrLf &
                   " Select #SA.Empresa_Id, #Sa.EndEmpresa_Id," & vbCrLf &
                   "        #Sa.Deposito,   #Sa.EndDeposito," & vbCrLf &
                   "        #Sa.Produto_Id, " & getCampoProduto("Prd") & "," & vbCrLf &
                   "        0,0,'Sem Movimentacao No Periodo'," & vbCrLf &
                   "        0,0,0,0,0," & vbCrLf &
                   "        Emp.Nome, Emp.Cidade, Emp.Estado, Emp.Reduzido," & vbCrLf &
                   "        Dep.Nome, Dep.Cidade, Dep.Estado, Dep.Reduzido," & vbCrLf &
                   "        #SA.EstoqueAnteriorPeriodo, #SA.EstoqueAnteriorExercicio" & vbCrLf &
                   "   from #SA" & vbCrLf &
                   "  Inner Join Clientes as Emp" & vbCrLf &
                   "     on Emp.Cliente_id  = #SA.Empresa_Id" & vbCrLf &
                   "    and Emp.Endereco_id = #Sa.EndEmpresa_Id" & vbCrLf &
                   "  Inner Join Clientes as Dep" & vbCrLf &
                   "     on Dep.Cliente_id  = #SA.Deposito" & vbCrLf &
                   "    and Dep.Endereco_id = #Sa.EndDeposito" & vbCrLf &
                   "  Inner Join Produtos Prd" & vbCrLf &
                   "     on Prd.Produto_id = #Sa.Produto_id" & vbCrLf &
                   "  Where #SA.EstoqueAnteriorExercicio > 0" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "    and Prd.Situacao in(1) " & vbCrLf
            Else
                Sql &= "    and Prd.Situacao not in(1) " & vbCrLf
            End If

            Sql &= "    and not exists(Select 1" & vbCrLf &
                   "                     from #temp" & vbCrLf &
                   "                    Where #temp.Produto_Id     = #Sa.Produto_Id" & vbCrLf &
                   " 					  and #temp.Empresa_Id     = #Sa.Empresa_Id" & vbCrLf &
                   " 					  and #temp.EndEmpresa_Id  = #Sa.EndEmpresa_Id" & vbCrLf &
                   " 					  and #temp.Deposito_Id    = #Sa.Deposito" & vbCrLf &
                   " 					  and #temp.EndDeposito_Id = #Sa.EndDeposito" & vbCrLf &
                   "                   )" & vbCrLf &
                   " )" & vbCrLf

            Sql2 = " select Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                            IIf(CkCDeposito.Checked = False, "Deposito_Id, EndDeposito_Id, ", "Empresa_Id as Deposito_Id, EndEmpresa_Id as EndDeposito_Id,") & vbCrLf &
                   "         Produto_Id," & vbCrLf &
                   "         Operacao, SubOperacao, Descricao," & vbCrLf &
                   "         EntradaSaida," & vbCrLf &
                   "         ValEntradaExercicio," & vbCrLf &
                   "         ValEntradaPeriodo," & vbCrLf &
                   "         ValSaidaExercicio," & vbCrLf &
                   "         ValSaidaPeriodo," & vbCrLf &
                   "         NomeEmpresa," & vbCrLf &
                   "         CidadeEmpresa," & vbCrLf &
                   "         EstadoEmpresa," & vbCrLf &
                   "         RedEmpresa," & vbCrLf &
                   IIf(CkCDeposito.Checked, "NomeEmpresa as NomeDeposito, CidadeEmpresa as CidadeDeposito, EstadoEmpresa as EstadoDeposito, RedEmpresa as RedDeposito,", "NomeDeposito, CidadeDeposito, EstadoDeposito, RedDeposito, ") & vbCrLf &
                   "         0 as EstoqueAnterior, " & vbCrLf &
                   "         NomeProduto " & vbCrLf &
                   "  into #temp2 " & vbCrLf &
                   "    FROM #temp " & vbCrLf &
                   " Order by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao, Suboperacao  "

            Sql2 &= " select Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                    "        case" & vbCrLf &
                    "            when Deposito_Id = Empresa_Id" & vbCrLf &
                    "                then '00000000000001'" & vbCrLf &
                    "                else Deposito_Id" & vbCrLf &
                    "        end as Deposito_Id, EndDeposito_Id, " & vbCrLf &
                   "         Produto_Id," & vbCrLf &
                   "         Operacao, SubOperacao, Descricao," & vbCrLf &
                   "         EntradaSaida," & vbCrLf &
                   "         ValEntradaExercicio," & vbCrLf &
                   "         ValEntradaPeriodo," & vbCrLf &
                   "         ValSaidaExercicio," & vbCrLf &
                   "         ValSaidaPeriodo," & vbCrLf &
                   "         NomeEmpresa," & vbCrLf &
                   "         CidadeEmpresa," & vbCrLf &
                   "         EstadoEmpresa," & vbCrLf &
                   "         RedEmpresa," & vbCrLf &
                   "         NomeDeposito, CidadeDeposito, EstadoDeposito, RedDeposito, " & vbCrLf &
                   "         EstoqueAnterior, " & vbCrLf &
                   "         NomeProduto " & vbCrLf &
                   "    FROM #temp2 " & vbCrLf &
                   " Order by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, Operacao, Suboperacao  "

            If ExcelDados Then

                Sql &= "SELECT Empresa_Id, EndEmpresa_Id, NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa, Deposito_Id, EndDeposito_Id, NomeDeposito, " & vbCrLf &
                       "CidadeDeposito, EstadoDeposito, RedDeposito, Produto_Id, NomeProduto, EntradaSaida, Operacao, SubOperacao, Descricao, " & vbCrLf &
                       "SUM(EstoqueAnteriorPeriodo)    As TotalEstoqueAnteriorPeriodo, SUM(EstoqueAnteriorExercicio)   As TotalEstoqueAnteriorExercicio, " & vbCrLf &
                       "SUM(ValEntradaPeriodo)          As TotalValEntradaPeriodo, SUM(ValEntradaExercicio)        As TotalValEntradaExercicio, " & vbCrLf &
                       "SUM(ValSaidaPeriodo)            As TotalValSaidaPeriodo, SUM(ValSaidaExercicio)          As TotalValSaidaExercicio, " & vbCrLf &
                       "SUM(EstoqueAnteriorExercicio) + (SUM(ValEntradaExercicio) - SUM(ValSaidaExercicio))       As TotalEstoque " & vbCrLf &
                       "into #VerExcel " & vbCrLf &
                       "FROM #temp" & vbCrLf &
                       "Group BY  Empresa_Id, EndEmpresa_Id, NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa, " & vbCrLf &
                       "Deposito_Id, EndDeposito_Id,NomeDeposito, CidadeDeposito, EstadoDeposito, RedDeposito, " & vbCrLf &
                       "Produto_Id, NomeProduto, EntradaSaida,Operacao, SubOperacao, Descricao" & vbCrLf &
                       "ORDER BY Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, EntradaSaida, Operacao, SubOperacao;" & vbCrLf &
                       "Select * from #VerExcel " & vbCrLf
            End If

            DS = Banco.ConsultaDataSet(Sql & Sql2, "EstoquePorOperacao")

            Sql2 = " select Empresa_Id,   EndEmpresa_Id, " & vbCrLf &
                    "        case" & vbCrLf &
                    "            when Deposito_Id = Empresa_Id" & vbCrLf &
                    "                then '00000000000001'" & vbCrLf &
                    "                else Deposito_Id" & vbCrLf &
                    "        end as Deposito_Id, EndDeposito_Id, Produto_Id, " & vbCrLf &
                    "        EstoqueAnteriorPeriodo, " & vbCrLf &
                    "        EstoqueAnteriorExercicio, " & vbCrLf &
                    "        SUM(ValEntradaPeriodo)   as TotalEntradasPeriodo, " & vbCrLf &
                    "        SUM(ValEntradaExercicio) as TotalEntradasExercicio, " & vbCrLf &
                    "        SUM(ValSaidaPeriodo)     as TotalSaidasPeriodo, " & vbCrLf &
                    "        SUM(ValSaidaExercicio)   as TotalSaidasExercicio " & vbCrLf &
                    "   FROM #temp " & vbCrLf &
                    "  group by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, EstoqueAnteriorPeriodo, EstoqueAnteriorExercicio " & vbCrLf &
                    "  Order by Empresa_Id, EndEmpresa_Id,Deposito_Id, EndDeposito_Id, Produto_Id " & vbCrLf
            DS.Merge(Banco.ConsultaDataSet(Sql & Sql2, "TotalEstoquePorOperacao"))

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return DS
    End Function

    Public Function Fiscal(ByVal ExcelDados As Boolean) As DataSet
        Try
            Dim Dia As String
            Dim DiaInicial As String
            Dia = CInt(IIf(TxtDiaMesFinal.Text = "", DateTime.DaysInMonth(CInt(ddlExercicio.Text), CInt(ddlMesFinal.Text)), TxtDiaMesFinal.Text))
            DiaInicial = CInt(IIf(TxtDiaMesInicial.Text = "", 1, TxtDiaMesInicial.Text))
            Dim Mes As String = ddlMesInicial.SelectedValue
            Dim Ano As String = ddlExercicio.SelectedValue
            Dim DataInicial = DiaInicial & " / " & ddlMesInicial.SelectedValue & " / " & ddlExercicio.SelectedValue
            Dim DataFinal = Dia & "/" & ddlMesFinal.SelectedValue & "/" & ddlExercicio.SelectedValue


            Sql = "SELECT sb.Empresa_Id, sb.EndEmpresa_Id," & IIf(CkCDeposito.Checked = False, "sb.Deposito_Id ", "Sb.Empresa_Id as Deposito_Id ") & "," & IIf(CkCDeposito.Checked = False, "sb.EndDeposito_Id ", "Sb.EndEmpresa_Id as EndDeposito_Id ") & ", sb.Produto_Id, sb.Operacao ,sb.SubOperacao,sb.Descricao,sb.EntradaSaida, " & vbCrLf &
                         "       Sum(sb.ValEntradaExercicio) as ValEntradaExercicio,Sum(sb.ValEntradaPeriodo) as ValEntradaPeriodo, Sum(sb.ValSaidaExercicio) as ValSaidaExercicio, Sum(sb.ValSaidaPeriodo) as ValSaidaPeriodo, " & vbCrLf &
                         "       sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa, " & IIf(CkCDeposito.Checked = False, "sb.NomeDeposito ", "Sb.NomeEmpresa as NomeDeposito ") & " ,   " & IIf(CkCDeposito.Checked = False, "sb.CidadeDeposito ", "Sb.CidadeEmpresa as CidadeDeposito ") & " , " & vbCrLf &
                         IIf(CkCDeposito.Checked = False, "sb.EstadoDeposito ", "Sb.EstadoEmpresa as EstadoDeposito ") & " , " & IIf(CkCDeposito.Checked = False, "sb.RedDeposito ", "Sb.RedEmpresa as RedDeposito ") & ", convert(decimal(18,4),0) as EstoqueAnteriorPeriodo,convert(decimal(18,4),0) as EstoqueAnteriorExercicio, sb.NomeProduto " & vbCrLf &
                         "	Into #Temp " & vbCrLf &
                         "	FROM ( " & vbCrLf &
                         "	      SELECT E.Empresa_Id, " & vbCrLf &
                         "			     E.EndEmpresa_Id, " & vbCrLf &
                         "			     E.Deposito_Id, " & vbCrLf &
                         "			     E.EndDeposito_Id, " & vbCrLf &
                         "               E.Produto_Id,  " & vbCrLf &
                         "			     E.Operacao_Id as Operacao, " & vbCrLf &
                         "               E.Suboperacao_Id as SubOperacao, " & vbCrLf &
                         "               SO.Descricao, " & vbCrLf &
                         "               SO.EntradaSaida, " & vbCrLf &
                         "               Isnull(SUM(E.Entradas),0) as ValEntradaExercicio, " & vbCrLf &
                         "               Isnull(SUM(E.Saidas),0) as ValSaidaExercicio, " & vbCrLf &
                         "			     Isnull(SUM(CASE WHEN (E.Movimento_Id BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') AND SO.EntradaSaida='E' THEN E.Entradas ELSE 0 END),0) AS ValEntradaPeriodo, " & vbCrLf &
                         "               Isnull(SUM(CASE WHEN (E.Movimento_Id BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') AND SO.EntradaSaida='S' THEN E.saidas   ELSE 0 END),0) AS ValSaidaPeriodo, " & vbCrLf &
                         "			     C.Nome AS NomeEmpresa, " & vbCrLf &
                         "			     C.Cidade AS CidadeEmpresa, " & vbCrLf &
                         "			     C.Estado AS EstadoEmpresa, " & vbCrLf &
                         "			     C.Reduzido AS RedEmpresa, " & vbCrLf &
                         "			     D.Nome AS NomeDeposito, " & vbCrLf &
                         "			     D.Cidade AS CidadeDeposito, " & vbCrLf &
                         "			     D.Estado AS EstadoDeposito, " & vbCrLf &
                         "			     D.Reduzido AS RedDeposito, " & vbCrLf &
                         "               " & getCampoProduto("Produtos") & " AS NomeProduto " & vbCrLf &
                         "	        FROM Producao E  " & vbCrLf &
                         "	       INNER JOIN SubOperacoes SO " & vbCrLf &
                         "		      ON E.Operacao_Id     = SO.Operacao_Id " & vbCrLf &
                         "	         AND E.SubOperacao_Id  = SO.Suboperacoes_Id " & vbCrLf &
                         "	         AND E.FisicoFiscal_Id = 2 " & vbCrLf &
                         "	         AND SO.EstoqueFiscal  = 'S' " & vbCrLf &
                         "	       INNER JOIN Clientes AS C  " & vbCrLf &
                         "		      ON E.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                         "		     AND E.EndEmpresa_Id = C.Endereco_Id  " & vbCrLf &
                         "   	   INNER JOIN Clientes AS D " & vbCrLf &
                         "		      ON E.Deposito_Id    = D.Cliente_Id " & vbCrLf &
                         "		     AND E.EndDeposito_Id = D.Endereco_Id " & vbCrLf &
                         "         INNER JOIN Produtos " & vbCrLf &
                         "            ON Produtos.Produto_Id = E.Produto_Id " & vbCrLf &
                         "         INNER JOIN GruposDeEstoques " & vbCrLf &
                         "            ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                         "		     AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            '"		     AND Produtos.Situacao   = 1 " & vbCrLf & _
            Sql &= "	       WHERE year(E.Movimento_Id) = " & CDate(DataInicial).Year & vbCrLf &
                         "           and E.Movimento_Id      <= '" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If chkOperacaoCusto.Checked Then
                Sql &= "           and isnull(SO.ApuracaoDeCustos,0) > 0 " & vbCrLf
            End If
            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "          AND E.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                       "          AND E.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
            End If
            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	         AND E.Deposito_Id    ='" & campo(0) & "'" & vbCrLf &
                       "	         AND E.EndDeposito_Id = " & campo(1) & " " & vbCrLf
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "E.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            Sql &= "           GROUP BY E.Empresa_Id, E.EndEmpresa_Id,E.Deposito_Id,EndDeposito_Id, E.Produto_Id, E.Operacao_Id, E.Suboperacao_Id,SO.Descricao,SO.EntradaSaida,  " & vbCrLf &
                   "                    C.Nome,C.Cidade, C.Estado, C.Reduzido, D.Nome , D.Cidade, D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & vbCrLf &
                   "	       UNION ALL " & vbCrLf &
                   "	      SELECT NF.Empresa_Id, " & vbCrLf &
                   "			     NF.EndEmpresa_Id, " & vbCrLf &
                   "			     NFxI.Deposito as Deposito_Id, " & vbCrLf &
                   "			     NFxI.EndDeposito as EndDeposito_Id, " & vbCrLf &
                   "			     NFxI.Produto_Id, " & vbCrLf &
                   "                 NFxI.Operacao, " & vbCrLf &
                   "                 NFxI.SubOperacao, " & vbCrLf &
                   "                 SO.Descricao, " & vbCrLf &
                   "                 SO.EntradaSaida, " & vbCrLf &
                   "                 isnull(SUM(CASE WHEN SO.EntradaSaida = 'E' THEN NFxI.QuantidadeFiscal END),0) AS ValEntradaExercicio, " & vbCrLf &
                   "                 isnull(SUM(CASE WHEN SO.EntradaSaida = 'S' THEN NFxI.QuantidadeFiscal END),0)  AS ValSaidaExercicio, " & vbCrLf &
                   "		         isnull(SUM(CASE WHEN SO.EntradaSaida = 'E' and (NF.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN NFxI.QuantidadeFiscal END),0) AS ValEntradaPeriodo, " & vbCrLf &
                   "			     isnull(SUM(CASE WHEN SO.EntradaSaida = 'S' and (NF.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN NFxI.QuantidadeFiscal END),0) AS ValSaidaPeriodo,  " & vbCrLf &
                   "			     C.Nome as NomeEmpresa, " & vbCrLf &
                   "			     C.Cidade as CidadeEmpresa,  " & vbCrLf &
                   "			     C.Estado as EstadoEmpresa, " & vbCrLf &
                   "			     C.Reduzido as RedEmpresa, " & vbCrLf &
                   "			     D.Nome as NomeDeposito, " & vbCrLf &
                   "			     D.Cidade as CidadeDeposito, " & vbCrLf &
                   "			     D.Estado as EstadoDeposito, " & vbCrLf &
                   "		  	     D.Reduzido as RedDeposito, " & vbCrLf &
                   "                 " & getCampoProduto("Produtos") & " AS NomeProduto      " & vbCrLf &
                   "	        FROM NotasFiscais NF  " & vbCrLf &
                   " 	       INNER JOIN NotasFiscaisXItens NFxI  " & vbCrLf &
                   "	          ON  NF.Empresa_Id     = NFxI.Empresa_Id  " & vbCrLf &
                   "	         AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
                   "	         AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
                   "	         AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
                   "	         AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
                   "	         AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
                   "	         AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
                   "	       INNER JOIN SubOperacoes SO " & vbCrLf &
                   "	          ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf &
                   "	         AND NFxI.SubOperacao = SO.Suboperacoes_Id " & vbCrLf &
                   "	         AND SO.EstoqueFiscal = 'S' " & vbCrLf &
                   "	       INNER JOIN  Clientes C " & vbCrLf &
                   "	          ON NFxI.Empresa_Id = C.Cliente_Id  " & vbCrLf &
                   "	         AND NFxI.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                   "	        LEFT JOIN  Clientes D " & vbCrLf &
                   "	          ON NFxI.Deposito = D.Cliente_Id  " & vbCrLf &
                   "	         AND NFxI.EndDeposito = D.Endereco_Id " & vbCrLf &
                   "           INNER JOIN  Produtos" & vbCrLf &
                   "              ON Produtos.Produto_Id = NFxI.Produto_Id " & vbCrLf &
                   "         INNER JOIN GruposDeEstoques " & vbCrLf &
                   "            ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                   "		     AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            ' "	         AND Produtos.Situacao   = 1 " & vbCrLf & _
            Sql &= "	       WHERE NF.Situacao        = 1  " & vbCrLf &
                   "	         AND NF.TipoDeDocumento = 1  "

            If rdAtivo.Checked Then
                Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
            End If

            If Not chkConsNotasComSerieEspecificas.Checked Then
                Sql &= " AND NF.Serie_id not in('101','102','103','104') " & vbCrLf
            End If
            'If chkConsNotasComSerieEspecificas.Checked Then
            '    Sql &= "  NF.Serie_id not in('501','502') " & vbCrLf
            'Else
            '    Sql &= "  NF.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
            'End If
            If chkOperacaoCusto.Checked Then
                Sql &= "           and isnull(SO.ApuracaoDeCustos,0) > 0 " & vbCrLf
            End If

            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	         AND NFxI.Deposito   = '" & campo(0) & "' "
                Sql &= "	         AND NFxI.EndDeposito =  " & campo(1) & " "
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NFxI.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	         AND NF.Empresa_Id    = '" & campo(0) & "' " & vbCrLf &
                       "	         AND NF.EndEmpresa_Id =  " & campo(1) & " " & vbCrLf
            End If

            Sql &= "           AND year(NF.Movimento) = " & CDate(DataInicial).Year & vbCrLf &
                   "           AND NF.Movimento      <= '" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf &
                   "	       GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NFxI.Deposito,	NFxI.EndDeposito, NFxI.Produto_Id, NFxI.Operacao ,  NFxI.SubOperacao, SO.Descricao,SO.EntradaSaida, " & vbCrLf &
                   "			        C.Nome, C.Cidade, C.Estado, C.Reduzido,D.Nome, D.Cidade , D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & vbCrLf

            If chkContraPartida.Checked Then
                Sql &= "	       UNION ALL " & vbCrLf &
                       "	      SELECT NF.Empresa_Id," & vbCrLf &
                       "			     NF.EndEmpresa_Id," & vbCrLf &
                       "			     NF.Destino as Deposito_Id, " & vbCrLf &
                       "			     NF.EndDestino as EndDeposito_Id, " & vbCrLf &
                       "			     NFxI.Produto_Id, " & vbCrLf &
                       "                 SOD.Operacao_id, " & vbCrLf &
                       "                 SOD.SubOperacoes_Id, " & vbCrLf &
                       "                 SOD.Descricao, " & vbCrLf &
                       "                 SOD.EntradaSaida, " & vbCrLf &
                       "                 isnull(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN NFxI.QuantidadeFiscal END),0) AS ValEntradaExercicio, " & vbCrLf &
                       "                 isnull(SUM(CASE WHEN SOD.EntradaSaida = 'S' THEN NFxI.QuantidadeFiscal END),0)  AS ValSaidaExercicio, " & vbCrLf &
                       "		         isnull(SUM(CASE WHEN SOD.EntradaSaida = 'E' and (NF.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN NFxI.QuantidadeFiscal END),0) AS ValEntradaPeriodo, " & vbCrLf &
                       "			     isnull(SUM(CASE WHEN SOD.EntradaSaida = 'S' and (NF.Movimento BETWEEN '" & DateValue(DataInicial).ToSqlDate() & "' AND '" & DateValue(DataFinal).ToSqlDate() & "') THEN NFxI.QuantidadeFiscal END),0) AS ValSaidaPeriodo," & vbCrLf &
                       "			     C.Nome as NomeEmpresa, " & vbCrLf &
                       "			     C.Cidade as CidadeEmpresa,  " & vbCrLf &
                       "			     C.Estado as EstadoEmpresa, " & vbCrLf &
                       "			     C.Reduzido as RedEmpresa, " & vbCrLf &
                       "			     D.Nome as NomeDeposito, " & vbCrLf &
                       "			     D.Cidade as CidadeDeposito, " & vbCrLf &
                       "			     D.Estado as EstadoDeposito, " & vbCrLf &
                       "		  	     D.Reduzido as RedDeposito, " & vbCrLf &
                       "                 " & getCampoProduto("Produtos") & " AS NomeProduto" & vbCrLf &
                       "	        FROM NotasFiscais NF  " & vbCrLf &
                       " 	       INNER JOIN NotasFiscaisXItens NFxI  " & vbCrLf &
                       "	          ON NF.Empresa_Id     = NFxI.Empresa_Id  " & vbCrLf &
                       "	         AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf &
                       "	         AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf &
                       "	         AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf &
                       "	         AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf &
                       "	         AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf &
                       "	         AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf &
                       "	       INNER JOIN SubOperacoes SO " & vbCrLf &
                       "	          ON NFxI.Operacao    = SO.Operacao_Id " & vbCrLf &
                       "	         AND NFxI.SubOperacao = SO.Suboperacoes_Id " & vbCrLf &
                       "	         AND SO.EstoqueFiscal = 'S' " & vbCrLf &
                       "         INNER JOIN SubOperacoes SOD" & vbCrLf &
                       "	          ON SOD.Operacao_Id     = SO.OperacaoDestino " & vbCrLf &
                       "	         AND SOD.Suboperacoes_Id = SO.SuboperacaoDestino " & vbCrLf &
                       "           AND SOD.EstoqueFiscal   = 'S'" & vbCrLf &
                       "	       INNER JOIN  Clientes C " & vbCrLf &
                       "	          ON NFxI.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                       "	         AND NFxI.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                       "	        LEFT JOIN  Clientes D " & vbCrLf &
                       "	          ON NF.Destino    = D.Cliente_Id  " & vbCrLf &
                       "	         AND NF.EndDestino = D.Endereco_Id " & vbCrLf &
                       "           INNER JOIN  Produtos" & vbCrLf &
                       "              ON Produtos.Produto_Id = NFxI.Produto_Id " & vbCrLf &
                       "         INNER JOIN GruposDeEstoques " & vbCrLf &
                       "            ON GruposDeEstoques.Grupo_Id = Produtos.Grupo " & vbCrLf &
                       "		     AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
                '"	         AND Produtos.Situacao   = 1 " & vbCrLf & _
                Sql &= "	       WHERE NF.Situacao        = 1 " & vbCrLf &
                       "	         AND NF.TipoDeDocumento = 1  "

                If rdAtivo.Checked Then
                    Sql &= "           AND Produtos.Situacao in(1) " & vbCrLf
                Else
                    Sql &= "           AND Produtos.Situacao not in(1) " & vbCrLf
                End If

                If Not chkConsNotasComSerieEspecificas.Checked Then
                    Sql &= "  AND NF.Serie_id not in('101','102','103','104') " & vbCrLf
                End If
                'If chkConsNotasComSerieEspecificas.Checked Then
                '    Sql &= "  NF.Serie_id not in('501','502') " & vbCrLf
                'Else
                '    Sql &= "  NF.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
                'End If
                If chkOperacaoCusto.Checked Then
                    Sql &= "           and isnull(SOD.ApuracaoDeCustos,0) > 0 " & vbCrLf
                End If

                If DdlDeposito.Text <> "" Then
                    Cliente = DdlDeposito.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "	         AND NF.Destino    = '" & campo(0) & "' "
                    Sql &= "	         AND NF.EndDestino =  " & campo(1) & " "
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NFxI.Produto_Id", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If DdlEmpresa.Text <> "" Then
                    Cliente = DdlEmpresa.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "	         AND NF.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                           "	         AND NF.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
                End If

                Sql &= "             AND year(NF.Movimento) =  " & CDate(DataInicial).Year & vbCrLf &
                       "             And NF.Movimento       <= '" & DateValue(DataFinal).ToSqlDate() & "'" & vbCrLf &
                       "             AND SO.Deposito = 'S'" & vbCrLf &
                       "	       GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Destino, NF.EndDestino, NFxI.Produto_Id, SOD.Operacao_Id, SOD.SubOperacoes_Id, SOD.Descricao, SOD.EntradaSaida" & vbCrLf &
                       "			        ,C.Nome, C.Cidade, C.Estado, C.Reduzido,D.Nome, D.Cidade , D.Estado , D.Reduzido, " & getCampoProduto("Produtos") & vbCrLf
            End If

            Sql &= "         ) sb " & vbCrLf
            If CkCDeposito.Checked Then
                Sql &= "	  GROUP BY sb.Empresa_Id, sb.EndEmpresa_Id, sb.Produto_Id, sb.Operacao ,sb.SubOperacao,sb.Descricao,sb.EntradaSaida, " & vbCrLf &
                       "               sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa, sb.NomeProduto; " & vbCrLf
            Else
                Sql &= "      GROUP BY sb.Empresa_Id, sb.EndEmpresa_Id,sb.Deposito_Id,sb.EndDeposito_Id, sb.Produto_Id," & vbCrLf &
                       "               sb.Operacao, sb.SubOperacao, sb.Descricao, sb.EntradaSaida, " & vbCrLf &
                       "               sb.NomeEmpresa,  sb.CidadeEmpresa,  sb.EstadoEmpresa,  sb.RedEmpresa," & vbCrLf &
                       "               sb.NomeDeposito,  sb.CidadeDeposito,  sb.EstadoDeposito,  sb.RedDeposito," & vbCrLf &
                       "               sb.NomeProduto;  " & vbCrLf
            End If


            Sql &= "              SELECT sb_saldo.Empresa_Id, sb_saldo.EndEmpresa_Id," & IIf(CkCDeposito.Checked, "sb_saldo.Empresa_Id as Deposito, sb_saldo.EndEmpresa_Id as EndDeposito,", "sb_saldo.Deposito, sb_saldo.EndDeposito,") & vbCrLf &
                   "                      sb_saldo.Produto_Id, Sum(sb_saldo.EstoqueAnteriorPeriodo) as EstoqueAnteriorPeriodo, Sum(sb_saldo.EstoqueAnteriorExercicio) as EstoqueAnteriorExercicio " & vbCrLf &
                   "                 INTO #SA" & vbCrLf &
                   "                 From (  " & vbCrLf &
                   "				       SELECT ESub.Empresa_Id," & vbCrLf &
                   "                              ESub.EndEmpresa_Id," & vbCrLf &
                   "                              ESub.Deposito_Id as Deposito ," & vbCrLf &
                   "                              ESub.EndDeposito_Id as EndDeposito," & vbCrLf &
                   "                              ESub.Produto_Id, " & vbCrLf &
                   "                              sum(case when year(ESub.Movimento_Id) < " & CDate(DataInicial).Year & " then isnull(ESub.Entradas,0) - isnull(ESub.saidas,0) else 0 end) as EstoqueAnteriorExercicio," & vbCrLf &
                   "                              SUM(isnull(ESub.Entradas,0) - isnull(ESub.saidas,0)) AS  EstoqueAnteriorPeriodo " & vbCrLf &
                   "	                     FROM Producao ESub  " & vbCrLf &
                   "                        INNER JOIN Produtos " & vbCrLf &
                   "                           ON Produtos.Produto_Id = ESub.Produto_Id " & vbCrLf &
                   "	                    INNER JOIN SubOperacoes SO " & vbCrLf &
                   "	        	           ON ESub.Operacao_Id     = SO.Operacao_Id " & vbCrLf &
                   "	                      AND ESub.SubOperacao_Id  = SO.Suboperacoes_Id " & vbCrLf &
                   "	                      AND ESub.FisicoFiscal_Id = 2  " & vbCrLf &
                   "	                      AND SO.EstoqueFiscal     = 'S' " & vbCrLf &
                   "	                    INNER JOIN Clientes AS C  " & vbCrLf &
                   "	                       ON  ESub.Empresa_Id   = C.Cliente_Id " & vbCrLf &
                   "	        	          AND ESub.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                   "	                    INNER JOIN Clientes AS D " & vbCrLf &
                   "		                   ON  ESub.Deposito_Id   = D.Cliente_Id " & vbCrLf &
                   "		                  AND ESub.EndDeposito_Id = D.Endereco_Id " & vbCrLf &
                   "		                Where ESub.Movimento_Id < '" & DateValue(DataInicial).ToSqlDate() & "' " & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "		                  And Produtos.Situacao in(1) " & vbCrLf
            Else
                Sql &= "		                  And Produtos.Situacao not in(1) " & vbCrLf
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "		                  AND ESub.Empresa_Id    ='" & campo(0) & "' " & vbCrLf &
                       "		                  AND ESub.EndEmpresa_Id = " & campo(1) & " " & vbCrLf
            End If
            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "		                  And ESub.Deposito_Id   ='" & campo(0) & "' "
                Sql &= "		                  AND ESub.EndDeposito_Id = " & campo(1) & " "
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "ESub.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If


            Sql &= "		               GROUP BY  ESub.Empresa_Id, ESub.EndEmpresa_Id,ESub.Deposito_Id,ESub.EndDeposito_Id,ESub.Produto_Id " & vbCrLf &
                   "                       UNION ALL " & vbCrLf &
                   "                      SELECT NF.Empresa_Id, NF.EndEmpresa_Id, NFxISub.Deposito, NFxISub.EndDeposito, NFxISub.Produto_Id, " & vbCrLf &
                   "                             ISNULL(SUM(CASE When year(NF.Movimento) < " & CDate(DataInicial).Year & vbCrLf &
                   "                                              then case when SO.EntradaSaida = 'E' THEN NFxISub.QuantidadeFiscal Else NFxISub.QuantidadeFiscal * -1 END" & vbCrLf &
                   "                                              else 0" & vbCrLf &
                   "                                         end), 0) AS EstoqueAnteriorExercicio, " & vbCrLf &
                   "                             ISNULL(SUM(CASE WHEN SO.EntradaSaida = 'E' THEN NFxISub.QuantidadeFiscal Else NFxISub.QuantidadeFiscal * -1 END), 0) AS EstoqueAnteriorPeriodo " & vbCrLf &
                   "	                    FROM NotasFiscais NF " & vbCrLf &
                   " 	                   INNER JOIN NotasFiscaisXItens NFxISub  " & vbCrLf &
                   "	                      ON NF.Empresa_Id      = NFxISub.Empresa_Id  " & vbCrLf &
                   "	                     AND NF.EndEmpresa_Id   = NFxISub.EndEmpresa_Id  " & vbCrLf &
                   "	                     AND NF.Cliente_Id      = NFxISub.Cliente_Id  " & vbCrLf &
                   "	                     AND NF.EndCliente_Id   = NFxISub.EndCliente_Id  " & vbCrLf &
                   "	                     AND NF.EntradaSaida_Id = NFxISub.EntradaSaida_Id " & vbCrLf &
                   "	                     AND NF.Serie_Id        = NFxISub.Serie_Id  " & vbCrLf &
                   "	                     AND NF.Nota_Id         = NFxISub.Nota_Id " & vbCrLf &
                   "	                   INNER JOIN SubOperacoes SO " & vbCrLf &
                   "	                      ON NFxISub.Operacao    = SO.Operacao_Id " & vbCrLf &
                   "	                     AND NFxISub.SubOperacao = SO.Suboperacoes_Id " & vbCrLf &
                   "	                     AND SO.EstoqueFiscal    = 'S' " & vbCrLf &
                   "	                   INNER JOIN Produtos Prd " & vbCrLf &
                   "	                      ON Prd.Produto_id = NFxISub.Produto_Id " & vbCrLf &
                   "                       INNER JOIN GruposDeEstoques " & vbCrLf &
                   "                          ON GruposDeEstoques.Grupo_Id = Prd.Grupo " & vbCrLf &
                   "		                 AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
            ' "	                     AND Prd.Situacao   = 1 " & vbCrLf & _
            Sql &= "	                   INNER JOIN  Clientes C " & vbCrLf &
                   "	                      ON NFxISub.Empresa_Id    = C.Cliente_Id  " & vbCrLf &
                   "	                     AND NFxISub.EndEmpresa_Id = C.Endereco_Id " & vbCrLf &
                   "	                    LEFT JOIN  Clientes D " & vbCrLf &
                   "	                      ON NFxISub.Deposito    = D.Cliente_Id  " & vbCrLf &
                   "	                     AND NFxISub.EndDeposito = D.Endereco_Id " & vbCrLf &
                   "	                   WHERE NF.Movimento    < '" & DateValue(DataInicial).ToSqlDate() & "'" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "           AND Prd.Situacao in(1) " & vbCrLf
            Else
                Sql &= "           AND Prd.Situacao not in(1) " & vbCrLf
            End If

            If Not chkConsNotasComSerieEspecificas.Checked Then
                Sql &= "  AND NF.Serie_id not in('101','102','103','104') " & vbCrLf
            End If
            'If chkConsNotasComSerieEspecificas.Checked Then
            '    Sql &= "  AND NF.Serie_id not in('501','502') " & vbCrLf
            'Else
            '    Sql &= "  AND NF.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
            'End If
            If DdlDeposito.Text <> "" Then
                Cliente = DdlDeposito.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "      AND NFxISub.Deposito   = '" & campo(0) & "' "
                Sql &= "	  AND NFxISub.EndDeposito =  " & campo(1) & " "
            End If

            If ucSelecaoProduto.TemSelecionado Then
                Dim RetornoProdutos As ArrayList
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NFxISub.Produto_Id", "")
                Sql &= " AND " & RetornoProdutos(0)
            End If

            If DdlEmpresa.Text <> "" Then
                Cliente = DdlEmpresa.SelectedValue
                campo = Cliente.Split("-")
                Sql &= "	                     AND NF.Empresa_Id    = '" & campo(0) & "' " & vbCrLf &
                       "	                     AND NF.EndEmpresa_Id =  " & campo(1) & " " & vbCrLf
            End If

            Sql &= "	                     AND NF.Situacao        = 1" & vbCrLf &
                   "	                     AND NF.TipoDeDocumento = 1" & vbCrLf &
                   "                 	   GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NFxISub.Deposito, NFxISub.EndDeposito, NFxISub.Produto_Id     " & vbCrLf

            If chkContraPartida.Checked Then
                Sql &= "                       UNION ALL " & vbCrLf &
                       "                      SELECT NF.Empresa_Id, NF.EndEmpresa_Id, nf.Destino, NF.EndDestino, NFxISub.Produto_Id," & vbCrLf &
                       "                             ISNULL(SUM(CASE When year(NF.Movimento) < " & CDate(DataInicial).Year & vbCrLf &
                       "                                              then case when SOD.EntradaSaida = 'E' THEN NFxISub.QuantidadeFiscal Else NFxISub.QuantidadeFiscal * -1 END" & vbCrLf &
                       "                                              else 0" & vbCrLf &
                       "                                         end), 0) AS EstoqueAnteriorExercicio, " & vbCrLf &
                       "                             ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN NFxISub.QuantidadeFiscal Else NFxISub.QuantidadeFiscal * -1 END), 0) AS EstoqueAnteriorPeriodo " & vbCrLf &
                       "	                    FROM NotasFiscais NF" & vbCrLf &
                       " 	                   INNER JOIN NotasFiscaisXItens NFxISub" & vbCrLf &
                       "	                      ON NF.Empresa_Id      = NFxISub.Empresa_Id" & vbCrLf &
                       "	                     AND NF.EndEmpresa_Id   = NFxISub.EndEmpresa_Id" & vbCrLf &
                       "	                     AND NF.Cliente_Id      = NFxISub.Cliente_Id" & vbCrLf &
                       "	                     AND NF.EndCliente_Id   = NFxISub.EndCliente_Id" & vbCrLf &
                       "	                     AND NF.EntradaSaida_Id = NFxISub.EntradaSaida_Id" & vbCrLf &
                       "	                     AND NF.Serie_Id        = NFxISub.Serie_Id" & vbCrLf &
                       "	                     AND NF.Nota_Id         = NFxISub.Nota_Id" & vbCrLf &
                       "	                   INNER JOIN SubOperacoes SO" & vbCrLf &
                       "	                      ON NFxISub.Operacao    = SO.Operacao_Id" & vbCrLf &
                       "	                     AND NFxISub.SubOperacao = SO.Suboperacoes_Id" & vbCrLf &
                       "	                     AND SO.EstoqueFiscal    = 'S'" & vbCrLf &
                       "	                   INNER JOIN SubOperacoes SOD" & vbCrLf &
                       "	                      ON SOD.Operacao_Id     = SO.OperacaoDestino" & vbCrLf &
                       "	                     AND SOD.Suboperacoes_Id = SO.SuboperacaoDestino" & vbCrLf &
                       "	                     AND SO.EstoqueFiscal    = 'S'" & vbCrLf &
                       "	                   INNER JOIN Produtos Prd " & vbCrLf &
                       "	                      ON Prd.Produto_id = NFxISub.Produto_Id " & vbCrLf &
                       "                       INNER JOIN GruposDeEstoques " & vbCrLf &
                       "                          ON GruposDeEstoques.Grupo_Id = Prd.Grupo " & vbCrLf &
                       "		                 AND GruposDeEstoques.MapaDeEstoque = 1 " & vbCrLf
                '"	                     AND Prd.Situacao   = 1 " & vbCrLf & _
                Sql &= "	                   INNER JOIN  Clientes C" & vbCrLf &
                       "	                      ON NFxISub.Empresa_Id    = C.Cliente_Id" & vbCrLf &
                       "	                     AND NFxISub.EndEmpresa_Id = C.Endereco_Id" & vbCrLf &
                       "	                    LEFT JOIN  Clientes D" & vbCrLf &
                       "	                      ON nf.Destino    = D.Cliente_Id" & vbCrLf &
                       "	                     AND nf.EndDestino = D.Endereco_Id " & vbCrLf &
                       "	                   WHERE NF.Situacao        = 1 " & vbCrLf &
                       "	                     AND NF.TipoDeDocumento = 1 " & vbCrLf

                If rdAtivo.Checked Then
                    Sql &= "           AND Prd.Situacao in(1) " & vbCrLf
                Else
                    Sql &= "           AND Prd.Situacao not in(1) " & vbCrLf
                End If

                If Not chkConsNotasComSerieEspecificas.Checked Then
                    Sql &= " AND NF.Serie_id not in('101','102','103','104') " & vbCrLf
                End If
                'If chkConsNotasComSerieEspecificas.Checked Then
                '    Sql &= "  NF.Serie_id not in('501','502') " & vbCrLf
                'Else
                '    Sql &= "  NF.Serie_id not in('501','502','101','102','103','104') " & vbCrLf
                'End If
                If DdlDeposito.Text <> "" Then
                    Cliente = DdlDeposito.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "      AND NF.Destino    = '" & campo(0) & "'" & vbCrLf &
                           "	  AND NF.EndDestino =  " & campo(1) & vbCrLf
                End If

                If ucSelecaoProduto.TemSelecionado Then
                    Dim RetornoProdutos As ArrayList
                    RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "NFxISub.Produto_Id", "")
                    Sql &= " AND " & RetornoProdutos(0)
                End If

                If DdlEmpresa.Text <> "" Then
                    Cliente = DdlEmpresa.SelectedValue
                    campo = Cliente.Split("-")
                    Sql &= "	                     AND NF.Empresa_Id    = '" & campo(0) & "' " & vbCrLf &
                           "	                     AND NF.EndEmpresa_Id =  " & campo(1) & " " & vbCrLf
                End If

                Sql &= "	                     AND NF.Movimento        < '" & DateValue(DataInicial).ToSqlDate() & "'" & vbCrLf &
                       "                         AND SO.Deposito         = 'S'" & vbCrLf &
                       "                         AND year(NF.Movimento) >= 2012" & vbCrLf &
                       "                 	   GROUP BY NF.Empresa_Id, NF.EndEmpresa_Id, NF.Destino, NF.EndDestino, NFxISub.Produto_Id" & vbCrLf
            End If

            Sql &= "                      ) sb_saldo " & vbCrLf

            If CkCDeposito.Checked Then
                Sql &= "       	group by sb_saldo.Empresa_Id,sb_saldo.EndEmpresa_Id,sb_saldo.Produto_Id " & vbCrLf
            Else
                Sql &= "       	group by sb_saldo.Empresa_Id, sb_saldo.EndEmpresa_Id ,sb_saldo.Deposito, sb_saldo.EndDeposito, sb_saldo.Produto_Id " & vbCrLf
            End If

            Sql &= " Update #temp set" & vbCrLf &
                   "     EstoqueAnteriorPeriodo   = #SA.EstoqueAnteriorPeriodo" & vbCrLf &
                   "    ,EstoqueAnteriorExercicio = #SA.EstoqueAnteriorExercicio" & vbCrLf &
                   "  from #temp " & vbCrLf &
                   " Inner Join #Sa" & vbCrLf &
                   "    ON #temp.Produto_Id     = #Sa.Produto_Id " & vbCrLf &
                   "   AND #temp.Empresa_Id     = #Sa.Empresa_Id " & vbCrLf &
                   "   AND #temp.EndEmpresa_Id  = #Sa.EndEmpresa_Id" & vbCrLf &
                   "   AND #temp.Deposito_Id    = #Sa.Deposito " & vbCrLf &
                   "   AND #temp.EndDeposito_Id = #Sa.EndDeposito " & vbCrLf

            Sql &= " Insert Into #Temp(Empresa_Id,  EndEmpresa_Id, " & vbCrLf &
                   " 				   Deposito_Id, EndDeposito_Id," & vbCrLf &
                   " 				   Produto_Id, NomeProduto," & vbCrLf &
                   "                   Operacao , SubOperacao, Descricao," & vbCrLf &
                   " 				   EntradaSaida, ValEntradaExercicio, ValEntradaPeriodo, ValSaidaExercicio, ValSaidaPeriodo," & vbCrLf &
                   " 				   NomeEmpresa,  CidadeEmpresa, EstadoEmpresa, RedEmpresa," & vbCrLf &
                   " 				   NomeDeposito,	CidadeDeposito, EstadoDeposito, RedDeposito," & vbCrLf &
                   " 				   EstoqueAnteriorPeriodo, EstoqueAnteriorExercicio)" & vbCrLf &
                   " (" & vbCrLf &
                   " Select #SA.Empresa_Id, #Sa.EndEmpresa_Id," & vbCrLf &
                   "        #Sa.Deposito,   #Sa.EndDeposito," & vbCrLf &
                   "        #Sa.Produto_Id, " & getCampoProduto("Prd") & "," & vbCrLf &
                   "        0,0,'Sem Movimentacao No Periodo'," & vbCrLf &
                   "        0,0,0,0,0," & vbCrLf &
                   "        Emp.Nome, Emp.Cidade, Emp.Estado, Emp.Reduzido," & vbCrLf &
                   "        Dep.Nome, Dep.Cidade, Dep.Estado, Dep.Reduzido," & vbCrLf &
                   "        #SA.EstoqueAnteriorPeriodo, #SA.EstoqueAnteriorExercicio" & vbCrLf &
                   "   from #SA" & vbCrLf &
                   "  Inner Join Clientes as Emp" & vbCrLf &
                   "     on Emp.Cliente_id  = #SA.Empresa_Id" & vbCrLf &
                   "    and Emp.Endereco_id = #Sa.EndEmpresa_Id" & vbCrLf &
                   "  Inner Join Clientes as Dep" & vbCrLf &
                   "     on Dep.Cliente_id  = #SA.Deposito" & vbCrLf &
                   "    and Dep.Endereco_id = #Sa.EndDeposito" & vbCrLf &
                   "  Inner Join Produtos Prd" & vbCrLf &
                   "     on Prd.Produto_id = #Sa.Produto_id" & vbCrLf &
                   "  Where #SA.EstoqueAnteriorExercicio > 0" & vbCrLf

            If rdAtivo.Checked Then
                Sql &= "    and Prd.Situacao in(1) " & vbCrLf
            Else
                Sql &= "    and Prd.Situacao not in(1) " & vbCrLf
            End If

            Sql &= "    and not exists(Select 1" & vbCrLf &
                   "                     from #temp" & vbCrLf &
                   "                    Where #temp.Produto_Id    = #Sa.Produto_Id" & vbCrLf &
                   " 					  and #temp.Empresa_Id     = #Sa.Empresa_Id" & vbCrLf &
                   " 					  and #temp.EndEmpresa_Id  = #Sa.EndEmpresa_Id" & vbCrLf &
                   " 					  and #temp.Deposito_Id    = #Sa.Deposito" & vbCrLf &
                   " 					  and #temp.EndDeposito_Id = #Sa.EndDeposito" & vbCrLf &
                   "                   )" & vbCrLf &
                   " )" & vbCrLf

            Sql2 = " select Empresa_Id, EndEmpresa_Id, " & vbCrLf &
            IIf(CkCDeposito.Checked = False, "Deposito_Id, ", "Empresa_Id as Deposito_Id, ") & vbCrLf &
            IIf(CkCDeposito.Checked = False, "EndDeposito_Id, ", "EndEmpresa_Id as EndDeposito_Id, ") & vbCrLf &
                   "        Produto_Id, NomeProduto, Operacao , SubOperacao, Descricao, " & vbCrLf &
                   "        EntradaSaida, ValEntradaExercicio, ValEntradaPeriodo, ValSaidaExercicio, ValSaidaPeriodo, " & vbCrLf &
                   "        NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa,  " & vbCrLf &
            IIf(CkCDeposito.Checked = False, "NomeDeposito, ", "NomeEmpresa as NomeDeposito, ") & vbCrLf &
            IIf(CkCDeposito.Checked = False, "CidadeDeposito, ", "CidadeEmpresa as CidadeDeposito, ") & vbCrLf &
            IIf(CkCDeposito.Checked = False, "EstadoDeposito, ", "EstadoEmpresa as EstadoDeposito, ") & vbCrLf &
            IIf(CkCDeposito.Checked = False, "RedDeposito, ", "RedEmpresa as RedDeposito, ") & vbCrLf &
                   "        0 as EstoqueAnterior " & vbCrLf &
                   "  into #temp2" & vbCrLf &
                   "  from #temp " & vbCrLf &
                   " Order by Empresa_Id, EndEmpresa_Id, Deposito_Id,  EndDeposito_Id, Produto_Id, Operacao, SubOperacao" & vbCrLf

            Sql2 &= " select Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                    "        case" & vbCrLf &
                    "            when Deposito_Id = Empresa_Id" & vbCrLf &
                    "                then '00000000000001'" & vbCrLf &
                    "                else Deposito_Id" & vbCrLf &
                    "        end as Deposito_Id, " & vbCrLf &
                    "        EndDeposito_Id, Produto_Id, NomeProduto, Operacao , SubOperacao, Descricao, " & vbCrLf &
                    "        EntradaSaida, ValEntradaExercicio, ValEntradaPeriodo, ValSaidaExercicio, ValSaidaPeriodo, " & vbCrLf &
                    "        NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa, " & vbCrLf &
                    "        NomeDeposito, CidadeDeposito, EstadoDeposito, RedDeposito, EstoqueAnterior " & vbCrLf &
                    " from #temp2 " & vbCrLf &
                    "  order by Empresa_Id, EndEmpresa_Id, Deposito_Id,  EndDeposito_Id, Produto_Id, Operacao, SubOperacao " & vbCrLf

            If ExcelDados Then
                Sql &= "SELECT Empresa_Id, EndEmpresa_Id, NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa, Deposito_Id, EndDeposito_Id, NomeDeposito, " & vbCrLf &
                       "CidadeDeposito, EstadoDeposito, RedDeposito, Produto_Id, NomeProduto, EntradaSaida, Operacao, SubOperacao, Descricao, " & vbCrLf &
                       "SUM(EstoqueAnteriorPeriodo)    As TotalEstoqueAnteriorPeriodo, SUM(EstoqueAnteriorExercicio)   As TotalEstoqueAnteriorExercicio, " & vbCrLf &
                       "SUM(ValEntradaPeriodo)          As TotalValEntradaPeriodo, SUM(ValEntradaExercicio)        As TotalValEntradaExercicio, " & vbCrLf &
                       "SUM(ValSaidaPeriodo)            As TotalValSaidaPeriodo, SUM(ValSaidaExercicio)          As TotalValSaidaExercicio, " & vbCrLf &
                       "SUM(0) As TotalEstoque " & vbCrLf &
                       "into #VerExcel " & vbCrLf &
                       "FROM #temp" & vbCrLf &
                       "Group BY  Empresa_Id, EndEmpresa_Id, NomeEmpresa, CidadeEmpresa, EstadoEmpresa, RedEmpresa, " & vbCrLf &
                       "Deposito_Id, EndDeposito_Id,NomeDeposito, CidadeDeposito, EstadoDeposito, RedDeposito, " & vbCrLf &
                       "Produto_Id, NomeProduto, EntradaSaida,Operacao, SubOperacao, Descricao" & vbCrLf &
                       "ORDER BY Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, EntradaSaida, Operacao, SubOperacao;" & vbCrLf &
                       "Select * from #VerExcel " & vbCrLf
            End If

            DS = Banco.ConsultaDataSet(Sql & Sql2, "EstoquePorOperacao")

            Dim sql3 As String
            sql3 = " select Empresa_Id,   EndEmpresa_Id, " & vbCrLf &
                    "        case" & vbCrLf &
                    "            when Deposito_Id = Empresa_Id" & vbCrLf &
                    "                then '00000000000001'" & vbCrLf &
                    "                else Deposito_Id" & vbCrLf &
                    "        end as Deposito_Id, EndDeposito_Id, Produto_Id, " & vbCrLf &
                   "        EstoqueAnteriorPeriodo, " & vbCrLf &
                   "        EstoqueAnteriorExercicio, " & vbCrLf &
                   "        SUM(ValEntradaPeriodo)   as TotalEntradasPeriodo, " & vbCrLf &
                   "        SUM(ValEntradaExercicio) as TotalEntradasExercicio, " & vbCrLf &
                   "        SUM(ValSaidaPeriodo)     as TotalSaidasPeriodo, " & vbCrLf &
                   "        SUM(ValSaidaExercicio)   as TotalSaidasExercicio " & vbCrLf &
                   "   FROM #temp " & vbCrLf &
                   "  group by Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, EstoqueAnteriorPeriodo, EstoqueAnteriorExercicio " & vbCrLf &
                   "  Order by Empresa_Id, EndEmpresa_Id,Deposito_Id, EndDeposito_Id, Produto_Id " & vbCrLf

            DS.Merge(Banco.ConsultaDataSet(Sql & sql3, "TotalEstoquePorOperacao"))

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return DS
    End Function

    Private Function getParam() As String
        Dim param As String = "Parametros: " & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            param &= "Empresa: " & DdlEmpresa.SelectedItem.Text.Split(".")(0) & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(TxtDiaMesInicial.Text) AndAlso Not String.IsNullOrWhiteSpace(ddlMesInicial.SelectedValue) AndAlso Not String.IsNullOrWhiteSpace(TxtDiaMesFinal.Text) AndAlso Not String.IsNullOrWhiteSpace(ddlMesFinal.SelectedValue) Then
            param &= "Movimento: " & TxtDiaMesInicial.Text & "-" & ddlMesInicial.SelectedItem.Text & "-" & ddlExercicio.SelectedItem.Text & " até " & TxtDiaMesFinal.Text & "-" & ddlMesFinal.SelectedItem.Text & "-" & ddlExercicio.SelectedItem.Text & vbCrLf
        End If

        param &= IIf(RadFisico.Checked, "Estoque: Fisico.", "Estoque: Fiscal.")

        Return param
    End Function

    Private Sub EmitirRelatorioDados()
        Try
            If Funcoes.VerificaPermissao("CustoMedio", "RELATORIO") Then
                Dim empresa = DdlEmpresa.SelectedValue.ToString.Split("-")
                Dim objEmpresa As New [Lib].Negocio.Cliente(empresa(0), empresa(1))
                Dim data As String = CStr(TxtDiaMesInicial.Text & " " & ddlMesInicial.SelectedItem.Text & " - " & TxtDiaMesFinal.Text & " " & ddlMesFinal.SelectedItem.Text & "/" & ddlExercicio.SelectedValue)
                Dim ds As New DataSet
                Dim dt As DataTable = New DataTable()
                Dim titulo As String = ""

                If RadFisico.Checked Then
                    ds = Fisico(True)
                ElseIf RadFiscal.Checked Then
                    ds = Fiscal(True)
                End If

                If ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                    Exit Sub
                End If

                dt = ds.Tables(0)

                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)

                        'criando planilha 
                        Dim worksheet As ExcelWorksheet

                        'criando título da planilha 
                        worksheet = package.Workbook.Worksheets.Add("Relatório de Estoques Por Operacoes")

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
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Relatório de Estoques Por Operacoes")
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        'criando linha que informa o período selecionado na página
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                        worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                        worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PERÍODO : " & data)
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                        worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                        rowIndex += 1

                        For Each col As DataColumn In dt.Columns
                            If col.ColumnName <> "Row" Then
                                worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                columnIndex += 1
                            End If
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A5:N" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        Dim lastProdutoId As Object = Nothing

                        Dim totalEstoque As Decimal = 0


                        ' criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In dt.Rows
                            columnIndex = 1

                            Dim currentProdutoId = row("Produto_Id")
                            Dim firstProdutoId As Boolean = False

                            If lastProdutoId Is Nothing OrElse Not lastProdutoId.Equals(currentProdutoId) Then
                                ' nova chave de produto → é o primeiro dessa chave
                                firstProdutoId = True
                                lastProdutoId = currentProdutoId
                            End If

                            For Each col As DataColumn In dt.Columns
                                If col.ColumnName = "TotalEstoqueAnteriorPeriodo" OrElse col.ColumnName = "TotalEstoqueAnteriorExercicio" Then
                                    If firstProdutoId Then
                                        ' exibe o valor normal
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)

                                        If col.ColumnName = "TotalEstoqueAnteriorExercicio" Then totalEstoque = row(col.ColumnName)
                                    Else
                                        ' zera pois não é o primeiro da chave
                                        worksheet.Cells(rowIndex, columnIndex).Value = 0

                                    End If
                                Else
                                    ' para outras colunas
                                    If col.ColumnName = "TotalEstoque" Then
                                        worksheet.Cells(rowIndex, columnIndex).Value = totalEstoque
                                    Else
                                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                    End If

                                    If col.ColumnName = "TotalValEntradaExercicio" Then totalEstoque += row(col.ColumnName)
                                    If col.ColumnName = "TotalValSaidaExercicio" Then totalEstoque -= row(col.ColumnName)

                                End If

                                columnIndex += 1
                            Next

                            'For Each col As DataColumn In dt.Columns
                            '    If col.ColumnName = "Produto_id" Then
                            '        If ID = row(col.ColumnName) Then

                            '        Else
                            '            Dim id = row(col.ColumnName)
                            '        End If

                            '        If col.ColumnName = "TotalEstoqueAnteriorPeriodo" Or "TotalEstoqueAnteriorExercicio" Then
                            '            worksheet.Cells(rowIndex, columnIndex).Value = "0,0000"
                            '        Else

                            '        End If
                            '    Else
                            '        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            '        columnIndex += 1
                            '    End If

                            'Next

                            ''formatando células datas
                            'worksheet.Cells(String.Format("Z{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                            'formatando células Peso
                            worksheet.Cells(String.Format("S{0}:Y{0}", rowIndex)).Style.Numberformat.Format = "#,####0.0000_ ;[Red]-#,####0.0000"
                            ''formatando células valores
                            'worksheet.Cells(String.Format("U{0}:X{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando quinta linha (cabeçalho)
                        worksheet.View.FreezePanes(6, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                'download do arquivo pelo browser
                'Download(fileName)
            Else
                MsgBox(Me.Page, "Usuário sempermissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("EstoquesPorOperacoes", "RELATORIO") Then
                Validar()
                If Mensagem = "" Then
                    Dim ds As DataSet = Nothing
                    If RadFisico.Checked Then
                        ds = Fisico(False)
                    ElseIf RadFiscal.Checked Then
                        ds = Fiscal(False)
                    End If

                    Dim parameters = New Dictionary(Of String, Object)()
                    parameters.Add("parameters", getParam())

                    Funcoes.BindReport(Me.Page, ds, "Cr_EstoquesPorOperacoes", eExportType.PDF, parameters)
                Else
                    MsgBox(Me.Page, Mensagem)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorioDados() 'Excel Dados
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
            Funcoes.Ajuda(Me.Page, "EstoquesPorOperacoes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class