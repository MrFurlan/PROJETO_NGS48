Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Drawing
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Partial Class PosicaoDeEstoques
    Inherits BasePage

    Private Sql As String
    Private DS As DataSet
    Private Mensagem As String
    Private Empresa() As String
    Private Deposito() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.ProducaoEstoque)
                If Funcoes.VerificaPermissao("PosicaoDeEstoques", "ACESSAR") Then
                    txtData.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidade()
                    VerificaUnidade()
                    CargaDepositos()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Producao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
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
        Next
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue.ToString, True)
    End Sub

    Private Sub CargaDepositos()
        ddl.Carregar(ddlDeposito, CarregarDDL.Tabela.Depositos, "", True)
    End Sub

    Private Sub Limpar()
        DdlUnidade.Enabled = True
        DdlEmpresa.Enabled = True
        ddlDeposito.Enabled = True
        RadFisico.Enabled = True
        RadFiscal.Enabled = True
        chkAlmoxarifado.Checked = False

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

    Function Validar() As Boolean
        If DdlUnidade.Text = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            Return False
        ElseIf DdlEmpresa.Text = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            Return False
        ElseIf Not IsDate(txtData.Text) Then
            MsgBox(Me.Page, "Data Inválida.")
            Return False
        Else
            Return True
        End If
    End Function

    Public Function Bissexto(ByVal intAno As Integer) As Boolean
        '
        ' verifica se um ano é bissexto
        '
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
    End Function

    Private Sub EmitirRelatorio(ByVal Pdf As Boolean, ByVal ExcelDados As Boolean)
        Try
            If Funcoes.VerificaPermissao("PosicaoDeEstoques", "RELATORIO") Then
                If Validar() Then
                    Dim SqlArray As New ArrayList
                    Dim i As Integer = 0
                    Dim Dia As String = ""
                    Dim Mes As String = Month(txtData.Text)
                    Dim Ano As String = Year(txtData.Text)
                    Dim strAnd As String = ""
                    Dim strGrupo As String = ""

                    Dia = Day(txtData.Text)

                    Dim FisicoFiscal As String = ""
                    If RadFisico.Checked = True Then
                        FisicoFiscal = "Fisico"
                    Else
                        FisicoFiscal = "Fiscal"
                    End If

                    Empresa = DdlEmpresa.SelectedValue.ToString.Split("-")
                    Dim objEmpresa As New [Lib].Negocio.Cliente(Empresa(0), Empresa(1))

                    Deposito = ddlDeposito.SelectedValue.ToString.Split("-")

                    Dim par As New ArrayList
                    If ucSelecaoProduto.TemSelecionado Then
                        par = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id")
                    End If

                    'Kitio
                    'If DdlProduto.SelectedIndex > 0 Then

                    'ElseIf DdlGrupo.SelectedIndex > 0 Then
                    '    If DdlGrupo.SelectedIndex > 0 Then strGrupo = " ProdutosConsulta.Grupo = '" & DdlGrupo.SelectedValue & "' "
                    'Else
                    '    For i = 0 To gridGrupoProduto.Rows.Count - 1
                    '        If CType(gridGrupoProduto.Rows(i).FindControl("chkGrupoProduto"), CheckBox).Checked = True Then
                    '            strGrupo &= strAnd & "ProdutosConsulta.Grupo = '" & gridGrupoProduto.Rows(i).Cells(1).Text & "' "
                    '            strAnd = " OR "
                    '        End If
                    '    Next
                    'End If

                    '**************************************************************************
                    '****************************** FISICO ************************************
                    '**************************************************************************
                    If RadFisico.Checked = True Then
                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo as Produto", "Consulta.Produto") & ", SUM(Consulta.Inicial) AS Inicial, " & vbCrLf
                        Else
                            Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, '" & objEmpresa.Codigo & "' as Deposito, " & objEmpresa.CodigoEndereco & "  as  EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo as Produto", "Consulta.Produto") & ", SUM(Consulta.Inicial) AS Inicial, " & vbCrLf
                        End If

                        Sql &= " " & IIf(strGrupo.Length > 0, "CONVERT(bit,Isnull(decimais.ControlarDecimais,0)) ", "Isnull(ProdutosConsulta.ControlarDecimais,0)") & " as ControlarDecimais, " & vbCrLf


                        Sql &= "       SUM(Consulta.EntradasDia) AS EntradasDia, SUM(Consulta.EntradasMes) AS EntradasMes, SUM(Consulta.EntradasAno) AS EntradasAno, " & vbCrLf &
                               "       SUM(Consulta.ProducaoDia) AS ProducaoDia, SUM(Consulta.ProducaoMes) AS ProducaoMes, SUM(Consulta.ProducaoAno) AS ProducaoAno, " & vbCrLf &
                               "       SUM(Consulta.ConsumoDia) AS ConsumoDia, SUM(Consulta.ConsumoMes) AS ConsumoMes, SUM(Consulta.ConsumoAno) AS ConsumoAno, " & vbCrLf &
                               "       SUM(Consulta.SaidasDia) AS SaidasDia, SUM(Consulta.SaidasMes) AS SaidasMes, SUM(Consulta.SaidasAno) AS SaidasAno, " & vbCrLf &
                               "       Empresa.Nome AS NomeDaEmpresa, Empresa.Cidade AS CidadeDaEmpresa, Empresa.Reduzido AS ReduzidoDaEmpresa, " & vbCrLf &
                               "       Empresa.Estado AS EstadoDaEmpresa, Depositos.Nome AS NomeDoDeposito, Depositos.Cidade AS CidadeDoDeposito, " & vbCrLf &
                               "       Depositos.Estado AS EstadoDoDeposito, Depositos.Reduzido AS ReduzidoDoDeposito, " & IIf(strGrupo.Length > 0, "GruposDeEstoques.Descricao", "ProdutosConsulta.Nome") & " as NomeDoProduto, " & vbCrLf &
                               "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao as NomeDoGrupo, isnull(ProdutosConsulta.EstoqueMinimo,0) AS EstoqueMinimo, isnull(PrdEstadoFisico.Descricao,'') AS DescricaoEstadoFisico" & vbCrLf &
                               "  FROM (" & vbCrLf &
                               "        Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto_Id as Produto, " & vbCrLf &
                               "               isnull(SUM(CASE  WHEN Year(Movimento) < " & Ano & "   THEN Entradas - Saidas END),0) As Inicial," & vbCrLf &
                               "		       0 as EntradasDia," & vbCrLf &
                               "		       0 as EntradasMes," & vbCrLf &
                               "	           0 as EntradasAno," & vbCrLf &
                               "	           isnull(SUM(CASE  WHEN day(Movimento)  = " & Dia & " And Month(Movimento) = " & Mes & " And year(Movimento) = " & Ano & " THEN Entradas  END),0) As ProducaoDia," & vbCrLf &
                               "               isnull(SUM(CASE  WHEN day(Movimento) <= " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Entradas  END),0) As ProducaoMes," & vbCrLf &
                               "		       isnull(SUM(CASE  WHEN Year(Movimento) = " & Ano & "                                                                      THEN Entradas  END),0) As ProducaoAno," & vbCrLf &
                               "               isnull(SUM(CASE  WHEN day(Movimento)  = " & Dia & " And Month(Movimento) = " & Mes & " And year(Movimento) = " & Ano & " THEN Saidas    END),0) As ConsumoDia," & vbCrLf &
                               "               isnull(SUM(CASE  WHEN day(Movimento) <= " & Dia & " And Month(Movimento) = " & Mes & " And year(Movimento) = " & Ano & " THEN Saidas    END),0) As ConsumoMes," & vbCrLf &
                               "	           isnull(SUM(CASE  WHEN Year(Movimento) = " & Ano & "                                                                      THEN Saidas    END),0) As ConsumoAno," & vbCrLf &
                               "	           0 as SaidasDia," & vbCrLf &
                               "	           0 as SaidasMes," & vbCrLf &
                               "	           0 as SaidasAno" & vbCrLf &
                               "         From (" & vbCrLf &
                               "               SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id As Deposito, Producao.EndDeposito_Id as EndDeposito, Producao.Produto_Id, Producao.Movimento_Id as Movimento," & vbCrLf &
                               "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf &
                               "                      ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 1 THEN Producao.Saidas END), 0) AS Saidas" & vbCrLf &
                               "                 FROM SubOperacoes " & vbCrLf &
                               "                INNER JOIN Producao " & vbCrLf &
                               "                   ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf &
                               "                  AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id" & vbCrLf &
                               "                INNER JOIN Produtos " & vbCrLf &
                               "                   ON Produtos.Produto_Id = Producao.Produto_Id " & vbCrLf &
                               "                INNER JOIN GruposDeEstoques " & vbCrLf &
                               "                   ON GruposDeEstoques.Grupo_id = Produtos.Grupo " & vbCrLf &
                               "                  AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                        '"                  AND Produtos.Situacao   = 1 " & vbCrLf & _
                        Sql &= "                WHERE SubOperacoes.EstoqueFisico   = 'S'" & vbCrLf &
                               "                  AND Producao.FisicoFiscal_Id     = 1" & vbCrLf &
                               "                  And Producao.Empresa_Id          = '" & objEmpresa.Codigo & "'" & vbCrLf &
                               "                  AND Producao.EndEmpresa_Id       = " & objEmpresa.CodigoEndereco & vbCrLf &
                               "                  AND Producao.Movimento_Id       <='" & Ano & "/" & Mes & "/" & Dia & "'"

                        If ddlDeposito.SelectedIndex > 0 Then
                            Sql &= "    And (Producao.Deposito_Id = '" & Deposito(0) & "') AND Producao.EndDeposito_Id = " & Deposito(1) & vbCrLf
                        End If

                        If ucSelecaoProduto.TemSelecionado Then
                            Sql &= " and " & par(0)
                        End If
                        'kitio
                        'If DdlProduto.SelectedIndex > 0 Then
                        '    Sql &= "    And (Producao.Produto_Id = '" & DdlProduto.SelectedValue & "') " & vbCrLf
                        'End If

                        Sql &= "                GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Movimento_Id" & vbCrLf &
                               "               ) as ConsultaProducao" & vbCrLf &
                               "          Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id" & vbCrLf &
                               "          UNION" & vbCrLf &
                               "         Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto, " & vbCrLf &
                               "		        isnull(SUM(CASE WHEN Year(Movimento) < " & Ano & "                                                                      THEN Entradas - Saidas END),0) As Inicial," & vbCrLf &
                               "                isnull(SUM(CASE WHEN day(Movimento)  = " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Entradas  END),0) As EntradasDia," & vbCrLf &
                               "	            isnull(SUM(CASE WHEN day(Movimento) <= " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Entradas  END),0) As EntradasMes," & vbCrLf &
                               "		        isnull(SUM(CASE WHEN Year(Movimento) = " & Ano & "                                                                      THEN Entradas  END),0) As EntradasAno," & vbCrLf &
                               "	            0 as ProducaoDia," & vbCrLf &
                               "	            0 as ProducaoMes," & vbCrLf &
                               "		        0 as ProducaoAno," & vbCrLf &
                               "	            0 as ConsumoDia," & vbCrLf &
                               "		        0 as ConsumoMes," & vbCrLf &
                               "                0 as ConsumoAno," & vbCrLf &
                               "		        isnull(SUM(CASE  WHEN day(Movimento)   = " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Saidas  END),0) As SaidasDia," & vbCrLf &
                               "                isnull(SUM(CASE  WHEN day(Movimento)  <= " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Saidas  END),0) As SaidasMes," & vbCrLf &
                               "                isnull(SUM(CASE  WHEN Year(Movimento)  = " & Ano & "                                                                      THEN Saidas  END),0) As SaidasAno" & vbCrLf &
                               "           From (" & vbCrLf &
                               "                 SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito, nfxi.Produto_Id as Produto, " & vbCrLf &
                               "                        NotasFiscais.Movimento," & vbCrLf &
                               "		                ISNULL(SUM(CASE WHEN SubOperacoes.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Entradas," & vbCrLf &
                               "		                ISNULL(SUM(CASE WHEN SubOperacoes.EntradaSaida = 'S' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Saidas" & vbCrLf &
                               "                   FROM NotasFiscaisxItens nfxi " & vbCrLf &
                               "                  INNER JOIN NotasFiscais " & vbCrLf &
                               "                     ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                               "                    AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                               "                    AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                               "                    AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                               "                    AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                               "                    AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                               "                    AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                               "                  Left JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf &
                               "                     ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                               "                    AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                               "                    AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                               "                    AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                               "                    AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                               "                    AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                               "                    AND nfxr.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                               "                  Left JOIN Romaneios " & vbCrLf &
                               "                     ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                               "                    AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                               "                    AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id" & vbCrLf &
                               "                  INNER JOIN SubOperacoes" & vbCrLf &
                               "                     ON SubOperacoes.Operacao_Id     = NotasFiscais.Operacao" & vbCrLf &
                               "                    AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao" & vbCrLf &
                               "                  INNER JOIN Produtos                            " & vbCrLf &
                               "                     ON Produtos.Produto_Id = nfxi.Produto_Id            " & vbCrLf &
                               "                INNER JOIN GruposDeEstoques " & vbCrLf &
                               "                   ON GruposDeEstoques.Grupo_id = Produtos.Grupo " & vbCrLf &
                               "                  AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                        '"                    AND Produtos.Situacao   = 1 " & vbCrLf & _
                        Sql &= "                  WHERE (SubOperacoes.EstoqueFisico  = 'S')" & vbCrLf &
                               "                    AND NotasFiscais.Situacao        = 1" & vbCrLf &
                               "                    AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf

                        If Not chkConsNotasComSerieEspecificas.Checked Then
                            Sql &= "               AND NotasFiscais.Serie_Id not in('101','102','103','104') " & vbCrLf
                        End If

                        Sql &= "                    AND (NotasFiscais.Empresa_Id = '" & objEmpresa.Codigo & "')" & vbCrLf &
                               "                    AND (NotasFiscais.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & ")" & vbCrLf &
                               "                    And (NotasFiscais.Deposito <> '')" & vbCrLf &
                               "                    AND (NotasFiscais.Movimento <= '" & Ano & "/" & Mes & "/" & Dia & "')  "

                        If ddlDeposito.SelectedIndex > 0 Then
                            Sql &= "                    And (NotasFiscais.Deposito = '" & Deposito(0) & "') AND NotasFiscais.EndDeposito = " & Deposito(1) & vbCrLf
                        End If

                        'kitio
                        'If DdlProduto.SelectedIndex > 0 Then
                        '    Sql &= "                    And (nfxi.Produto_Id = '" & DdlProduto.SelectedValue & "') " & vbCrLf
                        'End If
                        If ucSelecaoProduto.TemSelecionado Then
                            Sql &= " and " & par(0)
                        End If

                        Sql &= "                  GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Deposito, NotasFiscais.EndDeposito, nfxi.Produto_Id, NotasFiscais.Movimento" & vbCrLf

                        If chkContraPartida.Checked Then
                            '--****************************
                            '--*****  CONTRAPARTIDA FISICO  ******
                            '--****************************
                            Sql &= "    UNION ALL  " & vbCrLf &
                                   "   SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id,  NotasFiscais.Destino, NotasFiscais.EndDestino, nfxi.Produto_Id as Produto, " & vbCrLf &
                                   "          NotasFiscais.Movimento," & vbCrLf &
                                   "          ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'E' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Entradas," & vbCrLf &
                                   "          ISNULL(SUM(CASE WHEN SOD.EntradaSaida = 'S' THEN CASE WHEN Produtos.ControlarRomaneio = 'false' THEN nfxi.QuantidadeFisica ELSE Romaneios.PesoLiquido END END), 0) AS Saidas" & vbCrLf &
                                   "     FROM NotasFiscaisxItens nfxi " & vbCrLf &
                                   "    INNER JOIN NotasFiscais " & vbCrLf &
                                   "       ON nfxi.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                                   "      AND nfxi.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                                   "      AND nfxi.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                                   "      AND nfxi.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                                   "      AND nfxi.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                                   "      AND nfxi.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                                   "      AND nfxi.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                                   "    Left JOIN NotasFiscaisXRomaneios nfxr " & vbCrLf &
                                   "       ON nfxr.Empresa_Id      = NotasFiscais.Empresa_Id  " & vbCrLf &
                                   "      AND nfxr.EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id  " & vbCrLf &
                                   "      AND nfxr.Cliente_Id      = NotasFiscais.Cliente_Id  " & vbCrLf &
                                   "      AND nfxr.EndCliente_Id   = NotasFiscais.EndCliente_Id  " & vbCrLf &
                                   "      AND nfxr.EntradaSaida_Id = NotasFiscais.EntradaSaida_Id  " & vbCrLf &
                                   "      AND nfxr.Serie_Id        = NotasFiscais.Serie_Id  " & vbCrLf &
                                   "      AND nfxr.Nota_Id         = NotasFiscais.Nota_Id  " & vbCrLf &
                                   "    Left JOIN Romaneios " & vbCrLf &
                                   "       ON Romaneios.Empresa_Id    = nfxr.Empresa_Id " & vbCrLf &
                                   "      AND Romaneios.EndEmpresa_Id = nfxr.EndEmpresa_Id " & vbCrLf &
                                   "      AND Romaneios.Romaneio_Id   = nfxr.Romaneio_Id" & vbCrLf &
                                   "    INNER JOIN SubOperacoes" & vbCrLf &
                                   "       ON SubOperacoes.Operacao_Id     = NotasFiscais.Operacao" & vbCrLf &
                                   "      AND SubOperacoes.SubOperacoes_Id = NotasFiscais.SubOperacao" & vbCrLf &
                                   "    INNER JOIN SubOperacoes SOD                      " & vbCrLf &
                                   "       ON SOD.Operacao_Id     = SubOperacoes.OperacaoDestino     " & vbCrLf &
                                   "      AND SOD.Suboperacoes_Id = SubOperacoes.SuboperacaoDestino " & vbCrLf &
                                   "    INNER JOIN Produtos                            " & vbCrLf &
                                   "       ON Produtos.Produto_Id = nfxi.Produto_Id            " & vbCrLf &
                                   "    INNER JOIN GruposDeEstoques " & vbCrLf &
                                   "       ON GruposDeEstoques.Grupo_id       = Produtos.Grupo " & vbCrLf &
                                   "       AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                            '"      AND Produtos.Situacao   = 1 " & vbCrLf & _
                            Sql &= "    WHERE (SOD.EstoqueFisico           = 'S')" & vbCrLf &
                                   "      AND NotasFiscais.Situacao        = 1" & vbCrLf &
                                   "      AND NotasFiscais.TipoDeDocumento = 1" & vbCrLf &
                                   "      AND SubOperacoes.Deposito         = 'S'" & vbCrLf &
                                   "      AND year(NotasFiscais.Movimento) >= " & Ano & vbCrLf

                            If Not chkConsNotasComSerieEspecificas.Checked Then
                                Sql &= "               AND NotasFiscais.Serie_Id not in('101','102','103','104') " & vbCrLf
                            End If

                            Sql &= "                    AND (NotasFiscais.Empresa_Id = '" & objEmpresa.Codigo & "')" & vbCrLf &
                                   "                    AND (NotasFiscais.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & ")" & vbCrLf &
                                   "                    And (NotasFiscais.Deposito <> '')" & vbCrLf &
                                   "                    AND (NotasFiscais.Movimento <= '" & Ano & "/" & Mes & "/" & Dia & "')  "

                            If ddlDeposito.SelectedIndex > 0 Then
                                Sql &= "                    And (NotasFiscais.Destino = '" & Deposito(0) & "') AND NotasFiscais.EndDestino = " & Deposito(1) & vbCrLf
                            End If

                            If ucSelecaoProduto.TemSelecionado Then
                                Sql &= " and " & par(0)
                            End If
                            'kitio
                            'If DdlProduto.SelectedIndex > 0 Then
                            '    Sql &= "                    And (nfxi.Produto_Id = '" & DdlProduto.SelectedValue & "') " & vbCrLf
                            'End If

                            Sql &= "                  GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.Destino, NotasFiscais.EndDestino, nfxi.Produto_Id, NotasFiscais.Movimento" & vbCrLf



                            '--****************************
                            '--********  FIM  *************
                            '--**************************** 
                        End If


                        Sql &= "                 ) as ConsultaRomaneios" & vbCrLf &
                        "            Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto" & vbCrLf &
                        "           ) AS Consulta " & vbCrLf &
                        "      INNER JOIN Clientes AS Empresa ON Consulta.Empresa = Empresa.Cliente_Id AND Consulta.EndEmpresa = Empresa.Endereco_Id " & vbCrLf

                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql &= "      INNER JOIN Clientes AS Depositos ON Consulta.Deposito = Depositos.Cliente_Id AND Consulta.EndDeposito = Depositos.Endereco_Id " & vbCrLf
                        Else
                            Sql &= "      INNER JOIN Clientes AS Depositos ON '" & objEmpresa.Codigo & "' = Depositos.Cliente_Id AND " & objEmpresa.CodigoEndereco & " = Depositos.Endereco_Id " & vbCrLf
                        End If

                        Sql &= "      INNER JOIN Produtos AS ProdutosConsulta ON Consulta.Produto = ProdutosConsulta.Produto_Id " & vbCrLf &
                               "      LEFT JOIN EstadoFisicoIA AS PrdEstadoFisico ON PrdEstadoFisico.EstadoFisicoIA_Id = ProdutosConsulta.CodigoEstadoFisico" & vbCrLf &
                               "      INNER JOIN GruposDeEstoques ON ProdutosConsulta.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf
                        If strGrupo.Length > 0 Then 'Por Grupo de Produto
                            Sql &= "      INNER JOIN ( select grupo_id, " & vbCrLf &
                                   "                         case when exists(select 1 " & vbCrLf &
                                   "                                          from produtos p " & vbCrLf &
                                   "                                          where(p.grupo = grupo_id) " & vbCrLf &
                                   "                                          and p.controlardecimais =1 ) " & vbCrLf &
                                   "                         then 1 " & vbCrLf &
                                   "                         else 0 " & vbCrLf &
                                   "                         end controlardecimais " & vbCrLf &
                                   "                   from GruposDeEstoques " & vbCrLf &
                                   "                  ) decimais  " & vbCrLf &
                                   "       ON decimais.grupo_id = GruposDeEstoques.Grupo_Id  " & vbCrLf
                        End If

                        If rdAtivo.Checked Then
                            Sql &= "         WHERE ProdutosConsulta.Situacao in(1) " & vbCrLf
                        Else
                            Sql &= "         WHERE ProdutosConsulta.Situacao not in(1) " & vbCrLf
                        End If

                        Sql &= "        AND (" & IIf(strGrupo.Length > 0, "(" & strGrupo & ") AND (", "") & " Consulta.Inicial <> 0" & vbCrLf &
                               "         or Consulta.EntradasDia <> 0 " & vbCrLf &
                               "         or Consulta.EntradasMes <> 0" & vbCrLf &
                               "         or Consulta.EntradasAno <> 0 " & vbCrLf &
                               "         or Consulta.ProducaoDia <> 0" & vbCrLf &
                               "         or Consulta.ProducaoMes <> 0" & vbCrLf &
                               "         or Consulta.ProducaoAno <> 0" & vbCrLf &
                               "         or Consulta.ConsumoDia  <> 0" & vbCrLf &
                               "         or Consulta.ConsumoMes  <> 0" & vbCrLf &
                               "         or Consulta.ConsumoAno  <> 0 " & vbCrLf &
                               "         or Consulta.SaidasDia   <> 0" & vbCrLf &
                               "         or Consulta.SaidasMes   <> 0 " & vbCrLf &
                               "         or Consulta.SaidasAno   <> 0" & IIf(strGrupo.Length > 0, ")", "") & ")" & vbCrLf

                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql &= "      GROUP BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "", "Consulta.Produto,") & " Empresa.Nome, Empresa.Cidade, " & vbCrLf &
                                   "               Empresa.Reduzido, Empresa.Estado, Depositos.Nome, Depositos.Cidade, Depositos.Estado, Depositos.Reduzido, " & IIf(strGrupo.Length > 0, "", "ProdutosConsulta.Nome,") & " " & vbCrLf &
                                   "               ProdutosConsulta.Grupo, GruposDeEstoques.Descricao" & vbCrLf &
                                   " " & IIf(strGrupo.Length > 0, ",CONVERT(bit,Isnull(decimais.ControlarDecimais,0))", ",Isnull(ProdutosConsulta.ControlarDecimais,0), ProdutosConsulta.EstoqueMinimo, PrdEstadoFisico.Descricao") & " " & vbCrLf

                            If rdAlfabetico.Checked Then
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, ProdutosConsulta.Nome" & vbCrLf
                            Else
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo", "Consulta.Produto") & "" & vbCrLf
                            End If
                        Else
                            Sql &= "      GROUP BY Consulta.Empresa, Consulta.EndEmpresa, " & IIf(strGrupo.Length > 0, "", "Consulta.Produto,") & " Empresa.Nome, Empresa.Cidade, " & vbCrLf &
                                   "               Empresa.Reduzido, Empresa.Estado, Depositos.Nome, Depositos.Cidade, Depositos.Estado, Depositos.Reduzido, " & IIf(strGrupo.Length > 0, "", "ProdutosConsulta.Nome,") & " " & vbCrLf &
                                   "               ProdutosConsulta.Grupo, GruposDeEstoques.Descricao" & vbCrLf &
                                   " " & IIf(strGrupo.Length > 0, ",CONVERT(bit,Isnull(decimais.ControlarDecimais,0))", ",Isnull(ProdutosConsulta.ControlarDecimais,0), ProdutosConsulta.EstoqueMinimo, PrdEstadoFisico.Descricao") & " " & vbCrLf

                            If rdAlfabetico.Checked Then
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, ProdutosConsulta.Nome" & vbCrLf
                            Else
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo", "Consulta.Produto") & "" & vbCrLf
                            End If
                        End If
                    End If


                    '**************************************************************************
                    '****************************** FISCAL ************************************
                    '**************************************************************************
                    If RadFiscal.Checked = True Then

                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo as Produto", "Consulta.Produto") & ", SUM(Consulta.Inicial) AS Inicial, "
                        Else
                            Sql = "SELECT Consulta.Empresa, Consulta.EndEmpresa, '" & objEmpresa.Codigo & "' as Deposito, " & objEmpresa.CodigoEndereco & "  as EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo as Produto", "Consulta.Produto") & ", SUM(Consulta.Inicial) AS Inicial, "
                        End If

                        Sql &= " " & IIf(strGrupo.Length > 0, "CONVERT(bit,Isnull(decimais.ControlarDecimais,0)) ", "Isnull(ProdutosConsulta.ControlarDecimais,0)") & " as ControlarDecimais, " & vbCrLf

                        Sql &= "       SUM(Consulta.EntradasDia) AS EntradasDia, SUM(Consulta.EntradasMes) AS EntradasMes, SUM(Consulta.EntradasAno) AS EntradasAno, " & vbCrLf &
                               "       SUM(Consulta.ProducaoDia) AS ProducaoDia, SUM(Consulta.ProducaoMes) AS ProducaoMes, SUM(Consulta.ProducaoAno) AS ProducaoAno, " & vbCrLf &
                               "       SUM(Consulta.ConsumoDia) AS ConsumoDia, SUM(Consulta.ConsumoMes) AS ConsumoMes, SUM(Consulta.ConsumoAno) AS ConsumoAno, " & vbCrLf &
                               "       SUM(Consulta.SaidasDia) AS SaidasDia, SUM(Consulta.SaidasMes) AS SaidasMes, SUM(Consulta.SaidasAno) AS SaidasAno, " & vbCrLf &
                               "       Empresa.Nome AS NomeDaEmpresa, Empresa.Cidade AS CidadeDaEmpresa, Empresa.Reduzido AS ReduzidoDaEmpresa, " & vbCrLf &
                               "       Empresa.Estado AS EstadoDaEmpresa, Depositos.Nome AS NomeDoDeposito, Depositos.Cidade AS CidadeDoDeposito, " & vbCrLf &
                               "       Depositos.Estado AS EstadoDoDeposito, Depositos.Reduzido AS ReduzidoDoDeposito, " & vbCrLf

                        Sql &= "       case" & vbCrLf &
                               "          when isnull(ProdutosConsulta.CodigoEstadoFisico,0) > 0" & vbCrLf &
                               "             then " & IIf(strGrupo.Length > 0, "GruposDeEstoques.Descricao + ' - ' + PrdEstadoFisico.Descricao", "ProdutosConsulta.Nome + ' - ' + PrdEstadoFisico.Descricao") & vbCrLf &
                               "             else " & IIf(strGrupo.Length > 0, "GruposDeEstoques.Descricao", "ProdutosConsulta.Nome") & vbCrLf &
                               "          end as NomeDoProduto, " & vbCrLf

                        Sql &= "       ProdutosConsulta.Grupo, GruposDeEstoques.Descricao as NomeDoGrupo, isnull(ProdutosConsulta.EstoqueMinimo,0) AS EstoqueMinimo, isnull(PrdEstadoFisico.Descricao,'') AS DescricaoEstadoFisico" & vbCrLf &
                               "  FROM (" & vbCrLf &
                               "        Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto_Id as Produto, " & vbCrLf &
                               "               isnull(SUM(CASE  WHEN Year(Movimento) < " & Ano & "   THEN Entradas - Saidas END),0) As Inicial," & vbCrLf &
                               "		       0 as EntradasDia," & vbCrLf &
                               "		       0 as EntradasMes," & vbCrLf &
                               "	           0 as EntradasAno," & vbCrLf &
                               "	           isnull(SUM(CASE WHEN day(Movimento)  = " & Dia & " And Month(Movimento)  = " & Mes & " And year(Movimento) = " & Ano & " THEN Entradas  END),0) As ProducaoDia," & vbCrLf &
                               "               isnull(SUM(CASE WHEN day(Movimento) <= " & Dia & " And Month(Movimento)  = " & Mes & " And Year(Movimento) = " & Ano & " THEN Entradas  END),0) As ProducaoMes," & vbCrLf &
                               "		       isnull(SUM(CASE WHEN Year(Movimento) = " & Ano & "                                                                       THEN Entradas  END),0) As ProducaoAno," & vbCrLf &
                               "               isnull(SUM(CASE WHEN day(Movimento)  = " & Dia & " And Month(Movimento)  = " & Mes & " And year(Movimento) = " & Ano & " THEN Saidas    END),0) As ConsumoDia," & vbCrLf &
                               "               isnull(SUM(CASE WHEN day(Movimento) <= " & Dia & " And Month(Movimento)  = " & Mes & " And Year(Movimento) = " & Ano & " THEN Saidas    END),0) As ConsumoMes," & vbCrLf &
                               "	           isnull(SUM(CASE WHEN Year(Movimento) = " & Ano & "                                                                       THEN Saidas    END),0) As ConsumoAno," & vbCrLf &
                               "	           0 as SaidasDia," & vbCrLf &
                               "	           0 as SaidasMes," & vbCrLf &
                               "	           0 as SaidasAno" & vbCrLf &
                               "          From (" & vbCrLf &
                               "                SELECT Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id As Deposito, Producao.EndDeposito_Id as EndDeposito, Producao.Produto_Id, Producao.Movimento_Id as Movimento," & vbCrLf &
                               "                       ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 2 THEN Producao.Entradas END), 0) AS Entradas, " & vbCrLf &
                               "                       ISNULL(SUM(CASE WHEN Producao.FisicoFiscal_Id = 2  THEN Producao.Saidas END), 0) AS Saidas" & vbCrLf &
                               "                  FROM SubOperacoes " & vbCrLf &
                               "                 INNER JOIN Producao " & vbCrLf &
                               "                    ON SubOperacoes.Operacao_Id     = Producao.Operacao_Id " & vbCrLf &
                               "                   AND SubOperacoes.SubOperacoes_Id = Producao.SubOperacao_Id" & vbCrLf &
                               "                 INNER JOIN Produtos " & vbCrLf &
                               "                    ON Produtos.Produto_id = Producao.Produto_Id " & vbCrLf &
                               "                INNER JOIN GruposDeEstoques " & vbCrLf &
                               "                   ON GruposDeEstoques.Grupo_id = Produtos.Grupo " & vbCrLf

                        If Not chkAlmoxarifado.Checked Then
                            Sql &= "                  AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                        End If

                        '"                   AND Produtos.Situacao   = 1 " & vbCrLf & _

                        Sql &= "                 WHERE (Producao.FisicoFiscal_Id    = 2)" & vbCrLf &
                               "                   And (Producao.Empresa_Id         ='" & objEmpresa.Codigo & "')" & vbCrLf &
                               "                   AND Producao.EndEmpresa_Id       = " & objEmpresa.CodigoEndereco & vbCrLf &
                               "                   and Producao.Movimento_Id       <='" & Ano & "/" & Mes & "/" & Dia & "'" & vbCrLf

                        If chkAlmoxarifado.Checked Then
                            Sql &= "                  AND Produtos.Almoxarifado = 1" & vbCrLf
                        Else
                            Sql &= "                  AND SubOperacoes.EstoqueFiscal   ='S' " & vbCrLf
                        End If

                        If ddlDeposito.SelectedIndex > 0 Then
                            Sql &= "                   And (Producao.Deposito_Id = '" & Deposito(0) & "') AND Producao.EndDeposito_Id = " & Deposito(1)
                        End If

                        If ucSelecaoProduto.TemSelecionado Then
                            Sql &= " and " & par(0)
                        End If

                        Sql &= "                 GROUP BY Producao.Empresa_Id, Producao.EndEmpresa_Id, Producao.Deposito_Id, Producao.EndDeposito_Id, Producao.Produto_Id, Producao.Movimento_Id" & vbCrLf &
                               "                ) as Consulta" & vbCrLf &
                               "           Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id" & vbCrLf &
                               "           UNION ALL" & vbCrLf &
                               "          Select Empresa_Id as Empresa, EndEmpresa_Id as EndEmpresa, Deposito, EndDeposito, Produto_Id as Produto, " & vbCrLf &
                               "                 isnull(SUM(CASE WHEN Year(Movimento) < " & Ano & "   THEN Entradas - Saidas END),0) As Inicial," & vbCrLf &
                               "                 isnull(SUM(CASE WHEN EntradaSaida_Id = 'E' And day(Movimento)  = " & Dia & " And Month(Movimento) = " & Mes & " And year(Movimento) = " & Ano & " THEN Entradas  END),0) As EntradasDia," & vbCrLf &
                               "                 isnull(SUM(CASE WHEN EntradaSaida_Id = 'E' And day(Movimento) <= " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Entradas  END),0) As EntradasMes," & vbCrLf &
                               "                 isnull(SUM(CASE WHEN EntradaSaida_Id = 'E' And Year(Movimento) = " & Ano & "                                                                      THEN Entradas  END),0) As EntradasAno," & vbCrLf &
                               "	             0 as ProducaoDia," & vbCrLf &
                               "	             0 as ProducaoMes," & vbCrLf &
                               "	             0 as ProducaoAno," & vbCrLf &
                               "	             0 as ConsumoDia," & vbCrLf &
                               "	             0 as ConsumoMes," & vbCrLf &
                               "	             0 as ConsumoAno," & vbCrLf &
                               "	             isnull(SUM(CASE  WHEN EntradaSaida_Id = 'S' And day(Movimento)  = " & Dia & " And Month(Movimento) = " & Mes & " And year(Movimento) = " & Ano & " THEN Saidas  END),0) As SaidasDia," & vbCrLf &
                               "	             isnull(SUM(CASE  WHEN EntradaSaida_Id = 'S' And day(Movimento) <= " & Dia & " And Month(Movimento) = " & Mes & " And Year(Movimento) = " & Ano & " THEN Saidas  END),0) As SaidasMes," & vbCrLf &
                               "	             isnull(SUM(CASE  WHEN EntradaSaida_Id = 'S' And Year(Movimento) = " & Ano & "                                                                      THEN Saidas  END),0) As SaidasAno" & vbCrLf &
                               "           From (" & vbCrLf &
                               "                 SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, isnull(NotasFiscaisXItens.Deposito,NotasFiscais.Empresa_Id) as Deposito, isnull(NotasFiscaisXItens.EndDeposito,NotasFiscais.EndEmpresa_Id) as EndDeposito, " & vbCrLf &
                               "                        NotasFiscaisXItens.Produto_Id, NotasFiscais.Movimento, NotasFiscais.EntradaSaida_Id,  " & vbCrLf &
                               "                	    isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'E' THEN NotasFiscaisXItens.QuantidadeFiscal  END),0) As Entradas," & vbCrLf &
                               "                	    isnull(SUM(CASE  WHEN NotasFiscais.EntradaSaida_Id = 'S' THEN NotasFiscaisXItens.QuantidadeFiscal  END),0) As Saidas" & vbCrLf &
                               "                   FROM NotasFiscais " & vbCrLf &
                               "                  INNER JOIN NotasFiscaisXItens" & vbCrLf &
                               "                     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id  " & vbCrLf &
                               "                	AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id " & vbCrLf &
                               "                    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id  " & vbCrLf &
                               "                	AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
                               "                    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                               "                	AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id  " & vbCrLf &
                               "                    AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id " & vbCrLf &
                               "                  INNER JOIN SubOperacoes  " & vbCrLf &
                               "                     ON NotasFiscaisXItens.Operacao    = SubOperacoes.Operacao_Id  " & vbCrLf &
                               "                    AND NotasFiscaisXItens.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                               "                  INNER JOIN Produtos  " & vbCrLf &
                               "                     ON Produtos.Produto_id = NotasFiscaisXItens.Produto_Id  " & vbCrLf &
                               "                  INNER JOIN GruposDeEstoques " & vbCrLf &
                               "                     ON GruposDeEstoques.Grupo_id = Produtos.Grupo " & vbCrLf

                        If Not chkAlmoxarifado.Checked Then
                            Sql &= "                    AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                        End If

                        '"                    AND Produtos.Situacao   = 1 " & vbCrLf & _
                        Sql &= "                  WHERE (NotasFiscais.Situacao         = 1) " & vbCrLf &
                               "                    And (NotasFiscais.TipoDeDocumento  = 1)" & vbCrLf &
                               "                    And (NotasFiscais.Empresa_Id    = '" & objEmpresa.Codigo & "')" & vbCrLf &
                               "                    AND (NotasFiscais.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & ") " & vbCrLf &
                               "                    and NotasFiscais.Movimento     <='" & Ano & "/" & Mes & "/" & Dia & "'" & vbCrLf

                        If chkAlmoxarifado.Checked Then
                            Sql &= "                  AND Produtos.Almoxarifado = 1" & vbCrLf
                        Else
                            Sql &= "                  AND SubOperacoes.EstoqueFiscal   ='S' " & vbCrLf
                        End If

                        If ddlDeposito.SelectedIndex > 0 Then
                            Sql &= "                    And (NotasFiscaisXItens.Deposito = '" & Deposito(0) & "') AND NotasFiscaisXItens.EndDeposito = " & Deposito(1) & vbCrLf
                        End If

                        If ucSelecaoProduto.TemSelecionado Then
                            Sql &= " and " & par(0)
                        End If
                        'kitio
                        'If DdlProduto.SelectedIndex > 0 Then
                        '    Sql &= "                    And (NotasFiscaisXItens.Produto_Id = '" & DdlProduto.SelectedValue & "') " & vbCrLf
                        'End If

                        If Not chkConsNotasComSerieEspecificas.Checked Then
                            Sql &= "               AND NotasFiscais.Serie_Id not in('101','102','103','104') " & vbCrLf
                        End If

                        Sql &= "                   GROUP BY NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscaisXItens.Deposito, NotasFiscaisXItens.EndDeposito, NotasFiscaisXItens.Produto_Id, " & vbCrLf &
                               "                           NotasFiscais.Movimento, NotasFiscais.EntradaSaida_Id" & vbCrLf

                        If chkContraPartida.Checked Then
                            '--***********************************
                            '--*****  CONTRAPARTIDA FISCAL  ******
                            '--***********************************
                            Sql &= "    UNION ALL  " & vbCrLf &
                                   "    SELECT NotasFiscais.Empresa_Id, " & vbCrLf &
                                   "           NotasFiscais.EndEmpresa_Id, " & vbCrLf &
                                   "           NotasFiscais.Destino, " & vbCrLf &
                                   "           NotasFiscais.endDestino, " & vbCrLf &
                                   "           NotasFiscaisXItens.Produto_Id, " & vbCrLf &
                                   "           NotasFiscais.Movimento, " & vbCrLf &
                                   "           SOD.EntradaSaida, " & vbCrLf &
                                   "           isnull(SUM(CASE  WHEN SOD.EntradaSaida = 'E' THEN NotasFiscaisXItens.QuantidadeFiscal  END),0) As Entradas, " & vbCrLf &
                                   "           isnull(SUM(CASE  WHEN SOD.EntradaSaida = 'S' THEN NotasFiscaisXItens.QuantidadeFiscal  END),0) As Saidas " & vbCrLf &
                                   "      FROM NotasFiscais  " & vbCrLf &
                                   "     INNER JOIN NotasFiscaisXItens " & vbCrLf &
                                   "        ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id   " & vbCrLf &
                                   "       AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id  " & vbCrLf &
                                   "       AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id   " & vbCrLf &
                                   "       AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id  " & vbCrLf &
                                   "       AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id " & vbCrLf &
                                   "       AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id   " & vbCrLf &
                                   "       AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id  " & vbCrLf &
                                   "     INNER JOIN Produtos  " & vbCrLf &
                                   "        ON Produtos.Produto_id = NotasFiscaisXItens.Produto_Id  " & vbCrLf &
                                   "     INNER JOIN GruposDeEstoques " & vbCrLf &
                                   "        ON GruposDeEstoques.Grupo_id = Produtos.Grupo " & vbCrLf

                            If Not chkAlmoxarifado.Checked Then
                                Sql &= "       AND GruposDeEstoques.MapaDeEstoque = 1" & vbCrLf
                            End If

                            '"       AND Produtos.Situacao   = 1 " & vbCrLf & _
                            Sql &= "     INNER JOIN SubOperacoes so " & vbCrLf &
                                   "        ON NotasFiscaisXItens.Operacao    = so.Operacao_Id   " & vbCrLf &
                                   "       AND NotasFiscaisXItens.SubOperacao = so.SubOperacoes_Id " & vbCrLf
                            If Not chkAlmoxarifado.Checked Then
                                Sql &= "       AND so.EstoqueFiscal               = 'S' " & vbCrLf
                            End If

                            Sql &= "     INNER JOIN SubOperacoes SOD " & vbCrLf &
                                   "        ON SOD.Operacao_Id     = SO.OperacaoDestino  " & vbCrLf &
                                   "       AND SOD.Suboperacoes_Id = SO.SuboperacaoDestino  " & vbCrLf
                            If Not chkAlmoxarifado.Checked Then
                                Sql &= "       AND SOD.EstoqueFiscal   = 'S'   " & vbCrLf
                            End If

                            If ddlDeposito.SelectedIndex > 0 Then
                                Sql &= "                    And (NotasFiscais.Destino = '" & Deposito(0) & "') AND NotasFiscais.endDestino = " & Deposito(1) & vbCrLf
                            End If

                            If ucSelecaoProduto.TemSelecionado Then
                                Sql &= " and " & par(0)
                            End If
                            'kitio
                            'If DdlProduto.SelectedIndex > 0 Then
                            '    Sql &= "                    And (NotasFiscaisXItens.Produto_Id = '" & DdlProduto.SelectedValue & "') " & vbCrLf
                            'End If

                            If Not chkConsNotasComSerieEspecificas.Checked Then
                                Sql &= "               AND NotasFiscais.Serie_Id not in('101','102','103','104') " & vbCrLf
                            End If
                            Sql &= "     WHERE (NotasFiscais.Situacao        = 1)  " & vbCrLf &
                                   "       And (NotasFiscais.TipoDeDocumento = 1)" & vbCrLf &
                                   "       And (NotasFiscais.Empresa_Id    = '" & objEmpresa.Codigo & "')" & vbCrLf &
                                   "       AND (NotasFiscais.EndEmpresa_Id = " & objEmpresa.CodigoEndereco & ") " & vbCrLf &
                                   "       and NotasFiscais.Movimento     <='" & Ano & "/" & Mes & "/" & Dia & "' " & vbCrLf &
                                   "       AND SO.Deposito         = 'S'" & vbCrLf &
                                   "       AND year(NotasFiscais.Movimento) >= " & Ano & vbCrLf &
                                   "     GROUP BY NotasFiscais.Empresa_Id,  " & vbCrLf &
                                   "           NotasFiscais.EndEmpresa_Id, " & vbCrLf &
                                   "           NotasFiscais.Destino, " & vbCrLf &
                                   "           NotasFiscais.endDestino, " & vbCrLf &
                                   "           NotasFiscaisXItens.Produto_Id, " & vbCrLf &
                                   "           NotasFiscais.Movimento, " & vbCrLf &
                                   "           SOD.EntradaSaida " & vbCrLf


                            '--****************************
                            '--********  FIM  *************
                            '--**************************** 
                        End If


                        Sql &= "                 ) as Consulta" & vbCrLf &
                             "            Group By Empresa_Id, EndEmpresa_Id, Deposito, EndDeposito, Produto_Id" & vbCrLf &
                             "           ) AS Consulta " & vbCrLf &
                             "      INNER JOIN Clientes AS Empresa " & vbCrLf &
                             "         ON Consulta.Empresa = Empresa.Cliente_Id" & vbCrLf &
                             "        AND Consulta.EndEmpresa = Empresa.Endereco_Id " & vbCrLf

                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql &= "      INNER JOIN Clientes AS Depositos ON Consulta.Deposito = Depositos.Cliente_Id AND Consulta.EndDeposito = Depositos.Endereco_Id " & vbCrLf
                        Else
                            Sql &= "      INNER JOIN Clientes AS Depositos ON '" & objEmpresa.Codigo & "' = Depositos.Cliente_Id AND " & objEmpresa.CodigoEndereco & " = Depositos.Endereco_Id " & vbCrLf
                        End If

                        Sql &= "      INNER JOIN Produtos AS ProdutosConsulta ON Consulta.Produto = ProdutosConsulta.Produto_Id" & vbCrLf &
                               "      LEFT JOIN EstadoFisicoIA AS PrdEstadoFisico ON PrdEstadoFisico.EstadoFisicoIA_Id = ProdutosConsulta.CodigoEstadoFisico" & vbCrLf &
                               "      INNER JOIN GruposDeEstoques ON ProdutosConsulta.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf
                        If strGrupo.Length > 0 Then 'Por Grupo de Produto
                            Sql &= "      INNER JOIN ( select grupo_id, " & vbCrLf &
                                   "                         case when exists(select 1 " & vbCrLf &
                                   "                                          from produtos p " & vbCrLf &
                                   "                                          where(p.grupo = grupo_id) " & vbCrLf &
                                   "                                          and p.controlardecimais =1 ) " & vbCrLf &
                                   "                         then 1 " & vbCrLf &
                                   "                         else 0 " & vbCrLf &
                                   "                         end controlardecimais " & vbCrLf &
                                   "                   from GruposDeEstoques " & vbCrLf &
                                   "                  ) decimais  " & vbCrLf &
                                   "       ON decimais.grupo_id = GruposDeEstoques.Grupo_Id  " & vbCrLf
                        End If

                        If rdAtivo.Checked Then
                            Sql &= "         WHERE ProdutosConsulta.Situacao in(1) " & vbCrLf
                        Else
                            Sql &= "         WHERE ProdutosConsulta.Situacao not in(1) " & vbCrLf
                        End If

                        Sql &= "        AND (" & IIf(strGrupo.Length > 0, "(" & strGrupo & ") AND (", "") & " Consulta.Inicial <> 0" & vbCrLf &
                               "         or Consulta.EntradasDia <> 0" & vbCrLf &
                               "         or Consulta.EntradasMes <> 0" & vbCrLf &
                               "         or Consulta.EntradasAno <> 0" & vbCrLf &
                               "         or Consulta.ProducaoDia <> 0" & vbCrLf &
                               "         or Consulta.ProducaoMes <> 0" & vbCrLf &
                               "         or Consulta.ProducaoAno <> 0" & vbCrLf &
                               "         or Consulta.ConsumoDia  <> 0" & vbCrLf &
                               "         or Consulta.ConsumoMes  <> 0" & vbCrLf &
                               "         or Consulta.ConsumoAno  <> 0" & vbCrLf &
                               "         or Consulta.SaidasDia   <> 0" & vbCrLf &
                               "         or Consulta.SaidasMes   <> 0" & vbCrLf &
                               "         or Consulta.SaidasAno   <> 0" & IIf(strGrupo.Length > 0, ")", "") & ")" & vbCrLf

                        If CkCDeposito.Checked = False Then 'Consolidar Deposito
                            Sql &= "      GROUP BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "", "Consulta.Produto,") & " Empresa.Nome, Empresa.Cidade, " & vbCrLf &
                                   "               Empresa.Reduzido, Empresa.Estado, Depositos.Nome, Depositos.Cidade, Depositos.Estado, Depositos.Reduzido, " & IIf(strGrupo.Length > 0, "", "ProdutosConsulta.Nome,") & " " & vbCrLf &
                                   "               ProdutosConsulta.Grupo, ProdutosConsulta.CodigoEstadoFisico, GruposDeEstoques.Descricao, PrdEstadoFisico.Descricao " & vbCrLf &
                                   " " & IIf(strGrupo.Length > 0, ",CONVERT(bit,Isnull(decimais.ControlarDecimais,0))", ",Isnull(ProdutosConsulta.ControlarDecimais,0), ProdutosConsulta.EstoqueMinimo") & " " & vbCrLf

                            If rdAlfabetico.Checked Then
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, ProdutosConsulta.Nome" & vbCrLf
                            Else
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, Consulta.Deposito, Consulta.EndDeposito, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo", "Consulta.Produto") & "" & vbCrLf
                            End If
                        Else
                            Sql &= "      GROUP BY Consulta.Empresa, Consulta.EndEmpresa, " & IIf(strGrupo.Length > 0, "", "Consulta.Produto,") & " Empresa.Nome, Empresa.Cidade, " & vbCrLf &
                                   "               Empresa.Reduzido, Empresa.Estado, Depositos.Nome, Depositos.Cidade, Depositos.Estado, Depositos.Reduzido, " & IIf(strGrupo.Length > 0, "", "ProdutosConsulta.Nome,") & " " & vbCrLf &
                                   "               ProdutosConsulta.Grupo, ProdutosConsulta.CodigoEstadoFisico, GruposDeEstoques.Descricao, PrdEstadoFisico.Descricao " & vbCrLf &
                                   " " & IIf(strGrupo.Length > 0, ",CONVERT(bit,Isnull(decimais.ControlarDecimais,0))", ",Isnull(ProdutosConsulta.ControlarDecimais,0), ProdutosConsulta.EstoqueMinimo") & " " & vbCrLf

                            If rdAlfabetico.Checked Then
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, ProdutosConsulta.Nome" & vbCrLf
                            Else
                                Sql &= "      ORDER BY Consulta.Empresa, Consulta.EndEmpresa, " & IIf(strGrupo.Length > 0, "ProdutosConsulta.Grupo", "Consulta.Produto") & "" & vbCrLf
                            End If
                        End If
                    End If

                    DS = Banco.ConsultaDataSet(Sql, "PosicaoDeEstoques")

                    If DS.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
                        Exit Sub
                    End If

                    If ExcelDados Then
                        RelatorioExcelDados(Mes, Ano, objEmpresa)
                    Else
                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Nome", objEmpresa.Nome)
                        parameters.Add("Cidade", objEmpresa.Cidade & " - " & objEmpresa.CodigoEstado)
                        parameters.Add("CNPJ", "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo))
                        parameters.Add("INSCRI", "Inscr.Est.: " & objEmpresa.InscricaoEstadual)
                        parameters.Add("Periodo", "(" & FisicoFiscal & ")    " & "Atualizado até : " & Dia & "/" & Mes & "/" & Ano)

                        Funcoes.BindReport(Me.Page, DS, "Cr_PosicaoDeEstoques", IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub RelatorioExcelDados(ByVal Mes As String, ByVal Ano As String, ByVal objEmpresa As [Lib].Negocio.Cliente)
        Try


            Dim dsExcel As New DataSet

            dsExcel.Tables.Add("Dados")

            dsExcel.Tables(0).Columns.Add("Empresa")
            dsExcel.Tables(0).Columns.Add("EndEmpresa")
            dsExcel.Tables(0).Columns.Add("Deposito")
            dsExcel.Tables(0).Columns.Add("EndDeposito")
            dsExcel.Tables(0).Columns.Add("Produto")
            dsExcel.Tables(0).Columns.Add("NomeDoProduto")
            dsExcel.Tables(0).Columns.Add("EstoqueMinimo")
            dsExcel.Tables(0).Columns.Add("EstadoFisico")
            dsExcel.Tables(0).Columns.Add("EstoqueInicial")
            dsExcel.Tables(0).Columns.Add("EntradasDia")
            dsExcel.Tables(0).Columns.Add("EntradasMes")
            dsExcel.Tables(0).Columns.Add("EntradasAno")
            dsExcel.Tables(0).Columns.Add("ProducaoDia")
            dsExcel.Tables(0).Columns.Add("ProducaoMes")
            dsExcel.Tables(0).Columns.Add("ProducaoAno")
            dsExcel.Tables(0).Columns.Add("ConsumoDia")
            dsExcel.Tables(0).Columns.Add("ConsumoMes")
            dsExcel.Tables(0).Columns.Add("ConsumoAno")
            dsExcel.Tables(0).Columns.Add("SaidasDia")
            dsExcel.Tables(0).Columns.Add("SaidasMes")
            dsExcel.Tables(0).Columns.Add("SaidasAno")
            dsExcel.Tables(0).Columns.Add("Saldo")

            Dim i As Integer = 0
            For Each row In DS.Tables(0).Rows
                Dim StrValores(21) As String

                StrValores(0) = row("Empresa")
                StrValores(1) = row("EndEmpresa")
                StrValores(2) = row("Deposito")
                StrValores(3) = row("EndDeposito")
                StrValores(4) = row("Produto")
                StrValores(5) = row("NomeDoProduto")
                StrValores(6) = row("EstoqueMinimo")
                StrValores(7) = row("DescricaoEstadoFisico")
                StrValores(8) = row("Inicial")
                StrValores(9) = row("EntradasDia")
                StrValores(10) = row("EntradasMes")
                StrValores(11) = row("EntradasAno")
                StrValores(12) = row("ProducaoDia")
                StrValores(13) = row("ProducaoMes")
                StrValores(14) = row("ProducaoAno")
                StrValores(15) = row("ConsumoDia")
                StrValores(16) = row("ConsumoMes")
                StrValores(17) = row("ConsumoAno")
                StrValores(18) = row("SaidasDia")
                StrValores(19) = row("SaidasMes")
                StrValores(20) = row("SaidasAno")

                StrValores(21) = (row("Inicial") + row("EntradasAno") + row("ProducaoAno")) - row("ConsumoAno") - row("SaidasAno")

                dsExcel.Tables(0).Rows.Add(StrValores)
            Next

            Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then
                File.Delete(fileName)
            End If

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)
                    'criando planilha 
                    Dim worksheet As ExcelWorksheet

                    'criando título da planilha 
                    worksheet = package.Workbook.Worksheets.Add("Posição de Estoques")

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
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "POSIÇÃO DE ESTOQUES")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1

                    'criando linha que informa o período selecionado na página
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                    worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "ATUALIZADO ATÉ : " & txtData.Text & " (" & IIf(RadFiscal.Checked, " FISCAL", "FÍSICO") & ")")
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Merge = True
                    worksheet.Cells(String.Format("A{0}:H{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                    rowIndex += 1


                    'criando linha com o cabeçalho da planilha
                    For Each col As DataColumn In dsExcel.Tables(0).Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                        columnIndex += 1
                    Next


                    'criando auto filtro na planilha
                    worksheet.Cells("A5:H" & rowIndex).AutoFilter = True


                    'aplicando formatação nas células do cabeçalho
                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Tables(0).Columns.Count)
                        range.Style.Font.Bold = True
                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                        range.Style.Font.Color.SetColor(Color.White)
                    End Using
                    rowIndex += 1


                    ' criando conteúdo da planilha com os dados do dataset
                    For Each row As DataRow In dsExcel.Tables(0).Rows
                        columnIndex = 1
                        For Each col As DataColumn In dsExcel.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                            columnIndex += 1
                        Next

                        'formatando células valores
                        worksheet.Cells(String.Format("I{0}:V{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                        'aplicando formatação nas células do conteúdo
                        If rowIndex Mod 2 = 0 Then
                            Using range = worksheet.Cells(rowIndex, 1, rowIndex, dsExcel.Tables(0).Columns.Count)
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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkPdf.Click
        Try
            EmitirRelatorio(True, False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcel.Click
        Try
            EmitirRelatorio(False, False)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcelDados_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcelDados.Click
        Try
            EmitirRelatorio(False, True)
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
            Funcoes.Ajuda(Me.Page, "PosicaoDeEstoques")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class