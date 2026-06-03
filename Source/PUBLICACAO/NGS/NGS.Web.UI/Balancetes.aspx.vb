Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml
Imports OfficeOpenXml.Style
Imports System.Drawing

Public Class Balancetes
    Inherits BasePage

    Private i As Integer = 0
    Private DS As DataSet
    Private VModelo As Integer = 1
    Private VMoeda As Integer = 1
    Private VIsolarLotes As String = ""
    Private VIsolarCompensacao As String = ""
    Private VGrupos As String = ""
    Private VSelecao As String = ""
    Private VNivel As Integer = 5
    Private VNiveis As String = ""
    Private VNivelGP As Integer = 0
    Private VNivelDeCusto As Integer = 5

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Balancetes", "ACESSAR") Then
                ddlUnidade.Enabled = True
                ddlEmpresa.Enabled = True
                txtDataInicial.Enabled = True
                txtDataFinal.Enabled = True
                txtDataDeEmissao.Enabled = True
                txtIniciarNaFolha.Enabled = True
                lnkLimpar.Enabled = True
                lnkAjuda.Enabled = True
                Grupos.Enabled = True
                Niveis.Enabled = True
                Modelo.Enabled = True
                IsolarLotes.Enabled = True
                IsolarCompensacao.Enabled = True
                pnlCentroDeCusto.Parent.Visible = False
                pnlNivelDeProduto.Visible = False

                chkConsolidarEmpresa.Checked = True

                ucSelecaoProduto.Parent.Visible = False
                ddl.Carregar(ddlMoeda, CarregarDDL.Tabela.Moeda, "", False)
                ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                VerificaUnidade()
                CargaGrupos()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub VerificaUnidade()
        Dim sql As String = ""
        sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CarregarEmpresasComEnderecos()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            'VerEmpresas(Dr("AcessoEmpresa"), Dr("AcessoEndEmpresa"))
        Next
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidade.SelectedIndexChanged
        CarregarEmpresasComEnderecos()
    End Sub

    Private Sub CarregarEmpresasComEnderecos()
        Dim strSQL As String = " SELECT DISTINCT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
                               "   FROM GruposXEmpresas " & vbCrLf & _
                               "  INNER JOIN Clientes" & vbCrLf & _
                               "     ON GruposXEmpresas.Cliente_Id    = Clientes.Cliente_Id " & vbCrLf & _
                               "    AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                               "  Inner Join ClientesXEmpresas cxe" & vbCrLf & _
                               "     on cxe.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            strSQL &= " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "'"
        End If

        DS = Banco.ConsultaDataSet(strSQL, "Empresas")

        ddlEmpresa.Items.Clear()
        ddlEmpresa.DataTextField = "Descricao"
        ddlEmpresa.DataValueField = "Codigo"

        If DS IsNot Nothing AndAlso DS.Tables IsNot Nothing AndAlso DS.Tables.Count > 0 AndAlso DS.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In DS.Tables(0).Rows
                ddlEmpresa.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido"), Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))))
            Next
            ddlEmpresa.Items.Insert(0, New ListItem("", ""))
        End If
    End Sub

    Private Sub CargaGrupos()
        Dim sql As String = ""
        Dim Total As Integer = 0
        sql = "SELECT Left(Conta_Id, 7) As Codigo, Titulo as Descricao " & vbCrLf & _
              "FROM PlanoDeContas " & vbCrLf & _
              "WHERE Len(Conta_Id) = 7 " & vbCrLf & _
              "ORDER BY Conta_Id" & vbCrLf
        i = 0
        DS = Banco.ConsultaDataSet(sql, "Contas")
        Total = DS.Tables(0).Rows.Count

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Clientes").Tables(0).Rows
            If i <= Total / 2 Then
                SelecaoDeGrupos1.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            Else
                SelecaoDeGrupos2.Items.Add(New ListItem(Dr("Codigo") & " - " & Dr("Descricao"), Dr("Codigo")))
            End If
            i += 1
        Next
    End Sub

    Private Sub GeraBalanceteAuxiliar(ByVal ParametrosRelatorio As String, ByVal ParametrosEmpresa As ArrayList, pTipo As eTipoSaidaRelatorio)
        Try
            Dim UnidadeDeNegocio As String = ddlUnidade.SelectedValue
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim sql As String = ""
            Dim Ds_BalanceteAuxiliar As New DataSet

            Dim cont As Integer = 1

            '-----------Balancete --------------------------------------
            sql = "Select Conta as Conta_Id, PlanoDeContas.Titulo, SaldoAnterior, Debitos, Creditos, Conta" & vbCrLf & _
                  "  Into #Temp" & vbCrLf & _
                  "  From (" & vbCrLf & _
                  "        SELECT Left(r.Conta_Id, 7) As Conta," & vbCrLf

            If chkZeraContaDeResultado.Checked Then
                sql &= "               sum(CASE WHEN r.Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' and left(r.conta_Id,1) <> '3'" & vbCrLf & _
                       "                          THEN (r.DebitoOficial - r.CreditoOficial)" & vbCrLf & _
                       "                          ELSE 0 " & vbCrLf & _
                       "                    END) AS SaldoAnterior, " & vbCrLf
            Else
                sql &= "               sum(CASE WHEN r.Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "'" & vbCrLf & _
                       "                          THEN (r.DebitoOficial - r.CreditoOficial)" & vbCrLf & _
                       "                          ELSE 0 " & vbCrLf & _
                       "                    END) AS SaldoAnterior, " & vbCrLf
            End If

            sql &= "               sum(CASE WHEN r.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                          THEN r.DebitoOficial" & vbCrLf & _
                   "                          ELSE 0" & vbCrLf & _
                   "                    END) AS Debitos, " & vbCrLf & _
                   "               sum(CASE WHEN r.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                          THEN r.CreditoOficial" & vbCrLf & _
                   "                          ELSE 0" & vbCrLf & _
                   "                    END) AS Creditos " & vbCrLf & _
                   "          FROM Razao r" & vbCrLf

            If chkpiscofins56.Checked Then
                sql &= "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
                       "           AND r.Empresa_id      = nfi.Empresa_Id" & vbCrLf & _
                       "           AND r.EndEmpresa_id   = nfi.EndEmpresa_Id" & vbCrLf & _
                       "           AND r.Cliente_Nf      = nfi.Cliente_Id " & vbCrLf & _
                       "           AND r.EndCliente_Nf   = nfi.EndCliente_Id" & vbCrLf & _
                       "           AND r.EntradaSaida_Nf = nfi.EntradaSaida_Id" & vbCrLf & _
                       "           AND r.Serie_Nf        = nfi.Serie_Id" & vbCrLf & _
                       "           AND r.Numero_Nf       = nfi.Nota_Id " & vbCrLf & _
                       "           AND r.Produto_NF      = nfi.Produto_Id" & vbCrLf & _
                       "           AND r.Sequencia_NF    = nfi.Sequencia_id" & vbCrLf & _
                       "         Inner join OperacaoXEstado OE" & vbCrLf & _
                       "            on OE.Codigo_Id = nfi.OperacaoxEstado" & vbCrLf & _
                       "           AND OE.STPISCOFINS in (" & txtSituacaoPISCOFINS.Text & ")" & vbCrLf
            End If


            sql &= "         Where r.Movimento_ID <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "           And (Left(r.Conta_Id, 1) In" & VGrupos & ")" & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sql &= "           And isnull(r.UnidadeDeNegocio,'" & UnidadeDeNegocio & "') = '" & UnidadeDeNegocio & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sql &= "           And left(r.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sql &= "           And r.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sql &= "           And r.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                           "           AND r.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            If VSelecao <> "" Then sql &= "           And Left(r.Conta_Id, 7) In " & VSelecao & vbCrLf
            If VIsolarLotes <> "" Then sql &= "           And Not (year(r.movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and r.Lote_Id In " & VIsolarLotes & ")" & vbCrLf
            If VIsolarCompensacao <> "" Then sql &= " And Left(r.Conta_Id,3) Not In " & VIsolarCompensacao & vbCrLf

            sql &= "         Group  By Left(r.Conta_Id, 7)" & vbCrLf & _
                   "         Union" & vbCrLf & _
                   "        SELECT Left(r.Conta_Id, 9) As Conta, " & vbCrLf

            If chkZeraContaDeResultado.Checked Then
                sql &= "               sum(CASE WHEN r.Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' and left(r.conta_Id,1) <> '3'" & vbCrLf & _
                       "                          THEN (r.DebitoOficial - r.CreditoOficial)" & vbCrLf & _
                       "                          ELSE 0 " & vbCrLf & _
                       "                    END) AS SaldoAnterior, " & vbCrLf
            Else
                sql &= "               sum(CASE WHEN r.Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "'" & vbCrLf & _
                       "                          THEN (r.DebitoOficial - r.CreditoOficial)" & vbCrLf & _
                       "                          ELSE 0 " & vbCrLf & _
                       "                    END) AS SaldoAnterior, " & vbCrLf
            End If

            sql &= "               sum(CASE WHEN r.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                          THEN r.DebitoOficial" & vbCrLf & _
                   "                          ELSE 0" & vbCrLf & _
                   "                    END) AS Debitos, " & vbCrLf & _
                   "               sum(CASE WHEN r.Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "                          THEN r.CreditoOficial" & vbCrLf & _
                   "                          ELSE 0" & vbCrLf & _
                   "                    END) AS Creditos " & vbCrLf & _
                   "          FROM Razao r" & vbCrLf

            If chkpiscofins56.Checked Then
                sql &= "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
                       "           AND r.Empresa_id      = nfi.Empresa_Id" & vbCrLf & _
                       "           AND r.EndEmpresa_id   = nfi.EndEmpresa_Id" & vbCrLf & _
                       "           AND r.Cliente_Nf      = nfi.Cliente_Id " & vbCrLf & _
                       "           AND r.EndCliente_Nf   = nfi.EndCliente_Id" & vbCrLf & _
                       "           AND r.EntradaSaida_Nf = nfi.EntradaSaida_Id" & vbCrLf & _
                       "           AND r.Serie_Nf        = nfi.Serie_Id" & vbCrLf & _
                       "           AND r.Numero_Nf       = nfi.Nota_Id " & vbCrLf & _
                       "           AND r.Produto_NF      = nfi.Produto_Id" & vbCrLf & _
                       "           AND r.Sequencia_NF    = nfi.Sequencia_id" & vbCrLf & _
                       "         Inner join OperacaoXEstado OE" & vbCrLf & _
                       "            on OE.Codigo_Id = nfi.OperacaoxEstado" & vbCrLf & _
                       "           AND OE.STPISCOFINS in (" & txtSituacaoPISCOFINS.Text & ")" & vbCrLf
            End If

            sql &= "         where r.Movimento_Id <='" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "           AND Len(r.Conta_Id) = 9" & vbCrLf & _
                   "           AND (Left(r.Conta_Id, 1) In" & VGrupos & ")" & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sql &= "           And isnull(r.UnidadeDeNegocio,'" & UnidadeDeNegocio & "') = '" & UnidadeDeNegocio & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sql &= "           And left(r.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sql &= "           And r.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sql &= "           And r.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                           "           AND r.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            If VSelecao <> "" Then sql &= " And Left(r.Conta_Id, 7) In " & VSelecao & vbCrLf
            If VIsolarLotes <> "" Then sql &= " And Not (year(r.movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and r.Lote_Id  In " & VIsolarLotes & ")" & vbCrLf
            If VIsolarCompensacao <> "" Then sql &= " And Left(r.Conta_Id,3) Not In " & VIsolarCompensacao & vbCrLf

            sql &= "         Group  By Left(r.Conta_Id, 9)" & vbCrLf & _
                   "      ) AS Consulta " & vbCrLf & _
                   "  INNER JOIN PlanoDeContas" & vbCrLf & _
                   "     ON Replace(Consulta.Conta,'987654321', '') = PlanoDeContas.Conta_Id" & vbCrLf & _
                   "  Where Titulo <> ''" & vbCrLf

            If rdTodasAsContas.Checked Then
                sql &= "    And (Debitos <> 0 or Creditos <> 0 or SaldoAnterior <> 0) " & vbCrLf
            ElseIf rdComMovimentoNoPeriodo.Checked Then
                sql &= " and (Debitos <> 0 or Creditos <> 0)" & vbCrLf
            ElseIf rdSomenteComSaldoAtual.Checked Then
                sql &= " and (SaldoAnterior + abs(Debitos) - abs(Creditos) <> 0)" & vbCrLf
            End If

            sql &= "	Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   "		 (" & vbCrLf & _
                   "		   SELECT LEFT(T.Conta_Id, 5)," & vbCrLf & _
                   "				  PC.Titulo," & vbCrLf & _
                   "				  SUM(T.SaldoAnterior) AS SaldoAnterior," & vbCrLf & _
                   "				  SUM(T.Debitos) AS Debitos," & vbCrLf & _
                   "				  SUM(T.Creditos) AS Creditos," & vbCrLf & _
                   "				  LEFT(T.Conta_Id, 5)" & vbCrLf & _
                   "			 FROM #Temp T" & vbCrLf & _
                   "			INNER JOIN PlanoDeContas PC" & vbCrLf & _
                   "			   ON LEFT(T.Conta_Id, 5) = PC.Conta_Id" & vbCrLf & _
                   "			WHERE LEN(T.Conta_Id) = 7" & vbCrLf & _
                   "			GROUP BY LEFT(T.Conta_Id, 5)," & vbCrLf & _
                   "					 PC.Titulo" & vbCrLf & _
                   "		 )" & vbCrLf & _
                   "	Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   "		 (" & vbCrLf & _
                   "		   SELECT LEFT(T.Conta_Id, 3)," & vbCrLf & _
                   "				  PC.Titulo," & vbCrLf & _
                   "				  SUM(T.SaldoAnterior) AS SaldoAnterior," & vbCrLf & _
                   "				  SUM(T.Debitos) AS Debitos," & vbCrLf & _
                   "				  SUM(T.Creditos) AS Creditos," & vbCrLf & _
                   "				  LEFT(T.Conta_Id, 3)" & vbCrLf & _
                   "			 FROM #Temp T" & vbCrLf & _
                   "			INNER JOIN PlanoDeContas PC" & vbCrLf & _
                   "			   ON LEFT(T.Conta_Id, 3) = PC.Conta_Id" & vbCrLf & _
                   "			WHERE LEN(T.Conta_Id) = 5" & vbCrLf & _
                   "			GROUP BY LEFT(T.Conta_Id, 3)," & vbCrLf & _
                   "					 PC.Titulo" & vbCrLf & _
                   "		 )" & vbCrLf & _
                   "	Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   "		 (" & vbCrLf & _
                   "		   SELECT LEFT(T.Conta_Id, 1), " & vbCrLf & _
                   "				  PC.Titulo," & vbCrLf & _
                   "				  SUM(T.SaldoAnterior) AS SaldoAnterior, " & vbCrLf & _
                   "				  SUM(T.Debitos) AS Debitos," & vbCrLf & _
                   "				  SUM(T.Creditos) AS Creditos," & vbCrLf & _
                   "				  LEFT(T.Conta_Id, 1)" & vbCrLf & _
                   "			 FROM #Temp T" & vbCrLf & _
                   "			INNER JOIN PlanoDeContas PC" & vbCrLf & _
                   "			   ON LEFT(T.Conta_Id, 1) = PC.Conta_Id" & vbCrLf & _
                   "			WHERE LEN(T.Conta_Id) = 3" & vbCrLf & _
                   "			GROUP BY LEFT(T.Conta_Id, 1)," & vbCrLf & _
                   "					 PC.Titulo" & vbCrLf & _
                   "		 )" & vbCrLf

            '---PlanoDeContas Apuracao--------------------------------------------------
            If String.IsNullOrWhiteSpace(VNiveis) Then VNiveis = "9"

            sql &= "	select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta" & vbCrLf & _
                   "	  From #temp" & vbCrLf & _
                   "     Where len(Conta_Id) in (" & VNiveis & ")" & vbCrLf & _
                   "	 order by conta" & vbCrLf

            '---Totalizadores----------------------------------------------------------
            sql &= "  select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos " & vbCrLf & _
                   "    from #temp " & vbCrLf & _
                   "   where LEN(Conta_Id) = 1" & vbCrLf & _
                   "  order by conta_Id" & vbCrLf

            Ds_BalanceteAuxiliar = Banco.ConsultaDataSet(sql, "Consulta")
            Ds_BalanceteAuxiliar.Tables("Consulta").TableName = "PlanoDeContas"
            Ds_BalanceteAuxiliar.Tables("Consulta1").TableName = "Totais"

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", "Relatório De Balancete Auxiliar.")
            parameters.Add("Parametros", ParametrosRelatorio)
            parameters.Add("IniciarPagina", txtIniciarNaFolha.Text)

            Funcoes.BindReport(Me.Page, Ds_BalanceteAuxiliar, "Cr_BalanceteAuxiliar", IIf(pTipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(pTipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getDataSetCompleto() As DataSet
        Try
            Dim sql As String = ""
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")

            '-----------Balancete --------------------------------------
            sql = "Select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta_Id as Conta " & vbCrLf & _
                  "  into #Temp" & vbCrLf & _
                  "  From (" & vbCrLf & _
                  "        SELECT LEFT(Razao.Conta_Id, 7) AS Conta," & vbCrLf & _
                  "              ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo," & vbCrLf

            If chkZeraContaDeResultado.Checked Then
                sql &= "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' and left(Razao.Conta_id,1) <> '3' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf
            Else
                sql &= "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf
            End If

            sql &= "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                   "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                   "               LEFT(Razao.Conta_Id, 7) AS Conta_Id" & vbCrLf & _
                   "          FROM Razao " & vbCrLf

            If chkpiscofins56.Checked Then
                sql &= "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
                       "            ON Razao.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                       "           AND Razao.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                       "           AND Razao.Cliente_Nf      = nfi.Cliente_Id " & vbCrLf & _
                       "           AND Razao.EndCliente_Nf   = nfi.EndCliente_Id" & vbCrLf & _
                       "           AND Razao.EntradaSaida_Nf = nfi.EntradaSaida_Id" & vbCrLf & _
                       "           AND Razao.Serie_Nf        = nfi.Serie_Id" & vbCrLf & _
                       "           AND Razao.Numero_Nf       = nfi.Nota_Id " & vbCrLf & _
                       "           AND Razao.Produto_NF      = nfi.Produto_Id" & vbCrLf & _
                       "           AND Razao.Sequencia_NF    = nfi.Sequencia_id" & vbCrLf & _
                       "         Inner join OperacaoXEstado OE" & vbCrLf & _
                       "            on OE.Codigo_Id = nfi.OperacaoxEstado" & vbCrLf & _
                       "           AND OE.STPISCOFINS in (" & txtSituacaoPISCOFINS.Text & ")" & vbCrLf
            End If

            sql &= "         INNER JOIN PlanoDeContas " & vbCrLf & _
                   "            ON LEFT(Razao.Conta_Id, 7) = PlanoDeContas.Conta_Id" & vbCrLf & _
                   "         where Movimento_Id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                   "           And (Left(Razao.Conta_Id, 1) In" & VGrupos & ")" & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked AndAlso Not chkConsolidarUnidade.Checked Then
                sql &= "           And isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'"
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sql &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                           "           AND Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            If VSelecao <> "" Then sql &= "           And Left(Razao.Conta_Id, 7) In " & VSelecao & vbCrLf
            If VIsolarLotes <> "" Then sql &= "           And (year(movimento_Id) < " & CDate(txtDataInicial.Text).Year & " or (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & "))" & vbCrLf
            If VIsolarCompensacao <> "" Then sql &= "           And Left(Razao.Conta_Id,3) Not In " & VIsolarCompensacao & vbCrLf
            sql &= "         Group  By LEFT(Razao.Conta_Id, 7), PlanoDeContas.Titulo" & vbCrLf


            If VNivel = 5 Then
                sql &= "        Union" & vbCrLf & _
                       "        SELECT CASE WHEN LEN(Razao.Conta_Id) = 9" & vbCrLf & _
                       "                       THEN Razao.Conta_Id" & vbCrLf & _
                       "                       ELSE LEFT(Razao.Conta_Id, 9) + '-' + Razao.Cliente_Id + '-' + convert(varchar, Razao.EndCliente_Id)" & vbCrLf & _
                       "               END AS Conta," & vbCrLf & _
                       " 	           CASE WHEN LEN(Razao.Conta_Id) = 9" & vbCrLf & _
                       "                      THEN ('AAA' +  PlanoDeContas.Titulo + '# ')" & vbCrLf & _
                       "                      Else 'ZZZ' + Left(Clientes.Nome, 35) + ' - ' + left(Clientes.Cidade, 18) + '#' + Razao.Cliente_Id + '-' + CONVERT(varchar, Razao.EndCliente_Id)" & vbCrLf & _
                       "               END AS Titulo," & vbCrLf

                If chkZeraContaDeResultado.Checked Then
                    sql &= "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' and left(Razao.Conta_id,1) <> '3' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf
                Else
                    sql &= "               sum(CASE WHEN Movimento_ID < '" & txtDataInicial.Text.ToSqlDate() & "' THEN (DebitoOficial - CreditoOficial) ELSE 0 END) AS SaldoAnterior, " & vbCrLf
                End If

                sql &= "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN DebitoOficial ELSE 0 END) AS Debitos, " & vbCrLf & _
                       "               sum(CASE WHEN Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN CreditoOficial ELSE 0 END) AS Creditos, " & vbCrLf & _
                       "               Razao.Conta_Id" & vbCrLf & _
                       "          FROM Razao" & vbCrLf

                If chkpiscofins56.Checked Then
                    sql &= "         INNER JOIN NotasFiscaisXItens nfi" & vbCrLf & _
                           "            ON Razao.Empresa_Id      = nfi.Empresa_Id" & vbCrLf & _
                           "           AND Razao.EndEmpresa_Id   = nfi.EndEmpresa_Id" & vbCrLf & _
                           "           AND Razao.Cliente_Nf      = nfi.Cliente_Id " & vbCrLf & _
                           "           AND Razao.EndCliente_Nf   = nfi.EndCliente_Id" & vbCrLf & _
                           "           AND Razao.EntradaSaida_Nf = nfi.EntradaSaida_Id" & vbCrLf & _
                           "           AND Razao.Serie_Nf        = nfi.Serie_Id" & vbCrLf & _
                           "           AND Razao.Numero_Nf       = nfi.Nota_Id " & vbCrLf & _
                           "           AND Razao.Produto_NF      = nfi.Produto_Id" & vbCrLf & _
                           "           AND Razao.Sequencia_NF    = nfi.Sequencia_id" & vbCrLf & _
                           "         Inner join OperacaoXEstado OE" & vbCrLf & _
                           "            on OE.Codigo_Id = nfi.OperacaoxEstado" & vbCrLf & _
                           "           AND OE.STPISCOFINS in (" & txtSituacaoPISCOFINS.Text & ")" & vbCrLf
                End If

                sql &= "          LEFT JOIN PlanoDeContas" & vbCrLf & _
                       "            ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                       "          LEFT JOIN Clientes" & vbCrLf & _
                       "            ON Razao.Cliente_Id = Clientes.Cliente_Id " & vbCrLf & _
                       "           AND Razao.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                       "         where (Left(Razao.Conta_Id, 1) In" & VGrupos & ") " & vbCrLf

                If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                    sql &= "           And isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'"
                End If

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    If chkConsolidarEmpresa.Checked Then
                        sql &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                    ElseIf chkConsolidarSede.Checked Then
                        sql &= "           And Razao.Empresa_Id    = '" & Emp(0) & "'" & vbCrLf
                    Else
                        sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                               "           AND Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                    End If
                End If


                If VSelecao <> "" Then sql &= " And Left(Razao.Conta_Id, 7) In " & VSelecao & vbCrLf
                If VIsolarLotes <> "" Then sql &= "           And (year(movimento_Id) < " & CDate(txtDataInicial.Text).Year & " or (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & "))" & vbCrLf
                If VIsolarCompensacao <> "" Then sql &= " And Left(Razao.Conta_Id,3) Not In " & VIsolarCompensacao & vbCrLf

                sql &= "         Group BY CASE WHEN LEN(Razao.Conta_Id) = 9 THEN Razao.Conta_Id ELSE LEFT(Razao.Conta_Id, 9) + '-' + Razao.Cliente_Id + '-' + convert(varchar, Razao.EndCliente_Id) END," & vbCrLf & _
                       "                  CASE WHEN LEN(Razao.Conta_Id) = 9 THEN ('AAA' +  PlanoDeContas.Titulo + '# ') ELSE 'ZZZ' + Left(Clientes.Nome, 35) + ' - ' + left(Clientes.Cidade, 18) + '#' + Razao.Cliente_Id + '-' + CONVERT(varchar, Razao.EndCliente_Id) END, Razao.Conta_Id" & vbCrLf
            End If

            sql &= "       ) AS   Consulta " & vbCrLf & _
                   "  Where  Titulo <> ''" & vbCrLf

            If rdTodasAsContas.Checked Then
                sql &= "    And (Debitos <> 0 or Creditos <> 0 or SaldoAnterior <> 0) " & vbCrLf
            ElseIf rdComMovimentoNoPeriodo.Checked Then
                sql &= " and (Debitos <> 0 or Creditos <> 0)" & vbCrLf
            ElseIf rdSomenteComSaldoAtual.Checked Then
                sql &= " and (SaldoAnterior + abs(Debitos) - abs(Creditos) <> 0)" & vbCrLf
            End If

            sql &= " Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   " (" & vbCrLf & _
                   "        SELECT LEFT(#Temp.Conta_Id, 5) AS Conta_Id," & vbCrLf & _
                   "              ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo," & vbCrLf & _
                   "               sum(#Temp.SaldoAnterior) as SaldoAnterior," & vbCrLf & _
                   "               sum(#Temp.Debitos) AS Debitos," & vbCrLf & _
                   "               sum(#Temp.Creditos) AS Creditos," & vbCrLf & _
                   "               LEFT(#Temp.Conta_Id, 5) AS Conta" & vbCrLf & _
                   "          FROM #Temp" & vbCrLf & _
                   "         INNER JOIN PlanoDeContas" & vbCrLf & _
                   "            ON LEFT(#Temp.Conta_Id, 5) = PlanoDeContas.Conta_Id" & vbCrLf & _
                   "         where Len(#Temp.Conta_Id) = 7" & vbCrLf & _
                   "           and Left(#Temp.Titulo,3) = 'AAA' " & vbCrLf & _
                   "         Group By LEFT(#Temp.Conta_Id, 5), PlanoDeContas.Titulo" & vbCrLf & _
                   " )" & vbCrLf & _
                   " Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   " (" & vbCrLf & _
                   "        SELECT LEFT(#Temp.Conta_Id, 3) AS Conta," & vbCrLf & _
                   "              ('AAA' + PlanoDeContas.Titulo + '# ') as Titulo," & vbCrLf & _
                   "               sum(#Temp.SaldoAnterior) as SaldoAnterior," & vbCrLf & _
                   "               sum(#Temp.Debitos) AS Debitos," & vbCrLf & _
                   "               sum(#Temp.Creditos) AS Creditos," & vbCrLf & _
                   "               LEFT(#Temp.Conta_Id, 3) AS Conta" & vbCrLf & _
                   "          FROM #Temp" & vbCrLf & _
                   "         INNER JOIN PlanoDeContas" & vbCrLf & _
                   "            ON LEFT(#Temp.Conta_Id, 3) = PlanoDeContas.Conta_Id" & vbCrLf & _
                   "         where Len(#Temp.Conta_Id) =  5" & vbCrLf & _
                   "         Group By LEFT(#Temp.Conta_Id, 3), PlanoDeContas.Titulo" & vbCrLf & _
                   " )" & vbCrLf & _
                   " Insert into #Temp(Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, Conta)" & vbCrLf & _
                   " (" & vbCrLf & _
                   "        SELECT LEFT(#Temp.Conta_Id, 1) AS Conta," & vbCrLf & _
                   "               ('AAA' +  PlanoDeContas.Titulo + '# ') as Titulo ," & vbCrLf & _
                   "               sum(#Temp.SaldoAnterior) as SaldoAnterior," & vbCrLf & _
                   "               sum(#Temp.Debitos) AS Debitos," & vbCrLf & _
                   "               sum(#Temp.Creditos) AS Creditos," & vbCrLf & _
                   "               LEFT(#Temp.Conta_Id, 1) AS Conta" & vbCrLf & _
                   "          FROM #Temp" & vbCrLf & _
                   "         INNER JOIN PlanoDeContas" & vbCrLf & _
                   "            ON LEFT(#Temp.Conta_Id, 1) = PlanoDeContas.Conta_Id" & vbCrLf & _
                   "         where Len(#Temp.Conta_Id) =  3" & vbCrLf & _
                   "         Group By LEFT(#Temp.Conta_Id, 1), PlanoDeContas.Titulo" & vbCrLf & _
                   " )" & vbCrLf


            '---Plano De Contas Apuracao -----------------------------------------------
            If String.IsNullOrWhiteSpace(VNiveis) Then VNiveis = "9"

            sql &= " SELECT Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos, (SaldoAnterior + Debitos - Creditos) as SaldoFinal, Conta" & vbCrLf & _
                   "   FROM #Temp" & vbCrLf & _
                   "  Where len(Conta_Id) in (" & VNiveis & ")" & vbCrLf & _
                   "  Order by Conta, Titulo" & vbCrLf

            '---Totalizadores----------------------------------------------------------
            sql &= "  select Conta_Id, Titulo, SaldoAnterior, Debitos, Creditos " & vbCrLf & _
                   "    from #temp " & vbCrLf & _
                   "   where LEN(Conta_Id) = 1" & vbCrLf & _
                   "  order by conta_Id" & vbCrLf

            Dim Ds_BalanceteAuxiliar As DataSet = Banco.ConsultaDataSet(sql, "Consulta")
            Ds_BalanceteAuxiliar.Tables("Consulta").TableName = "PlanoDeContas"
            Ds_BalanceteAuxiliar.Tables("Consulta1").TableName = "Totais"

            Return Ds_BalanceteAuxiliar
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub GeraBalanceteAuxiliarCompleto(ByVal ParametrosRelatorio As String, ByVal ParametrosEmpresa As ArrayList, ptipo As eTipoSaidaRelatorio)
        Try
            Dim ContaAnterior As String = ""
            Dim DS As DataSet = getDataSetCompleto()

            If ptipo = eTipoSaidaRelatorio.PDF OrElse ptipo = eTipoSaidaRelatorio.EXCEL Then
                Dim param As New Dictionary(Of String, Object)
                param.Add("Parametros", ParametrosRelatorio)
                param.Add("Titulo", "Relatório Auxiliar Completo.")
                param.Add("IniciarPagina", txtIniciarNaFolha.Text)
                'Ds_BalanceteAuxiliar
                Funcoes.BindReport(Me.Page, DS, "Cr_BalanceteCompleto", IIf(ptipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, eExportType.ExcelCrystal), param)
            Else

                'Emitir Excel.xsls do office/ Relatorio padrao em lista.
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1
                Try
                    Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                    If File.Exists(fileName) Then File.Delete(fileName)

                    Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                        Using package As New ExcelPackage(arquivo)

                            'criando aba da planilha.
                            Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Balancete")

                            'Inserindo o Cabeçalho.
                            Dim objEmpresa As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                            worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Nome.ToUpper
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            rowIndex += 1
                            columnIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Cidade.ToUpper & "/" & objEmpresa.Estado.Descricao.ToUpper
                            rowIndex += 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & "-" & objEmpresa.Endereco
                            rowIndex += 2

                            'Inserindo o Título
                            worksheet.Cells(rowIndex, columnIndex).Value = "Balancete Auxiliar"
                            worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                            rowIndex += 2

                            'Inserindo informacoes adicionais
                            columnIndex = 1
                            worksheet.Cells(rowIndex, columnIndex).Value = "Período: " & txtDataInicial.Text & " À " & txtDataFinal.Text
                            columnIndex = DS.Tables(0).Columns.Count - 2
                            worksheet.Cells(rowIndex, columnIndex).Value = "Emissão: " & DateTime.Now
                            rowIndex += 1

                            'Inserindo as Colunas.
                            columnIndex = 1
                            For Each col As DataColumn In DS.Tables(0).Columns
                                If col.ColumnName <> "Conta" Then
                                    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                                Else
                                    If chkTitulos.Checked Then
                                        columnIndex += 1
                                        worksheet.Cells(rowIndex, columnIndex).Value = "Saldo Titulos"
                                        columnIndex += 1
                                        worksheet.Cells(rowIndex, columnIndex).Value = "Diferença"
                                    End If
                                End If
                                columnIndex += 1
                            Next

                            rowIndex += 1

                            'aplicando formatação nas células das Colunas
                            Using range = worksheet.Cells(rowIndex - 1, 1, rowIndex - 1, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 2, DS.Tables(0).Columns.Count - 1))
                                range.Style.Font.Bold = True
                                range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                                range.Style.Font.Color.SetColor(Color.White)
                            End Using

                            ' Exportando conteúdo da planilha com os dados da Tabela.
                            For Each row As DataRow In DS.Tables(0).Rows
                                columnIndex = 1
                                For Each col As DataColumn In DS.Tables(0).Columns
                                    If ContaAnterior <> "" Then
                                        If Len(row("Conta_Id")) < 9 AndAlso Len(ContaAnterior) > Len(row("Conta_Id")) Then
                                            rowIndex += 1
                                        End If
                                    ElseIf Len(row("Conta_Id")) = 7 AndAlso Len(ContaAnterior) = 7 Then
                                        If row("Conta_Id") <> ContaAnterior Then
                                            rowIndex += 1
                                        End If
                                    ElseIf Len(row("Conta_Id")) = 5 AndAlso Len(ContaAnterior) = 5 Then
                                        If row("Conta_Id") <> ContaAnterior Then
                                            rowIndex += 1
                                        End If
                                    End If

                                    If Len(row("Conta_Id")) < 8 AndAlso row("Conta_Id") <> ContaAnterior Then
                                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 2, DS.Tables(0).Columns.Count - 1))
                                            range.Style.Font.Bold = True
                                        End Using
                                    End If

                                    If col.ColumnName <> "Conta" Then
                                        If col.ColumnName = "Titulo" Then
                                            worksheet.Cells(rowIndex, columnIndex).Value = Replace(Replace(Replace(row(col.ColumnName), "AAA", ""), "#", " "), "ZZZ", "")
                                        Else
                                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        End If
                                    Else
                                        columnIndex += 1

                                        Dim cliente As String = row("Titulo").ToString.Split("#")(1)

                                        If Not String.IsNullOrWhiteSpace(cliente) AndAlso chkTitulos.Checked Then
                                            If Len(row("Conta_Id")) = 7 Then
                                                If Len(cliente) > 4 Then
                                                    Select Case row("Conta_Id")
                                                        Case "1010201", "1010202", "2010101", "2010104", "2010107", "2010109", "2010111"
                                                            Dim Conta As String = row("Conta_Id")
                                                            Dim saldo As Decimal
                                                            Dim SaldoReferencial As Decimal

                                                            SaldoDeTitulos(ddlEmpresa.SelectedValue.Split("-")(0), cliente, Conta, saldo)

                                                            If saldo < 0 Then
                                                                saldo = saldo * -1
                                                            End If

                                                            If row("SaldoFinal") < 0 Then
                                                                SaldoReferencial = row("SaldoFinal") * -1
                                                            Else
                                                                SaldoReferencial = row("SaldoFinal")
                                                            End If

                                                            worksheet.Cells(rowIndex, columnIndex).Value = saldo
                                                            columnIndex += 1
                                                            worksheet.Cells(rowIndex, columnIndex).Value = SaldoReferencial - saldo
                                                    End Select
                                                End If
                                            End If
                                        End If
                                    End If

                                    ContaAnterior = row("Conta_Id")
                                    columnIndex += 1

                                Next

                                worksheet.Cells(rowIndex, 3, rowIndex, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 2, DS.Tables(0).Columns.Count - 1)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                'Formatacoes de celulas
                                If rowIndex - 1 Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex - 1, 1, rowIndex - 1, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 2, DS.Tables(0).Columns.Count - 1))
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If

                                If rowIndex Mod 2 = 0 Then
                                    Using range = worksheet.Cells(rowIndex, 1, rowIndex, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 2, DS.Tables(0).Columns.Count - 1))
                                        range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                        range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                    End Using
                                End If
                                rowIndex += 1
                            Next

                            'adicionando o resumo
                            rowIndex += 2
                            Dim linha As Integer = rowIndex

                            For Each row As DataRow In DS.Tables(1).Rows
                                columnIndex = 1
                                For Each col As DataColumn In DS.Tables(1).Columns
                                    If col.ColumnName <> "Conta_Id" Then
                                        If col.ColumnName = "Titulo" Then
                                            worksheet.Cells(rowIndex, columnIndex).Value = Replace(Replace(Replace(row(col.ColumnName), "AAA", ""), "#", " "), "ZZZ", "")
                                        ElseIf IsNumeric(row(col.ColumnName)) AndAlso Not row(col.ColumnName).ToString.Contains(",") Then
                                            worksheet.Cells(rowIndex, columnIndex).Value = Convert.ToInt64(row(col.ColumnName))
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "0"
                                        ElseIf IsNumeric(row(col.ColumnName)) AndAlso row(col.ColumnName).ToString.Contains(",") Then
                                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                            worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                                        Else
                                            worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                        End If
                                    End If
                                    columnIndex += 1
                                Next
                                worksheet.Cells(rowIndex, columnIndex).Value = row("SaldoAnterior") + row("Debitos") - row("Creditos")
                                worksheet.Cells(rowIndex, columnIndex).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                                worksheet.Cells(rowIndex, columnIndex - 4, rowIndex, columnIndex).Style.Font.Bold = True
                                rowIndex += 1
                            Next

                            worksheet.Cells(rowIndex, 2).Value = "Total"
                            worksheet.Cells(rowIndex, 3).Formula = String.Format("SUM({0}:{1})", worksheet.Cells(linha, 3).Address, worksheet.Cells(rowIndex - 1, 3).Address)
                            worksheet.Cells(rowIndex, 4).Formula = String.Format("SUM({0}:{1})", worksheet.Cells(linha, 4).Address, worksheet.Cells(rowIndex - 1, 4).Address)
                            worksheet.Cells(rowIndex, 5).Formula = String.Format("SUM({0}:{1})", worksheet.Cells(linha, 5).Address, worksheet.Cells(rowIndex - 1, 5).Address)
                            worksheet.Cells(rowIndex, 6).Formula = String.Format("SUM({0}:{1})", worksheet.Cells(linha, 6).Address, worksheet.Cells(rowIndex - 1, 6).Address)

                            worksheet.Cells(rowIndex, 3, rowIndex, 6).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'criando auto filtro na planilha
                            worksheet.Cells(8, 1, 8, IIf(chkTitulos.Checked, DS.Tables(0).Columns.Count + 1, DS.Tables(0).Columns.Count - 1)).AutoFilter = True

                            'setando autofit nas células da planilha
                            worksheet.Cells.AutoFitColumns(0)

                            'congelando primeira linham
                            worksheet.View.FreezePanes(9, 1)

                            worksheet.Column(1).Width = 20

                            'salvando planilha do excel
                            package.Save()
                        End Using
                    End Using

                    'download do arquivo pelo browser
                    Funcoes.AbrirExcel(Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
                Catch ex As Exception
                    Throw New Exception(ex.Message & " - linha: " & rowIndex & " - coluna: " & columnIndex)
                End Try
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getListCentroDeCusto() As List(Of String)
        Try
            Dim lst As New List(Of String)

            For Each item As String In lstCentroDeCusto.GetSelectedValues()
                lst.Add(item)
            Next
            Return lst
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Private Sub SaldoDeTitulos(ByVal Empresa As String, ByVal Cliente As String, ByVal Conta As String, ByRef Saldo As Decimal)
        Try
            If IsNumeric(Cliente.Split("-")(0)) Then
                Dim sql As String = ""

                If FinanceiroNovo Then
                    sql &= "   Select isnull(Sum(Apagar - AReceber), 0) as Saldo From (                         " & vbCrLf & _
                            "   Select SUM(tc.ValorOficial) as APagar, 0 as AReceber                             " & vbCrLf & _
                            "      from Titulos t                                                                " & vbCrLf & _
                            "     Inner Join TitulosxContaContabil tc                                            " & vbCrLf & _
                            "        on tc.Titulo_Id = t.Titulo_Id                                               " & vbCrLf & _
                            "   	and tc.Conta_Id = t.ContaContabilCliFor                                      " & vbCrLf & _
                            "   Where t.RecPag = 'P'                                                             " & vbCrLf & _
                            "   	And t.Empresa = '" & Empresa & "'" & vbCrLf & _
                            "   	And t.CliFor = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                            "   	And t.EnderecoCliFor = " & Cliente.Split("-")(1) & vbCrLf & _
                            "   	and t.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   	And T.ContaContabilCliFor = '" & Conta & "'" & vbCrLf & _
                            "   	And (t.Provisao = 3 or (t.Provisao = 1 And t.DataBaixa > '" & txtDataFinal.Text.ToSqlDate() & "'))      " & vbCrLf & _
                            "   	And t.Situacao = 1                                                           " & vbCrLf & _
                            "   UNION ALL                                                                        " & vbCrLf & _
                            "   Select 0 as APagar, SUM(tc.ValorOficial) as AReceber                             " & vbCrLf & _
                            "      from Titulos t                                                                " & vbCrLf & _
                            "     Inner Join TitulosxContaContabil tc                                            " & vbCrLf & _
                            "        on tc.Titulo_Id = t.Titulo_Id                                               " & vbCrLf & _
                            "   	and tc.Conta_Id = t.ContaContabilCliFor                                      " & vbCrLf & _
                            "   Where t.RecPag = 'R'" & vbCrLf & _
                            "   	And t.Empresa = '" & Empresa & "'" & vbCrLf & _
                            "   	And t.CliFor = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                            "   	And t.EnderecoCliFor = " & Cliente.Split("-")(1) & vbCrLf & _
                            "   	and t.Movimento <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   	And T.ContaContabilCliFor = '" & Conta & "'" & vbCrLf & _
                            "   	And (t.Provisao = 3 or (t.Provisao = 1 And t.DataBaixa > '" & txtDataFinal.Text.ToSqlDate() & "'))      " & vbCrLf & _
                            "   	And t.Situacao = 1" & vbCrLf & _
                            "   ) as Consulta	      " & vbCrLf
                Else
                    sql = " Select isnull(Sum(Apagar - AReceber), 0) as Saldo From (" & vbCrLf & _
                          " select  Sum(ValorDoDocumento) as APagar, 0 as AReceber from ContasAPagar" & vbCrLf & _
                          "        where Empresa = '" & Empresa & "'" & vbCrLf & _
                          "         And Cliente = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                          "         And EndCliente = " & Cliente.Split("-")(1) & vbCrLf & _
                          "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                          "         And ContaContabilCliente = '" & Conta & "'" & vbCrLf & _
                          "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                          "         And Situacao = 1" & vbCrLf & _
                          " Union" & vbCrLf & _
                          " select 0 as APagar, Sum(ValorDoDocumento) as AReceber from ContasAReceber" & vbCrLf & _
                          "        where Empresa = '" & Empresa & "'" & vbCrLf & _
                          "         And Cliente = '" & Cliente.Split("-")(0) & "'" & vbCrLf & _
                          "         And EndCliente = " & Cliente.Split("-")(1) & vbCrLf & _
                          "         And UsuarioInclusaoData <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                          "         And ContaContabilCliente = '" & Conta & "'" & vbCrLf & _
                          "         And (Provisao = 2 or (Provisao = 1 And Baixa > '" & txtDataFinal.Text.ToSqlDate() & "'))" & vbCrLf & _
                          "         And Situacao = 1" & vbCrLf & _
                          "         ) as Consulta" & vbCrLf
                End If

                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "SaldoTitulo")

                For Each row As DataRow In ds.Tables(0).Rows
                    Saldo = row("Saldo")
                Next

            End If
        Catch ex As Exception

        End Try

    End Sub

    Private Sub GeraCentrosDeCustos(ByVal ParametrosDoRelatorio As String, ByVal ParametrosEmpresa As ArrayList, pTipo As eTipoSaidaRelatorio)
        Try
            Dim DS_Anterior As New DataSet
            Dim DS_DebitoCredito As New DataSet
            Dim DS_balancete As New DataSet
            Dim sqldebitocredito As String
            Dim Ds_BalanceCentroDeCusto As New DataSet

            'Dim Ds_Fechamento As New DataSet
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")

            Dim grupos As String = String.Empty
            Dim grup As String = String.Empty

            For i As Integer = 0 To lstCentroCustoSelecionados.Items.Count - 1
                Select Case lstCentroCustoSelecionados.Items(i).Value
                    Case 0 : grupos &= " Safra as Grupo" & i + 1 & ","
                    Case 1
                        grupos &= " UnidadeDeNegocio + ' - ' + NomeUnidade as Grupo" & i + 1 & ","
                        grup &= " UnidadeDeNegocio + ' - ' + NomeUnidade,"
                    Case 2
                        grupos &= " Empresa_Id + convert(char,EndEmpresa_id) + ' - ' + NomeEmpresa as Grupo" & i + 1 & ","
                        grup &= " Empresa_Id + convert(char,EndEmpresa_id) + ' - ' + NomeEmpresa,"
                    Case 3
                        grupos &= " CentroDeCusto_Id + ' - ' + Descricao as Grupo" & i + 1 & ","
                        grup &= " CentroDeCusto_Id + ' - ' + Descricao,"
                End Select
            Next

            If lstCentroCustoSelecionados.Items.Count = 0 Then
                grupos &= " CentroDeCusto_Id + ' - ' + Descricao as Grupo" & i + 1 & ","
                grup &= " CentroDeCusto_Id + ' - ' + Descricao,"
            End If

            Select Case lstCentroCusto.Items.Count
                Case 0 : grupos &= " '' as grupo4,"
                Case 1 : grupos &= " '' as grupo4,"
                Case 2 : grupos &= " '' as grupo3, '' as grupo4,"
                Case 3 : grupos &= " '' as grupo2, '' as grupo3, '' as grupo4,"
                Case 4 : grupos &= " '' as grupo1, '' as grupo2, '' as grupo3, '' as grupo4,"
            End Select

            'Calcuala os debitos e os creditos 
            sqldebitocredito = "Select SB.Safra, isnull(SB.UnidadeDeNegocio,'00') as UnidadeDeNegocio, isnull(UN.Nome,'NENHUMA') as NomeUnidade, SB.Empresa_Id, SB.EndEmpresa_Id, Empresa.Nome + ' / ' + Empresa.Cidade + '-' + Empresa.Estado as NomeEmpresa, SB.CentroDeCusto_Id, SB.Descricao, SB.Conta_Id, PL.Titulo," & vbCrLf & _
                               "       sum(CASE WHEN SB.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' then SB.TemNF else 0 end) as TemNF," & vbCrLf & _
                               "       sum(CASE WHEN SB.Movimento < '" & txtDataInicial.Text.ToSqlDate() & "' and year(SB.Movimento) = " & CDate(txtDataInicial.Text).Year & " then sb.Debitos - sb.Creditos else 0 end) as SaldoAnterior," & vbCrLf & _
                               "       sum(case when SB.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN SB.Debitos else 0 end) as Debitos," & vbCrLf & _
                               "       Sum(case when Sb.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' THEN SB.Creditos else 0 end) as Creditos " & vbCrLf & _
                               "  Into #DebitosECreditos " & vbCrLf & _
                               "  From" & vbCrLf & _
                               "	   (" & vbCrLf & _
                               "		SELECT Balancete.Movimento_id as Movimento," & vbCrLf & _
                               "               isnull(P.Safra,'NENHUMA') as Safra," & vbCrLf & _
                               "               Balancete.Movimento_Id, " & vbCrLf & _
                               "               Balancete.UnidadeDeNegocio," & vbCrLf & _
                               "			   Balancete.Empresa_Id," & vbCrLf & _
                               "			   Balancete.EndEmpresa_Id," & vbCrLf & _
                               "			   left(CCD.CentroDeCusto_Id," & VNivelDeCusto & ") as CentroDeCusto_Id," & vbCrLf & _
                               "			   CCD.Descricao, " & vbCrLf & _
                               "			   Balancete.Conta_Id, " & vbCrLf & _
                               "               Balancete.Lote_Id," & vbCrLf & _
                               "               Balancete.Sequencia_Id," & vbCrLf & _
                               "			   Balancete.DebitoOficial as Debitos, " & vbCrLf & _
                               "               Balancete.CreditoOficial AS Creditos, " & vbCrLf & _
                               "               case when Balancete.numero_NF > 0 then 1 else 0 end as TemNF" & vbCrLf & _
                               "		  FROM Razao AS Balancete" & vbCrLf & _
                               "		  LEFT JOIN GruposxEmpresas GE" & vbCrLf & _
                               "			ON GE.Cliente_Id    = Balancete.Empresa_Id" & vbCrLf & _
                               "		   AND GE.EndCliente_Id = Balancete.EndEmpresa_Id" & vbCrLf & _
                               "		   AND GE.Empresa_id    = Balancete.UnidadeDeNegocio " & vbCrLf & _
                               "		 INNER JOIN CentrosDeCustos " & vbCrLf & _
                               "			ON CentrosDeCustos.CentroDeCusto_Id = Balancete.Custo " & vbCrLf & _
                               "		  LEFT JOIN CentrosDeCustos CCD" & vbCrLf & _
                               "			ON CCD.CentroDeCusto_Id = left(CentrosDeCustos.CentroDeCusto_Id," & VNivelDeCusto & ")" & vbCrLf & _
                               "          LEFT Join Pedidos P" & vbCrLf & _
                               "            ON Balancete.Empresa_id    = P.Empresa_id" & vbCrLf & _
                               "           AND Balancete.EndEmpresa_Id = P.EndEmpresa_id" & vbCrLf & _
                               "           AND Balancete.Pedido        = P.Pedido_Id" & vbCrLf & _
                               "          Left JOIN NotasFiscaisXItensXRateio NFxIxR" & vbCrLf & _
                               "        	ON Balancete.Empresa_Id      = NFxIxR.Empresa_Id " & vbCrLf & _
                               "           AND Balancete.EndEmpresa_Id   = NFxIxR.EndEmpresa_Id " & vbCrLf & _
                               "           AND Balancete.Cliente_NF      = NFxIxR.Cliente_Id " & vbCrLf & _
                               "           AND Balancete.EndCliente_NF   = NFxIxR.EndCliente_Id " & vbCrLf & _
                               "           AND Balancete.EntradaSaida_NF = NFxIxR.EntradaSaida_Id " & vbCrLf & _
                               "           AND Balancete.Serie_NF        = NFxIxR.Serie_Id " & vbCrLf & _
                               "           AND Balancete.Numero_Nf       = NFxIxR.Nota_Id " & vbCrLf & _
                               "           AND Balancete.Produto_NF      = NFxIxR.Produto_Id " & vbCrLf & _
                               "           AND Balancete.Sequencia_NF    = NFxIxR.Sequencia_Id" & vbCrLf & _
                               "         Where len(Balancete.Custo)    = 5" & vbCrLf & _
                               "           And Balancete.Rateado = 0" & vbCrLf & _
                               "           And Balancete.Lote_Id not in (7500,7600)" & vbCrLf & _
                               "           AND NFxIxR.Empresa_Id is null" & vbCrLf & _
                               "         UNION ALL" & vbCrLf & _
                               "		select NF.Movimento," & vbCrLf & _
                               "               P.Safra, " & vbCrLf & _
                               "               NF.Movimento, " & vbCrLf & _
                               "               NFxIxR.UnidadeDeNegocioRateio_id," & vbCrLf & _
                               "			   NFxIxR.EmpresaRateio_Id," & vbCrLf & _
                               "			   NFxIxR.EndEmpresaRateio_Id," & vbCrLf & _
                               "		       left(NFxIxR.CentroDeCusto_Id," & VNivelDeCusto & ") as CentroDeCusto_Id," & vbCrLf & _
                               "			   CC.Descricao, " & vbCrLf & _
                               "			   OE.DebitaConta as Conta_Id," & vbCrLf & _
                               "               0," & vbCrLf & _
                               "               row_number()over(Order by NF.movimento)," & vbCrLf & _
                               "		       NFxIxR.Valor as Debitos," & vbCrLf & _
                               "		       0 as Creditos," & vbCrLf & _
                               "               1 as TemNF " & vbCrLf & _
                               "		  from notasFiscais NF WITH(INDEX = Idx_Movimento )" & vbCrLf & _
                               "		 Inner Join NotasFiscaisXItens NFxI " & vbCrLf & _
                               "			ON NF.Empresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
                               "		   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
                               "		   AND NF.Cliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
                               "		   AND NF.EndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
                               "		   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
                               "		   AND NF.Serie_Id        = NFxI.Serie_Id " & vbCrLf & _
                               "		   AND NF.Nota_Id         = NFxI.Nota_Id " & vbCrLf & _
                               "		 INNER JOIN NotasFiscaisXEncargos NFxE" & vbCrLf & _
                               "			ON NFxI.Empresa_Id      = NFxE.Empresa_Id " & vbCrLf & _
                               "		   AND NFxI.EndEmpresa_Id   = NFxE.EndEmpresa_Id " & vbCrLf & _
                               "		   AND NFxI.Cliente_Id      = NFxE.Cliente_Id " & vbCrLf & _
                               "		   AND NFxI.EndCliente_Id   = NFxE.EndCliente_Id " & vbCrLf & _
                               "		   AND NFxI.EntradaSaida_Id = NFxE.EntradaSaida_Id " & vbCrLf & _
                               "		   AND NFxI.Serie_Id        = NFxE.Serie_Id " & vbCrLf & _
                               "		   AND NFxI.Nota_Id         = NFxE.Nota_Id " & vbCrLf & _
                               "		   AND NFxI.Produto_Id      = NFxE.Produto_Id " & vbCrLf & _
                               "		   AND NFxI.Sequencia_Id    = NFxE.Sequencia_id" & vbCrLf & _
                               "		   AND NFxE.Encargo_Id      = 'PRODUTO'" & vbCrLf & _
                               "		 INNER JOIN NotasFiscaisXItensXRateio NFxIxR" & vbCrLf & _
                               "			ON NFxI.Empresa_Id      = NFxIxR.Empresa_Id " & vbCrLf & _
                               "		   AND NFxI.EndEmpresa_Id   = NFxIxR.EndEmpresa_Id " & vbCrLf & _
                               "		   AND NFxI.Cliente_Id      = NFxIxR.Cliente_Id " & vbCrLf & _
                               "		   AND NFxI.EndCliente_Id   = NFxIxR.EndCliente_Id " & vbCrLf & _
                               "		   AND NFxI.EntradaSaida_Id = NFxIxR.EntradaSaida_Id " & vbCrLf & _
                               "		   AND NFxI.Serie_Id        = NFxIxR.Serie_Id " & vbCrLf & _
                               "		   AND NFxI.Nota_Id         = NFxIxR.Nota_Id " & vbCrLf & _
                               "		   AND NFxI.Produto_Id      = NFxIxR.Produto_Id " & vbCrLf & _
                               "		   AND NFxI.Sequencia_Id    = NFxIxR.Sequencia_Id" & vbCrLf & _
                               "		  LEFT JOIN CentrosDeCustos CC" & vbCrLf & _
                               "			ON CC.CentroDeCusto_Id = left(NFxIxR.CentroDeCusto_Id," & VNivelDeCusto & ")" & vbCrLf & _
                               "		 INNER JOIN OperacaoXEstadoXEncargo OE" & vbCrLf & _
                               "			ON OE.Codigo_Id  = NFxI.OperacaoxEstado " & vbCrLf & _
                               "		   AND OE.Encargo_Id = 'PRODUTO'" & vbCrLf & _
                               "		 Inner Join Pedidos P" & vbCrLf & _
                               "			ON P.Empresa_Id    = NF.Empresa_Id" & vbCrLf & _
                               "           And P.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf & _
                               "           And P.Pedido_Id     = NF.Pedido" & vbCrLf & _
                               "		 )  as SB" & vbCrLf & _
                               " INNER JOIN PlanoDeContas PL" & vbCrLf & _
                               "    ON SB.Conta_Id = PL.Conta_Id " & vbCrLf & _
                               "  LEFT Join Clientes UN" & vbCrLf & _
                               "    on UN.Cliente_id  = SB.UnidadeDeNegocio" & vbCrLf & _
                               "   and UN.Endereco_id = 0" & vbCrLf & _
                               " INNER Join Clientes Empresa" & vbCrLf & _
                               "    on Empresa.Cliente_id  = SB.Empresa_id" & vbCrLf & _
                               "   and Empresa.Endereco_id = SB.EndEmpresa_Id" & vbCrLf & _
                               " WHERE SB.Movimento_Id <= '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                               "   and sb.CentroDeCusto_Id is not null " & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sqldebitocredito &= "   AND isnull(SB.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqldebitocredito &= "           And left(SB.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sqldebitocredito &= "           And SB.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sqldebitocredito &= "           And SB.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                                        "           AND SB.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            sqldebitocredito &= "AND (Left(SB.Conta_Id, 1) In" & VGrupos & ") " & vbCrLf

            If lstCentroDeCusto.GetSelectedValues().Count > 0 Then
                sqldebitocredito &= " AND SB.CentroDeCusto_Id in ('" & String.Join("','", getListCentroDeCusto()) & "')" & vbCrLf
            End If

            sqldebitocredito &= " GROUP BY SB.Safra, SB.UnidadeDeNegocio, UN.Nome, SB.Empresa_Id, SB.EndEmpresa_Id,  Empresa.Nome + ' / ' + Empresa.Cidade + '-' + Empresa.Estado, SB.CentroDeCusto_Id, SB.Descricao, SB.Conta_Id, PL.Titulo" & vbCrLf

            sqldebitocredito &= " Select SB.*" & vbCrLf & _
                                "   Into #tt" & vbCrLf & _
                                "   From (" & vbCrLf & _
                                "         select S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,1) as conta_Id, PC.Titulo, sum(S.SaldoAnterior) as SaldoAnterior, sum(S.Debitos) as Debitos, sum(S.Creditos) as Creditos, 0 TemNF" & vbCrLf & _
                                "           from #DebitosECreditos S" & vbCrLf & _
                                "           LEFT JOIN planodecontas PC" & vbCrLf & _
                                "             ON PC.Conta_Id = left(S.Conta_Id,1)" & vbCrLf & _
                                "          group by S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,1) , PC.Titulo" & vbCrLf & _
                                "          union " & vbCrLf & _
                                "         select S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,3) as conta_Id, PC.Titulo, sum(S.SaldoAnterior) as SaldoAnterior, sum(S.Debitos) as Debitos, sum(S.Creditos) as Creditos, 0 TemNF" & vbCrLf & _
                                "           from #DebitosECreditos S" & vbCrLf & _
                                "           LEFT JOIN planodecontas PC" & vbCrLf & _
                                "             ON PC.Conta_Id = left(S.Conta_Id,3)" & vbCrLf & _
                                "          group by S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,3) , PC.Titulo" & vbCrLf & _
                                "          union" & vbCrLf & _
                                "         select S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,5) as conta_Id, PC.Titulo,  sum(S.SaldoAnterior) as SaldoAnterior, sum(S.Debitos) as Debitos, sum(S.Creditos) as Creditos, 0 TemNF" & vbCrLf & _
                                "           from #DebitosECreditos S" & vbCrLf & _
                                "           LEFT JOIN planodecontas PC" & vbCrLf & _
                                "             ON PC.Conta_Id = left(S.Conta_Id,5)" & vbCrLf & _
                                "          group by S.Safra, S.UnidadeDeNegocio, S.NomeUnidade, S.Empresa_Id, S.EndEmpresa_Id, S.NomeEmpresa, S.CentroDeCusto_Id, S.Descricao, left(S.Conta_Id,5) , PC.Titulo" & vbCrLf & _
                                "          Union" & vbCrLf & _
                                "         select Safra, UnidadeDeNegocio, NomeUnidade, Empresa_Id, EndEmpresa_Id, NomeEmpresa, CentroDeCusto_Id, Descricao, Conta_id, Titulo, SaldoAnterior, Debitos, Creditos, TemNF" & vbCrLf & _
                                "           from #DebitosECreditos" & vbCrLf & _
                                " ) as SB " & vbCrLf

            For i As Integer = 0 To lstCentroCusto.Items.Count - 1
                Select Case lstCentroCusto.Items(i).Value
                    Case 0 : sqldebitocredito &= " Update #tt set Safra = '' "
                    Case 1 : sqldebitocredito &= " Update #tt set UnidadeDeNegocio = '', NomeUnidade = '' "
                    Case 2 : sqldebitocredito &= " Update #tt set Empresa_Id = '', EndEmpresa_id = 0, NomeEmpresa = '' "
                End Select
            Next

            '------------------------- Principal -------------------------------------------------------------------
            sqldebitocredito &= " Select " & grupos & " Safra, UnidadeDeNegocio, NomeUnidade, Empresa_Id, EndEmpresa_Id, NomeEmpresa, CentroDeCusto_Id, Descricao, Conta_id, Titulo,  sum(TemNF) as TemNF,  sum(SaldoAnterior) as SaldoAnterior, sum(Debitos) as Debitos, sum(Creditos) as Creditos" & vbCrLf & _
                                "   from #tt" & vbCrLf & _
                                "  Group by " & grup & " Safra, UnidadeDeNegocio, NomeUnidade, Empresa_Id, EndEmpresa_Id, NomeEmpresa, CentroDeCusto_Id, Descricao, Conta_id, Titulo" & vbCrLf & _
                                "  having  sum(SaldoAnterior) <> 0 or sum(Debitos) <> 0 or sum(Creditos) <> 0 "

            '------------------------- Totais -------------------------------------------------------------------
            sqldebitocredito &= " Select Conta_id, Titulo,  sum(SaldoAnterior) as SaldoAnterior, sum(Debitos) as Debitos, sum(Creditos) as Creditos" & vbCrLf & _
                                "   from #tt" & vbCrLf & _
                                "  Group by Conta_id, Titulo order by Conta_id"

            Ds_BalanceCentroDeCusto = Banco.ConsultaDataSet(sqldebitocredito, "Consulta")
            Ds_BalanceCentroDeCusto.Tables("Consulta").TableName = "PlanoDeContas"
            Ds_BalanceCentroDeCusto.Tables("Consulta1").TableName = "Totais"

            Dim sqlcomposicao As String = ""
            sqlcomposicao = "Select NF.Empresa_Id as EmpresaDaNota," & vbCrLf & _
                            "       NF.EndEmpresa_Id as EndEmpresaDaNota," & vbCrLf & _
                            "       NF.Pedido," & vbCrLf & _
                            "       NF.Observacoes," & vbCrLf & _
                            "       NF.Movimento," & vbCrLf & _
                            "       NfxR.UnidadeDeNegocioRateio_Id," & vbCrLf & _
                            "       isnull(NfxR.EmpresaRateio_Id,NF.Empresa_Id) as EmpresaRateio_Id," & vbCrLf & _
                            "       case when not NfxR.EmpresaRateio_Id is null then NfxR.EndEmpresaRateio_Id else NF.EndEmpresa_Id end as EndEmpresaRateio_Id," & vbCrLf & _
                            "       NF.Cliente_id," & vbCrLf & _
                            "       NF.EndCliente_Id," & vbCrLf & _
                            "       NF.Nota_id," & vbCrLf & _
                            "       NF.Serie_Id," & vbCrLf & _
                            "       NF.EntradaSaida_Id," & vbCrLf & _
                            "       Nfi.Produto_Id," & vbCrLf & _
                            "       nfe.Encargo_Id as Encargo_Id, " & vbCrLf & _
                            "       isnull(NfxR.CentroDeCusto_Id,nfe.centrodecusto) as CentroDeCusto_Id," & vbCrLf & _
                            "       OxE.DebitaConta," & vbCrLf & _
                            "       OxE.CreditaConta," & vbCrLf & _
                            "       OxE.Sinal," & vbCrLf & _
                            "       isnull(NfxR.Valor,nfi.valor) as Valor," & vbCrLf & _
                            "       OxE.Aliquota as Percentual" & vbCrLf & _
                            "   Into #T1" & vbCrLf & _
                            "  from NotasFiscais NF WITH(INDEX = Idx_Movimento)" & vbCrLf & _
                            "  INNER Join NotasfiscaisXItens Nfi" & vbCrLf & _
                            "     ON NF.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
                            "    AND NF.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
                            "    AND NF.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
                            "    AND NF.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
                            "    AND NF.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
                            "    AND NF.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
                            "    AND NF.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
                            "  INNER Join Notasfiscaisxencargos Nfe" & vbCrLf & _
                            "     ON Nfi.Empresa_Id      = Nfe.Empresa_Id" & vbCrLf & _
                            "    AND Nfi.EndEmpresa_Id   = Nfe.EndEmpresa_Id" & vbCrLf & _
                            "    AND Nfi.Cliente_Id      = Nfe.Cliente_Id" & vbCrLf & _
                            "    AND Nfi.EndCliente_Id   = Nfe.EndCliente_Id" & vbCrLf & _
                            "    AND Nfi.EntradaSaida_Id = Nfe.EntradaSaida_Id" & vbCrLf & _
                            "    AND Nfi.Serie_Id        = Nfe.Serie_Id" & vbCrLf & _
                            "    AND Nfi.Nota_Id         = Nfe.Nota_Id" & vbCrLf & _
                            "    AND Nfi.Produto_Id      = Nfe.Produto_Id" & vbCrLf & _
                            "    AND Nfi.sequencia_Id    = Nfe.sequencia_Id" & vbCrLf & _
                            "  Inner join (Select Empresa_Id," & vbCrLf & _
                            "                     EndEmpresa_Id," & vbCrLf & _
                            "                     Cliente_Id," & vbCrLf & _
                            "                     EndCliente_Id," & vbCrLf & _
                            "                     EntradaSaida_Id," & vbCrLf & _
                            "                     Serie_Id," & vbCrLf & _
                            "                     Nota_Id," & vbCrLf & _
                            "                     produto_id," & vbCrLf & _
                            "                     sequencia_Id," & vbCrLf & _
                            "                     CentroDeCusto" & vbCrLf & _
                            "                From Notasfiscaisxencargos" & vbCrLf & _
                            "                Where Encargo_Id = 'PRODUTO'" & vbCrLf & _
                            "               ) SB_PRODUTO" & vbCrLf & _
                            "     ON Nfi.Empresa_Id      = SB_PRODUTO.Empresa_Id" & vbCrLf & _
                            "    AND Nfi.EndEmpresa_Id   = SB_PRODUTO.EndEmpresa_Id" & vbCrLf & _
                            "    AND Nfi.Cliente_Id      = SB_PRODUTO.Cliente_Id" & vbCrLf & _
                            "    AND Nfi.EndCliente_Id   = SB_PRODUTO.EndCliente_Id" & vbCrLf & _
                            "    AND Nfi.EntradaSaida_Id = SB_PRODUTO.EntradaSaida_Id" & vbCrLf & _
                            "    AND Nfi.Serie_Id        = SB_PRODUTO.Serie_Id" & vbCrLf & _
                            "    AND Nfi.Nota_Id         = SB_PRODUTO.Nota_Id" & vbCrLf & _
                            "    AND Nfi.Produto_Id      = SB_PRODUTO.Produto_Id" & vbCrLf & _
                            "    AND Nfi.sequencia_Id    = SB_PRODUTO.sequencia_Id" & vbCrLf & _
                            "   left JOIN NotasFiscaisXItensXRateio NfxR" & vbCrLf & _
                            "     ON NfxR.Empresa_Id      = Nfi.Empresa_Id" & vbCrLf & _
                            "    AND NfxR.EndEmpresa_Id   = Nfi.EndEmpresa_Id" & vbCrLf & _
                            "    AND NfxR.Cliente_Id      = Nfi.Cliente_Id" & vbCrLf & _
                            "    AND NfxR.EndCliente_Id   = Nfi.EndCliente_Id" & vbCrLf & _
                            "    AND NfxR.EntradaSaida_Id = Nfi.EntradaSaida_Id" & vbCrLf & _
                            "    AND NfxR.Serie_Id        = Nfi.Serie_Id" & vbCrLf & _
                            "    AND NfxR.Nota_Id         = Nfi.Nota_Id" & vbCrLf & _
                            "  inner Join OperacaoXEstadoXEncargo OxE" & vbCrLf & _
                            "     on OxE.Codigo_Id     = Nfi.OperacaoXEstado" & vbCrLf & _
                            "    and OxE.Encargo_Id = Nfe.Encargo_Id" & vbCrLf & _
                            "  inner join Produtos P" & vbCrLf & _
                            "     on P.Produto_id = Nfi.Produto_Id" & vbCrLf & _
                            "  inner join Clientes C" & vbCrLf & _
                            "     on C.Cliente_Id  = NF.Cliente_ID" & vbCrLf & _
                            "    and C.Endereco_Id = NF.EndCliente_Id" & vbCrLf & _
                            "  inner join Pedidos Pe" & vbCrLf & _
                            "     on Nf.Empresa_id    = pe.Empresa_Id" & vbCrLf & _
                            "    and Nf.EndEmpresa_id = pe.EndEmpresa_id" & vbCrLf & _
                            "    and NF.Pedido        = Pe.Pedido_Id" & vbCrLf

            If Not chkAnaliticoPorNota.Checked Then
                sqlcomposicao &= " Where 1 = 2 "
            Else
                sqlcomposicao &= "  where NF.Movimento between '" & txtDataInicial.Text.ToSqlDate() & "' and '" & txtDataFinal.Text.ToSqlDate() & "'" & vbCrLf & _
                                 "    and (SB_PRODUTO.centrodecusto > 0 or  nfxr.empresa_id is not null)" & vbCrLf

                If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                    sqlcomposicao &= "   AND isnull(NfxR.UnidadeDeNegocioRateio_Id,Pe.UnidadedeNegocio) = '" & ddlUnidade.SelectedValue & "'" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    If chkConsolidarEmpresa.Checked Then
                        sqlcomposicao &= "           And left(isnull(NfxR.EmpresaRateio_Id,NF.Empresa_Id),8)  = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                    ElseIf chkConsolidarSede.Checked Then
                        sqlcomposicao &= "           And isnull(NfxR.EmpresaRateio_Id,NF.Empresa_Id) ='" & Emp(0) & "'" & vbCrLf
                    Else
                        sqlcomposicao &= "           And isnull(NfxR.EmpresaRateio_Id,NF.Empresa_Id) ='" & Emp(0) & "'" & vbCrLf & _
                                         "           AND case when not NfxR.EmpresaRateio_Id is null then NfxR.EndEmpresaRateio_Id else NF.EndEmpresa_Id end = " & Emp(1) & vbCrLf
                    End If
                End If

                sqlcomposicao &= "AND (Left(isnull(OxE.DebitaConta,''), 1) In" & VGrupos & " or Left(isnull(OxE.CreditaConta,''), 1) In" & VGrupos & ")" & vbCrLf

                If lstCentroDeCusto.GetSelectedValues().Count > 0 Then
                    sqlcomposicao &= " And left(isnull(NfxR.CentroDeCusto_Id,nfe.centrodecusto)," & VNivelDeCusto & ") in ('" & String.Join("','", getListCentroDeCusto()) & "')" & vbCrLf
                    sqldebitocredito &= " AND SB.CentroDeCusto_Id in ('" & String.Join("','", getListCentroDeCusto()) & "')" & vbCrLf
                End If
            End If

            '*****************************************************************************************************
            sqlcomposicao &= " select Pe.Safra," & vbCrLf & _
                             "        Pe.UnidadedeNegocio," & vbCrLf & _
                             "        #T1.EmpresaDaNota," & vbCrLf & _
                             "        #T1.EndEmpresaDaNota," & vbCrLf & _
                             "        #T1.Observacoes," & vbCrLf & _
                             "        #T1.Movimento," & vbCrLf & _
                             "        isnull(#T1.UnidadeDeNegocioRateio_Id,Pe.UnidadedeNegocio) as UnidadeDeNegocioRateio_Id," & vbCrLf & _
                             "        #T1.EmpresaRateio_Id," & vbCrLf & _
                             "        #T1.EndEmpresaRateio_Id," & vbCrLf & _
                             "        #T1.Cliente_id," & vbCrLf & _
                             "        #T1.EndCliente_Id," & vbCrLf & _
                             "        C.Nome," & vbCrLf & _
                             "        #T1.Nota_id," & vbCrLf & _
                             "        #T1.Serie_Id," & vbCrLf & _
                             "        #T1.EntradaSaida_Id," & vbCrLf & _
                             "        #T1.Produto_Id," & vbCrLf & _
                             "        P.Nome as DescProduto," & vbCrLf & _
                             "        #T1.Encargo_Id, " & vbCrLf & _
                             "        #T1.CentroDeCusto_Id," & vbCrLf & _
                             "        #T1.DebitaConta," & vbCrLf & _
                             "        #T1.CreditaConta," & vbCrLf & _
                             "        #T1.Sinal," & vbCrLf & _
                             "        #T1.Valor," & vbCrLf & _
                             "        #T1.Percentual  " & vbCrLf & _
                             "   Into #Temp" & vbCrLf & _
                             "   from #T1" & vbCrLf & _
                             "  inner join Clientes C" & vbCrLf & _
                             "     on C.Cliente_Id  = #T1.Cliente_ID" & vbCrLf & _
                             "    and C.Endereco_Id = #T1.EndCliente_Id" & vbCrLf & _
                             "  inner join Pedidos Pe" & vbCrLf & _
                             "     on PE.Empresa_Id    = #T1.EmpresaDaNota" & vbCrLf & _
                             "    and pe.EndEmpresa_id = #T1.EndEmpresaDaNota" & vbCrLf & _
                             "    and Pe.Pedido_Id     = #T1.Pedido" & vbCrLf & _
                             "  Inner Join Produtos p" & vbCrLf & _
                             "     on P.Produto_Id = #T1.Produto_Id" & vbCrLf & _
                             "  Where #T1.Encargo_Id in('PRODUTO','PIS','COFINS')" & vbCrLf

            '*****************************************************************************************************
            sqlcomposicao &= " Update #temp set" & vbCrLf & _
                             "  valor = round(valor * percentual / 100,2)" & vbCrLf & _
                             " where encargo_Id in ('PIS','COFINS')" & vbCrLf

            sqlcomposicao &= " Update #temp set" & vbCrLf & _
                             "   valor = sb.Valor" & vbCrLf & _
                             "  from #Temp" & vbCrLf & _
                             " inner join (Select UnidadeDeNegocioRateio_Id," & vbCrLf & _
                             "                    EmpresaRateio_Id," & vbCrLf & _
                             "                    Cliente_Id," & vbCrLf & _
                             "                    EndCliente_Id," & vbCrLf & _
                             "                    Nota_Id," & vbCrLf & _
                             "                    Serie_Id," & vbCrLf & _
                             "                    EntradaSaida_Id," & vbCrLf & _
                             "                    Produto_Id," & vbCrLf & _
                             "                    CentroDeCusto_Id," & vbCrLf & _
                             "                    Sum(case" & vbCrLf & _
                             "                           when sinal = '=' then 0" & vbCrLf & _
                             "                           when sinal = '-' then Valor * -1" & vbCrLf & _
                             "                           when sinal = '+' then valor" & vbCrLf & _
                             "                        end) as Valor" & vbCrLf & _
                             "               from #Temp" & vbCrLf & _
                             "              group by UnidadeDeNegocioRateio_Id," & vbCrLf & _
                             "					   EmpresaRateio_Id," & vbCrLf & _
                             "					   Cliente_Id," & vbCrLf & _
                             "					   EndCliente_Id," & vbCrLf & _
                             "					   Nota_Id," & vbCrLf & _
                             "					   Serie_Id," & vbCrLf & _
                             "					   EntradaSaida_Id," & vbCrLf & _
                             "					   Produto_Id," & vbCrLf & _
                             "                     CentroDeCusto_Id" & vbCrLf & _
                             "             ) sb" & vbCrLf & _
                             "    on #Temp.UnidadeDeNegocioRateio_Id = sb.UnidadeDeNegocioRateio_Id" & vbCrLf & _
                             "   and #Temp.EmpresaRateio_Id          = sb.EmpresaRateio_Id" & vbCrLf & _
                             "   and #Temp.Cliente_Id                = sb.Cliente_Id" & vbCrLf & _
                             "   and #Temp.EndCliente_Id             = sb.EndCliente_Id" & vbCrLf & _
                             "   and #Temp.Nota_Id                   = sb.Nota_Id" & vbCrLf & _
                             "   and #Temp.Serie_Id                  = sb.Serie_Id" & vbCrLf & _
                             "   and #Temp.EntradaSaida_Id           = sb.EntradaSaida_Id" & vbCrLf & _
                             "   and #Temp.Produto_Id                = sb.Produto_Id" & vbCrLf & _
                             "   and #Temp.CentroDeCusto_Id          = sb.CentroDeCusto_Id" & vbCrLf & _
                             " Where Encargo_Id = 'PRODUTO'" & vbCrLf




            Dim pSafra As Boolean = False
            Dim pUnidade As Boolean = False
            Dim pEmpresa As Boolean = False

            For i As Integer = 0 To lstCentroCusto.Items.Count - 1
                Select Case lstCentroCusto.Items(i).Value
                    Case 0 : pSafra = True 'sqlcomposicao &= " Update #tt set Safra = '' "
                    Case 1 : pUnidade = True 'sqlcomposicao &= " Update #tt set UnidadeDeNegocio = ''"
                    Case 2 : pEmpresa = True 'sqlcomposicao &= " Update #tt set Empresa_Id = '', EndEmpresa_id = 0 "
                End Select
            Next

            sqlcomposicao &= "Select " & vbCrLf & IIf(pSafra, "'' as Safra", "Safra") & "," & IIf(pUnidade, "'' as UnidadeDeNegocio", "UnidadeDeNegocioRateio_Id as UnidadeDeNegocio") & "," & IIf(pEmpresa, "'' as Empresa_Id, 0 as EndEmpresa_id", "EmpresaRateio_id as Empresa_Id, EndEmpresaRateio_Id as EndEmpresa_Id") & "," & vbCrLf & _
                             "       Cliente_Id, EndCliente_Id, Nome as DescCliente, Observacoes as Observacao, Movimento as data, Nota_Id, Serie_Id, EntradaSaida_Id, Produto_Id, DescProduto, Encargo_id," & vbCrLf & _
                             "       left(CentroDeCusto_Id," & VNivelDeCusto & ") as CentroDeCusto_Id, Conta as Conta_Id, Credito,  Debito" & vbCrLf & _
                             "  from (" & vbCrLf & _
                             "		  SELECT Safra, UnidadedeNegocio, EmpresaDaNota, EndEmpresaDaNota, Observacoes,Movimento, UnidadeDeNegocioRateio_Id, EmpresaRateio_id, EndEmpresaRateio_Id, Cliente_Id, EndCliente_Id, Nome, Nota_Id, Serie_Id, EntradaSaida_Id, Produto_Id, DescProduto, Encargo_id, CentroDeCusto_Id, DebitaConta as Conta, 0.00 as credito,  Valor as Debito" & vbCrLf & _
                             "		    FROM #temp" & vbCrLf & _
                             "		   where len(DebitaConta) > 0" & vbCrLf & _
                             "		   union all" & vbCrLf & _
                             "		  SELECT Safra, UnidadedeNegocio, EmpresaDaNota, EndEmpresaDaNota, Observacoes,Movimento, UnidadeDeNegocioRateio_Id, EmpresaRateio_id, EndEmpresaRateio_Id, Cliente_Id, EndCliente_Id, Nome, Nota_Id, Serie_Id, EntradaSaida_Id, Produto_Id, DescProduto, Encargo_id, CentroDeCusto_Id, CreditaConta as Conta, Valor as credito, 0.00 as Debito" & vbCrLf & _
                             "		    FROM #temp" & vbCrLf & _
                             "		   where len(CreditaConta) > 0" & vbCrLf & _
                             "       ) sb" & vbCrLf
            Dim dsComposicao As DataSet
            dsComposicao = Banco.ConsultaDataSet(sqlcomposicao, "Composicao")
            Ds_BalanceCentroDeCusto.Merge(dsComposicao)

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", "Relatório De Anexo Por Centro De Custos.")
            parameters.Add("Parametros", ParametrosDoRelatorio & "Lote <> 7500 e 7600")
            parameters.Add("AnaliticoPorNota", chkAnaliticoPorNota.Checked)
            parameters.Add("ordem", IIf(lstCentroCustoSelecionados.Items.Count = 0, 1, lstCentroCustoSelecionados.Items.Count))
            parameters.Add("IniciarPagina", txtIniciarNaFolha.Text)


            Dim Nomereport As String = "Cr_BalanceteCentrodeCusto"
            If pTipo = eTipoSaidaRelatorio.EXCEL_DADOS Then
                Nomereport = "Cr_BalanceteCentroDeCustoDados"
            End If

            'Funcoes.BindReport(Me.Page, Ds_BalanceCentroDeCusto, "Cr_BalanceteCentrodeCusto", IIf(pTipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(pTipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
            Funcoes.BindReport(Me.Page, Ds_BalanceCentroDeCusto, Nomereport, IIf(pTipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(pTipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub GeraAnexoPorCliente(ByVal ParametrosRelatorio As String, ByVal ParametrosEmpresa As ArrayList, ptipo As eTipoSaidaRelatorio)
        Try
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim ds_AnexoCliente As New DataSet
            Dim sql As String

            sql = "SELECT R.Conta_Id, PC.Titulo," & vbCrLf & _
                  "       SUM(R.SaldoAnterior) AS SaldoAnterior," & vbCrLf & _
                  "       SUM(R.Debitos) AS Debitos," & vbCrLf & _
                  "       SUM(R.Creditos) AS Creditos," & vbCrLf & _
                  "       R.Cliente_Id," & vbCrLf & _
                  "       C.Nome AS NomeCliente," & vbCrLf & _
                  "       R.EndCliente_Id AS Endereco_Id " & vbCrLf & _
                  "  FROM (SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id," & vbCrLf & _
                  "               Razao.Conta_Id," & vbCrLf & _
                  "               Razao.Cliente_Id, Razao.EndCliente_Id," & vbCrLf & _
                  "               ISNULL(SUM(Razao.DebitoOficial - Razao.CreditoOficial), 0) AS SaldoAnterior," & vbCrLf & _
                  "               0 AS Debitos," & vbCrLf & _
                  "               0 AS Creditos " & vbCrLf & _
                  "          FROM Razao " & vbCrLf & _
                  "         WHERE Razao.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sql &= "           AND isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sql &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                           "           and Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            If VSelecao <> "" Then sql &= "           And Left(Razao.Conta_Id, 7) In " & VSelecao
            sql &= "           AND (Left(Razao.Conta_Id, 1) In" & VGrupos & ") "
            If VIsolarLotes <> "" Then sql &= "           And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            sql &= "         GROUP BY Razao.Empresa_Id, Razao.EndEmpresa_Id, Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id " & vbCrLf & _
                   "        HAVING(ISNULL(SUM(Razao.DebitoOficial - Razao.CreditoOficial), 0) <> 0) " & vbCrLf & _
                   "         UNION " & vbCrLf & _
                   "        SELECT Razao.Empresa_Id, Razao.EndEmpresa_Id, Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id, 0 AS SaldoAnterior,SUM(Razao.DebitoOficial) AS Debitos, SUM(Razao.CreditoOficial) AS Creditos " & vbCrLf & _
                   "          FROM Razao " & vbCrLf & _
                   "         WHERE Razao.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sql &= "           AND isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'" & vbCrLf
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sql &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sql &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                           "           And Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            If VSelecao <> "" Then sql &= "           And Left(Razao.Conta_Id, 7) In " & VSelecao
            If VIsolarLotes <> "" Then sql &= "           And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            sql &= "           AND (Left(Razao.Conta_Id, 1) In" & VGrupos & ") " & vbCrLf & _
                   "         GROUP BY Razao.Empresa_Id, Razao.EndEmpresa_Id, Razao.Conta_Id, Razao.Cliente_Id, Razao.EndCliente_Id" & vbCrLf & _
                   "        ) R " & vbCrLf & _
                   "   INNER JOIN PlanoDeContas PC " & vbCrLf & _
                   "      ON R.Conta_Id = PC.Conta_Id " & vbCrLf & _
                   "   INNER JOIN Clientes C " & vbCrLf & _
                   "      ON R.Cliente_Id    = C.Cliente_Id " & vbCrLf & _
                   "     AND R.EndCliente_Id = C.Endereco_Id " & vbCrLf & _
                   "   GROUP BY R.Conta_Id, R.Cliente_Id, R.EndCliente_Id, PC.Titulo, C.Nome " & vbCrLf

            If rdComMovimentoNoPeriodo.Checked Then
                sql &= " having SUM(R.Debitos) > 0 or SUM(R.Creditos) > 0 or SUM(R.SaldoAnterior)  + SUM(R.Debitos) - SUM(R.Creditos) <> 0" & vbCrLf
            ElseIf rdSomenteComSaldoAtual.Checked Then
                sql &= " having  SUM(R.SaldoAnterior)  + SUM(R.Debitos) - SUM(R.Creditos) <> 0" & vbCrLf
            End If
            ds_AnexoCliente = Banco.ConsultaDataSet(sql, "Razao")

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("titulo", "Relatório De Anexo Por Clientes.")
            parameters.Add("Parametros", ParametrosRelatorio)
            parameters.Add("IniciarPagina", txtIniciarNaFolha.Text)

            Funcoes.BindReport(Me.Page, ds_AnexoCliente, "Cr_AnexoPorCliente", IIf(ptipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(ptipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub GeraAnexoPorContaProduto(ByVal ParametrosRelatorio As String, ByVal ParametrosEmpresa As ArrayList, pTipo As eTipoSaidaRelatorio)
        Try
            Dim sqlPorContaProduto As String
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim sqlplano As String
            Dim DS_Anterior1 As New DataSet
            Dim Ds_DebitosCreditos1 As DataSet
            Dim tbBalanceteCentro As DataTable = New DataTable("Razao")
            Dim Ds_AnexoProduto As New DataSet
            Dim Ds_PlanoDeContas As New DataSet
            Dim drConta As DataRow
            Dim drAnexo As DataRow
            Dim cont As Integer = 1
            Dim RetornoProdutos As ArrayList

            sqlPorContaProduto = " select nome, produto_Id, conta_id,titulo,sum(saldoanterior) as saldoanterior,sum(debitos) as debitos,sum(creditos) as creditos, left(conta_id,7) as Grupo, DescGrupo,right(conta_id,2) as conta_Id2 " & vbCrLf & _
                                 "   from( " & vbCrLf

            If VNivelGP = 0 Then
                sqlPorContaProduto &= "        SELECT Produtos.Nome, Produtos.Produto_Id,"
            Else
                sqlPorContaProduto &= "        SELECT GE.Descricao as Nome, ge.Grupo_Id as Produto_Id,"
            End If

            sqlPorContaProduto &= "               PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, " & vbCrLf & _
                                  "               SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS SaldoAnterior, 0.00 as Debitos, 0.00 as Creditos,PXG.Titulo as DescGrupo " & vbCrLf & _
                                  "          FROM Razao " & vbCrLf & _
                                  "          LEFT JOIN Produtos " & vbCrLf & _
                                  "            ON Razao.Produto = Produtos.Produto_Id " & vbCrLf & _
                                  "         inner join GruposDeEstoques GE " & vbCrLf

            Select Case VNivelGP
                Case 0 : sqlPorContaProduto &= "            ON Produtos.Grupo         = ge.Grupo_Id"
                Case 1 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,1) = ge.Grupo_Id"
                Case 2 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,2) = ge.Grupo_Id"
                Case 3 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,3) = ge.Grupo_Id"
                Case 5 : sqlPorContaProduto &= "            ON Produtos.Grupo         = ge.Grupo_Id"
            End Select


            sqlPorContaProduto &= "         INNER JOIN PlanoDeContas " & vbCrLf & _
                                  "            ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                                  "         INNER JOIN PlanoDeContas PXG" & vbCrLf & _
                                  "            ON left(Razao.Conta_Id,7 )= PXG.Conta_Id " & vbCrLf & _
                                  "         WHERE Razao.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                sqlPorContaProduto &= " AND " & RetornoProdutos(0)
                ParametrosRelatorio &= RetornoProdutos(1)
            End If

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sqlPorContaProduto &= "           AND isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'"
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqlPorContaProduto &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sqlPorContaProduto &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sqlPorContaProduto &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                                          "           And Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            sqlPorContaProduto &= "AND    (Left(PlanoDeContas.Conta_Id, 1) In" & VGrupos & ") "
            If VSelecao <> "" Then sqlPorContaProduto &= " And Left(Razao.Conta_Id, 7) In " & VSelecao
            If VIsolarLotes <> "" Then sqlPorContaProduto &= " And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            If VNivelGP = 0 Then
                sqlPorContaProduto &= " GROUP BY Produtos.Nome, Produtos.Produto_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PXG.Titulo " & vbCrLf
            Else
                sqlPorContaProduto &= " GROUP BY GE.Descricao, ge.Grupo_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PXG.Titulo " & vbCrLf
            End If

            cont = 1
            sqlPorContaProduto &= " Union All " & vbCrLf

            If VNivelGP = 0 Then
                sqlPorContaProduto &= "        SELECT Produtos.Nome, Produtos.Produto_Id,"
            Else
                sqlPorContaProduto &= "        SELECT GE.Descricao as Nome, ge.Grupo_Id as Produto_Id,"
            End If

            sqlPorContaProduto &= "       PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, 0.00 as SaldoAnterior, " & vbCrLf & _
                                  "       SUM(Razao.DebitoOficial) AS Debitos, SUM(Razao.CreditoOficial) AS Creditos,PXG.Titulo as DescGrupo " & vbCrLf & _
                                  "  FROM Razao " & vbCrLf & _
                                  "  LEFT JOIN Produtos " & vbCrLf & _
                                  "    ON Razao.Produto = Produtos.Produto_Id " & vbCrLf & _
                                  " inner join GruposDeEstoques GE " & vbCrLf

            Select Case VNivelGP
                Case 0 : sqlPorContaProduto &= "            ON Produtos.Grupo         = ge.Grupo_Id"
                Case 1 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,1) = ge.Grupo_Id"
                Case 2 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,2) = ge.Grupo_Id"
                Case 3 : sqlPorContaProduto &= "            ON left(Produtos.Grupo,3) = ge.Grupo_Id"
                Case 5 : sqlPorContaProduto &= "            ON Produtos.Grupo         = ge.Grupo_Id"
            End Select

            sqlPorContaProduto &= " INNER JOIN PlanoDeContas " & vbCrLf & _
                                  "    ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                                  " INNER JOIN PlanoDeContas PXG" & vbCrLf & _
                                  "    ON left(Razao.Conta_Id,7 )= PXG.Conta_Id " & vbCrLf & _
                                  " WHERE LEN(Razao.Conta_Id) = 9" & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                sqlPorContaProduto &= " AND " & RetornoProdutos(0)
            End If

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sqlPorContaProduto &= "   AND isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'"
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqlPorContaProduto &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sqlPorContaProduto &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sqlPorContaProduto &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                                          "           And Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            sqlPorContaProduto &= " AND Razao.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

            sqlPorContaProduto &= " AND    (Left(PlanoDeContas.Conta_Id, 1) In" & VGrupos & ") "
            If VSelecao <> "" Then sqlPorContaProduto &= " And Left(Razao.Conta_Id, 7) In " & VSelecao
            If VIsolarLotes <> "" Then sqlPorContaProduto &= " And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            If VNivelGP = 0 Then
                sqlPorContaProduto &= " GROUP BY Produtos.Nome, Produtos.Produto_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PXG.Titulo " & vbCrLf
            Else
                sqlPorContaProduto &= " GROUP BY GE.Descricao, ge.Grupo_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, PXG.Titulo " & vbCrLf
            End If

            sqlPorContaProduto &= "      ) consulta  " & vbCrLf & _
                                  " Where produto_id is not null " & vbCrLf

            If rdTodasAsContas.Checked Then
                sqlPorContaProduto &= "    And (Debitos <> 0 or Creditos <> 0 or SaldoAnterior <> 0) " & vbCrLf
            ElseIf rdComMovimentoNoPeriodo.Checked Then
                sqlPorContaProduto &= " and (Debitos <> 0 or Creditos <> 0)" & vbCrLf
            ElseIf rdSomenteComSaldoAtual.Checked Then
                sqlPorContaProduto &= " and (SaldoAnterior + abs(Debitos) - abs(Creditos) <> 0)" & vbCrLf
            End If

            sqlPorContaProduto &= " group by  nome, produto_Id,conta_id,titulo, DescGrupo  " & vbCrLf & _
                                  " order by Grupo, Produto_Id, conta_Id " & vbCrLf

            Ds_DebitosCreditos1 = Banco.ConsultaDataSet(sqlPorContaProduto, "Razao")

            tbBalanceteCentro.Columns.Add("Nome", GetType(String))
            tbBalanceteCentro.Columns.Add("Produto_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("Conta_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("Titulo", GetType(String))
            tbBalanceteCentro.Columns.Add("ContaGrupo_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("TituloGrupo", GetType(String))
            tbBalanceteCentro.Columns.Add("SaldoAnterior", GetType(Decimal))
            tbBalanceteCentro.Columns.Add("Debitos", GetType(Decimal))
            tbBalanceteCentro.Columns.Add("Creditos", GetType(Decimal))
            tbBalanceteCentro.Columns.Add("Grupo", GetType(String))
            tbBalanceteCentro.Columns.Add("DescGrupo", GetType(String))
            tbBalanceteCentro.Columns.Add("Conta_Id2", GetType(String))
            Ds_AnexoProduto.Tables.Add(tbBalanceteCentro)
            Ds_AnexoProduto.Merge(Ds_DebitosCreditos1)

            sqlplano = "SELECT Conta_Id, Titulo "
            sqlplano &= "FROM PlanoDeContas "
            sqlplano &= "WHERE LEN(Conta_Id) = 9 "
            sqlplano &= "And Produto = 'S'"

            Ds_PlanoDeContas = Banco.ConsultaDataSet(sqlplano, "PlanoDeContas")

            For Each drAnexo In Ds_AnexoProduto.Tables(0).Rows
                For Each drConta In Ds_PlanoDeContas.Tables(0).Rows
                    If Trim(drAnexo(2)) = Trim(drConta(0)) Then
                        drAnexo(4) = drConta(0)
                        drAnexo(5) = drConta(1)
                    End If
                Next drConta
            Next drAnexo

            For Each drAnexo In Ds_AnexoProduto.Tables(0).Rows
                If IsDBNull(drAnexo(6)) = True Then
                    drAnexo(6) = 0
                End If

                If IsDBNull(drAnexo(7)) = True Then
                    drAnexo(7) = 0
                End If

                If IsDBNull(drAnexo(8)) = True Then
                    drAnexo(8) = 0
                End If

                If IsDBNull(drAnexo(4)) = True And IsDBNull(drAnexo(5)) = True Then
                    drAnexo.Delete()
                End If
            Next drAnexo

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", "Relatório De Anexo Por Contas/Produtos.")
            parameters.Add("Parametros", ParametrosRelatorio)
            parameters.Add("IniciarPagina", txtIniciarNaFolha.Text)

            Funcoes.BindReport(Me.Page, Ds_AnexoProduto, "Cr_AnexoPorProduto", IIf(pTipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(pTipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub GeraAnexoPorProdutoConta(ByVal ParametrosRelatorio As String, ByVal ParametrosEmpresa As ArrayList, pTipo As eTipoSaidaRelatorio)
        Try
            Dim sqlPorProdutoConta As String
            Dim Emp As String() = ddlEmpresa.SelectedValue.Split("-")
            Dim sqlplano As String
            Dim DS_Anterior1 As New DataSet
            Dim Ds_DebitosCreditos1 As DataSet
            Dim tbBalanceteCentro As DataTable = New DataTable("Razao")
            Dim Ds_AnexoProduto As New DataSet
            Dim Ds_PlanoDeContas As New DataSet
            Dim drConta As DataRow
            Dim drAnexo As DataRow
            Dim cont As Integer = 1
            Dim RetornoProdutos As ArrayList

            sqlPorProdutoConta = "select nome, produto_Id,conta_id,titulo,sum(saldoanterior) as saldoanterior,sum(debitos) as debitos,sum(creditos) as creditos " & vbCrLf & _
                                 "  from( " & vbCrLf

            If VNivelGP = 0 Then
                sqlPorProdutoConta &= "        SELECT Produtos.Nome, Produtos.Produto_Id,"
            Else
                sqlPorProdutoConta &= "        SELECT GE.Descricao as Nome, ge.Grupo_Id as Produto_Id,"
            End If


            sqlPorProdutoConta &= "        PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, " & vbCrLf & _
                                  "        SUM(Razao.DebitoOficial - Razao.CreditoOficial) AS SaldoAnterior, 0.00 as Debitos, 0.00 as Creditos " & vbCrLf & _
                                  "   FROM Razao " & vbCrLf & _
                                  "   LEFT JOIN Produtos " & vbCrLf & _
                                  "     ON Razao.Produto = Produtos.Produto_Id " & vbCrLf & _
                                  "  inner join GruposDeEstoques GE " & vbCrLf

            Select Case VNivelGP
                Case 0 : sqlPorProdutoConta &= "     ON Produtos.Grupo         = ge.Grupo_Id"
                Case 1 : sqlPorProdutoConta &= "     ON left(Produtos.Grupo,1) = ge.Grupo_Id"
                Case 2 : sqlPorProdutoConta &= "     ON left(Produtos.Grupo,2) = ge.Grupo_Id"
                Case 3 : sqlPorProdutoConta &= "     ON left(Produtos.Grupo,3) = ge.Grupo_Id"
                Case 5 : sqlPorProdutoConta &= "     ON Produtos.Grupo         = ge.Grupo_Id"
            End Select

            sqlPorProdutoConta &= "  INNER JOIN PlanoDeContas " & vbCrLf & _
                                  "     ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                                  "  WHERE Razao.Movimento_Id < '" & txtDataInicial.Text.ToSqlDate() & "' " & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                sqlPorProdutoConta &= " AND " & RetornoProdutos(0)
                ParametrosRelatorio &= RetornoProdutos(1)
            End If

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sqlPorProdutoConta &= "           And isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') = '" & ddlUnidade.SelectedValue & "'"
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqlPorProdutoConta &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sqlPorProdutoConta &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sqlPorProdutoConta &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                                          "           And Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            sqlPorProdutoConta &= "AND    (Left(PlanoDeContas.Conta_Id, 1) In" & VGrupos & ") "
            If VSelecao <> "" Then sqlPorProdutoConta &= " And Left(Razao.Conta_Id, 7) In " & VSelecao
            If VIsolarLotes <> "" Then sqlPorProdutoConta &= " And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            If VNivelGP = 0 Then
                sqlPorProdutoConta &= " GROUP BY Produtos.Nome, Produtos.Produto_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo " & vbCrLf
            Else
                sqlPorProdutoConta &= " GROUP BY GE.Descricao, ge.Grupo_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo " & vbCrLf
            End If

            sqlPorProdutoConta &= " Union All " & vbCrLf
            cont = 1

            If VNivelGP = 0 Then
                sqlPorProdutoConta &= "        SELECT Produtos.Nome, Produtos.Produto_Id,"
            Else
                sqlPorProdutoConta &= "        SELECT GE.Descricao as Nome, ge.Grupo_Id as Produto_Id,"
            End If

            sqlPorProdutoConta &= "       PlanoDeContas.Conta_Id, PlanoDeContas.Titulo, 0.00 as SaldoAnterior, " & vbCrLf & _
                                  "       SUM(Razao.DebitoOficial) AS Debitos, SUM(Razao.CreditoOficial) AS Creditos " & vbCrLf & _
                                  "  FROM Razao " & vbCrLf & _
                                  "  LEFT JOIN Produtos " & vbCrLf & _
                                  "    ON Razao.Produto = Produtos.Produto_Id " & vbCrLf & _
                                  " inner join GruposDeEstoques GE " & vbCrLf

            Select Case VNivelGP
                Case 0 : sqlPorProdutoConta &= "    ON Produtos.Grupo         = ge.Grupo_Id"
                Case 1 : sqlPorProdutoConta &= "    ON left(Produtos.Grupo,1) = ge.Grupo_Id"
                Case 2 : sqlPorProdutoConta &= "    ON left(Produtos.Grupo,2) = ge.Grupo_Id"
                Case 3 : sqlPorProdutoConta &= "    ON left(Produtos.Grupo,3) = ge.Grupo_Id"
                Case 5 : sqlPorProdutoConta &= "    ON Produtos.Grupo         = ge.Grupo_Id"
            End Select

            sqlPorProdutoConta &= " INNER JOIN PlanoDeContas " & vbCrLf & _
                                  "    ON Razao.Conta_Id = PlanoDeContas.Conta_Id " & vbCrLf & _
                                  " WHERE LEN(Razao.Conta_Id) = 9" & vbCrLf

            If ucSelecaoProduto.TemSelecionado Then
                RetornoProdutos = ucSelecaoProduto.GetSqlEParametrosRelatorio("Produtos.Grupo", "Produtos.Produto_Id", "")
                sqlPorProdutoConta &= " AND " & RetornoProdutos(0)
            End If

            If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then
                sqlPorProdutoConta &= "           And isnull(Razao.UnidadeDeNegocio,'" & ddlUnidade.SelectedValue & "') ='" & ddlUnidade.SelectedValue & "'"
            End If

            If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                If chkConsolidarEmpresa.Checked Then
                    sqlPorProdutoConta &= "           And left(Razao.Empresa_Id,8) = '" & Emp(0).Substring(0, 8) & "'" & vbCrLf
                ElseIf chkConsolidarSede.Checked Then
                    sqlPorProdutoConta &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf
                Else
                    sqlPorProdutoConta &= "           And Razao.Empresa_Id    ='" & Emp(0) & "'" & vbCrLf & _
                                          "           And Razao.EndEmpresa_Id = " & Emp(1) & vbCrLf
                End If
            End If

            sqlPorProdutoConta &= "AND Razao.Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

            sqlPorProdutoConta &= "AND    (Left(PlanoDeContas.Conta_Id, 1) In" & VGrupos & ") "
            If VSelecao <> "" Then sqlPorProdutoConta &= " And Left(Razao.Conta_Id, 7) In " & VSelecao
            If VIsolarLotes <> "" Then sqlPorProdutoConta &= " And (year(movimento_Id) = " & CDate(txtDataInicial.Text).Year & " and Lote_Id Not In " & VIsolarLotes & ")"

            If VNivelGP = 0 Then
                sqlPorProdutoConta &= " GROUP BY Produtos.Nome, Produtos.Produto_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo" & vbCrLf
            Else
                sqlPorProdutoConta &= " GROUP BY GE.Descricao, ge.Grupo_Id, PlanoDeContas.Conta_Id, PlanoDeContas.Titulo " & vbCrLf
            End If

            sqlPorProdutoConta &= "      ) consulta " & vbCrLf & _
                                  " Where produto_id is not null" & vbCrLf

            If rdTodasAsContas.Checked Then
                sqlPorProdutoConta &= "    And (Debitos <> 0 or Creditos <> 0 or SaldoAnterior <> 0) " & vbCrLf
            ElseIf rdComMovimentoNoPeriodo.Checked Then
                sqlPorProdutoConta &= " and (Debitos <> 0 or Creditos <> 0)" & vbCrLf
            ElseIf rdSomenteComSaldoAtual.Checked Then
                sqlPorProdutoConta &= " and (SaldoAnterior + abs(Debitos) - abs(Creditos) <> 0)" & vbCrLf
            End If

            sqlPorProdutoConta &= " group by  nome, produto_Id,conta_id,titulo " & vbCrLf

            Ds_DebitosCreditos1 = Banco.ConsultaDataSet(sqlPorProdutoConta, "Razao")

            tbBalanceteCentro.Columns.Add("Nome", GetType(String))
            tbBalanceteCentro.Columns.Add("Produto_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("Conta_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("Titulo", GetType(String))
            tbBalanceteCentro.Columns.Add("ContaGrupo_Id", GetType(String))
            tbBalanceteCentro.Columns.Add("TituloGrupo", GetType(String))
            tbBalanceteCentro.Columns.Add("SaldoAnterior", GetType(Decimal))
            tbBalanceteCentro.Columns.Add("Debitos", GetType(Decimal))
            tbBalanceteCentro.Columns.Add("Creditos", GetType(Decimal))
            Ds_AnexoProduto.Tables.Add(tbBalanceteCentro)
            Ds_AnexoProduto.Merge(Ds_DebitosCreditos1)

            sqlplano = "SELECT Conta_Id, Titulo "
            sqlplano &= "FROM PlanoDeContas "
            sqlplano &= "WHERE LEN(Conta_Id) = 9 "
            sqlplano &= "And Produto = 'S'"

            Ds_PlanoDeContas = Banco.ConsultaDataSet(sqlplano, "PlanoDeContas")

            For Each drAnexo In Ds_AnexoProduto.Tables(0).Rows
                For Each drConta In Ds_PlanoDeContas.Tables(0).Rows
                    If Trim(drAnexo(2)) = Trim(drConta(0)) Then
                        drAnexo(4) = drConta(0)
                        drAnexo(5) = drConta(1)
                    End If
                Next drConta
            Next drAnexo

            For Each drAnexo In Ds_AnexoProduto.Tables(0).Rows
                If IsDBNull(drAnexo(6)) = True Then
                    drAnexo(6) = 0
                End If

                If IsDBNull(drAnexo(7)) = True Then
                    drAnexo(7) = 0
                End If

                If IsDBNull(drAnexo(8)) = True Then
                    drAnexo(8) = 0
                End If

                If IsDBNull(drAnexo(4)) = True And IsDBNull(drAnexo(5)) = True Then
                    drAnexo.Delete()
                End If
            Next drAnexo

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("Titulo", "Relatório De Anexo Por Produto Conta.")
            parameters.Add("Parametros", ParametrosRelatorio)
            parameters.Add("IniciarPagina", txtIniciarNaFolha.Text)

            Funcoes.BindReport(Me.Page, Ds_AnexoProduto, "Cr_AnexoPorProdutoConta", IIf(pTipo = eTipoSaidaRelatorio.PDF, eExportType.PDF, IIf(pTipo = eTipoSaidaRelatorio.EXCEL, eExportType.ExcelCrystal, eExportType.ExcelCrystalDados)), parameters)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub VerGrupos()
        VGrupos = "('"
        Dim i As Integer
        For i = 0 To Grupos.Items.Count - 1
            If Grupos.Items(i).Selected Then
                If Len(VGrupos) > 2 Then
                    VGrupos &= ",'" & Grupos.Items(i).Value() & "'"
                Else
                    VGrupos &= Grupos.Items(i).Value() & "'"
                End If
            End If
        Next
        If Len(VGrupos) > 2 Then
            VGrupos &= ")"
        End If
    End Sub

    Private Sub VerNivel()
        VNivel = 5
        VNiveis = ""
        Dim i As Integer
        For i = 0 To Niveis.Items.Count - 1
            If Niveis.Items(i).Selected Then
                If VNiveis.Length > 0 Then VNiveis &= ","
                Select Case i
                    Case 0 : VNiveis &= "1"
                    Case 1 : VNiveis &= "3"
                    Case 2 : VNiveis &= "5"
                    Case 3 : VNiveis &= "7"
                    Case 4 : VNiveis &= "9"
                End Select
                VNivel = Niveis.Items(i).Value()
            End If
        Next
    End Sub

    Private Sub VerNivelGrupoDeEstoque()
        VNivelGP = 0
        Dim i As Integer
        For i = 0 To rdGrupoProduto.Items.Count - 1
            If rdGrupoProduto.Items(i).Selected Then
                VNivelGP = rdGrupoProduto.Items(i).Value()
            End If
        Next
    End Sub

    Private Sub VerNivelDeCusto()
        VNivelDeCusto = 5
        Dim i As Integer
        For i = 0 To NiveisDeCusto.Items.Count - 1
            If NiveisDeCusto.Items(i).Selected Then
                VNivelDeCusto = NiveisDeCusto.Items(i).Value()
            End If
        Next
    End Sub

    Private Sub VerModelo()
        VModelo = 1
        Dim i As Integer
        For i = 0 To Modelo.Items.Count - 1
            If Modelo.Items(i).Selected Then
                VModelo = Modelo.Items(i).Value()
            End If
        Next
    End Sub

    Private Sub VerMoeda()
        VMoeda = ddlMoeda.SelectedValue
    End Sub

    Private Sub VerIsolarLotes()
        VIsolarLotes = "("
        Dim i As Integer
        For i = 0 To IsolarLotes.Items.Count - 1
            If IsolarLotes.Items(i).Selected Then
                If Len(VIsolarLotes) > 1 Then
                    VIsolarLotes &= "," & IsolarLotes.Items(i).Value()
                Else
                    VIsolarLotes &= IsolarLotes.Items(i).Value()
                End If
            End If
        Next
        If Len(VIsolarLotes) > 1 Then
            VIsolarLotes &= ")"
        Else
            VIsolarLotes = ""
        End If
    End Sub

    Private Sub VerIsolarCompensacao()
        Dim controle As Integer
        VIsolarCompensacao = ""
        If IsolarCompensacao.Items(controle).Selected Then
            VIsolarCompensacao = "('"
        End If


        Dim i As Integer
        For i = 0 To IsolarCompensacao.Items.Count - 1
            If IsolarCompensacao.Items(i).Selected Then
                If Len(VIsolarCompensacao) > 2 Then
                    VIsolarCompensacao &= ",'" & IsolarCompensacao.Items(i).Value() & "'"
                Else
                    VIsolarCompensacao &= IsolarCompensacao.Items(i).Value() & "'"
                End If
            End If
        Next
        If Len(VIsolarCompensacao) > 1 Then
            VIsolarCompensacao &= ")"
        End If
    End Sub

    Private Sub VerSelecaoDeGrupos()
        VSelecao = "('"
        Dim i As Integer = 0
        For i = 0 To SelecaoDeGrupos1.Items.Count - 1
            If SelecaoDeGrupos1.Items(i).Selected Then
                If Len(VSelecao) > 2 Then
                    VSelecao &= ",'" & SelecaoDeGrupos1.Items(i).Value() & "'"
                Else
                    VSelecao &= SelecaoDeGrupos1.Items(i).Value() & "'"
                End If
            End If
        Next
        i = 0
        For i = 0 To SelecaoDeGrupos2.Items.Count - 1
            If SelecaoDeGrupos2.Items(i).Selected Then
                If Len(VSelecao) > 2 Then
                    VSelecao &= ",'" & SelecaoDeGrupos2.Items(i).Value() & "'"
                Else
                    VSelecao &= SelecaoDeGrupos2.Items(i).Value() & "'"
                End If
            End If
        Next

        If Len(VSelecao) > 2 Then
            VSelecao &= ")"
        Else
            VSelecao = ""
        End If
    End Sub

    Private Sub Parametros()
        VerGrupos()
        VerNivelGrupoDeEstoque()
        VerNivel()
        VerNivelDeCusto()
        VerModelo()
        VerMoeda()
        VerIsolarLotes()
        VerIsolarCompensacao()
        VerSelecaoDeGrupos()
    End Sub

    Protected Sub Modelo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles Modelo.SelectedIndexChanged
        Try
            If Modelo.SelectedValue = 2 Then
                pnlCentroDeCusto.Parent.Visible = True

                NiveisDeCusto.Enabled = True
                VerCentroDeCusto()
            Else
                pnlCentroDeCusto.Parent.Visible = False
            End If

            If Modelo.SelectedValue = 1 Or Modelo.SelectedValue = 5 Then
                chkZeraContaDeResultado.Visible = True
                chkpiscofins56.Visible = True
                txtSituacaoPISCOFINS.Visible = True
            Else
                chkZeraContaDeResultado.Visible = False
                chkpiscofins56.Visible = False
                txtSituacaoPISCOFINS.Visible = False
            End If

            If Modelo.SelectedValue = 4 Or Modelo.SelectedValue = 6 Then
                pnlNivelDeProduto.Visible = True
                ucSelecaoProduto.Parent.Visible = True
            Else
                pnlNivelDeProduto.Visible = False
                ucSelecaoProduto.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub VerCentroDeCusto()
        Dim where As String = ""
        Select Case NiveisDeCusto.SelectedIndex
            Case 0
                where = "  LEN(CentroDeCusto_id) = 1 "
            Case 1
                where = "  LEN(CentroDeCusto_id) = 3 "
            Case 2
                where = "  LEN(CentroDeCusto_id) > 3 "
        End Select

        Dim Lista As New [Lib].Negocio.ListCentroDeCusto(True, where)

        lstCentroDeCusto.Items.Clear()
        lstCentroDeCusto.DataTextField = "Descricao"
        lstCentroDeCusto.DataValueField = "Codigo"

        For Each row As CentroDeCusto In Lista
            lstCentroDeCusto.Items.Add(New ListItem(row.Codigo & " - " & row.Descricao, row.Codigo))
        Next
    End Sub

    Private Sub Limpar()
        'Session("ObjEmpresa") = Nothing
        SelecaoDeGrupos1.Items.Clear()
        SelecaoDeGrupos2.Items.Clear()
        CargaGrupos()

        txtDataInicial.Text = Format(Today, "01/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
        txtDataDeEmissao.Text = Format(Today, "dd/MM/yyyy")
        txtIniciarNaFolha.Text = 1
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub NiveisDeCusto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            VerCentroDeCusto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgAdiciona_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs) Handles imgAdiciona.Click
        Try
            If lstCentroCusto.SelectedIndex = -1 Then Exit Sub

            Dim item As New ListItem

            item.Text = lstCentroCusto.SelectedItem.Text
            item.Value = lstCentroCusto.SelectedItem.Value

            lstCentroCusto.Items.Remove(lstCentroCusto.SelectedItem)
            lstCentroCustoSelecionados.Items.Add(item)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgRemove_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try
            If lstCentroCustoSelecionados.SelectedIndex = -1 Then Exit Sub

            Dim item As New ListItem

            item.Text = lstCentroCustoSelecionados.SelectedItem.Text
            item.Value = lstCentroCustoSelecionados.SelectedItem.Value

            lstCentroCustoSelecionados.Items.Remove(lstCentroCustoSelecionados.SelectedItem)
            lstCentroCusto.Items.Add(item)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub EmitirRelatorio(pTipo As eTipoSaidaRelatorio)
        Try
            If CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
                MsgBox(Me.Page, "A data final não pode ser menor que a data inicial")
                Exit Sub
            ElseIf chkpiscofins56.Checked And String.IsNullOrWhiteSpace(txtSituacaoPISCOFINS.Text) Then
                MsgBox(Me.Page, "Digite pelo menos uma Situação Pis/Cofins.")
                txtSituacaoPISCOFINS.Focus()
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("Balancetes", "RELATORIO") Then
                Parametros()

                Dim ParametrosRelatorio As String = "PARAMETROS: " & vbCrLf
                If ddlUnidade.SelectedIndex > 0 AndAlso Not chkConsolidarUnidade.Checked Then ParametrosRelatorio &= "Unidade de Negocio: " & ddlUnidade.SelectedItem.Text & vbCrLf

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    If chkConsolidarEmpresa.Checked Then
                        ParametrosRelatorio &= "Empresa Consolidada: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                    ElseIf chkConsolidarSede.Checked Then
                        ParametrosRelatorio &= " Empresa Consolidada por setor: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                    Else
                        ParametrosRelatorio &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
                    End If
                End If
                ParametrosRelatorio &= "Periodo de: " & CDate(txtDataInicial.Text).ToString("dd/MM/yyyy") & " a " & CDate(txtDataFinal.Text).ToString("dd/MM/yyyy") & vbCrLf
                If rdSomenteComSaldoAtual.Checked Then ParametrosRelatorio &= "Somente Contas com Saldo Atual" & vbCrLf
                If rdComMovimentoNoPeriodo.Checked Then ParametrosRelatorio &= "Somente Contas com Movimento No Periodo" & vbCrLf
                If VIsolarLotes.Length > 0 Then ParametrosRelatorio &= "Isolar Lotes: " & VIsolarLotes & vbCrLf
                If VIsolarCompensacao.Length > 0 Then ParametrosRelatorio &= "Isolar Compensacao: " & VIsolarCompensacao & vbCrLf
                If lstCentroDeCusto.GetSelectedValues().Count > 0 And VModelo = 2 Then ParametrosRelatorio &= "Centro de Custo: " & String.Join(", ", getListCentroDeCusto()) & vbCrLf
                ParametrosRelatorio &= "Moeda: " & ddlMoeda.SelectedItem.Text & vbCrLf

                Dim Par As New ArrayList

                If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
                    Dim objEmp As New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
                    Par.Add(objEmp.Codigo)
                    Par.Add(objEmp.Nome)
                    Par.Add(objEmp.Cidade)
                    Par.Add(objEmp.CodigoEstado)
                Else
                    Par.Add("")
                    Par.Add("")
                    Par.Add("")
                    Par.Add("")
                End If

                Select Case VModelo
                    Case 1
                        GeraBalanceteAuxiliar(ParametrosRelatorio, Par, pTipo)
                    Case 2
                        GeraCentrosDeCustos(ParametrosRelatorio, Par, pTipo)
                    Case 3
                        GeraAnexoPorCliente(ParametrosRelatorio, Par, pTipo)
                    Case 4
                        GeraAnexoPorContaProduto(ParametrosRelatorio, Par, pTipo)
                    Case 5
                        GeraBalanceteAuxiliarCompleto(ParametrosRelatorio, Par, pTipo)
                    Case 6
                        GeraAnexoPorProdutoConta(ParametrosRelatorio, Par, pTipo)
                End Select
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório!")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkPdf_Click(sender As Object, e As EventArgs) Handles lnkPdf.Click
        EmitirRelatorio(eTipoSaidaRelatorio.PDF) 'Pdf
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        EmitirRelatorio(eTipoSaidaRelatorio.EXCEL) 'Excel
    End Sub

    Protected Sub lnkExcelDados_Click(sender As Object, e As EventArgs) Handles lnkExcelDados.Click
        EmitirRelatorio(eTipoSaidaRelatorio.EXCEL_DADOS) 'Excel Dados
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Balancetes")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class