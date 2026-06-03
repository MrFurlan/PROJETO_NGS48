Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class CurvaABC
    Inherits BasePage

    Dim ListGrupo As [Lib].Negocio.ListGrupoProduto
    Dim ParametrosDaConsulta As New StringBuilder

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CurvaABC", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    ddl.Carregar(ddlAno, CarregarDDL.Tabela.Ano, "2009;10;C", False)
                    ddl.Carregar(ddlMarca, CarregarDDL.Tabela.Marca, "", True)
                    ddlEmpresa.SelectedValue = HttpContext.Current.Session("ssEmpresa") & "-" & HttpContext.Current.Session("ssEndEmpresa")

                    txtDataInicial.Text = "01/" & Today.ToString("MM") & "/" & Today.ToString("yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getDataSetCurvaABC() As DataSet
        Dim TamGrupo As Integer
        Dim sql As String = ""

        If ListGrupo Is Nothing Then
            TamGrupo = 5
        Else
            TamGrupo = ListGrupo(0).Codigo.Length
        End If

        sql = " SELECT ROW_NUMBER( ) OVER (order by sum(case" & vbCrLf & _
               "                                          when SO.devolucao = 'S'" & vbCrLf & _
               "                                            then case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then NxE.Valor      else NxE.Valor * - 1 end" & vbCrLf & _
               "                                            else case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then NxE.Valor * -1 else NxE.Valor       end" & vbCrLf & _
               "                                        end) Desc) as NumLinha, " & vbCrLf

        If rdCliente.Checked Then
            If chkConsolidarCliente.Checked Then
                sql &= "        left(NF.Cliente_Id,8) as Cliente_Id, " & vbCrLf & _
                       "        0 as EndCliente_Id," & vbCrLf & _
                       "        (Select top 1 Nome from clientes where left(clientes.cliente_Id,8) = left(NF.Cliente_Id,8)) as Nome,  " & vbCrLf & _
                       "        '' as Bairro," & vbCrLf & _
                       "        '' as Cidade," & vbCrLf & _
                       "        '' as Estado," & vbCrLf
            Else
                sql &= "        NF.Cliente_Id, " & vbCrLf & _
                       "        NF.EndCliente_Id, " & vbCrLf & _
                       "        C.Nome, " & vbCrLf & _
                       "        C.Bairro," & vbCrLf & _
                       "        C.Cidade," & vbCrLf & _
                       "        C.Estado AS estado," & vbCrLf
            End If

        Else
            sql &= "        GE.Grupo_Id AS Grupo," & vbCrLf & _
                   "        GE.Descricao as DescGrupo," & vbCrLf & _
                   "        Prd.Produto_Id as Produto," & vbCrLf & _
                   "        Prd.Nome," & vbCrLf
        End If

        sql &= "        COUNT(NF.Nota_Id) AS Notas," & vbCrLf

        If rbReal.Checked Then
            sql &= "        sum(case" & vbCrLf & _
                   "              when SO.devolucao = 'S'" & vbCrLf & _
                   "                then case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then NxE.Valor      else NxE.Valor * - 1 end" & vbCrLf & _
                   "                else case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then NxE.Valor * -1 else NxE.Valor       end" & vbCrLf & _
                   "            end) Total," & vbCrLf
        Else
            sql &= "        sum(case" & vbCrLf & _
                   "              when SO.devolucao = 'S'" & vbCrLf & _
                   "                then case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then round((NxE.Valor / Co.Indice),2)      else round((NxE.Valor * - 1 / Co.Indice),2) end" & vbCrLf & _
                   "                else case when nxe.encargo_id in ('PIS','COFINS','ISS'" & IIf(rdVendas.Checked, ",'ICMS'", "") & ") then round((NxE.Valor * -1 / Co.Indice),2) else round((NxE.Valor / Co.Indice),2)       end" & vbCrLf & _
                   "            end) Total," & vbCrLf
        End If

        sql &= "       convert(numeric(18,2),0) as Acumulado " & vbCrLf & _
               "  into #Temp" & vbCrLf & _
               "  FROM NotasFiscais NF" & vbCrLf & _
               " INNER JOIN Pedidos P" & vbCrLf & _
               "    ON P.Empresa_id    = NF.Empresa_Id" & vbCrLf & _
               "   And P.EndEmpresa_Id = NF.EndEmpresa_Id" & vbCrLf & _
               "   And P.Pedido_Id     = NF.Pedido" & vbCrLf & _
               " INNER JOIN Cotacoes Co" & vbCrLf & _
               "    ON Co.Indexador_id = 3" & vbCrLf & _
               "   And Co.Data_Id      = NF.Movimento" & vbCrLf & _
               " INNER JOIN Clientes C" & vbCrLf & _
               "    ON NF.Cliente_Id    = C.Cliente_Id" & vbCrLf & _
               "   AND NF.EndCliente_Id = C.Endereco_Id" & vbCrLf & _
               " INNER JOIN NotasFiscaisXItens NFxI" & vbCrLf & _
               "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               " INNER Join NotasFiscaisXEncargos NxE" & vbCrLf & _
               "    ON NxE.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf & _
               "   AND NxE.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf & _
               "   AND NxE.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf & _
               "   AND NxE.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf & _
               "   AND NxE.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf & _
               "   AND NxE.Serie_Id        = NFxI.Serie_Id" & vbCrLf & _
               "   AND NxE.Nota_Id         = NFxI.Nota_Id" & vbCrLf & _
               "   AND NxE.Produto_Id      = NFxI.Produto_Id" & vbCrLf & _
               "   AND NxE.cfop_Id         = NFxI.cfop_Id" & vbCrLf & _
               "   AND NxE.Sequencia_Id    = NFxI.Sequencia_Id" & vbCrLf & _
               "   AND NxE.Encargo_Id   " & IIf(rbBruto.Checked, "= 'PRODUTO'", IIf(rbLiquido.Checked, "='LIQUIDO'", IIf(rdCompras.Checked, " in ('PRODUTO','PIS','IPI','COFINS','ISS')", " in ('PRODUTO','PIS','IPI','COFINS','ISS','ICMS')"))) & vbCrLf & _
               " INNER JOIN Produtos Prd" & vbCrLf & _
               "    ON NFxI.Produto_Id = Prd.Produto_Id" & vbCrLf & _
               " Inner Join GruposDeEstoques GE" & vbCrLf & _
               "    on GE.Grupo_Id = Left(Prd.Grupo," & TamGrupo & ")" & vbCrLf & _
               " Inner Join Operacoes OP" & vbCrLf & _
               "    on op.operacao_id = NFxI.Operacao" & vbCrLf & _
               " INNER JOIN SubOperacoes SO" & vbCrLf & _
               "    ON NFxI.Operacao    = SO.Operacao_Id" & vbCrLf & _
               "   AND NFxI.SubOperacao = SO.SubOperacoes_Id" & vbCrLf & _
               " WHERE NF.situacao      = 1" & vbCrLf

        If rdPeriodo.Checked Then
            sql &= "   AND NF.Movimento between '" & CDate(txtDataInicial.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataFinal.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
        Else
            sql &= "   And P.Safra in (Select Safra_Id from safras where year(vencimento) = " & ddlAno.SelectedValue & ")" & vbCrLf
        End If

        If rdTodosMercados.Checked Then
            sql &= "   AND OP.Classe   in " & IIf(rdCompras.Checked, "('" & eClassesOperacoes.COMPRAS.ToString & "','" & eClassesOperacoes.IMPORTACOES.ToString & "')", "('" & eClassesOperacoes.VENDAS.ToString & "','" & eClassesOperacoes.EXPORTACOES.ToString & "')") & vbCrLf
        ElseIf rdInterno.Checked Then
            sql &= "   AND OP.Classe    = " & IIf(rdCompras.Checked, "'" & eClassesOperacoes.COMPRAS.ToString & "'", "'" & eClassesOperacoes.VENDAS.ToString & "'") & vbCrLf
        ElseIf rdExterno.Checked Then
            sql &= "   AND OP.Classe    = " & IIf(rdCompras.Checked, "'" & eClassesOperacoes.IMPORTACOES.ToString & "'", "'" & eClassesOperacoes.EXPORTACOES.ToString & "'") & vbCrLf
        End If

        sql &= "   AND SO.classe  not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.CONTAEORDEM.ToString & "')" & vbCrLf

        If ucSelecaoProdutoABC.TemSelecionado Then
            Dim par As ArrayList
            par = ucSelecaoProdutoABC.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")

            sql &= "   AND " & par(0) & vbCrLf
        End If


        Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")
        If chkConsolidar.Checked Then
            sql &= "   AND (left(NF.Empresa_Id,8)  = '" & emp(0).Substring(0, 8) & "')" & vbCrLf
        Else
            sql &= "   AND NF.Empresa_Id     ='" & emp(0) & "'" & vbCrLf & _
                   "   AND NF.EndEmpresa_Id  = " & emp(1)
        End If


        If rdCliente.Checked Then
            If chkConsolidarCliente.Checked Then
                sql &= " GROUP BY left(NF.Cliente_Id,8)" & vbCrLf
            Else
                sql &= " GROUP BY NF.Cliente_Id, NF.EndCliente_Id, C.Nome, C.Bairro, C.Cidade, C.Estado" & vbCrLf
            End If
        Else
            sql &= " GROUP BY GE.Grupo_Id, GE.Descricao, Prd.Produto_Id, Prd.Nome" & vbCrLf
        End If

        sql &= " Update #Temp set " & vbCrLf & _
               "   Acumulado = (Select sum(T2.Total)" & vbCrLf & _
               "                  from #Temp T2" & vbCrLf & _
               "                 where T2.NumLinha <= #Temp.NumLinha )" & vbCrLf & _
               " select sum(Total) as TotalGeral," & vbCrLf & _
               "        sum(Total) * " & Str((ddlCurvaA.SelectedValue / 100)) & " as TotalA," & vbCrLf & _
               "        sum(Total) * " & Str((ddlCurvaB.SelectedValue / 100)) & " as TotalB," & vbCrLf & _
               "        sum(Total) * " & Str((ddlCurvaC.SelectedValue / 100)) & " as TotalC " & vbCrLf & _
               "   into #Totais" & vbCrLf & _
               "   from #temp" & vbCrLf


        If rdCliente.Checked Then
            sql &= " select T.Cliente_Id, T.EndCliente_Id, T.Nome, T.Bairro," & vbCrLf & _
                   "        T.Cidade, T.Estado, T.Notas, T.Total, T.Acumulado," & vbCrLf & _
                   "        round(T.Total     * 100 / TotalGeral,2) as PercIndividual," & vbCrLf & _
                   "        round(T.Acumulado * 100 / TotalGeral,2) as PercAcumuludo," & vbCrLf
        Else
            sql &= "select T.Grupo, T.DescGrupo, T.Nome, T.Produto," & vbCrLf & _
                   "       T.Notas, T.Total, T.Acumulado, " & vbCrLf & _
                   "        round(T.Total     * 100 / TotalGeral,2) as PercAcumuludo," & vbCrLf & _
                   "        round(T.Acumulado * 100 / TotalGeral,2) as PercIndividual," & vbCrLf
        End If

        sql &= "        case" & vbCrLf & _
            "          when Acumulado < TotalA then 'A'" & vbCrLf & _
            "          when Acumulado < TotalA + TotalB then 'B'" & vbCrLf & _
            "          else 'C' " & vbCrLf & _
            "        End ABC" & vbCrLf & _
            "   Into #curva" & vbCrLf & _
            "   from #Temp T" & vbCrLf & _
            "  inner join #Totais" & vbCrLf & _
            "     on 1 = 1 " & vbCrLf & _
            "  order by Total desc;" & vbCrLf & _
            " Select * from #Curva;" & vbCrLf


        sql &= "select SUM(case" & vbCrLf & _
               "                   when ABC = 'A'" & vbCrLf & _
               "                     Then 1.00" & vbCrLf & _
               "                     else 0.00" & vbCrLf & _
               "                  end) * 100 / COUNT(*) as PercA," & vbCrLf & _
               "       SUM(case" & vbCrLf & _
               "                   when ABC = 'B'" & vbCrLf & _
               "                     Then 1.00" & vbCrLf & _
               "                     else 0.00" & vbCrLf & _
               "                 end) * 100 / COUNT(*) as PercB," & vbCrLf & _
               "       SUM(case" & vbCrLf & _
               "                   when ABC = 'C'" & vbCrLf & _
               "                     Then 1.00" & vbCrLf & _
               "                     else 0.00" & vbCrLf & _
               "                 end) * 100 / COUNT(*) as PercC," & vbCrLf & _
               "       T.TotalA," & vbCrLf & _
               "       T.TotalB," & vbCrLf & _
               "       T.TotalC," & vbCrLf & _
               "       T.TotalGeral" & vbCrLf & _
               "  from #curva C" & vbCrLf & _
               " inner join #Totais T" & vbCrLf & _
               "    on 1 = 1" & vbCrLf & _
               " group by T.TotalA,T.TotalB,T.TotalC, T.TotalGeral" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "Consulta")
    End Function

    Private Sub Parametros()
        ParametrosDaConsulta.AppendLine("Parâmetros da Consulta")

        'Empresa consolidada
        If chkConsolidar.Checked Then
            ParametrosDaConsulta.AppendLine("Empresa Consolidada.")
        End If

        'Período
        If rdPeriodo.Checked Then
            ParametrosDaConsulta.AppendLine("Período: " & CDate(txtDataInicial.Text).ToString("dd/MM/yyyy") & " a " & CDate(txtDataFinal.Text).ToString("dd/MM/yyyy"))
        End If

        'Ano Safra
        If rdAnoSafra.Checked Then
            ParametrosDaConsulta.AppendLine("Ano Safra: " & ddlAno.SelectedValue)
        End If

        'Posição
        ParametrosDaConsulta.Append("Posição: ")
        If rdCliente.Checked Then
            If chkConsolidarCliente.Checked Then
                ParametrosDaConsulta.AppendLine("Por Cliente Consolidado")
            Else
                ParametrosDaConsulta.AppendLine("Por Cliente")
            End If
        Else
            ParametrosDaConsulta.AppendLine("Por Produto")
        End If

        'Operações
        ParametrosDaConsulta.AppendLine("Operações: " & IIf(rdCompras.Checked, "Compras", "Vendas"))

        'Mercado
        ParametrosDaConsulta.Append("Mercado: ")
        If rdTodosMercados.Checked Then
            ParametrosDaConsulta.AppendLine("Todos Marcados")
        ElseIf rdInterno.Checked Then
            ParametrosDaConsulta.AppendLine("Interno")
        ElseIf rdExterno.Checked Then
            ParametrosDaConsulta.AppendLine("Externo")
        End If

        'Produtos
        If ucSelecaoProdutoABC.TemSelecionado Then
            Dim par As ArrayList
            par = ucSelecaoProdutoABC.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id")
            ParametrosDaConsulta.AppendLine(par(1))
        End If

        'Curva ABC
        ParametrosDaConsulta.AppendLine("Curva: " & "A: " & ddlCurvaA.SelectedValue & "  B: " & ddlCurvaB.SelectedValue & "  C: " & ddlCurvaC.SelectedValue)
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
            If Funcoes.VerificaPermissao("CurvaABC", "RELATORIO") Then
                If CInt(ddlCurvaA.SelectedValue) + CInt(ddlCurvaB.SelectedValue) + CInt(ddlCurvaC.SelectedValue) > 100 Then
                    MsgBox(Me.Page, "A Soma Da Curva A B C nao pode ser MAIOR que 100.")
                    Exit Sub
                ElseIf rdPeriodo.Checked Then
                    If String.IsNullOrEmpty(txtDataInicial.Text) OrElse String.IsNullOrWhiteSpace(txtDataInicial.Text) Then
                        MsgBox(Me.Page, "Data inicial não foi informada.")
                        Exit Sub
                    ElseIf String.IsNullOrEmpty(txtDataFinal.Text) OrElse String.IsNullOrWhiteSpace(txtDataFinal.Text) Then
                        MsgBox(Me.Page, "Data final não foi informada.")
                        Exit Sub
                    ElseIf CDate(txtDataFinal.Text) < CDate(txtDataInicial.Text) Then
                        MsgBox(Me.Page, "Data final não pode ser menor que data inicial.")
                        Exit Sub
                    End If
                End If

                Parametros()

                Dim ds As DataSet = getDataSetCurvaABC()
                ds.Tables(0).TableName = IIf(rdCliente.Checked, "CurvaABC", "CurvaABCProduto")
                ds.Tables(1).TableName = "Resumo"

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("VP1", ddlCurvaA.SelectedItem.Value.ToString())
                parameters.Add("VP2", ddlCurvaB.SelectedItem.Value.ToString())
                parameters.Add("VP3", ddlCurvaC.SelectedItem.Value.ToString())
                parameters.Add("NomeRelatorio", String.Format("Curva A-B-C de {0} por {1}", IIf(rdCompras.Checked, "Compras", "Vendas"), IIf(rdCliente.Checked, "Cliente", "Produto")))
                parameters.Add("ParametrosDaConsulta", ParametrosDaConsulta.ToString())

                If rdCliente.Checked Then parameters.Add("Moeda", IIf(rbReal.Checked, "Moeda: Reais", "Moeda: Dólar"))

                Funcoes.BindReport(Me.Page, ds, IIf(rdCliente.Checked, "Cr_CurvaABC", "Cr_CurvaABC_Produto"), IIf(Pdf, eExportType.PDF, eExportType.ExcelCrystal), parameters, True)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlMarca_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlMarca.SelectedIndexChanged
        Try
            If ddlMarca.SelectedIndex = 0 Then
                ucSelecaoProdutoABC.WhereProduto = ""
                ucSelecaoProdutoABC.CarregarNivel(1)
            Else
                ucSelecaoProdutoABC.WhereProduto = "Marca = " & ddlMarca.SelectedValue
                ucSelecaoProdutoABC.CarregarNivel(1)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CurvaABC")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class