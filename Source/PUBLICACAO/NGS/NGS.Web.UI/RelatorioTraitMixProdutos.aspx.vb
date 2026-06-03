Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared

Public Class RelatorioTraitMixProdutos
    Inherits BasePage

    Private descricao As String
    Private parametros As String = ""

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Revenda)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("RelatorioTraitMixProdutos", "ACESSAR") Then
                ddl.Carregar(ddlMarcaProd, CarregarDDL.Tabela.Marca, "")
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "")
                ddl.Carregar(ddlSafra, CarregarDDL.Tabela.Safra, "")
                ddl.Carregar(lstClasseOp, CarregarDDL.Tabela.ClasseDeOperacao, "isnull(operacao,0) = 1", False)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Revenda.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Function GetSql() As String
        Dim sql As String = ""
        Dim strSoma As String = ""
        Dim PPeriodo As String = ""
        Dim tipoValor As String = ""

        '************Verificar o coluna a ser somada****************

        If rbQtde.Checked Then
            If rbNota.Checked Then
                strSoma = "QuantidadeFisica"
                descricao = "Qtde"
                tipoValor = "Qtde/Nota"
            Else
                strSoma = "QtdeDeEmbalagem"
                descricao = "Qtde Emb"
                tipoValor = "Qtde/Embalagem"
            End If
        Else
            strSoma = "Valor"
            descricao = "Valor"
            tipoValor = "Valor"
        End If

        sql = "         declare @strSoma decimal(18,2) " & vbCrLf & _
                "       set @strSoma =  0" & vbCrLf & _
                "                           " & vbCrLf & _
                "       SELECT   GE.Grupo_Id," & vbCrLf & _
                "                Ge.Descricao," & vbCrLf & _
                "                Prd.Produto_Id, " & vbCrLf & _
                "                Prd.Nome, " & vbCrLf & _
                "       sum( case when SO.devolucao = 'S' " & vbCrLf & _
                "               then NFI." & strSoma & " * -1" & vbCrLf & _
                "	            else NFI." & strSoma & vbCrLf & _
                "               end) as valor" & vbCrLf & _
                "               into #temp" & vbCrLf & _
                "                                       " & vbCrLf & _
                "       FROM  NotasFiscais NF" & vbCrLf & _
                "       INNER JOIN  NotasFiscaisXItens NFI ON NF.Empresa_Id = NFI.Empresa_Id AND NF.EndEmpresa_Id = NFI.EndEmpresa_Id AND " & vbCrLf & _
                "                   NF.Cliente_Id = NFI.Cliente_Id AND NF.EndCliente_Id = NFI.EndCliente_Id AND " & vbCrLf & _
                "                   NF.EntradaSaida_Id = NFI.EntradaSaida_Id AND NF.Serie_Id = NFI.Serie_Id AND " & vbCrLf & _
                "                   NF.Nota_Id = NFI.Nota_Id " & vbCrLf & _
                "       INNER JOIN Pedidos P ON NF.Empresa_Id = P.Empresa_Id AND NF.EndEmpresa_Id = P.EndEmpresa_Id AND " & vbCrLf & _
                "                   NF.Pedido = P.Pedido_Id " & vbCrLf & _
                "       INNER JOIN Produtos Prd ON NFI.Produto_Id = Prd.Produto_Id " & vbCrLf & _
                "       inner join Gruposdeestoques GE on Prd.grupo = GE.Grupo_ID " & vbCrLf & _
                "       INNER JOIN SubOperacoes SO ON NFI.Operacao = SO.Operacao_Id AND NFI.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
                "       INNER JOIN Operacoes OP ON OP.Operacao_Id =  SO.Operacao_Id " & vbCrLf & _
                "   Where NF.Situacao = 1" & vbCrLf & _
                "       AND SO.Classe <> '" & eClassesOperacoes.GLOBAL.ToString & "'" & vbCrLf

        Dim op As String = String.Join("','", lstClasseOp.GetSelecteds)
        If op.Length > 0 Then
            sql &= "   AND op.classe in ('" & op & "')"
            parametros &= "Classe: " & op
        End If


        If ddlEmpresa.SelectedIndex > 0 Then
            Dim emp As String() = ddlEmpresa.SelectedValue.Split("-")
            parametros &= "Empresa: " & ddlEmpresa.SelectedItem.Text & vbCrLf
            If chkConsolidarEmpresa.Checked Then
                sql &= " and Left(NF.Empresa_id, 8) = '" & emp(0).Substring(0, 8) & "'"

            Else
                sql &= " and NF.Empresa_id = '" & emp(0) & "'"
                sql &= " and NF.EndEmpresa_id = " & emp(1)
            End If
        End If

        If ddlMarcaProd.SelectedIndex > 0 Then
            parametros &= "Marca do Produto: " & ddlMarcaProd.SelectedItem.Text & vbCrLf
        End If

        If ddlSafra.SelectedIndex > 0 Then
            sql &= " and P.Safra = '" & ddlSafra.SelectedValue & "'"
            parametros &= "Safra: " & ddlSafra.SelectedItem.Text & vbCrLf
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim ret As New ArrayList
            ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("Prd.Grupo", "Prd.Produto_Id", "", True)
            sql &= " and " & ret(0) & vbCrLf
            parametros &= "Produtos: " & ret(1)
        End If

        If chkPerido.Checked Then
            If IsDate(txtDataDe.Text) And IsDate(txtDataAte.Text) Then
                sql &= " and NF.DataDaNota between '" & CDate(txtDataDe.Text).ToString("yyyy-MM-dd") & "' and '" & CDate(txtDataAte.Text).ToString("yyyy-MM-dd") & "' "
                parametros &= "Período de: " & txtDataDe.Text & " até " & txtDataAte.Text & vbCrLf

            Else
                MsgBox(Me.Page, "Informe a data correta")
                Return ""
            End If
        End If
        parametros &= "Tipo Soma: " & tipoValor & vbCrLf

        sql &= "group by GE.Grupo_Id,     " & vbCrLf & _
               "         Ge.Descricao,    " & vbCrLf & _
               "         Prd.Produto_Id,  " & vbCrLf & _
               "         Prd.Nome         " & vbCrLf & _
               "                          " & vbCrLf & _
               "set @strSoma = (select SUM(valor) from #temp) " & vbCrLf & _
               "                          " & vbCrLf & _
               "select Produto_Id,	Nome, valor, valor *100 / @strSoma as Participacao, Grupo_Id, Descricao from #temp order by 4" & vbCrLf & _
               "                          " & vbCrLf & _
               "select grupo_Id, Descricao, SUM(Valor) as SomaValor, SUM(valor) * 100 /@strSoma as SomaParticipacao " & vbCrLf & _
               "  from #Temp              " & vbCrLf & _
               "    group by grupo_Id, Descricao " & vbCrLf & _
               "                          " & vbCrLf & _
               "   DROP table #temp       " & vbCrLf

        Return sql
    End Function

    Protected Sub ddlMarcaProd_SelectedIndexChanged(ByVal sender As Object, ByVal e As EventArgs) Handles ddlMarcaProd.SelectedIndexChanged
        Try
            If ddlMarcaProd.SelectedIndex = 0 Then
                ucSelecaoProduto.WhereProduto = ""
            Else
                ucSelecaoProduto.WhereProduto = "marca = " & ddlMarcaProd.SelectedValue
            End If

            ucSelecaoProduto.CarregarNivel(1)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkPerido_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles chkPerido.CheckedChanged
        Try
            pnlPeriodo.Visible = chkPerido.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbQtde_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbQtde.CheckedChanged
        Try
            rbNota.Visible = rbQtde.Checked
            rbEmbalagem.Visible = rbQtde.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rbValor_CheckedChanged(ByVal sender As Object, ByVal e As EventArgs) Handles rbValor.CheckedChanged
        Try
            rbNota.Visible = rbQtde.Checked
            rbEmbalagem.Visible = rbQtde.Checked
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Dim rpt As New ReportDocument()

        Try
            If Funcoes.VerificaPermissao("RelatorioTraitMixProdutos", "RELATORIO") Then
                Dim ds As DataSet
                ds = Banco.ConsultaDataSet(GetSql(), "Consulta")
                ds.Tables(0).TableName = "NotasFiscais"
                ds.Tables(1).TableName = "SomasGerais"

                Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))
                ds.Tables(0).Columns.Add("Imagem", GetType(System.Byte()))
                ds.Tables(0).Columns.Add("NomeEmp", GetType(String))
                ds.Tables(0).Columns.Add("CidadeEstado", GetType(String))

                For Each row As DataRow In ds.Tables(0).Rows
                    row("Imagem") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
                    row("NomeEmp") = HttpContext.Current.Session("ssNomeEmpresa") & " (" & Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("ssEmpresa")) & ")"
                    row("CidadeEstado") = HttpContext.Current.Session("ssCidadeEmpresa") & "/" & HttpContext.Current.Session("ssEstadoEmpresa")
                Next

                rpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/Cr_RelatorioTraitMixProdutos.rpt")
                rpt.Load(rpt.FileName, OpenReportMethod.OpenReportByDefault)
                Dim UrlArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = HttpContext.Current.Server.MapPath(UrlArquivo)

                rpt.SetDataSource(ds)

                Dim param As New Dictionary(Of String, Object)()
                param.Add("DescColuna", descricao)
                param.Add("ParamConsulta", parametros)
                Funcoes.BindParameters(rpt, param)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                rpt.ExportToDisk(ExportFormatType.PortableDocFormat, NomeArquivo)

                If System.IO.File.Exists(NomeArquivo) Then
                    Funcoes.AbrirArquivo(Page, UrlArquivo)
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatorio.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            rpt.Close()
            rpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioTraitMixProdutos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class