Imports NGS.Lib.Negocio

Public Class AtivosXBaixas
    Inherits BasePage

    Private Sql As String

#Region "Evenths"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("AtivosXBaixas", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    CargaUnidadeDeNegocio()
                    CargaGruposDeContas()
                    ddl.Carregar(DdlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "Len(CentroDeCusto_Id) = 5", True)
                    ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.GruposDeAtivos, "", True)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Patrimonio.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            If Not String.IsNullOrWhiteSpace(DdlEmpresaOrigem.SelectedValue) Then
                If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
                    CargaAtivos()
                Else
                    MsgBox(Me.Page, "Informe o grupo do bem.")
                End If
            Else
                MsgBox(Me.Page, "Informe a Empresa a qual pertence o bem.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
            HabilitarCampos()
            txtDataBaixa.Enabled = True
            txtValorBaixa.Enabled = True
            txtMotivoBaixa.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkBaixar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkBaixar.Click
        Try
            If Funcoes.VerificaPermissao("AtivosXBaixas", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList()
                    Dim sql As String = String.Empty
                    Dim calculo As Decimal
                    Dim saldo As Decimal

                    Dim qtdeDiasMes As Integer = DateTime.DaysInMonth(CDate(txtDataBaixa.Text).Year, CDate(txtDataBaixa.Text).Month)

                    If ddlDepreciar.SelectedValue = "S" Then calculo = Math.Round(((((CDec(txtValorOriginal.Text) * ViewState("Indice")) / 100) / 12) / qtdeDiasMes) * CDate(txtDataBaixa.Text).Day, 2)

                    saldo = CDec(txtValorDepreciado.Text) + calculo

                    If saldo > CDec(txtValorOriginal.Text) Then
                        calculo = CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)
                    End If

                    sql = "Update Ativos Set DataDaBaixa   = '" & txtDataBaixa.Text.ToSqlDate() & "'" & vbCrLf & _
                          "                 ,ValorDaBaixa  =  " & Str(CDec(txtValorBaixa.Text.Trim())) & vbCrLf & _
                          "                 ,MotivoDaBaixa = '" & txtMotivoBaixa.Text.Trim().ToUpper() & "'" & vbCrLf & _
                          "                 ,QuemAlterou   = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                          "                 ,QuandoAlterou = getDate()" & vbCrLf & _
                          "                 ,Atualizacao   = '" & txtDataBaixa.Text.ToSqlDate() & "'" & vbCrLf & _
                          "       Where Situacao = 1" & vbCrLf & _
                          "           And Empresa_Id = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                          "           And Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                          "           And Codigo_Id = " & CInt(txtCodigo.Text) & vbCrLf & _
                          "           And Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf
                    SqlArray.Add(sql)

                    If calculo > 0 Then
                        sql = " INSERT INTO AtivosXMovimentos (Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Movimento_Id, Empresa, EndEmpresa, Quotas, Indice, Valor, Processo)" & vbCrLf & _
                        "           VALUES ('" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "', " & "'" & DdlGrupo.SelectedValue & "'," & txtCodigo.Text.Trim & ", " & txtSequencia.Text.Trim & ", " & vbCrLf & _
                        "                   '" & CDate(txtDataBaixa.Text).ToString("yyyy/MM/dd") & "', '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "', " & DdlEmpresaDest.SelectedValue.Split("-")(1) & ", " & vbCrLf & _
                                     Str(calculo) & ", " & "1, " & Str(calculo) & ", 'NORMAL')" & vbCrLf
                        SqlArray.Add(sql)

                        getSqlContabilizacaoDeprec(calculo, SqlArray)
                    End If

                    getSqlContabilzacaoBaixa(SqlArray)

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Baixa efetuada com Sucesso.", eTitulo.Sucess)
                        GerarRelatorioBaixa()
                        Limpar()
                        HabilitarCampos()
                    End If

                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridAtivos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Grupo As String = GridAtivos.SelectedRow.Cells(1).Text()
            Dim Codigo As String = GridAtivos.SelectedRow.Cells(2).Text()
            Dim Sequencia As String = GridAtivos.SelectedRow.Cells(3).Text()

            Sql = " SELECT  Ativos.Empresa_Id, Ativos.Grupo_Id, Ativos.Codigo_Id, Ativos.Sequencia_Id, Ativos.DataAquisicao, Ativos.InicioDeUso, Ativos.Descricao, " & vbCrLf & _
                  "         Ativos.UnidadeDeNegocio, Ativos.Empresa, Ativos.EndEmpresa, Ativos.Conta, Ativos.CentroDeCusto, CentrosDeCustos.Descricao AS NomeDoCusto, " & vbCrLf & _
                  "         Ativos.Historico, Ativos.ValorOriginal, Ativos.Atualizacao, Ativos.Depreciar, Ativos.Seguro, Ativos.DataDaBaixa, " & vbCrLf & _
                  "         isnull(ativos.ValorDaBaixa, 0) as ValorDaBaixa, isnull(Ativos.MotivoDaBaixa, '') as MotivoDaBaixa, at.ValorDepreciado, " & vbCrLf & _
                  "         isnull(atInic.ValorDepreciadoInicial,0) as ValorDepreciadoInicial, GruposDeAtivos.PercentualDepreciacao as Indice" & vbCrLf & _
                  " FROM    Ativos" & vbCrLf & _
                  "     INNER JOIN GruposDeAtivos" & vbCrLf & _
                  "         ON Ativos.Grupo_Id = GruposDeAtivos.Grupo_Id" & vbCrLf & _
                  "     Inner Join(                                                                                         " & vbCrLf & _
                  "                 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado  " & vbCrLf & _
                  "         	        FROM AtivosXMovimentos                                                              " & vbCrLf & _
                  "                     group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                              " & vbCrLf & _
                  "         	) as at                                                                                     " & vbCrLf & _
                  "          On at.Empresa_Id = Ativos.Empresa_Id                                                            " & vbCrLf & _
                  "         And at.Grupo_Id = ativos.Grupo_Id                                                               " & vbCrLf & _
                  "         and at.Codigo_Id = Ativos.Codigo_Id                                                             " & vbCrLf & _
                  "         and at.Sequencia_Id = Ativos.Sequencia_Id                                                       " & vbCrLf & _
                   "     Left Join(                                                                                         " & vbCrLf & _
                  "                 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciadoInicial  " & vbCrLf & _
                  "         	        FROM AtivosXMovimentos  Where Processo = 'Inicial'                                                   " & vbCrLf & _
                  "                     group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                              " & vbCrLf & _
                  "         	) as atInic                                                                                     " & vbCrLf & _
                  "          On atInic.Empresa_Id = Ativos.Empresa_Id                                                            " & vbCrLf & _
                  "         And atInic.Grupo_Id = ativos.Grupo_Id                                                               " & vbCrLf & _
                  "         and atInic.Codigo_Id = Ativos.Codigo_Id                                                             " & vbCrLf & _
                  "         and atInic.Sequencia_Id = Ativos.Sequencia_Id                                                       " & vbCrLf & _
                  "     LEFT JOIN   CentrosDeCustos " & vbCrLf & _
                  "         ON Ativos.CentroDeCusto = CentrosDeCustos.CentroDeCusto_Id" & vbCrLf & _
                  " WHERE   Ativos.Empresa_Id   = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                  "     AND Ativos.Grupo_Id     = '" & Grupo & "'" & vbCrLf & _
                  "     AND Ativos.Codigo_Id    =  " & Codigo & vbCrLf & _
                  "     AND Ativos.Sequencia_Id =  " & Sequencia & vbCrLf

            For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows
                DdlGrupo.SelectedIndex = DdlGrupo.Items.IndexOf(DdlGrupo.Items.FindByValue(row("Grupo_Id")))
                ddlUnidade.SelectedIndex = ddlUnidade.Items.IndexOf(ddlUnidade.Items.FindByValue(row("UnidadeDeNegocio")))

                DdlEmpresaDest.SelectedValue = String.Format("{0}-{1}", row("Empresa"), row("EndEmpresa"))

                DdlGrupoDeContas.SelectedIndex = DdlGrupoDeContas.Items.IndexOf(DdlGrupoDeContas.Items.FindByValue(Left(row("Conta"), 7)))
                CarregarContas()
                DdlContaContabil.SelectedIndex = DdlContaContabil.Items.IndexOf(DdlContaContabil.Items.FindByValue(row("Conta")))
                DdlCentroDeCusto.SelectedIndex = DdlCentroDeCusto.Items.IndexOf(DdlCentroDeCusto.Items.FindByValue(row("CentroDeCusto")))

                txtCodigo.Text = Format(row("Codigo_Id"), "0000")
                txtSequencia.Text = Format(row("Sequencia_Id"), "000")
                txtAquisicao.Text = row("DataAquisicao")
                txtInicioDeUso.Text = row("InicioDeUso")
                txtDescricao.Text = row("Descricao")
                txtHistorico.Text = row("Historico").ToString.ToUpper
                txtValorOriginal.Text = Format(row("ValorOriginal"), "###,###,##0.00")
                ddlDepreciar.SelectedValue = row("Depreciar")
                txtAtualizacao.Text = row("Atualizacao")

                If Not IsDBNull(row("DataDaBaixa")) Then
                    txtDataBaixa.Text = row("DataDaBaixa")
                    txtValorBaixa.Text = row("ValorDaBaixa")
                    txtMotivoBaixa.Text = row("MotivoDaBaixa").ToString.ToUpper()

                    txtDataBaixa.Enabled = False
                    txtValorBaixa.Enabled = False
                    txtMotivoBaixa.Enabled = False
                    lnkBaixar.Parent.Visible = False
                    lnkDesfazerBaixa.Parent.Visible = True
                    lnkReemitirDocBaixa.Parent.Visible = True
                Else
                    txtDataBaixa.Enabled = True
                    txtValorBaixa.Enabled = True
                    txtMotivoBaixa.Enabled = True

                    lnkBaixar.Parent.Visible = True
                    txtDataBaixa.Focus()
                End If

                If row("Seguro") = "TRUE" Then
                    chkSeguro.Checked = True
                Else
                    chkSeguro.Checked = False
                End If

                lnkConsultar.Parent.Visible = False

                txtValorDepreciado.Text = Format(row("ValorDepreciado"), "###,###,##0.00")
                txtValorDepreciadoInicial.Text = Format(row("ValorDepreciadoInicial"), "###,###,##0.00")
                ViewState("Indice") = row("Indice")
                HabilitarCampos(False)
            Next

            TabContainer1.ActiveTabIndex = 0

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidade.SelectedIndex > 0 Then
                CargaEmpresaOrigem()
                DdlEmpresaDest.Items.Clear()
            Else
                DdlEmpresaOrigem.Items.Clear()
                DdlEmpresaDest.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupoDeContas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If DdlGrupoDeContas.SelectedIndex > 0 Then CarregarContas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ATIVOSXBAIXAS")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDesfazerBaixa_Click(sender As Object, e As EventArgs) Handles lnkDesfazerBaixa.Click
        Try
            Dim sqlArray As New ArrayList()

            Dim sql As String = "DELETE AtivosXMovimentos  " & vbCrLf & _
                                " WHERE Empresa_Id   = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'         " & vbCrLf & _
                                "   AND Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                         " & vbCrLf & _
                                "   AND Codigo_Id    = " & CInt(txtCodigo.Text) & "                             " & vbCrLf & _
                                "   AND Sequencia_Id = " & CInt(txtSequencia.Text) & "                          " & vbCrLf & _
                                "   AND Movimento_Id = '" & txtDataBaixa.Text.ToSqlDate() & "'" & vbCrLf
            sqlArray.Add(sql)

            sql &= "UPDATE Ativos                                                              " & vbCrLf & _
                   "   SET motivodabaixa = NULL,                                               " & vbCrLf & _
                   "       valordabaixa  = NULL,                                               " & vbCrLf & _
                   "       datadabaixa   = NULL,                                               " & vbCrLf & _
                   "       atualizacao   = (SELECT MAX(Movimento_Id) FROM AtivosXMovimentos    " & vbCrLf & _
                   "                         WHERE Empresa_Id   = a.Empresa_Id                 " & vbCrLf & _
                   "  						   AND Grupo_Id     = a.Grupo_Id                   " & vbCrLf & _
                   "  						   AND Codigo_Id    = a.Codigo_Id                  " & vbCrLf & _
                   "  						   AND Sequencia_Id = a.Sequencia_Id)              " & vbCrLf & _
                   "  FROM Ativos a                                                            " & vbCrLf & _
                   " WHERE a.Empresa_Id   = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'   " & vbCrLf & _
                   "   AND a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                   " & vbCrLf & _
                   "   AND a.Codigo_Id    = " & CInt(txtCodigo.Text) & "                       " & vbCrLf & _
                   "   AND a.Sequencia_Id = " & CInt(txtSequencia.Text) & "                    " & vbCrLf
            sqlArray.Add(sql)

            sql &= "   DELETE Razao                 " & vbCrLf & _
                   "    WHERE Empresa_Id = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                   "      And EndEmpresa_Id = '" & DdlEmpresaDest.SelectedValue.Split("-")(1) & "'" & vbCrLf & _
                   "      And Lote_id          = 52 " & vbCrLf & _
                   "      AND Movimento_Id     = '" & txtDataBaixa.Text.ToSqlDate() & "'" & vbCrLf & _
                   "      And AtivoImobilizado = '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "'" & vbCrLf
            sqlArray.Add(sql)

            If Banco.GravaBanco(sqlArray) Then
                MsgBox(Me.Page, "Baixa cancelada com Sucesso.", eTitulo.Sucess)
                Limpar()
                HabilitarCampos()
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkReemitirDocBaixa_Click(sender As Object, e As EventArgs) Handles lnkReemitirDocBaixa.Click
        Try
            GerarRelatorioBaixa(True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresaOrigem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresaOrigem.SelectedIndexChanged
        Try
            CargaEmpresaDestino()
            CargaGrupos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub getSqlContabilzacaoBaixa(ByRef SqlArray As ArrayList)
        Dim numero As Integer = getAgrupador(DdlEmpresaDest.SelectedValue.Split("-")(0), DdlEmpresaDest.SelectedValue.Split("-")(1), CDate(txtDataBaixa.Text))

        Sql = "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                        Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                        PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "   Select a.Empresa, a.EndEmpresa,  a.Conta as ContaDoBem,                                                                               " & vbCrLf & _
              "          '' as Cliente, 0 as EndCliente,  a.DataDaBaixa as Movimento, 52 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)    " & vbCrLf & _
              "                                                                                      from Razao                                         " & vbCrLf & _
              "                                                                                     Where Empresa_Id    = a.Empresa                     " & vbCrLf & _
              "                                                                                       and EndEmpresa_Id = a.EndEmpresa                  " & vbCrLf & _
              "                                                                                       and Lote_Id       = 52" & vbCrLf & _
              "                                                                                       and Movimento_id  = a.DataDaBaixa) as Sequencia,  " & vbCrLf & _
              "          a.CentroDeCusto AS Custo, 3 as Indexador, a.DataDaBaixa as DataMoeda, 0 as DebitoOficial, a.ValorOriginal as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "          'BAIXA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: ' + convert(varchar,a.DataDaBaixa ,103) + CASE WHEN a.ValorDaBaixa <> 0 THEN ' - PELO VALOR DE: ' + cast(a.ValorDaBaixa as varchar) ELSE '' END as Historico," & vbCrLf & _
              "          'P' as PrevistoRealizado, 'AtivosXBaixas' as Processo, '" & UsuarioServidor.NomeUsuario & "' as UsuarioInclusao, GETDATE() as UsuarioInclusaoData , a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                                                                                      (Select isnull(MAX(Sequencia_Id) + 1, 1)   " & vbCrLf & _
              "                                                                                         from Razao                                                     " & vbCrLf & _
              "                                                                                        where Empresa_Id    = a.Empresa                                 " & vbCrLf & _
              "                                                                                          and EndEmpresa_Id = a.EndEmpresa                              " & vbCrLf & _
              "                                                                                          and Lote_Id       = 52" & vbCrLf & _
              "                                                                                          and Movimento_id  = a.DataDaBaixa) as Agrupador  " & vbCrLf & _
              "    from Ativos a                                                                                                                                            " & vbCrLf & _
              "   Inner Join AtivosXContas axc                                                                                                                                 " & vbCrLf & _
              "     On a.Empresa     = axc.Empresa_Id                                                                                                                          " & vbCrLf & _
              "     And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                                    " & vbCrLf & _
              "     And a.Conta      = axc.Conta_Id                                                                                                                                 " & vbCrLf & _
              "   Inner Join(                                                                                                                                                  " & vbCrLf & _
              "   		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                                    " & vbCrLf & _
              "   		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                                     " & vbCrLf & _
              "   		 ) as axm                                                                                                                                            " & vbCrLf & _
              "      On axm.Empresa_Id   = a.Empresa_Id                                                                                                                     " & vbCrLf & _
              "     And axm.Grupo_Id     = a.Grupo_Id                                                                                                                       " & vbCrLf & _
              "     and axm.Codigo_Id    = a.Codigo_Id                                                                                                                      " & vbCrLf & _
              "     and axm.Sequencia_Id = a.Sequencia_Id                                                                                                                   " & vbCrLf & _
              "   Where a.Empresa_Id   = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'                                                     " & vbCrLf & _
              "     And a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                                                                    " & vbCrLf & _
              "     And a.Codigo_Id    = " & CInt(txtCodigo.Text) & "                                                                       " & vbCrLf & _
              "     And a.Sequencia_Id = " & CInt(txtSequencia.Text) & "                                                                 " & vbCrLf & _
              "     And (a.ValorOriginal > 0)" & vbCrLf & _
              "                                                                                                                                         " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                        Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                        PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "  Select a.Empresa, a.EndEmpresa,  isnull(axc.DepreciacaoCredito, '') as ContaDepreciacao,                                                                      " & vbCrLf & _
              "         '' as Cliente, 0 as EndCliente,  a.DataDaBaixa as Movimento, 52 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)                   " & vbCrLf & _
              "                                                                                     from Razao                                                                 " & vbCrLf & _
              "                                                                                    where Empresa_Id    = a.Empresa                                             " & vbCrLf & _
              "                                                                                      and EndEmpresa_Id = a.EndEmpresa                                          " & vbCrLf & _
              "                                                                                      and Lote_Id       = 52" & vbCrLf & _
              "                                                                                      and Movimento_id  = a.DataDaBaixa) as Sequencia,  " & vbCrLf & _
              "         a.CentroDeCusto AS Custo, 3 as Indexador, a.DataDaBaixa as DataMoeda, axm.ValorDepreciado as DebitoOficial, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "         'BAIXA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: ' + convert(varchar,a.DataDaBaixa ,103) + CASE WHEN a.ValorDaBaixa <> 0 THEN ' - PELO VALOR DE: ' + cast(a.ValorDaBaixa as varchar) ELSE '' END as Historico, " & vbCrLf & _
              "         'P' as PrevistoRealizado, 'AtivosXBaixas' as Processo, '" & UsuarioServidor.NomeUsuario & "' as UsuarioInclusao, GETDATE() as UsuarioInclusaoData, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado," & vbCrLf & _
              "                                                                                      (Select isnull(MAX(Sequencia_Id), 1)   " & vbCrLf & _
              "                                                                                         from Razao                                                     " & vbCrLf & _
              "                                                                                        where Empresa_Id    = a.Empresa                                 " & vbCrLf & _
              "                                                                                          and EndEmpresa_Id = a.EndEmpresa                              " & vbCrLf & _
              "                                                                                          and Lote_Id       = 52" & vbCrLf & _
              "                                                                                          and Movimento_id  = a.DataDaBaixa) as Agrupador  " & vbCrLf & _
              "   	from Ativos a                                                                                                " & vbCrLf & _
              "   Inner Join AtivosXContas axc                                                                                     " & vbCrLf & _
              "   	  On a.Empresa    = axc.Empresa_Id                                                                              " & vbCrLf & _
              "   	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                        " & vbCrLf & _
              "   	 And a.Conta      = axc.Conta_Id                                                                                  " & vbCrLf & _
              "   Inner Join(                                                                                                      " & vbCrLf & _
              "   	  SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado        " & vbCrLf & _
              "   	    FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                         " & vbCrLf & _
              "   	  ) as axm                                                                                                " & vbCrLf & _
              "      On axm.Empresa_Id   = a.Empresa_Id                                                                         " & vbCrLf & _
              "     And axm.Grupo_Id     = a.Grupo_Id                                                                           " & vbCrLf & _
              "     and axm.Codigo_Id    = a.Codigo_Id                                                                          " & vbCrLf & _
              "     and axm.Sequencia_Id = a.Sequencia_Id                                                                       " & vbCrLf & _
              "   Where a.Empresa_Id      = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'                                                     " & vbCrLf & _
              "     And a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                                                                    " & vbCrLf & _
              "     And a.Codigo_Id    = " & CInt(txtCodigo.Text) & "                                                                       " & vbCrLf & _
              "     And a.Sequencia_Id = " & CInt(txtSequencia.Text) & "                                                                 " & vbCrLf & _
              "     And (axm.ValorDepreciado > 0)" & vbCrLf & _
              "                                                                                                                                         " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                        Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                        PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "  Select a.Empresa, a.EndEmpresa, AXC.ContaDeBaixa AS ContaResidual,                                               " & vbCrLf & _
              "         '' as Cliente, 0 as EndCliente,  a.DataDaBaixa as Movimento, 52 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)   " & vbCrLf & _
              "                                                                                     from Razao                                                     " & vbCrLf & _
              "                                                                                        where Empresa_Id    = a.Empresa                                 " & vbCrLf & _
              "                                                                                          and EndEmpresa_Id = a.EndEmpresa                              " & vbCrLf & _
              "                                                                                          and Lote_Id       = 52" & vbCrLf & _
              "                                                                                          and Movimento_id  = a.DataDaBaixa) as Sequencia,  " & vbCrLf & _
              "         a.CentroDeCusto AS Custo, 3 as Indexador, a.DataDaBaixa as DataMoeda, a.ValorOriginal - axm.ValorDepreciado as ValorResidual, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "         'BAIXA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: ' + convert(varchar,a.DataDaBaixa ,103) + CASE WHEN a.ValorDaBaixa <> 0 THEN ' - PELO VALOR DE: ' + cast(a.ValorDaBaixa as varchar) ELSE '' END as Historico, " & vbCrLf & _
              "         'P' as PrevistoRealizado, 'AtivosXBaixas' as Processo, '" & UsuarioServidor.NomeUsuario & "' as UsuarioInclusao, GETDATE() as UsuarioInclusaoData, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                                                                                      (Select isnull(MAX(Sequencia_Id) - 1, 1)   " & vbCrLf & _
              "                                                                                         from Razao                                                     " & vbCrLf & _
              "                                                                                        where Empresa_Id    = a.Empresa                                 " & vbCrLf & _
              "                                                                                          and EndEmpresa_Id = a.EndEmpresa                              " & vbCrLf & _
              "                                                                                          and Lote_Id       = 52" & vbCrLf & _
              "                                                                                          and Movimento_id  = a.DataDaBaixa) as Agrupador  " & vbCrLf & _
              "   	from Ativos a                                                                                                         " & vbCrLf & _
              "   Inner Join AtivosXContas axc                                                                                              " & vbCrLf & _
              "   	  On a.Empresa    = axc.Empresa_Id                                                                                       " & vbCrLf & _
              "   	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                 " & vbCrLf & _
              "   	 And a.Conta      = axc.Conta_Id                                                                                           " & vbCrLf & _
              "   Inner Join(                                                                                                               " & vbCrLf & _
              "   		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                 " & vbCrLf & _
              "   		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                  " & vbCrLf & _
              "   		 ) as axm                                                                                                         " & vbCrLf & _
              "   	  On axm.Empresa_Id   = a.Empresa_Id                                                                                  " & vbCrLf & _
              "   	 And axm.Grupo_Id     = a.Grupo_Id                                                                                    " & vbCrLf & _
              "   	 and axm.Codigo_Id    = a.Codigo_Id                                                                                   " & vbCrLf & _
              "   	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                " & vbCrLf & _
              "   Where a.Empresa_Id      = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'                                                     " & vbCrLf & _
              "   	 And a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                                                                    " & vbCrLf & _
              "   	 And a.Codigo_Id    = " & CInt(txtCodigo.Text) & "                                                                       " & vbCrLf & _
              "   	 And a.Sequencia_Id = " & CInt(txtSequencia.Text) & "                                                                 " & vbCrLf & _
              "      And ((a.ValorOriginal - axm.ValorDepreciado) > 0)" & vbCrLf
        SqlArray.Add(Sql)
    End Sub

    Private Sub getSqlContabilizacaoDeprec(ByVal valorDepreciado As Decimal, ByRef sqlArray As ArrayList)
        Try

            Dim numero As Integer = getAgrupador(DdlEmpresaDest.SelectedValue.Split("-")(0), DdlEmpresaDest.SelectedValue.Split("-")(1), CDate(txtDataBaixa.Text))

            Sql = "Select a.Empresa, a.EndEmpresa, isnull(axc.DepreciacaoDebito, '') as Debita, isnull(axc.DepreciacaoCredito, '') as Credita, " & vbCrLf & _
                  " 	   a.Codigo_Id, a.Sequencia_Id, a.Grupo_Id,  a.CentroDeCusto AS Custo                                                 " & vbCrLf & _
                  " 	from Ativos a                                                                                                          " & vbCrLf & _
                  " 	Inner Join AtivosXContas axc                                                                                           " & vbCrLf & _
                  " 		 On a.Empresa        = axc.Empresa_Id                                                                              " & vbCrLf & _
                  " 		And a.EndEmpresa     = axc.EndEmpresa_Id                                                                           " & vbCrLf & _
                  " 		And a.Conta          = axc.Conta_Id                                                                                " & vbCrLf & _
                  " 	Where a.Empresa_Id     = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                  " 		And a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                  " 		And a.Codigo_Id    = " & CInt(txtCodigo.Text) & vbCrLf & _
                  " 		And a.Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf

            For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows

                If Not String.IsNullOrWhiteSpace(row("Debita")) Then
                    Sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, " & vbCrLf & _
                          "                   Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico," & vbCrLf & _
                          "                   PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                          "           VALUES ('" & row("Empresa") & "', " & row("EndEmpresa") & ", '" & row("Debita") & "', '', 0" & ", '" & txtDataBaixa.Text.ToSqlDate() & "', 52, (" & getNumerador(DdlEmpresaDest.SelectedValue.Split("-")(0), DdlEmpresaDest.SelectedValue.Split("-")(1), CDate(txtDataBaixa.Text)) & "), " & vbCrLf & _
                          "                    " & row("Custo") & ", 3" & ", '" & txtDataBaixa.Text.ToSqlDate() & "', " & Str(valorDepreciado) & ", 0, 0, 0, " & vbCrLf & _
                          "                   'DEPRECIAÇÃO PARCIAL DO BEM:  " & row("Grupo_Id") & " - " & row("Codigo_Id").ToString("0000") & " - " & row("Sequencia_Id").ToString("000") & " ATE O DIA: " & txtDataBaixa.Text.ToSqlDate() & "', " & vbCrLf & _
                          "                   'P', 'AtivosXBaixas', '" & UsuarioServidor.NomeUsuario & "', getdate(), " & vbCrLf & _
                          "                   '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "', " & numero & ")" & vbCrLf
                    sqlArray.Add(Sql)
                End If

                If Not String.IsNullOrWhiteSpace(row("Credita")) Then
                    Sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, " & vbCrLf & _
                          "                   Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, " & vbCrLf & _
                          "                   PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                          "           VALUES ('" & row("Empresa") & "', " & row("EndEmpresa") & ", '" & row("Credita") & "', '', 0" & ", '" & txtDataBaixa.Text.ToSqlDate() & "', 52, (" & getNumerador(DdlEmpresaDest.SelectedValue.Split("-")(0), DdlEmpresaDest.SelectedValue.Split("-")(1), CDate(txtDataBaixa.Text)) & "), " & vbCrLf & _
                          "                    " & row("Custo") & ", 3" & ", '" & txtDataBaixa.Text.ToSqlDate() & "', 0, " & Str(valorDepreciado) & ", 0, 0, " & vbCrLf & _
                          "                   'DEPRECIAÇÃO PARCIAL DO BEM:  " & row("Grupo_Id") & " - " & row("Codigo_Id").ToString("0000") & " - " & row("Sequencia_Id").ToString("000") & " ATE O DIA: " & txtDataBaixa.Text.ToSqlDate() & "', " & vbCrLf & _
                          "                   'P', 'AtivosXBaixas', '" & UsuarioServidor.NomeUsuario & "', getdate(), " & vbCrLf & _
                          "                   '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "', " & numero & ")" & vbCrLf
                    sqlArray.Add(Sql)
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getNumerador(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal Movimento As DateTime) As String
        Dim sql As String = "Select isnull(MAX(Sequencia_Id) + 1, 1)              " & vbCrLf & _
                            "  from Razao " & vbCrLf & _
                            " where Empresa_Id    = '" & Empresa & "'" & vbCrLf & _
                            "   and EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                            "   and Lote_Id       = 52 " & vbCrLf & _
                            "   and Movimento_id  = '" & Movimento.ToSqlDate() & "'" & vbCrLf

        Return sql
    End Function

    Private Function getAgrupador(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal Movimento As DateTime) As Integer
        Dim sql As String = "Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
                            "  from Razao " & vbCrLf & _
                            " where Empresa_Id    = '" & Empresa & "'" & vbCrLf & _
                            "   and EndEmpresa_Id = " & EndEmpresa & vbCrLf & _
                            "   and Lote_Id       = 52 " & vbCrLf & _
                            "   and Movimento_id  = '" & Movimento.ToSqlDate() & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Tabela")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count Then
            For Each row As DataRow In ds.Tables(0).Rows
                Return row("maxSequencia")
            Next
        End If
        Return 1
    End Function

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
            MsgBox(Me.Page, "Grupo é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtAquisicao.Text) Then
            MsgBox(Me.Page, "Data de Aquisição é obrigatório")
            Return False
        ElseIf Not IsDate(txtAquisicao.Text) Then
            MsgBox(Me.Page, "Data de Aquisição não é uma data válida")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtInicioDeUso.Text) Then
            MsgBox(Me.Page, "Data de Inicio de Uso é obrigatório")
            Return False
        ElseIf Not IsDate(txtInicioDeUso.Text) Then
            MsgBox(Me.Page, "Data de Inicio não é uma data válida")
            Return False
        ElseIf CDate(txtAquisicao.Text) > CDate(txtInicioDeUso.Text) Then
            MsgBox(Me.Page, "Data de Aquisição não pode ser maior que a data de inicio de uso.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresaOrigem.SelectedValue) Then
            MsgBox(Me.Page, "Empresa a qual o bem pertence é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresaDest.SelectedValue) Then
            MsgBox(Me.Page, "Empresa a qual o bem esta localizado é obrigatório")
            Return False
        ElseIf Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) <> Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) Then
            MsgBox(Me.Page, "Empresa onde o bem esta localizado, não é uma empresa filial da empresa a qual o bem pertence.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlGrupoDeContas.SelectedValue) Then
            MsgBox(Me.Page, "Grupo de Contas é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlContaContabil.SelectedValue) Then
            MsgBox(Me.Page, "Conta é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlCentroDeCusto.SelectedValue) Then
            MsgBox(Me.Page, "Centro de Custo é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtHistorico.Text) Then
            MsgBox(Me.Page, "Histórico é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtValorOriginal.Text) OrElse CDbl(txtValorOriginal.Text) = 0 Then
            MsgBox(Me.Page, "Valor original é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtAtualizacao.Text) Then
            MsgBox(Me.Page, "Data de atualização é obrigatório.")
            Return False
        ElseIf Not IsDate(txtAtualizacao.Text) Then
            MsgBox(Me.Page, "Data de atualização não é uma data válida.")
            Return False
        ElseIf CDate(txtAquisicao.Text) > CDate(txtAtualizacao.Text) Then
            MsgBox(Me.Page, "Data de aquisição não pode ser maior que a data de atualização.")
            Return False
        End If
        Return ValidarBaixa()
    End Function

    Private Function ValidarBaixa() As Boolean
        If String.IsNullOrWhiteSpace(txtDataBaixa.Text) Then
            MsgBox(Me.Page, "Informe a data da baixa.")
            Return False
        ElseIf Not IsDate(txtDataBaixa.Text) Then
            MsgBox(Me.Page, "Data da baixa não é uma data válida.")
            Return False
        ElseIf CDate(txtDataBaixa.Text) < CDate(txtInicioDeUso.Text) Then
            MsgBox(Me.Page, "Data da baixa não pode ser menor que data de início de uso.")
            Return False
        ElseIf CDate(txtDataBaixa.Text) < CDate(txtAtualizacao.Text) Then
            MsgBox(Me.Page, "Data da baixa não pode ser menor que data de atualização.")
            Return False
        ElseIf (CDate(txtAtualizacao.Text).Year = CDate(txtDataBaixa.Text).Year) AndAlso (CDate(txtAtualizacao.Text).Month = CDate(txtDataBaixa.Text).Month) Then
            MsgBox(Me.Page, "Impossível dar baixa neste mês, pois o processo ja foi realizado. Elimine o processo deste mês para dar baixa.")
            Return False
        ElseIf ddlDepreciar.SelectedValue = "S" AndAlso (CDate(txtAtualizacao.Text).AddMonths(1).Month <> CDate(txtDataBaixa.Text).Month And ((CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)) > 0)) _
            OrElse (CDate(txtAtualizacao.Text).AddMonths(1).Year <> CDate(txtDataBaixa.Text).Year And ((CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)) > 0)) Then
            MsgBox(Me.Page, "Não é possivel baixar o bem com mês de baixa diferente de: " & MonthName(CDate(txtAtualizacao.Text).AddMonths(1).Month) & " de: " & CDate(txtAtualizacao.Text).AddMonths(1).Year)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtMotivoBaixa.Text) Then
            MsgBox(Me.Page, "Informe o Motivo de Baixa.")
            Return False
        ElseIf Len(txtMotivoBaixa.Text) < 5 Then
            MsgBox(Me.Page, "Informe no mínimo 5 caracteres para o Motivo da Baixa.")
            Return False
        End If

        If String.IsNullOrWhiteSpace(txtValorBaixa.Text) Then
            txtValorBaixa.Text = 0
        End If
        Return True
    End Function

    Private Function HabilitarCampos(Optional ByVal Enable As Boolean = True) As Boolean
        DdlGrupo.Enabled = Enable
        DdlEmpresaDest.Enabled = Enable
        DdlGrupo.Enabled = Enable
    End Function

    Private Sub CargaGruposDeContas()
        ddl.Carregar(DdlGrupoDeContas, CarregarDDL.Tabela.PlanoDeContas, "(len(conta_id) = 7 and left(conta_id, 1) = '1')", True)
    End Sub

    Private Sub CargaUnidadeDeNegocio()
        ddlUnidade.SelectedIndex = ddlUnidade.Items.IndexOf(ddlUnidade.Items.FindByValue(UsuarioServidor.Usuario.AcessoUnidade))
        CarregarEmpresas(DdlEmpresaOrigem, ddlUnidade.SelectedValue, "", True)
        DdlEmpresaOrigem.SelectedIndex = DdlEmpresaOrigem.Items.IndexOf(DdlEmpresaOrigem.Items.FindByValue(UsuarioServidor.Usuario.AcessoEmpresa & "-" & UsuarioServidor.Usuario.AcessoEnderecoEmpresa))
        If DdlEmpresaOrigem.SelectedIndex > 0 Then
            CargaGrupos()
            CargaEmpresaDestino()
        End If
    End Sub
    Private Sub CargaEmpresaOrigem()
        CarregarEmpresas(DdlEmpresaOrigem, ddlUnidade.SelectedValue, "", True)
    End Sub

    Private Sub CargaEmpresaDestino()
        CarregarEmpresas(DdlEmpresaDest, "", DdlEmpresaOrigem.SelectedValue, False)
    End Sub

    Private Sub CarregarEmpresas(ByVal ddl As DropDownList, ByVal unidade As String, ByVal Empresa As String, ByVal Matriz As Boolean)
        Dim sql As String = " SELECT DISTINCT c.Cliente_Id as Codigo, c.Endereco_Id, c.Reduzido, c.Nome, c.Cidade, c.Estado  " & vbCrLf & _
                               " 	FROM GruposXEmpresas gxe                                                                    " & vbCrLf & _
                               "    INNER JOIN Clientes c                                                                       " & vbCrLf & _
                               "       ON gxe.Cliente_Id     = c.Cliente_Id                                                     " & vbCrLf & _
                               "      AND gxe.EndCliente_Id = c.Endereco_Id                                                     " & vbCrLf & _
                               "    Inner Join ClientesXEmpresas cxe                                                            " & vbCrLf & _
                               " 	  on cxe.Empresa_Id = c.Cliente_Id                                                          " & vbCrLf & _
                               " 	 and cxe.EndEmpresa_Id = c.Endereco_Id                                                      " & vbCrLf & _
                               "    Where 1 = 1                                                                     " & vbCrLf
        If Matriz Then
            sql &= "And cxe.Matriz = 'S'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Empresa) Then
            sql &= "And Left(cxe.Empresa_Id, 8) = '" & Left(Empresa, 8) & "'"
        End If

        If Not String.IsNullOrWhiteSpace(unidade) Then
            sql &= "And gxe.Empresa_Id = '" & unidade & "'" & vbCrLf
        End If



        Dim ds = Banco.ConsultaDataSet(sql, "Empresas")

        ddl.Items.Clear()
        ddl.DataTextField = "Descricao"
        ddl.DataValueField = "Codigo"

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                ddl.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".") & " - " & Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".") & " " & Dr("Estado") & " " & Funcoes.AlinharEsquerda(Funcoes.FormatarCpfCnpj(Dr("Codigo")), 18, ".") & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido"), Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))))
            Next
        End If

        Funcoes.InserirLinhaEmBranco(ddl, 0)
    End Sub

    Private Sub CarregarContas()
        ddl.Carregar(DdlContaContabil, CarregarDDL.Tabela.PlanoDeContas, "conta_id like '" & DdlGrupoDeContas.SelectedValue & "%' and len(Conta_Id) = 9", True)
    End Sub

    Private Sub CargaGrupos()
        ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.GruposDeAtivos, DdlEmpresaOrigem.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = String.Empty
        txtSequencia.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtHistorico.Text = String.Empty
        txtValorOriginal.Text = String.Empty
        txtValorDepreciado.Text = String.Empty
        ddlDepreciar.SelectedValue = "S"
        chkSeguro.Checked = False

        DdlEmpresaDest.Items.Clear()

        VerificaUnidade()

        DdlGrupoDeContas.SelectedIndex = 0

        If DdlGrupo.Items.Count > 0 Then DdlGrupo.SelectedIndex = 0

        DdlContaContabil.Items.Clear()
        DdlCentroDeCusto.SelectedIndex = 0

        txtValorDepreciadoInicial.Text = String.Empty

        txtAquisicao.Text = String.Empty
        txtAtualizacao.Text = String.Empty
        txtInicioDeUso.Text = String.Empty

        txtMotivoBaixa.Text = String.Empty
        txtValorBaixa.Text = String.Empty
        txtDataBaixa.Text = String.Empty

        lnkConsultar.Parent.Visible = True
        lnkDesfazerBaixa.Parent.Visible = False
        lnkReemitirDocBaixa.Parent.Visible = False
        lnkBaixar.Parent.Visible = False
        TabContainer1.ActiveTabIndex = 0

        GridAtivos.DataSource = Nothing
        GridAtivos.DataBind()

        If DdlEmpresaOrigem.SelectedIndex > 0 Then CargaEmpresaDestino()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresaOrigem.Enabled = False
        End If
    End Sub

    Private Sub VerificaUnidade()
        Dim Sql As String = ""
        Sql = "SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf &
            "       isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf &
            "       isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf &
            "  from Usuarios" & vbCrLf &
            " where Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            ddl.Carregar(DdlEmpresaOrigem, CarregarDDL.Tabela.Empresas, Dr("AcessoUnidade"), True)
            DdlEmpresaOrigem.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next

        ddlUnidade.Enabled = False
        DdlEmpresaOrigem.Enabled = False
    End Sub

    Private Sub CargaAtivos()
        Sql = " SELECT Grupo_Id as Grupo, Codigo_Id as Codigo, Sequencia_Id as Sequencia, Descricao, DataAquisicao as Aquisicao, ValorOriginal as Valor,  " & vbCrLf & _
              " case " & vbCrLf & _
              "    when DataDaBaixa is null " & vbCrLf & _
              "       then 'N' " & vbCrLf & _
              "       else 'S' " & vbCrLf & _
              "    end Baixado " & vbCrLf & _
              "" & vbCrLf & _
              "   FROM Ativos" & vbCrLf & _
              "  WHERE Situacao = 1 " & vbCrLf & _
              "    and Empresa_Id = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
              "    And Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlEmpresaDest.SelectedValue) Then
            Sql &= " And Empresa = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf
        End If

        Sql &= " Order by Grupo_Id, Codigo_Id, Sequencia_Id"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")
        GridAtivos.DataSource = ds
        GridAtivos.DataBind()

        If ds.Tables(0).Rows.Count > 0 Then
            TabContainer1.ActiveTabIndex = 1

            Dim i As Integer = 0

            While i < GridAtivos.Rows.Count
                If ds.Tables(0).Rows(i).Item("Baixado").ToString = "S" Then GridAtivos.Rows(i).ForeColor = Drawing.Color.Red

                i += 1
            End While
        End If

    End Sub

    Private Sub GerarRelatorioBaixa(Optional ByVal reemissao As Boolean = False)
        Dim sql As String = "   select a.codigo_Id, a.Grupo_Id, ga.Descricao as DescricaoGrupo, a.Sequencia_Id, a.Descricao, a.Historico,  " & vbCrLf & _
                            "          a.DataAquisicao, a.InicioDeUso, a.ValorOriginal, SUM(isnull(am.Valor, 0)) as valorDepreciado,       " & vbCrLf & _
                            "          a.ValorOriginal - SUM(isnull(am.Valor, 0)) as ValorResidual, ga.PercentualDepreciacao,              " & vbCrLf & _
                            "          a.ValorDaBaixa, a.DataDaBaixa, a.MotivoDaBaixa                                                      " & vbCrLf & _
                            "     from Ativos a                                                                                            " & vbCrLf & _
                            "    Inner Join AtivosXMovimentos am                                                                           " & vbCrLf & _
                            "   	on am.Empresa_Id   = a.Empresa_Id                                                                      " & vbCrLf & _
                            "      and am.Codigo_Id    = a.Codigo_Id                                                                       " & vbCrLf & _
                            "      and am.Grupo_Id     = a.Grupo_Id                                                                        " & vbCrLf & _
                            "      and am.Sequencia_Id = a.Sequencia_Id                                                                    " & vbCrLf & _
                            "    Inner Join GruposDeAtivos ga                                                                              " & vbCrLf & _
                            "   	on ga.Grupo_Id = a.Grupo_Id		                                                                       " & vbCrLf & _
                            "   Where a.Empresa_Id = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'                             " & vbCrLf & _
                            "     and a.Grupo_Id = '" & DdlGrupo.SelectedValue & "'                                                        " & vbCrLf & _
                            "     and a.Codigo_Id = " & CInt(txtCodigo.Text) & vbCrLf & _
                            "     And a.Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf & _
                            "   group by a.codigo_Id, a.Grupo_Id, ga.Descricao, a.Sequencia_Id, a.Descricao, a.Historico,                  " & vbCrLf & _
                            "            a.DataAquisicao, a.InicioDeUso, a.ValorOriginal, ga.PercentualDepreciacao,                        " & vbCrLf & _
                            "            a.ValorDaBaixa, a.DataDaBaixa, a.MotivoDaBaixa                                                    " & vbCrLf & _
                            "                                                                                                              " & vbCrLf & _
                            "   select Conta_Id, Historico, DebitoOficial as Debito, CreditoOficial as Credito, Sequencia_Id               " & vbCrLf & _
                            "     from Razao                                                                                               " & vbCrLf & _
                            "    WHERE Empresa_Id       = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
                            "      And EndEmpresa_Id    = '" & DdlEmpresaDest.SelectedValue.Split("-")(1) & "'" & vbCrLf & _
                            "      And Lote_id          = 52 " & vbCrLf & _
                            "      AND Movimento_Id     = '" & txtDataBaixa.Text.ToSqlDate() & "'" & vbCrLf & _
                            "      And AtivoImobilizado = '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "'" & vbCrLf & _
                            "    order by UsuarioInclusaoData " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Ativos")

        Dim param As New Dictionary(Of String, Object)
        param.Add("VerificaReemissao", IIf(reemissao, "Reemissão de documento.", ""))

        If ds IsNot Nothing AndAlso ds.Tables.Count = 2 Then ds.Tables(1).TableName = "Razao"

        Funcoes.BindReport(Me.Page, ds, "Cr_AtivosXBaixas", eExportType.PDF, param)
    End Sub

#End Region

End Class