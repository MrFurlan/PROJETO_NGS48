Imports NGS.Lib.Negocio

Public Class AtivosXTransferencia
    Inherits BasePage

    Private Sql As String

#Region "Evenths"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Ativos", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    CargaUnidadeDeNegocio()
                    CargaCentroDeCusto()
                    CargaGruposDeContas()
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
            HabilitarCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkTransferir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkTransferir.Click
        Try
            If Funcoes.VerificaPermissao("Ativos", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList()
                    Dim sql As String = String.Empty
                    Dim CalculoDepreciacao As Decimal
                    Dim saldo As Decimal
                    Dim DepreciacaoMensal As Decimal
                    Dim DepreciacaoParcial As Decimal

                    'Dim qtdeDiasMes As Integer = DateTime.DaysInMonth(CDate(txtDataTransferencia.Text).Year, CDate(txtDataTransferencia.Text).Month)
                    Dim qtdeDiasMes As Integer = 30 - CDate(txtDataTransferencia.Text).Day



                    'If CDate(row("InicioDeUso")).Month = CDate(txtDataProxAtualizacao.Text).Month And CDate(row("InicioDeUso")).Year = CDate(txtDataProxAtualizacao.Text).Year Then 'Mes parcial
                    '    dias = (30 - diaInicioDeUsoBase) + 1
                    '    calculo = Math.Round((((row("Valor") * row("Indice")) / 100) / 360) * dias, 2)
                    'Else 'Mes completo
                    '    calculo = Math.Round((((row("Valor") * row("Indice")) / 100) / 360) * 30, 2)
                    'End If

                    '****************************************************************************************************************************

                    ''saldo = row("Depreciado") + calculo


                    'calculo = Math.Round(((((CDec(txtValorOriginal.Text) * ViewState("Indice")) / 100) / 12) / qtdeDiasMes) * CDate(txtDataTransferencia.Text).Day, 2)

                    DepreciacaoMensal = Math.Round(CDec(txtValorOriginal.Text) * ViewState("Indice") / 100 / 360 * 30, 2, MidpointRounding.AwayFromZero)
                    DepreciacaoParcial = Math.Round(CDec(txtValorOriginal.Text) * ViewState("Indice") / 100 / 360 * (30 - qtdeDiasMes), 2, MidpointRounding.AwayFromZero)

                    CalculoDepreciacao = DepreciacaoMensal - DepreciacaoParcial  'Math.Round(CDec(txtValorOriginal.Text) * ViewState("Indice") / 100 / 360 * qtdeDiasMes, 2, MidpointRounding.AwayFromZero)


                    saldo = CDec(txtValorDepreciado.Text) + CalculoDepreciacao

                    If saldo > CDec(txtValorOriginal.Text) Then
                        CalculoDepreciacao = CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)
                    End If

                    If CalculoDepreciacao > 0 Then
                        sql = " INSERT INTO AtivosXMovimentos (Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Movimento_Id, Empresa, EndEmpresa, Quotas, Indice, Valor, Processo)" & vbCrLf & _
                        "           VALUES ('" & Left(ViewState("EmpresaOrigem").Split("-")(0), 8) & "', " & "'" & DdlGrupo.SelectedValue & "'," & txtCodigo.Text.Trim & ", " & txtSequencia.Text.Trim & ", " & vbCrLf & _
                        "                   '" & txtDataTransferencia.Text.ToSqlDate() & "', '" & ViewState("EmpresaOrigem").Split("-")(0) & "', " & ViewState("EmpresaOrigem").Split("-")(1) & ", " & vbCrLf & _
                                     Str(CalculoDepreciacao) & ", " & "1, " & Str(CalculoDepreciacao) & ", 'TRANSFERENCIA')" & vbCrLf
                        SqlArray.Add(sql)

                        getSqlContabilizacao(CalculoDepreciacao, SqlArray)
                    End If

                    getSqlContabilizacaoTransferencia(SqlArray)

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Transferencia efetuada com Sucesso.", eTitulo.Sucess)
                        GerarRelatorioTransferencia(CDate(txtDataTransferencia.Text))
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

    Private Sub habilitarBotoes(ByVal habilitar As Boolean)
        lnkTransferir.Parent.Visible = habilitar
        lnkDesfazerTrasferencia.Parent.Visible = habilitar
        lnkReemitirDocTrans.Parent.Visible = habilitar
        lnkConsultar.Parent.Visible = Not habilitar
    End Sub

    Protected Sub GridAtivos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Grupo As String = GridAtivos.SelectedRow.Cells(1).Text()
            Dim Codigo As String = GridAtivos.SelectedRow.Cells(2).Text()
            Dim Sequencia As String = GridAtivos.SelectedRow.Cells(3).Text()

            Sql = " SELECT A.Empresa_Id, A.Grupo_Id, A.Codigo_Id, A.Sequencia_Id, A.DataAquisicao, A.InicioDeUso, A.Descricao, " & vbCrLf & _
                  "        A.UnidadeDeNegocio, A.Empresa, A.EndEmpresa, A.Conta, A.CentroDeCusto, Cc.Descricao AS NomeDoCusto, " & vbCrLf & _
                  "        A.Historico, A.ValorOriginal, A.Atualizacao, A.Depreciar, A.Seguro, A.DataDaBaixa, " & vbCrLf & _
                  "        ISNULL(A.ValorDaBaixa, 0) as ValorDaBaixa, ISNULL(A.MotivoDaBaixa, '') as MotivoDaBaixa, at.ValorDepreciado, " & vbCrLf & _
                  "        ISNULL(atInic.ValorDepreciadoInicial,0) as ValorDepreciadoInicial, Ga.PercentualDepreciacao as Indice" & vbCrLf & _
                  "   FROM Ativos A" & vbCrLf & _
                  "  INNER JOIN GruposDeAtivos Ga" & vbCrLf & _
                  "     ON A.Grupo_Id = Ga.Grupo_Id" & vbCrLf & _
                  "  INNER JOIN (" & vbCrLf & _
                  "               SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado  " & vbCrLf & _
                  "         	    FROM AtivosXMovimentos " & vbCrLf & _
                  "                GROUP BY Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id" & vbCrLf & _
                  "         	) as at" & vbCrLf & _
                  "     ON at.Empresa_Id = A.Empresa_Id" & vbCrLf & _
                  "    AND at.Grupo_Id = A.Grupo_Id" & vbCrLf & _
                  "    AND at.Codigo_Id = A.Codigo_Id" & vbCrLf & _
                  "    AND at.Sequencia_Id = A.Sequencia_Id" & vbCrLf & _
                   "  LEFT JOIN( " & vbCrLf & _
                  "              SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciadoInicial  " & vbCrLf & _
                  "         	   FROM AtivosXMovimentos  " & vbCrLf & _
                  "               WHERE Processo = 'Inicial'" & vbCrLf & _
                  "               GROUP BY Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id" & vbCrLf & _
                  "         	) as atInic" & vbCrLf & _
                  "     ON atInic.Empresa_Id   = A.Empresa_Id" & vbCrLf & _
                  "    AND atInic.Grupo_Id     = A.Grupo_Id" & vbCrLf & _
                  "    AND atInic.Codigo_Id    = A.Codigo_Id" & vbCrLf & _
                  "    AND atInic.Sequencia_Id = A.Sequencia_Id" & vbCrLf & _
                  "   LEFT JOIN CentrosDeCustos Cc " & vbCrLf & _
                  "     ON A.CentroDeCusto = Cc.CentroDeCusto_Id" & vbCrLf & _
                  "  WHERE A.Empresa_Id   = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                  "    AND A.Grupo_Id     = '" & Grupo & "'" & vbCrLf & _
                  "    AND A.Codigo_Id    =  " & Codigo & vbCrLf & _
                  "    AND A.Sequencia_Id =  " & Sequencia & vbCrLf



            'Dim row As DataRow = Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows(0)

            'With row
            '    DdlGrupo.SelectedIndex = DdlGrupo.Items.IndexOf(DdlGrupo.Items.FindByValue(("Grupo_Id")))
            '    ddlUnidade.SelectedIndex = ddlUnidade.Items.IndexOf(ddlUnidade.Items.FindByValue(("UnidadeDeNegocio")))

            '    DdlGrupoDeContas.SelectedIndex = DdlGrupoDeContas.Items.IndexOf(DdlGrupoDeContas.Items.FindByValue(Left(("Conta"), 7)))
            '    CarregarContas()
            '    DdlContaContabil.SelectedIndex = DdlContaContabil.Items.IndexOf(DdlContaContabil.Items.FindByValue(("Conta")))
            '    DdlCentroDeCusto.SelectedIndex = DdlCentroDeCusto.Items.IndexOf(DdlCentroDeCusto.Items.FindByValue(("CentroDeCusto")))

            '    txtCodigo.Text = Format(("Codigo_Id"), "0000")
            '    txtSequencia.Text = Format(("Sequencia_Id"), "000")
            '    txtAquisicao.Text = ("DataAquisicao")
            '    txtInicioDeUso.Text = Format(("InicioDeUso"), "yyyy/MM/dd")
            '    txtDescricao.Text = ("Descricao")
            '    txtMotivoTransferencia.Text = ("MotivoDaBaixa").ToString().ToUpper()
            '    txtValorOriginal.Text = ("ValorOriginal").ToString()
            '    ddlDepreciar.SelectedValue = ("Depreciar")
            '    txtAtualizacao.Text = ("Atualizacao")
            '    txtHistorico.Text = ("Historico").ToString().ToUpper()
            '    'DdlEmpresaDest.SelectedValue = String.Format("{0}-{1}", ("Empresa"), ("EndEmpresa"))
            '    'ViewState("EmpresaOrigem") = String.Format("{0}-{1}", ("Empresa"), ("EndEmpresa"))

            '    txtDataTransferencia.Focus()

            '    If ("Seguro") = "TRUE" Then
            '        chkSeguro.Checked = True
            '    Else
            '        chkSeguro.Checked = False
            '    End If

            '    txtValorDepreciado.Text = ("ValorDepreciado").ToString()
            '    txtValorDepreciadoInicial.Text = ("ValorDepreciadoInicial").ToString()
            '    ViewState("Indice") = ("Indice")
            '    HabilitarCampos(False)
            '    habilitarBotoes(True)
            'End With

            For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows
                DdlGrupo.SelectedIndex = DdlGrupo.Items.IndexOf(DdlGrupo.Items.FindByValue(row("Grupo_Id")))
                ddlUnidade.SelectedIndex = ddlUnidade.Items.IndexOf(ddlUnidade.Items.FindByValue(row("UnidadeDeNegocio")))

                DdlGrupoDeContas.SelectedIndex = DdlGrupoDeContas.Items.IndexOf(DdlGrupoDeContas.Items.FindByValue(Left(row("Conta"), 7)))
                CarregarContas()
                DdlContaContabil.SelectedIndex = DdlContaContabil.Items.IndexOf(DdlContaContabil.Items.FindByValue(row("Conta")))
                DdlCentroDeCusto.SelectedIndex = DdlCentroDeCusto.Items.IndexOf(DdlCentroDeCusto.Items.FindByValue(row("CentroDeCusto")))

                txtCodigo.Text = Format(row("Codigo_Id"), "0000")
                txtSequencia.Text = Format(row("Sequencia_Id"), "000")
                txtAquisicao.Text = row("DataAquisicao")
                txtInicioDeUso.Text = row("InicioDeUso")
                txtDescricao.Text = row("Descricao")
                txtMotivoTransferencia.Text = row("MotivoDaBaixa").ToString().ToUpper()
                txtValorOriginal.Text = row("ValorOriginal").ToString()
                ddlDepreciar.SelectedValue = row("Depreciar")
                txtAtualizacao.Text = row("Atualizacao")
                txtHistorico.Text = row("Historico").ToString().ToUpper()
                DdlEmpresaDest.SelectedValue = String.Format("{0}-{1}", row("Empresa"), row("EndEmpresa"))
                ViewState("EmpresaOrigem") = String.Format("{0}-{1}", row("Empresa"), row("EndEmpresa"))

                txtDataTransferencia.Focus()

                If row("Seguro") = "TRUE" Then
                    chkSeguro.Checked = True
                Else
                    chkSeguro.Checked = False
                End If

                txtValorDepreciado.Text = row("ValorDepreciado").ToString()
                txtValorDepreciadoInicial.Text = row("ValorDepreciadoInicial").ToString()
                ViewState("Indice") = row("Indice")
                HabilitarCampos(False)
                habilitarBotoes(True)
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

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkConsultar.Click
        Try
            If Not String.IsNullOrWhiteSpace(DdlEmpresaOrigem.SelectedValue) Then
                If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
                    CargaAtivos()
                Else
                    MsgBox(Me.Page, "Informe o Grupo.")
                End If
            Else
                MsgBox(Me.Page, "Informe a Empresa a qual pertence o bem.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ATIVOSXTRANSFERENCIA")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkDesfazerTrasferencia_Click(sender As Object, e As EventArgs) Handles lnkDesfazerTrasferencia.Click
        Try
            If Not validaDesfazTransferencia() Then
                MsgBox(Me.Page, "Impossível desfazer transferência. Verifique se ja foram realizados processos para a empresa atual.")
            Else
                Dim sqlArray As New ArrayList()

                Dim sql As String = "DELETE AtivosXMovimentos  " & vbCrLf & _
                                    " WHERE Empresa_Id   = '" & Left(ViewState("EmpresaOrigem").Split("-")(0), 8) & "'" & vbCrLf & _
                                    "   AND Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                                    "   AND Codigo_Id    =  " & CInt(txtCodigo.Text) & vbCrLf & _
                                    "   AND Sequencia_Id =  " & CInt(txtSequencia.Text) & vbCrLf & _
                                    "   AND Movimento_Id > '" & CDate(txtAtualizacao.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                                    "   AND Processo =  'TRANSFERENCIA'" & vbCrLf
                sqlArray.Add(sql)

                sql = "Update Ativos" & vbCrLf & _
                      "   Set Empresa = (SELECT top 1 Empresa " & vbCrLf & _
                      "                    FROM AtivosXMovimentos " & vbCrLf & _
                      "                   WHERE Grupo_Id     = '" & DdlGrupo.SelectedValue & "' " & vbCrLf & _
                      "                     AND Codigo_Id    =  " & CInt(txtCodigo.Text) & vbCrLf & _
                      "                     AND Sequencia_Id =  " & CInt(txtSequencia.Text) & vbCrLf & _
                      "                   GROUP BY Empresa, EndEmpresa, Movimento_Id " & vbCrLf & _
                      "                   ORDER BY Movimento_Id desc),                " & vbCrLf & _
                      "       EndEmpresa = (SELECT top 1 EndEmpresa                       " & vbCrLf & _
                      "                       FROM AtivosXMovimentos                    " & vbCrLf & _
                      "                      WHERE Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                      "                        AND Codigo_Id    =  " & CInt(txtCodigo.Text) & vbCrLf & _
                      "                        AND Sequencia_Id =  " & CInt(txtSequencia.Text) & vbCrLf & _
                      "                      GROUP BY Empresa, EndEmpresa, Movimento_Id " & vbCrLf & _
                      "                      ORDER BY Movimento_Id desc),                " & vbCrLf & _
                      "       Motivodabaixa = NULL,                                               " & vbCrLf & _
                      "       InicioDeUso   = '" & CDate(txtAquisicao.Text).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "       atualizacao   = (SELECT MAX(Movimento_Id) FROM AtivosXMovimentos    " & vbCrLf & _
                      "                         Where Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                      "                           AND Codigo_Id    =  " & CInt(txtCodigo.Text) & vbCrLf & _
                      "                           AND Sequencia_Id =  " & CInt(txtSequencia.Text) & ")" & vbCrLf & _
                      " WHERE Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                      "   AND Codigo_Id    =  " & CInt(txtCodigo.Text) & vbCrLf & _
                      "   AND Sequencia_Id =  " & CInt(txtSequencia.Text) & vbCrLf

                sqlArray.Add(sql)

                sql &= "   DELETE Razao                                                                                               " & vbCrLf & _
                       "    WHERE lote_id          = 53 " & vbCrLf & _
                       "      AND Movimento_Id     = '" & CDate(txtAtualizacao.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                       "      AND AtivoImobilizado = '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "'"
                sqlArray.Add(sql)

                If Banco.GravaBanco(sqlArray) Then
                    MsgBox(Me.Page, "transferencia cancelada com Sucesso.", eTitulo.Sucess)
                    Limpar()
                    HabilitarCampos()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkReemitirDocTrans_Click(sender As Object, e As EventArgs) Handles lnkReemitirDocTrans.Click
        Try
            GerarRelatorioTransferencia(CDate(txtAtualizacao.Text), True)
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

    Private Sub getSqlContabilizacaoTransferencia(ByRef SqlArray As ArrayList)
        Sql = " declare @Empresa varchar(18), @grupo varchar(3), @Codigo int, @Sequencia int;" & vbCrLf & _
              "  select @Empresa = '" & ViewState("EmpresaOrigem").Split("-")(0) & "', @grupo = '" & DdlGrupo.SelectedValue & "', @Codigo = " & CInt(txtCodigo.Text) & ", @Sequencia = " & CInt(txtSequencia.Text) & ";" & vbCrLf & _
              "                                                                                                                                                                                                        " & vbCrLf & _
              "  INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                     Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                     PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)        " & vbCrLf & _
              "  Select a.Empresa, a.EndEmpresa, a.Conta, '' as Cliente, 0 as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, " & vbCrLf & _
              "         53 as Lote, ( Select isnull(MAX(Sequencia_Id) + 1, 1)           " & vbCrLf & _
              "                         from Razao                                        " & vbCrLf & _
              "                        Where Empresa_Id    = a.Empresa                    " & vbCrLf & _
              "                          and EndEmpresa_Id = a.EndEmpresa                 " & vbCrLf & _
              "                          and Lote_Id       = 53) as Sequencia, " & vbCrLf & _
              "         a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, 0 as DebitoOficial, a.ValorOriginal as CreditoOficial, " & vbCrLf & _
              "         0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "         'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "         'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, " & vbCrLf & _
              "         a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "         (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "            from Razao " & vbCrLf & _
              "           where Empresa_Id    = a.Empresa" & vbCrLf & _
              "             and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "             and Lote_Id         = 53" & vbCrLf & _
              "             and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento  " & vbCrLf & _
              "    from Ativos a                                                                                                               " & vbCrLf & _
              "   Inner Join AtivosXContas axc                                                                                                 " & vbCrLf & _
              "      On a.Empresa = axc.Empresa_Id                                                                                             " & vbCrLf & _
              "     And a.EndEmpresa = axc.EndEmpresa_Id                                                                                       " & vbCrLf & _
              "     And a.Conta = axc.Conta_Id                                                                                                 " & vbCrLf & _
              "   Inner Join( SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                 " & vbCrLf & _
              "          	    FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                  " & vbCrLf & _
              "    ) as axm                                                                                                                    " & vbCrLf & _
              "      On axm.Empresa_Id   = a.Empresa_Id                                                                                        " & vbCrLf & _
              "     And axm.Grupo_Id     = a.Grupo_Id                                                                                          " & vbCrLf & _
              "     and axm.Codigo_Id    = a.Codigo_Id                                                                                         " & vbCrLf & _
              "     and axm.Sequencia_Id = a.Sequencia_Id                                                                                      " & vbCrLf & _
              "   Where a.Empresa_Id   =   Left(@Empresa, 8)                                                                                   " & vbCrLf & _
              "     And a.Grupo_Id     = @grupo                                                                                                " & vbCrLf & _
              "     And a.Codigo_Id    = @Codigo                                                                                               " & vbCrLf & _
              "     And a.Sequencia_Id = @Sequencia                                                                                            " & vbCrLf & _
              "                                                                                                                                " & vbCrLf & _
              "  INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,       " & vbCrLf & _
              "                     Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,          " & vbCrLf & _
              "                     PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
              "  Select a.Empresa, a.EndEmpresa, AXC.ContaDeTransferencia,                                                                     " & vbCrLf & _
              "         '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "' as Cliente, " & DdlEmpresaDest.SelectedValue.Split("-")(1) & " as EndCliente," & vbCrLf & _
              "         '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)            " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, a.ValorOriginal as DebitoOficial, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                        'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                        'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                                " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                  " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                              " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                        " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                  " & vbCrLf & _
              "       Inner Join(                                                                                                                                   " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                        " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                         " & vbCrLf & _
              "       		 ) as axm                                                                                                                                " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                         " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                           " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                       " & vbCrLf & _
              "      Where a.Empresa_Id     = Left(@Empresa, 8)                                                                                                                " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                 " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                                " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                             " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                      PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, isnull(axc.DepreciacaoCredito, '') as ContaDepreciacao,                                                           " & vbCrLf & _
              "           '' as Cliente, 0 as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)    " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, axm.ValorDepreciado as DebitoOficial, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                                'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                                'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                                " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                  " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                              " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                        " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                  " & vbCrLf & _
              "       Inner Join(                                                                                                                                   " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                        " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                         " & vbCrLf & _
              "       		 ) as axm                                                                                                                                " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                         " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                           " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                       " & vbCrLf & _
              "       Where a.Empresa_Id    = Left(@Empresa, 8)                                                                                                               " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                 " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                                " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                             " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                      PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, isnull(axc.ContaDeTransferencia, '') as ContaDepreciacao,                                                          " & vbCrLf & _
              "           '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "' as Cliente, " & DdlEmpresaDest.SelectedValue.Split("-")(1) & " as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)    " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, 0 as DebitoOficial, axm.ValorDepreciado as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                                        'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                                        'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                               " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                 " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                             " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                       " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                 " & vbCrLf & _
              "       Inner Join(                                                                                                                                  " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                       " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                        " & vbCrLf & _
              "       		 ) as axm                                                                                                                               " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                        " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                         " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                      " & vbCrLf & _
              "       Where a.Empresa_Id    = Left(@Empresa, 8)                                                                                                              " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                               " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                            " & vbCrLf
        SqlArray.Add(Sql)

        Sql = "Update Ativos Set MotivoDaBaixa = '" & txtMotivoTransferencia.Text.Trim.ToUpper() & "'" & vbCrLf & _
              "                 ,QuemAlterou   = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
              "                 ,QuandoAlterou = getDate()" & vbCrLf & _
              "                 ,InicioDeUso   = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "                 ,Atualizacao   = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "                 ,Empresa       = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
              "                 ,EndEmpresa    = " & DdlEmpresaDest.SelectedValue.Split("-")(1) & vbCrLf & _
              "       Where Empresa_Id   = '" & Left(ViewState("EmpresaOrigem").Split("-")(0), 8) & "'" & vbCrLf & _
              "         And Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
              "         And Codigo_Id    = " & CInt(txtCodigo.Text) & vbCrLf & _
              "         And Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf
        SqlArray.Add(Sql)

        Sql = "declare  @Empresa varchar(18), @grupo varchar(3), @Codigo int, @Sequencia int;" & vbCrLf & _
              "select   @Empresa = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "', @grupo = '" & DdlGrupo.SelectedValue & "', @Codigo = " & CInt(txtCodigo.Text) & ", @Sequencia = " & CInt(txtSequencia.Text) & ";" & vbCrLf & _
              "" & vbLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                        Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                        PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, a.Conta,                                                                               " & vbCrLf & _
              "           '' as Cliente, 0 as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)           " & vbCrLf & _
              "                                                                                          from Razao                                        " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                    " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                 " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia, " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, a.ValorOriginal as DebitoOficial, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "'as Historico," & vbCrLf & _
              "                'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                                " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                  " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                              " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                        " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                  " & vbCrLf & _
              "       Inner Join(                                                                                                                                   " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                        " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                         " & vbCrLf & _
              "       		 ) as axm                                                                                                                                " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                         " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                           " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                       " & vbCrLf & _
              "      Where a.Empresa_Id   =   Left(@Empresa, 8)                                                                                                                " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                 " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                                " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                             " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                      PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, AXC.ContaDeTransferencia,                                                                                        " & vbCrLf & _
              "           '" & ViewState("EmpresaOrigem").Split("-")(0) & "' as Cliente, " & ViewState("EmpresaOrigem").Split("-")(1) & " as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)            " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, 0 as DebitoOficial, a.ValorOriginal as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                        'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                        'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                                " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                  " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                              " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                        " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                  " & vbCrLf & _
              "       Inner Join(                                                                                                                                   " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                        " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                         " & vbCrLf & _
              "       		 ) as axm                                                                                                                                " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                         " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                           " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                       " & vbCrLf & _
              "      Where a.Empresa_Id     = Left(@Empresa, 8)                                                                                                                " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                 " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                                " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                             " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                      PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, isnull(axc.DepreciacaoCredito, '') as ContaDepreciacao,                                                           " & vbCrLf & _
              "           '' as Cliente, 0 as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)    " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, 0 as DebitoOficial, axm.ValorDepreciado as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                                'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                                'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                                " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                  " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                              " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                        " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                  " & vbCrLf & _
              "       Inner Join(                                                                                                                                   " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                        " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                         " & vbCrLf & _
              "       		 ) as axm                                                                                                                                " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                         " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                           " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                       " & vbCrLf & _
              "       Where a.Empresa_Id    = Left(@Empresa, 8)                                                                                                               " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                 " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                                " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                             " & vbCrLf & _
              "                                                                                                                                                     " & vbCrLf & _
              "   INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id,               " & vbCrLf & _
              "                      Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico,                  " & vbCrLf & _
              "                      PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)                               " & vbCrLf & _
              "    Select a.Empresa, a.EndEmpresa, isnull(axc.ContaDeTransferencia, '') as ContaDepreciacao,                                                          " & vbCrLf & _
              "           '" & ViewState("EmpresaOrigem").Split("-")(0) & "' as Cliente, '" & ViewState("EmpresaOrigem").Split("-")(1) & "' as EndCliente,  '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Movimento, 53 as Lote, (Select isnull(MAX(Sequencia_Id) + 1, 1)    " & vbCrLf & _
              "                                                                                          from Razao                                                 " & vbCrLf & _
              "                                                                                         Where Empresa_Id    = a.Empresa                             " & vbCrLf & _
              "                                                                                           and EndEmpresa_Id = a.EndEmpresa                          " & vbCrLf & _
              "                                                                                           and Lote_Id       = 53) as Sequencia,          " & vbCrLf & _
              "              a.CentroDeCusto AS Custo, 3 as Indexador, '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as DataMoeda, axm.ValorDepreciado as DebitoOficial, 0 as CreditoOficial, 0 as DebitoMoeda, 0 as CreditoMoeda, " & vbCrLf & _
              "                                        'TRANSFERENCIA DO BEM: ' + a.Grupo_Id + ' - ' + cast(a.Codigo_Id as varchar) + ' - ' + cast(a.Sequencia_Id as varchar) + ' - ATE O DIA: " & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "' as Historico," & vbCrLf & _
              "                                        'P' as PrevistoRealizado, 'AtivosXTransferencia' as Processo, '" & UsuarioServidor.NomeUsuario & "' AS UsuarioInclusaoData, GETDATE() as UsuarioInclusao, a.Grupo_Id + '-' + CAST(a.Codigo_Id as varchar) + '-' + cast(a.Sequencia_Id as varchar) as AtivoImobilizado, " & vbCrLf & _
              "                 (Select isnull(MAX(Sequencia_Id) + 1, 1) as maxSequencia" & vbCrLf & _
              "                    from Razao " & vbCrLf & _
              "                   where Empresa_Id    = a.Empresa" & vbCrLf & _
              "                     and EndEmpresa_Id   = a.EndEmpresa" & vbCrLf & _
              "                     and Lote_Id         = 53" & vbCrLf & _
              "                     and Movimento_id    = '" & CDate(txtDataTransferencia.Text).ToString("yyyy-MM-dd") & "') as AgrupadorDeLancamento" & vbCrLf & _
              "       	from Ativos a                                                                                                                               " & vbCrLf & _
              "       Inner Join AtivosXContas axc                                                                                                                 " & vbCrLf & _
              "       	  On a.Empresa = axc.Empresa_Id                                                                                                             " & vbCrLf & _
              "       	 And a.EndEmpresa = axc.EndEmpresa_Id                                                                                                       " & vbCrLf & _
              "         And a.Conta = axc.Conta_Id                                                                                                                 " & vbCrLf & _
              "       Inner Join(                                                                                                                                  " & vbCrLf & _
              "       		 SELECT Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, COALESCE(sum(Valor), 0) as ValorDepreciado                                       " & vbCrLf & _
              "       		   FROM AtivosXMovimentos group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id                                                        " & vbCrLf & _
              "       		 ) as axm                                                                                                                               " & vbCrLf & _
              "       	  On axm.Empresa_Id   = a.Empresa_Id                                                                                                        " & vbCrLf & _
              "       	 And axm.Grupo_Id     = a.Grupo_Id                                                                                                          " & vbCrLf & _
              "       	 and axm.Codigo_Id    = a.Codigo_Id                                                                                                         " & vbCrLf & _
              "       	 and axm.Sequencia_Id = a.Sequencia_Id                                                                                                      " & vbCrLf & _
              "       Where a.Empresa_Id    = Left(@Empresa, 8)                                                                                                              " & vbCrLf & _
              "       	 And a.Grupo_Id     = @grupo                                                                                                                " & vbCrLf & _
              "       	 And a.Codigo_Id    = @Codigo                                                                                                               " & vbCrLf & _
              "       	 And a.Sequencia_Id = @Sequencia                                                                                                            " & vbCrLf
        SqlArray.Add(Sql)
    End Sub

    Private Sub getSqlContabilizacao(ByVal valorDepreciado As Decimal, ByRef sqlArray As ArrayList)
        Try
            Dim numero As Integer = getNumeroAgrupador(ViewState("EmpresaOrigem").Split("-")(0), ViewState("EmpresaOrigem").Split("-")(1), CDate(txtDataTransferencia.Text))

            Sql = "Select a.Empresa, a.EndEmpresa, isnull(axc.DepreciacaoDebito, '') as Debita, isnull(axc.DepreciacaoCredito, '') as Credita, " & vbCrLf & _
                  " 	   a.Codigo_Id, a.Sequencia_Id, a.Grupo_Id,  a.CentroDeCusto AS Custo                                                  " & vbCrLf & _
                  " 	from Ativos a                                                                                                          " & vbCrLf & _
                  " 	Inner Join AtivosXContas axc                                                                                           " & vbCrLf & _
                  " 		 On a.Empresa        = axc.Empresa_Id                                                                              " & vbCrLf & _
                  " 		And a.EndEmpresa     = axc.EndEmpresa_Id                                                                           " & vbCrLf & _
                  " 		And a.Conta          = axc.Conta_Id                                                                                " & vbCrLf & _
                  " 	Where a.Empresa_Id     = '" & Left(ViewState("EmpresaOrigem").Split("-")(0), 8) & "'" & vbCrLf & _
                  " 		And a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                  " 		And a.Codigo_Id    = " & CInt(txtCodigo.Text) & vbCrLf & _
                  " 		And a.Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf

            For Each row As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows

                If Not String.IsNullOrWhiteSpace(row("Debita")) Then
                    Sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, " & vbCrLf & _
                          "                   Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico," & vbCrLf & _
                          "                   PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                          "           VALUES ('" & row("Empresa") & "', " & row("EndEmpresa") & ", '" & row("Debita") & "', '', 0" & ", '" & txtDataTransferencia.Text.ToSqlDate() & "', 53, (" & getNumerador(ViewState("EmpresaOrigem").Split("-")(0), ViewState("EmpresaOrigem").Split("-")(1), CDate(txtDataTransferencia.Text)) & "), " & vbCrLf & _
                          "                    " & row("Custo") & ", 3" & ", '" & txtDataTransferencia.Text.ToSqlDate() & "', " & Str(valorDepreciado) & ", 0, 0, 0, " & vbCrLf & _
                          "                   'DEPRECIACAO PARCIAL DO BEM:  " & row("Grupo_Id") & " - " & row("Codigo_Id").ToString("0000") & " - " & row("Sequencia_Id").ToString("000") & " ATE O DIA: " & txtDataTransferencia.Text.ToSqlDate() & "', " & vbCrLf & _
                          "                   'P', 'AtivosXTransferencia', '" & UsuarioServidor.NomeUsuario & "', getdate(), " & vbCrLf & _
                          "                   '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "', " & numero & ")" & vbCrLf
                    sqlArray.Add(Sql)
                End If

                If Not String.IsNullOrWhiteSpace(row("Credita")) Then
                    Sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, " & vbCrLf & _
                          "                   Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, " & vbCrLf & _
                          "                   PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                          "           VALUES ('" & row("Empresa") & "', " & row("EndEmpresa") & ", '" & row("Credita") & "', '', 0" & ", '" & txtDataTransferencia.Text.ToSqlDate() & "', 53, (" & getNumerador(ViewState("EmpresaOrigem").Split("-")(0), ViewState("EmpresaOrigem").Split("-")(1), CDate(txtDataTransferencia.Text)) & "), " & vbCrLf & _
                          "                    " & row("Custo") & ", 3" & ", '" & txtDataTransferencia.Text.ToSqlDate() & "', 0, " & Str(valorDepreciado) & ", 0, 0, " & vbCrLf & _
                          "                   'DEPRECIACAO PARCIAL DO BEM:  " & row("Grupo_Id") & " - " & row("Codigo_Id").ToString("0000") & " - " & row("Sequencia_Id").ToString("000") & " ATE O DIA: " & txtDataTransferencia.Text.ToSqlDate() & "', " & vbCrLf & _
                          "                   'P', 'AtivosXTransferencia', '" & UsuarioServidor.NomeUsuario & "', getdate(), " & vbCrLf & _
                          "                   '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "', " & numero & ")" & vbCrLf
                    sqlArray.Add(Sql)
                End If
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getNumerador(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal movimento As DateTime) As String
        Dim sql As String = "Select isnull(MAX(Sequencia_Id) + 1, 1)              " & vbCrLf & _
                            "  from Razao " & vbCrLf & _
                            " where Empresa_Id    =  '" & Empresa & "'" & vbCrLf & _
                            "   and EndEmpresa_Id =  " & EndEmpresa & vbCrLf & _
                            "   and Lote_Id       = 53 " & vbCrLf & _
                            "   and movimento_id  = '" & movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf
        Return sql
    End Function

    Private Function getNumeroAgrupador(ByVal Empresa As String, ByVal EndEmpresa As Integer, ByVal movimento As DateTime) As Integer
        Dim sql As String = "Select isnull(MAX(Sequencia_Id) + 1, 1) as num             " & vbCrLf & _
                            "  from Razao " & vbCrLf & _
                            " where Empresa_Id    =  '" & Empresa & "'" & vbCrLf & _
                            "   and EndEmpresa_Id =  " & EndEmpresa & vbCrLf & _
                            "   and Lote_Id       = 53 " & vbCrLf & _
                            "   and movimento_id  = '" & movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "tabela")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Return row("num")
            Next
        End If
        Return 1
    End Function

    Private Function ExisteContaTransf() As Boolean
        Dim sql As String = "select distinct ContaDeTransferencia from ativosxcontas where isnull(ContaDeTransferencia, '') <> '' and empresa_id = '" & DdlEmpresaOrigem.SelectedValue.Split("-")(0) & "' and endempresa_id = " & DdlEmpresaOrigem.SelectedValue.Split("-")(1)

        Return Banco.ConsultaDataSet(sql, "ContaTransferencia").Tables(0).Rows.Count > 0
    End Function

    Private Function ValidarTransferencia() As Boolean
        If String.IsNullOrWhiteSpace(txtDataTransferencia.Text) Then
            MsgBox(Me.Page, "Informe a data da Transferência.")
            Return False
        ElseIf Not IsDate(txtDataTransferencia.Text) Then
            MsgBox(Me.Page, "Data da Transferência não é uma data válida.")
            Return False
        ElseIf CDate(txtDataTransferencia.Text) < CDate(txtAtualizacao.Text) Then
            MsgBox(Me.Page, "Data da Transferência não pode ser menor que data de última atualização.")
            Return False
        ElseIf (CDate(txtAtualizacao.Text).Year = CDate(txtDataTransferencia.Text).Year) AndAlso (CDate(txtAtualizacao.Text).Month = CDate(txtDataTransferencia.Text).Month) Then
            MsgBox(Me.Page, "Impossível efetuar transferência neste mês. Processo mensal ja efetuado. Elimine o processo deste mês para efetuar transferência.")
            Return False
        ElseIf (CDate(txtAtualizacao.Text).AddMonths(1).Month <> CDate(txtDataTransferencia.Text).Month And ((CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)) > 0)) _
            OrElse (CDate(txtAtualizacao.Text).AddMonths(1).Year <> CDate(txtDataTransferencia.Text).Year And ((CDec(txtValorOriginal.Text) - CDec(txtValorDepreciado.Text)) > 0)) Then
            MsgBox(Me.Page, "Não é possivel transferir o bem com mês de baixa diferente de: " & MonthName(CDate(txtAtualizacao.Text).AddMonths(1).Month) & " de: " & CDate(txtAtualizacao.Text).AddMonths(1).Year)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtMotivoTransferencia.Text) Then
            MsgBox(Me.Page, "Informe o Motivo da Transferência.")
            Return False
        ElseIf Len(txtMotivoTransferencia.Text) < 10 Then
            MsgBox(Me.Page, "Informe no mínimo 10 caracteres para o Motivo da Transferência.")
            Return False
        ElseIf ViewState("EmpresaOrigem") = DdlEmpresaDest.SelectedValue Then
            MsgBox(Me.Page, "Informe uma empresa diferente para transferência.")
            Return False
        End If
        Return True
    End Function

    Private Function validaDesfazTransferencia() As Boolean
        Dim sql As String = "   select a.Empresa, am.Empresa, MAX(am.Movimento_Id) as Movimento_Id from Ativos a  " & vbCrLf & _
                            "    Inner Join AtivosXMovimentos am                                                  " & vbCrLf & _
                            "   	On am.Empresa_Id = a.Empresa_Id                                               " & vbCrLf & _
                            "      and am.Grupo_Id = a.Grupo_Id                                                   " & vbCrLf & _
                            "      and am.Codigo_Id = a.Codigo_Id                                                 " & vbCrLf & _
                            "      and am.Sequencia_Id = a.Sequencia_Id                                           " & vbCrLf & _
                            "      and a.Atualizacao = am.Movimento_Id                                            " & vbCrLf & _
                            "    Where a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                          " & vbCrLf & _
                            "      and a.Codigo_Id    = " & CInt(txtCodigo.Text) & "                              " & vbCrLf & _
                            "      and a.Sequencia_Id = " & CInt(txtSequencia.Text) & "                           " & vbCrLf & _
                            "      and a.Atualizacao = '" & CDate(txtAtualizacao.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                            "    group by a.Empresa, am.Empresa                                                   " & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Transferencia")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            If ds.Tables(0).Rows(0)(0).ToString() <> ds.Tables(0).Rows(0)(1).ToString() Then
                Return True
            End If
        End If
        Return False
    End Function

    Private Function HabilitarCampos(Optional ByVal Enable As Boolean = True) As Boolean
        DdlGrupo.Enabled = Enable
        txtAquisicao.Enabled = Enable
        txtInicioDeUso.Enabled = Enable
        txtDescricao.Enabled = Enable
        DdlGrupo.Enabled = Enable
        DdlGrupoDeContas.Enabled = Enable
        DdlContaContabil.Enabled = Enable
        DdlCentroDeCusto.Enabled = Enable
        txtHistorico.Enabled = Enable
        txtValorOriginal.Enabled = Enable
        txtValorDepreciadoInicial.Enabled = Enable
        ddlDepreciar.Enabled = Enable
        txtAtualizacao.Enabled = Enable
        chkSeguro.Enabled = Enable
    End Function

    Private Sub CargaGruposDeContas()
        ddl.Carregar(DdlGrupoDeContas, CarregarDDL.Tabela.PlanoDeContas, "(len(conta_id) = 7 and left(conta_id, 1) = '1')", True)
    End Sub

    Private Sub CargaCentroDeCusto()
        ddl.Carregar(DdlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "Len(CentroDeCusto_Id) = 5", True)
    End Sub

    Private Sub CarregarContas()
        ddl.Carregar(DdlContaContabil, CarregarDDL.Tabela.PlanoDeContas, "conta_id like '" & DdlGrupoDeContas.SelectedValue & "%' and len(Conta_Id) = 9", True)
    End Sub

    Private Sub CargaGrupos()
        ddl.Carregar(DdlGrupo, CarregarDDL.Tabela.GruposDeAtivos, DdlEmpresaOrigem.SelectedValue, True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If Not ExisteContaTransf() Then
            MsgBox(Me.Page, "Atenção, não existe conta de transferencia cadastrada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
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
        Return ValidarTransferencia()
    End Function

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

        DdlGrupoDeContas.SelectedIndex = 0

        If DdlGrupo.Items.Count > 0 Then DdlGrupo.SelectedIndex = 0

        DdlContaContabil.Items.Clear()
        DdlCentroDeCusto.SelectedIndex = 0

        txtValorDepreciadoInicial.Text = String.Empty

        txtAquisicao.Text = String.Empty
        txtAtualizacao.Text = String.Empty
        txtInicioDeUso.Text = String.Empty

        txtMotivoTransferencia.Text = String.Empty
        txtDataTransferencia.Text = String.Empty
        habilitarBotoes(False)

        TabContainer1.ActiveTabIndex = 0

        If DdlEmpresaOrigem.SelectedIndex > 0 Then CargaEmpresaDestino()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresaOrigem.Enabled = False
        End If
    End Sub

    Private Sub CargaAtivos()
        Sql = " SELECT a.Grupo_Id as Grupo, a.Codigo_Id as Codigo, a.Sequencia_Id as Sequencia, a.Descricao, a.DataAquisicao as Aquisicao, a.ValorOriginal as Valor, a.DataDaBaixa  " & vbCrLf & _
              "   FROM Ativos a                                                                                                                                                     " & vbCrLf & _
              "  Inner Join AtivosXMovimentos am                                                                                                                                    " & vbCrLf & _
              "     On am.Empresa_Id    = a.Empresa_Id                                                                                                                              " & vbCrLf & _
              "    and am.Grupo_Id     = a.Grupo_Id                                                                                                                                 " & vbCrLf & _
              "    and am.Codigo_Id    = a.Codigo_Id                                                                                                                                " & vbCrLf & _
              "    and am.Sequencia_Id = a.Sequencia_Id                                                                                                                             " & vbCrLf & _
              "  where Situacao = 1 and a.DataDaBaixa  is null                                                                                                                      " & vbCrLf & _
              "    and a.Empresa_Id = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlEmpresaDest.SelectedValue) Then
            Sql &= " And a.Empresa = '" & DdlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
            Sql &= " And a.Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf
        End If

        Sql &= "  group by a.Grupo_Id, a.Codigo_Id, a.Sequencia_Id, a.Descricao, a.DataAquisicao, a.ValorOriginal,  a.DataDaBaixa                                                    " & vbCrLf & _
               "  order by a.Grupo_Id, a.Codigo_Id, a.Sequencia_Id, a.Descricao, a.DataAquisicao, a.ValorOriginal                                                                    " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")

        If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            GridAtivos.DataSource = ds
            GridAtivos.DataBind()
            TabContainer1.ActiveTabIndex = 1
        Else
            MsgBox(Me.Page, "Nenhum resultado encontrado.")
        End If
    End Sub

    Private Sub GerarRelatorioTransferencia(ByVal data As DateTime, Optional ByVal reemissao As Boolean = False)
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
                            "   	on ga.Grupo_Id   = a.Grupo_Id		                                                                       " & vbCrLf & _
                            "   Where a.Empresa_Id   = '" & Left(DdlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'                             " & vbCrLf & _
                            "     and a.Grupo_Id     = '" & DdlGrupo.SelectedValue & "'                                                        " & vbCrLf & _
                            "     and a.Codigo_Id    = " & CInt(txtCodigo.Text) & vbCrLf & _
                            "     And a.Sequencia_Id = " & CInt(txtSequencia.Text) & vbCrLf & _
                            "   group by a.codigo_Id, a.Grupo_Id, ga.Descricao, a.Sequencia_Id, a.Descricao, a.Historico,                  " & vbCrLf & _
                            "            a.DataAquisicao, a.InicioDeUso, a.ValorOriginal, ga.PercentualDepreciacao,                        " & vbCrLf & _
                            "            a.ValorDaBaixa, a.DataDaBaixa, a.MotivoDaBaixa                                                    " & vbCrLf & _
                            "                                                                                                              " & vbCrLf & _
                            "   select case when c.Reduzido is not null" & vbCrLf & _
                            "               then CAST(r.Conta_Id as varchar) + '-' +  REPLICATE('0', 5 - LEN(c.Reduzido)) + RTrim(c.Reduzido)" & vbCrLf & _
                            "               else r.Conta_Id" & vbCrLf & _
                            "          end as Conta_Id," & vbCrLf & _
                            "          r.Historico, r.DebitoOficial as Debito, r.CreditoOficial as Credito, r.Sequencia_Id" & vbCrLf & _
                            "     from Razao r                                                                                             " & vbCrLf & _
                            "     Left Join Clientes c" & vbCrLf & _
                            "       on c.Cliente_Id = r.Cliente_Id" & vbCrLf & _
                            "      and c.Endereco_Id = r.EndCliente_Id" & vbCrLf & _
                            "    WHERE r.Lote_id          = 53 " & vbCrLf & _
                            "      AND r.Movimento_Id     = '" & data.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                            "      And r.AtivoImobilizado = '" & DdlGrupo.SelectedValue & "-" & CInt(txtCodigo.Text) & "-" & CInt(txtSequencia.Text) & "'" & vbCrLf & _
                            "    order by r.UsuarioInclusaoData " & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Ativos")

        If ds IsNot Nothing AndAlso ds.Tables.Count = 2 Then ds.Tables(1).TableName = "Razao"

        Funcoes.BindReport(Me.Page, ds, "Cr_AtivosXTransferencia", eExportType.PDF)
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

        If Not String.IsNullOrWhiteSpace(unidade) Then
            sql &= "And gxe.Empresa_Id = '" & unidade & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Empresa) Then
            sql &= "And Left(cxe.Empresa_Id, 8) = '" & Left(Empresa, 8) & "'"
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
#End Region

End Class
