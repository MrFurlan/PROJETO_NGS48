Imports System.IO
Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class Ativos
    Inherits BasePage

    Private Sql As String

#Region "Evenths"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then

                If Not UsuarioServidor.KeyCodeActive Then
                    MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Patrimonio.aspx", eTitulo.Info)
                    Exit Sub
                End If

                If Funcoes.VerificaPermissao("Ativos", "ACESSAR") Then
                    ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    CargaUnidadeDeNegocio()
                    ddl.Carregar(ddlCentroDeCusto, CarregarDDL.Tabela.CentroDeCusto, "Len(CentroDeCusto_Id) = 5", True)
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

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
            HabilitarCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Ativos", "GRAVAR") Then
                Dim SqlArray As New ArrayList
                Dim at As New Ativo()

                If Validar() Then
                    Sql = "INSERT Into Ativos(Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Situacao, UnidadeDeNegocio, Empresa, EndEmpresa, " & vbCrLf & _
                          "                   CentroDeCusto, Conta, Descricao, Historico, DataAquisicao, InicioDeUso, " & vbCrLf & _
                          "                   ValorOriginal, Depreciar, Atualizacao, QuemLancou, QuandoLancou, Seguro) " & vbCrLf & _
                          "            Values( '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "', '" & DdlGrupo.SelectedValue & "', " & CInt(txtCodigo.Text) & ", " & vbCrLf & _
                          "                     " & CInt(txtSequencia.Text) & ", " & "1" & ", '" & ddlUnidade.SelectedValue & "', '" & ddlEmpresaDest.SelectedValue.ToString.Split("-")(0) & "', " & vbCrLf & _
                          "                     " & ddlEmpresaDest.SelectedValue.Split("-")(1) & ", '" & ddlCentroDeCusto.SelectedValue & "', '" & ddlContaContabil.SelectedValue & "', " & vbCrLf & _
                          "                    '" & txtDescricao.Text.Trim & "', '" & txtHistorico.Text.Trim.ToUpper() & "', '" & Format(CDate(txtAquisicao.Text), "yyyy/MM/dd") & "', " & vbCrLf & _
                          "                    '" & Format(CDate(txtInicioDeUso.Text), "yyyy/MM/dd") & "', " & Str(CDec(txtValorOriginal.Text)) & ", " & vbCrLf & _
                          "                    '" & ddlDepreciar.SelectedValue & "', '" & Format(CDate(txtAtualizacao.Text), "yyyy/MM/dd") & "', " & vbCrLf & _
                          "                    '" & UsuarioServidor.NomeUsuario & "', '" & Format(Today, "yyyy/MM/dd") & "', " & vbCrLf & _
                          "                    '" & IIf(chkSeguro.Checked, "TRUE", "FALSE") & "')" & vbCrLf
                    SqlArray.Add(Sql)

                    If txtValorDepreciadoInicial.Text = "" Then txtValorDepreciadoInicial.Text = "0,00"

                    Sql = "INSERT Into AtivosXMovimentos (Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Movimento_Id, Valor, Processo, Empresa, EndEmpresa)" & vbCrLf & _
                          "         Values( '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "', '" & DdlGrupo.SelectedValue & "', " & CInt(txtCodigo.Text) & ", " & CInt(txtSequencia.Text) & ", " & vbCrLf & _
                          "                 '" & Format(CDate(txtInicioDeUso.Text), "yyyy/MM/dd") & "', " & vbCrLf & _
                          "                  " & Str(CDec(txtValorDepreciadoInicial.Text)) & ",'INICIAL'," & vbCrLf & _
                          "                 '" & ddlEmpresaDest.SelectedValue.ToString.Split("-")(0) & "', " & ddlEmpresaDest.SelectedValue.ToString.Split("-")(1) & ")" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        If chkReaproveitarDados.Checked Then
                            IncrementaCodigo()
                            CargaAtivos()
                            txtValorOriginal.Text = "0,00"
                            txtValorDepreciadoInicial.Text = "0,00"
                        Else
                            Limpar()
                        End If

                        HabilitarCampos()
                        txtSequencia.Text = "000"
                        MsgBox(Me.Page, "Registro Gravado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Ativos", "ALTERAR") Then
                If Validar() Then
                    Dim SqlArray As New ArrayList

                    Sql = "Update Ativos " & vbCrLf & _
                      "     Set   DataAquisicao   = '" & txtAquisicao.Text.ToSqlDate() & "'" & vbCrLf & _
                      "         , InicioDeUso   = '" & txtInicioDeUso.Text.ToSqlDate() & "'" & vbCrLf & _
                      "         , Descricao     = '" & txtDescricao.Text.Trim & "'" & vbCrLf & _
                      "         , Empresa       = '" & ddlEmpresaDest.SelectedValue.ToString.Split("-")(0) & "'" & vbCrLf & _
                      "         , EndEmpresa    =  " & ddlEmpresaDest.SelectedValue.ToString.Split("-")(1) & vbCrLf & _
                      "         , Conta         = '" & ddlContaContabil.SelectedValue & "'" & vbCrLf & _
                      "         , CentroDeCusto =  " & ddlCentroDeCusto.SelectedValue & vbCrLf & _
                      "         , Historico     = '" & txtHistorico.Text.Trim.ToUpper() & "'" & vbCrLf & _
                      "         , ValorOriginal =  " & Str(CDec(txtValorOriginal.Text)) & vbCrLf & _
                      "         , Atualizacao   = '" & txtAtualizacao.Text.ToSqlDate() & "'" & vbCrLf & _
                      "         , Depreciar     = '" & ddlDepreciar.SelectedValue & "'" & vbCrLf & _
                      "         , QuemAlterou   = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                      "         , QuandoAlterou = '" & Now.ToSqlDate() & "'" & vbCrLf & _
                      "         , Seguro        = '" & IIf(chkSeguro.Checked, "TRUE", "FALSE") & "'" & vbCrLf & _
                      "     WHERE Empresa_Id     = '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                      "         AND Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                      "         AND Codigo_Id    =  " & txtCodigo.Text & vbCrLf & _
                      "         AND Sequencia_Id =  " & txtSequencia.Text & vbCrLf
                    SqlArray.Add(Sql)

                    Sql = "Update AtivosXMovimentos " & vbCrLf &
                          "     Set   Valor             = " & Str(CDec(txtValorDepreciadoInicial.Text)) & vbCrLf &
                      "               ,Movimento_id     = '" & txtInicioDeUso.Text.ToSqlDate() & "'" & vbCrLf &
                      "     WHERE Empresa_Id            = '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf &
                      "         AND Grupo_Id            = '" & DdlGrupo.SelectedValue & "'" & vbCrLf &
                      "         AND Codigo_Id           =  " & txtCodigo.Text & vbCrLf &
                      "         AND Sequencia_Id        =  " & txtSequencia.Text & vbCrLf &
                      "         AND Processo            = 'INICIAL'" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        If chkReaproveitarDados.Checked Then
                            IncrementaCodigo()
                            CargaAtivos()

                            txtValorOriginal.Text = "0,00"
                            txtValorDepreciadoInicial.Text = "0,00"
                            lnkNovo.Parent.Visible = True
                            lnkAtualizar.Parent.Visible = False
                            lnkExcluir.Parent.Visible = False

                        Else
                            Limpar()
                        End If

                        HabilitarCampos()
                        txtSequencia.Text = "000"
                        MsgBox(Me.Page, "Registro Atualizado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Ativos", "EXCLUIR") Then
                If ExisteMovimentos() Then
                    MsgBox(Me.Page, "Registro contém movimentos. Exclusão negada.")
                ElseIf String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigo.Text) OrElse String.IsNullOrWhiteSpace(txtSequencia.Text) Then
                    MsgBox(Me.Page, "É necessário Selecionar um bem a ser depreciado.")
                Else
                    Sql = "UPDATE Ativos SET Situacao = 3" & vbCrLf & _
                          "                 , QuemExcluiu   = '" & UsuarioServidor.NomeUsuario & "'" & vbCrLf & _
                          "                 , QuandoExcluiu = '" & Now.ToSqlDate() & "'" & vbCrLf & _
                          "     WHERE Empresa_Id     = '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                          "         AND Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                          "         AND Codigo_Id    =  " & txtCodigo.Text & vbCrLf & _
                          "         AND Sequencia_Id =  " & txtSequencia.Text

                    If Banco.GravaBanco(Sql) Then
                        Limpar()
                        HabilitarCampos()
                        MsgBox(Me.Page, "Registro Excluído com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridAtivos_RowDataBound(sender As Object, e As GridViewRowEventArgs) Handles GridAtivos.RowDataBound
        Try
            If e.Row.RowType.Equals(DataControlRowType.DataRow) Then
                If e.Row.Cells(7).Text.Equals("Sim") Then
                    e.Row.ForeColor = Drawing.Color.Red
                    e.Row.Attributes.Add(" data-ToolTip", "default")
                    e.Row.Attributes.Add("title", "Ativo baixado.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub GridAtivos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim Grupo As String = GridAtivos.SelectedRow.Cells(1).Text()
            Dim Codigo As String = GridAtivos.SelectedRow.Cells(2).Text()
            Dim Sequencia As String = GridAtivos.SelectedRow.Cells(3).Text()

            ddlUnidade.Enabled = False
            DdlEmpresaOrigem.Enabled = False
            DdlGrupo.Enabled = False

            Sql = " SELECT  Ativos.Empresa_Id, Ativos.Grupo_Id, Ativos.Codigo_Id, Ativos.Sequencia_Id, Ativos.DataAquisicao, Ativos.InicioDeUso, Ativos.Descricao, " & vbCrLf & _
                  "         Ativos.UnidadeDeNegocio, Ativos.Empresa, Ativos.EndEmpresa, Ativos.Conta, Ativos.CentroDeCusto, CentrosDeCustos.Descricao AS NomeDoCusto, " & vbCrLf & _
                  "         Ativos.Historico, Ativos.ValorOriginal, Ativos.Atualizacao, Ativos.Depreciar, Ativos.Seguro, Ativos.DataDaBaixa " & vbCrLf & _
                  " FROM    Ativos" & vbCrLf & _
                  "     LEFT JOIN   CentrosDeCustos " & vbCrLf & _
                  "         ON Ativos.CentroDeCusto = CentrosDeCustos.CentroDeCusto_Id" & vbCrLf & _
                  " WHERE   Ativos.Empresa_Id   = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                  "     AND Ativos.Grupo_Id     = '" & Grupo & "'" & vbCrLf & _
                  "     AND Ativos.Codigo_Id    =  " & Codigo & vbCrLf & _
                  "     AND Ativos.Sequencia_Id =  " & Sequencia & vbCrLf

            Dim row As DataRow = Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows(0)

            DdlGrupo.SelectedValue = row("Grupo_Id")
            ddlUnidade.SelectedValue = row("UnidadeDeNegocio")

            ddlEmpresaDest.SelectedValue = String.Format("{0}-{1}", row("Empresa"), row("EndEmpresa"))
            ddl.Carregar(ddlContaContabil, CarregarDDL.Tabela.AtivosXContas, ddlEmpresaDest.SelectedValue.Split("-")(0))

            ddlContaContabil.SelectedIndex = ddlContaContabil.Items.IndexOf(ddlContaContabil.Items.FindByValue(row("Conta")))
            If Not ddlContaContabil.SelectedIndex > 0 Then
                MsgBox(Me.Page, "A conta contábil " & row("Conta") & " cadastrada no Bem não foi encontrada na lista, verifique a Tabela AtivosXContas se a conta está referenciada com a Filial de Locação.")
            End If

            ddlCentroDeCusto.SelectedIndex = ddlCentroDeCusto.Items.IndexOf(ddlCentroDeCusto.Items.FindByValue(row("CentroDeCusto")))

            txtCodigo.Text = Format(row("Codigo_Id"), "0000")
            txtSequencia.Text = Format(row("Sequencia_Id"), "000")
            txtAquisicao.Text = row("DataAquisicao")
            txtInicioDeUso.Text = row("InicioDeUso")
            txtDescricao.Text = row("Descricao")
            txtHistorico.Text = row("Historico")
            txtValorOriginal.Text = row("ValorOriginal").ToString()
            ddlDepreciar.SelectedValue = row("Depreciar")

            txtAtualizacao.Text = row("Atualizacao")
            txtDataBaixa.Text = IIf(IsDBNull(row("DataDaBaixa")), "", row("DataDaBaixa"))

            If row("Seguro") = "TRUE" Then
                chkSeguro.Checked = True
            Else
                chkSeguro.Checked = False
            End If

            Sql = " SELECT COALESCE(sum(Valor), 0) as Valor" & vbCrLf & _
                  "     FROM AtivosXMovimentos" & vbCrLf & _
                  "     WHERE Empresa_Id = '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                  "         AND Grupo_Id  = '" & Grupo & "'" & vbCrLf & _
                  "         AND Codigo_Id = " & Codigo & vbCrLf & _
                  "         AND Sequencia_Id = " & Sequencia & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "AtivosXMovimentos").Tables(0).Rows
                txtValorDepreciado.Text = Dr("Valor").ToString()
            Next

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql & " And Processo = 'INICIAL' ", "AtivosXMovimentos").Tables(0).Rows
                txtValorDepreciadoInicial.Text = Dr("Valor").ToString()
            Next

            If Not String.IsNullOrWhiteSpace(txtDataBaixa.Text) AndAlso IsDate(txtDataBaixa.Text) Then
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
            Else
                lnkAtualizar.Parent.Visible = True
                lnkExcluir.Parent.Visible = True
            End If

            lnkNovo.Parent.Visible = False

            If Not String.IsNullOrWhiteSpace(txtDataBaixa.Text) AndAlso IsDate(txtDataBaixa.Text) Then
                HabilitarCampos(False)
                ddlSequencia.Enabled = False
            ElseIf ExisteMovimentos() Then
                HabilitarCampos(False, True)
            End If

            TabContainer1.ActiveTabIndex = 0

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString())
        End Try
    End Sub

    Protected Sub txtSequencia_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlSequencia.Click
        Try
            If String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                MsgBox(Me.Page, "É necessário informar o Grupo e Código.")
            Else
                IncrementaSequencia()

                HabilitarCampos()

                txtAquisicao.Text = String.Empty
                txtInicioDeUso.Text = String.Empty
                txtDescricao.Text = String.Empty
                txtHistorico.Text = String.Empty
                txtValorOriginal.Text = String.Empty
                txtValorDepreciado.Text = String.Empty
                ddlDepreciar.SelectedValue = "N"

                txtAtualizacao.Text = String.Empty
                lnkNovo.Parent.Visible = True
                lnkAtualizar.Parent.Visible = False
                lnkExcluir.Parent.Visible = False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlUnidade.SelectedIndex > 0 Then
                CarregarEmpresas(DdlEmpresaOrigem, ddlUnidade.SelectedValue, "", True)
                ddlEmpresaDest.Items.Clear()
            Else
                DdlEmpresaOrigem.Items.Clear()
                ddlEmpresaDest.Items.Clear()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
                If Not String.IsNullOrWhiteSpace(DdlEmpresaOrigem.SelectedValue) Then
                    IncrementaCodigo()
                    txtSequencia.Text = "000"
                    CargaAtivos()
                    CarregarDataDaUltimaAtualizacao(DdlEmpresaOrigem.SelectedValue, chkSeguro.Checked)
                Else
                    MsgBox(Me.Page, "Informe a Empresa a qual pertence o bem.")
                    DdlGrupo.SelectedValue = String.Empty
                End If
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
    Protected Sub DdlEmpresaOrigem_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresaOrigem.SelectedIndexChanged
        Try
            CargaEmpresaDestino()
            CarregarDataDaUltimaAtualizacao(DdlEmpresaOrigem.SelectedValue, chkSeguro.Checked)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Ativos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub chkSeguro_CheckedChanged(sender As Object, e As EventArgs) Handles chkSeguro.CheckedChanged
        Try
            CarregarDataDaUltimaAtualizacao(DdlEmpresaOrigem.SelectedValue, CType(sender, CheckBox).Checked)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ddlEmpresaDest_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlEmpresaDest.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlEmpresaDest.SelectedValue) Then
                ddl.Carregar(ddlContaContabil, CarregarDDL.Tabela.AtivosXContas, ddlEmpresaDest.SelectedValue.Split("-")(0))
            Else
                ddlContaContabil.Items.Clear()
            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Sub HabilitarCampos(Optional ByVal Enable As Boolean = True, Optional ByVal ExisteMovimentos As Boolean = False)
        DdlGrupo.Enabled = Enable
        txtInicioDeUso.Enabled = Enable

        If ExisteMovimentos Then
            txtDescricao.Enabled = True
            ddlCentroDeCusto.Enabled = True
            txtHistorico.Enabled = True
            lnkExcluir.Parent.Visible = False
        Else
            txtDescricao.Enabled = Enable
            ddlCentroDeCusto.Enabled = Enable
            txtHistorico.Enabled = Enable
            lnkExcluir.Parent.Visible = True
        End If

        ddlEmpresaDest.Enabled = Enable
        DdlGrupo.Enabled = Enable
        ddlContaContabil.Enabled = Enable

        txtValorOriginal.Enabled = Enable
        txtValorDepreciadoInicial.Enabled = Enable
        ddlDepreciar.Enabled = Enable
        txtAtualizacao.Enabled = Enable
        chkSeguro.Enabled = Enable
    End Sub
    Private Sub CargaUnidadeDeNegocio()
        ddlUnidade.SelectedIndex = ddlUnidade.Items.IndexOf(ddlUnidade.Items.FindByValue(UsuarioServidor.Usuario.AcessoUnidade))
        CarregarEmpresas(DdlEmpresaOrigem, ddlUnidade.SelectedValue, "", True)
        DdlEmpresaOrigem.SelectedIndex = DdlEmpresaOrigem.Items.IndexOf(DdlEmpresaOrigem.Items.FindByValue(UsuarioServidor.Usuario.AcessoEmpresa & "-" & UsuarioServidor.Usuario.AcessoEnderecoEmpresa))
        If DdlEmpresaOrigem.SelectedIndex > 0 Then CargaEmpresaDestino()
    End Sub
    Private Sub CargaEmpresaDestino()
        CarregarEmpresas(ddlEmpresaDest, "", DdlEmpresaOrigem.SelectedValue, False)
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

    Private Sub IncrementaCodigo()
        Sql = "SELECT ISNULL(Max(Codigo_Id), 0 ) + 1 as Codigo" & vbCrLf & _
              "     FROM Ativos" & vbCrLf & _
              "     WHERE Empresa_Id = " & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & vbCrLf & _
              "         And Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows
            txtCodigo.Text = Format(Dr("Codigo"), "0000")
        Next
    End Sub

    Private Sub IncrementaSequencia()
        Sql = "SELECT isnull(Max(Sequencia_Id), 0) + 1 as Sequencia" & vbCrLf & _
              "     FROM Ativos" & vbCrLf & _
              "     WHERE Empresa_Id = " & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & vbCrLf & _
              "         And Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
              "         And Codigo_Id = " & txtCodigo.Text & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Ativos").Tables(0).Rows
            txtSequencia.Text = Format(Dr("sequencia"), "000")
        Next
    End Sub

    Private Function Validar() As Boolean
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
        ElseIf Trim(txtDescricao.Text).Length > 500 Then
            MsgBox(Me.Page, "Descrição não pode ter mais que 500 caracteres")
            Return False
        ElseIf Trim(txtHistorico.Text).Length > 500 Then
            MsgBox(Me.Page, "Histórico não pode ter mais que 500 caracteres")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlUnidade.SelectedValue) Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlEmpresaOrigem.SelectedValue) Then
            MsgBox(Me.Page, "Empresa a qual o bem pertence é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEmpresaDest.SelectedValue) Then
            MsgBox(Me.Page, "Empresa a qual o bem esta localizado é obrigatório")
            Return False
        ElseIf Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) <> Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) Then
            MsgBox(Me.Page, "Empresa onde o bem esta localizado, não é uma empresa filial da empresa a qual o bem pertence.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlContaContabil.SelectedValue) Then
            MsgBox(Me.Page, "Conta é obrigatório")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlCentroDeCusto.SelectedValue) Then
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
        End If
        Return True
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

        ddlEmpresaDest.Items.Clear()

        ddlSequencia.Enabled = True

        VerificaUnidade()

        If DdlGrupo.Items.Count > 0 Then DdlGrupo.SelectedIndex = 0

        ddlContaContabil.Items.Clear()
        ddlCentroDeCusto.SelectedIndex = 0

        txtValorDepreciadoInicial.Text = String.Empty

        txtAquisicao.Text = Today.ToString("dd/MM/yyyy")
        txtInicioDeUso.Text = Today.ToString("dd/MM/yyyy")

        txtAtualizacao.Text = String.Empty
        txtAtualizacao.Enabled = True

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        GridAtivos.DataBind()

        If DdlEmpresaOrigem.SelectedIndex > 0 Then CargaEmpresaDestino()

        TabContainer1.ActiveTabIndex = 0
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
        Sql = "SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "       isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "       isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              "  from Usuarios" & vbCrLf & _
              " where Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            ddl.Carregar(DdlEmpresaOrigem, CarregarDDL.Tabela.Empresas, Dr("AcessoUnidade"), True)
            DdlEmpresaOrigem.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next

        ddlUnidade.Enabled = False
        DdlEmpresaOrigem.Enabled = False
    End Sub

    Private Function ExisteMovimentos() As Boolean
        Dim sql As String = "  Select top 1 Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id From AtivosXMovimentos" & vbCrLf & _
                            "     WHERE Processo     = 'Normal'" & vbCrLf & _
                            "       And Empresa_Id   = '" & Left(ddlEmpresaDest.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                            "       AND Grupo_Id     = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
                            "       AND Codigo_Id    =  " & txtCodigo.Text & vbCrLf & _
                            "       AND Sequencia_Id =  " & txtSequencia.Text & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    Private Sub CargaAtivos()
        Sql = " SELECT Grupo_Id as Grupo, Codigo_Id as Codigo, Sequencia_Id as Sequencia, Descricao, DataAquisicao as Aquisicao, ValorOriginal as Valor, " & vbCrLf & _
              "        case when DataDaBaixa is not null then 'Sim' else 'Não' end as baixado" & vbCrLf & _
              "   FROM Ativos" & vbCrLf & _
              "  WHERE Situacao = 1 " & vbCrLf & _
              "    and Empresa_Id = '" & Left(DdlEmpresaOrigem.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresaDest.SelectedValue) Then
            Sql &= " And Empresa = '" & ddlEmpresaDest.SelectedValue.Split("-")(0) & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
            Sql &= "    And Grupo_Id = '" & DdlGrupo.SelectedValue & "'" & vbCrLf
        End If

        Sql &= " Order by Grupo_Id, Codigo_Id, Sequencia_Id"

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Contas")
        GridAtivos.DataSource = ds
        GridAtivos.DataBind()
        If ds.Tables(0).Rows.Count > 0 AndAlso Not chkReaproveitarDados.Checked Then TabContainer1.ActiveTabIndex = 1
    End Sub

    Private Sub CarregarDataDaUltimaAtualizacao(ByVal EmpresaOrigem As String, ByVal Seguro As Boolean)
        Dim sql As String = String.Empty
        sql &= "SELECT ISNULL(MAX(am.Movimento_Id),dateadd(day, - day(max(a.Iniciodeuso)), max(a.Iniciodeuso))) UltimaAtualizacao" & vbCrLf & _
               "  FROM Ativos a " & vbCrLf & _
               "  LEFT JOIN AtivosXMovimentos am " & vbCrLf & _
               "    ON a.Sequencia_Id = am.Sequencia_Id" & vbCrLf & _
               "   AND a.Codigo_Id    = am.Codigo_Id" & vbCrLf & _
               "   AND a.Grupo_Id     = am.Grupo_Id" & vbCrLf & _
               "   And a.Empresa_Id   = am.Empresa_Id" & vbCrLf & _
               " WHERE am.Empresa_Id = '" & Left(EmpresaOrigem.Split("-")(0), 8) & "'" & vbCrLf & _
               "   AND am.Processo = 'NORMAL'" & vbCrLf & _
               "   And a.Seguro    = '" & IIf(Seguro, "TRUE", "FALSE") & "'" & vbCrLf & _
               "   AND isnull(a.Depreciar, 'N') = 'S'" & vbCrLf & _
               "   AND a.DataDaBaixa is null " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

        txtAtualizacao.Enabled = False
        If ds IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(ds.Tables(0).Rows(0)(0).ToString()) Then
            txtAtualizacao.Text = ds.Tables(0).Rows(0)(0)
        Else
            txtAtualizacao.Enabled = True
        End If
    End Sub

#End Region

End Class