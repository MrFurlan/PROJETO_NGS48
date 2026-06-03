Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class AtivosXProcessos
    Inherits BasePage

    Dim SqlArray As New ArrayList

#Region "Events"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Patrimonio)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("AtivosXProcessos", "ACESSAR") Then
                CarregarUnidade()
                VerificaUnidade()
                UltimaAtualizacao()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Patrimonio.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            CarregarEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            UltimaAtualizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkProcessar_Click(sender As Object, e As EventArgs) Handles lnkProcessar.Click
        Try
            If ValidaCampos() Then
                Dim sql As String = " DELETE From AtivosXMovimentos" & vbCrLf & _
                                    "   FROM Ativos" & vbCrLf & _
                                    "  inner join AtivosXMovimentos" & vbCrLf & _
                                    "     ON Ativos.Empresa_Id   = AtivosXMovimentos.Empresa_Id " & vbCrLf & _
                                    "    AND Ativos.Grupo_Id     = AtivosXMovimentos.Grupo_Id " & vbCrLf & _
                                    "    AND Ativos.Codigo_Id    = AtivosXMovimentos.Codigo_Id " & vbCrLf & _
                                    "    AND Ativos.Sequencia_Id = AtivosXMovimentos.Sequencia_Id " & vbCrLf & _
                                    "  WHERE Left(Ativos.Empresa, 8) = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                    "    AND AtivosXMovimentos.Processo            = 'NORMAL' " & vbCrLf & _
                                    "    And Ativos.Situacao = 1" & vbCrLf & _
                                    "    And Isnull(Ativos.Depreciar, 'N')  = 'S'" & vbCrLf & _
                                    "    AND MONTH(AtivosXMovimentos.Movimento_Id) = " & Month(CDate(txtDataProxAtualizacao.Text)) & vbCrLf & _
                                    "    AND Year(AtivosXMovimentos.Movimento_Id)  = " & Year(CDate(txtDataProxAtualizacao.Text)) & vbCrLf & _
                                    "    And (Ativos.DataDaBaixa is null or Ativos.DataDaBaixa > '" & txtDataProxAtualizacao.Text.ToSqlDate() & "')" & vbCrLf

                If chkSeguro.Checked Then
                    sql &= " AND Ativos.Seguro = 'TRUE'" & vbCrLf
                Else
                    sql &= " AND Ativos.Seguro = 'FALSE'" & vbCrLf
                End If
                SqlArray.Add(sql)

                Dim ds As DataSet = getDataSetAtivos()
                Dim dias As Integer
                Dim Saldo As Decimal
                Dim Calculo As Decimal

                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        Dim diaInicioDeUsoBase As Integer = IIf(CDate(row("InicioDeUso")).Day = 31, 30, CDate(row("InicioDeUso")).Day)

                        If CDate(row("InicioDeUso")).Month = CDate(txtDataProxAtualizacao.Text).Month And CDate(row("InicioDeUso")).Year = CDate(txtDataProxAtualizacao.Text).Year Then 'Mes parcial
                            'dias = (30 - diaInicioDeUsoBase) + 1
                            dias = (30 - diaInicioDeUsoBase)
                            Calculo = Math.Round((((row("Valor") * row("Indice")) / 100) / 360) * dias, 2, MidpointRounding.AwayFromZero)
                        Else 'Mes completo
                            Calculo = Math.Round((((row("Valor") * row("Indice")) / 100) / 360) * 30, 2, MidpointRounding.AwayFromZero)
                        End If

                        '****************************************************************************************************************************

                        Saldo = row("Depreciado") + Calculo

                        If Saldo > row("Valor") Then
                            Calculo = row("Valor") - row("Depreciado")
                        End If

                        If Calculo > 0 Then
                            sql = "" & vbCrLf & _
                                  " INSERT INTO AtivosXMovimentos (Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id, Movimento_Id, Empresa, EndEmpresa, Quotas, Indice, Valor, Processo)" & vbCrLf & _
                                  "  VALUES ('" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "', " & "'" & row("Grupo") & "'," & row("Codigo") & ", " & row("Sequencia") & ", " & vbCrLf & _
                                  "          '" & txtDataProxAtualizacao.Text.ToSqlDate() & "', '" & row("Empresa") & "', " & row("EndEmpresa") & ", "
                            sql &= Str(CDec(Calculo)) & ", " & "1, " & Str(CDec(Calculo)) & ", 'NORMAL')" & vbCrLf
                            SqlArray.Add(sql)

                            sql = "Update Ativos Set Atualizacao = '" & txtDataProxAtualizacao.Text.ToSqlDate() & "'" & vbCrLf & _
                                  "     Where Empresa_Id  = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                  "         And Grupo_Id  = '" & row("Grupo") & "'" & vbCrLf & _
                                  "         And Codigo_Id = " & row("Codigo") & vbCrLf & _
                                  "         And Sequencia_Id = " & row("Sequencia") & vbCrLf

                            SqlArray.Add(sql)
                        End If
                    Next

                    If Banco.GravaBanco(SqlArray) Then
                        Contabiliza()
                    End If
                Else
                    MsgBox(Me.Page, "Não foi encontrado registros para o processo.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkEliminar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
                MsgBox(Me.Page, "Informe a empresa.")
            ElseIf String.IsNullOrWhiteSpace(txtDataUltAtualizacao.Text) Then
                MsgBox(Me.Page, "Informe a data da Ultima Atualização para eliminar.")
            ElseIf Not IsDate(txtDataUltAtualizacao.Text) Then
                MsgBox(Me.Page, "Data de Ultima Atualização não é uma data válida.")

            ElseIf Not Funcoes.VerificaAcessoMensal(DdlEmpresa.SelectedValue.Split("-")(0), DdlEmpresa.SelectedValue.Split("-")(1), txtDataUltAtualizacao.Text, "CONTABIL") Then
                MsgBox(Me.Page, "Movimento Contábil já fechado para esta data!")

            Else
                If VerificaExistenciaAtivosXMovimento(CDate(txtDataUltAtualizacao.Text)) Then
                    Dim sql As String = " DELETE From AtivosXMovimentos" & vbCrLf & _
                                        "   FROM Ativos" & vbCrLf & _
                                        "  inner join AtivosXMovimentos" & vbCrLf & _
                                        "     ON Ativos.Empresa_Id   = AtivosXMovimentos.Empresa_Id " & vbCrLf & _
                                        "    AND Ativos.Grupo_Id     = AtivosXMovimentos.Grupo_Id " & vbCrLf & _
                                        "    AND Ativos.Codigo_Id    = AtivosXMovimentos.Codigo_Id " & vbCrLf & _
                                        "    AND Ativos.Sequencia_Id = AtivosXMovimentos.Sequencia_Id " & vbCrLf & _
                                        "  WHERE Left(Ativos.Empresa, 8) = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                        "    AND AtivosXMovimentos.Processo            = 'NORMAL' " & vbCrLf & _
                                        "    And Ativos.Situacao = 1" & vbCrLf & _
                                        "    And Isnull(Ativos.Depreciar, 'N')  = 'S'" & vbCrLf & _
                                        "    AND MONTH(AtivosXMovimentos.Movimento_Id) = " & Month(CDate(txtDataUltAtualizacao.Text)) & vbCrLf & _
                                        "    AND Year(AtivosXMovimentos.Movimento_Id)  = " & Year(CDate(txtDataUltAtualizacao.Text)) & vbCrLf & _
                                        "    And (Ativos.DataDaBaixa is null or Ativos.DataDaBaixa > '" & txtDataUltAtualizacao.Text.ToSqlDate() & "')" & vbCrLf

                    If chkSeguro.Checked Then
                        sql &= " AND Ativos.Seguro = 'TRUE'" & vbCrLf
                    Else
                        sql &= " AND Ativos.Seguro = 'FALSE'" & vbCrLf
                    End If
                    SqlArray.Add(sql)

                    sql = " DELETE From Razao" & vbCrLf & _
                          "     WHERE Empresa_Id Like '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "%'" & vbCrLf & _
                          "         AND MONTH(Movimento_Id) = " & Month(CDate(txtDataUltAtualizacao.Text)) & vbCrLf & _
                          "         AND Year(Movimento_Id)  = " & Year(CDate(txtDataUltAtualizacao.Text)) & vbCrLf

                    If chkSeguro.Checked Then
                        sql &= " AND Lote_Id = 51 " & vbCrLf
                    Else
                        sql &= " AND Lote_Id = 50 " & vbCrLf
                    End If
                    SqlArray.Add(sql)

                    Dim ds As DataSet = getDataSetAtivos()

                    For Each row As DataRow In ds.Tables(0).Rows
                        sql = "Update Ativos Set Atualizacao = dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, dateadd(month, -1,'" & txtDataUltAtualizacao.Text.ToSqlDate() & "')), 126)+'01')))" & vbCrLf & _
                              "     Where   Empresa_Id     = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                              "         And Grupo_Id     = '" & row("Grupo") & "'" & vbCrLf & _
                              "         And Codigo_Id    = " & row("Codigo") & vbCrLf & _
                              "         And Sequencia_Id = " & row("Sequencia") & vbCrLf
                        SqlArray.Add(sql)
                    Next

                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Sucesso na Eliminação.")
                        UltimaAtualizacao()
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

#End Region

#Region "Methods"

    Private Sub CarregarUnidade()
        Try
            Dim Sql As String = ""
            Sql = "SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
                  "       isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
                  "       isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
                  "  from Usuarios" & vbCrLf & _
                  " where Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

            ddl.Carregar(ddlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
                ddlUnidade.SelectedValue = Dr("AcessoUnidade")
                ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, Dr("AcessoUnidade"), True)
                DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            Next

            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub CarregarEmpresa()
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidade.SelectedValue, True)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub VerificaUnidade()
        Try
            Funcoes.VerificaUnidadeDeNegocio(ddlUnidade, DdlEmpresa)
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub UltimaAtualizacao()
        Try
            Dim sql As String = "SELECT " & vbCrLf & _
                                "        case" & vbCrLf & _
                                "           when AtivosXMovimentos.Movimento_Id < dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, AtivosXMovimentos.Movimento_Id), 126)+'01'))) " & vbCrLf & _
                                "           then AtivosXMovimentos.Movimento_Id - DAY(AtivosXMovimentos.Movimento_Id) + 1 " & vbCrLf & _
                                "           else Isnull(AtivosXMovimentos.Movimento_Id,  " & vbCrLf & _
                                " 	        ((select min(cons.IniciodeUso) as Dt from ativos cons " & vbCrLf & _
                                "               WHERE cons.Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                " 		        And cons.Seguro    = 'FALSE' " & vbCrLf & _
                                "               and isnull(cons.Depreciar, 'N') = 'S' " & vbCrLf & _
                                "               and cons.DataDaBaixa is null)) " & vbCrLf & _
                                "               ) end as Data " & vbCrLf & _
                                " Into #Temp " & vbCrLf & _
                                "       FROM      Ativos INNER JOIN " & vbCrLf & _
                                "                 AtivosXMovimentos ON Ativos.Empresa_Id = AtivosXMovimentos.Empresa_Id AND Ativos.Grupo_Id = AtivosXMovimentos.Grupo_Id AND Ativos.Codigo_Id = AtivosXMovimentos.Codigo_Id AND  " & vbCrLf & _
                                "                 Ativos.Sequencia_Id = AtivosXMovimentos.Sequencia_Id And AtivosXMovimentos.Processo = 'NORMAL' " & vbCrLf & _
                                "       WHERE     Ativos.Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                " 		        AND (Ativos.Seguro = 'FALSE')  " & vbCrLf & _
                                " 		        AND (ISNULL(Ativos.Depreciar, 'N') = 'S')  " & vbCrLf & _
                                " 		        AND (Ativos.DataDaBaixa IS NULL) " & vbCrLf & _
                                " Select Top 1 *,  " & vbCrLf & _
                                "               Case when Data < dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, Data), 126)+'01')))  " & vbCrLf & _
                                "  	  	            then  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, Data),126)+'01'))) " & vbCrLf & _
                                "  	                else  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 1, Data),126)+'01'))) " & vbCrLf & _
                                "        end as Nova " & vbCrLf & _
                                " From #Temp " & vbCrLf & _
                                " order by data Desc"


            '--ALTERADO POR nERI EM 22/11/2021 CONFORME SQL ACIMA
            '"SELECT top 1" & vbCrLf & _
            '"       case" & vbCrLf & _
            '"         when am.Movimento_Id < dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, am.Movimento_Id), 126)+'01')))" & vbCrLf & _
            '"           then am.Movimento_Id - DAY(am.Movimento_Id) + 1" & vbCrLf & _
            '"  	        else am.Movimento_Id" & vbCrLf & _
            '"       end as Data," & vbCrLf & _
            '"       case" & vbCrLf & _
            '"         when am.Movimento_Id < dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, am.Movimento_Id), 126)+'01'))) " & vbCrLf & _
            '"  	  	    then  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, am.Movimento_Id),126)+'01')))" & vbCrLf & _
            '"  	        else  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 1, am.Movimento_Id),126)+'01')))" & vbCrLf & _
            '"       end as Nova" & vbCrLf & _
            '"  From AtivosXMovimentos am" & vbCrLf & _
            '" Inner Join Ativos a                        " & vbCrLf & _
            '"    ON a.Sequencia_Id = am.Sequencia_Id    " & vbCrLf & _
            '"   AND a.Codigo_Id    = am.Codigo_Id       " & vbCrLf & _
            '"   AND a.Grupo_Id     = am.Grupo_Id        " & vbCrLf & _
            '"   And a.Empresa_Id   = am.Empresa_Id      " & vbCrLf & _
            '" WHERE am.Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
            '"   AND am.Processo = 'NORMAL'" & vbCrLf & _
            '"   And a.Seguro    = '" & IIf(chkSeguro.Checked, "TRUE", "FALSE") & "'" & vbCrLf & _
            '"   and isnull(a.Depreciar, 'N') = 'S'" & vbCrLf & _
            '"   and a.DataDaBaixa is null " & vbCrLf & _
            '" ORDER BY am.Movimento_Id DESC" & vbCrLf


            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each Dr As DataRow In ds.Tables(0).Rows
                    txtDataUltAtualizacao.Text = Format(Dr("Data"), "dd/MM/yyyy")
                    txtDataProxAtualizacao.Text = Format(Dr("Nova"), "dd/MM/yyyy")
                    ViewState("UltAtualizacao") = Format(Dr("Data"), "dd/MM/yyyy")
                Next
            Else
                sql = "select min(Atualizacao) as Atualizacao" & vbCrLf & _
                        "into #Temp" & vbCrLf & _
                        "From Ativos" & vbCrLf & _
                        "WHERE Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                        "  AND (Seguro = 'FALSE')  " & vbCrLf & _
                        "  AND (ISNULL(Depreciar, 'N') = 'S')  " & vbCrLf & _
                        "" & vbCrLf & _
                        "select min(Atualizacao) as Data," & vbCrLf & _
                        "    case" & vbCrLf & _
                        "		when Atualizacao < dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, Atualizacao), 126)+'01')))" & vbCrLf & _
                        "			then  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 0, Atualizacao),126)+'01')))" & vbCrLf & _
                        "			else  dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 1, Atualizacao),126)+'01')))" & vbCrLf & _
                        "		end as Nova" & vbCrLf & _
                        "from #Temp" & vbCrLf & _
                        "group by Atualizacao"

                ds = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

                If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    For Each Dr As DataRow In ds.Tables(0).Rows
                        txtDataUltAtualizacao.Text = Format(Dr("Data"), "dd/MM/yyyy")
                        txtDataProxAtualizacao.Text = Format(Dr("Nova"), "dd/MM/yyyy")
                        ViewState("UltAtualizacao") = Format(Dr("Data"), "dd/MM/yyyy")
                    Next
                Else
                    txtDataUltAtualizacao.Text = String.Empty
                    txtDataProxAtualizacao.Text = String.Empty
                End If
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Sub Contabiliza()
        Try
            Dim sql As String = " DELETE Razao" & vbCrLf & _
                                "  WHERE Empresa_Id Like '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "%'" & vbCrLf & _
                                "    AND MONTH(Movimento_Id) = " & Month(CDate(txtDataProxAtualizacao.Text)) & vbCrLf & _
                                "    AND YEAR(Movimento_Id)  = " & Year(CDate(txtDataProxAtualizacao.Text)) & vbCrLf

            If chkSeguro.Checked Then
                sql &= " AND Lote_Id = 51 " & vbCrLf
            Else
                sql &= " AND Lote_Id = 50 " & vbCrLf
            End If
            SqlArray.Add(sql)

            sql = "  SELECT A.Empresa, " & vbCrLf & _
                  "         A.EndEmpresa AS Endereco," & vbCrLf & _
                  "         A.CentroDeCusto AS Custo, " & vbCrLf & _
                  "         Cc.Descricao AS NomeCusto," & vbCrLf & _
                  "         A.Conta, " & vbCrLf & _
                  "         A.Empresa_Id, " & vbCrLf & _
                  "         A.Grupo_Id, " & vbCrLf & _
                  "         A.Codigo_Id, " & vbCrLf & _
                  "         A.Sequencia_Id, " & vbCrLf & _
                  "         Ga.Descricao," & vbCrLf & _
                  "         Isnull(Ac.DepreciacaoDebito, '') as Debita, " & vbCrLf & _
                  "         Isnull(Ac.DepreciacaoCredito, '') as Credita," & vbCrLf & _
                  "         A.Historico as HistoricoAtivo, " & vbCrLf & _
                  "         A.Seguro, " & vbCrLf & _
                  "         SUM(Am.Valor) AS Valor" & vbCrLf & _
                  "    FROM Ativos A " & vbCrLf & _
                  "   INNER JOIN AtivosXMovimentos Am" & vbCrLf & _
                  "      ON A.Empresa_Id   = Am.Empresa_Id " & vbCrLf & _
                  "     AND A.Grupo_Id     = Am.Grupo_Id  " & vbCrLf & _
                  "     AND A.Codigo_Id    = Am.Codigo_Id " & vbCrLf & _
                  "     AND A.Sequencia_Id = Am.Sequencia_Id " & vbCrLf & _
                  "   INNER JOIN GruposDeAtivos Ga " & vbCrLf & _
                  "      ON A.Grupo_Id = Ga.Grupo_Id" & vbCrLf & _
                  "   INNER JOIN CentrosDeCustos Cc" & vbCrLf & _
                  "      ON A.CentroDeCusto = Cc.CentroDeCusto_Id" & vbCrLf & _
                  "    LEFT OUTER JOIN AtivosXContas Ac" & vbCrLf & _
                  "      ON A.Empresa    = Ac.Empresa_Id " & vbCrLf & _
                  "     AND A.EndEmpresa = Ac.EndEmpresa_Id" & vbCrLf & _
                  "     AND A.Conta      = Ac.Conta_Id" & vbCrLf & _
                  "   WHERE Left(A.Empresa, 8) = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "' " & vbCrLf & _
                  "     AND A.Situacao          = 1" & vbCrLf & _
                  "     AND Am.Valor <> 0 " & vbCrLf & _
                  "     AND Am.Movimento_Id = '" & txtDataProxAtualizacao.Text.ToSqlDate() & "'" & vbCrLf & _
                  "     AND ISNULL(am.Processo, '') IN ('INICIAL', 'NORMAL')" & vbCrLf & _
                  "     AND A.Seguro = '" & IIf(chkSeguro.Checked, "TRUE", "FALSE") & "'" & vbCrLf & _
                  "     AND (A.DataDaBaixa is null or A.DataDaBaixa > '" & txtDataProxAtualizacao.Text.ToSqlDate() & "')" & vbCrLf & _
                  "   GROUP BY A.Empresa, A.EndEmpresa, A.CentroDeCusto, A.Empresa_Id, " & vbCrLf & _
                  "         A.Grupo_Id, A.Codigo_Id, A.Sequencia_Id, " & vbCrLf & _
                  "         Ga.Descricao, Cc.Descricao, A.Conta, " & vbCrLf & _
                  "         Ac.DepreciacaoDebito, Ac.DepreciacaoCredito, A.seguro, A.Historico" & vbCrLf

            Dim EmpresaAnt As String = String.Empty
            Dim Sequencia As Integer
            Dim Ativo As String = String.Empty

            For Each row As DataRow In Banco.ConsultaDataSet(sql, "Ativos").Tables(0).Select("", "Empresa")
                If row("Empresa") <> EmpresaAnt Then
                    Sequencia = 0
                    EmpresaAnt = row("Empresa")
                End If

                Ativo = row("Grupo_Id") & "-" & row("Codigo_Id") & "-" & row("sequencia_Id")
                'Ativo = String.Empty

                If Not String.IsNullOrWhiteSpace(row("Debita")) Then
                    sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, " & vbCrLf & _
                          "                   Custo, Indexador, DataMoeda, DebitoOficial, CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, " & vbCrLf & _
                          "                   Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                          "           VALUES ('" & row("Empresa") & "', " & row("Endereco") & ", '" & row("Debita") & "', '', 0" & ", '" & txtDataProxAtualizacao.Text.ToSqlDate() & "', " & vbCrLf & _
                          "                    " & IIf(chkSeguro.Checked, 51, 50) & vbCrLf

                    Sequencia += 1
                    sql &= ", " & Sequencia & ", " & row("Custo") & ", 3" & ", '" & txtDataProxAtualizacao.Text.ToSqlDate() & "', " & Str(CDec(row("Valor"))) & ", 0, 0, 0" & vbCrLf

                    If chkSeguro.Checked Then
                        sql &= ", '" & row("HistoricoAtivo") & IIf(row("Seguro") = "TRUE", " / Apropriacao", " / Depreciacao") & " Mes " & Month(CDate(txtDataProxAtualizacao.Text)) & "/" & Year(CDate(txtDataProxAtualizacao.Text)) & "'" & vbCrLf
                    Else
                        sql &= ", '" & IIf(row("Seguro") = "TRUE", "Apropriacao", "Depreciacao") & " Mes " & Month(CDate(txtDataProxAtualizacao.Text)) & "/" & Year(CDate(txtDataProxAtualizacao.Text)) & " - " & row("Descricao") & "'" & vbCrLf
                    End If
                    sql &= ", 'P', 'AtivosXProcessos', '" & UsuarioServidor.NomeUsuario & "', getdate(), '" & Ativo & "', " & Sequencia & ")" & vbCrLf
                    SqlArray.Add(sql)
                End If

                If Not String.IsNullOrWhiteSpace(row("Credita")) Then
                    sql = "INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Custo, Indexador, DataMoeda, DebitoOficial, " & vbCrLf & _
                           "                  CreditoOficial, DebitoMoeda, CreditoMoeda, Historico, PrevistoRealizado, Processo, UsuarioInclusao, UsuarioInclusaoData, AtivoImobilizado, AgrupadorDeLancamento)" & vbCrLf & _
                           "         VALUES  ('" & row("Empresa") & "', " & row("Endereco") & ", '" & row("Credita") & "', '', 0, '" & txtDataProxAtualizacao.Text.ToSqlDate() & "', " & IIf(chkSeguro.Checked, 51, 50) & vbCrLf

                    Sequencia += 1
                    sql &= ", " & Sequencia & ", " & row("Custo") & ", 3, '" & txtDataProxAtualizacao.Text.ToSqlDate() & "', 0, " & Str(CDec(row("Valor"))) & ", 0, 0"

                    If chkSeguro.Checked Then
                        sql &= ", '" & row("HistoricoAtivo") & IIf(row("Seguro") = "TRUE", " / Apropriacao", " / Depreciacao") & " Mes " & Month(CDate(txtDataProxAtualizacao.Text)) & "/" & Year(CDate(txtDataProxAtualizacao.Text)) & "'"
                    Else
                        sql &= ", '" & IIf(row("Seguro") = "TRUE", "Apropriacao", "Depreciacao") & " Mes " & Month(CDate(txtDataProxAtualizacao.Text)) & "/" & Year(CDate(txtDataProxAtualizacao.Text)) & "'"
                    End If

                    sql &= ", 'P', 'AtivosXProcessos','" & UsuarioServidor.NomeUsuario & "', getdate(), '" & Ativo & "'," & Sequencia - 1 & ")"
                    SqlArray.Add(sql)
                End If
            Next

            If Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, "Processo realizado com Sucesso.", eTitulo.Sucess)
                UltimaAtualizacao()
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Private Function getDataSetAtivos() As DataSet
        Dim sql As String = "SELECT A.DataDaBaixa, A.Empresa, A.EndEmpresa, A.Grupo_Id AS Grupo," & vbCrLf & _
                            "       A.Codigo_Id AS Codigo," & vbCrLf & _
                            "       A.Sequencia_Id AS Sequencia," & vbCrLf & _
                            "       A.InicioDeUso," & vbCrLf & _
                            "       A.ValorOriginal as Valor, " & vbCrLf & _
                            "       ISNULL(SUM(AM.Valor), 0) AS Depreciado," & vbCrLf & _
                            "       GA.PercentualDepreciacao as Indice" & vbCrLf & _
                            "  FROM Ativos A" & vbCrLf & _
                            " INNER JOIN GruposDeAtivos GA " & vbCrLf & _
                            "    ON A.Grupo_Id = GA.Grupo_Id " & vbCrLf & _
                            "  LEFT OUTER JOIN AtivosXMovimentos AM " & vbCrLf & _
                            "    ON A.Sequencia_Id = AM.Sequencia_Id" & vbCrLf & _
                            "   AND A.Codigo_Id    = AM.Codigo_Id  " & vbCrLf & _
                            "   AND A.Grupo_Id     = AM.Grupo_Id" & vbCrLf & _
                            "   AND A.Empresa_Id   = AM.Empresa_Id" & vbCrLf & _
                            " WHERE Left(A.Empresa, 8) = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                            "   AND isnull(A.Depreciar,'N')  = 'S' " & vbCrLf & _
                            "   AND A.Situacao   = 1 " & vbCrLf & _
                            "   AND ISNULL(AM.Processo, '') IN ('INICIAL', 'NORMAL')" & vbCrLf & _
                            "   AND GA.PercentualDepreciacao <> 0 " & vbCrLf & _
                            "   AND A.DataAquisicao < '" & txtDataProxAtualizacao.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   AND A.InicioDeUso < '" & txtDataProxAtualizacao.Text.ToSqlDate() & "'" & vbCrLf & _
                            "   AND (A.DataDaBaixa is null or A.DataDaBaixa > '" & CDate(txtDataProxAtualizacao.Text).ToString("yyyy-MM-dd") & "')" & vbCrLf

        If chkSeguro.Checked Then
            sql &= " AND A.Seguro = 'TRUE'" & vbCrLf
        Else
            sql &= " AND A.Seguro = 'FALSE'" & vbCrLf
        End If

        sql &= "  GROUP BY  A.DataDaBaixa, A.Empresa, A.EndEmpresa, A.Grupo_Id, A.Codigo_Id, A.Sequencia_Id,  A.Empresa_Id, " & vbCrLf & _
                "           A.InicioDeUso, A.ValorOriginal, GA.PercentualDepreciacao"

        Return Banco.ConsultaDataSet(sql, "Ativo")
    End Function


    Private Function ValidaCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataUltAtualizacao.Text) OrElse String.IsNullOrWhiteSpace(txtDataProxAtualizacao.Text) Then
            MsgBox(Me.Page, "Informe a data da última  e próxima atualização.")
            Return False
        ElseIf Not IsDate(txtDataUltAtualizacao.Text) Then
            MsgBox(Me.Page, "Data da ultima atualização não é uma data válida.")
            Return False
        ElseIf Not IsDate(txtDataProxAtualizacao.Text) Then
            MsgBox(Me.Page, "Data da próxima atualização não é uma data válida.")
            Return False
            'Rever caso quando a ultima atualização é o ano anterior (Neri) 
            '###################################################################################################################
            'para correção foi adicionado 1 mes a data de ultima atualizacao, sendo assim as datas devem estar no mesmo ano.
        ElseIf CDate(txtDataUltAtualizacao.Text).AddMonths(1).Year <> CDate(txtDataProxAtualizacao.Text).Year Then
            MsgBox(Me.Page, "As data de proxima atualização não pode ultrapassar 1 mês a data de Ultima atualização.")
            Return False
        ElseIf CDate(txtDataProxAtualizacao.Text) < CDate(txtDataUltAtualizacao.Text) Then
            MsgBox(Me.Page, "Data da Próxima atualização não pode ser menor que a data da última atualização.")
            Return False
            '#################################################################################################################################
            'Não sei o por que que foi retirado esta validação, pois o mesmo é necessário para que não efetuem processos que ja foram rodadas, - descomentei.
        ElseIf ViewState("UltAtualizacao") IsNot Nothing AndAlso CDate(ViewState("UltAtualizacao")) < CDate(txtDataUltAtualizacao.Text) Then
            MsgBox(Me.Page, "Não é permitido processar o próximo mês.")
            Return False
            'Rever caso em que o ultimo mes é o mes 12 e o próximo é o mes 01 >> (Neri) 
            '#####################################################################################################################################
            'para correção foi adicionado 1 mes a data de ultima atualizacao, para que o mesmo seja verificado mes e ano.
        ElseIf (CDate(txtDataUltAtualizacao.Text).AddMonths(1).Month <> CDate(txtDataProxAtualizacao.Text).Month) OrElse (CDate(txtDataUltAtualizacao.Text).AddMonths(1).Year <> CDate(txtDataProxAtualizacao.Text).Year) Then
            MsgBox(Me.Page, "A data da próxima atualização deve conter exato(s) 1 mês de diferença em relação a data de última atualização.")
            Return False
        ElseIf DateTime.DaysInMonth(CDate(txtDataUltAtualizacao.Text).Year, CDate(txtDataUltAtualizacao.Text).Month) <> CDate(txtDataUltAtualizacao.Text).Day Then
            MsgBox(Me.Page, "O Dia informado da ultima atualização não é o último dia do mês.")
            Return False
        ElseIf DateTime.DaysInMonth(CDate(txtDataProxAtualizacao.Text).Year, CDate(txtDataProxAtualizacao.Text).Month) <> CDate(txtDataProxAtualizacao.Text).Day Then
            MsgBox(Me.Page, "O Dia informado da próxima atualização não é o último dia do mês.")
            Return False
            'Esta dando erro de data nula precisa rever essa função (Neri) 
            'Para empresas que ainda não executou processos, inativa elseif
        ElseIf getUltimoMesProcessado() IsNot Nothing AndAlso getUltimoMesProcessado() < CDate(txtDataUltAtualizacao.Text) Then
            MsgBox(Me.Page, "Atenção. existem meses anteriores a ser processado. Ultimo Mês Processado: " & IIf(IsDBNull(getUltimoMesProcessado()), "", MonthName(Month(getUltimoMesProcessado()))))
            Return False
        End If
        Return True
    End Function

    Private Function VerificaExistenciaAtivosXMovimento(ByVal Data As DateTime) As Boolean
        Dim sql As String = "  Select * From AtivosXMovimentos" & vbCrLf & _
                              "     WHERE   Empresa_Id Like '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "%' " & vbCrLf & _
                              "         AND Processo            = 'NORMAL'" & vbCrLf & _
                              "         AND MONTH(Movimento_Id) = " & Month(Data) & vbCrLf & _
                              "         AND Year(Movimento_Id)  = " & Year(Data) & vbCrLf
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

        If ds Is Nothing OrElse ds.Tables Is Nothing AndAlso ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Não encontrado dados para exclusão.")
            Return False
        End If
        Return True
    End Function

    Private Function getUltimoMesProcessado() As Date?
        Dim sql As String = "  Select max(Movimento_Id) as UltMovimento From AtivosXMovimentos" & vbCrLf & _
                            "       WHERE   Empresa_Id Like '" & Left(DdlEmpresa.SelectedValue, 8) & "' " & vbCrLf & _
                            "           AND Processo            = 'NORMAL'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            If Not IsDBNull(ds.Tables(0).Rows(0)(0)) Then
                Return ds.Tables(0).Rows(0)(0)
            Else
                Return Nothing
            End If
        End If
    End Function

#End Region

    Protected Sub chkSeguro_CheckedChanged(sender As Object, e As EventArgs) Handles chkSeguro.CheckedChanged
        Try
            UltimaAtualizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class