Imports System.IO
Imports System.Data
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class SinteseDaDepreciacao
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("SinteseDaDepreciacao", "ACESSAR") Then
                    ddl.Carregar(DdlUnidade, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    VerificaUnidade()
                    UltimaAtualizacao()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "Patrimonio.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub VerificaUnidade()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidade, DdlEmpresa, True)
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, DdlUnidade.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub UltimaAtualizacao()
        Try
            Dim sql As String = "  SELECT top 1 am.Movimento_Id as Data, dateadd(day,-1,(dateadd(month,1,convert(char(08),DATEADD(MONTH, 1, am.Movimento_Id),126)+'01'))) AS Nova" & vbCrLf & _
                                "         From AtivosXMovimentos am" & vbCrLf & _
                                "           Inner Join Ativos a                        " & vbCrLf & _
                                "           	ON a.Sequencia_Id  = am.Sequencia_Id    " & vbCrLf & _
                                "           	AND a.Codigo_Id    = am.Codigo_Id       " & vbCrLf & _
                                "           	AND a.Grupo_Id     = am.Grupo_Id        " & vbCrLf & _
                                "           	And a.Empresa_Id   = am.Empresa_Id      " & vbCrLf & _
                                "       WHERE am.Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                "         AND am.Processo = 'NORMAL'" & vbCrLf & _
                                "         And a.Seguro = '" & IIf(chkSeguro.Checked, "TRUE", "FALSE") & "'" & vbCrLf & _
                                "         and isnull(a.Depreciar, 'N') = 'S'" & vbCrLf & _
                                "           and a.DataDaBaixa is null " & vbCrLf & _
                                "     ORDER BY am.Movimento_Id DESC" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "AtivosXMovimentos")

            If ds IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each Dr As DataRow In ds.Tables(0).Rows
                    txtDataFinal.Text = Format(Dr("Data"), "dd/MM/yyyy")
                Next
            Else
                txtDataFinal.Text = String.Empty
            End If

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            UltimaAtualizacao()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("SinteseDaDepreciacao", "RELATORIO") Then
                If txtDataFinal.Text = "" Then
                    MsgBox(Me.Page, "Data da última atualização é obrigatária.")
                Else

                    Dim DataFinal As String = Format(CDate(txtDataFinal.Text), "yyyy/MM/dd")
                    Dim PDataFinal As String = Format(CDate(txtDataFinal.Text), "dd/MM/yy")
                    Dim Mes As Integer = Month(DataFinal)
                    Dim Ano As Integer = Year(DataFinal)

                    Dim Sql As String = "" & vbCrLf & _
                          "   SELECT Ativos.Empresa, Ativos.EndEmpresa AS Endereco, Ativos.CentroDeCusto AS Custo, CentrosDeCustos.Descricao AS NomeCusto, Ativos.Conta, " & vbCrLf & _
                          "          Ativos.Grupo_Id AS Grupo, GruposDeAtivos.Descricao, SUM(AtivosXMovimentos.Valor) AS Valor, Clientes.Nome, Clientes.Cidade, Clientes.Estado, " & vbCrLf & _
                          "          Isnull(AtivosXContas.DepreciacaoDebito, '') AS Debito, Isnull(AtivosXContas.DepreciacaoCredito, '') AS Credito" & vbCrLf & _
                          "     FROM Ativos " & vbCrLf & _
                          "    INNER JOIN AtivosXMovimentos " & vbCrLf & _
                          "       ON Ativos.Empresa_Id   = AtivosXMovimentos.Empresa_Id " & vbCrLf & _
                          "      And Ativos.Grupo_Id     = AtivosXMovimentos.Grupo_Id " & vbCrLf & _
                          "      And Ativos.Codigo_Id    = AtivosXMovimentos.Codigo_Id " & vbCrLf & _
                          "      AND Ativos.Sequencia_Id = AtivosXMovimentos.Sequencia_Id " & vbCrLf & _
                          "    INNER JOIN GruposDeAtivos " & vbCrLf & _
                          "       ON Ativos.Grupo_Id     = GruposDeAtivos.Grupo_Id " & vbCrLf & _
                          "    INNER JOIN CentrosDeCustos " & vbCrLf & _
                          "       ON Ativos.CentroDeCusto = CentrosDeCustos.CentroDeCusto_Id " & vbCrLf & _
                          "    INNER JOIN Clientes " & vbCrLf & _
                          "       ON Ativos.Empresa    = Clientes.Cliente_Id " & vbCrLf & _
                          "      AND Ativos.EndEmpresa = Clientes.Endereco_Id " & vbCrLf & _
                          "     LEFT JOIN AtivosXContas  " & vbCrLf & _
                          "       ON Ativos.Empresa    = AtivosXContas.Empresa_Id " & vbCrLf & _
                          "      AND Ativos.EndEmpresa = AtivosXContas.EndEmpresa_Id " & vbCrLf & _
                          "      AND Ativos.Conta      = AtivosXContas.Conta_Id " & vbCrLf & _
                          "    WHERE Ativos.Empresa_Id = '" & Left(DdlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                          "      AND Ativos.Situacao   = 1 " & vbCrLf & _
                          "      AND AtivosXMovimentos.Valor <> 0 " & vbCrLf & _
                          "      AND MONTH(AtivosXMovimentos.Movimento_Id) = " & Mes & vbCrLf & _
                          "      AND Year(AtivosXMovimentos.Movimento_Id)  = " & Ano & vbCrLf & _
                          "      AND AtivosXMovimentos.Processo = 'NORMAL'" & vbCrLf

                    If chkSeguro.Checked Then
                        Sql &= " AND Ativos.Seguro = 'TRUE' "
                    Else
                        Sql &= " AND Ativos.Seguro = 'FALSE' "
                    End If
                    Sql &= " GROUP BY Ativos.Empresa, Ativos.EndEmpresa, Ativos.CentroDeCusto, Ativos.Grupo_Id, GruposDeAtivos.Descricao," & vbCrLf & _
                           "          CentrosDeCustos.Descricao, Ativos.Conta, Clientes.Nome, Clientes.Cidade," & vbCrLf & _
                           "          Clientes.Estado, AtivosXContas.DepreciacaoDebito, AtivosXContas.DepreciacaoCredito" & vbCrLf

                    Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "SinteseDaDepreciacao")

                    If ds.Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Nenhum registro encontrado para esta data.")
                    Else
                        Dim parameters = New Dictionary(Of String, Object)()
                        parameters.Add("Periodo", "Referente Mês " & Mes.ToString() & "/" & Ano)

                        Funcoes.BindReport(Me.Page, ds, "Cr_SinteseDaDepreciacao", eExportType.PDF, parameters)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "SinteseDaDepreciacao")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class