Imports NGS.Lib.Negocio

Public Class ExtratoDeBens
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Patrimonio)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ExtratoDeBens", "ACESSAR") Then
                    ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas)
                    ddl.Carregar(ddlGrupoDeAtivo, CarregarDDL.Tabela.GruposDeAtivos)
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    'txtData1.Text = String.Format("01/{0}/{1}", Now.Month.ToString("00"), Now.Year)
                    'txtData2.Text = Now.ToString("dd/MM/yyyy")
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página!", "~/Patrimonio.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If validaCampos() Then
                Dim ds As DataSet = getDataSet()
                gridAtivos.DataSource = ds
                gridAtivos.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If validaCampos() Then
                Dim ds As DataSet = getDataSet()
                Dim param As New Dictionary(Of String, Object)
                param.Add("ParametroConsulta", "")
                Funcoes.BindReport(Me.Page, ds, "Cr_ExtratoDeBens", eExportType.PDF, param)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub imgExtratoDeBens_Click(sender As Object, e As ImageClickEventArgs)
        Try
            If validaCampos() Then
                Dim lnkConsultar As ImageButton = CType(sender, ImageButton)
                Dim row As GridViewRow = CType(lnkConsultar.NamingContainer, GridViewRow)
                Dim ds As DataSet = getDataSet(row.Cells(1).Text.Split(" - ")(0))

                Dim param As New Dictionary(Of String, Object)
                param.Add("ParametroConsulta", "")
                Funcoes.BindReport(Me.Page, ds, "Cr_ExtratoDeBens", eExportType.PDF, param)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            LimparCampos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ExtratoDeBens")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Function validaCampos()
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlGrupoDeAtivo.SelectedValue) Then
            MsgBox(Me.Page, "informe o grupo.")
            Return False
        End If
        Return True
    End Function

    Private Function getDataSet(Optional codigo As Integer = 0) As DataSet
        Dim sql As String = "select cast(a.Empresa_Id as integer) as Empresa_Id," & vbCrLf &
                            "       (select top 1 dbo.FormatarCliente(Cliente_Id, Endereco, Nome, Cidade, Estado)" & vbCrLf &
                            "          from Clientes" & vbCrLf &
                            "		  Where left(Cliente_Id, 8) = a.Empresa_Id" & vbCrLf &
                            "	    ) as DescricaoEmpresa," & vbCrLf &
                            "	    a.Grupo_Id, ga.Descricao as DescricaoGrupo, a.Codigo_Id, a.Sequencia_Id, cast(a.Codigo_Id as varchar) + ' - ' + a.Descricao as DescricaoBem," & vbCrLf &
                            "	    a.Empresa, a.EndEmpresa, dbo.FormatarCliente(a.Empresa, a.EndEmpresa, emp.Nome, emp.Cidade, emp.Estado) as EmpresaLocacao," & vbCrLf &
                            "	    a.Conta, pc.Titulo as DescricaoConta, a.CentroDeCusto, cc.CentroDeCusto_Id + ' - ' + cc.Descricao as DescricaoCentroDeCusto," & vbCrLf &
                            "	    a.Historico, a.DataAquisicao, a.InicioDeUso, ga.PercentualDepreciacao, case when a.Seguro = 'True' then 'Sim' Else 'Não' End as Seguro," & vbCrLf &
                            "	    a.ValorOriginal, axm.ValorDepreciado," & vbCrLf &
                            "	    (select Sum(Valor)" & vbCrLf &
                            "	       From AtivosXMovimentos" & vbCrLf &
                            "		  where Empresa_Id = a.Empresa_Id" & vbCrLf &
                            "		    and Grupo_Id   = a.Grupo_Id " & vbCrLf &
                            "		    and Codigo_Id  = a.Codigo_Id" & vbCrLf &
                            "		    and Processo   = 'INICIAL'" & vbCrLf &
                            "	    ) ValorDepreciadoInicial," & vbCrLf &
                            "       case when a.DataDaBaixa is null then 'Não' else 'Sim' end as Baixado" & vbCrLf &
                            "	    into #Ativo" & vbCrLf &
                            "  from Ativos a" & vbCrLf &
                            " Inner Join (Select Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id," & vbCrLf &
                            "                    Sum(Valor) as ValorDepreciado" & vbCrLf &
                            "  			    From AtivosXMovimentos" & vbCrLf &
                            "			   group by Empresa_Id, Grupo_Id, Codigo_Id, Sequencia_Id" & vbCrLf &
                            "			 ) as axm" & vbCrLf &
                            "	 on axm.Empresa_Id = a.Empresa_Id" & vbCrLf &
                            "   and axm.Grupo_Id   = a.Grupo_Id" & vbCrLf &
                            "   and axm.Codigo_Id  = a.Codigo_Id" & vbCrLf &
                            " Inner Join Clientes emp" & vbCrLf &
                            "    on emp.Cliente_Id  = a.Empresa" & vbCrLf &
                            "   and emp.Endereco_Id = a.EndEmpresa" & vbCrLf &
                            " Inner Join GruposDeAtivos ga" & vbCrLf &
                            "    on ga.Grupo_Id = a.Grupo_Id" & vbCrLf &
                            " Inner Join AtivosXContas axc" & vbCrLf &
                            "    on axc.Conta_Id = a.Conta" & vbCrLf &
                            " Inner Join PlanodeContas pc" & vbCrLf &
                            "    on pc.Conta_Id = axc.Conta_Id" & vbCrLf &
                            "  Left Join CentrosDeCustos cc" & vbCrLf &
                            "    on cc.CentroDeCusto_Id = a.CentroDeCusto" & vbCrLf &
                            " where a.Empresa_Id = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf &
                            "   and a.Grupo_Id   = '" & ddlGrupoDeAtivo.SelectedValue & "'" & vbCrLf

        If codigo > 0 Then sql &= "   and a.Codigo_Id  = " & codigo & vbCrLf

        sql &= "   and a.Situacao   = 1" & vbCrLf &
               "" & vbCrLf &
               "select * from #Ativo" & vbCrLf &
               "" & vbCrLf &
               "select cast(axm.Empresa_Id as integer) as Empresa_Id, axm.Grupo_Id, axm.Codigo_Id, axm.Sequencia_Id, axm.Movimento_Id, axm.Valor as Depreciado," & vbCrLf &
               "       case when axm.Processo = 'INICIAL' then 'Saldo Inicial' else 'Depreciação Mensal' end Historico," & vbCrLf &
               "       (select Sum(valor)" & vbCrLf &
               "                        from AtivosXMovimentos" & vbCrLf &
               "                        where Empresa_Id = axm.Empresa_Id" & vbCrLf &
               "		   and Grupo_Id      = axm.grupo_Id" & vbCrLf &
               "		   and Codigo_Id     = axm.Codigo_Id" & vbCrLf &
               "		   and Sequencia_Id  = axm.Sequencia_Id" & vbCrLf &
               "		   and Movimento_Id <= axm.Movimento_Id" & vbCrLf &
               "	   ) as Acumulado," & vbCrLf &
               "	   a.ValorOriginal -" & vbCrLf &
               "	   (select Sum(valor)" & vbCrLf &
               "                        from AtivosXMovimentos" & vbCrLf &
               "                        where Empresa_Id = axm.Empresa_Id" & vbCrLf &
               "		   and Grupo_Id      = axm.grupo_Id" & vbCrLf &
               "		   and Codigo_Id     = axm.Codigo_Id" & vbCrLf &
               "		   and Sequencia_Id  = axm.Sequencia_Id" & vbCrLf &
               "		   and Movimento_Id <= axm.Movimento_Id" & vbCrLf &
               "	   ) as Residual" & vbCrLf &
               "  from #Ativo a" & vbCrLf &
               "  Inner Join AtivosXMovimentos axm" & vbCrLf &
               "    on axm.Empresa_Id = a.Empresa_Id" & vbCrLf &
               "   and axm.Grupo_Id   = a.Grupo_Id" & vbCrLf &
               "   and axm.Codigo_Id  = a.Codigo_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Ativo")
        ds.Tables("Ativo1").TableName = "AtivosXMovimento"

        Return ds
    End Function

    Private Sub LimparCampos()
        Funcoes.VerificaEmpresa(ddlEmpresa)
        ddlGrupoDeAtivo.SelectedValue = ""
        'txtData1.Text = String.Format("01/{0}/{1}", Now.Month.ToString("00"), Now.Year)
        'txtData2.Text = Now.ToString("dd/MM/yyyy")
        'chkPeriodo.Checked = False
        'txtData1.Parent.Visible = False
        gridAtivos.DataBind()

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

End Class