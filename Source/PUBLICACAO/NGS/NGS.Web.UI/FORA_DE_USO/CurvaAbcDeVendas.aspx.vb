Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CurvaAbcDeVendas
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim Mensagem As String
    Dim Cliente() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Comercial)
        If Not IsPostBack AndAlso IsConnect Then
            If Funcoes.VerificaPermissao("CurvaAbcDeVendas", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaMes()
                CargaAno()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Sub CargaMes()
        ddlMes.Items.Add(New ListItem("Janeiro", 1))
        ddlMes.Items.Add(New ListItem("Fevereiro", 2))
        ddlMes.Items.Add(New ListItem("Março", 3))
        ddlMes.Items.Add(New ListItem("Abril", 4))
        ddlMes.Items.Add(New ListItem("Maio", 5))
        ddlMes.Items.Add(New ListItem("Junho", 6))
        ddlMes.Items.Add(New ListItem("Julho", 7))
        ddlMes.Items.Add(New ListItem("Agosto", 8))
        ddlMes.Items.Add(New ListItem("Setembro", 9))
        ddlMes.Items.Add(New ListItem("Outubro", 10))
        ddlMes.Items.Add(New ListItem("Novembro", 11))
        ddlMes.Items.Add(New ListItem("Dezembro", 12))
    End Sub

    Sub CargaAno()
        ddlAno.Items.Add(New ListItem("2010", 2010))
        ddlAno.Items.Add(New ListItem("2011", 2011))
        ddlAno.Items.Add(New ListItem("2012", 2012))
        ddlAno.Items.Add(New ListItem("2013", 2013))
        ddlAno.Items.Add(New ListItem("2014", 2014))
        ddlAno.Items.Add(New ListItem("2015", 2015))
        ddlAno.Items.Add(New ListItem("2016", 2016))
        ddlAno.Items.Add(New ListItem("2017", 2017))
    End Sub

    Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & _
              "FROM Clientes C " & _
              "INNER JOIN ClientesXTipos CT " & _
              "ON C.Cliente_Id = CT.Cliente_Id " & _
              "WHERE CT.Tipo_Id = 050 " & _
              "ORDER BY Nome"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade,"
        Sql &= "        isnull(AcessoEmpresa, '') as AcessoEmpresa,"
        Sql &= "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa"
        Sql &= " from Usuarios"
        Sql &= " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado "
        Sql &= " FROM   GruposXEmpresas INNER JOIN"
        Sql &= " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id"
        Sql &= " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' "

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0

    End Sub

    Sub Limpar()

    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        CargaEmpresas()
    End Sub

    Function Validar()

        Mensagem = ""

        If DdlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório..."
            Return Mensagem
        End If

        If DdlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório..."
            Return Mensagem
        End If

        Return Mensagem

    End Function

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Validar()
        If Mensagem = "" Then
            Try
                Dim DataInicial As String = ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-01"


                Sql = " if(object_id('tempdb..#Temp') is not null) begin drop table #Temp end;"

                Sql &= " SELECT Clientes.Nome as cliente,  NotasFiscais.Cliente_Id,"
                Sql &= "        Count(NotasFiscais.Nota_id) as QuantidadeNF,"
                Sql &= " 	    Sum(NotasFiscaisXItens.Valor) as Valor,"
                Sql &= " 	    convert(money,0.00) as ValorTotal"
                Sql &= " into #temp"

                Sql &= " FROM    NotasFiscais    INNER JOIN   Clientes ON  NotasFiscais.Cliente_Id = Clientes.Cliente_Id"
                Sql &= " 	    And NotasFiscais.EndCliente_Id = Clientes.Endereco_Id"
                Sql &= " 	    INNER JOIN  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id"
                Sql &= " 	    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id"
                Sql &= " 	    AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id"
                Sql &= " 	    AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id"
                Sql &= " 	    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id"
                Sql &= "        AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id"
                Sql &= "        AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id"
                Sql &= " Where  (Year(NotasFiscais.Movimento) = Year(getdate()))"
                Sql &= "        And NotasFiscais.EntradaSaida_Id = 'S'"
                Sql &= "        And NotasFiscais.Situacao = 1"

                Sql &= " Group	By Clientes.Reduzido, NotasFiscais.Cliente_Id,  Clientes.Nome"
                Sql &= " Order	By Valor desc"

                Sql &= " update	#temp set ValorTotal =("
                Sql &= "        Select Case Convert(money, sum(Valor))"
                Sql &= " from	( SELECT Clientes.Nome as cliente,"
                Sql &= "        Sum(NotasFiscaisXItens.Valor) as Valor"
                Sql &= " FROM   NotasFiscais"
                Sql &= " 		INNER JOIN   Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id"
                Sql &= "        AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id"
                Sql &= "        INNER JOIN  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id"
                Sql &= " 	    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id"
                Sql &= " 	    AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id"
                Sql &= " 		AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id"
                Sql &= " 	    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id"
                Sql &= " 	    AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id"
                Sql &= " 	    AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id"
                Sql &= " Where (Year(NotasFiscais.Movimento) = Year(getdate()))"
                Sql &= " 	    And NotasFiscais.EntradaSaida_Id = 'S'"
                Sql &= " 	    And NotasFiscais.Situacao = 1"

                Sql &= " Group	By NotasFiscais.Cliente_Id, Clientes.Nome  ) as Consulta                    )"
                Sql &= " "
                Sql &= " update	#temp set cliente =  (Select cliente  from ("
                Sql &= " SELECT	top 1 Clientes.Nome as cliente"
                Sql &= "        from Clientes"
                Sql &= " Where	left(#temp.Cliente_Id,8) = left(Cliente_Id,8) ) as consulta )"
                Sql &= " select Cliente, "
                Sql &= " 		sum(quantidadeNF) as quantidadeNF,   "
                Sql &= " 		sum(Valor) as Valor,ValorTotal,  "
                Sql &= " 		sum(Valor * 100 / ValorTotal) AS ParcelaIndividual"
                Sql &= " from	#temp"
                Sql &= "        where(valor > 100)"
                Sql &= " group	by Cliente,ValorTotal  "
                Sql &= " order	by Valor desc "

                GridAbc.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
                GridAbc.DataBind()


            Catch ex As Exception
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "alert('Nemhum Registro Encontrado...');", True)
            End Try
        Else
            ScriptManager.RegisterClientScriptBlock(Me, Me.DdlUnidade.GetType(), Guid.NewGuid().ToString(), "alert('" & Mensagem & "');", True)
        End If
    End Sub

    Protected Sub LinkLimpar_Click(sender As Object, e As EventArgs) Handles LinkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Dim NomeArquivo As String = "Manual/PedidosEntreguesDeVendas.mht"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
    End Sub
End Class