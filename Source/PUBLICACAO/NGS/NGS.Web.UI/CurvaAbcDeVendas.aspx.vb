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
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CurvaAbcDeVendas", "ACESSAR") Then
                CargaUnidade()
                VerificaUnidade()
                CargaMes()
                CargaAno()
                Limpar()
                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaMes()
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

    Private Sub CargaAno()
        ddlAno.Items.Add(New ListItem("2013", 2013))
        ddlAno.Items.Add(New ListItem("2014", 2014))
        ddlAno.Items.Add(New ListItem("2015", 2015))
        ddlAno.Items.Add(New ListItem("2016", 2016))
        ddlAno.Items.Add(New ListItem("2017", 2017))
        ddlAno.Items.Add(New ListItem("2018", 2018))
        ddlAno.Items.Add(New ListItem("2019", 2019))
        ddlAno.Items.Add(New ListItem("2020", 2020))
        ddlAno.Items.Add(New ListItem("2021", 2021))
        ddlAno.Items.Add(New ListItem("2022", 2022))
        ddlAno.Items.Add(New ListItem("2023", 2023))
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf &
              "FROM Clientes C " &
              "INNER JOIN ClientesXTipos CT " & vbCrLf &
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf &
              "WHERE CT.Tipo_Id = 050 " & vbCrLf &
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' " & vbCrLf

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

    Private Sub Limpar()

    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Validar()
        Mensagem = ""

        If DdlUnidade.Text = "" Then
            Mensagem = "Unidade de negócio é obrigatório."
            Return Mensagem
        End If

        If DdlEmpresa.Text = "" Then
            Mensagem = "Empresa é obrigatório."
            Return Mensagem
        End If

        Return Mensagem
    End Function

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Validar()
            If Mensagem = "" Then

                Dim DataInicial As String = ddlAno.SelectedValue & "-" & ddlMes.SelectedValue & "-01"


                Sql = " if(object_id('tempdb..#Temp') is not null) begin drop table #Temp end;"

                Sql &= " SELECT Clientes.Nome as cliente,  NotasFiscais.Cliente_Id," & vbCrLf & _
                       "        Count(NotasFiscais.Nota_id) as QuantidadeNF," & vbCrLf & _
                       " 	    Sum(NotasFiscaisXItens.Valor) as Valor," & vbCrLf & _
                       " 	    convert(money,0.00) as ValorTotal" & vbCrLf & _
                       " into #temp" & vbCrLf

                Sql &= " FROM    NotasFiscais    INNER JOIN   Clientes ON  NotasFiscais.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
                       " 	    And NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                       " 	    INNER JOIN  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                       "        AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                       "        AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                       " Where  (Year(NotasFiscais.Movimento) = Year(getdate()))" & vbCrLf & _
                       "        And NotasFiscais.EntradaSaida_Id = 'S'" & vbCrLf & _
                       "        And NotasFiscais.Situacao = 1" & vbCrLf

                Sql &= " Group	By Clientes.Reduzido, NotasFiscais.Cliente_Id,  Clientes.Nome" & vbCrLf & _
                       " Order	By Valor desc" & vbCrLf

                Sql &= " update	#temp set ValorTotal =(" & vbCrLf & _
                       "        Select Case Convert(money, sum(Valor))" & vbCrLf & _
                       " from	( SELECT Clientes.Nome as cliente," & vbCrLf & _
                       "        Sum(NotasFiscaisXItens.Valor) as Valor" & vbCrLf & _
                       " FROM   NotasFiscais" & vbCrLf & _
                       " 		INNER JOIN   Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id" & vbCrLf & _
                       "        AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                       "        INNER JOIN  NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id" & vbCrLf & _
                       " 		AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id" & vbCrLf & _
                       " 	    AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf & _
                       " Where (Year(NotasFiscais.Movimento) = Year(getdate()))" & vbCrLf & _
                       " 	    And NotasFiscais.EntradaSaida_Id = 'S'" & vbCrLf & _
                       " 	    And NotasFiscais.Situacao = 1" & vbCrLf

                Sql &= " Group	By NotasFiscais.Cliente_Id, Clientes.Nome  ) as Consulta                    )" & vbCrLf & _
                       " " & vbCrLf & _
                       " update	#temp set cliente =  (Select cliente  from (" & vbCrLf & _
                       " SELECT	top 1 Clientes.Nome as cliente" & vbCrLf & _
                       "        from Clientes" & vbCrLf & _
                       " Where	left(#temp.Cliente_Id,8) = left(Cliente_Id,8) ) as consulta )" & vbCrLf & _
                       " select Cliente, " & vbCrLf & _
                       " 		sum(quantidadeNF) as quantidadeNF,   " & vbCrLf & _
                       " 		sum(Valor) as Valor,ValorTotal,  " & vbCrLf & _
                       " 		sum(Valor * 100 / ValorTotal) AS ParcelaIndividual" & vbCrLf & _
                       " from	#temp" & vbCrLf & _
                       "        where(valor > 100)" & vbCrLf & _
                       " group	by Cliente,ValorTotal  " & vbCrLf & _
                       " order	by Valor desc " & vbCrLf

                GridAbc.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
                GridAbc.DataBind()
            Else
                MsgBox(Me.Page, Mensagem, eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CurvaABC")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class