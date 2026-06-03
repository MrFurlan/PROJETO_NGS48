Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ProdutosDeTerceiros
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Mensagem As String
    Dim Cliente As String
    Dim campo() As String
    Dim Porcentagem As String
    Dim Empresa() As String
    Dim Deposito() As String


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Custos)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("ProdutosDeTerceiros", "ACESSAR") Then
                    CargaUnidade()
                    VerificaUnidade()
                    CargaDepositos()
                    CargaGrupos()
                    ddl.Carregar(DdlAno, CarregarDDL.Tabela.Ano, "2016;10;C", False)
                    Limpar()
                    ddlUnidade.Focus()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf & _
              "FROM Clientes C " & vbCrLf & _
              "INNER JOIN ClientesXTipos CT " & vbCrLf & _
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf & _
              "WHERE CT.Tipo_Id = 050 " & vbCrLf & _
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresas()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        'Limpar()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            ddlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next

        ddlEmpresa.Items.Insert(0, "")
        ddlEmpresa.SelectedIndex = 0
    End Sub

    Private Sub CargaDepositos()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlDeposito.Items.Clear()

        Sql = " SELECT  Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, " & vbCrLf & _
              "         Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
              " FROM    Clientes INNER JOIN" & vbCrLf & _
              "         ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id " & vbCrLf & _
              "     AND Clientes.Endereco_Id = ClientesXTipos.Endereco_Id" & vbCrLf & _
              " where   ClientesXTipos.Tipo_Id = 3" & vbCrLf & _
              " Order by  Clientes.Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            DdlDeposito.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlDeposito.Items.Insert(0, "")
        DdlDeposito.SelectedIndex = 0
    End Sub

    Private Sub CargaGrupos()
        DdlGrupo.Items.Clear()

        Sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao" & vbCrLf & _
              " FROM         GruposDeEstoques INNER JOIN" & vbCrLf & _
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf & _
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) order by GruposDeEstoques.Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlGrupo.Items.Add(New ListItem(Dr("Grupo_Id") & "-" & Dr("Descricao"), Dr("Grupo_Id")))
        Next

        DdlGrupo.Items.Insert(0, "")
        DdlGrupo.SelectedIndex = 0
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarProduto()
        DdlProduto.Items.Clear()

        Sql = "SELECT Produto_Id, Nome FROM Produtos WHERE Grupo = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
              " Order by Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlProduto.Items.Add(New ListItem(Dr("Produto_Id") & "-" & Dr("Nome"), Dr("Produto_Id")))
        Next

        DdlProduto.Items.Insert(0, "")
        DdlProduto.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        DdlDeposito.SelectedIndex = 0
        DdlGrupo.SelectedIndex = 0
        DdlProduto.Items.Clear()
        'DdlCusto.SelectedIndex = 0
        DdlMes.SelectedValue = Now.Month
        DdlAno.SelectedValue = Now.Year

        'txtQuantidade.Text = ""
        'txtValorDaMercadoria.Text = ""
        'txtValorDoFrete.Text = ""
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Function Valida() As Boolean
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Campo empresa é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlDeposito.SelectedValue) Then
            MsgBox(Me.Page, "Campo depósito é obrigatório.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlGrupo.SelectedValue) Then
            MsgBox(Me.Page, "Campo grupo é obrigatório.")
            Return False
        End If
        Return True
    End Function

    Protected Sub DdlMes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub DdlAno_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Dim crpt As New ReportDocument()

        Try
            If Valida() Then
                Dim Ds As New DataSet
                Empresa = ddlEmpresa.SelectedValue.Split("-")
                Deposito = DdlDeposito.SelectedValue.Split("-")

                Sql = "SELECT  R.Produto, P.Nome as NomeProduto,  " & vbCrLf & _
                        "R.Empresa_Id,R.EndEmpresa_Id,E.Nome as NomeEmpresa,  " & vbCrLf & _
                        "R.Deposito, R.EndDeposito,D.Nome as NomeDeposito,  " & vbCrLf & _
                        "Sum(DebitoQuantidade) as DebitoQuantidade,  " & vbCrLf & _
                        "Sum(CreditoQuantidade) as CreditoQuantidade,  " & vbCrLf & _
                        "Sum(DebitoOficial) as DebitoOficial,  " & vbCrLf & _
                        "Sum(CreditoOficial) as CreditoOficial, " & vbCrLf & _
                        "Sum(Isnull(Sb.SaldoAnteriorOficial,0)) as SaldoAnteriorOficial, " & vbCrLf & _
                        "Sum(Isnull(Sb.SaldoAnteriorOficial,0)) - Sum(DebitoOficial) + Sum(CreditoOficial) as SaldoOficial, " & vbCrLf & _
                        "Sum(Isnull(Sb1.SaldoAnteriorQuantidade,0)) as SaldoAnteriorQuantidade, " & vbCrLf & _
                        "Sum(Isnull(Sb1.SaldoAnteriorQuantidade,0)) - Sum(DebitoQuantidade) + Sum(CreditoQuantidade) as SaldoQuantidade, " & vbCrLf & _
                        "E.Cidade as EmpresaCidade,E.Estado as EmpresaEstado,E.Reduzido as EmpresaReduzido, " & vbCrLf & _
                        "D.Cidade as DepositoCidade, D.Estado as DepositoEstado, D.Reduzido as DepositoReduzido " & vbCrLf & _
                        "FROM Razao R   " & vbCrLf & _
                        "INNER JOIN Clientes AS E   " & vbCrLf & _
                        "ON R.Empresa_Id = E.Cliente_Id  " & vbCrLf & _
                        "AND R.EndEmpresa_Id = E.Endereco_Id  " & vbCrLf & _
                        "INNER JOIN Clientes AS D  " & vbCrLf & _
                        "ON R.Deposito = D.Cliente_Id  " & vbCrLf & _
                        "AND R.EndDeposito = D.Endereco_Id  " & vbCrLf & _
                        "INNER JOIN  Produtos P  " & vbCrLf & _
                        "ON R.Produto = P.Produto_Id  " & vbCrLf & _
                        "Left Join (  " & vbCrLf & _
                        "Select Produto,Empresa_Id,EndEmpresa_Id, Deposito, EndDeposito, Sum((DebitoOficial- CreditoOficial)) as SaldoAnteriorOficial " & vbCrLf & _
                        "From Razao " & vbCrLf & _
                        "Where isnull(Situacao,0) = '10' " & vbCrLf & _
                        "AND Conta_Id= '116010030' " & vbCrLf & _
                        "AND Movimento_Id < '" & DdlAno.SelectedValue + "-" + DdlMes.SelectedValue & "-01' " & vbCrLf & _
                        "Group by Produto,Empresa_Id,EndEmpresa_Id, Deposito, EndDeposito	 " & vbCrLf & _
                        " )sb " & vbCrLf & _
                        "On sb.Produto=R.Produto " & vbCrLf & _
                        "AND sb.Empresa_Id = R.Empresa_Id " & vbCrLf & _
                        "AND sb.EndEmpresa_Id=R.EndEmpresa_Id " & vbCrLf & _
                        "AND sb.Deposito=R.Deposito " & vbCrLf & _
                        "AND sb.EndDeposito=R.EndDeposito " & vbCrLf & _
                        "Left Join ( " & vbCrLf & _
                        "Select Produto,Empresa_Id,EndEmpresa_Id, Deposito, EndDeposito, Sum((DebitoQuantidade- CreditoQuantidade)) as SaldoAnteriorQuantidade " & vbCrLf & _
                        "From Razao " & vbCrLf & _
                        "Where isnull(Situacao,0) = '10' " & vbCrLf & _
                        "AND Conta_Id= '116010030' " & vbCrLf & _
                        "AND Movimento_Id < '" & DdlAno.SelectedValue + "-" + DdlMes.SelectedValue & "-01' " & vbCrLf & _
                        "Group by Produto,Empresa_Id,EndEmpresa_Id, Deposito, EndDeposito	 " & vbCrLf & _
                        ")sb1 " & vbCrLf & _
                        "On sb1.Produto=R.Produto " & vbCrLf & _
                        "AND sb1.Empresa_Id = R.Empresa_Id " & vbCrLf & _
                        "AND sb1.EndEmpresa_Id=R.EndEmpresa_Id " & vbCrLf & _
                        "AND sb1.Deposito=R.Deposito " & vbCrLf & _
                        "AND sb1.EndDeposito=R.EndDeposito " & vbCrLf & _
                        "Where isnull(R.Situacao,0) = 10 " & vbCrLf & _
                        "AND Conta_Id= '116010030' " & vbCrLf & _
                        "AND (month(Movimento_Id)= '" & DdlMes.SelectedValue & "') AND Year(Movimento_Id)= '" & DdlAno.SelectedValue & "' " & vbCrLf

                If ddlEmpresa.Text <> "" Then
                    If ChkTodasAsFiliais.Checked = True Then
                        Sql &= " AND R.Empresa_ID  Like  '" & Empresa(0).Substring(0, 8) & "%'  "
                    Else
                        Sql &= " And R.Empresa_ID = '" & Empresa(0) & "' "
                        Sql &= " And R.EndEmpresa_Id = " & Empresa(1)
                    End If
                End If
                If DdlDeposito.Text <> "" Then
                    Sql &= " And R.Deposito = '" & Deposito(0) & "' "
                    Sql &= " And R.EndDeposito = " & Deposito(1)
                End If
                If DdlProduto.Text <> "" Then
                    Sql &= " And R.Produto = '" & DdlProduto.SelectedValue & "' "
                End If
                Sql &= "Group By  " & vbCrLf & _
                        "R.Empresa_Id,R.EndEmpresa_Id,E.Nome, " & vbCrLf & _
                        "R.Deposito,R.EndDeposito,D.Nome,  " & vbCrLf & _
                        "E.Cidade ,E.Estado ,E.Reduzido , " & vbCrLf & _
                        "D.Cidade, D.Estado, D.Reduzido,R.Produto, P.Nome " & vbCrLf & _
                        "Order by  " & vbCrLf & _
                        "R.Empresa_Id,R.EndEmpresa_Id, " & vbCrLf & _
                        "R.Deposito, R.EndDeposito, R.Produto, P.Nome " & vbCrLf


                Ds = Banco.ConsultaDataSet(Sql, "ProdutosDeTerceiros")

                '---Definição do Relatorio Crystal Report---------
                If RadPorProduto.Checked = True Then
                    crpt.FileName = Server.MapPath("~/Reports/Cr_ProdutosDeTerceirosXProdutos.rpt")
                Else
                    crpt.FileName = Server.MapPath("~/Reports/Cr_ProdutosDeTerceirosXFilial.rpt")
                End If

                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(Ds)
                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As ParameterFieldDefinitions
                Dim crparameterfielddefinition As ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("Empresa")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Periodo")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = "Relativo a " & Funcoes.MesPorExtenso(DdlMes.Text) & " de " & DdlAno.SelectedValue
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo2 & "');", True)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Limpar()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ProdutosDeTerceiros")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class