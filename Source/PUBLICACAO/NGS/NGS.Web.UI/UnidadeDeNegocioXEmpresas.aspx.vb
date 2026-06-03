Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class UnidadeDeNegocioXEmpresas
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("UnidadeDeNegocioXEmpresas", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarUnidadeDeNegocio()
                CarregarEmpresas()
                AtualizarGrid()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
    End Sub

    Private Sub CarregarEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXTipos, "1", True)
    End Sub

    Private Function ValidarCampos() As Boolean
        If ddlUnidadeDeNegocio.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Unidade não foi informada!")
            Return False
        ElseIf ddlEmpresa.SelectedIndex <= 0 Then
            MsgBox(Me.Page, "Empresa não foi informada!")
            Return False
        Else
            Return True
        End If
    End Function

    Private Sub AtualizarGrid()
        Sql = "SELECT GruposXEmpresas.Empresa_Id AS Unidade, CliXUni.Nome AS NomeUnidade, GruposXEmpresas.Cliente_Id AS Empresa, " & vbCrLf & _
              "		GruposXEmpresas.EndCliente_Id AS EndEmpresa, CliXEmp.Nome AS NomeEmpresa, CliXEmp.Cidade " & vbCrLf & _
              "FROM  GruposXEmpresas " & vbCrLf & _
              "	INNER JOIN Clientes AS CliXUni " & vbCrLf & _
              "			ON  GruposXEmpresas.Empresa_Id    = CliXUni.Cliente_Id " & vbCrLf & _
              "			AND GruposXEmpresas.EndEmpresa_Id = CliXUni.Endereco_Id " & vbCrLf & _
              "	INNER JOIN Clientes AS CliXEmp " & vbCrLf & _
              "			ON  GruposXEmpresas.Cliente_Id    = CliXEmp.Cliente_Id " & vbCrLf & _
              "			AND GruposXEmpresas.EndCliente_Id = CliXEmp.Endereco_Id " & vbCrLf

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(Sql, "GruposXEmpresas")

        gridUnidadeDeNegocioXEmpresas.DataSource = ds
        gridUnidadeDeNegocioXEmpresas.DataBind()
    End Sub

    Private Sub Limpar()
        ddlUnidadeDeNegocio.SelectedIndex = 0
        ddlEmpresa.SelectedIndex = 0
        ddlUnidadeDeNegocio.Enabled = True
        ddlEmpresa.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function getparam() As String
        Dim Texto As String = ""
        If Not String.IsNullOrWhiteSpace(ddlUnidadeDeNegocio.SelectedValue) Then
            Texto &= "Unidade: " & ddlUnidadeDeNegocio.SelectedValue & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Texto &= "Empresa: " & ddlEmpresa.SelectedValue & vbCrLf
        End If
        Return Texto
    End Function

    Protected Sub gridUnidadeDeNegocioXEmpresas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ddlUnidadeDeNegocio.SelectedValue = gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(1).Text()
            ddlEmpresa.SelectedValue = gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(3).Text() & "-" & gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(4).Text()
            ddlUnidadeDeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeDeNegocioXEmpresas", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim SqlArray As New ArrayList
                    Dim Empresa() As String = ddlEmpresa.SelectedValue.ToString.Split("-")
                    Sql = "INSERT Into GruposXEmpresas(Empresa_id, EndEmpresa_id, Cliente_id , EndCliente_id, Descricao) " & vbCrLf & _
                          " Values ('" & ddlUnidadeDeNegocio.SelectedValue & "',0,'" & Empresa(0) & "'," & Empresa(1) & ",'" & ddlUnidadeDeNegocio.SelectedItem.Text & "')" & vbCrLf
                    SqlArray.Add(Sql)
                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        AtualizarGrid()
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("UnidadeDeNegocioXEmpresas", "EXCLUIR") Then
                Dim SqlArray As New ArrayList

                Sql = "DELETE FROM GruposXEmpresas " & vbCrLf & _
                      " WHERE Empresa_id    = '" & gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(1).Text() & "'" & vbCrLf & _
                      "   And EndEmpresa_Id = 0 " & vbCrLf & _
                      "   And Cliente_id    = '" & gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(3).Text() & "'" & vbCrLf & _
                      "   And EndCliente_id = " & gridUnidadeDeNegocioXEmpresas.SelectedRow.Cells(4).Text() & vbCrLf

                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) Then
                    Limpar()
                    AtualizarGrid()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try

            Sql = "SELECT GruposXEmpresas.Empresa_Id + ' - ' + CliXUni.Nome AS Unidade,       " & vbCrLf & _
                  "GruposXEmpresas.Cliente_Id + ' - ' + CliXEmp.Nome + ' - ' + CliXEmp.Cidade AS Empresa   " & vbCrLf & _
                  "       FROM GruposXEmpresas                                                          " & vbCrLf & _
                  "	INNER JOIN Clientes AS CliXUni                                                       " & vbCrLf & _
                  "ON  GruposXEmpresas.Empresa_Id    = CliXUni.Cliente_Id                                  " & vbCrLf & _
                  "	AND GruposXEmpresas.EndEmpresa_Id = CliXUni.Endereco_Id                              " & vbCrLf & _
                  "	INNER JOIN Clientes AS CliXEmp                                                       " & vbCrLf & _
                  "	ON  GruposXEmpresas.Cliente_Id    = CliXEmp.Cliente_Id                               " & vbCrLf & _
                  "	AND GruposXEmpresas.EndCliente_Id = CliXEmp.Endereco_Id                              " & vbCrLf
            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(Sql, "UnidadeDeNegocioXEmpresas")

            Dim parameters = New Dictionary(Of String, Object)()
            parameters.Add("ConsultaParam", getparam())

            Funcoes.BindReport(Me.Page, ds, "Cr_UnidadeDeNegocioXEmpresas", eExportType.PDF, parameters)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "UnidadeDeNegocioXEmpresas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class