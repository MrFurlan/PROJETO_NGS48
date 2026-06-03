Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class DatasNaoProgramaveis
    Inherits BasePage

#Region "Events"
    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load, Me.Load
        Me.setMenu(eModulo.Gestao)
        Try
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("DatasNaoProgramaveis", "ACESSAR") Then
                    HID.Value = Guid.NewGuid().ToString
                    ddl.Carregar(DdlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                    Funcoes.VerificaEmpresa(DdlEmpresa)
                    BindGridView()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub DdlEmpresa_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresa.SelectedIndexChanged
        Try
            BindGridView()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub grd_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles grd.SelectedIndexChanged
        Try
            DdlEmpresa.SelectedValue = Server.HtmlDecode(grd.SelectedRow.Cells(1).Text()) & "-0"
            txtData.Text = Server.HtmlDecode(grd.SelectedRow.Cells(3).Text())
            txtDescricao.Text = Server.HtmlDecode(grd.SelectedRow.Cells(4).Text())

            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtData.Enabled = False
            DdlEmpresa.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If validarCampos() Then
                Dim sql As String = "update DatasNaoProgramaveis set Descricao = '" & txtDescricao.Text.Trim() & "'" & vbCrLf & _
                                    " where Data_Id = '" & txtData.Text.ToSqlDate() & "'" & IIf(Not chkTodasFiliais.Checked, " and Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' and EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1), "")

                If Banco.GravaBanco(sql) Then
                    MsgBox(Me.Page, "Alteração(ões) realizada(s) com sucesso.")
                    Limpar()
                    BindGridView()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If validarCampos() Then
                Dim sql As String
                If chkTodasFiliais.Checked Then
                    sql = "insert into DatasNaoProgramaveis (Empresa_Id, EndEmpresa_Id, Data_Id, Descricao)                         " & vbCrLf & _
                          " select Empresa_Id, EndEmpresa_Id, '" & txtData.Text.ToSqlDate() & "', '" & txtDescricao.Text.Trim() & "'" & vbCrLf & _
                          "   from ClientesXEmpresas ce                                                                             " & vbCrLf & _
                          "  where not exists (select *                                                                             " & vbCrLf & _
                          "                      from datasnaoprogramaveis dnp                                                      " & vbCrLf & _
                          " 				    where dnp.empresa_id     = ce.empresa_id                                             " & vbCrLf & _
                          "                       and dnp.endempresa_id = ce.endempresa_id                                          " & vbCrLf & _
                          " 				      and dnp.data_id        = '" & txtData.Text.ToSqlDate() & "')                       " & vbCrLf
                Else
                    sql = "insert into DatasNaoProgramaveis (Empresa_Id, EndEmpresa_Id, Data_Id, Descricao) " & vbCrLf & _
                          "                           Values('" & DdlEmpresa.SelectedValue.Split("-")(0) & "','" & DdlEmpresa.SelectedValue.Split("-")(1) & "','" & txtData.Text.ToSqlDate & "','" & txtDescricao.Text.Trim & "')"
                End If

                If Banco.GravaBanco(sql) Then
                    MsgBox(Me.Page, "Inclusão(ões) realizada(s) com sucesso.")
                    Limpar()
                    BindGridView()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If validarCampos() Then
                Dim sql = " delete DatasNaoProgramaveis" & vbCrLf & _
                          " where Data_Id = '" & txtData.Text.ToSqlDate() & "'" & IIf(Not chkTodasFiliais.Checked, " and Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' and EndEmpresa_Id = " & DdlEmpresa.SelectedValue.Split("-")(1), "")

                If Banco.GravaBanco(sql) Then
                    MsgBox(Me.Page, "Exclusão(ões) realizada(s) com sucesso.")
                    Limpar()
                    BindGridView()
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("DatasNaoProgramaveis", "RELATORIO") Then


                Dim sql As String = "SELECT Dp.Empresa_Id,                                                     " & vbCrLf & _
                                    "       Dp.EndEmpresa_Id,                                                  " & vbCrLf & _
                                    "       CONVERT(VARCHAR(10), Dp.Data_Id, 103) as Data_Id,                  " & vbCrLf & _
                                    "	    Dp.Descricao,                                                      " & vbCrLf & _
                                    "       Cl.Nome,                                                           " & vbCrLf & _
                                    "       Cl.Cidade,                                                         " & vbCrLf & _
                                    "	    Cl.Estado,                                                         " & vbCrLf & _
                                    "	    Cl.Reduzido                                                        " & vbCrLf & _
                                    "  FROM DatasNaoProgramaveis Dp                                            " & vbCrLf & _
                                    " INNER JOIN Clientes Cl                                                   " & vbCrLf & _
                                    "    ON Dp.Empresa_Id COLLATE Latin1_General_CI_AS = Cl.Cliente_Id         " & vbCrLf & _
                                    "   AND Dp.EndEmpresa_Id = Cl.Endereco_Id                                  " & vbCrLf

                If DdlEmpresa.SelectedValue.Length > 0 Then
                    sql &= " AND Dp.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf
                End If

                sql &= " ORDER BY year(Dp.Data_Id) desc, Dp.Data_Id asc, Cl.Reduzido asc  " & vbCrLf

                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "DatasNaoProgramaveis")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Parametros", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_DatasNaoProgramaveis", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "DatasNaoProgramaveis")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

#Region "Methods"
    Private Sub BindGridView()
        Dim sql As String = "SELECT Dp.Empresa_Id,                                                     " & vbCrLf & _
                            "       Dp.EndEmpresa_Id,                                                  " & vbCrLf & _
                            "       CONVERT(VARCHAR(10), Dp.Data_Id, 103) as Data_Id,                  " & vbCrLf & _
                            "	    Dp.Descricao,                                                      " & vbCrLf & _
                            "       Cl.Nome,                                                           " & vbCrLf & _
                            "       Cl.Cidade,                                                         " & vbCrLf & _
                            "	    Cl.Estado,                                                         " & vbCrLf & _
                            "	    Cl.Reduzido,                                                       " & vbCrLf & _
                            "       Cl.Cidade + '/' + Cl.Estado as CidadeUF                            " & vbCrLf & _
                            "  FROM DatasNaoProgramaveis Dp                                            " & vbCrLf & _
                            " INNER JOIN Clientes Cl                                                   " & vbCrLf & _
                            "    ON Dp.Empresa_Id COLLATE Latin1_General_CI_AS = Cl.Cliente_Id         " & vbCrLf & _
                            "   AND Dp.EndEmpresa_Id = Cl.Endereco_Id                                  " & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            sql &= " AND Dp.Empresa_Id = '" & DdlEmpresa.SelectedValue.Split("-")(0) & "' " & vbCrLf
        End If
        sql &= " ORDER BY year(Dp.Data_Id) desc, Dp.Data_Id asc, Cl.Reduzido asc  " & vbCrLf

        grd.DataSource = Banco.ConsultaDataSet(sql, "TblDatasNaoProgramaveis")
        grd.DataBind()
    End Sub

    Private Sub Limpar()
        txtData.Text = String.Empty
        txtDescricao.Text = String.Empty
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        txtData.Enabled = True
        DdlEmpresa.Enabled = True
        Funcoes.VerificaEmpresa(DdlEmpresa)
    End Sub

    Private Function validarCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtData.Text) Then
            MsgBox(Me.Page, "Informe a data.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        End If
        Return True
    End Function

    Private Function getParam() As String
        Dim param As String = ""

        If Not String.IsNullOrWhiteSpace(DdlEmpresa.SelectedValue) Then
            param &= "Empresa: " & DdlEmpresa.SelectedItem.Text.Split(".")(0) & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtData.Text) Then
            param &= "Data: " & txtData.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text
        End If

        Return param
    End Function

#End Region
End Class