Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CentrosDeCustos
    Inherits BasePage

    Dim Sql As String
    Dim DS As DataSet

    Private Property ListCentroDeCusto As ListCentroDeCusto
        Get
            Return Session("ListCentroDeCusto" + HID.Value)
        End Get
        Set(ByVal value As ListCentroDeCusto)
            Session("ListCentroDeCusto" + HID.Value) = value
        End Set
    End Property

    Private Property objCentroDeCusto() As CentroDeCusto
        Get
            Return CType(Session("objCentroDeCusto" + HID.Value), CentroDeCusto)
        End Get
        Set(ByVal value As CentroDeCusto)
            Session("objCentroDeCusto" + HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CentrosDeCustos", "ACESSAR") Then
                    CarregarCentrosDeCustos()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregarCentrosDeCustos()
        If Funcoes.VerificaPermissao("CentrosDeCustos", "LEITURA") Then
            If ListCentroDeCusto Is Nothing Then
                ListCentroDeCusto = New ListCentroDeCusto(True)
            End If

            GridCentroDeCusto.DataSource = ListCentroDeCusto
            GridCentroDeCusto.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub Limpar()
        Try
            txtCodigo.Text = String.Empty
            txtDescricao.Text = String.Empty
            txtCodigo.Enabled = True
            lnkNovo.Parent.Visible = True
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
            ddlAtivo.SelectedValue = String.Empty
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "Código: " & txtCodigo.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text
        End If
        Return param
    End Function

    Protected Sub GridCentrosDeCustos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try

            Dim Pos = ListCentroDeCusto.FindIndex(Function(s) s.Codigo = GridCentroDeCusto.SelectedRow.Cells(1).Text())

            txtCodigo.Text = ListCentroDeCusto(Pos).Codigo
            txtDescricao.Text = ListCentroDeCusto(Pos).Descricao
            ddlAtivo.SelectedValue = IIf(ListCentroDeCusto(Pos).Ativo, 1, 0)
            txtCodigo.Enabled = False
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CentrosDeCustos", "GRAVAR") Then
                If ValidaCampos() Then

                    objCentroDeCusto = New CentroDeCusto()
                    objCentroDeCusto.Codigo = txtCodigo.Text
                    objCentroDeCusto.Descricao = txtDescricao.Text
                    objCentroDeCusto.Ativo = ddlAtivo.SelectedValue

                    objCentroDeCusto.IUD = "I"
                    If objCentroDeCusto.Salvar() Then
                        ListCentroDeCusto.Add(objCentroDeCusto)
                        MsgBox(Me.Page, "Registro Gravado com Sucesso.", eTitulo.Sucess)
                    End If
                    Limpar()
                    CarregarCentrosDeCustos()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos(Optional ByVal Novo As Boolean = True) As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Informe o código.")
            Return False
        ElseIf ListCentroDeCusto.FindIndex(Function(s) s.Codigo = txtCodigo.Text) > 0 AndAlso Novo Then
            MsgBox(Me.Page, "Centro de Custo Já Cadastrado.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Informe a descrição.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlAtivo.SelectedValue) Then
            MsgBox(Me.Page, "Selecione Ativo ou Inativo.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("CentrosDeCustos", "ALTERAR") Then
                If ValidaCampos(False) Then

                    Dim pos As Integer = ListCentroDeCusto.FindIndex(Function(s) s.Codigo = txtCodigo.Text)

                    ListCentroDeCusto(pos).Descricao = txtDescricao.Text
                    ListCentroDeCusto(pos).Ativo = ddlAtivo.SelectedValue
                    ListCentroDeCusto(pos).IUD = "U"
                    If ListCentroDeCusto(pos).Salvar Then
                        MsgBox(Me.Page, "Registro Alterado com Sucesso.", eTitulo.Sucess)
                    End If

                    Limpar()
                    CarregarCentrosDeCustos()
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
            If Funcoes.VerificaPermissao("CentrosDeCustos", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
                    MsgBox(Me.Page, "Selecione um Centro de Custo para Excluir.", eTitulo.Info)
                Else
                    Dim pos As Integer = ListCentroDeCusto.FindIndex(Function(s) s.Codigo = txtCodigo.Text)

                    If ListCentroDeCusto(pos).JaUtilizado() Then
                        MsgBox(Me.Page, "Centro de Custo já utilizado em alguma nota, rateio de nota e ou contabilização por isso não pode ser excluído.", eTitulo.Info)
                    Else
                        ListCentroDeCusto(pos).IUD = "D"
                        If ListCentroDeCusto(pos).Salvar() Then
                            ListCentroDeCusto.Remove(objCentroDeCusto)
                            MsgBox(Me.Page, "Registro Excluído com Sucesso.", eTitulo.Sucess)
                        End If
                        Limpar()
                        CarregarCentrosDeCustos()
                    End If
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
            If Funcoes.VerificaPermissao("CentrosDeCustos", "RELATORIO") Then
                Sql = " Select CentroDeCusto_Id as Codigo, Descricao From CentrosDeCustos" & vbCrLf
                If Not String.IsNullOrWhiteSpace(ddlAtivo.SelectedValue) Then Sql &= " Where Ativo = " & ddlAtivo.SelectedValue & vbCrLf
                Sql &= " Order by CentroDeCusto_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Funcoes.BindReport(Me.Page, DS, "Cr_CentroDeCustos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CentrosDeCustos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub GridCentrosDeCustos_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridCentroDeCusto.RowDataBound
        Try
            If e.Row.RowType = DataControlRowType.DataRow Then
                If CBool(e.Row.Cells(3).Text) Then
                    e.Row.Cells(3).Text = "Sim"
                Else
                    e.Row.Cells(3).Text = "Não"
                    e.Row.ForeColor = Drawing.Color.Red
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    <System.Web.Services.WebMethod> _
    Public Shared Function MeuMetodo(ByVal desc As String) As String
        Dim qtdeclick As Integer
        qtdeclick = CInt(desc.Split("-")(1).Trim) + 1
        Return "click - " & qtdeclick
    End Function

    Protected Sub GridCentrosDeCustos_PreRender(sender As Object, e As EventArgs) Handles GridCentroDeCusto.PreRender
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridCentroDeCusto.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridCentroDeCusto.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridCentroDeCusto.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If
    End Sub
End Class