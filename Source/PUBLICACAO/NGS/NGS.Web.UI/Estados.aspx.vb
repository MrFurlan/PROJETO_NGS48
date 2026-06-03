Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Estados
    Inherits BasePage

    Dim Sql As String
    Dim ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Estados", "ACESSAR") Then
                    CarregarEstados()
                    Limpar()
                    For i As Integer = 1 To i = 10
                        MsgBox(Me.Page, i.ToString(), eTitulo.Info)
                    Next
                    'If Not String.Join("", ListUrl).Contains("Request.Url.ToString") Then
                    '    ListUrl.Add(Request.Url.ToString)
                    'End If
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
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

    Private Sub CarregarEstados()
        Dim nome As String = String.Empty
        Dim lstEstados As New [Lib].Negocio.Estados(True)

        GridEstados.DataSource = lstEstados
        GridEstados.DataBind()
    End Sub

    'Public Property ListUrl() As List(Of String)
    '    Get
    '        If Session("ListUrl") Is Nothing Then Session("ListUrl") = New List(Of String)
    '        Return CType(Session("ListUrl"), List(Of String))
    '    End Get
    '    Set(ByVal value As List(Of String))
    '        Session("ListUrl") = value
    '    End Set
    'End Property

    Private Sub Limpar()
        'Dim list As New List(Of String)
        'Dim existe As Boolean = False

        'If ListUrl IsNot Nothing AndAlso ListUrl.Count > 0 Then
        '    For Each item As String In ListUrl
        '        If item = Request.Url.ToString Then MsgBox(Me.Page, item)
        '    Next
        'End If

        txtCodigo.Text = ""
        txtRegiao.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            param &= "UF: " & txtCodigo.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtRegiao.Text) Then
            param &= "Região: " & txtRegiao.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            param &= "Descrição: " & txtDescricao.Text & "."
        End If

        Return param
    End Function

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtCodigo.Text) Then
            MsgBox(Me.Page, "Código do Estado não foi informado!", eTitulo.Info)

            MsgBox(Me.Page, "Região do Estado não foi informada!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Descrição do Estado não foi informada!", eTitulo.Info)
            Return False
        End If
        Return True
    End Function

    Protected Sub GridEstados_PreRender(sender As Object, e As EventArgs)
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridEstados.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridEstados.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridEstados.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If

    End Sub

    Protected Sub GridEstados_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridEstados.SelectedRow.Cells(1).Text()
            txtRegiao.Text = GridEstados.SelectedRow.Cells(2).Text()
            txtDescricao.Text = GridEstados.SelectedRow.Cells(3).Text()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
            txtCodigo.Enabled = False
            txtRegiao.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Estados", "GRAVAR") Then
                If ValidarCampos() Then
                    Dim objEstado As New Estado()
                    objEstado.Codigo = Trim(txtCodigo.Text)
                    objEstado.Descricao = RTrim(txtDescricao.Text)
                    objEstado.Regiao = Trim(txtRegiao.Text)
                    objEstado.IUD = "I"
                    If objEstado.Salvar Then
                        MsgBox(Me.Page, "Informação inserida com sucesso.", eTitulo.Sucess)
                        Limpar()
                        CarregarEstados()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar Registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        If Funcoes.VerificaPermissao("Estados", "ALTERAR") Then
            If ValidarCampos() Then
                Dim objEstado As New Estado(txtCodigo.Text)
                objEstado.Descricao = RTrim(txtDescricao.Text)
                objEstado.Regiao = Trim(txtRegiao.Text)
                objEstado.IUD = "U"
                If objEstado.Salvar Then
                    MsgBox(Me.Page, "Informação alterada com sucesso.", eTitulo.Sucess)
                    Limpar()
                    CarregarEstados()
                End If
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para alterar Registro")
        End If
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Estados", "EXCLUIR") Then
                If ValidarCampos() Then
                    Dim objEstado As New Estado(txtCodigo.Text)
                    If objEstado.VerificaClientes() Then
                        MsgBox(Me.Page, "Estado não pode ser excluído por estar relacionado no cadastro de Clientes", eTitulo.Info)
                    Else
                        objEstado.Descricao = RTrim(txtDescricao.Text)
                        objEstado.Regiao = Trim(txtRegiao.Text)
                        objEstado.IUD = "D"
                        If objEstado.Salvar Then
                            MsgBox(Me.Page, "Informação excluída com sucesso.", eTitulo.Sucess)
                            Limpar()
                            CarregarEstados()
                        End If
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
            If Funcoes.VerificaPermissao("Estados", "RELATORIO") Then
                Sql = "  SELECT Estado_Id as UF, Regiao, Descricao" & vbCrLf & _
                      " FROM Estados " & vbCrLf & _
                      " ORDER BY Estado_Id" & vbCrLf
                ds = Banco.ConsultaDataSet(Sql, "Estados")

                Funcoes.BindReport(Me.Page, ds, "Cr_Estados", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para visualizar o relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Estados")
            'Dim cod As String = FuncoesStrings.CodificarPara64Bits(txtDescricao.Text)
            'Dim decod As String = FuncoesStrings.DecodificarDe64Bits(cod)
            'MsgBox(Me.Page, cod & " / " & decod, eTitulo.Info, False)           
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class