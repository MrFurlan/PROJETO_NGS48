Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Marcas
    Inherits BasePage

    Private Sql As String
    Private SqlArray As New ArrayList
    Private DS As DataSet
    Private objMarca As [Lib].Negocio.Marca

    Public Property SessaoMarca As [Lib].Negocio.Marca
        Get
            Return Session("objMarca" & HID.Value)
        End Get
        Set(ByVal value As [Lib].Negocio.Marca)
            Session("objMarca" & HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("Marcas", "ACESSAR") Then
                CarregarMarcas()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarMarcas()
        If Funcoes.VerificaPermissao("Marcas", "LEITURA") Then
            Dim objListaMarcas As New ListMarca(True)
            GridMarcas.DataSource = objListaMarcas
            GridMarcas.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridMarcas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridMarcas.SelectedIndexChanged
        Try
            Limpar()
            txtCodigo.Text = GridMarcas.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridMarcas.SelectedRow.Cells(2).Text()
            objMarca = New [Lib].Negocio.Marca(GridMarcas.SelectedRow.Cells(1).Text())
            SessaoMarca = objMarca
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaMarca() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "A Descrição da Marca é Obrigatória.")
            txtDescricao.Focus()
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Marcas", "GRAVAR") Then
                If ValidaMarca() Then
                    objMarca = New [Lib].Negocio.Marca()
                    objMarca.Descricao = Trim(txtDescricao.Text.ToUpper)
                    objMarca.IUD = "I"
                    If Not objMarca.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarMarcas()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Marcas", "ALTERAR") Then
                If ValidaMarca() Then
                    objMarca = SessaoMarca
                    objMarca.Descricao = Trim(txtDescricao.Text.ToUpper)
                    objMarca.IUD = "U"

                    If Not objMarca.Salvar Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        Limpar()
                        CarregarMarcas()
                    End If
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
            If Funcoes.VerificaPermissao("Marcas", "EXCLUIR") Then
                objMarca = SessaoMarca

                'If objMarca.LiberaExclusaoDaMarca Then
                objMarca.IUD = "D"
                If Not objMarca.Salvar Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Limpar()
                    CarregarMarcas()
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
            If Funcoes.VerificaPermissao("Marcas", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Marca_Id AS varchar))) + CAST(Marca_Id AS varchar) as Codigo, Descricao From Marca Order by Marca_Id"
                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters As New Dictionary(Of String, Object)
                parameters.Add("Titulo", "Relatório De Marcas De Produto.")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")
              
                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permisão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Marcas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class