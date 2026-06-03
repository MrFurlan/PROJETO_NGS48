Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class TiposDeClientes
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet
    Dim objTipoDeCliente As [Lib].Negocio.TipoDeCliente
    Dim ListTiposDeClientes As ListTipoDeCliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("TiposDeClientes", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                CarregarTiposDeClientes()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarTiposDeClientes()
        If Funcoes.VerificaPermissao("TiposDeClientes", "LEITURA") Then
            ListTiposDeClientes = New ListTipoDeCliente(True)
            ListTiposDeClientes.Sort(Function(Tp1 As TipoDeCliente, tp2 As TipoDeCliente) Tp1.CodigoTipo.CompareTo(tp2.CodigoTipo))
            GridTiposDeClientes.DataSource = ListTiposDeClientes
            GridTiposDeClientes.DataBind()
        End If
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtCodigo.Enabled = True
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
    End Sub

    Protected Sub GridTiposDeClientes_PreRender(sender As Object, e As EventArgs)
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridTiposDeClientes.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridTiposDeClientes.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridTiposDeClientes.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If

    End Sub

    Protected Sub GridTiposDeClientes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            txtCodigo.Text = GridTiposDeClientes.SelectedRow.Cells(1).Text()
            txtDescricao.Text = GridTiposDeClientes.SelectedRow.Cells(2).Text()
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
            If Funcoes.VerificaPermissao("TiposDeClientes", "GRAVAR") Then
                objTipoDeCliente = New TipoDeCliente()
                objTipoDeCliente.IUD = "I"
                objTipoDeCliente.CodigoTipo = txtCodigo.Text
                objTipoDeCliente.Descricao = txtDescricao.Text

                If objTipoDeCliente.Salvar Then
                    Limpar()
                    CarregarTiposDeClientes()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("TiposDeClientes", "ALTERAR") Then
                objTipoDeCliente = New TipoDeCliente(txtCodigo.Text)
                objTipoDeCliente.IUD = "U"
                objTipoDeCliente.CodigoTipo = txtCodigo.Text
                objTipoDeCliente.Descricao = txtDescricao.Text

                If objTipoDeCliente.Salvar Then
                    Limpar()
                    CarregarTiposDeClientes()
                Else
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
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
            If Funcoes.VerificaPermissao("TiposDeClientes", "EXCLUIR") Then
                objTipoDeCliente = New TipoDeCliente(txtCodigo.Text)
                objTipoDeCliente.IUD = "D"

                If objTipoDeCliente.Salvar Then
                    Limpar()
                    CarregarTiposDeClientes()
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
            If Funcoes.VerificaPermissao("TiposDeClientes", "RELATORIO") Then
                Sql = "Select REPLICATE('0', 3 - LEN(CAST(Tipo_Id AS varchar))) + CAST(Tipo_Id AS varchar) as Codigo, Descricao From TiposdeClientes Order by Tipo_Id"

                DS = Banco.ConsultaDataSet(Sql, "Tabelas")

                Dim parameters = New Dictionary(Of String, Object)()
                parameters.Add("Titulo", "Tabela de Tipos de Clientes")
                parameters.Add("Codigo", "Código")
                parameters.Add("Descricao", "Descrição")

                Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "TiposDeClientes")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class