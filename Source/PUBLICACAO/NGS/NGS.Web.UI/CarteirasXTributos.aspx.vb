Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CarteirasXTributos
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim DS As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CarteirasXTributos", "ACESSAR") Then
                If Funcoes.VerificaPermissao("CarteirasXTributos", "LEITURA") Then
                    CarregarCarteirasXTributos()
                    CarregarCarteiras()
                    CarregarTributos()
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CarregarCarteirasXTributos()
        Sql = "  SELECT     CarteirasXTributos.Carteira_Id AS Codigo, ComprasXProdutos.Descricao, CarteirasXTributos.Tributo_ID as Tributo" & vbCrLf & _
                " FROM      CarteirasXTributos INNER JOIN" & vbCrLf & _
                " ComprasXProdutos ON CarteirasXTributos.Carteira_Id = ComprasXProdutos.Produto_Id" & vbCrLf & _
                " ORDER BY Codigo" & vbCrLf

        GridCarteirasXTributos.DataSource = Banco.ConsultaDataSet(Sql, "Consulta")
        GridCarteirasXTributos.DataBind()
    End Sub

    Private Sub CarregarCarteiras()
        DdlCarteira.Items.Clear()

        Sql = "SELECT Produto_Id as Codigo, Descricao FROM ComprasXProdutos " & vbCrLf & _
              " where len(contaclientes) = 0 " & vbCrLf & _
              " Order by Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlCarteira.Items.Add(New ListItem(Funcoes.AlinharEsquerda(Dr("Descricao"), 50, ".") & "-" & Dr("Codigo"), Dr("Codigo")))
        Next

        DdlCarteira.Items.Insert(0, "")
        DdlCarteira.SelectedIndex = 0

    End Sub

    Private Sub CarregarTributos()
        DdlTributo.Items.Clear()

        Sql = "SELECT Encargo_Id as Codigo, Encargo_Id as Descricao FROM Encargos " & vbCrLf & _
              " Order by Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlTributo.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        DdlTributo.Items.Insert(0, "")
        DdlTributo.SelectedIndex = 0
    End Sub

    Private Sub Limpar()
        DdlCarteira.Enabled = True
        DdlTributo.Enabled = True

        DdlCarteira.SelectedIndex = 0
        DdlTributo.SelectedIndex = 0

        lnkNovo.Parent.Visible = True
        lnkExcluir.Parent.Visible = False
    End Sub



    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlCarteira.SelectedValue) Then
            param &= "Carteira: " & DdlCarteira.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(DdlTributo.SelectedValue) Then
            param &= "Encargo: " & DdlTributo.SelectedValue
        End If

        Return param
    End Function

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(DdlCarteira.SelectedValue) Then
            MsgBox(Me.Page, "Carteira não foi informado!", eTitulo.Info)
            Return False
        ElseIf String.IsNullOrWhiteSpace(DdlTributo.SelectedValue) Then
            MsgBox(Me.Page, "Encargo não foi informada!", eTitulo.Info)
            Return False
        End If

        Return True
    End Function

    Protected Sub GridCarteirasXTributos_PreRender(sender As Object, e As EventArgs)
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridCarteirasXTributos.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridCarteirasXTributos.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridCarteirasXTributos.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If

    End Sub

    Protected Sub GridCarteirasXTributos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            DdlCarteira.SelectedValue = GridCarteirasXTributos.SelectedRow.Cells(1).Text()
            DdlTributo.SelectedValue = GridCarteirasXTributos.SelectedRow.Cells(3).Text()

            DdlCarteira.Enabled = False
            DdlTributo.Enabled = False

            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = True

            DdlTributo.Focus()

            ValidarConta()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub ValidarConta()
        Try
            Dim sql As String = " select Carteira, Tributo from ContasAPagar " & vbCrLf &
                                " where Carteira = '" & DdlCarteira.SelectedValue & "'" & vbCrLf &
                                " and Tributo = '" & DdlTributo.SelectedValue & "'" & vbCrLf &
                                " select Carteira, Tributo from ContasAReceber " & vbCrLf &
                                " where Carteira = '" & DdlCarteira.SelectedValue & "'" & vbCrLf &
                                " and Tributo = '" & DdlTributo.SelectedValue & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Razao")

            If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                lnkNovo.Parent.Visible = False
                lnkExcluir.Parent.Visible = False

                MsgBox(Me.Page, "Conta não pode ser alterado/excluída pois foi contabilizada no Razão.")
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CarteirasXTributos", "GRAVAR") Then
                If ValidarCampos() Then
                    Sql = "INSERT Into CarteirasXTributos (Carteira_Id, Tributo_Id) " & vbCrLf & _
                                          " Values('" & DdlCarteira.SelectedValue & "' " & vbCrLf & _
                                          ",'" & UCase(DdlTributo.SelectedValue) & "')" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CarregarCarteirasXTributos()
                        MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
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
            If Funcoes.VerificaPermissao("CarteirasXTributos", "EXCLUIR") Then
                If ValidarCampos() Then
                    Sql = "DELETE FROM CarteirasXTributos" & vbCrLf & _
                                          " WHERE Carteira_Id = '" & DdlCarteira.SelectedValue & "' And Tributo_Id = '" & DdlTributo.SelectedValue & "'" & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) Then
                        Limpar()
                        CarregarCarteirasXTributos()
                        MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
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
            If Funcoes.VerificaPermissao("CarteirasXTributos", "RELATORIO") Then
                Sql = " SELECT     ComprasXProdutos.Produto_Id as Carteira, ComprasXProdutos.Descricao as NomeDaCarteira, Encargos.Encargo_id as Encargo, Encargos.Descricao As NomeDoEncargo" & vbCrLf & _
                      " FROM         ComprasXProdutos INNER JOIN" & vbCrLf & _
                      " CarteirasXTributos ON ComprasXProdutos.Produto_Id = CarteirasXTributos.Carteira_Id INNER JOIN" & vbCrLf & _
                      " Encargos ON CarteirasXTributos.Tributo_ID = Encargos.Encargo_id" & vbCrLf
                DS = Banco.ConsultaDataSet(Sql, "CarteirasXTributos")

                Funcoes.BindReport(Me.Page, DS, "Cr_CarteirasXTributos", eExportType.PDF)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CarteirasXTributos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class