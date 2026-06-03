Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Pagamentos
    Inherits BasePage

    Dim SqlArray As New ArrayList
    Dim Sql As String = ""
    Dim ds As DataSet
    Dim Row As DataRow

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Pagamentos", "ACESSAR") Then
                    Limpar()
                    CargaPagamentos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaPagamentos()
        If Funcoes.VerificaPermissao("Pagamentos", "LEITURA") Then
            Sql = "Select Pagamento_Id as Codigo, " & vbCrLf &
                  "Descricao, Parcelas, isnull(AVista,0) as Avista, " & vbCrLf &
                  "isnull(Antecipado,0) as Antecipado, ISNULL(VencimentoPedido,0) AS VencimentoPedido, " & vbCrLf &
                  "Case" & vbCrLf &
                  "   When isnull(AVista,0) = 0" & vbCrLf &
                  "      then 2" & vbCrLf &
                  "	     else 1" & vbCrLf &
                  "   End As Ordem" & vbCrLf &
                  "into #Pagamentos" & vbCrLf &
                  "from Pagamentos " & vbCrLf &
                  "" & vbCrLf &
                  "select Codigo, Descricao, Parcelas, Avista, Antecipado, VencimentoPedido from #Pagamentos" & vbCrLf &
                  "order by Ordem, Codigo" & vbCrLf

            ds = Banco.ConsultaDataSet(Sql, "Pagamentos")

            GridPagamentos.DataSource = ds
            GridPagamentos.DataBind()

            Dim n As Integer

            For Each cp In ds.Tables(0).Rows
                n = cp("Codigo")
            Next

            n += 1

            txtCodigo.Text = n

        Else
            MsgBox(Me.Page, Session("ssMessage"))
        End If
    End Sub

    Private Sub GravaPagamentos()
        If Funcoes.VerificaPermissao("Pagamentos", "GRAVAR") Then
            Dim i As Integer

            If Not ExcluirPagamentos() Then
                Exit Sub
            End If

            Dim dias As Integer

            While i < GridParcelas.Rows.Count
                Dim diatxt = CType(GridParcelas.Rows(i).FindControl("txtDias"), TextBox)

                If Not IsNumeric(diatxt.Text()) Or Trim(diatxt.Text()) = "" Then
                    MsgBox(Me.Page, "Vencimento da " & i + 1 & "° parcela não informada/invalida.")
                    Exit Sub
                Else
                    dias = CInt(diatxt.Text())

                    Sql = "INSERT INTO PagamentosXParcelas (Pagamento_Id, Sequencia_Id, Dias)" & vbCrLf & _
                          " VALUES (" & CInt(txtCodigo.Text) & ", " & i + 1 & ", " & dias & ");"

                    SqlArray.Add(Sql)
                    i += 1
                End If
            End While

            Sql = "INSERT INTO Pagamentos (Pagamento_Id, Descricao, Parcelas, AVista, Antecipado, VencimentoPedido)" & vbCrLf
            Sql &= "VALUES (" & CInt(txtCodigo.Text) & ",'" & txtDescricao.Text & "'" & "," & CInt(txtParcelas.Text) & "," & CByte(chkAVista.Checked) & "," & CByte(chkAntecipado.Checked) & "," & CByte(chkVencimentoPedido.Checked) & ");" & vbCrLf
            SqlArray.Add(Sql)

            If Not Banco.GravaBanco(SqlArray) Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                MsgBox(Me.Page, "Registro gravado/atualizado com Sucesso.", eTitulo.Sucess)
                Limpar()
            End If
        Else
        MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
        End If
    End Sub

    Protected Sub GridOperacoes_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer = 0

        txtCodigo.Text = GridPagamentos.SelectedRow.Cells(1).Text()
        txtDescricao.Text = GridPagamentos.SelectedRow.Cells(2).Text()
        txtParcelas.Text = GridPagamentos.SelectedRow.Cells(3).Text()
        chkAVista.Checked = CType(GridPagamentos.SelectedRow.FindControl("checkbox1"), CheckBox).Checked
        chkAntecipado.Checked = CType(GridPagamentos.SelectedRow.FindControl("checkbox2"), CheckBox).Checked
        chkVencimentoPedido.Checked = CType(GridPagamentos.SelectedRow.FindControl("checkbox3"), CheckBox).Checked

        Sql = "Select Sequencia_Id as Sequencia from PagamentosXParcelas " & vbCrLf & _
              "  WHERE Pagamento_Id = " & txtCodigo.Text & " order by Sequencia_Id" & vbCrLf
        ds = Banco.ConsultaDataSet(Sql, "Pagamentos")

        GridParcelas.DataSource = Banco.ConsultaDataSet(Sql, "Pagamentos")
        GridParcelas.DataBind()

        Sql = "Select Dias from PagamentosXParcelas " & vbCrLf & _
              "  WHERE Pagamento_Id = " & txtCodigo.Text & " order by Sequencia_Id" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Parcelas").Tables(0).Rows
            If Not IsDBNull(Dr("Dias")) Then
                CType(GridParcelas.Rows(i).FindControl("txtDias"), TextBox).Text = Dr("Dias")
            End If
            i += 1
        Next

        lnkNovo.Parent.Visible = False
        lnkAtualizar.Parent.Visible = True

        If Session("ssNomeUsuario") = "FURLAN" Then
            lnkExcluir.Parent.Visible = True
        End If

    End Sub

    Private Sub Limpar()
        txtCodigo.Text = ""
        txtDescricao.Text = ""
        txtParcelas.Text = ""
        chkAVista.Checked = False
        chkAntecipado.Checked = False
        chkVencimentoPedido.Checked = False
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False
        GridParcelas.DataSource = Nothing
        GridParcelas.DataBind()
    End Sub

    Private Function ExcluirPagamentos()
        If Funcoes.VerificaPermissao("Pagamentos", "EXCLUIR") Then


            Sql = "Select CondicaoPagamento from pedidos " & vbCrLf &
                  "WHERE CondicaoPagamento = " & txtCodigo.Text

            ds = Banco.ConsultaDataSet(Sql, "Pagamentos")

            If ds.Tables(0).Rows.Count > 0 Then
                MsgBox(Me.Page, "Condição de Pagamento não pode ser removida porque já foi utilizada no Pedido.")
                Return False
            End If

            Sql = " DELETE Pagamentos " & vbCrLf &
                  "  WHERE Pagamento_Id = " & txtCodigo.Text & vbCrLf
            SqlArray.Add(Sql)

            Sql = " DELETE PagamentosXParcelas " & vbCrLf &
                  "  WHERE Pagamento_Id = " & txtCodigo.Text & vbCrLf
            SqlArray.Add(Sql)

            If Banco.GravaBanco(SqlArray) Then
                Return True
            Else
                MsgBox(Me.Page, Session("ssMessage"))
                Return False
            End If
        Else
            MsgBox(Me.Page, "Usuário sem permissão para excluir o registro.")
            Return False
        End If
    End Function

    Private Sub CriaParcelas()
        Dim i As Integer = 0
        Dim dsParcelas As New DataSet
        Dim dtParcelas As DataTable

        dtParcelas = New DataTable("Parcelas")
        dtParcelas.Columns.Add("Sequencia", Type.GetType("System.String")).DefaultValue = ""

        While i < CInt(txtParcelas.Text)
            Row = dtParcelas.NewRow()
            Row("Sequencia") = i + 1
            dtParcelas.Rows.Add(Row)
            i += 1
        End While

        GridParcelas.DataSource = dtParcelas
        GridParcelas.DataBind()

    End Sub

    Protected Sub cmdParcelas_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim nParcelas As String = Trim(txtParcelas.Text)

            If nParcelas.Length = 0 Then
                MsgBox(Me.Page, "Quantidade de Parcelas não foi informada.")
                Exit Sub
            End If

            If CInt(nParcelas) < 1 Then
                MsgBox(Me.Page, "Quantidade de Parcelas não foi informada.")
                Exit Sub
            End If


            If chkVencimentoPedido.Checked AndAlso CInt(nParcelas) > 1 Then
                MsgBox(Me.Page, "Para Vencimento no Pedido quantidade de Parcelas deve ser 1(UMA).")
                Exit Sub
            End If

            CriaParcelas()
            txtParcelas.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Pagamentos", "GRAVAR") Then

                Dim nParcelas As String = Trim(txtParcelas.Text)

                If chkVencimentoPedido.Checked AndAlso CInt(nParcelas) > 1 Then
                    MsgBox(Me.Page, "Para Vencimento no Pedido quantidade de Parcelas deve ser 1(UMA).")
                    Exit Sub
                End If

                GravaPagamentos()
                CargaPagamentos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("Pagamentos", "ALTERAR") Then
                GravaPagamentos()
                Limpar()
                CargaPagamentos()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("Pagamentos", "EXCLUIR") Then
                If ExcluirPagamentos() Then
                    MsgBox(Me.Page, "Registro excluido com Sucesso.", eTitulo.Sucess)
                    Limpar()
                    CargaPagamentos()
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
            CargaPagamentos()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim Ds As DataSet
            If Funcoes.VerificaPermissao("Pagamentos", "RELATORIO") Then
                Sql = "SELECT Pagamento_Id as Codigo, Descricao, Parcelas " & vbCrLf & _
                      "FROM  Pagamentos" & vbCrLf
                Ds = Banco.ConsultaDataSet(Sql, "Pagamentos")

                Dim parameters = New Dictionary(Of String, Object)()

                Funcoes.BindReport(Me.Page, Ds, "Cr_Pagamentos", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Pagamentos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class