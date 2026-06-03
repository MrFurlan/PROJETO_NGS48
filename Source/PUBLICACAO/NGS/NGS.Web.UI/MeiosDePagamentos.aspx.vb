Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class MeiosDePagamentos
    Inherits BasePage
    Dim objMeiosPagamento As MeiosPagamento

    Protected Sub Page_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("MeiosDePagamentos", "ACESSAR") AndAlso Request.Url.Host.ToUpper.Contains("LOCALHOST") Then
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub salvarobj()
        Session("objMeiosDePagamentos") = objMeiosPagamento
    End Sub

    Private Sub carregarobj()
        objMeiosPagamento = Session("objMeiosDePagamentos")
    End Sub

    Private Sub Limpar()
        txtCodigo.Text = String.Empty
        txtDescricao.Text = String.Empty
        txtCodigo.Enabled = True

        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        lnkExcluir.Parent.Visible = False

        atualizarGrid()
    End Sub

    Private Sub gravar(IUD As String)
        Try
            carregarobj()
            If IUD = "I" Then
                If IsNumeric(txtCodigo.Text) Then
                    objMeiosPagamento = New MeiosPagamento
                    objMeiosPagamento.IUD = "I"
                    objMeiosPagamento.Codigo = txtCodigo.Text
                    objMeiosPagamento.Descricao = txtDescricao.Text
                    objMeiosPagamento.Ativo = IIf(RadSim.Checked, True, False)
                    objMeiosPagamento.UsuarioInclusao = Session("ssNomeUsuario")
                End If
            ElseIf IUD = "U" Then
                objMeiosPagamento.IUD = "U"
                objMeiosPagamento.Descricao = txtDescricao.Text
                objMeiosPagamento.Ativo = IIf(RadSim.Checked, True, False)
                objMeiosPagamento.UsuarioAlteracao = Session("ssNomeUsuario")
            ElseIf IUD = "D" Then
                objMeiosPagamento = New MeiosPagamento(txtCodigo.Text)
                objMeiosPagamento.IUD = "D"
                objMeiosPagamento.UsuarioAlteracao = Session("ssNomeUsuario")
            Else
                MsgBox(Me.Page, "Valor invalido para IUD.")
                Exit Sub
            End If
            objMeiosPagamento.salvar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub relatorio()
        If GridMeiosDePagamentos Is Nothing OrElse Not GridMeiosDePagamentos.Rows.Count > 0 Then
            MsgBox(Me.Page, "Nenhum resultado Encontrado!")
        Else
            Dim rowIndex As Integer = 1
            Dim columnIndex As Integer = 1

            Dim fileName As String = HttpContext.Current.Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

            If File.Exists(fileName) Then File.Delete(fileName)

            Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                Using package As New ExcelPackage(arquivo)

                    'criando aba da planilha.
                    Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Meios de Pagamento")

                    'Inserindo o Cabeçalho.
                    Dim objEmpresa As Cliente = New Cliente(Session("ssEmpresa"), Session("ssEndEmpresa"))
                    worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Nome.ToUpper
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    rowIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = objEmpresa.Cidade.ToUpper & "/" & objEmpresa.Estado.Descricao.ToUpper
                    rowIndex += 1
                    worksheet.Cells(rowIndex, columnIndex).Value = "CNPJ: " & Funcoes.FormatarCpfCnpj(objEmpresa.Codigo) & "-" & objEmpresa.Endereco
                    rowIndex += 1

                    worksheet.Cells(rowIndex, columnIndex).Value = "Meios de Pagamento"
                    worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                    rowIndex += 1

                    'Inserindo informacoes adicionaisW
                    worksheet.Cells(rowIndex, columnIndex).Value = "Emissão: " & DateTime.Now
                    rowIndex += 1

                    For column = 1 To GridMeiosDePagamentos.Columns.Count - 1
                        worksheet.Cells(rowIndex, columnIndex).Value = GridMeiosDePagamentos.Columns(column)
                        columnIndex += 1
                    Next

                    rowIndex += 1

                    For row = 0 To GridMeiosDePagamentos.Rows.Count - 1
                        For column = 1 To columnIndex
                            worksheet.Cells(rowIndex, column).Value = GridMeiosDePagamentos.Rows(row).Cells(column).Text.ToString()
                        Next
                        rowIndex += 1
                    Next

                    rowIndex += 1

                    'setando autofit nas células da planilha
                    worksheet.Cells.AutoFitColumns(0)

                    'congelando primeira linham
                    worksheet.View.FreezePanes(9, 1)

                    worksheet.Column(1).Width = 20

                    'salvando planilha do excel
                    package.Save()
                End Using
            End Using

            'download do arquivo pelo browser
            Funcoes.AbrirExcel(Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        End If
    End Sub

    Private Sub atualizarGrid()
        Dim objListMeiosPagamento As New ListMeiosPagamento()

        If Not objListMeiosPagamento.Count > 0 Then
            GridMeiosDePagamentos.DataSource = Nothing
        Else
            GridMeiosDePagamentos.DataSource = objListMeiosPagamento.ToArray
        End If

        GridMeiosDePagamentos.DataBind()
    End Sub

    Private Sub selecionar(codEncargo As Integer)
        Limpar()
        objMeiosPagamento = New MeiosPagamento(codEncargo)
        salvarobj()

        If objMeiosPagamento Is Nothing Then
            MsgBox(Me.Page, "Encargo " & codEncargo & " não foi encontrado.")
        Else
            If objMeiosPagamento.Ativo = "S" Then
                RadSim.Checked = True
                RadNao.Checked = False
            Else
                RadSim.Checked = False
                RadNao.Checked = True
            End If
            txtCodigo.Enabled = False
            lnkNovo.Parent.Visible = False
            lnkExcluir.Parent.Visible = True
            lnkAtualizar.Parent.Visible = True
            txtCodigo.Text = objMeiosPagamento.Codigo
            txtDescricao.Text = objMeiosPagamento.Descricao
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As System.EventArgs) Handles lnkNovo.Click
        gravar("I")
        Limpar()
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As System.EventArgs) Handles lnkAtualizar.Click
        gravar("U")
        Limpar()
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As System.EventArgs) Handles lnkExcluir.Click
        gravar("D")
        Limpar()
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As System.EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As System.EventArgs) Handles lnkRelatorio.Click
        Try
            relatorio()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridMeiosDePagamentos_SelectedIndexChanged(sender As Object, e As System.EventArgs) Handles GridMeiosDePagamentos.SelectedIndexChanged
        Try
            Dim codEncargo As Integer = CInt(GridMeiosDePagamentos.Rows(GridMeiosDePagamentos.SelectedIndex).Cells(1).Text)
            selecionar(codEncargo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As System.EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "MeioDePagamentos")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class