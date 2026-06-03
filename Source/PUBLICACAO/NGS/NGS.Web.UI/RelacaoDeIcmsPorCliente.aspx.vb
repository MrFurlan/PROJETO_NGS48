Imports System.IO
Imports System.Data
Imports System.Drawing
Imports Microsoft.VisualBasic
Imports OfficeOpenXml.Style
Imports OfficeOpenXml
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelacaoDeIcmsPorCliente
    Inherits BasePage

    Dim Sql As String = ""
    Dim Sqla As String = ""
    Dim Row As DataRow
    Dim SqlArray As New ArrayList
    Dim Empresa() As String
    Dim i As Integer = 0
    Dim ii As Integer = 0
    Dim PeriodoInicial As String
    Dim PeriodoFinal As String
    Dim EmpresaNome As String = ""
    Dim EmpresaCNPJ As String = ""
    Dim EmpresaInscricao As String = ""
    Dim EmpresaMestre As String
    Dim sqlAux As String = ""
    Dim ds As New DataSet
    Dim dsAux As New DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Fiscal)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelacaoDeIcmsPorCliente", "ACESSAR") Then
                    txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
                    txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    CargaUnidade()
                    VerificaUnidade()
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Fiscal.aspx")
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

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                DdlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
            Next
        End If

        DdlUnidade.Items.Insert(0, "")
        DdlUnidade.SelectedIndex = 0
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                DdlUnidade.SelectedValue = Dr("AcessoUnidade")
                CargaEmpresas()
                DdlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            Next
        End If
    End Sub

    Private Sub CargaEmpresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresa.Items.Clear()

        HttpContext.Current.Session("EmpresaIcms") = ""
        HttpContext.Current.Session("ProcessoIcms") = 0

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              "  FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              "  Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              "  Where GruposXEmpresas.Empresa_Id = '" & DdlUnidade.SelectedValue & "' " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Clientes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each Dr As DataRow In ds.Tables(0).Rows
                Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
                Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
                DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
            Next
        End If

        DdlEmpresa.Items.Insert(0, "")
        DdlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub DdlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function Validar()
        Dim ok As Boolean = True

        If DdlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a unidade de negócio.")
            ok = False
        End If
        If DdlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            ok = False
        End If
        If txtDataInicial.Text = "" Then
            MsgBox(Me.Page, "Informe o período inicial.")
            ok = False
        End If
        If txtDataFinal.Text = "" Then
            MsgBox(Me.Page, "Informe o período final.")
            ok = False
        End If

        Return ok
    End Function

    Private Sub Limpar()
        txtDataInicial.Text = ""
        txtDataFinal.Text = ""
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidade.Enabled = False
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Private Function getDataSet() As DataSet
        Empresa = DdlEmpresa.SelectedValue.Split("-")
        EmpresaMestre = Left(Empresa(0), 8)
        PeriodoInicial = Format(DateValue(txtDataInicial.Text), "yyyy-MM-dd")
        PeriodoFinal = Format(DateValue(txtDataFinal.Text), "yyyy-MM-dd")

        Sql = "     SELECT     Clientes.Cliente_Id AS Cliente, Clientes.Nome AS NomeDoCliente, Produtos.Produto_Id AS Produto, Produtos.Nome AS NomeDoProduto," & vbCrLf & _
              "                       NotasFiscaisXEncargos.Serie_Id AS Serie, NotasFiscaisXEncargos.Nota_Id AS Nota, NotasFiscais.Movimento, NotasFiscaisXEncargos.Base, NotasFiscaisXEncargos.Percentual," & vbCrLf & _
              "                       NotasFiscaisXEncargos.Valor" & vbCrLf & _
              "     FROM              NotasFiscais INNER JOIN" & vbCrLf & _
              "                       NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id AND" & vbCrLf & _
              "                       NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id AND NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id AND" & vbCrLf & _
              "                       NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id AND NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND" & vbCrLf & _
              "                       NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id INNER JOIN" & vbCrLf & _
              "                       NotasFiscaisXEncargos ON NotasFiscaisXItens.Empresa_Id = NotasFiscaisXEncargos.Empresa_Id AND" & vbCrLf & _
              "                       NotasFiscaisXItens.EndEmpresa_Id = NotasFiscaisXEncargos.EndEmpresa_Id AND NotasFiscaisXItens.Cliente_Id = NotasFiscaisXEncargos.Cliente_Id AND" & vbCrLf & _
              "                       NotasFiscaisXItens.EndCliente_Id = NotasFiscaisXEncargos.EndCliente_Id AND NotasFiscaisXItens.EntradaSaida_Id = NotasFiscaisXEncargos.EntradaSaida_Id AND" & vbCrLf & _
              "                       NotasFiscaisXItens.Serie_Id = NotasFiscaisXEncargos.Serie_Id AND NotasFiscaisXItens.Nota_Id = NotasFiscaisXEncargos.Nota_Id AND" & vbCrLf & _
              "                       NotasFiscaisXItens.Produto_Id = NotasFiscaisXEncargos.Produto_Id AND NotasFiscaisXItens.CFOP_Id = NotasFiscaisXEncargos.CFOP_Id AND" & vbCrLf & _
              "                       NotasFiscaisXItens.Sequencia_Id = NotasFiscaisXEncargos.Sequencia_id INNER JOIN" & vbCrLf & _
              "                       Clientes ON NotasFiscais.Cliente_Id = Clientes.Cliente_Id AND NotasFiscais.EndCliente_Id = Clientes.Endereco_Id INNER JOIN" & vbCrLf & _
              "                       Produtos ON NotasFiscaisXItens.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
              "       Where           NotasFiscais.Empresa_Id = '" & Empresa(0) & "'" & vbCrLf & _
              "                   And (NotasFiscais.Movimento  between '" & PeriodoInicial & "' AND '" & PeriodoFinal & "') AND (NotasFiscais.Situacao = 1)  " & vbCrLf & _
              "                   AND (NotasFiscais.EntradaSaida_Id = 'S') AND (NotasFiscaisXEncargos.Encargo_Id = 'ICMS')" & vbCrLf & _
              "                   And NotasFiscaisXEncargos.Valor  > 0" & vbCrLf & _
              "     Order By  Clientes.Nome" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Notas")
        Return ds
    End Function

    Private Sub GerarExcel()
        Try
            If Funcoes.VerificaPermissao("RelacaoDeIcmsPorCliente", "RELATORIO") Then
                Dim ds As DataSet = getDataSet()
                Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

                If File.Exists(fileName) Then
                    File.Delete(fileName)
                End If

                Using arquivo As New FileStream(fileName, FileMode.CreateNew)
                    Using package As New ExcelPackage(arquivo)
                        'criando planilha títulos
                        Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("RELAÇÃO DE ICMS POR CLIENTE")

                        'criando linha com o cabeçalho da planilha
                        Dim rowIndex As Integer = 1
                        Dim columnIndex As Integer = 1

                        'criando linha com o cabeçalho da planilha
                        For Each col As DataColumn In ds.Tables(0).Columns
                            worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                            columnIndex += 1
                        Next

                        'criando auto filtro na planilha
                        worksheet.Cells("A1:J" & rowIndex).AutoFilter = True

                        'aplicando formatação nas células do cabeçalho
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                            range.Style.Font.Bold = True
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                            range.Style.Font.Color.SetColor(Color.White)
                        End Using
                        rowIndex += 1

                        'criando conteúdo da planilha com os dados do dataset
                        For Each row As DataRow In ds.Tables(0).Rows
                            columnIndex = 1
                            For Each col As DataColumn In ds.Tables(0).Columns
                                worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                                columnIndex += 1
                            Next

                            'aplicando formatação decimal nos campos de valores
                            worksheet.Cells(String.Format("H{0}:H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("I{0}:I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                            worksheet.Cells(String.Format("J{0}:J{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                            'aplicando formatação nas células do conteúdo
                            If rowIndex Mod 2 = 0 Then
                                Using range = worksheet.Cells(rowIndex, 1, rowIndex, ds.Tables(0).Columns.Count)
                                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                                End Using
                            End If
                            rowIndex += 1
                        Next

                        'setando autofit nas células da planilha
                        worksheet.Cells.AutoFitColumns(0)

                        'congelando primeira linha (cabeçalho)
                        worksheet.View.FreezePanes(2, 1)

                        'salvando planilha do excel
                        package.Save()
                    End Using
                End Using

                'download do arquivo pelo browser
                Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            GerarExcel()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelacaoDeIcmsPorCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class