Imports System.Drawing
Imports System.IO
Imports NGS.Lib.Negocio
Imports OfficeOpenXml
Imports OfficeOpenXml.Style

Public Class CnpjsSefaz
    Inherits BasePage

    Private objCliente As New [Lib].Negocio.Cliente
    Private objMunicipio As [Lib].Negocio.Municipio

    Private Property ListCnpjSefaz As ListCnpjSefaz
        Get
            Return Session("ListCnpjSefaz" + HID.Value)
        End Get
        Set(ByVal value As ListCnpjSefaz)
            Session("ListCnpjSefaz" + HID.Value) = value
        End Set
    End Property

    Private Property objCnpjSefaz() As CnpjSefaz
        Get
            Return CType(Session("objCnpjSefaz" + HID.Value), CnpjSefaz)
        End Get
        Set(ByVal value As CnpjSefaz)
            Session("objCnpjSefaz" + HID.Value) = value
        End Set
    End Property

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("CnpjsSefaz", "ACESSAR") Then
                    ddl.Carregar(ddlEstado, CarregarDDL.Tabela.Estados, "", True)
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaCodMunicipios.SetarHID(HID.Value)
                    CarregarCnpjsSefaz()
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

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objMunicipioCli" & HID.Value) Is Nothing Then
                objMunicipio = CType(Session("objMunicipioCli" & HID.Value), [Lib].Negocio.Municipio)
                txtCidade.Text = objMunicipio.CodigoMunicipio
                Session.Remove("objMunicipioCli" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub CarregarCnpjsSefaz()
        If Funcoes.VerificaPermissao("CnpjsSefaz", "LEITURA") Then
            If ListCnpjSefaz Is Nothing Then
                ListCnpjSefaz = New ListCnpjSefaz(True)
            End If

            GridCnpjSefaz.DataSource = ListCnpjSefaz
            GridCnpjSefaz.DataBind()
        Else
            MsgBox(Me.Page, "Usuário sem permissão para leitura registro.", eTitulo.Info)
        End If
    End Sub

    Private Sub SalvarSessionCnpjSefaz()
        ViewState("objCnpjSefaz") = objCnpjSefaz
    End Sub

    Private Sub RecuperarSessionCnpjSefaz()
        objCnpjSefaz = CType(ViewState("objCnpjSefaz"), [Lib].Negocio.CnpjSefaz)
    End Sub

    Private Sub Limpar()
        Try
            txtCnpj.Text = String.Empty
            txtCnpj.Enabled = True
            txtCidade.Text = String.Empty
            ddlEstado.SelectedIndex = 0
            lnkNovo.Parent.Visible = True
            lnkAtualizar.Parent.Visible = False
            lnkExcluir.Parent.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function ValidaCampos(Optional ByVal Novo As Boolean = True) As Boolean
        If String.IsNullOrWhiteSpace(txtCnpj.Text) Then
            MsgBox(Me.Page, "Informe o CNPJ.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtCidade.Text) Then
            MsgBox(Me.Page, "Informe a cidade.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(ddlEstado.SelectedValue) Then
            MsgBox(Me.Page, "Selecione Ativo ou Estado.")
            Return False
        End If
        Return True
    End Function

    Private Sub EmitirRelatorio()
        ListCnpjSefaz = New ListCnpjSefaz(True)
        Dim obj As Usuario = UsuarioServidor.Carregar(Session("ssNomeUsuario"))
        Dim dt As DataTable = New DataTable()

        Dim ds As DataSet = ListCnpjSefaz

        dt = ds.Tables(0)

        If dt.Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum registro encontrado para essa seleção.")
            Exit Sub
        End If

        Dim fileName As String = Server.MapPath("~/PlanilhasExcel/" & Funcoes.GeraNomeArquivo & ".xlsx")

        If File.Exists(fileName) Then
            File.Delete(fileName)
        End If

        Using arquivo As New FileStream(fileName, FileMode.CreateNew)
            Using package As New ExcelPackage(arquivo)

                'criando planilha e título
                Dim worksheet As ExcelWorksheet = package.Workbook.Worksheets.Add("Relatório de produtos.")

                'criando linha com o cabeçalho da planilha
                Dim rowIndex As Integer = 1
                Dim columnIndex As Integer = 1

                'criando linha que informa o nome da empresa e o cnpj
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0} - {1}", obj.Empresa.Nome, Funcoes.FormatarCpfCnpj(obj.Empresa.Codigo))
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa a cidade e o estado da empresa
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}/{1}", obj.Empresa.Cidade, obj.Empresa.CodigoEstado)
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o título do relatório
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Red)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "PRODUTOS")
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando linha que informa o período selecionado na página
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Bold = True
                worksheet.Cells(rowIndex, columnIndex).Style.Font.Color.SetColor(Color.Black)
                worksheet.Cells(rowIndex, columnIndex).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                worksheet.Cells(rowIndex, columnIndex).Value = String.Format("{0}", "Data : " & Now().ToString("dd-MM-yyyy"))
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Merge = True
                worksheet.Cells(String.Format("A{0}:D{0}", rowIndex)).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left
                rowIndex += 1

                'criando cabeçalho da planilha com os dados do dataset
                'criando linha com o cabeçalho da planilha
                For Each col As DataColumn In dt.Columns
                    worksheet.Cells(rowIndex, columnIndex).Value = col.ColumnName
                    columnIndex += 1
                Next

                'criando auto filtro na planilha
                worksheet.Cells("A5:F" & rowIndex).AutoFilter = True

                'formatando células numéricas
                worksheet.Cells(String.Format("H{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"
                worksheet.Cells(String.Format("I{0}", rowIndex)).Style.Numberformat.Format = "#,##0.00_ ;[Red]-#,##0.00"

                'aplicando formatação nas células do cabeçalho
                Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                    range.Style.Font.Bold = True
                    range.Style.Fill.PatternType = ExcelFillStyle.Solid
                    range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(79, 129, 189))
                    range.Style.Font.Color.SetColor(Color.White)
                End Using
                rowIndex += 1

                ' criando conteúdo da planilha com os dados do dataset
                For Each row As DataRow In dt.Rows
                    columnIndex = 1

                    For Each col As DataColumn In dt.Columns
                        worksheet.Cells(rowIndex, columnIndex).Value = row(col.ColumnName)
                        columnIndex += 1
                    Next

                    'formatando células datas
                    worksheet.Cells(String.Format("A{0}", rowIndex)).Style.Numberformat.Format = "dd/MM/yyyy"

                    'aplicando formatação nas células do conteúdo
                    If rowIndex Mod 2 = 0 Then
                        Using range = worksheet.Cells(rowIndex, 1, rowIndex, dt.Columns.Count)
                            range.Style.Fill.PatternType = ExcelFillStyle.Solid
                            range.Style.Fill.BackgroundColor.SetColor(Color.FromArgb(153, 204, 255))
                        End Using
                    End If
                    rowIndex += 1
                Next

                'setando autofit nas células da planilha
                worksheet.Cells.AutoFitColumns(0)

                'congelando quinta linha (cabeçalho)
                worksheet.View.FreezePanes(6, 1)

                'salvando planilha do excel
                package.Save()
            End Using
        End Using
        Funcoes.AbrirExcel(Me.Page, "PlanilhasExcel/" & Path.GetFileName(fileName))
        'download do arquivo pelo browser
        'Download(fileName)
    End Sub

    Protected Sub ddlEstado_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEstado.SelectedIndexChanged
        Try
            If Not String.IsNullOrWhiteSpace(ddlEstado.SelectedValue) Then
                txtCidade.Text = String.Empty
                Session("ssUF" & HID.Value) = ddlEstado.SelectedValue
                ucConsultaCodMunicipios.Limpar()
                Popup.ConsultaDeMunicipios(Me.Page, "objMunicipioCli" & HID.Value)
            Else
                txtCidade.Text = String.Empty
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub GridCnpjSefaz_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim i = ListCnpjSefaz.FindIndex(Function(s) s.Cnpj = GridCnpjSefaz.SelectedRow.Cells(1).Text())

            txtCnpj.Text = ListCnpjSefaz(i).Cnpj
            txtCidade.Text = ListCnpjSefaz(i).Cidade
            ddlEstado.SelectedValue = ListCnpjSefaz(i).Estado
            txtCnpj.Enabled = False
            txtCidade.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridCnpjSefaz_PreRender(sender As Object, e As EventArgs) Handles GridCnpjSefaz.PreRender
        ' You only need the following 2 lines of code if you are not 
        ' using an ObjectDataSource of SqlDataSource


        If GridCnpjSefaz.Rows.Count > 0 Then
            'This replaces <td> with <th> and adds the scope attribute
            GridCnpjSefaz.UseAccessibleHeader = True

            'This will add the <thead> and <tbody> elements
            GridCnpjSefaz.HeaderRow.TableSection = TableRowSection.TableHeader

            'This adds the <tfoot> element. 
            'Remove if you don't have a footer row
            'GridEstados.FooterRow.TableSection = TableRowSection.TableFooter
        End If
    End Sub

    Protected Sub txtCnpj_TextChanged(sender As Object, e As EventArgs) Handles txtCnpj.TextChanged
        Try
            'PESQUISA NA RECEITA FEDERAL DADOS DA EMPRESA
            If (txtCnpj.Text.RemoveMask.Length = 14) Then
                Dim novoCliente As New Cliente(txtCnpj.Text.RemoveMask)

                If novoCliente.Nome.Length = 0 Then
                    MsgBox(Me.Page, "A Receita Federal não disponibilizou as informações desse CNPJ.")
                    Exit Sub
                End If
                Limpar()
                objCliente = novoCliente
                txtCnpj.Text = objCliente.Codigo
                txtCidade.Text = objCliente.Cidade
                ddlEstado.SelectedValue = objCliente.CodigoEstado
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CnpjsSefaz", "GRAVAR") Then
                If ValidaCampos() Then
                    objCnpjSefaz = New CnpjSefaz()
                    objCnpjSefaz.Cnpj = txtCnpj.Text
                    objCnpjSefaz.Cidade = txtCidade.Text
                    objCnpjSefaz.Estado = ddlEstado.SelectedValue
                    objCnpjSefaz.IUD = "I"

                    If objCnpjSefaz.Salvar() Then
                        ListCnpjSefaz.Add(objCnpjSefaz)
                        MsgBox(Me.Page, "Registro Gravado com Sucesso.", eTitulo.Sucess)
                    End If
                    CarregarCnpjsSefaz()
                    Limpar()
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
            If Funcoes.VerificaPermissao("CnpjsSefaz", "ALTERAR") Then
                If ValidaCampos(False) Then
                    Dim i As Integer = ListCnpjSefaz.FindIndex(Function(s) s.Cnpj = txtCnpj.Text)
                    ListCnpjSefaz(i).Cidade = txtCidade.Text
                    ListCnpjSefaz(i).Estado = ddlEstado.SelectedValue
                    ListCnpjSefaz(i).IUD = "U"

                    If ListCnpjSefaz(i).Salvar Then
                        MsgBox(Me.Page, "Registro Alterado com Sucesso.", eTitulo.Sucess)
                    End If
                    CarregarCnpjsSefaz()
                    Limpar()
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
            If Funcoes.VerificaPermissao("CnpjsSefaz", "EXCLUIR") Then
                If String.IsNullOrWhiteSpace(txtCnpj.Text) Then
                    MsgBox(Me.Page, "Selecione um registro para Excluir.", eTitulo.Info)
                Else
                    Dim i As Integer = ListCnpjSefaz.FindIndex(Function(s) s.Cnpj = txtCnpj.Text)
                    ListCnpjSefaz(i).IUD = "D"

                    If ListCnpjSefaz(i).Salvar() Then
                        ListCnpjSefaz.Remove(objCnpjSefaz)
                        MsgBox(Me.Page, "Registro Excluído com Sucesso.", eTitulo.Sucess)
                    End If
                    CarregarCnpjsSefaz()
                    Limpar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRelatorio.Click
        Try
            EmitirRelatorio()
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
End Class