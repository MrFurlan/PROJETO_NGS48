Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Reflection

Public Class Contratos
    Inherits BasePage

    Public Property ListaDeContratos As [Lib].Negocio.ListContrato
        Get
            Return CType(Session("ListaContratos" & HID.Value), ListContrato)
        End Get
        Set(ByVal value As [Lib].Negocio.ListContrato)
            Session("ListaContratos" & HID.Value) = value
        End Set
    End Property

    Public Property objContrato As [Lib].Negocio.Contrato
        Get
            Return CType(Session("objContrato" & HID.Value), Contrato)
        End Get
        Set(ByVal value As [Lib].Negocio.Contrato)
            Session("objContrato" & HID.Value) = value
        End Set
    End Property

    Public Property objArquivoPostado As Byte()
        Get
            Return Session("objArquivoPostado" & HID.Value)
        End Get
        Set(ByVal value As Byte())
            Session("objArquivoPostado" & HID.Value) = value
        End Set
    End Property


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Contratos", "ACESSAR") Then
                    Limpar()
                    CarregarContratos()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        HID.Value = Guid.NewGuid().ToString
        txtCodigo.Text = String.Empty
        txtNomeDoArquivo.Text = String.Empty
        txtDescricao.Text = String.Empty
        lnkNovo.Parent.Visible = True
        lnkAtualizar.Parent.Visible = False
        'lnkExcluir.Parent.Visible = False
        Session.Remove("objArquivoPostado" & HID.Value)
        Session.Remove("objContrato" & HID.Value)
        Session.Remove("ListaContratos" & HID.Value)
        CarregarContratos()
    End Sub

    Private Sub CarregarContratos()
        If Funcoes.VerificaPermissao("Contratos", "LEITURA") Then
            ListaDeContratos = New ListContrato(True)
            gridContratos.DataSource = ListaDeContratos
            gridContratos.DataBind()
        End If
    End Sub

    Private Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDescricao.Text) Then
            MsgBox(Me.Page, "Digite a Descrição", eTitulo.Info)
            txtDescricao.Focus()
            Return False
        ElseIf objArquivoPostado Is Nothing Then
            MsgBox(Me.Page, "Selecione um arquivo.", eTitulo.Info)
            txtNomeDoArquivo.Focus()
            Return False
        Else
            Return True
        End If
    End Function

    Protected Sub btnAdicionar_Click(sender As Object, e As EventArgs) Handles btnAdicionar.Click
        If fupArquivo.HasFile Then
            Dim NomeDoArquivo As String = Path.GetFileName(fupArquivo.PostedFile.FileName)
            Dim TamanhoDoArquivo As Long = fupArquivo.PostedFile.ContentLength
            Dim extensao As String = Path.GetExtension(NomeDoArquivo)
            Dim contentType As String = String.Empty

            If Not extensao.ToLower.Equals(".docx") Then
                MsgBox(Me.Page, "São permitidos apenas documentos de extensão docx.")
                Exit Sub
            End If
            objArquivoPostado = fupArquivo.FileBytes
            txtNomeDoArquivo.Text = NomeDoArquivo
        Else
            MsgBox(Me.Page, "Selecione um arquivo.")
        End If
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("Contratos", "GRAVAR") Then
                If ValidarCampos() Then

                    If objContrato Is Nothing Then
                        objContrato = New Contrato()
                    End If
                    objContrato.IUD = "I"
                    objContrato.Descricao = txtDescricao.Text()
                    objContrato.Arquivo = objArquivoPostado
                    If Not objContrato.Salvar() Then
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        Limpar()
                        'CarregarContratos()
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            'If Funcoes.VerificaPermissao("Marcas", "ALTERAR") Then

            Dim Pos = ListaDeContratos.FindIndex(Function(s) s.Codigo = New Guid(txtCodigo.Text))

            ListaDeContratos(Pos).IUD = "U"
            ListaDeContratos(Pos).Descricao = txtDescricao.Text
            If objArquivoPostado IsNot Nothing Then
                ListaDeContratos(Pos).Arquivo = objArquivoPostado
            End If

            If Not ListaDeContratos(Pos).Salvar Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                'CarregarContratos()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    'Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
    '    Try
    '        'If Funcoes.VerificaPermissao("Marcas", "EXCLUIR") Then
    '        Dim Pos = ListaDeContratos.FindIndex(Function(s) s.Codigo = New Guid(txtCodigo.Text))

    '        ListaDeContratos(Pos).IUD = "D"
    '        If Not ListaDeContratos(Pos).Salvar Then
    '            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
    '        Else
    '            Limpar()
    '            'CarregarContratos()
    '        End If
    '        'Else
    '        'MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
    '        'End If

    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message, eTitulo.Erro)
    '    End Try
    'End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            'If Funcoes.VerificaPermissao("Marcas", "RELATORIO") Then
            Dim Sql As String = "SELECT Contrato_Id AS codigo, Descricao" & vbCrLf & _
                                "  FROM Contrato " & vbCrLf & _
                                " ORDER BY Contrato_Id " & vbCrLf

            Dim DS As DataSet = Banco.ConsultaDataSet(Sql, "Tabelas")

            Dim parameters As New Dictionary(Of String, Object)
            parameters.Add("Titulo", "Relatório De Marcas De Produto.")
            parameters.Add("Codigo", "Código")
            parameters.Add("Descricao", "Descrição")

            Funcoes.BindReport(Me.Page, DS, "Cr_Tabelas", eExportType.PDF, parameters)

            'Else
            '    MsgBox(Me.Page, "Usuário sem permisão para emitir relatório.")
            'End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "Contrato")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

    Protected Sub gridContratos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles gridContratos.SelectedIndexChanged
        Try
            Dim Pos = ListaDeContratos.FindIndex(Function(s) s.Codigo = New Guid(gridContratos.SelectedRow.Cells(1).Text()))

            txtCodigo.Text = ListaDeContratos(Pos).Codigo.ToString
            txtDescricao.Text = ListaDeContratos(Pos).Descricao
            txtDescricao.Focus()
            lnkNovo.Parent.Visible = False
            lnkAtualizar.Parent.Visible = True
            'lnkExcluir.Parent.Visible = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub imgDownload_Click(sender As Object, e As ImageClickEventArgs)
        Try
            Dim imgArquivo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgArquivo.NamingContainer, GridViewRow)

            Dim NomeDoArquivo As String = ListaDeContratos(row.RowIndex).Codigo.ToString() & ".docx"
            Dim CaminhoNomeArquivo As String = Server.MapPath("~/Files/" & NomeDoArquivo)
            If Dir(CaminhoNomeArquivo).Length > 0 Then Kill(CaminhoNomeArquivo)

            Dim Fs As FileStream = New FileStream(CaminhoNomeArquivo, FileMode.Create)
            Fs.Write(ListaDeContratos(row.RowIndex).Arquivo, 0, ListaDeContratos(row.RowIndex).Arquivo.Length)
            Fs.Flush()
            Fs.Close()

            Funcoes.AbrirArquivo(Me.Page, "Files/" & NomeDoArquivo)
            'Funcoes.AbrirExcel(Me.Page, CaminhoNomeArquivo)

            'DownloadDoc(NomeDoArquivo)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub DownloadDoc(ByVal NomeDoArquivo As String)
        Dim Files As FilesManager = New FilesManager()
        Files.IsLocal = True
        Files.FilesFolder = Server.MapPath("~/Files/")


        Files.openFile(NomeDoArquivo, New eTipoDeDocumento)
    End Sub

    Protected Sub imgExcluir_Click(sender As Object, e As ImageClickEventArgs)
        'SessaoRecuperaNotaFiscal()
        'Dim img As ImageButton = CType(sender, ImageButton)
        'Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

        'Dim i As Integer = objNotaFiscal.Itens.FindIndex(Function(s) s.CodigoProduto = row.Cells(1).Text)

        'If objNotaFiscal.IUD = "I" Or objNotaFiscal.Itens(i).IUD = "I" Then
        '    objNotaFiscal.Itens.Remove(objNotaFiscal.Itens(i))
        'Else
        '    objNotaFiscal.Itens(i).IUD = "D"
        'End If
        'SessaoSalvaNotaFiscal()
        'AtualizarItensNoGrid()
        'If FinanceiroNovo Then AtualizaValorParcelamento()


        Try
            'If Funcoes.VerificaPermissao("Marcas", "EXCLUIR") Then

            Dim img As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(img.NamingContainer, GridViewRow)

            Dim Pos = ListaDeContratos.FindIndex(Function(s) s.Codigo = New Guid(row.Cells(1).Text()))

            ListaDeContratos(Pos).IUD = "D"
            If Not ListaDeContratos(Pos).Salvar Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                'CarregarContratos()
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            'End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try

    End Sub
End Class