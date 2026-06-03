Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Drawing
Imports OfficeOpenXml.Style
Imports OfficeOpenXml

Public Class ImportaListaDePrecos
    Inherits BasePage

    Private Mensagem As String
    Dim SqlArray As New ArrayList
    Dim strSQL As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("ImportaListaDePrecos", "ACESSAR") Then
                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, "", True)
                txtMovimnto.Text = Format(Today, "dd/MM/yyyy")
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            filUpload.Value = ""
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkImportar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkImportar.Click
        Try
            Importar()
        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try
    End Sub

    Private Sub Importar()

        If String.IsNullOrWhiteSpace(filUpload.PostedFile.FileName) Then
            MsgBox(Me.Page, "Arquivo não foi selecionado.", eTitulo.Info)
            Exit Sub
        End If

        Dim extArquivo As String() = filUpload.PostedFile.FileName.Split(".")

        If extArquivo(1).ToUpper = "XLS" OrElse extArquivo(1).ToUpper = "XLSX" Then
            'LIBERA
        Else
            MsgBox(Me.Page, "Extensão do Arquivo deve ser XLS ou XLSX, caso tenha dúvida entre em contato com o Suporte.", eTitulo.Info)
            Exit Sub
        End If

        Dim infoarquivo As New IO.FileInfo(filUpload.PostedFile.FileName)
        If Upload(infoarquivo) Then
            If ValidaDados(infoarquivo) Then
                ValidarArquivo(infoarquivo)
            Else
                MsgBox(Me.Page, Mensagem.ToString, eTitulo.Erro)
            End If
        Else
            MsgBox(Me.Page, Mensagem.ToString, eTitulo.Erro)
        End If
    End Sub

    Private Function Upload(ByVal infoarquivo As Object) As Boolean
        Try
            'Verificamos se tem alguma coisa postada 
            If Not IsNothing(filUpload.PostedFile) Then
                'Pegamos as informacoes do arquivo postado 
                'Definimos onde ele será salvo 
                Dim strCaminho As String = Server.MapPath("Files/") & infoarquivo.Name

                If File.Exists(strCaminho) Then File.Delete(strCaminho)

                'Salvamos o mesmo 
                filUpload.PostedFile.SaveAs(strCaminho)
            Else
                MsgBox(Me.Page, "Selecione um arquivo.", eTitulo.Info)
                Return False
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
        Return True
    End Function

    Public Function ValidaDados(ByVal infoarquivo As Object) As Boolean
        Try
            If ddlEmpresa.SelectedIndex = 0 Then
                Mensagem = "Empresa não foi selecionanda."
                Return False
            ElseIf CDate(txtMovimnto.Text).Year < Now().Year Then
                Mensagem = "Ano informado não pode ser menor que o atual."
                Return False
            ElseIf CDate(txtMovimnto.Text).Month < Now().Month Then
                Mensagem = "Mês informado não pode ser menor que o atual."
                Return False
            ElseIf infoarquivo.Name.ToString.Length = 0 Then
                Mensagem = "Arquivo não foi informado."
                Return False
            Else
                Return True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Function

    Private Sub ValidarArquivo(ByVal infoarquivo As Object)
        If infoarquivo Is Nothing OrElse String.IsNullOrWhiteSpace(infoarquivo.Name) Then
            MsgBox(Me.Page, "Arquivo inválido ou não foi encontrado.", eTitulo.Erro)
            Exit Sub
        End If

        Dim caminhoArquivoExcel As String = Server.MapPath("Files/") & infoarquivo.Name
        Dim nomePlanilhaExcel As String = "Planilha1" & "$"

        Dim conexaoOleDb As OleDbConnection = Nothing
        Dim ds As DataSet
        Dim cmd As OleDbDataAdapter

        Try

            Dim arrEmpresa() As String
            arrEmpresa = ddlEmpresa.SelectedValue.Split("-")

            Dim extArquivo As String() = filUpload.PostedFile.FileName.Split(".")

            If extArquivo(1).ToUpper = "XLS" Then
                conexaoOleDb = New OleDbConnection("provider=Microsoft.Jet.OLEDB.4.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 8.0;HDR=Yes'")
            ElseIf extArquivo(1).ToUpper = "XLSX" Then
                conexaoOleDb = New OleDbConnection("provider=Microsoft.ACE.OLEDB.12.0;Data Source= " & caminhoArquivoExcel & ";Extended Properties='Excel 12.0;HDR=Yes'")
            End If

            cmd = New OleDbDataAdapter("Select * from [" & nomePlanilhaExcel & "]", conexaoOleDb)
            cmd.TableMappings.Add("Table", "tabelaExcel")
            ds = New DataSet
            cmd.Fill(ds)

            SqlArray.Clear()

            If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                For Each row As DataRow In ds.Tables(0).Rows

                    If IsDBNull(row("TABELA")) Then
                        Continue For
                    End If

                    If String.IsNullOrWhiteSpace(row("TABELA")) OrElse row("TABELA") = "0" Then
                        MsgBox(Me.Page, "Tabela não pode estar vazia ou ser igual a 0(Zero).", eTitulo.Info)
                        Exit Sub
                    End If

                    If Not arrEmpresa(0) = row("EMPRESA") Then
                        MsgBox(Me.Page, "Arquivo selecionado para importação não corresponde com empresa selecionada.", eTitulo.Info)
                        Exit Sub
                    End If

                    If String.IsNullOrWhiteSpace(row("CODIGO")) Then
                        MsgBox(Me.Page, "Código do Produto não foi informado.", eTitulo.Info)
                        Exit Sub
                    End If

                    Dim codProduto As New Produto(row("CODIGO"))

                    If String.IsNullOrWhiteSpace(codProduto.Nome) Then
                        MsgBox(Me.Page, "Código do Produto não foi encontrado. Código: " & row("CODIGO"), eTitulo.Info)
                        Exit Sub
                    End If

                    If String.IsNullOrWhiteSpace(row("VALOR")) OrElse row("VALOR") = 0 Then
                        MsgBox(Me.Page, "Valor não pode estar vazio ou ser igual a 0(Zero).", eTitulo.Info)
                        Exit Sub
                    End If

                    strSQL = "DELETE ProdutosXPrecos " & vbCrLf &
                                     "WHERE Tabela_Id   = " & row("TABELA") & vbCrLf &
                                     "AND Cliente_Id    = '" & arrEmpresa(0) & "'" & vbCrLf &
                                     "AND EndCliente_Id = " & arrEmpresa(1) & vbCrLf &
                                     "AND Produto_Id    = '" & row("CODIGO") & "'" & vbCrLf &
                                     "AND Data_Id       = '" & Now().ToString("yyyy-MM-dd") & "';"
                    SqlArray.Add(strSQL)

                    strSQL = "INSERT INTO ProdutosXPrecos (Tabela_Id, Cliente_Id, EndCliente_Id, Produto_Id, " & vbCrLf &
                             "                             Data_Id, Moeda, MargemMenor, MargemMaior, Valor, ValorDolar, FixoOperacional, UsuarioInclusao, UsuarioInclusaoData) " & vbCrLf &
                             "VALUES (" & row("TABELA") & ",'" & arrEmpresa(0) & "'," & arrEmpresa(1) & ",'" & row("CODIGO") & "'," & vbCrLf &
                             "'" & CDate(txtMovimnto.Text).ToString("yyyy-MM-dd") & "',1,0,0," & Str(row("VALOR")) & ",0,0,'" & Session("ssNomeUsuario") & "','" & Now().ToString("yyyy-MM-dd") & "');"

                    SqlArray.Add(strSQL)
                Next

                If SqlArray.Count > 0 Then
                    If Banco.GravaBanco(SqlArray) Then
                        MsgBox(Me.Page, "Processo concluido com sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    End If
                Else
                    MsgBox(Me.Page, "Lista não encontrada para importação.", eTitulo.Info)
                End If
            Else
                MsgBox(Me.Page, "Lista não encontrada no arquivo.", eTitulo.Erro)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

    End Sub


    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ImportaListaDePrecos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class