Imports System.Data
Imports System.Data.SqlClient
Imports System.Data.SqlTypes
Imports System.IO
Imports System.Web.HttpFileCollection
Imports System.Web.UI.WebControls.WebControl
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.Xml

Partial Public Class ucFile
    Inherits BaseUserControl

#Region "Property"

    Dim lblCodigo As Label
    Dim lblIUD As Label
    Dim lblDescricao As Label

#End Region

#Region "Event"

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack AndAlso fupArquivo.HasFile Then
            AdicionarArquivo()
        End If
    End Sub

    Protected Sub lnkDownload_Click(sender As Object, e As EventArgs)
        Dim lnkDownload As LinkButton = CType(sender, LinkButton)
        Dim item As DataListItem = CType(lnkDownload.NamingContainer, DataListItem)
        lblIUD = item.FindControl("lblIUD")
        lblDescricao = item.FindControl("lblIUD")
        Download(lblIUD.Text, lblDescricao.Text)
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs)
        Dim lnkExcluir As LinkButton = CType(sender, LinkButton)
        Dim item As DataListItem = CType(lnkExcluir.NamingContainer, DataListItem)
        Dim lblIUD As Label = item.FindControl("lblIUD")
        lblIUD.Text = "D"
        dlsArquivo.DataSource = CarregarLista()
        dlsArquivo.DataBind()

        If TypeOf Me.Page Is NotasFiscaisGerais Then
            CType(Me.Page, NotasFiscaisGerais).RemoverArquivoBD()
        ElseIf TypeOf Me.Page Is NotaFiscalXItens Then
            CType(Me.Page, NotaFiscalXItens).RemoverArquivoBD()
        ElseIf TypeOf Me.Page Is ConhecimentoDeTransporte Then
            CType(Me.Page, ConhecimentoDeTransporte).RemoverArquivoBD()
        End If

    End Sub

    Protected Sub imbDoc_Click(sender As Object, e As System.Web.UI.ImageClickEventArgs)
        Dim imbDoc As ImageButton = CType(sender, ImageButton)
        Dim item As DataListItem = CType(imbDoc.NamingContainer, DataListItem)
        lblCodigo = item.FindControl("lblCodigo")
        lblDescricao = item.FindControl("lblDescricao")
        Download(lblCodigo.Text, lblDescricao.Text)
    End Sub

    Protected Sub AdicionarArquivo()

        If fupArquivo.HasFile Then

            Dim listFile As [Lib].Negocio.ListArquivo = CarregarLista()
            Dim filePath As String = fupArquivo.PostedFile.FileName
            Dim fileName As String = Path.GetFileName(filePath)
            Dim extensao As String = Path.GetExtension(fileName).ToLower()
            Dim contentType As String = String.Empty
            lblErroMsg.Text = String.Empty

            ' Definir content-type
            Select Case extensao
                Case ".jpg" : contentType = "image/jpg"
                Case ".png" : contentType = "image/png"
                Case ".gif" : contentType = "image/gif"
                Case ".pdf" : contentType = "application/pdf"
                Case ".xml" : contentType = "application/xml"
            End Select

            If extensao = ".xml" Then

                Dim match As Match = Regex.Match(fileName, "\d{44}")

                If match.Success Then

                    Dim modelo As String = match.Value.Substring(20, 2) ' posição 21 e 22 (zero-based)
                    Select Case modelo
                        Case "55"
                            fileName = String.Format("{0}-nfe.xml", match.Value)
                            filePath = String.Format("{0}-nfe.xml", match.Value)
                        Case "57"
                            fileName = String.Format("{0}-cte.xml", match.Value)
                            filePath = String.Format("{0}-cte.xml", match.Value)
                    End Select

                Else
                    lblErroMsg.Text = "Algum problema com o arquivo! Entre em contato com o suporte."
                    Exit Sub
                End If

            End If

            If Not String.IsNullOrWhiteSpace(contentType) Then
                Dim caminhoArquivos As String = Server.MapPath("~/Files/")
                Dim caminhoCompleto As String = Path.Combine(caminhoArquivos, fileName)

                ' Remove arquivo anterior se já existir
                If File.Exists(caminhoCompleto) Then
                    Dim tentativas As Integer = 0
                    Dim maxTentativas As Integer = 5
                    Dim deletado As Boolean = False

                    While tentativas < maxTentativas AndAlso Not deletado
                        Try
                            Using fs As New FileStream(caminhoCompleto, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                                fs.Close()
                            End Using

                            File.Delete(caminhoCompleto)
                            deletado = True
                        Catch
                            tentativas += 1
                            Threading.Thread.Sleep(200) ' Aguarda antes de tentar de novo
                        End Try
                    End While

                    If Not deletado Then
                        lblErroMsg.Text = "Arquivo está em uso e não pôde ser excluído."
                        Exit Sub
                    End If
                End If

                ' Verifica se já está na lista (pendente)
                Dim temArquivo As Boolean = listFile.Any(Function(arq) arq.Descricao = fileName AndAlso arq.IUD = "I")

                ' Salva novo arquivo
                fupArquivo.SaveAs(caminhoCompleto)

                ' Adiciona na lista
                If Not temArquivo Then
                    listFile.Add(New [Lib].Negocio.Arquivo() With {
                    .IUD = "I",
                    .Descricao = filePath,
                    .Codigo = String.Empty
                })
                End If

                ' Recarrega controle
                dlsArquivo.DataSource = listFile
                dlsArquivo.DataBind()

                ' Processa XML, se aplicável
                If contentType = "application/xml" Then
                    Try
                        Dim xmlDoc As New XmlDocument
                        Using fs As New FileStream(caminhoCompleto, FileMode.Open, FileAccess.Read, FileShare.Read)
                            xmlDoc.Load(fs)
                        End Using

                        Dim chaveNFe As String = String.Empty
                        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 OrElse xmlDoc.GetElementsByTagName("protCTe").Count = 1 Then
                            If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
                                chaveNFe = xmlDoc.GetElementsByTagName("protNFe").GetNode("protNFe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chNFe").InnerText
                            Else
                                chaveNFe = xmlDoc.GetElementsByTagName("protCTe").GetNode("protCTe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chCTe").InnerText
                            End If

                            If chaveNFe.Length = 44 Then
                                Dim modeloNFe As String = Mid(fileName, 21, 2)
                                Dim nomeDestino As String = fileName

                                If modeloNFe = "55" AndAlso Not fileName.Contains("-nfe") Then
                                    nomeDestino = chaveNFe & "-nfe.xml"
                                ElseIf modeloNFe = "57" AndAlso Not fileName.Contains("-CTe") Then
                                    nomeDestino = chaveNFe & "-CTe.xml"
                                End If

                                ' Cópia para backup
                                Dim backupDir As String = "C:/Alfasig/LeituraNFe/-download/"
                                Directory.CreateDirectory(backupDir)
                                Dim destinoPath As String = Path.Combine(backupDir, nomeDestino)

                                If Not File.Exists(destinoPath) Then
                                    ' Garante que o arquivo não esteja em uso
                                    Dim tentativas As Integer = 0
                                    While tentativas < 5
                                        Try
                                            Using fsTest As New FileStream(caminhoCompleto, FileMode.Open, FileAccess.Read, FileShare.Read)
                                                ' Arquivo liberado
                                            End Using
                                            File.Copy(caminhoCompleto, destinoPath)
                                            Exit While
                                        Catch
                                            Threading.Thread.Sleep(200)
                                            tentativas += 1
                                        End Try
                                    End While
                                End If

                                CType(Me.Page, IBasePage).Carregar(Path.GetFileNameWithoutExtension(fileName))
                            Else
                                lblErroMsg.Text = "Arquivo XML não corresponde a uma NFe válida."
                                listFile.RemoveAt(listFile.Count - 1)
                            End If
                        Else
                            lblErroMsg.Text = "Arquivo XML sem protocolo."
                        End If
                    Catch ex As Exception
                        lblErroMsg.Text = "Erro ao processar XML: " & ex.Message
                        listFile.RemoveAt(listFile.Count - 1)
                    End Try
                End If

                ' Atualiza banco
                If TypeOf Me.Page Is NotasFiscaisGerais Then
                    CType(Me.Page, NotasFiscaisGerais).AdicionarArquivoBD()
                ElseIf TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).AdicionarArquivoBD()
                ElseIf TypeOf Me.Page Is NotasDeTerceiro Then
                    CType(Me.Page, NotasDeTerceiro).AdicionarArquivoBD()
                End If
            Else
                lblErroMsg.Text = "Somente arquivos com extensão .jpg, .png, .gif, .pdf ou .xml são permitidos."
            End If
        End If
    End Sub

    Protected Sub AdicionarArquivo_bkp()
        If fupArquivo.HasFile Then
            Dim listFile As [Lib].Negocio.ListArquivo = CarregarLista()
            Dim filePath As String = fupArquivo.PostedFile.FileName
            Dim fileName As String = Path.GetFileName(filePath)
            Dim extensao As String = Path.GetExtension(fileName)
            Dim contentType As String = String.Empty
            lblErroMsg.Text = String.Empty
            Select Case extensao.ToLower
                Case ".jpg"
                    contentType = "image/jpg"
                    Exit Select
                Case ".png"
                    contentType = "image/png"
                    Exit Select
                Case ".gif"
                    contentType = "image/gif"
                    Exit Select
                Case ".pdf"
                    contentType = "application/pdf"
                Case ".xml"
                    contentType = "application/xml"
                    Exit Select
            End Select

            If Not String.IsNullOrWhiteSpace(contentType) Then

                Dim strCaminho As String = Server.MapPath("Files/") & fileName

                If File.Exists(strCaminho) Then
                    File.Delete(strCaminho)
                End If

                Dim temArquivo As Boolean
                For Each arq In listFile
                    If arq.Descricao = fileName And arq.IUD = "I" Then
                        temArquivo = True
                    End If
                Next

                'Salva o arquivo em Files
                fupArquivo.SaveAs(strCaminho)

                If Not temArquivo Then
                    listFile.Add(New [Lib].Negocio.Arquivo() With {
                                         .IUD = "I",
                                         .Descricao = filePath,
                                         .Codigo = String.Empty})
                End If

                dlsArquivo.DataSource = listFile
                dlsArquivo.DataBind()

                fupArquivo.Dispose()

                'Para preencher os campos da nota em NotasFiscaisGerais.aspx
                If contentType = "application/xml" Then
                    Try

                        Dim xmlDoc As New XmlDocument
                        xmlDoc.Load(Server.MapPath(String.Format("~/Files/{0}", fileName)))

                        If xmlDoc.GetElementsByTagName("protNFe").Count = 1 OrElse xmlDoc.GetElementsByTagName("protCTe").Count = 1 Then

                            Dim ChaveNFe As String = String.Empty

                            If xmlDoc.GetElementsByTagName("protNFe").Count = 1 Then
                                ChaveNFe = xmlDoc.GetElementsByTagName("protNFe").GetNode("protNFe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chNFe").InnerText
                            Else
                                ChaveNFe = xmlDoc.GetElementsByTagName("protCTe").GetNode("protCTe").ChildNodes.GetNode("infProt").ChildNodes.GetNode("chCTe").InnerText
                            End If

                            If ChaveNFe.Count = 44 Then

                                Dim ModeloNFe As String = Mid(fileName, 21, 2)
                                Dim arquivo As String = String.Empty

                                If ModeloNFe.Equals("55") Then

                                    If Not fileName.Contains("-nfe") Then
                                        arquivo = fileName
                                    Else
                                        'strCaminho = Server.MapPath("Files/") & ChaveNFe & "-nfe.xml"
                                        arquivo = ChaveNFe & "-nfe.xml"
                                    End If

                                ElseIf ModeloNFe.Equals("57") Then

                                    If fileName.Contains("-CTe") Then
                                        arquivo = fileName
                                    Else
                                        'strCaminho = Server.MapPath("Files/") & ChaveNFe & "-CTe.xml"
                                        arquivo = ChaveNFe & "-CTe.xml"
                                    End If
                                End If

                                If File.Exists(strCaminho) Then
                                    Dim sourceDir As String = Server.MapPath("~/Files/")
                                    Dim backupDir As String = "C:/Alfasig/LeituraNFe/-download/"

                                    If Not File.Exists(backupDir & arquivo) Then
                                        File.Copy(Path.Combine(sourceDir, fileName), Path.Combine(backupDir, arquivo))
                                    End If
                                End If


                                CType(Me.Page, IBasePage).Carregar(Path.GetFileName(fileName).Replace(Path.GetExtension(fileName), ""))
                            Else
                                lblErroMsg.Text = "Arquivo XML não corresponde a uma NFe"
                                listFile.RemoveAt(listFile.Count - 1)
                            End If
                        Else
                            lblErroMsg.Text = "Arquivo XML sem o Protocolo."
                        End If
                    Catch ex As Exception
                        lblErroMsg.Text = "Arquivo XML não corresponde a uma NFe"
                        listFile.RemoveAt(listFile.Count - 1)
                    End Try
                End If

                If TypeOf Me.Page Is NotasFiscaisGerais Then
                    CType(Me.Page, NotasFiscaisGerais).AdicionarArquivoBD()
                ElseIf TypeOf Me.Page Is NotaFiscalXItens Then
                    CType(Me.Page, NotaFiscalXItens).AdicionarArquivoBD()
                ElseIf TypeOf Me.Page Is NotasDeTerceiro Then
                    CType(Me.Page, NotasDeTerceiro).AdicionarArquivoBD()
                End If

            Else
                lblErroMsg.Text = "Arquivos somente com extensão .jpg .png .gif .pdf"
            End If
        End If
    End Sub

#End Region

#Region "Methods"

    Public Sub Salvar(ByRef listArq As [Lib].Negocio.ListArquivo)
        Dim listArqAux = New [Lib].Negocio.ListArquivo
        Dim i As Integer = 0
        For Each _Item In dlsArquivo.Items
            lblCodigo = _Item.FindControl("lblCodigo")
            lblIUD = _Item.FindControl("lblIUD")
            lblDescricao = _Item.FindControl("lblDescricao")
            If lblIUD.Text.Equals("I") AndAlso File.Exists(Server.MapPath("Files/") & Server.HtmlDecode(lblDescricao.Text)) Then
                Dim fs As Stream = New FileStream(Server.MapPath("Files/") & Server.HtmlDecode(lblDescricao.Text), FileMode.Open, FileAccess.Read)
                Dim br As New BinaryReader(fs)
                Dim bytes As Byte() = br.ReadBytes(fs.Length)

                listArqAux.Add(New [Lib].Negocio.Arquivo() With {
                                     .IUD = lblIUD.Text,
                                     .Codigo = lblCodigo.Text,
                                     .Descricao = lblDescricao.Text,
                                     .Arquivo = bytes})
            Else
                Dim arquivo As Byte() = listArq.Find(Function(x) x.Codigo = lblCodigo.Text).Arquivo

                listArqAux.Add(New [Lib].Negocio.Arquivo() With {
                                    .IUD = lblIUD.Text,
                                    .Codigo = lblCodigo.Text,
                                    .Descricao = lblDescricao.Text,
                                    .Arquivo = arquivo})
            End If
            i += 1
        Next

        listArq = listArqAux

    End Sub

    Public Sub Baixar(ByVal pCodigo As String, ByVal pDescricao As String)
        Download(pCodigo, pDescricao)
    End Sub


    Private Sub RegistrarLog(mensagem As String)
        Dim caminhoLog As String = Server.MapPath("~/Logs/log.txt")
        Using sw As New StreamWriter(caminhoLog, True)
            sw.WriteLine(DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & " - " & mensagem)
        End Using
    End Sub

    Protected Sub Download(ByVal pCodigo As String, ByVal pDescricao As String)

        Try

            If Not String.IsNullOrWhiteSpace(Server.HtmlDecode(pCodigo)) Then

                Dim ds As New DataSet
                Dim Sql As String
                Sql = "SELECT  Arquivo AS ByteFile, Descricao FROM VW_Documentos WHERE Arquivo_Id = '" & pCodigo & "';"

                ds = Banco.ConsultaReport(Sql, "Arquivos")

                For Each row As DataRow In ds.Tables(0).Rows

                    Dim filePath As String = Server.MapPath("Files/") & pDescricao
                    Dim objContext As Byte() = DirectCast(row("ByteFile"), Byte())
                    Dim fName As String = row("Descricao").ToString()

                    If File.Exists(filePath) Then
                        File.Delete(filePath)
                    End If
                    System.IO.File.WriteAllBytes(filePath, objContext)

                Next

            End If

            Dim fileInfo As FileInfo = New FileInfo(Server.MapPath("Files/") & pDescricao)
            Response.Clear()
            Response.BufferOutput = False
            Response.AddHeader("Content-Disposition", "attachment; filename=" + fileInfo.Name.Replace(" ", ""))
            Response.AddHeader("Content-Length", fileInfo.Length.ToString())
            Response.ContentType = "application/octet-stream"
            Response.WriteFile(fileInfo.FullName)
            Response.Flush()
            HttpContext.Current.ApplicationInstance.CompleteRequest()

        Catch ex As Exception
            RegistrarLog("Erro: " & ex.Message)
        End Try
    End Sub

    Protected Sub Download_BKP(ByVal pCodigo As String, ByVal pDescricao As String)

        Try

            '' Necessário atribuir permissões a AppPool 
            'CREATE Login [IIS APPPOOL\DefaultAppPool] FROM WINDOWS;

            'USE [DOCUMENTOS];
            'CREATE USER [IIS APPPOOL\DefaultAppPool] FOR LOGIN [IIS APPPOOL\DefaultAppPool];

            'ALTER ROLE [db_owner] ADD MEMBER [IIS APPPOOL\DefaultAppPool];
            'GRANT SELECT, INSERT, UPDATE, DELETE ON SCHEMA:: dbo TO [IIS APPPOOL\DefaultAppPool];


            If Not String.IsNullOrWhiteSpace(Server.HtmlDecode(pCodigo)) Then
                Dim cLocal As String() = UsuarioServidor.EnderecoLocal.Split(";")
                Dim cs As String = cLocal(0) & ";" & "Initial Catalog=Documentos" & ";" & cLocal(2) & ";" & cLocal(3) & "; Integrated Security=true;"
                'Dim cs As String = [Lib].Negocio.UsuarioServidor.Conexao & "; Integrated Security=true;"

                'Dim cs As String = "Data Source=(local);Initial Catalog=Documentos;Integrated Security=True;"
                'Data Source=SRVNGS; Initial Catalog=Verde; User Id=n@$; Password=pwd_ngs123

                Using con As New SqlConnection(cs)
                    con.Open()
                    Dim txn As SqlTransaction = con.BeginTransaction()
                    Dim sql As String = "SELECT Arquivo.PathName(), GET_FILESTREAM_TRANSACTION_CONTEXT(), Descricao FROM Arquivo WHERE Arquivo_Id = @Arquivo_Id"
                    Dim cmd As New SqlCommand(sql, con, txn)
                    cmd.Parameters.AddWithValue("@Arquivo_Id", pCodigo)
                    Dim rdr As SqlDataReader = cmd.ExecuteReader()

                    If rdr.Read() Then
                        Dim filePath As String = rdr(0).ToString()
                        Dim objContext As Byte() = DirectCast(rdr(1), Byte())
                        Dim fName As String = rdr(2).ToString()

                        Using sfs As New SqlFileStream(filePath, objContext, FileAccess.Read)
                            'RegistrarLog("Iniciando o download do arquivo: " & fName)

                            Response.Clear()
                            Response.BufferOutput = False
                            Response.ContentType = "application/octet-stream"
                            Response.AddHeader("Content-Disposition", "attachment; filename=" & fName.Replace(" ", ""))
                            Response.AddHeader("Content-Length", sfs.Length.ToString())
                            Dim bufferSize As Integer = 8192
                            Dim buffer As Byte() = New Byte(bufferSize - 1) {}
                            Dim bytesRead As Integer = 0
                            Dim totalBytesRead As Long = 0

                            Do
                                bytesRead = sfs.Read(buffer, 0, buffer.Length)
                                If bytesRead > 0 Then
                                    Response.OutputStream.Write(buffer, 0, bytesRead)
                                    Response.Flush()
                                    totalBytesRead += bytesRead
                                End If
                            Loop While bytesRead > 0

                            HttpContext.Current.ApplicationInstance.CompleteRequest()

                            'RegistrarLog("Download concluído com sucesso, tamanho enviado: " & totalBytesRead.ToString())
                            'RegistrarLog("Diretório de origem: " & filePath)
                            'RegistrarLog("Tamanho esperado do arquivo: " & sfs.Length.ToString())
                        End Using
                    Else
                        'RegistrarLog("Arquivo não encontrado para o código: " & pCodigo)
                    End If

                    rdr.Close()
                    txn.Commit()
                End Using
            Else
                Dim fileInfo As FileInfo = New FileInfo(Server.MapPath("Files/") & pDescricao)
                Response.Clear()
                Response.BufferOutput = False
                Response.AddHeader("Content-Disposition", "attachment; filename=" + fileInfo.Name.Replace(" ", ""))
                Response.AddHeader("Content-Length", fileInfo.Length.ToString())
                Response.ContentType = "application/octet-stream"
                Response.WriteFile(fileInfo.FullName)
                Response.Flush()
                HttpContext.Current.ApplicationInstance.CompleteRequest()

                'RegistrarLog("Download de arquivo estático concluído: " & fileInfo.Name)
            End If
        Catch ex As Exception
            RegistrarLog("Erro: " & ex.Message)
        End Try
    End Sub

    Protected Sub Download_old(ByVal pCodigo As String, ByVal pDescricao As String)
        Try
            If Not String.IsNullOrWhiteSpace(Server.HtmlDecode(pCodigo)) Then
                'Dim cs As String = [Lib].Negocio.UsuarioServidor.Conexao & "; Integrated Security=true;"

                'Dim cs As String = "Data Source=(local);Initial Catalog=Documentos;Integrated Security=True;"
                'Data Source=SRVNGS; Initial Catalog=Verde; User Id=n@$; Password=pwd_ngs123

                ''Testando - Furlan - 10/10/2024
                Dim cLocal As String() = UsuarioServidor.EnderecoLocal.Split(";")
                Dim cs As String = cLocal(0) & ";" & "Initial Catalog=Documentos" & ";" & cLocal(2) & ";" & cLocal(3) & "; Integrated Security=true;"

                Using con As New SqlConnection(cs)
                    con.Open()
                    Dim txn As SqlTransaction = con.BeginTransaction()
                    'Dim sql As String = "SELECT Arquivo.PathName(), GET_FILESTREAM_TRANSACTION_CONTEXT(), Descricao FROM VW_Documentos WHERE Arquivo_Id = '" & pCodigo & "'"
                    Dim sql As String = "SELECT Arquivo.PathName(), GET_FILESTREAM_TRANSACTION_CONTEXT(), Descricao FROM Arquivo WHERE Arquivo_Id = '" & pCodigo & "'"
                    Dim cmd As New SqlCommand(sql, con, txn)
                    Dim rdr As SqlDataReader = cmd.ExecuteReader()
                    Response.BufferOutput = True

                    While rdr.Read()
                        Dim filePath As String = rdr(0).ToString()
                        Dim objContext As Byte() = DirectCast(rdr(1), Byte())
                        Dim fName As String = rdr(2).ToString()

                        Dim sfs As New SqlFileStream(filePath, objContext, System.IO.FileAccess.Read)

                        Dim buffer As Byte() = New Byte(CInt(sfs.Length) - 1) {}
                        sfs.Read(buffer, 0, buffer.Length)
                        sfs.Close()

                        Dim filename As String = Server.MapPath("Files/") & fName

                        MsgBox(Me.Page, filename & " #1 ", eTitulo.Info)

                        If File.Exists(filename) Then
                            File.Delete(filename)
                        End If

                        Dim fs As New System.IO.FileStream(filename, FileMode.Create, FileAccess.Write, FileShare.Write)
                        'Dim fs As New System.IO.FileStream(filename, FileMode.Open, FileAccess.ReadWrite, FileShare.ReadWrite)
                        fs.Write(buffer, 0, buffer.Length)
                        'fs.Flush()
                        fs.Close()


                        MsgBox(Me.Page, filename & " #2 ", eTitulo.Info)
                        Response.Clear()
                        Response.ClearHeaders()
                        Response.AddHeader("content-length", buffer.Length.ToString())
                        Response.ContentType = "text/plain"
                        Response.AppendHeader("content-disposition", "attachment;filename=" & fName.Replace(" ", ""))
                        Response.OutputStream.Write(buffer, 0, buffer.Length)
                        MsgBox(Me.Page, filename & " #3 ", eTitulo.Info)

                        'Response.Flush()
                        HttpContext.Current.ApplicationInstance.CompleteRequest()


                    End While

                    rdr.Close()
                    txn.Commit()
                    con.Close()
                End Using
            Else
                Dim fileInfo As FileInfo = New FileInfo(Server.MapPath("Files/") & pDescricao)
                Response.Clear()
                Response.AddHeader("Content-Disposition", "attachment; filename=" + fileInfo.Name.Replace(" ", ""))
                Response.AddHeader("Content-Length", fileInfo.Length.ToString())
                Response.ContentType = "text/plain"
                Response.WriteFile(fileInfo.FullName)
                HttpContext.Current.ApplicationInstance.CompleteRequest()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Bind(ByRef listArq As [Lib].Negocio.ListArquivo)
        dlsArquivo.DataSource = listArq
        dlsArquivo.DataBind()
    End Sub

    Private Function CarregarLista() As [Lib].Negocio.ListArquivo
        Dim listArq As New [Lib].Negocio.ListArquivo
        For Each _Item In dlsArquivo.Items
            lblCodigo = _Item.FindControl("lblCodigo")
            lblIUD = _Item.FindControl("lblIUD")
            lblDescricao = _Item.FindControl("lblDescricao")

            listArq.Add(New [Lib].Negocio.Arquivo() With { _
                                 .IUD = lblIUD.Text, _
                                 .Codigo = lblCodigo.Text, _
                                 .Descricao = lblDescricao.Text})
        Next
        Return listArq
    End Function

    Public Sub Clear()
        dlsArquivo.DataSource = Nothing
        dlsArquivo.DataBind()
        lblErroMsg.Text = String.Empty
    End Sub

    Protected Sub Item_Bound(sender As Object, e As DataListItemEventArgs)
        If e.Item.ItemType = ListItemType.Item OrElse e.Item.ItemType = ListItemType.AlternatingItem Then
            Dim imb As ImageButton = e.Item.FindControl("imbDoc")
            lblIUD = e.Item.FindControl("lblIUD")
            Dim item() As String = DirectCast(e.Item.DataItem, NGS.Lib.Negocio.Arquivo).Descricao.Split(".")
            Dim panArq As Panel = e.Item.FindControl("panArquivo")
            If item(1).ToUpper.Equals("XML") Then
                imb.ImageUrl = "~/images/xml16x16.png"
            ElseIf item(1).ToUpper.Equals("PDF") Then
                imb.ImageUrl = "~/images/icopdf16X16.jpg"
            Else
                imb.ImageUrl = "~/images/icoJpg.gif"
            End If
            imb.ToolTip = String.Format("Download - {0}.{1}", item(0), item(1))

            If lblIUD.Text = "D" Then
                panArq.Style.Add("display", "none")
            End If
            ScriptManager.GetCurrent(Me.Page).RegisterPostBackControl(CType(e.Item.FindControl("imbDoc"), ImageButton))
        End If
    End Sub

    Protected Sub Item_Command(sender As Object, e As DataListCommandEventArgs)
        If e.CommandName = "Delete" Then
            lblIUD = e.Item.FindControl("lblIUD")
            lblIUD.Text = "D"
            dlsArquivo.DataSource = CarregarLista()
            dlsArquivo.DataBind()
        End If

    End Sub

#End Region

End Class

