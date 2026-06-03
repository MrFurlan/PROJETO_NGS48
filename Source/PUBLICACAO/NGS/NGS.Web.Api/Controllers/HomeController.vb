Imports System.Threading.Tasks

<AllowAnonymous>
Public Class HomeController
    Inherits System.Web.Mvc.Controller

    Function Index() As ActionResult

        ViewData("Title") = "Home Page"

        Return View()
    End Function

    ' Action para abrir e imprimir o PDF
    Public Function PrintPdf() As ActionResult

        'http://localhost:44323/home/printpdf
        ' Ajuste o caminho para refletir o caminho virtual correto
        Dim pdfUrl As String = Url.Content("http://localhost:44323/files/41240816977573000100550010000368581041504944.pdf")


        Dim filePath As String = "D:\Projetos\NGS\PROJETO-NGS\Source\PUBLICACAO\NGS\NGS.Web.UI\Files\teste.pdf"

        ' Verifica se o arquivo existe
        If System.IO.File.Exists(filePath) Then
            ' Lê o arquivo para ser enviado ao cliente
            Dim fileBytes As Byte() = System.IO.File.ReadAllBytes(filePath)
            Dim fileName As String = "teste.pdf"

            ' Retorna o arquivo como um download
            Return File(fileBytes, "application/pdf", fileName)
        Else
            ' Se o arquivo não for encontrado, redireciona para uma página de erro ou mostra uma mensagem
            Return HttpNotFound("O arquivo PDF não foi encontrado.")
        End If

    End Function

    Function TextView() As ActionResult

        ViewData("Title") = "Home Page"

        Dim baseUrl As String = "https://example.com" ' Substitua pela URL base da sua API
        Dim username As String = "your_username"
        Dim password As String = "your_password"
        Dim xmlId As String = "xml_id_here" ' Substitua pelo ID do XML que deseja baixar
        Dim outputFilePath As String = "C:\path\to\save\file\downloaded.xml"

        Try
            ' Chama a função para obter o token de acesso do módulo ApiModule
            'Dim accessToken As String = ApiModuleDownloadMIC.GetAccessTokenAsync(baseUrl, username, password).Result
            ' Chama a função para fazer o download do XML do módulo ApiModule
            'ApiModuleDownloadMIC.DownloadXmlAsync(baseUrl, accessToken, xmlId, outputFilePath).Wait()

            ViewBag.Message = "XML downloaded successfully."

        Catch ex As Exception
            ViewBag.Message = $"Error: {ex.Message}"
        End Try

        Return View()
    End Function

End Class
