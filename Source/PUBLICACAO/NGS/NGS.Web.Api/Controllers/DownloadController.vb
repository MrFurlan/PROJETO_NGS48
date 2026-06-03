Imports System.Security.Claims
Imports System.Text
Imports System.Web.Http
Imports System.Web
Imports System.Net.Http
Imports System.Net
Imports NGS.Lib.Negocio
Imports System.Threading.Tasks
Imports Ionic.Zip

Public Class DownloadController
    Inherits ApiController

    Private ReadOnly _httpContext As HttpContextBase

    Public Sub New()

    End Sub

    Public Sub New(httpContext As HttpContextBase)
        _httpContext = httpContext
    End Sub

    <HttpGet>
    <AllowAnonymous>
    <Route("api/Download/ZIP")>
    Public Function Zip(<FromUri> path As String) As HttpResponseMessage

        'http://localhost:44323/api/Download/Zip?path=caminho
        'https://www.nuget.org/packages/Ionic.Zip/

        'Dim encodedPath As String = EncodeBase64(path)
        Dim decodedPath As String = DecodeBase64(path)

        Using zipFile As New ZipFile()

            zipFile.AlternateEncodingUsage = ZipOption.AsNecessary

            zipFile.AddDirectory(decodedPath, "Arquivos")

            'zipFile.AddDirectoryByName("Files")
            'zipFile.AddFile(filePath, "Files")

            HttpContext.Current.Response.Clear()
            HttpContext.Current.Response.BufferOutput = False
            Dim zipName As String = [String].Format("Arq_{0}.zip", DateTime.Now.ToString("yyyyMMddHHmmss"))
            HttpContext.Current.Response.ContentType = "application/zip"
            HttpContext.Current.Response.AddHeader("content-disposition", "attachment; filename=" + zipName)
            zipFile.Save(HttpContext.Current.Response.OutputStream)
            'HttpContext.Current.Response.End()
        End Using

        Return Request.CreateResponse(HttpStatusCode.OK)

    End Function

    Public Function DecodeBase64(ByVal base64EncodedData As String) As String
        Dim base64EncodedBytes As Byte() = Convert.FromBase64String(base64EncodedData)
        Return Encoding.UTF8.GetString(base64EncodedBytes)
    End Function

    Public Function EncodeBase64(ByVal plainText As String) As String
        Dim plainTextBytes As Byte() = Encoding.UTF8.GetBytes(plainText)
        Return Convert.ToBase64String(plainTextBytes)
    End Function

End Class
