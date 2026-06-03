Imports System.IO
Imports System.Xml
Imports System.Web
Imports NGS.Lib.Negocio


Public Class teste
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        'Dim dir As New System.IO.DirectoryInfo(Server.MapPath("~/Files"))
        'Dim arq As System.IO.FileInfo
        'For Each arq In dir.GetFiles("*.pdf*")
        '    Response.Write(arq.Name & "<br>")
        'Next
        Dim returnFiles As ArrayList = GetFiles(Server.MapPath("~/Files"))

        lst.DataSource = returnFiles
        lst.DataBind()
    End Sub

    Private Sub Files_List()
        Try
            ' The GetFiles() variable will contain 
            ' all files of the C:\MyFolderName directory.
            Dim GetFiles() As String = Nothing
            GetFiles = System.IO.Directory.GetFiles(Server.MapPath("~/Files/"))

            ' The FileInfo variable will contain 
            ' all file information (name, size, etc.).
            Dim FileInfo As System.IO.FileInfo

            ' I loop over on GetFiles object
            ' until all files are read.
            For Each File As String In GetFiles

                ' I get file information.
                FileInfo = New System.IO.FileInfo(File)

                ' I show the file name.
                Response.Write("File Name" & FileInfo.Name)
                Response.Write("<br>")
                Response.Write("File extension" & FileInfo.Extension)
                Response.Write("<br>")
            Next
        Catch ex As Exception

            ' I show an error message if the sub generates an error.
            Response.Write("Error message: " & ex.Message)
        End Try
    End Sub

    Private Function GetFiles(pathFolder As String) As ArrayList
        Dim returnFiles As New ArrayList()

        Dim dirInfo As New DirectoryInfo(pathFolder)
        If dirInfo.Exists Then

            Dim filesInfo As FileSystemInfo() = dirInfo.GetFileSystemInfos()

            For Each fil As FileSystemInfo In filesInfo

                returnFiles.Add(fil.FullName)
                If fil.Attributes = FileAttributes.Directory Then returnFiles.AddRange(GetFiles(fil.FullName))

            Next
        End If

        Return returnFiles
    End Function


    Private Sub EnviarRequisicao()

        Try
            'FRAMEWORK 2 (WEB SERVICE)
            'Dim ws As New Balanca.Balanca() With { _
            '    .Url = String.Format("http://{0}/LerPeso/Balanca.asmx", lblAcesso.Text) _
            '}

            'Return ws.CapturarPeso(lblModulo.Text, True, True, 4096, 1, lblPorta.Text, hdnBaudRateBalanca.Value)

            'FRAMEWORK 4 (WCF)
            'Dim binding As New BasicHttpBinding()
            'Dim ipAddress As String = String.Format("http://{0}/LerPeso/Service.svc", lblAcesso.Text)
            'Dim endpointAddress As New EndpointAddress(ipAddress)

            'Dim ws As New LerPeso.ServiceClient(binding, endpointAddress)

            ''Dim pesoTemporario As Integer = 30000
            'Return ws.CapturarPeso(lblModulo.Text, True, True, 4096, 1, lblPorta.Text, hdnBaudRateBalanca.Value, hdnDataBits.Value, hdnParity.Value, hdnStopBits.Value)

            Dim xml As New StringBuilder
            'xml.Append("POST /ws/cadconsultacadastro/cadconsultacadastro2.asmx HTTP/1.1" & ControlChars.CrLf)
            'xml.Append("Host:sef.sefaz.rs.gov.br" & ControlChars.CrLf)
            'xml.Append("Content-Type: application/soap+xml; charset=utf-8" & ControlChars.CrLf)
            'xml.Append("Content -Length : length()" & ControlChars.CrLf)
            xml.Append("<?xml version=""1.0"" encoding=""utf-8""?> " & ControlChars.CrLf)
            xml.Append("<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" & ControlChars.CrLf)
            xml.Append("<soap12:Header>" & ControlChars.CrLf)
            xml.Append("<nfeCabecMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/CadConsultaCadastro2" > "" & ControlChars.CrLf)
            xml.Append("<cUF></cUF>" & ControlChars.CrLf)
            xml.Append("<versaoDados></versaoDados>" & ControlChars.CrLf)
            xml.Append("</nfeCabecMsg>" & ControlChars.CrLf)
            xml.Append("</soap12:Header>" & ControlChars.CrLf)
            xml.Append("<soap12:Body>" & ControlChars.CrLf)
            xml.Append("<nfeDadosMsg xmlns=""http://www.portalfiscal.inf.br/nfe/wsdl/CadConsultaCadastro2"">xml</nfeDadosMsg>" & ControlChars.CrLf)
            xml.Append("</soap12:Body>" & ControlChars.CrLf)
            xml.AppendLine("</soap12:Envelope>" & ControlChars.CrLf)

            Dim XmlDocument As New XmlDocument()
            XmlDocument.LoadXml(xml.ToString())

            Dim XmlRetorno As New XmlDocument()

            Dim sf As New br.gov.rs.sefaz.sef.CadConsultaCadastro2
            sf.nfeCabecMsgValue.cUF = "MT"
            sf.nfeCabecMsgValue.versaoDados = "3.10"
            sf.Url = "https://sef.sefaz.rs.gov.br/ws/cadconsultacadastro/cadconsultacadastro2.asmx"

            XmlRetorno = sf.consultaCadastro2(XmlDocument)

        Catch ex As Exception
            MsgBox(Me.Page, Funcoes.EliminarCaracteresEspeciais(ex.Message))
        End Try

    End Sub


    'Private Function CriarXml() As Xml
    '    Dim xml As New StringBuilder
    '    Dim writer As New XmlTextWriter("product.xml", System.Text.Encoding.UTF8)
    '    writer.WriteStartDocument(True)
    '    writer.Formatting = Formatting.Indented
    '    writer.Indentation = 2
    '    writer.WriteStartElement("Table")
    '    createNode(1, "Product 1", "1000", writer)
    '    createNode(2, "Product 2", "2000", writer)
    '    createNode(3, "Product 3", "3000", writer)
    '    createNode(4, "Product 4", "4000", writer)
    '    writer.WriteEndElement()
    '    writer.WriteEndDocument()
    '    writer.Close()

    '    Dim x As New XmlDocument()


    '    Return writer

    'End Function

    'Private Sub createNode(ByVal pID As String, ByVal pName As String, ByVal pPrice As String, ByVal writer As XmlTextWriter)
    '    writer.WriteStartElement("Product")
    '    writer.WriteStartElement("Product_id")
    '    writer.WriteString(pID)
    '    writer.WriteEndElement()
    '    writer.WriteStartElement("Product_name")
    '    writer.WriteString(pName)
    '    writer.WriteEndElement()
    '    writer.WriteStartElement("Product_price")
    '    writer.WriteString(pPrice)
    '    writer.WriteEndElement()
    '    writer.WriteEndElement()
    'End Sub


    Protected Sub btnEnviar_Click(sender As Object, e As EventArgs) Handles btnEnviar.Click
        EnviarRequisicao()
    End Sub
End Class