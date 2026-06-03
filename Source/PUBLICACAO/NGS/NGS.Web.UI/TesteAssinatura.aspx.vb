Imports System
Imports System.Collections.Generic
Imports System.Linq
Imports System.Text
Imports System.Threading.Tasks
Imports System.Security.Cryptography
Imports System.Security.Cryptography.X509Certificates
Imports System.Security.Cryptography.Xml
Imports System.Xml
'Imports System.Deployment.Internal.CodeSigning
Imports System.IO
Public Class TesteAssinatura
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
     
    End Sub

    Protected Sub lnkNovo_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkNovo.Click
        Processar()



    End Sub

    Private Sub Processar()
        Try
            CryptoConfig.AddAlgorithm(GetType(RSAPKCS1SHA256SignatureDescription), "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256")
            ' Generate a signing key.
            Dim Key As RSACryptoServiceProvider = New RSACryptoServiceProvider

            Dim p_NumeroSerialCertificado As String = SelecionarCertificado()

            If String.IsNullOrEmpty(p_NumeroSerialCertificado) Then
                MsgBox(Me.Page, "Certificado não foi selecionado")
                Exit Sub
            End If

            SignXmlFile(Server.MapPath("~/SpedReinf/teste.xml"), Server.MapPath("~/SpedReinf/signedExample.xml"), Key, p_NumeroSerialCertificado)
            'Console.WriteLine("XML file signed.")
            ' Verify the signature of the signed XML.
            'Console.WriteLine("Verifying signature...")
            Dim result As Boolean = VerifyXmlFile(Server.MapPath("~/SpedReinf/signedExample.xml"), Key)
            ' Display the results of the signature verification to
            ' the console.
            If result Then
                MsgBox(Me.Page, "The XML signature is valid.")
            Else
                MsgBox(Me.Page, "The XML signature is not valid.")
            End If
        Catch ex As Exception

        End Try
    End Sub

    Public Function SelecionarCertificado() As String
        Try
            'Representa um certificado X509
            Dim objCertificadoX509 As New X509Certificate2

            'Representa o local onde os certificados X509 são armazenados
            'Seleciona os certificados locais e do usuário atual
            Dim getCertificadosX509 As New X509Store("MY", StoreLocation.CurrentUser)
            getCertificadosX509.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)

            'Representa uma coleção de objetos X509Certificate2
            Dim objColecaoCertificadosX509 As New X509Certificate2Collection

            'Abre a caixa de diálogo com os certificados diponiveis
            objColecaoCertificadosX509 = X509Certificate2UI.SelectFromCollection(getCertificadosX509.Certificates, _
            "Certificado(s) Digital(is) disponível(is)", "Selecione o certificado digital para uso no aplicativo", _
            X509SelectionFlag.SingleSelection)

            'Verifica se existe algum certificado selecionado
            If objColecaoCertificadosX509.Count > 0 Then
                'Mostra o assunto do certficado selecionado
                'lblCertificado.Text = objColecaoCertificadosX509.Item(0).Subject.ToString
                'Adiciona o número do serial na variável
                Return objColecaoCertificadosX509.Item(0).SerialNumber.ToString
            End If
        Catch ex As Exception
            Return String.Empty
        End Try
    End Function

    Public Sub SignXmlFile(ByVal FileName As String, ByVal SignedFileName As String, ByVal Key As RSA, ByVal p_NumeroSerialCertificado As String)
        'Representa uma coleção de objetos X509Certificate2
        Dim objColecaoCertificadosX509 As X509Certificate2Collection = Nothing

        'Instância que representa um certificado X509
        Dim objCertificadoX509 As New X509Certificate2()

        'Representa o local onde os certificados X509 são armazenados e
        'seleciona os certificados locais e do usuário atual
        Dim getCertificadosX509 As New X509Store("MY", StoreLocation.CurrentUser)

        'Abrir em modo de leitura e apenas os certificados existentes
        getCertificadosX509.Open(OpenFlags.ReadOnly Or OpenFlags.OpenExistingOnly)

        'Procure por certificados usando o número serial como parâmetro, o último parâmetro indica para a
        'função retornar somente certificados válidos
        objColecaoCertificadosX509 = getCertificadosX509.Certificates.Find(X509FindType.FindBySerialNumber, p_NumeroSerialCertificado, True)

        'Selecionar o certificado e armazena na váriavel
        objCertificadoX509 = objColecaoCertificadosX509.Item(0)





        ' Create a new XML document.
        Dim doc As XmlDocument = New XmlDocument
        ' Load the passed XML file using its name.
        doc.Load(New XmlTextReader(FileName))
        ' Create a SignedXml object.
        Dim signedXml As SignedXml = New SignedXml(doc)
        ' Add the key to the SignedXml document.
        signedXml.SigningKey = Key

        signedXml.SignedInfo.SignatureMethod = "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256"
        ' Create a reference to be signed.

        Dim reference As Reference = New Reference
        reference.Uri = ""
        ' Add an enveloped transformation to the reference.           
        reference.AddTransform(New XmlDsigEnvelopedSignatureTransform)

        'reference.AddTransform(New XmlDsigExcC14NTransform)  'antes estava assim.. usar outro
        Dim xmlResolver As XmlUrlResolver = New XmlUrlResolver
        xmlResolver.Credentials = System.Net.CredentialCache.DefaultCredentials
        Dim xmlTransform As XmlDsigC14NTransform = New XmlDsigC14NTransform
        xmlTransform.Resolver = xmlResolver
        reference.AddTransform(xmlTransform)

        reference.DigestMethod = "http://www.w3.org/2001/04/xmlenc#sha256"
        ' Add the reference to the SignedXml object.
        signedXml.AddReference(reference)




        'A Classe KeyInfo representa o elemento <KeyInfo> da assinatura digital
        Dim KeyInfo As New KeyInfo
        KeyInfo.AddClause(New KeyInfoX509Data(objCertificadoX509))
        'Adiciona o objeto KeyInfo na assinatura
        signedXml.KeyInfo = KeyInfo





        ' Compute the signature.
        signedXml.ComputeSignature()
        ' Get the XML representation of the signature and save
        ' it to an XmlElement object.
        Dim xmlDigitalSignature As XmlElement = signedXml.GetXml
        ' Append the element to the XML document.
        doc.DocumentElement.AppendChild(doc.ImportNode(xmlDigitalSignature, True))
        If (TypeOf doc.FirstChild Is XmlDeclaration) Then
            doc.RemoveChild(doc.FirstChild)
        End If

        ' Save the signed XML document to a file specified
        ' using the passed string.
        Dim xmltw As XmlTextWriter = New XmlTextWriter(SignedFileName, New UTF8Encoding(False))
        doc.WriteTo(xmltw)
        xmltw.Close()
    End Sub

    ' Verify the signature of an XML file against an asymetric
    ' algorithm and return the result.
    Public Function VerifyXmlFile(ByVal Name As String, ByVal Key As RSA) As Boolean
        ' Create a new XML document.
        Dim xmlDocument As XmlDocument = New XmlDocument
        ' Load the passed XML file into the document.
        xmlDocument.Load(Name)
        ' Create a new SignedXml object and pass it
        ' the XML document class.
        Dim signedXml As SignedXml = New SignedXml(xmlDocument)
        ' Find the "Signature" node and create a new
        ' XmlNodeList object.
        Dim nodeList As XmlNodeList = xmlDocument.GetElementsByTagName("Signature")
        ' Load the signature node.
        signedXml.LoadXml(CType(nodeList(0), XmlElement))
        ' Check the signature and return the result.
        Return signedXml.CheckSignature(Key)
    End Function

    Public Class RSAPKCS1SHA256SignatureDescription
        Inherits SignatureDescription

        ''' <summary>
        ''' Registers the http://www.w3.org/2001/04/xmldsig-more#rsa-sha256 algorithm
        ''' with the .NET CrytoConfig registry. This needs to be called once per
        ''' appdomain before attempting to validate SHA256 signatures.
        ''' </summary>
        Public Shared Sub Register()
            CryptoConfig.AddAlgorithm(GetType(RSAPKCS1SHA256SignatureDescription), "http://www.w3.org/2001/04/xmldsig-more#rsa-sha256")
        End Sub

        ''' <summary>
        ''' .NET calls this parameterless ctor
        ''' </summary>
        Public Sub New()
            MyBase.New()
            KeyAlgorithm = "System.Security.Cryptography.RSACryptoServiceProvider"
            DigestAlgorithm = "System.Security.Cryptography.SHA256Managed"
            FormatterAlgorithm = "System.Security.Cryptography.RSAPKCS1SignatureFormatter"
            DeformatterAlgorithm = "System.Security.Cryptography.RSAPKCS1SignatureDeformatter"
        End Sub

        Public Overrides Function CreateDeformatter(ByVal key As AsymmetricAlgorithm) As AsymmetricSignatureDeformatter
            Dim asymmetricSignatureDeformatter = CType(CryptoConfig.CreateFromName(DeformatterAlgorithm), AsymmetricSignatureDeformatter)
            asymmetricSignatureDeformatter.SetKey(key)
            asymmetricSignatureDeformatter.SetHashAlgorithm("SHA256")
            Return asymmetricSignatureDeformatter
        End Function

        Public Overrides Function CreateFormatter(ByVal key As AsymmetricAlgorithm) As AsymmetricSignatureFormatter
            Dim asymmetricSignatureFormatter = CType(CryptoConfig.CreateFromName(FormatterAlgorithm), AsymmetricSignatureFormatter)
            asymmetricSignatureFormatter.SetKey(key)
            asymmetricSignatureFormatter.SetHashAlgorithm("SHA256")
            Return asymmetricSignatureFormatter
        End Function
    End Class

End Class