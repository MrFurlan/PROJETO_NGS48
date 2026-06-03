Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports System.IO

Public Class FilesManager

#Region "properties"

    Private _IsLocal As Boolean = False
    Public Property IsLocal() As Boolean
        Get
            Return _IsLocal
        End Get
        Set(ByVal value As Boolean)
            _IsLocal = value
        End Set
    End Property

    Private _FilesFolder As String
    Public Property FilesFolder() As String
        Get
            Return _FilesFolder
        End Get
        Set(ByVal value As String)
            _FilesFolder = value
        End Set
    End Property

#End Region

#Region "contructors"

    Public Sub New()
        Me.New(True)
    End Sub

    Public Sub New(tryConnection As Boolean)
        If tryConnection AndAlso Not IsLocal Then
            Dim ws = getWebService()
            ws.TryConnection()
            Return
        End If

        FilesFolder = HttpContext.Current.Server.MapPath("~/Files")
        If Not Directory.Exists(FilesFolder) Then
            Directory.CreateDirectory(FilesFolder)
        End If
    End Sub

#End Region

#Region "public methods"

    Public Function saveFile(path As String, systemFileName As String, id As Int32, ByVal doc As eTipoDeDocumento) As Boolean
        If IsLocal Then
            Return saveFileLocal(path, systemFileName, id)
        End If
        Return saveFileRemote(path, systemFileName, id, doc)
    End Function

    Public Function getFile(fileName As String, ByVal doc As eTipoDeDocumento) As Byte()
        If IsLocal Then
            Return getFileLocal(fileName)
        End If
        Return getFileRemote(fileName, doc)
    End Function

    Public Function getFileXml(fileName As String) As Byte()
        Dim ws = getWebService()
        Return ws.GetFileXml(fileName)
    End Function

    Public Sub openFile(fileName As String, ByVal doc As eTipoDeDocumento)
        If String.IsNullOrEmpty(fileName) Then
            Return
        End If
        Dim temp = getTempFile(fileName, doc)
        If temp Is Nothing Then
            Throw New FileNotFoundException("File not found: " & Convert.ToString(fileName), fileName)
        End If
        Process.Start(temp)
    End Sub

    Public Function getRemoteFileName(fileName As String, ByVal doc As eTipoDeDocumento) As String
        If String.IsNullOrEmpty(fileName) Then
            Return Nothing
        End If
        Dim ws = getWebService()
        Return ws.GetRemoteFileName(fileName, getTypeDoc(doc))
    End Function

    Public Sub saveFileAs(fileName As String, newPath As String, ByVal doc As eTipoDeDocumento)
        If String.IsNullOrEmpty(fileName) Then
            Return
        End If
        Dim temp = getTempFile(fileName, doc)
        If temp Is Nothing Then
            Throw New FileNotFoundException("File not found: " & Convert.ToString(fileName), fileName)
        End If
        File.Copy(temp, newPath, True)
    End Sub

    Public Function removeFile(fileName As String, ByVal doc As eTipoDeDocumento) As Boolean
        If IsLocal Then
            Return removeFileLocal(fileName)
        End If
        Return removeFileRemote(fileName, getTypeDoc(doc))
    End Function

    Public Function getTempFile(fileName As [String], ByVal doc As eTipoDeDocumento) As [String]
        Return getTempFile(fileName, True, doc)
    End Function

    Public Function getTempFile(fileName As [String], [readOnly] As [Boolean], ByVal doc As eTipoDeDocumento) As [String]
        Dim bytes = getFile(fileName, doc)
        If bytes Is Nothing Then
            Return Nothing
        End If

        Dim tempPath = getTempPath()
        Dim tempFile = Path.GetFileNameWithoutExtension(Path.GetTempFileName()) + Path.GetExtension(fileName)
        Dim temp = Path.Combine(tempPath, tempFile)
        Dim fs = New FileStream(temp, FileMode.Create, FileAccess.Write)
        fs.Write(bytes, 0, bytes.Length)
        fs.Close()

        If [readOnly] Then
            File.SetAttributes(temp, FileAttributes.[ReadOnly] Or FileAttributes.Archive Or FileAttributes.Temporary)
        Else
            File.SetAttributes(temp, FileAttributes.Archive Or FileAttributes.Temporary)
        End If

        Return temp
    End Function

    Public Function getTempPath() As String
        If Directory.Exists(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp")) Then
            Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp")
        End If
        Directory.CreateDirectory(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp"))
        Return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Temp")
    End Function

    Public Function IsConnect() As Boolean
        Dim ws = getWebService()
        Return ws.IsConnect()
    End Function

#End Region

#Region "private methods"

    Private Function saveFileLocal(path As String, systemFileName As String, id As Int32) As Boolean
        If Not File.Exists(path) Then
            Return False
        End If
        Dim newFilepath = System.IO.Path.Combine(FilesFolder, systemFileName)
        File.Copy(path, newFilepath, True)
        Return True
    End Function

    Private Function saveFileRemote(path As String, systemFileName As String, id As Int32, ByVal doc As eTipoDeDocumento) As Boolean
        If Not File.Exists(path) Then
            Return False
        End If

        Dim fs = New FileStream(path, FileMode.Open, FileAccess.Read)
        Dim bytes = New Byte(fs.Length - 1) {}
        fs.Read(bytes, 0, bytes.Length)
        fs.Close()

        Dim ws = getWebService()
        ws.SaveFile(bytes, systemFileName, id, getTypeDoc(doc))
        Return True
    End Function

    Private Function getFileLocal(fileName As String) As Byte()
        If String.IsNullOrEmpty(fileName) Then
            Return Nothing
        End If

        Dim file = Path.Combine(FilesFolder, fileName)
        If Not System.IO.File.Exists(file) Then
            Return Nothing
        End If

        Dim fs = New FileStream(file, FileMode.Open, FileAccess.Read)
        Dim bytes = New Byte(fs.Length - 1) {}
        fs.Read(bytes, 0, bytes.Length)
        fs.Close()
        Return bytes
    End Function

    Private Function getFileRemote(fileName As String, ByVal doc As eTipoDeDocumento) As Byte()
        If String.IsNullOrEmpty(fileName) Then
            Return Nothing
        End If

        Dim ws = getWebService()
        Return ws.GetFile(fileName, getTypeDoc(doc))
    End Function

    Private Function removeFileLocal(fileName As String) As Boolean
        If String.IsNullOrEmpty(fileName) Then
            Return False
        End If

        Dim file = Path.Combine(FilesFolder, fileName)
        If Not System.IO.File.Exists(file) Then
            Return False
        End If

        Dim newFilename = Convert.ToString(fileName) & ".removed"
        System.IO.File.Move(file, Path.Combine(FilesFolder, newFilename))
        Return True
    End Function

    Private Function removeFileRemote(fileName As String, ByVal doc As eTipoDeDocumento) As Boolean
        If String.IsNullOrEmpty(fileName) Then
            Return False
        End If
        Dim ws = getWebService()
        Return ws.RemoveFile(fileName, getTypeDoc(doc))
    End Function

    Private Function getWebService() As FilesServer.Services
        Dim ws As New FilesServer.Services()
        Return ws
    End Function

    Private Function getTypeDoc(ByVal doc As eTipoDeDocumento) As FilesServer.TIPO_DOCUMENTO
        If doc = eTipoDeDocumento.CTRC OrElse doc = eTipoDeDocumento.CT_E Then
            Return FilesServer.TIPO_DOCUMENTO.CONHECIMENTO_DE_TRANSPORTE
        ElseIf doc = eTipoDeDocumento.ManifestoEletronico Then
            Return FilesServer.TIPO_DOCUMENTO.MANIFESTO_ELETRONICO
        End If
        Return FilesServer.TIPO_DOCUMENTO.NOTA_FISCAL
    End Function

#End Region

End Class
