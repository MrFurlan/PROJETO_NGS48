Imports System.IO
Imports System.Security.Cryptography
Imports System.Web.Hosting
Imports System.Xml.Serialization

Public Class DataAcess

    Public Shared Function ReadDataFromXmlFile() As DataModel

        Try

            Dim filePath As String = HostingEnvironment.MapPath(String.Format("~/Files/{0}.xml", "config"))
            ' Verifique se o arquivo existe
            If Not File.Exists(filePath) Then
                Return New DataModel
            End If

            ' Desserialize o XML para uma lista de objetos DataModel
            Dim dataList As New DataModel
            Dim serializer As New XmlSerializer(GetType(DataModel))
            Using reader As New StreamReader(filePath)
                dataList = CType(serializer.Deserialize(reader), DataModel)
            End Using

            Return dataList
        Catch ex As Exception
            ' Lide com erros de leitura do arquivo
            Return New DataModel
        End Try

    End Function

    Public Shared Function CreateMD5Hash(input As String) As String
        ' Converte a string de entrada em um array de bytes
        Dim inputBytes As Byte() = Encoding.ASCII.GetBytes(input)

        ' Cria uma instância de MD5
        Using md5 As MD5 = MD5.Create()
            ' Calcula o hash MD5 da entrada
            Dim hashBytes As Byte() = md5.ComputeHash(inputBytes)

            ' Converte o array de bytes do hash para uma string hexadecimal
            Dim sb As New StringBuilder()
            For Each b As Byte In hashBytes
                sb.Append(b.ToString("x2"))
            Next

            ' Retorna a string hexadecimal
            Return sb.ToString()
        End Using
    End Function

End Class
