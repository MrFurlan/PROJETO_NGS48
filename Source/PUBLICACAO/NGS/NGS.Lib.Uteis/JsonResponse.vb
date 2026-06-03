' Classe principal para deserialização
Imports Newtonsoft.Json

Public Class JsonResponse
    Public Property message As String

    ' Método para deserializar o JSON
    Public Shared Function DeserializeJson(json As String) As JsonResponse
        Return JsonConvert.DeserializeObject(Of JsonResponse)(json)
    End Function

    ' Método para deserializar o conteúdo da propriedade "message"
    Public Shared Function GetErrorMessage(json As String) As ErrorMessage

        ' Remove qualquer prefixo como "Erro HTTP:" ou espaços
        json = json.Trim()
        If json.StartsWith("Erro HTTP:", StringComparison.OrdinalIgnoreCase) Then
            json = json.Substring("Erro HTTP:".Length).Trim()
        End If

        Return JsonConvert.DeserializeObject(Of ErrorMessage)(json)

    End Function

End Class


Public Class ApiResponse
    Public Property Success As Boolean
    Public Property Message As String
    Public Property Data As String
End Class

Public Structure RespostaInterpretada
    Public Property Sucesso As Boolean
    Public Property Mensagem As String
End Structure

Public Class DadosNotaFiscal
    Public Property Serie As String
    Public Property Numero As String
End Class