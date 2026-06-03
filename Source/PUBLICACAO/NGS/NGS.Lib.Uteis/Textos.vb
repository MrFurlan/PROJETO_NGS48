Imports Microsoft.VisualBasic
Imports System.ComponentModel
Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class Textos

    Public Shared Function DuplicarApostrofos(ByVal Texto As String) As String
        Return Texto.Replace("'", "''")
    End Function

    Public Shared Function RetirarAcentos(ByVal StringAcentuado As String) As String
        Dim strSemAcentos As String = StringAcentuado

        strSemAcentos = Regex.Replace(strSemAcentos, "[áàâãª]", "a")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ÁÀÂÃ]", "A")
        strSemAcentos = Regex.Replace(strSemAcentos, "[éèê]", "e")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ÉÈÊ]", "e")
        strSemAcentos = Regex.Replace(strSemAcentos, "[íìî]", "i")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ÍÌÎ]", "I")
        strSemAcentos = Regex.Replace(strSemAcentos, "[óòôõ°º]", "o")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ÓÒÔÕ]", "O")
        strSemAcentos = Regex.Replace(strSemAcentos, "[úùû]", "u")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ÚÙÛ]", "U")
        strSemAcentos = Regex.Replace(strSemAcentos, "[ç]", "c")
        strSemAcentos = Regex.Replace(strSemAcentos, "[Ç]", "C")

        Return strSemAcentos
    End Function

    Shared Function GetEnumDescription(ByVal EnumConstant As [Enum]) As String
        Dim fi As FieldInfo = EnumConstant.GetType().GetField(EnumConstant.ToString())

        If fi Is Nothing Then
            Return EnumConstant.ToString()
        End If

        Dim attr() As DescriptionAttribute = DirectCast(
       fi.GetCustomAttributes(GetType(DescriptionAttribute), False),
          DescriptionAttribute())
        If attr.Length > 0 Then
            Return attr(0).Description
        Else
            Return EnumConstant.ToString()
        End If
    End Function

End Class