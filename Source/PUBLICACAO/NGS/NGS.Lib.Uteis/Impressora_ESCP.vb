Imports Microsoft.VisualBasic

Public Class Impressora_ESCP

    Public Const PRINT_START_COMMAND = ChrW(27)
    Public Const PRINT_DOUBLEWIDTH_ON = ChrW(14)
    Public Const PRINT_CONDENSED_ON = ChrW(15)
    Public Const PRINT_CONDENSED_OFF = ChrW(18)
    Public Const PRINT_STYLE_BOLD_ON = PRINT_START_COMMAND & ChrW(69)
    Public Const PRINT_STYLE_BOLD_OFF = PRINT_START_COMMAND & ChrW(70)
    Public Const PRINT_CARACTER_10DPI = PRINT_START_COMMAND & "P"c
    Public Const PRINT_TYPE_DRAFT = PRINT_START_COMMAND & "x"c & "0"c
    Public Const PRINT_PAGE_SIZE = PRINT_START_COMMAND & "C"c
    Public Const PRINT_6_LPI = PRINT_START_COMMAND & ChrW(50)
    Public Const PRINT_8_LPI = PRINT_START_COMMAND & ChrW(48)
    Public Const NEW_LINE_6_LPI = PRINT_6_LPI & ControlChars.Lf
    Public Const NEW_LINE_8_LPI = PRINT_8_LPI & ControlChars.Lf

    Public Function TamanhoPagina(ByVal Linhas As Integer) As String
        Return PRINT_PAGE_SIZE & ChrW(Linhas)
    End Function

    Public Function MetadePagina() As String
        Return PRINT_PAGE_SIZE & ChrW(33)
    End Function

End Class
