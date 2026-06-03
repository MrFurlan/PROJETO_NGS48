Imports System.Text
Imports System.Globalization
Imports System.Text.RegularExpressions
Imports OfficeOpenXml
Imports System.IO
Imports System.Security.Cryptography

Public Class FuncoesStrings

    Public Shared Function CodificarPara64Bits(ByVal Valor As String) As String
        Try
            Dim byteCodificar As Byte() = ASCIIEncoding.ASCII.GetBytes(Valor)
            Dim strValorConvertido As String = Convert.ToBase64String(byteCodificar)
            Return strValorConvertido
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function DecodificarDe64Bits(ByVal ValorCodificado As String) As String
        Try
            Dim byteCodificar As Byte() = Convert.FromBase64String(ValorCodificado)
            Dim strValorConvertido As String = ASCIIEncoding.ASCII.GetString(byteCodificar)
            Return strValorConvertido
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function CodificarPara128Bits(ByVal Valor As String) As String
        Try
            Return ""
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function LerExcelParaDataSet(caminhoArquivo As String) As DataSet
        'ExcelPackage.LicenseContext = LicenseContext.NonCommercial

        Dim ds As New DataSet()
        Dim dt As New DataTable("tabelaExcel")

        Using package As New ExcelPackage(New FileInfo(caminhoArquivo))
            Dim planilha = package.Workbook.Worksheets.FirstOrDefault()
            If planilha Is Nothing Then
                Throw New Exception("Nenhuma planilha encontrada no arquivo.")
            End If

            ' Adiciona colunas
            For col = 1 To planilha.Dimension.End.Column
                Dim nomeColuna = planilha.Cells(1, col).Text.Trim()
                If dt.Columns.Contains(nomeColuna) Then
                    nomeColuna &= "_" & col ' evita nome duplicado
                End If
                dt.Columns.Add(nomeColuna)
            Next

            ' Adiciona linhas
            For row = 2 To planilha.Dimension.End.Row
                Dim novaLinha = dt.NewRow()
                For col = 1 To dt.Columns.Count
                    novaLinha(col - 1) = planilha.Cells(row, col).Text.Trim()
                Next
                dt.Rows.Add(novaLinha)
            Next
        End Using

        ds.Tables.Add(dt)
        Return ds
    End Function

    Public Shared Function DecodificarDe128Bits(ByVal Valor As String) As String
        Try
            Return ""
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function BytesToString(ByVal Input As Byte()) As String
        Dim Result As New System.Text.StringBuilder(Input.Length * 2)
        Dim Part As String
        For Each b As Byte In Input
            Part = Conversion.Hex(b)
            If Part.Length = 1 Then Part = "0" & Part
            Result.Append(Part)
        Next
        Return "0x" & Result.ToString()
    End Function

    Public Shared Function NormalizeLote(ByVal lote As String, ByVal empresa As String) As String

        ' Remove letras e símbolos, mantendo apenas os números
        Dim onlyNumbers As String = Regex.Replace(lote, "\D", "")
        Dim normalizeString As String

        Dim empresas As List(Of String) = New List(Of String) From {"05272759"}

        If empresas.Contains(Left(empresa, 8)) Then
            normalizeString = lote
        Else
            ' Converte para inteiro para remover zeros à esquerda
            normalizeString = Long.Parse(onlyNumbers).ToString()
        End If


        Return normalizeString

    End Function

    Public Shared Function GerarNumeroSeguro(digitos As Integer) As String
        ' Array de bytes, um para cada dígito desejado
        Dim bytes(digitos - 1) As Byte

        ' Preenche com valores aleatórios
        Using rng As RandomNumberGenerator = RandomNumberGenerator.Create()
            rng.GetBytes(bytes)
        End Using

        ' Converte cada byte em um dígito (0 a 9)
        Dim sb As New StringBuilder()
        For Each b In bytes
            sb.Append((b Mod 10).ToString())
        Next

        Return sb.ToString()
    End Function

End Class