Public Class FuncoesLogs
    Private _operacao As String
    Private _tipo As Integer
    Private _obj As Object
    Private _empresa As String

    Public Sub New(tipo As Integer, empresa As String)
        _tipo = tipo
        _empresa = empresa
    End Sub

    Public Sub RegistrarLog(operacao As String, obj As Object)
        _obj = obj
        _operacao = operacao

        Select Case _tipo
            Case 1
                NotaFiscalXItens()
            Case 2
                NotasFiscaisGerais()
        End Select
    End Sub

    Private Function NotaFiscalXItens()
        Dim info As String = _obj.Cliente.Codigo & " - " &
                             _obj.Cliente.Nome & " - " &
                             _obj.Cliente.Cidade & "/" & _obj.Cliente.Estado.Descricao & " - " &
                             _obj.Codigo & " - " &
                             _obj.Serie & " - "


        If _operacao = "I" Then
            info &= " GRAVOU NOTA FISCAL "
        ElseIf _operacao = "U" Then
            info &= " ATUALIZOU NOTA FISCAL "
        ElseIf _operacao = "D" Then
            info &= " DELETOU NOTA FISCAL "
        End If

        Dim log As Logs = New Logs()
        log.Log_id = _empresa
        log.Tipo = "NotaFiscalXItens"
        log.GetLog(info)
        If log.Gravar() Then
            Return True
        End If
        Return False
    End Function

    Private Function NotasFiscaisGerais()
        Dim info As String = _obj.Cliente.Codigo & " - " &
                             _obj.Cliente.Nome & " - " &
                             _obj.Cliente.Cidade & "/" & _obj.Cliente.Estado.Descricao & " - " &
                             _obj.Codigo & " - " &
                             _obj.Serie & " - "


        If _operacao = "I" Then
            info &= " GRAVOU NOTA FISCAL "
        ElseIf _operacao = "U" Then
            info &= " ATUALIZOU NOTA FISCAL "
        ElseIf _operacao = "D" Then
            info &= " DELETOU NOTA FISCAL "
        End If

        Dim log As Logs = New Logs()
        log.Log_id = _empresa
        log.Tipo = "NotasFiscaisGerais"
        log.GetLog(info)
        If log.Gravar() Then
            Return True
        End If
        Return False
    End Function
End Class
