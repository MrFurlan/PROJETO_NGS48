Imports Microsoft.VisualBasic
Imports System.Xml
Imports System.Data
Imports System.Web

Public NotInheritable Class BACEN

    Private Sub New()
    End Sub

    Public Shared Function getUltimosValores(ByVal startIndex As Integer, ByVal endIndex As Integer) As List(Of String)
        Dim lst As New List(Of String)
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim obj As br.gov.bcb.WSSerieVO = ws.getUltimosValoresSerieVO(startIndex, endIndex)
            lst = obj.valores.Select(Function(s) s.svalor).ToList()
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
        Return lst
    End Function

    Public Shared Function getValor(ByVal data As DateTime) As String
        Try
            Dim valor As String = String.Empty
            Dim diff As Integer = IIf(DateTime.Now.Date.Subtract(data.Date).Days <= 0, 1, DateTime.Now.Date.Subtract(data.Date).Days)
            If (diff > 20) Then
                Throw New Exception("Não é possível consultar a cotação do dólar na data informada!")
            End If

            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim obj As br.gov.bcb.WSSerieVO = ws.getUltimosValoresSerieVO(1, diff)
            For Each vlr As br.gov.bcb.WSValorSerieVO In obj.valores
                If (vlr.dia = data.Day AndAlso vlr.mes = data.Month AndAlso vlr.ano = data.Year) Then
                    valor = vlr.svalor
                End If
            Next

            If (String.IsNullOrWhiteSpace(valor)) Then
                Dim vlr As br.gov.bcb.WSValorSerieVO = obj.valores.Where(Function(x) New DateTime(x.ano, x.mes, x.dia) = obj.valores.Max(Function(s) New DateTime(s.ano, s.mes, s.dia))).FirstOrDefault()
                If (vlr IsNot Nothing) Then
                    valor = vlr.svalor
                End If
            End If
            Return valor
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getValor(ByVal serie As String, ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim valor As Decimal = ws.getValor(Long.Parse(serie), data.ToShortDateString())
            Return valor
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getValores(ByVal data As DateTime) As Dictionary(Of String, Decimal)
        Try
            Dim valores As New Dictionary(Of String, Decimal)
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrVenda As Decimal = ws.getValor(Long.Parse("1"), data.ToShortDateString())
            Dim vlrCompra As Decimal = ws.getValor(Long.Parse("10813"), data.ToShortDateString())
            Dim vlrMedio As Decimal = (vlrVenda + vlrCompra) / 2

            Console.WriteLine("Taxa de câmbio - Livre - Dólar americano (compra): " & vlrCompra)
            Console.WriteLine("Taxa de câmbio - Livre - Dólar médio PTAX: " & vlrMedio)
            Console.WriteLine("Taxa de câmbio - Livre - Dólar americano (venda): " & vlrVenda)

            valores("vlrCompra") = vlrCompra
            valores("vlrMedio") = vlrMedio
            valores("vlrVenda") = vlrVenda
            Return valores
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getVlrCompra(ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrCompra As Decimal = ws.getValor(Long.Parse("10813"), data.ToShortDateString())
            Return vlrCompra
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getVlrVenda(ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrVenda As Decimal = ws.getValor(Long.Parse("1"), data.ToShortDateString())
            Return vlrVenda
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getVlrMedio(ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrVenda As Decimal = ws.getValor(Long.Parse("1"), data.ToShortDateString())
            Dim vlrCompra As Decimal = ws.getValor(Long.Parse("10813"), data.ToShortDateString())
            Dim vlrMedio As Decimal = (vlrVenda + vlrCompra) / 2
            Return vlrMedio
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoValor() As String
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim obj As br.gov.bcb.WSSerieVO = ws.getUltimosValoresSerieVO(1, 1)
            Return obj.valores(0).svalor
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrCompra() As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim obj As br.gov.bcb.WSSerieVO = ws.getUltimoValorVO(Long.Parse("10813"))
            Dim vlrCompra As Decimal = Convert.ToDecimal(obj.ultimoValor.svalor.Replace(".", ","))
            Return vlrCompra
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrVenda() As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim obj As br.gov.bcb.WSSerieVO = ws.getUltimoValorVO(Long.Parse("1"))
            Dim vlrVenda As Decimal = Convert.ToDecimal(obj.ultimoValor.svalor.Replace(".", ","))
            Return vlrVenda
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrMedio() As Decimal
        Try
            Dim vlrCompra As Decimal = getUltimoVlrCompra()
            Dim vlrVenda As Decimal = getUltimoVlrVenda()
            Return (vlrCompra + vlrVenda) / 2
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrCompra(ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrCompra As Decimal = Decimal.Zero
            If (data.Date <= DateTime.Now.Date) Then
                If (data.Date.DayOfWeek = DayOfWeek.Sunday) Then
                    Try
                        vlrCompra = ws.getValor(Long.Parse("10813"), data.AddDays(-2).ToShortDateString())
                    Catch ex As Exception
                        vlrCompra = getUltimoVlrCompra()
                    End Try
                ElseIf (data.Date.DayOfWeek = DayOfWeek.Saturday) Then
                    Try
                        vlrCompra = ws.getValor(Long.Parse("10813"), data.AddDays(-1).ToShortDateString())
                    Catch ex As Exception
                        vlrCompra = getUltimoVlrCompra()
                    End Try
                Else
                    If (data.Date.DayOfWeek = DayOfWeek.Monday) Then
                        vlrCompra = getUltimoVlrCompra()
                    Else
                        vlrCompra = Decimal.Zero
                    End If
                End If
            End If
            Return vlrCompra
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrVenda(ByVal data As DateTime) As Decimal
        Try
            Dim ws As br.gov.bcb.FachadaWSSGSService = New br.gov.bcb.FachadaWSSGSService()
            Dim vlrVenda As Decimal = Decimal.Zero
            If (data.Date <= DateTime.Now.Date) Then
                If (data.Date.DayOfWeek = DayOfWeek.Sunday) Then
                    Try
                        vlrVenda = ws.getValor(Long.Parse("1"), data.AddDays(-2).ToShortDateString())
                    Catch ex As Exception
                        vlrVenda = getUltimoVlrVenda()
                    End Try
                ElseIf (data.Date.DayOfWeek = DayOfWeek.Saturday) Then
                    Try
                        vlrVenda = ws.getValor(Long.Parse("1"), data.AddDays(-1).ToShortDateString())
                    Catch ex As Exception
                        vlrVenda = getUltimoVlrVenda()
                    End Try
                Else
                    If (data.Date.DayOfWeek = DayOfWeek.Monday) Then
                        vlrVenda = getUltimoVlrVenda()
                    Else
                        vlrVenda = Decimal.Zero
                    End If
                End If
            End If
            Return vlrVenda
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

    Public Shared Function getUltimoVlrMedio(ByVal data As DateTime) As Decimal
        Try
            Dim vlrCompra As Decimal = getUltimoVlrCompra(data)
            Dim vlrVenda As Decimal = getUltimoVlrVenda(data)
            Dim vlrMedio As Decimal = (vlrCompra + vlrVenda) / 2
            Return vlrMedio
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Function

End Class