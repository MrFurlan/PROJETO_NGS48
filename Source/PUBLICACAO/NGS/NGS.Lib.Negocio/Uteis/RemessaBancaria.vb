Imports System.IO
Imports System.Text

Public Class RemessaBancaria

#Region "Contrutores"
    Public Sub New()
        Detalhes = New List(Of StringBuilder)
        Header = New StringBuilder
        Trailler = New StringBuilder
    End Sub
#End Region

#Region "Propriedades"

    Private _Header As StringBuilder
    Public Property Header() As StringBuilder
        Get
            Return _Header
        End Get
        Set(ByVal value As StringBuilder)
            _Header = value
        End Set
    End Property

    Private _Detalhe As StringBuilder
    Public Property Detalhe() As StringBuilder
        Get
            Return _Detalhe
        End Get
        Set(ByVal value As StringBuilder)
            _Detalhe = value
        End Set
    End Property

    Private Property Detalhes() As List(Of StringBuilder)

    Private _Trailler As StringBuilder
    Public Property Trailler() As StringBuilder
        Get
            Return _Trailler
        End Get
        Set(ByVal value As StringBuilder)
            _Trailler = value
        End Set
    End Property
#End Region

#Region "Métodos"

    Public Sub GerarArquivoDeRemessa(ByVal Path As String)
        'Kill(Path)

        Dim strm = New StreamWriter(Path, True)

        strm.WriteLine(_Header)

        For Each d In Detalhes
            strm.WriteLine(d)
        Next

        strm.WriteLine(_Trailler)

        strm.Close()
    End Sub

    Public Sub AddDetalhe(ByVal sb As StringBuilder)
        Detalhes.Add(sb)
    End Sub

    Public Function Format(
       ByVal BeginPosition As Integer,
       ByVal EndPosition As Integer,
       ByVal IsNumber As Boolean,
       ByVal Value As String) As String

        Return FitStringLength(
                 Value,
                 (EndPosition + 1) - BeginPosition,
                 (EndPosition + 1) - BeginPosition,
                  IIf(IsNumber, "0", " "),
                  0,
                  True,
                  True,
                  IsNumber).ToString().ToUpper()
    End Function

    Private Function FitStringLength(ByVal SringToBeFit As String, ByVal maxLength As Integer, ByVal minLength As Integer,
                                           ByVal FitChar As Char, ByVal maxStartPosition As Integer, ByVal maxTest As Boolean,
                                           ByVal minTest As Boolean, ByVal isNumber As Boolean) As String
        Try
            Dim result As String = ""

            If maxTest = True Then
                ' max
                If SringToBeFit.Length > maxLength Then
                    result += SringToBeFit.Substring(maxStartPosition, maxLength)
                End If
            End If

            If minTest = True Then
                ' min
                If SringToBeFit.Length <= minLength Then
                    If isNumber = True Then
                        result += DirectCast((New String(FitChar, (minLength - SringToBeFit.Length)) + SringToBeFit), String)
                    Else
                        result += DirectCast((SringToBeFit + New String(FitChar, (minLength - SringToBeFit.Length))), String)
                    End If
                End If
            End If
            Return result
        Catch ex As Exception
            Dim tmpEx As New Exception("Problemas ao Formatar a string. String = " & SringToBeFit, ex)
            Throw tmpEx
        End Try
    End Function

#End Region

End Class

