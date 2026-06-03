Imports System.Net
Imports System.Web.Http


Public Class ContatosController
        Inherits ApiController

        Public Property ContatosRepo As IContatosRepositorio

        ' Parameterless constructor
        Public Sub New()
        End Sub

        ' Constructor with dependency
        Public Sub New(ByVal _repo As IContatosRepositorio)
            ContatosRepo = _repo
        End Sub

        <HttpGet>
        Public Function GetTodos() As IEnumerable(Of Contato)
            Return ContatosRepo.GetTodos()
        End Function

        Public Function GetValue(ByVal id As Integer) As String
            Return "value"
        End Function

        Public Sub PostValue(<FromBody()> ByVal value As String)
        End Sub

        Public Sub PutValue(ByVal id As Integer, <FromBody()> ByVal value As String)
        End Sub

        Public Sub DeleteValue(ByVal id As Integer)
        End Sub

    End Class

