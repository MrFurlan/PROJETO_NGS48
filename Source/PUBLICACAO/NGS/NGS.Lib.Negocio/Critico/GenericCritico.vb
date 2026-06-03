<Serializable()>
Public Class ListGenericCritico
    Inherits List(Of GenericCritico)
#Region "Construtures"
    Public Sub New(ativo As Boolean)
        consulta("Ativo = " & ativo)
    End Sub

    Public Sub New(usuario As String)
        consulta("Usuario_ID = '" & usuario & "'")
    End Sub

    Public Sub New(movimento As DateTime)
        consulta("movimento = '" & movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'")
    End Sub

    Public Sub New(tipo As eCritico)
        consulta("tipo = " & tipo)
    End Sub

    Private Sub consulta(where As String)
        Dim sql As String

        sql = "SELECT * FROM Critico.Main " & vbCrLf & _
              "WHERE " & where

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "Critico.Main")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim objGenericCritico As New GenericCritico

            objGenericCritico.ID = row("ID")
            objGenericCritico.Ativo = row("Ativo")
            objGenericCritico.Usuario_ID = row("Usuario_ID")
            objGenericCritico.Movimento_ID = row("Movimento_ID")
            objGenericCritico.Tipo_ID = row("Tipo_ID")
            objGenericCritico.Mensagem = row("Mensagem")

            Me.add(objGenericCritico)
        Next
    End Sub
#End Region
End Class

<Serializable()>
Public Class GenericCritico
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()
    End Sub
#End Region

#Region "Fields"
    Private _IUD As Char()
    Private _ID As Integer
    Private _Ativo As Boolean
    Private _Usuario_ID As String
    Private _Movimento_ID As DateTime
    Private _Tipo_ID As eCritico
    Private _Mensagem As String
#End Region

#Region "Propriedades"
    Public Property IUD As Char()
        Get
            Return _IUD
        End Get
        Set(value As Char())
            _IUD = value
        End Set
    End Property

    Public Property ID As Integer
        Get
            Return _ID
        End Get
        Set(value As Integer)
            _ID = value
        End Set
    End Property

    Public Property Ativo As Boolean
        Get
            Return _Ativo
        End Get
        Set(value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property Usuario_ID As String
        Get
            Return _Usuario_ID
        End Get
        Set(value As String)
            _Usuario_ID = value
        End Set
    End Property

    Public Property Movimento_ID As Date
        Get
            Return _Movimento_ID
        End Get
        Set(value As Date)
            _Movimento_ID = value
        End Set
    End Property

    Public Property Tipo_ID As eCritico
        Get
            Return _Tipo_ID
        End Get
        Set(value As eCritico)
            _Tipo_ID = value
        End Set
    End Property

    Public Property Mensagem As String
        Get
            Return _Mensagem
        End Get
        Set(value As String)
            _Mensagem = value
        End Set
    End Property
#End Region

#Region "Metodos"
    Public Function genericLog() As Boolean
        Try
            Dim sql As String

            If _IUD = "I" Then
                sql = "INSERT INTO Critico.Main (Usuario_ID, Movimento_ID, Tipo_ID, Ativo, Mensagem) " & vbCrLf & _
                      "VALUES ('" & _Usuario_ID & "','" & _Movimento_ID.ToString("yyyy-MM-dd HH:mm:ss") & "','" & _Tipo_ID & "', 1, '" & _Mensagem & "')"
            ElseIf _IUD = "D" Then
                sql = " UPDATE Critico.Main SET Ativo = 0, DataAlteracao = '" & DateTime.Now() & "'" & vbCrLf & _
                      " WHERE ID = " & _ID
            Else
                Return False
            End If

            Dim banco As AcessaBanco = New AcessaBanco()
            banco.GravaBanco(sql)

            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
#End Region
End Class