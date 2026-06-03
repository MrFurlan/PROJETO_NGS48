Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListEmbalagem
    Inherits List(Of Embalagem)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        If CarregarDados Then
            Dim sql As String = "Select Embalagem_Id, Descricao, isnull(EmbalagemIndea,'') as EmbalagemIndea from Embalagens"
            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Embalagem")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Em As New Embalagem
                Em.Codigo = row("Embalagem_Id")
                Em.Descricao = row("Descricao")
                Em.EmbalagemIndea = row("EmbalagemIndea")
                Me.Add(Em)
            Next

        End If
    End Sub
End Class

<Serializable()> _
Public Class Embalagem

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Embalagem_Id, Descricao, Isnull(EmbalagemIndea,'') AS EmbalagemIndea from Embalagens where Embalagem_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Embalagem")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Embalagem_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _EmbalagemIndea = ds.Tables(0).Rows(0)("EmbalagemIndea")
        End If
    End Sub

    Public Sub New(ByVal pCodigoIndea As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Embalagem_Id, Descricao, EmbalagemIndea from Embalagens where EmbalagemIndea = '" & pCodigoIndea & "'"

        ds = Banco.ConsultaDataSet(sql, "Embalagem")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Embalagem_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _EmbalagemIndea = ds.Tables(0).Rows(0)("EmbalagemIndea")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Descricao As String = ""
    Private _EmbalagemIndea As String = ""
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property EmbalagemIndea() As String
        Get
            Return _EmbalagemIndea
        End Get
        Set(ByVal value As String)
            _EmbalagemIndea = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO Embalagens(Embalagem_Id, Descricao, EmbalagemIndea) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "','" & _EmbalagemIndea & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Embalagens SET" & vbCrLf & _
                      "    Descricao      ='" & _Descricao & "'" & vbCrLf & _
                      "   ,EmbalagemIndea ='" & _EmbalagemIndea & "'" & vbCrLf & _
                      "  WHERE Embalagem_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Embalagens" & vbCrLf & _
                      "  WHERE Embalagem_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
