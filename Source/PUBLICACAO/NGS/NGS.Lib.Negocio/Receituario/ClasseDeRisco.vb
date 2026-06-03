Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClasseDeRisco
    Inherits List(Of ClasseDeRisco)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal OrderBy As String = "")
        If CarregarDados Then
            Dim sql As String = "Select ClasseDeRisco_Id, Descricao From ClasseDeRisco "
            If OrderBy.Length > 0 Then
                sql &= "Order By " & OrderBy
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "ClasseDeRisco")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cr As New ClasseDeRisco
                Cr.Codigo = row("ClasseDeRisco_Id")
                Cr.Descricao = row("Descricao")
                Me.Add(Cr)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class ClasseDeRisco

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet("Select ClasseDeRisco_Id, Descricao From ClasseDeRisco Where ClasseDeRisco_Id = " & pCodigo, "ClasseDeRisco")

        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("ClasseDeRisco_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
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
                Sql = " INSERT INTO ClasseDeRisco (ClasseDeRisco_Id, Descricao) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ClasseDeRisco SET" & vbCrLf & _
                      "   Descricao        ='" & _Descricao & "'" & vbCrLf & _
                      "  WHERE ClasseDeRisco_Id =" & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ClasseDeRisco" & vbCrLf & _
                      "  WHERE ClasseDeRisco_Id =" & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
