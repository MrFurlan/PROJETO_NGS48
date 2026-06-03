Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClasseToxicologica
    Inherits List(Of ClasseToxicologica)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal OrderBy As String = "")
        If CarregarDados Then
            Dim sql As String = "Select ClasseTox_Id AS Codigo, Descricao, isnull(CorClasse,'') as CorClasse, isnull(CodCorClasse,'') as CodCorClasse  From ClasseToxicologica"
            If OrderBy.Length > 0 Then
                sql &= "Order By " & OrderBy
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "ClasseToxicologica")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ct As New ClasseToxicologica
                Ct.Codigo = row("Codigo")
                Ct.Descricao = row("Descricao")
                Ct.Cor = row("CorClasse")
                Ct.CodigoCor = row("CodCorClasse")
                Me.Add(Ct)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class ClasseToxicologica

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet("Select ClasseTox_Id, Descricao, isnull(CorClasse,'') as CorClasse, isnull(CodCorClasse,'') as CodCorClasse  From ClasseToxicologica Where ClasseTox_Id = " & pCodigo, "ClasseToxicologica")

        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("ClasseTox_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _Cor = ds.Tables(0).Rows(0)("CorClasse")
            _CodigoCor = ds.Tables(0).Rows(0)("CodCorClasse")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _Cor As String = ""
    Private _CodigoCor As String = ""
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

    Public Property Cor() As String
        Get
            Return _Cor
        End Get
        Set(ByVal value As String)
            _Cor = value
        End Set
    End Property

    Public Property CodigoCor() As String
        Get
            Return _CodigoCor
        End Get
        Set(ByVal value As String)
            _CodigoCor = value
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
                Sql = " INSERT INTO ClasseToxicologica (ClasseTox_Id, Descricao,CorClasse,CodCorClasse) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "','" & _Cor & "','" & _CodigoCor & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ClasseToxicologica SET" & vbCrLf & _
                      "   Descricao        ='" & _Descricao & "'" & vbCrLf & _
                      "  ,CorClasse        ='" & _Cor & "'" & vbCrLf & _
                      "  ,CodCorClasse     ='" & _CodigoCor & "'" & vbCrLf & _
                      "  WHERE ClasseTox_Id =" & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ClasseToxicologica" & vbCrLf & _
                      "  WHERE ClasseTox_Id =" & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class

