Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClasseDeOperacao
    Inherits List(Of ClasseDeOperacao)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False, Optional ByVal Where As String = "")
        If Not Carregar Then Exit Sub

        Dim objBanco As New AcessaBanco()
        Dim strSQL As String = "SELECT Classe_Id, Descricao, isnull(Operacao,0) as Operacao, isnull(SubOperacao,0) as SubOperacao" & _
                               "  FROM ClassesDeOperacoes "
        If Where.Length > 0 Then strSQL &= " Where " & Where
        strSQL &= " ORDER BY Descricao"
        Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "ClasseDeOperacao")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Co As New ClasseDeOperacao
            Co.Codigo = row("Classe_Id")
            Co.Descricao = row("Descricao")
            Co.Operacao = row("Operacao")
            Co.SubOperacao = row("SubOperacao")
            Add(Co)
        Next
        objBanco = Nothing
    End Sub


#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Row As ClasseDeOperacao In Me
            If Row.IUD <> "" Then
                Row.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClasseDeOperacao
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As String)
        Dim objBanco As New AcessaBanco()
        Dim strSQL As String = "SELECT Classe_Id, Descricao, isnull(Operacao,0) as Operacao,  isnull(SubOperacao,0) as SubOperacao" & _
                               "  FROM ClassesDeOperacoes" & _
                               " Where Classe_Id = '" & Codigo & "' "
        Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "ClasseDeOperacao")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim Co As New ClasseDeOperacao
            Co.Codigo = row("Classe_Id")
            Co.Descricao = row("Descricao")
            Co.Operacao = row("Operacao")
            Co.SubOperacao = row("SubOperacao")
        Next
        objBanco = Nothing
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As String
    Private _Descricao As String
    Private _Operacao As Boolean
    Private _SubOperacao As Boolean
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

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
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

    Public Property Operacao() As Boolean
        Get
            Return _Operacao
        End Get
        Set(ByVal value As Boolean)
            _Operacao = value
        End Set
    End Property

    Public Property SubOperacao() As Boolean
        Get
            Return _SubOperacao
        End Get
        Set(ByVal value As Boolean)
            _SubOperacao = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into ClassesDeOperacoes(Classe_Id, Descricao, Operacao, SubOperacao)" & vbCrLf & _
                      " Values ('" & _Codigo & "','" & _Descricao & "'," & CByte(_Operacao) & "," & CByte(_SubOperacao) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClassesDeOperacoes set" & vbCrLf & _
                      "  Descricao   ='" & _Descricao & "'" & vbCrLf & _
                      " ,Operacao    = " & CByte(_Operacao) & vbCrLf & _
                      " ,SubOperacao = " & CByte(_SubOperacao) & vbCrLf & _
                      " Where Classe_id = '" & _Codigo & "'"
            Case "D"
                sql = " Delete ClassesDeOperacoes" & vbCrLf & _
                      "	 Where Classe_Id      ='" & _Codigo & "'" & vbCrLf
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
