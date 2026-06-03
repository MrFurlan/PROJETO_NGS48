Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ListMeiosPagamento
    Inherits List(Of MeiosPagamento)
    Dim ds As DataSet
    Dim sql As String
    Dim objBanco As New AcessaBanco

#Region "Constructor"
    Public Sub New()
        Try
            sql = " SELECT Codigo_Id, Descricao, Ativo, ISNULL(QuemLancou, '') UsuarioInclusao, ISNULL(QuandoLancou, '') UsuarioInclusaoData, " & vbCrLf & _
                  " ISNULL(QuemAlterou, '') UsuarioAlteracao, ISNULL(QuandoAlterou, '') UsuarioAlteracaoData FROM MeiosDePagamento "

            ds = objBanco.ConsultaDataSet(sql, "MeiosDePagamento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objMeiosPagamento As New MeiosPagamento
                objMeiosPagamento.Codigo = row("Codigo_Id")
                objMeiosPagamento.Descricao = row("Descricao")
                objMeiosPagamento.Ativo = row("Ativo").Equals("S")
                objMeiosPagamento.UsuarioInclusao = row("UsuarioInclusao")
                objMeiosPagamento.UsuarioInclusaoData = row("UsuarioInclusaoData")
                objMeiosPagamento.UsuarioAlteracao = row("UsuarioAlteracao")
                objMeiosPagamento.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
                Me.Add(objMeiosPagamento)
            Next
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region

#Region "Methods"
    Public Sub salvar()
        For Each MeiosPagamento As MeiosPagamento In Me
            MeiosPagamento.salvar()
        Next
    End Sub
#End Region
End Class

<Serializable()> _
Public Class MeiosPagamento
    Implements IBaseEntity

#Region "Constructor"
    Dim ds As DataSet
    Dim sql As String
    Dim objBanco As New AcessaBanco

    Public Sub New()

    End Sub

    Public Sub New(codigo As String)
        Try
            sql = " SELECT Codigo_Id, Descricao, Ativo, ISNULL(QuemLancou, '') UsuarioInclusao, ISNULL(QuandoLancou, '') UsuarioInclusaoData, " & vbCrLf & _
                  " ISNULL(QuemAlterou, '') UsuarioAlteracao, ISNULL(QuandoAlterou, '') UsuarioAlteracaoData FROM MeiosDePagamento " & vbCrLf & _
                  " WHERE Codigo_Id = " & codigo

            ds = objBanco.ConsultaDataSet(sql, "MeiosDePagamento")

            For Each row As DataRow In ds.Tables(0).Rows
                _Codigo_Id = row("Codigo_Id")
                _Descricao = row("Descricao")
                _Ativo = row("Ativo").Equals("S")
                _UsuarioInclusao = row("UsuarioInclusao")
                _UsuarioInclusaoData = row("UsuarioInclusaoData")
                _UsuarioAlteracao = row("UsuarioAlteracao")
                _UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Next
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region


#Region "Fields"
    Private _IUD As String
    Private _Codigo_Id As Integer
    Private _Descricao As String
    Private _Ativo As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As datetime
#End Region

#Region "Property"
    Public WriteOnly Property IUD As String
        Set(value As String)
            _IUD = value
        End Set
    End Property
    Public Property Codigo As Integer
        Get
            Return _Codigo_Id
        End Get
        Set(value As Integer)
            _Codigo_Id = value
        End Set
    End Property
    Public Property Descricao As String
        Get
            Return _Descricao
        End Get
        Set(value As String)
            _Descricao = value
        End Set
    End Property
    Public Property Ativo As String
        Get
            Return IIf(_Ativo, "S", "N")
        End Get
        Set(value As String)
            _Ativo = value
        End Set
    End Property
    Public Property UsuarioInclusao As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(value As String)
            _UsuarioInclusao = value
        End Set
    End Property
    Public Property UsuarioInclusaoData As String
        Get
            If _UsuarioInclusaoData = "01/01/1900 00:00:00" Then
                Return ""
            End If
            Return _UsuarioInclusaoData
        End Get
        Set(value As String)
            _UsuarioInclusaoData = value
        End Set
    End Property
    Public Property UsuarioAlteracao As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(value As String)
            _UsuarioAlteracao = value
        End Set
    End Property
    Public Property UsuarioAlteracaoData As String
        Get
            If _UsuarioAlteracaoData = "01/01/1900 00:00:00" Then
                Return ""
            End If
            Return _UsuarioAlteracaoData
        End Get
        Set(value As String)
            _UsuarioAlteracaoData = value
        End Set
    End Property
    Private ReadOnly Property Data As String
        Get
            Return Date.Now.ToString("yyyy-MM-dd HH:mm:ss")
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub salvar()
        Try
            If Me._IUD = "I" Then
                sql = " INSERT INTO MeiosDePagamento VALUES (" & Me.Codigo & ",'" & Me.Descricao & "','" & Me.UsuarioInclusao & "','" & Me.Data & "', NULL, NULL,'" & Me.Ativo & "')"
            ElseIf Me._IUD = "U" Then
                sql = " UPDATE MeiosDePagamento SET Ativo='" & Me.Ativo & "',Descricao='" & Me.Descricao & "',QuemAlterou='" & Me.UsuarioAlteracao & "',QuandoAlterou='" & Me.Data & "' WHERE Codigo_Id=" & Me.Codigo
            ElseIf Me._IUD = "D" Then
                sql = " UPDATE MeiosDePagamento SET Ativo=0,QuemAlterou='" & Me.UsuarioAlteracao & "',QuandoAlterou='" & Me.Data & "' WHERE Codigo_Id=" & Me.Codigo
            Else
                MsgBox("VALOR IUD INVALIDO.")
                Me._IUD = String.Empty
                Exit Sub
            End If

            objBanco.GravaBanco(sql)
        Catch ex As Exception
            MsgBox("Houve um erro ao consultar as alterações. Erro: " & ex.Message)
        End Try
    End Sub
#End Region

End Class
