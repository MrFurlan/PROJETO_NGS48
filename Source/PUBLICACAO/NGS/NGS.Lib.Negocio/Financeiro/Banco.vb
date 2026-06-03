Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListBanco
    Inherits List(Of Banco)

#Region "Fields"
    Public Erro As Exception
#End Region

#Region "Construtor"
    Public Sub New(Optional ByVal carregar As Boolean = False)
        If Not carregar Then Exit Sub
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Banco_Id, Descricao, QuemLancou, QuandoLancou, QuemAlterou, QuandoAlterou, Ativo, isnull(LiquidacaoDias,0) as LiquidacaoDias" & _
                                   "  FROM Bancos " & _
                                   " ORDER BY Banco_Id"

            Dim dsBancos As DataSet = objBanco.ConsultaDataSet(strSQL, "Bancos")

            For Each drBanco As DataRow In dsBancos.Tables(0).Rows
                Dim objBancoConta As New Banco()

                objBancoConta.Codigo = drBanco("Banco_Id")
                objBancoConta.Descricao = drBanco("Descricao").ToString()
                If Not drBanco.IsNull("QuemLancou") Then objBancoConta.QuemLancou = drBanco("QuemLancou").ToString()
                If Not drBanco.IsNull("QuandoLancou") Then objBancoConta.QuandoLancou = Convert.ToDateTime(drBanco("QuandoLancou"))
                If Not drBanco.IsNull("QuemAlterou") Then objBancoConta.QuemAlterou = drBanco("QuemAlterou").ToString()
                If Not drBanco.IsNull("QuandoAlterou") Then objBancoConta.QuandoAlterou = Convert.ToDateTime(drBanco("QuandoAlterou"))
                objBancoConta.Ativo = (drBanco("Ativo").ToString() = "S")
                objBancoConta.LiquidacaoDias = drBanco("LiquidacaoDias")

                Me.Add(objBancoConta)
            Next

        Catch ex As Exception
            Me.Erro = ex
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

End Class

<Serializable()> _
Public Class Banco
    Implements IBaseEntity

#Region "Fields"
    Private _Codigo As Integer
    Private _Descricao As String
    Private _QuemLancou As String
    Private _QuandoLancou As DateTime
    Private _QuemAlterou As String
    Private _QuandoAlterou As DateTime
    Private _Ativo As Boolean
    Private _LiquidacaoDias As Integer

    Public Erro As Exception
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Me.Codigo = Codigo
        Selecionar(Codigo)
    End Sub
#End Region

#Region "Property"
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

    Public Property QuemLancou() As String
        Get
            Return _QuemLancou
        End Get
        Set(ByVal value As String)
            _QuemLancou = value
        End Set
    End Property

    Public Property QuandoLancou() As DateTime
        Get
            Return _QuandoLancou
        End Get
        Set(ByVal value As DateTime)
            _QuandoLancou = value
        End Set
    End Property

    Public Property QuemAlterou() As String
        Get
            Return _QuemAlterou
        End Get
        Set(ByVal value As String)
            _QuemAlterou = value
        End Set
    End Property

    Public Property QuandoAlterou() As DateTime
        Get
            Return _QuandoAlterou
        End Get
        Set(ByVal value As DateTime)
            _QuandoAlterou = value
        End Set
    End Property

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property LiquidacaoDias As Integer
        Get
            Return _LiquidacaoDias
        End Get
        Set(value As Integer)
            _LiquidacaoDias = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Banco_Id, Descricao, QuemLancou, QuandoLancou, QuemAlterou, QuandoAlterou, Ativo, isnull(LiquidacaoDias,0) as LiquidacaoDias " & _
                                   "  FROM Bancos " & _
                                   " WHERE Banco_Id = " & Codigo.ToString()

            Dim dsBancos As DataSet = objBanco.ConsultaDataSet(strSQL, "Bancos")

            If dsBancos.Tables(0).Rows.Count > 0 Then
                Dim drBanco As DataRow = dsBancos.Tables(0).Rows(0)

                Me.Codigo = Convert.ToInt32(drBanco("Banco_Id"))
                Me.Descricao = drBanco("Descricao").ToString()
                If Not drBanco.IsNull("QuemLancou") Then Me.QuemLancou = drBanco("QuemLancou").ToString()
                If Not drBanco.IsNull("QuandoLancou") Then Me.QuandoLancou = Convert.ToDateTime(drBanco("QuandoLancou"))
                If Not drBanco.IsNull("QuemAlterou") Then Me.QuemAlterou = drBanco("QuemAlterou").ToString()
                If Not drBanco.IsNull("QuandoAlterou") Then Me.QuandoAlterou = Convert.ToDateTime(drBanco("QuandoAlterou"))
                Me.Ativo = (drBanco("Ativo").ToString() = "S")
                Me.LiquidacaoDias = drBanco("LiquidacaoDias")
            End If

            Return True
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function
#End Region

End Class