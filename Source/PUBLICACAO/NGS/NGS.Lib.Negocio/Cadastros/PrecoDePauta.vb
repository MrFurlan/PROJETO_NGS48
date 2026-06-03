Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPrecoDePauta
    Inherits List(Of PrecoDePauta)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False)
        If Not Carregar Then Exit Sub
        CarregarRegistros()
    End Sub

    Public Sub New(ByVal CodigoProduto As String, Optional ByVal CodigoEstado As String = "")
        CarregarRegistros(CodigoProduto, CodigoEstado)
    End Sub
#End Region

#Region "Methods"
    Private Sub CarregarRegistros(Optional ByVal CodigoProduto As String = "", Optional ByVal CodigoEstado As String = "")
        Dim sql As String
        sql = " SELECT Produto_Id, Estado_Id, Data_Id, Preco" & vbCrLf & _
              "   FROM PrecoDePauta" & vbCrLf & _
              "  Where 1 = 1 "

        If CodigoProduto.Trim.Length > 0 Then
            sql &= "  And Produto_id = '" & CodigoProduto & "'" & vbCrLf
        End If

        If CodigoEstado.Trim.Length > 0 Then
            sql &= "  And Estado_Id = '" & CodigoEstado & "'" & vbCrLf
        End If

        sql &= " Order By Produto_Id, Estado_Id, Data_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "PrecoDePauta")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PP As New PrecoDePauta
            PP.CodigoProduto = row("Produto_Id")
            PP.CodigoEstado = row("Estado_Id")
            PP.Data = row("Data_Id")
            PP.Preco = row("Preco")
            Me.Add(PP)
        Next


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
        For Each PP As PrecoDePauta In Me
            If PP.IUD <> "" Then
                PP.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PrecoDePauta
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigoEstado As String
    Private _Estado As Estado
    Private _Data As DateTime
    Private _Preco As Decimal
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

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public ReadOnly Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
    End Property

    Public Property CodigoEstado() As String
        Get
            Return _CodigoEstado
        End Get
        Set(ByVal value As String)
            _CodigoEstado = value
            _Estado = Nothing
        End Set
    End Property

    Public ReadOnly Property Estado() As Estado
        Get
            If _Estado Is Nothing And _CodigoEstado.Length > 0 Then _Estado = New Estado(_CodigoEstado)
            Return _Estado
        End Get
    End Property

    Public Property Data() As DateTime
        Get
            Return _Data
        End Get
        Set(ByVal value As DateTime)
            _Data = value
        End Set
    End Property

    Public Property Preco() As Decimal
        Get
            Return _Preco
        End Get
        Set(ByVal value As Decimal)
            _Preco = value
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "Insert into PrecoDePauta(Produto_Id, Estado_Id, Data_Id, Preco)" & vbCrLf & _
                      " values('" & _CodigoProduto & "','" & _CodigoEstado & "','" & _Data.ToString("yyyy-MM-dd") & "'," & Str(_Preco) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update PrecoDePauta set " & vbCrLf & _
                      "    Preco = " & Str(_Preco) & vbCrLf & _
                      " Where Produto_Id ='" & _CodigoProduto & "'" & vbCrLf & _
                      "   and Estado_Id  ='" & _CodigoEstado & "'" & vbCrLf & _
                      "   and Data_Id    ='" & _Data.ToString("yyyy-MM-dd") & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete PrecoDePauta " & vbCrLf & _
                      " Where Produto_Id ='" & _CodigoProduto & "'" & vbCrLf & _
                      "   and Estado_Id  ='" & _CodigoEstado & "'" & vbCrLf & _
                      "   and Data_Id    ='" & _Data.ToString("yyyy-MM-dd") & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class