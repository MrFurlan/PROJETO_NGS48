Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXArrendante
    Inherits List(Of ClienteXArrendante)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CXA.Cliente_Id, CXA.EndCliente_Id, CXA.Arrendante_Id, CXA.EndArrendante_Id, Arrendante.Nome AS NomeArrendante, CXA.DataContrato_Id, CXA.DataVencimento, isnull(CXA.Area,0) as Area, CXA.Matricula_Id, CXA.Observacao, isnull(CXA.CustoArrendante,0) as CustoArrendante" & vbCrLf & _
              "  FROM ClienteXArrendante CXA " & vbCrLf & _
              " INNER JOIN Clientes AS Arrendante " & vbCrLf & _
              "    ON CXA.Arrendante_Id    = Arrendante.Cliente_Id " & vbCrLf & _
              "   AND CXA.EndArrendante_Id = Arrendante.Endereco_Id " & vbCrLf & _
              " Where CXA.Cliente_Id       ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and CXA.EndCliente_Id    = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXArrendante")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxA As New ClienteXArrendante(pcliente)
            CxA.CodigoArrendante = row("Arrendante_Id")
            CxA.EndArrendante = row("EndArrendante_Id")
            CxA.NomeArrendante = row("NomeArrendante")
            CxA.DataContrato = row("DataContrato_Id")
            CxA.DataVencimento = row("DataVencimento")
            CxA.Matricula = row("Matricula_Id")
            CxA.Area = row("Area")
            CxA.Observacao = row("Observacao")
            CxA.CustoArrendante = row("CustoArrendante")
            Me.Add(CxA)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Cliente As Cliente
#End Region

#Region "Property"
    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxA As ClienteXArrendante In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxA.IUD = _Cliente.IUD
            If CxA.IUD <> "" Then
                CxA.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXArrendante
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoArrendante As String = ""
    Private _EndArrendante As Integer
    Private _NomeArrendante As String
    Private _Arrendante As Cliente
    Private _Matricula As String = ""
    Private _DataContrato As DateTime
    Private _Area As Decimal
    Private _DataVencimento As DateTime
    Private _Observacao As String
    Private _CustoArrendante As Decimal
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

    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property CodigoArrendante() As String
        Get
            Return _CodigoArrendante
        End Get
        Set(ByVal value As String)
            _CodigoArrendante = value
            _Arrendante = Nothing
        End Set
    End Property

    Public Property EndArrendante() As Integer
        Get
            Return _EndArrendante
        End Get
        Set(ByVal value As Integer)
            _EndArrendante = value
            _Arrendante = Nothing
        End Set
    End Property

    Public Property NomeArrendante() As String
        Get
            Return _NomeArrendante
        End Get
        Set(ByVal value As String)
            _NomeArrendante = value
        End Set
    End Property

    Public Property Arrendante() As Cliente
        Get
            If _Arrendante Is Nothing And _CodigoArrendante.Trim.Length > 0 Then _Arrendante = New Cliente(_CodigoArrendante, _EndArrendante)
            Return _Arrendante
        End Get
        Set(ByVal value As Cliente)
            _Arrendante = value
        End Set
    End Property

    Public Property Matricula() As String
        Get
            Return _Matricula
        End Get
        Set(ByVal value As String)
            _Matricula = value
        End Set
    End Property

    Public Property DataContrato() As DateTime
        Get
            Return _DataContrato
        End Get
        Set(ByVal value As DateTime)
            _DataContrato = value
        End Set
    End Property

    Public Property Area() As Decimal
        Get
            Return _Area
        End Get
        Set(ByVal value As Decimal)
            _Area = value
        End Set
    End Property

    Public Property DataVencimento() As DateTime
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As DateTime)
            _DataVencimento = value
        End Set
    End Property

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property

    Public Property CustoArrendante() As Decimal
        Get
            Return _CustoArrendante
        End Get
        Set(ByVal value As Decimal)
            _CustoArrendante = value
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
                sql = "Insert into ClienteXArrendante(Cliente_Id, EndCliente_Id, Arrendante_Id, EndArrendante_Id,  DataContrato_Id, DataVencimento, Matricula_Id, Area, Observacao, CustoArrendante)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoArrendante & "'," & _EndArrendante & ",'" & _DataContrato.ToString("yyyy-MM-dd") & "','" & _DataVencimento.ToString("yyyy-MM-dd") & "','" & _Matricula & "'," & Str(_Area) & ",'" & _Observacao & "'," & Str(_CustoArrendante) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXArrendante set " & vbCrLf & _
                      "    DataVencimento  ='" & _DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Area            = " & Str(_Area) & vbCrLf & _
                      "   ,Observacao      ='" & _Observacao & "'" & vbCrLf & _
                      "   ,CustoArrendante = " & Str(_CustoArrendante) & vbCrLf & _
                      " Where Cliente_Id       ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id    = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Arrendante_Id    ='" & _CodigoArrendante & "'" & vbCrLf & _
                      "   and EndArrendante_Id = " & _EndArrendante & vbCrLf & _
                      "   and DataContrato_Id  ='" & _DataContrato.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   And Matricula_Id     ='" & _Matricula & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXArrendante " & vbCrLf & _
                      " Where Cliente_Id       ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id    = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Arrendante_Id    ='" & _CodigoArrendante & "'" & vbCrLf & _
                      "   and EndArrendante_Id = " & _EndArrendante & vbCrLf & _
                      "   and DataContrato_Id  ='" & _DataContrato.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   And Matricula_Id     ='" & _Matricula & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class