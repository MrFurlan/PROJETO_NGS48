Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXSafra
    Inherits List(Of ClienteXSafra)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "SELECT CxS.Cliente_Id, CxS.EndCliente_Id, CxS.Safra_Id, CxS.Cultura_Id, " & vbCrLf & _
              "       C.Descricao AS NomeCultura, CxS.AreaPlantada, CxS.Produtividade, " & vbCrLf & _
              "       CxS.ConsumoProprio, CxS.Comprometido, CxS.EstimativaEntrega, CxS.Observacao " & vbCrLf & _
              "  FROM ClienteXSafra AS CxS " & vbCrLf & _
              "       INNER JOIN Cultura AS C " & vbCrLf & _
              "       ON CxS.Cultura_Id = C.Cultura_Id " & vbCrLf & _
              " Where CxS.Cliente_Id    = '" & pcliente.Codigo & "' " & vbCrLf & _
              "   and CxS.EndCliente_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "ClienteXSafra")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CxA As New ClienteXSafra(pcliente)
            CxA.CodigoSafra = row("Safra_Id")
            CxA.CodigoCultura = row("Cultura_Id")
            CxA.NomeCultura = row("NomeCultura")
            CxA.AreaPlantada = row("AreaPlantada")
            CxA.Produtividade = row("Produtividade")
            CxA.ConsumoProprio = row("ConsumoProprio")
            CxA.EstimativaEntrega = row("EstimativaEntrega")
            CxA.Observacao = row("Observacao")
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
        For Each CxS As ClienteXSafra In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxS.IUD = _Cliente.IUD
            If CxS.IUD <> "" Then
                CxS.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXSafra
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
    Private _CodigoSafra As String
    Private _Safra As Safra
    Private _CodigoCultura As Integer
    Private _NomeCultura As String
    Private _Cultura As Cultura
    Private _AreaPlantada As Decimal
    Private _Produtividade As Decimal
    Private _ConsumoProprio As Decimal
    Private _Comprometido As Decimal
    Private _EstimativaEntrega As Decimal
    Private _Observacao As String
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

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property Safra() As Safra
        Get
            If _Safra Is Nothing And _CodigoSafra.Trim.Length > 0 Then _Safra = New Safra(_CodigoSafra)
            Return _Safra
        End Get
        Set(ByVal value As Safra)
            _Safra = value
        End Set
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
        End Set
    End Property

    Public Property NomeCultura() As String
        Get
            Return _NomeCultura
        End Get
        Set(ByVal value As String)
            _NomeCultura = value
        End Set
    End Property

    Public Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And _CodigoCultura Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura
        End Get
        Set(ByVal value As Cultura)
            _Cultura = value
        End Set
    End Property

    Public Property AreaPlantada() As Decimal
        Get
            Return _AreaPlantada
        End Get
        Set(ByVal value As Decimal)
            _AreaPlantada = value
        End Set
    End Property

    Public Property Produtividade() As Decimal
        Get
            Return _Produtividade
        End Get
        Set(ByVal value As Decimal)
            _Produtividade = value
        End Set
    End Property

    Public Property ConsumoProprio() As Decimal
        Get
            Return _ConsumoProprio
        End Get
        Set(ByVal value As Decimal)
            _ConsumoProprio = value
        End Set
    End Property

    Public Property Comprometido() As Decimal
        Get
            Return _Comprometido
        End Get
        Set(ByVal value As Decimal)
            _Comprometido = value
        End Set
    End Property

    Public Property EstimativaEntrega() As Decimal
        Get
            Return _EstimativaEntrega
        End Get
        Set(ByVal value As Decimal)
            _EstimativaEntrega = value
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
                sql = "Insert into ClienteXSafra (Cliente_Id, EndCliente_Id, Safra_Id, Cultura_Id, AreaPlantada, Produtividade, ConsumoProprio, Comprometido, EstimativaEntrega, Observacao)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _CodigoSafra & "'," & _CodigoCultura & "," & Str(_AreaPlantada) & "," & Str(_Produtividade) & "," & Str(_ConsumoProprio) & "," & Str(_Comprometido) & "," & Str(_EstimativaEntrega) & ",'" & _Observacao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXSafra set " & vbCrLf & _
                      "     AreaPlantada     = " & Str(_AreaPlantada) & vbCrLf & _
                      "   ,Produtividade     = " & Str(_Produtividade) & vbCrLf & _
                      "   ,ConsumoProprio    = " & Str(_ConsumoProprio) & vbCrLf & _
                      "   ,Comprometido      = " & Str(_Comprometido) & vbCrLf & _
                      "   ,EstimativaEntrega = " & Str(_EstimativaEntrega) & vbCrLf & _
                      "   ,Observacao        ='" & _Observacao & "'" & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Safra_Id      ='" & _CodigoSafra & "'" & vbCrLf & _
                      "   and Cultura_Id    = " & _CodigoCultura
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXSafra " & vbCrLf & _
                      " Where Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and Safra_Id      ='" & _CodigoSafra & "'" & vbCrLf & _
                      "   and Cultura_Id    = " & _CodigoCultura
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class