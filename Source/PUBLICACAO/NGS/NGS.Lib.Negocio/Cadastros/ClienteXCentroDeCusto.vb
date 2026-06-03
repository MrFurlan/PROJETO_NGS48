Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXCentroDeCusto
    Inherits List(Of ClienteXCentroDeCusto)

#Region "Contrutor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim sql As String

        sql = "SELECT Cliente_Id, EndCliente_Id, CentroDeCusto_Id, PercentualFixo" & vbCrLf & _
              "  FROM ClienteXCentroDeCusto" & vbCrLf & _
              " Where Cliente_Id    ='" & pCliente.Codigo & "'" & vbCrLf & _
              "   AND EndCliente_Id = " & pCliente.CodigoEndereco

        ds = Banco.ConsultaDataSet(sql, "CC")

        If Not ds Is Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim CC As New ClienteXCentroDeCusto(pCliente)
                CC.CodigoCentroDeCusto = row("CentroDeCusto_Id")
                CC.PercentualFixo = row("PercentualFixo")
                Me.Add(CC)
            Next
        End If

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
        For Each CxC As ClienteXCentroDeCusto In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXCentroDeCusto
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoCentroDeCusto As String
    Private _CentroDeCusto As CentroDeCusto
    Private _PercentualFixo As Decimal
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

    Public Property CodigoCentroDeCusto() As String
        Get
            Return _CodigoCentroDeCusto
        End Get
        Set(ByVal value As String)
            _CodigoCentroDeCusto = value
            _CentroDeCusto = Nothing
        End Set
    End Property

    Public ReadOnly Property CentroDeCusto() As CentroDeCusto
        Get
            If _CentroDeCusto Is Nothing And CodigoCentroDeCusto.Length > 0 Then _CentroDeCusto = New CentroDeCusto(CodigoCentroDeCusto)
            Return _CentroDeCusto
        End Get
    End Property

    Public ReadOnly Property DescricaoCentroCusto()
        Get
            Return CentroDeCusto.Descricao
        End Get
    End Property

    Public Property PercentualFixo() As Decimal
        Get
            Return _PercentualFixo
        End Get
        Set(ByVal value As Decimal)
            _PercentualFixo = value
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
                Sql = " INSERT INTO ClienteXCentroDeCusto(Cliente_Id, EndCliente_Id, CentroDeCusto_Id, PercentualFixo) " & vbCrLf & _
                      " VALUES ('" & Cliente.Codigo & "'," & Cliente.CodigoEndereco & "," & CodigoCentroDeCusto & "," & Str(PercentualFixo) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ClienteXCentroDeCusto SET" & vbCrLf & _
                      "    PercentualFixo        =" & Str(PercentualFixo) & vbCrLf & _
                      "  WHERE Cliente_Id       ='" & Cliente.Codigo & "'" & vbCrLf & _
                      "    AND EndCliente_Id    = " & Cliente.CodigoEndereco & vbCrLf & _
                      "    AND CentroDeCusto_Id ='" & CodigoCentroDeCusto & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ClienteXCentroDeCusto" & vbCrLf & _
                      "  WHERE Cliente_Id       ='" & Cliente.Codigo & "'" & vbCrLf & _
                      "    AND EndCliente_Id    = " & Cliente.CodigoEndereco & vbCrLf & _
                      "    AND CentroDeCusto_Id ='" & CodigoCentroDeCusto & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
