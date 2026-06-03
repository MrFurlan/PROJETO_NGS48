Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListComissoes
    Inherits List(Of Comissoes)
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()
    End Sub

    Public Sub New(ByVal pedido As [Lib].Negocio.Pedido)
        Dim sql As String = "SELECT * FROM Comissoes " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND Empresa_Id = '" & pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "AND EndEmpresa_Id = '" & pedido.EnderecoEmpresa & "' " & vbCrLf & _
                            "AND Pedido_Id = '" & pedido.Codigo & "' "

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "Comissoes")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New Comissoes
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Pedido_Id = row("Pedido_Id")
                obj.Representante_Id = row("Representante_Id")
                obj.EndRepresentante_Id = row("EndRepresentante_Id")
                obj.Percentual = row("Percentual")
                obj.ValorComissao = row("ValorComissao")
                obj.Principal = row("Principal")
                obj.PercentualFixo = row("PercentualFixo")
                obj.Liquidado = row("Liquidado")
                Me.Add(obj)
            Next
        End If
    End Sub
#End Region

#Region "Métodos"
    Public Function Salvar() As Boolean
        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        Return db.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each obj As Comissoes In Me
            If Not String.IsNullOrWhiteSpace(obj.IUD) Then
                obj.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class Comissoes
    Implements IBaseEntity

#Region "Propriedades"
    Private _IUD As String
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Private _Empresa_Id As String
    Public Property Empresa_Id() As String
        Get
            Return _Empresa_Id
        End Get
        Set(ByVal value As String)
            _Empresa_Id = value
        End Set
    End Property

    Private _EndEmpresa_Id As Integer
    Public Property EndEmpresa_Id() As Integer
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa_Id = value
        End Set
    End Property

    Private _Pedido_Id As String
    Public Property Pedido_Id() As String
        Get
            Return _Pedido_Id
        End Get
        Set(ByVal value As String)
            _Pedido_Id = value
        End Set
    End Property

    Private _Representante_Id As String
    Public Property Representante_Id() As String
        Get
            Return _Representante_Id
        End Get
        Set(ByVal value As String)
            _Representante_Id = value
        End Set
    End Property

    Private _EndRepresentante_Id As Integer
    Public Property EndRepresentante_Id() As Integer
        Get
            Return _EndRepresentante_Id
        End Get
        Set(ByVal value As Integer)
            _EndRepresentante_Id = value
        End Set
    End Property

    Private _Percentual As Decimal
    Public Property Percentual() As Decimal
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _Percentual = value
        End Set
    End Property

    Private _ValorComissao As Decimal
    Public Property ValorComissao() As Decimal
        Get
            Return _ValorComissao
        End Get
        Set(ByVal value As Decimal)
            _ValorComissao = value
        End Set
    End Property

    Private _Principal As String
    Public Property Principal() As String
        Get
            Return _Principal
        End Get
        Set(ByVal value As String)
            _Principal = value
        End Set
    End Property

    Private _PercentualFixo As Boolean
    Public Property PercentualFixo() As Boolean
        Get
            Return _PercentualFixo
        End Get
        Set(ByVal value As Boolean)
            _PercentualFixo = value
        End Set
    End Property

    Private _Liquidado As Boolean
    Public Property Liquidado() As Boolean
        Get
            Return _Liquidado
        End Get
        Set(ByVal value As Boolean)
            _Liquidado = value
        End Set
    End Property
#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        If String.IsNullOrWhiteSpace(IUD) Then Return True
        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        Return db.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " INSERT INTO Comissoes (Empresa_Id, EndEmpresa_Id, Pedido_Id, Representante_Id, EndRepresentante_Id, Percentual, ValorComissao, Principal, PercentualFixo, Liquidado) "
                sql &= "VALUES ('" & _Empresa_Id & "', '" & _EndEmpresa_Id & "', '" & _Pedido_Id & "', '" & _Representante_Id & "', '" & _EndRepresentante_Id & "', '" & Str(_Percentual) & "', '" & Str(_ValorComissao) & "', '" & _Principal & "', '" & _PercentualFixo & "', '" & _Liquidado & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE Comissoes SET Percentual = '" & Str(_Percentual) & "', ValorComissao = '" & Str(_Percentual) & "', Principal = '" & _Principal & "', PercentualFixo = '" & _PercentualFixo & "', Liquidado = '" & _Liquidado & "' "
                sql &= " WHERE Empresa_Id = '" & _Empresa_Id & "'"
                sql &= " AND EndEmpresa_Id = '" & _EndEmpresa_Id & "'"
                sql &= " AND Pedido_Id = '" & _Pedido_Id & "'"
                sql &= " AND Representante_Id = '" & _Representante_Id & "'"
                sql &= " AND EndRepresentante_Id = '" & _EndRepresentante_Id & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE FROM Comissoes"
                sql &= " WHERE Empresa_Id = '" & _Empresa_Id & "'"
                sql &= " AND EndEmpresa_Id = '" & _EndEmpresa_Id & "'"
                sql &= " AND Pedido_Id = '" & _Pedido_Id & "'"
                sql &= " AND Representante_Id = '" & _Representante_Id & "'"
                sql &= " AND EndRepresentante_Id = '" & _EndRepresentante_Id & "'"
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class
