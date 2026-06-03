Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListComissoesXBaixas
    Inherits List(Of ComissoesXBaixas)
    Implements IBaseEntity

#Region "Construtores"
    Public Sub New()
    End Sub

    Public Sub New(ByVal pedido As [Lib].Negocio.Pedido)
        Dim sql As String = "SELECT * FROM ComissoesXBaixas " & vbCrLf & _
                            "WHERE 1=1 " & vbCrLf & _
                            "AND Empresa_Id = '" & pedido.CodigoEmpresa & "' " & vbCrLf & _
                            "AND EndEmpresa_Id = '" & pedido.EnderecoEmpresa & "' " & vbCrLf & _
                            "AND Pedido_Id = '" & pedido.Codigo & "' "

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "ComissoesXBaixas")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New ComissoesXBaixas
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Pedido_Id = row("Pedido_Id")
                obj.Representante_Id = row("Representante_Id")
                obj.EndRepresentante_Id = row("EndRepresentante_Id")
                obj.EmpresaNota_Id = row("EmpresaNota_Id")
                obj.EndEmpresaNota_Id = row("EndEmpresaNota_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                obj.Serie_Id = row("Serie_Id")
                obj.Nota_Id = row("Nota_Id")
                obj.Valor = row("Valor")
                Me.Add(obj)
            Next
        End If
    End Sub

    Public Sub New(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal)
        Dim sql As String = "SELECT * FROM ComissoesXBaixas " & vbCrLf & _
                           "WHERE 1=1 " & vbCrLf & _
                           "AND EmpresaNota_Id = '" & objNotaFiscal.CodigoEmpresa & "' " & vbCrLf & _
                           "AND EndEmpresaNota_Id = '" & objNotaFiscal.EnderecoEmpresa & "' " & vbCrLf & _
                           "AND EntradaSaida_Id = '" & IIf(objNotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                           "AND Serie_Id = '" & objNotaFiscal.Serie & "' " & vbCrLf & _
                           "AND Nota_Id = '" & objNotaFiscal.Codigo & "' "

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "ComissoesXBaixas")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New ComissoesXBaixas
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Pedido_Id = row("Pedido_Id")
                obj.Representante_Id = row("Representante_Id")
                obj.EndRepresentante_Id = row("EndRepresentante_Id")
                obj.EmpresaNota_Id = row("EmpresaNota_Id")
                obj.EndEmpresaNota_Id = row("EndEmpresaNota_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                obj.Serie_Id = row("Serie_Id")
                obj.Nota_Id = row("Nota_Id")
                obj.Valor = row("Valor")
                Me.Add(obj)
            Next
        End If
    End Sub

    Public Sub New(ByVal objComissao As [Lib].Negocio.Comissoes)
        Dim sql As String = "SELECT * FROM ComissoesXBaixas " & vbCrLf & _
                           "WHERE 1=1 " & vbCrLf & _
                           "AND Empresa_Id = '" & objComissao.Empresa_Id & "' " & vbCrLf & _
                           "AND EndEmpresa_Id = '" & objComissao.EndEmpresa_Id & "' " & vbCrLf & _
                           "AND Pedido_Id = '" & objComissao.Pedido_Id & "' " & vbCrLf & _
                           "AND Representante_Id = '" & objComissao.Representante_Id & "' " & vbCrLf & _
                           "AND EndRepresentante_Id = '" & objComissao.EndRepresentante_Id & "' "

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "ComissoesXBaixas")

        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New ComissoesXBaixas
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Pedido_Id = row("Pedido_Id")
                obj.Representante_Id = row("Representante_Id")
                obj.EndRepresentante_Id = row("EndRepresentante_Id")
                obj.EmpresaNota_Id = row("EmpresaNota_Id")
                obj.EndEmpresaNota_Id = row("EndEmpresaNota_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                obj.Serie_Id = row("Serie_Id")
                obj.Nota_Id = row("Nota_Id")
                obj.Valor = row("Valor")
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
        For Each obj As ComissoesXBaixas In Me
            If Not String.IsNullOrWhiteSpace(obj.IUD) Then
                obj.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ComissoesXBaixas
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

    Private _EmpresaNota_Id As String
    Public Property EmpresaNota_Id() As String
        Get
            Return _EmpresaNota_Id
        End Get
        Set(ByVal value As String)
            _EmpresaNota_Id = value
        End Set
    End Property

    Private _EndEmpresaNota_Id As Integer
    Public Property EndEmpresaNota_Id() As Integer
        Get
            Return _EndEmpresaNota_Id
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaNota_Id = value
        End Set
    End Property

    Private _EntradaSaida_Id As eEntradaSaida
    Public Property EntradaSaida_Id() As eEntradaSaida
        Get
            Return _EntradaSaida_Id
        End Get
        Set(ByVal value As eEntradaSaida)
            _EntradaSaida_Id = value
        End Set
    End Property

    Private _Serie_Id As String
    Public Property Serie_Id() As String
        Get
            Return _Serie_Id
        End Get
        Set(ByVal value As String)
            _Serie_Id = value
        End Set
    End Property

    Private _Nota_Id As String
    Public Property Nota_Id() As String
        Get
            Return _Nota_Id
        End Get
        Set(ByVal value As String)
            _Nota_Id = value
        End Set
    End Property

    Private _Valor As Decimal
    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
        End Set
    End Property

    Private _NotaFiscal As [Lib].Negocio.NotaFiscal
    Public Property NotaFiscal() As [Lib].Negocio.NotaFiscal
        Get
            Return _NotaFiscal
        End Get
        Set(ByVal value As [Lib].Negocio.NotaFiscal)
            _NotaFiscal = value
        End Set
    End Property
#End Region

#Region "Construtores"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objNotaFiscal As [Lib].Negocio.NotaFiscal)
        _NotaFiscal = objNotaFiscal
    End Sub
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
                sql = " INSERT INTO ComissoesXBaixas (Empresa_Id, EndEmpresa_Id, Pedido_Id, Representante_Id, EndRepresentante_Id, EmpresaNota_Id, EndEmpresaNota_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Valor) "
                sql &= "VALUES ('" & _Empresa_Id & "', '" & _EndEmpresa_Id & "', '" & _Pedido_Id & "', '" & _Representante_Id & "', '" & _EndRepresentante_Id & "', '" & NotaFiscal.CodigoEmpresa & "', '" & NotaFiscal.EnderecoEmpresa & "', '" & IIf(NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & NotaFiscal.Serie & "', '" & NotaFiscal.Codigo & "', " & Str(_Valor) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE ComissoesXBaixas SET Valor = " & Str(_Valor) & " "
                sql &= " WHERE Empresa_Id = '" & _Empresa_Id & "'"
                sql &= " AND EndEmpresa_Id = '" & _EndEmpresa_Id & "'"
                sql &= " AND Pedido_Id = '" & _Pedido_Id & "'"
                sql &= " AND Representante_Id = '" & _Representante_Id & "'"
                sql &= " AND EndRepresentante_Id = '" & _EndRepresentante_Id & "'"
                sql &= " AND EmpresaNota_Id = '" & NotaFiscal.CodigoEmpresa & "'"
                sql &= " AND EndEmpresaNota_Id = '" & NotaFiscal.EnderecoEmpresa & "'"
                sql &= " AND EntradaSaida_Id = '" & IIf(NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"
                sql &= " AND Serie_Id = '" & NotaFiscal.Serie & "'"
                sql &= " AND Nota_Id = '" & NotaFiscal.Codigo & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE FROM ComissoesXBaixas"
                sql &= " WHERE Empresa_Id = '" & _Empresa_Id & "'"
                sql &= " AND EndEmpresa_Id = '" & _EndEmpresa_Id & "'"
                sql &= " AND Pedido_Id = '" & _Pedido_Id & "'"
                sql &= " AND Representante_Id = '" & _Representante_Id & "'"
                sql &= " AND EndRepresentante_Id = '" & _EndRepresentante_Id & "'"
                sql &= " AND EmpresaNota_Id = '" & NotaFiscal.CodigoEmpresa & "'"
                sql &= " AND EndEmpresaNota_Id = '" & NotaFiscal.EnderecoEmpresa & "'"
                sql &= " AND EntradaSaida_Id = '" & IIf(NotaFiscal.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'"
                sql &= " AND Serie_Id = '" & NotaFiscal.Serie & "'"
                sql &= " AND Nota_Id = '" & NotaFiscal.Codigo & "'"
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class