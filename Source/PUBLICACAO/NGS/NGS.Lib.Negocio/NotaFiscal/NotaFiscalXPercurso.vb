Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Drawing
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXPercurso
    Inherits List(Of NotaFiscalXPercurso)
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New(ByVal nf As [Lib].Negocio.NotaFiscal)
        Dim sql As String = "SELECT * FROM NotasFiscaisXPercursos WHERE 1=1 " & vbCrLf & _
        "AND Empresa_Id = '" & nf.CodigoEmpresa & "' " & vbCrLf & _
        "AND EndEmpresa_Id = '" & nf.EnderecoEmpresa & "' " & vbCrLf & _
        "AND Cliente_Id = '" & nf.CodigoCliente & "' " & vbCrLf & _
        "AND EndCliente_Id = '" & nf.EnderecoCliente & "' " & vbCrLf & _
        "AND EntradaSaida_Id = '" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
        "AND Serie_Id = '" & nf.Serie & "' " & vbCrLf & _
        "AND Nota_Id = '" & nf.Codigo & "' " & vbCrLf

        Dim db As New AcessaBanco()
        Dim ds As DataSet = db.ConsultaDataSet(sql, "NotasFiscaisXPercursos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New NotaFiscalXPercurso()
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Cliente_Id = row("Cliente_Id")
                obj.EndCliente_Id = row("EndCliente_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                obj.Nota_Id = row("Nota_Id")
                obj.Serie_Id = row("Serie_Id")
                obj.Estado_Id = row("Estado_Id")
                obj.Ordem = row("Ordem")
                obj.NotaFiscal = nf
                Me.Add(obj)
            Next
        End If
    End Sub

    Public Sub New(ByVal e As [Lib].Negocio.Estado)
        Dim sql As String = "SELECT * FROM NotasFiscaisXPercursos WHERE 1=1 " & vbCrLf & _
        "AND Estado_Id = '" & e.Codigo & "' " & vbCrLf

        Dim db As New AcessaBanco()
        Dim ds As DataSet = db.ConsultaDataSet(sql, "NotasFiscaisXPercursos")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 And ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New NotaFiscalXPercurso()
                obj.Empresa_Id = row("Empresa_Id")
                obj.EndEmpresa_Id = row("EndEmpresa_Id")
                obj.Cliente_Id = row("Cliente_Id")
                obj.EndCliente_Id = row("EndCliente_Id")
                obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                obj.Nota_Id = row("Nota_Id")
                obj.Serie_Id = row("Serie_Id")
                obj.Estado_Id = row("Estado_Id")
                obj.Ordem = row("Ordem")
                Me.Add(obj)
            Next
        End If
    End Sub

#End Region

#Region "Métodos"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As NotaFiscalXPercurso In Me
            If String.IsNullOrWhiteSpace(item.IUD) Then
                item.IUD = item.NotaFiscal.IUD
            End If

            If item.IUD Is Nothing OrElse String.IsNullOrWhiteSpace(item.IUD) Then
                item.IUD = "I"
            End If

            If Not String.IsNullOrWhiteSpace(item.IUD) Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class NotaFiscalXPercurso
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()
    End Sub

#End Region

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

    Private _NotaFiscal As NotaFiscal
    Public Property NotaFiscal() As NotaFiscal
        Get
            If _NotaFiscal Is Nothing Then
                Dim nf As New [Lib].Negocio.NotaFiscal()
                nf.CodigoEmpresa = _Empresa_Id
                nf.EnderecoEmpresa = _EndEmpresa_Id
                nf.CodigoCliente = _Cliente_Id
                nf.EnderecoCliente = _EndCliente_Id
                nf.EntradaSaida = _EntradaSaida_Id
                nf.Serie = _Serie_Id
                nf.Codigo = _Nota_Id
                Return New [Lib].Negocio.NotaFiscal(nf)
            End If
            Return _NotaFiscal
        End Get
        Set(ByVal value As NotaFiscal)
            _NotaFiscal = value
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

    Private _EndEmpresa_Id As String
    Public Property EndEmpresa_Id() As String
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As String)
            _EndEmpresa_Id = value
        End Set
    End Property

    Private _Cliente_Id As String
    Public Property Cliente_Id() As String
        Get
            Return _Cliente_Id
        End Get
        Set(ByVal value As String)
            _Cliente_Id = value
        End Set
    End Property

    Private _EndCliente_Id As String
    Public Property EndCliente_Id() As String
        Get
            Return _EndCliente_Id
        End Get
        Set(ByVal value As String)
            _EndCliente_Id = value
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

    Private _Estado_Id As String
    Public Property Estado_Id() As String
        Get
            Return _Estado_Id
        End Get
        Set(ByVal value As String)
            _Estado_Id = value
        End Set
    End Property

    Private _Ordem As Integer
    Public Property Ordem() As Integer
        Get
            Return _Ordem
        End Get
        Set(ByVal value As Integer)
            _Ordem = value
        End Set
    End Property

#End Region

#Region "Métodos"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        Return db.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " DELETE NotasFiscaisXPercursos" & vbCrLf & _
                     "  WHERE 1=1 " & vbCrLf & _
                     "  AND Empresa_Id = '" & Empresa_Id & "' " & vbCrLf & _
                     "  AND EndEmpresa_Id = '" & EndEmpresa_Id & "' " & vbCrLf & _
                     "  AND Cliente_Id = '" & Cliente_Id & "' " & vbCrLf & _
                     "  AND EndCliente_Id = '" & EndCliente_Id & "' " & vbCrLf & _
                     "  AND EntradaSaida_Id = '" & IIf(EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                     "  AND Serie_Id = '" & Serie_Id & "' " & vbCrLf & _
                     "  AND Nota_Id = '" & Nota_Id & "' " & vbCrLf & _
                     "  AND Estado_Id = '" & Estado_Id & "' " & vbCrLf
                Sqls.Add(Sql)

                Sql = " INSERT INTO NotasFiscaisXPercursos (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Estado_Id, Ordem) " & vbCrLf & _
                      " VALUES ('" & Empresa_Id & "','" & EndEmpresa_Id & "','" & Cliente_Id & "','" & EndCliente_Id & "','" & IIf(EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "','" & Serie_Id & "','" & Nota_Id & "','" & Estado_Id & "','" & Ordem & "')"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE NotasFiscaisXPercursos" & vbCrLf & _
                      "  WHERE 1=1 " & vbCrLf & _
                      "  AND Empresa_Id = '" & Empresa_Id & "' " & vbCrLf & _
                      "  AND EndEmpresa_Id = '" & EndEmpresa_Id & "' " & vbCrLf & _
                      "  AND Cliente_Id = '" & Cliente_Id & "' " & vbCrLf & _
                      "  AND EndCliente_Id = '" & EndCliente_Id & "' " & vbCrLf & _
                      "  AND EntradaSaida_Id = '" & IIf(EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf & _
                      "  AND Serie_Id = '" & Serie_Id & "' " & vbCrLf & _
                      "  AND Nota_Id = '" & Nota_Id & "' " & vbCrLf & _
                      "  AND Estado_Id = '" & Estado_Id & "' " & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
