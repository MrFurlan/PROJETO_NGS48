Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class NotaFiscalXDestino

#Region "Fields"
    Private _NF As NotaFiscal
    Private _ItemNotaFiscal As NotaFiscalXItem
    Private _IUD As String
    Private _PesoBruto As Decimal
    Private _Desconto As Decimal
    Private _PesoLiquido As Decimal
    Private _Movimento As DateTime
    Private _Sinistro As Boolean
    Private _TarifaFrete As Decimal
#End Region

#Region "Property"
    Public Property NF() As NotaFiscal
        Get
            Return _NF
        End Get
        Set(ByVal value As NotaFiscal)
            _NF = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property PesoBruto() As Decimal
        Get
            Return _PesoBruto
        End Get
        Set(ByVal value As Decimal)
            _PesoBruto = value
        End Set
    End Property

    Public Property Desconto() As Decimal
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Decimal)
            _Desconto = value
        End Set
    End Property

    Public Property PesoLiquido() As Decimal
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Decimal)
            _PesoLiquido = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property Sinistro() As Boolean
        Get
            Return _Sinistro
        End Get
        Set(ByVal value As Boolean)
            _Sinistro = value
        End Set
    End Property

    Public Property TarifaFrete() As Decimal
        Get
            Return _TarifaFrete
        End Get
        Set(ByVal value As Decimal)
            _TarifaFrete = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal pNF As NotaFiscal, Optional ByVal CarregarDados As Boolean = False)
        Me.NF = pNF

        If CarregarDados Then
            CarregarNotaFiscalXDestino()
        End If
    End Sub
#End Region

#Region "Methods"

    Private Sub CarregarNotaFiscalXDestino()
        Try
            Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
                                "       PesoBruto, Desconto, PesoLiquido, Movimento, Sinistro " & vbCrLf & _
                                "  FROM NotasXDestinos " & vbCrLf & _
                                " WHERE Empresa_Id      ='" & Me.NF.CodigoEmpresa & "' " & vbCrLf & _
                                "   AND EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                                "   AND Cliente_Id      ='" & Me.NF.CodigoCliente & "' " & vbCrLf & _
                                "   AND EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                                "   AND EntradaSaida_Id ='" & Me.NF.EntradaSaida.ToString.Substring(0, 1) & "' " & vbCrLf & _
                                "   AND Serie_Id        ='" & Me.NF.Serie & "' " & vbCrLf & _
                                "   AND Nota_id         = " & Me.NF.Codigo & vbCrLf
            Dim db As New AcessaBanco
            Dim ds As DataSet = db.ConsultaDataSet(sql, "NotasXDestinos")

            For Each row As DataRow In ds.Tables(0).Rows
                Me.PesoBruto = row("PesoBruto")
                Me.Desconto = row("Desconto")
                Me.PesoLiquido = row("PesoLiquido")
                Me.Movimento = row("Movimento")
                Me.Sinistro = Not IsDBNull(row("Sinistro")) AndAlso row("Sinistro") = "S"
            Next
        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)

        Dim sql As String

        If IUD Is Nothing AndAlso Me.NF.IUD = "D" Then
            IUD = Me.NF.IUD
        End If

        Select Case IUD
            Case "I"
                sql = "INSERT INTO NotasXDestinos (Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                                                  "Cliente_Id, EndCliente_Id, " & vbCrLf &
                                                  "EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf &
                                                  "PesoBruto, Desconto, " & vbCrLf &
                                                  "PesoLiquido, Movimento, Sinistro, TarifaFrete)" & vbCrLf &
                       "                   VALUES ('" & Me.NF.CodigoEmpresa & "', " & Me.NF.EnderecoEmpresa & ", " & vbCrLf &
                                                  "'" & Me.NF.CodigoCliente & "', " & Me.NF.EnderecoCliente & ", " & vbCrLf &
                                                  "'" & Me.NF.EntradaSaida.ToString.Substring(0, 1) & "', '" & Me.NF.Serie & "', " & Me.NF.Codigo & ", " & vbCrLf &
                                                  Str(Me.PesoBruto) & "," & Str(Me.Desconto) & ", " & vbCrLf &
                                                  "" & Str(Me.PesoLiquido) & ",'" & CDate(Me.Movimento).ToString("yyyy-MM-dd") & "', '" & IIf(Me.Sinistro, "S", "N") & "', " & Str(Me.TarifaFrete) & ")" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotasXDestinos " & vbCrLf &
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf &
                      "	  and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & Me.NF.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf &
                      "   and Serie_Id        ='" & Me.NF.Serie & "';"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
