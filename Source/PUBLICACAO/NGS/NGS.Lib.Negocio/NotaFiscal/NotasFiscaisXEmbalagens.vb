Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class NotasFiscaisXEmbalagens

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pNota As NotaFiscal)
        Dim sql As String
        Dim Banco As New AcessaBanco

        sql = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Quantidade, Especie, Marca, Numero, PesoBruto, PesoLiquido" & vbCrLf & _
              "  FROM NotasXEmbalagens" & vbCrLf & _
              " WHERE Empresa_Id      ='" & pNota.Empresa.Codigo & "'" & vbCrLf & _
              "   AND EndEmpresa_Id   = " & pNota.Empresa.CodigoEndereco & vbCrLf & _
              "   AND Cliente_Id      ='" & pNota.Cliente.Codigo & "'" & vbCrLf & _
              "   AND EndCliente_Id   = " & pNota.Cliente.CodigoEndereco & vbCrLf & _
              "   AND EntradaSaida_Id ='" & pNota.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
              "   AND Serie_Id        ='" & pNota.Serie & "'" & vbCrLf & _
              "   AND Nota_Id         = " & pNota.Codigo & _
              "   AND Nota_Id > 0"

        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "Embalagem")

        For Each row As DataRow In ds.Tables(0).Rows
            Empresa = row("Empresa_Id")
            EndEmpresa = row("EndEmpresa_Id")
            Cliente = row("Cliente_Id")
            EndCliente = row("EndCliente_Id")
            EntradaSaida = row("EntradaSaida_Id")
            Serie = row("Serie_Id")
            Nota = row("Nota_Id")
            Quantidade = row("Quantidade")
            Especie = row("Especie")
            Marca = row("Marca")
            Numero = row("Numero")
            PesoBruto = row("PesoBruto")
            PesoLiquido = row("PesoLiquido")
        Next


    End Sub
#End Region

#Region "Fields"
    Private _Empresa As String
    Private _EndEmpresa As Integer
    Private _Cliente As String
    Private _EndCliente As Integer
    Private _EntradaSaida As String
    Private _Serie As String
    Private _Nota As Integer
    Private _Quantidade As Decimal
    Private _Especie As String
    Private _Marca As String
    Private _Numero As String
    Private _PesoBruto As Decimal
    Private _PesoLiquido As Decimal
#End Region

#Region "Property"

    Public Property Empresa() As String
        Get
            Return _Empresa
        End Get
        Set(ByVal value As String)
            _Empresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property Cliente() As String
        Get
            Return _Cliente
        End Get
        Set(ByVal value As String)
            _Cliente = value
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property Nota() As Integer
        Get
            Return _Nota
        End Get
        Set(ByVal value As Integer)
            _Nota = value
        End Set
    End Property

    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Especie() As String
        Get
            Return _Especie
        End Get
        Set(ByVal value As String)
            _Especie = value
        End Set
    End Property

    Public Property Marca() As String
        Get
            Return _Marca
        End Get
        Set(ByVal value As String)
            _Marca = value
        End Set
    End Property

    Public Property Numero() As String
        Get
            Return _Numero
        End Get
        Set(ByVal value As String)
            _Numero = value
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

    Public Property PesoLiquido() As Decimal
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Decimal)
            _PesoLiquido = value
        End Set
    End Property
#End Region

End Class