Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClientexTipo
    Inherits List(Of ClientexTipo)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT TC.Tipo_Id " & vbCrLf & _
                  "  FROM ClientesXTipos CT " & vbCrLf & _
                  " INNER JOIN TiposDeClientes TC " & vbCrLf & _
                  "    ON TC.Tipo_Id = CT.Tipo_Id " & vbCrLf & _
                  " WHERE CT.Cliente_Id = '" & pCliente.Codigo & "'" & vbCrLf & _
                  "   AND CT.Endereco_Id = " & pCliente.CodigoEndereco & vbCrLf & _
                  " ORDER BY TC.Tipo_Id"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ClientesXTipo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim tip As New ClientexTipo(pCliente)
                tip.CodigoTipo = row("Tipo_Id")
                Me.Add(tip)
            Next
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try

    End Sub
#End Region

#Region "Fields"
    Public _Cliente As Cliente
    Public Erro As Exception
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxC As ClientexTipo In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub


#End Region

End Class

<Serializable()> _
Public Class ClientexTipo
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Cliente As Cliente)
        Me.Cliente = Cliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _CodigoTipo As Integer
    Private _Tipo As TipoDeCliente
    Private _DescricaoTipo As String

    Public Erro As Exception
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

    Public Property CodigoTipo() As Integer
        Get
            Return _CodigoTipo
        End Get
        Set(ByVal value As Integer)
            _CodigoTipo = value
            _Tipo = Nothing
            _DescricaoTipo = ""
        End Set
    End Property

    Public Property TipoDeCliente() As TipoDeCliente
        Get
            If Me.CodigoTipo > 0 And _Tipo Is Nothing Then _Tipo = New TipoDeCliente(_CodigoTipo)
            Return _Tipo
        End Get
        Set(ByVal value As TipoDeCliente)
            _Tipo = value
        End Set
    End Property

    Public ReadOnly Property DescricaoTipo() As String
        Get
            If _DescricaoTipo = "" And _CodigoTipo > 0 Then _DescricaoTipo = TipoDeCliente.Descricao
            Return _DescricaoTipo
        End Get
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
                sql = "INSERT INTO ClientesxTipos (Cliente_Id, Endereco_Id, Tipo_Id) " & vbCrLf & _
                      "VALUES ('" & _Cliente.Codigo & "', " & _Cliente.CodigoEndereco & "," & _CodigoTipo & ")"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE ClientesxTipos " & _
                      " WHERE Cliente_Id       ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND Endereco_Id      = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Tipo_Id          = " & _CodigoTipo

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class