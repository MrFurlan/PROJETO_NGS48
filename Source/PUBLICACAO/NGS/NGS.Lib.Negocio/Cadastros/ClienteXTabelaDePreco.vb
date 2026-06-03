Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXTabelaDePreco
    Inherits List(Of ClienteXTabelaDePreco)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT CT.Cliente_Id, CT.Tabela_Id, T.Descricao " & vbCrLf &
                  "  FROM ClientesXTabelasDePrecos CT " & vbCrLf &
                  " INNER JOIN TabelaDePrecos T " & vbCrLf &
                  "    ON T.Codigo_Id = CT.Tabela_Id " & vbCrLf &
                  " WHERE CT.Cliente_Id = '" & pCliente.Codigo & "'" & vbCrLf &
                  "   AND CT.Endereco_Id = " & pCliente.CodigoEndereco & vbCrLf &
                  " ORDER BY CT.Tabela_Id"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ClientesXTabelasDePrecos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim tp As New ClienteXTabelaDePreco(pCliente)
                tp.CodigoTabela = row("Tabela_Id")
                tp.DescricaoTipo = row("Descricao")
                Me.Add(tp)
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
        For Each CxT As ClienteXTabelaDePreco In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxT.IUD = _Cliente.IUD
            If CxT.IUD <> "" Then
                CxT.SalvarSql(Sqls)
            End If
        Next
    End Sub


#End Region

End Class

<Serializable()> _
Public Class ClienteXTabelaDePreco
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
    Private _CodigoTabela As Integer
    Private _TabelaDePreco As TabelaDePreco
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

    Public Property CodigoTabela() As Integer
        Get
            Return _CodigoTabela
        End Get
        Set(ByVal value As Integer)
            _CodigoTabela = value
            _TabelaDePreco = Nothing
        End Set
    End Property

    Public Property TabelaDePreco() As TabelaDePreco
        Get
            If Me.CodigoTabela > 0 And _TabelaDePreco Is Nothing Then _TabelaDePreco = New TabelaDePreco(_CodigoTabela)
            Return _TabelaDePreco
        End Get
        Set(ByVal value As TabelaDePreco)
            _TabelaDePreco = value
        End Set
    End Property

    Public Property DescricaoTipo() As String
        Get
            Return _DescricaoTipo
        End Get
        Set(ByVal value As String)
            _DescricaoTipo = value
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
                sql = "INSERT INTO ClientesXTabelasDePrecos (Cliente_Id, Endereco_Id, Tabela_Id) " & vbCrLf & _
                      "VALUES ('" & _Cliente.Codigo & "', " & _Cliente.CodigoEndereco & "," & _CodigoTabela & ")"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE ClientesXTabelasDePrecos " & _
                      " WHERE Cliente_Id       ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND Endereco_Id      = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Tabela_Id        = " & _CodigoTabela

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class