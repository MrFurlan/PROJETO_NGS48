Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClientexDependente
    Inherits List(Of ClientexDependente)

#Region "Construtor"
    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = " SELECT Cliente_Id, EndCliente_Id, Nome_Id, TipoDeDependente, RG, CPF, DataNascimento, Profissao, isnull(CustoAno,0) as CustoAno" & vbCrLf & _
                  "   FROM ClienteXDependente" & vbCrLf & _
                  "  Where Cliente_ID    ='" & pCliente.Codigo & "'" & vbCrLf & _
                  "    AND EndCliente_Id = " & pCliente.CodigoEndereco

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ClienteXDependente")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim CD As New ClientexDependente(pCliente)

                CD.Nome = row("Nome_Id")
                CD.TipoDeDependente = row("TipoDeDependente")
                CD.RG = row("RG")
                CD.CPF = row("CPF")
                CD.DataNascimento = row("DataNascimento")
                CD.Profissao = row("Profissao")
                CD.CustoAno = row("CustoAno")
                Me.Add(CD)
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
        For Each CxC As ClientexDependente In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClientexDependente
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
    Private _Nome As String
    Private _TipoDeDependente As String
    Private _RG As String
    Private _CPF As String
    Private _DataNascimento As DateTime
    Private _Profissao As String
    Private _CustoAno As Decimal

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

    Public Property Nome() As String
        Get
            Return _Nome
        End Get
        Set(ByVal value As String)
            _Nome = value
        End Set
    End Property

    Public Property TipoDeDependente() As String
        Get
            Return _TipoDeDependente
        End Get
        Set(ByVal value As String)
            _TipoDeDependente = value
        End Set
    End Property

    Public Property RG() As String
        Get
            Return _RG
        End Get
        Set(ByVal value As String)
            _RG = value
        End Set
    End Property

    Public Property CPF() As String
        Get
            Return _CPF
        End Get
        Set(ByVal value As String)
            _CPF = value
        End Set
    End Property

    Public Property DataNascimento() As DateTime
        Get
            Return _DataNascimento
        End Get
        Set(ByVal value As DateTime)
            _DataNascimento = value
        End Set
    End Property

    Public Property Profissao() As String
        Get
            Return _Profissao
        End Get
        Set(ByVal value As String)
            _Profissao = value
        End Set
    End Property

    Public Property CustoAno() As Decimal
        Get
            Return _CustoAno
        End Get
        Set(ByVal value As Decimal)
            _CustoAno = value
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

                sql = "INSERT INTO ClienteXDependente(Cliente_Id, EndCliente_Id, Nome_Id, TipoDeDependente, RG, CPF, DataNascimento, Profissao, CustoAno) " & vbCrLf & _
                      "VALUES ('" & _Cliente.Codigo & "', " & _Cliente.CodigoEndereco & ",'" & _Nome & "','" & _TipoDeDependente & "','" & _RG & "','" & _CPF & "','" & _DataNascimento.ToString("yyyy-MM-dd") & "','" & _Profissao & "'," & Str(_CustoAno) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE ClienteXDependente SET" & _
                      "    TipoDeDependente ='" & _TipoDeDependente & "'" & vbCrLf & _
                      "   ,RG               ='" & _RG & "'" & vbCrLf & _
                      "   ,CPF              ='" & _CPF & "'" & vbCrLf & _
                      "   ,DataNascimento   ='" & _DataNascimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Profissao        ='" & _Profissao & "'" & vbCrLf & _
                      "   ,CustoAno         = " & Str(_CustoAno) & vbCrLf & _
                      " WHERE Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Nome_Id       ='" & _Nome & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE ClienteXDependente " & _
                      " WHERE Cliente_Id    ='" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   AND EndCliente_Id = " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Nome_Id       ='" & _Nome & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class