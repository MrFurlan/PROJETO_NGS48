Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAnaliseDeCreditoCliente
    Inherits List(Of AnaliseDeCreditoCliente)

#Region "Construtor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
        If pAnaliseDeCredito.CodigoAnalise > 0 Then
            Dim ds As DataSet
            Dim sql As String = ""
            sql = "SELECT an.Analise_Id, an.Ano_Id, an.DefinicaoAno_Id, " & vbCrLf & _
                  "       an.Cliente_Id, " & vbCrLf & _
                  "       (Select top 1 c.Nome from clientes c where c.cliente_id = an.Cliente_Id) as Nome" & vbCrLf & _
                  "  FROM AnaliseDeCreditoCliente an" & vbCrLf & _
                  " Where an.Analise_Id      = " & pAnaliseDeCredito.CodigoAnalise & vbCrLf & _
                  "   and an.Ano_Id          = " & pAnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                  "   and an.DefinicaoAno_Id = " & pAnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno

            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoCliente")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim AnC As New AnaliseDeCreditoCliente(pAnaliseDeCredito)
                AnC.CodigoCliente = row("Cliente_Id")
                AnC.NomeCliente = row("Nome")
                Me.Add(AnC)
            Next
        End If
    End Sub
#End Region

#Region "Fields"
    Private _AnaliseDeCredito As AnaliseDeCredito
#End Region

#Region "Property"
    Public ReadOnly Property AnaliseDeCredito() As AnaliseDeCredito
        Get
            Return _AnaliseDeCredito
        End Get
    End Property

    Public ReadOnly Property ClientesParaSql() As String
        Get
            Dim cli As String = ""
            For Each row As AnaliseDeCreditoCliente In Me
                cli &= IIf(cli.Length = 0, "", ",") & "'" & row.CodigoCliente & "'"
            Next
            Return IIf(cli <> "", cli, "''")
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each cli As AnaliseDeCreditoCliente In Me
            If AnaliseDeCredito.IUD = "D" Or AnaliseDeCredito.IUD = "I" Then cli.IUD = AnaliseDeCredito.IUD
            If cli.IUD <> "" Then
                cli.SalvarSql(Sqls)
            End If
        Next
    End Sub

    Public Function AdicionarCliente(ByVal pCliente As Cliente) As Boolean
        For Each row As AnaliseDeCreditoCliente In Me
            If row.CodigoCliente = pCliente.Codigo Then
                Return False
            End If
        Next

        Dim CliXAna As New AnaliseDeCreditoCliente(AnaliseDeCredito)
        CliXAna.CodigoCliente = pCliente.Codigo
        CliXAna.NomeCliente = pCliente.Nome
        Me.Add(CliXAna)
        Return True
    End Function
#End Region

End Class

<Serializable()> _
Public Class AnaliseDeCreditoCliente
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _AnaliseDeCredito As AnaliseDeCredito
    Private _CodigoCliente As String = ""
    Private _NomeCliente As String = ""
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

    Public Property AnaliseDeCredito() As AnaliseDeCredito
        Get
            Return _AnaliseDeCredito
        End Get
        Set(ByVal value As AnaliseDeCredito)
            _AnaliseDeCredito = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property NomeCliente() As String
        Get
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = ""
        Select Case Me.IUD
            Case "I"
                strSQL &= " Insert Into AnaliseDeCreditoCliente(Analise_Id, Ano_Id, DefinicaoAno_Id, Cliente_Id)" & vbCrLf & _
                          " Values (" & AnaliseDeCredito.CodigoAnalise & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & ",'" & Me.CodigoCliente & "')"
                Sqls.Add(strSQL)
            Case "D"
                strSQL = " Delete AnaliseDeCreditoCliente" & vbCrLf & _
                         "  Where Analise_Id      = " & AnaliseDeCredito.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & _
                         "    and Cliente_Id      = " & Me.CodigoCliente
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
