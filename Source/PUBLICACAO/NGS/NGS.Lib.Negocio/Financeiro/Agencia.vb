Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAgencia
    Inherits List(Of Agencia)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim sql As String = "Select Ag.Banco_Id, Bancos.Descricao As NomeBanco, Ag.Agencia_Id, Ag.Digito_Id, Ag.Praca, Ag.Cidade, Ag.Estado_Id " & vbCrLf & _
                "  From Agencias AS Ag " & vbCrLf & _
                "  INNER JOIN Bancos " & vbCrLf & _
                "     ON Ag.Banco_id = Bancos.Banco_Id " & vbCrLf & _
                "  ORDER BY Ag.Banco_Id, Ag.Agencia_Id"

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Agencias")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ag As New Agencia
                Ag.CodigoBanco = row("Banco_Id")
                Ag.NomeBanco = row("NomeBanco")
                Ag.Agencia = row("Agencia_Id")
                Ag.DigitoAgencia = row("Digito_Id")
                Ag.Praca = row("Praca")
                Ag.Cidade = row("Cidade")
                Ag.Estado = row("Estado_Id")
                Me.Add(Ag)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Agencia
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pBanco As Integer, ByVal pAgencia As Integer, ByVal pDigito As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Ag.Banco_Id, Bancos.Descricao As NomeBanco, Ag.Agencia_Id, Ag.Digito_Id, Ag.Praca, Ag.Cidade, Ag.Estado_Id " & vbCrLf & _
                            "  From Agencias AS Ag " & vbCrLf & _
                            "  INNER JOIN Bancos " & vbCrLf & _
                            "     ON Ag.Banco_id = Bancos.Banco_Id " & vbCrLf & _
                            "  WHERE Ag.Banco_Id   = " & pBanco & vbCrLf & _
                            "    AND Ag.Agencia_Id = " & pAgencia & vbCrLf & _
                            "    AND Ag.Digito_id  = '" & pDigito & "'"

        ds = Banco.ConsultaDataSet(sql, "Agencias")
        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoBanco = ds.Tables(0).Rows(0)("Banco_Id")
            _NomeBanco = ds.Tables(0).Rows(0)("NomeBanco")
            _Agencia = ds.Tables(0).Rows(0)("Agencia_Id")
            _DigitoAgencia = ds.Tables(0).Rows(0)("Digito_Id")
            _Praca = ds.Tables(0).Rows(0)("Praca")
            _Cidade = ds.Tables(0).Rows(0)("Cidade")
            _Estado = ds.Tables(0).Rows(0)("Estado_Id")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoBanco As Integer
    Private _NomeBanco As String
    Private _Banco As Banco
    Private _Agencia As Integer
    Private _DigitoAgencia As String
    Private _Praca As String
    Private _Cidade As String
    Private _Estado As String
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

    Public Property CodigoBanco() As Integer
        Get
            Return _CodigoBanco
        End Get
        Set(ByVal value As Integer)
            _CodigoBanco = value
        End Set
    End Property

    Public Property NomeBanco() As String
        Get
            Return _NomeBanco
        End Get
        Set(ByVal value As String)
            _NomeBanco = value
        End Set
    End Property

    Public Property Banco() As Banco
        Get
            If Me.CodigoBanco > 0 And _Banco Is Nothing Then _Banco = New Banco(Me.CodigoBanco)
            Return _Banco
        End Get
        Set(ByVal value As Banco)
            _Banco = value
        End Set
    End Property

    Public Property Agencia() As Integer
        Get
            Return _Agencia
        End Get
        Set(ByVal value As Integer)
            _Agencia = value
        End Set
    End Property

    Public Property DigitoAgencia() As String
        Get
            Return _DigitoAgencia
        End Get
        Set(ByVal value As String)
            _DigitoAgencia = value
        End Set
    End Property

    Public Property Praca() As String
        Get
            Return _Praca
        End Get
        Set(ByVal value As String)
            _Praca = value
        End Set
    End Property

    Public Property Cidade() As String
        Get
            Return _Cidade
        End Get
        Set(ByVal value As String)
            _Cidade = value
        End Set
    End Property

    Public Property Estado() As String
        Get
            Return _Estado
        End Get
        Set(ByVal value As String)
            _Estado = value
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
                Sql = " INSERT INTO Agencias (Banco_Id, Agencia_Id, Digito_Id, Praca, Cidade, Estado_Id) " & vbCrLf & _
                      " VALUES (" & _CodigoBanco & "," & _Agencia & ",'" & _DigitoAgencia & "','" & _Praca & "','" & _Cidade & "','" & _Estado & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Agencias SET" & vbCrLf & _
                      "    Praca      = '" & _Praca & "'" & vbCrLf & _
                      "    ,Cidade    = '" & _Cidade & "'" & vbCrLf & _
                      "    ,Estado_Id = '" & _Estado & "'" & vbCrLf & _
                      "  WHERE Banco_Id   = " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia_Id = " & _Agencia & vbCrLf & _
                      "    AND Digito_Id  = '" & _DigitoAgencia & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Agencias " & vbCrLf & _
                      "  WHERE Banco_Id   = " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia_Id = " & _Agencia & vbCrLf & _
                      "    AND Digito_Id  = '" & _DigitoAgencia & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class