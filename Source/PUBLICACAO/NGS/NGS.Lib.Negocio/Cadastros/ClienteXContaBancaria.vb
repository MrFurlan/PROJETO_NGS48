Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXContaBancaria
    Inherits List(Of ClienteXContaBancaria)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pCliente As Cliente, Optional ByVal SomenteContasAtivas As Boolean = False)
        _Cliente = pCliente

        Dim Banco As New AcessaBanco()

        Try
            Dim sql As String
            sql = " SELECT CXCB.Banco_Id, Bancos.Descricao As NomeBanco, CXCB.Agencia_Id, CXCB.DigitoAgencia_Id, CXCB.ContaCorrente_Id, CXCB.DigitoConta_Id,Isnull(CXCB.TipoConta,'C') as TipoConta, isnull(CXCB.Observacoes,'') as Observacoes, " & vbCrLf & _
                  "       Isnull(Agencias.Praca,'') as Praca, Isnull(Agencias.Cidade,'') as Cidade, Isnull(Agencias.Estado_Id,'') as Estado, ISNULL(CXCB.Ativo,1) Ativo " & vbCrLf & _
                  "   FROM ClientesXContasBancarias CXCB" & vbCrLf & _
                  "  LEFT JOIN Agencias " & vbCrLf & _
                  "     ON CXCB.Banco_id         = Agencias.Banco_Id " & vbCrLf & _
                  "    AND CXCB.Agencia_id       = Agencias.Agencia_Id " & vbCrLf & _
                  "    AND CXCB.DigitoAgencia_id = Agencias.Digito_Id " & vbCrLf & _
                  "  INNER JOIN Bancos " & vbCrLf & _
                  "     ON CXCB.Banco_id         = Bancos.Banco_Id " & vbCrLf & _
                  "  WHERE CXCB.Cliente_Id  ='" & pCliente.Codigo & "'" & vbCrLf & _
                  "    AND CXCB.Endereco_Id = " & pCliente.CodigoEndereco & vbCrLf
            If SomenteContasAtivas Then
                sql &= "    AND CXCB.Ativo = 1 " & vbCrLf
            End If
            sql &= "  ORDER BY CXCB.Banco_Id, CXCB.Agencia_Id"

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "ClientesXContasBancarias")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim objConta As New ClienteXContaBancaria(pCliente)

                objConta.CodigoBanco = row("Banco_Id")
                objConta.NomeBanco = row("NomeBanco")
                objConta.CodigoAgencia = row("Agencia_Id")
                objConta.DigitoAgencia = row("DigitoAgencia_Id")
                objConta.ContaCorrente = row("ContaCorrente_Id")
                objConta.DigitoConta = row("DigitoConta_Id")
                objConta.TipoConta = row("TipoConta")
                objConta.Observacoes = row("Observacoes")
                objConta.Praca = row("Praca")
                objConta.Cidade = row("Cidade")
                objConta.Estado = row("Estado")
                objConta.Ativo = row("Ativo")

                Me.Add(objConta)
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
        For Each CxC As ClienteXContaBancaria In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXContaBancaria
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
    Private _CodigoBanco As Integer
    Private _NomeBanco As String
    Private _Banco As Banco
    Private _CodigoAgencia As Integer
    Private _DigitoAgencia As String
    Private _Agencia As Agencia
    Private _ContaCorrente As String
    Private _DigitoConta As String
    Private _TipoConta As String
    Private _Observacoes As String
    Private _Praca As String
    Private _Cidade As String
    Private _Estado As String
    Private _Ativo As Boolean


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

    Public Property CodigoAgencia() As Integer
        Get
            Return _CodigoAgencia
        End Get
        Set(ByVal value As Integer)
            _CodigoAgencia = value
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

    Public Property Agencia() As Agencia
        Get
            If Me.CodigoBanco > 0 And _
               Me.CodigoAgencia > 0 And _
               _Agencia Is Nothing Then _Agencia = New Agencia(Me.CodigoBanco, Me.CodigoAgencia, Me.DigitoAgencia)
            Return _Agencia
        End Get
        Set(ByVal value As Agencia)
            _Agencia = value
        End Set
    End Property

    Public Property ContaCorrente() As String
        Get
            Return _ContaCorrente
        End Get
        Set(ByVal value As String)
            _ContaCorrente = value
        End Set
    End Property

    Public Property DigitoConta() As String
        Get
            Return _DigitoConta
        End Get
        Set(ByVal value As String)
            _DigitoConta = value
        End Set
    End Property

    Public Property TipoConta() As String
        Get
            Return _TipoConta
        End Get
        Set(ByVal value As String)
            _TipoConta = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
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

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
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
                sql = "INSERT INTO ClientesXContasBancarias (Cliente_Id, Endereco_Id, Banco_Id, Agencia_Id, DigitoAgencia_Id, ContaCorrente_Id, DigitoConta_Id, TipoConta, Observacoes, Ativo) " & vbCrLf & _
                      "VALUES ('" & Me.Cliente.Codigo & "', " & Me.Cliente.CodigoEndereco & "," & Me.CodigoBanco & ", " & Me.CodigoAgencia & ", '" & Me.DigitoAgencia & "'," & _
                      "'" & Me.ContaCorrente & "', '" & Me.DigitoConta & "', '" & Me.TipoConta & "', '" & Me.Observacoes & "', '" & Me.Ativo & "')"
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "UPDATE ClientesXContasBancarias " & _
                      "   SET Observacoes      = '" & Me.Observacoes & "' , " & vbCrLf & _
                      "       TipoConta        ='" & Me.TipoConta & "', " & vbCrLf & _
                      "       Ativo            ='" & Me.Ativo & "' " & vbCrLf & _
                      " WHERE Cliente_Id       = '" & Me.Cliente.Codigo & "'" & vbCrLf & _
                      "   AND Endereco_Id      =  " & Me.Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Banco_Id         =  " & Me.CodigoBanco & vbCrLf & _
                      "   AND Agencia_Id       =  " & Me.CodigoAgencia & vbCrLf & _
                      "   AND DigitoAgencia_Id = '" & Me.DigitoAgencia & "'" & vbCrLf & _
                      "   AND ContaCorrente_Id = '" & Me.ContaCorrente & "'" & vbCrLf & _
                      "   AND DigitoConta_Id   = '" & Me.DigitoConta & "'" & vbCrLf

                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                sql = "DELETE ClientesXContasBancarias " & _
                      " WHERE Cliente_Id       = '" & Me.Cliente.Codigo & "'" & vbCrLf & _
                      "   AND Endereco_Id      =  " & Me.Cliente.CodigoEndereco & vbCrLf & _
                      "   AND Banco_Id         =  " & Me.CodigoBanco & vbCrLf & _
                      "   AND Agencia_Id       =  " & Me.CodigoAgencia & vbCrLf & _
                      "   AND DigitoAgencia_Id = '" & Me.DigitoAgencia & "'" & vbCrLf & _
                      "   AND ContaCorrente_Id = '" & Me.ContaCorrente & "'" & vbCrLf & _
                      "   AND DigitoConta_Id   = '" & Me.DigitoConta & "'"
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Dim objAgencia As New Agencia()

        If _Agencia Is Nothing OrElse _Agencia.Agencia = 0 Then
            objAgencia.IUD = "I"
        Else
            objAgencia.IUD = "U"
        End If

        objAgencia.CodigoBanco = _CodigoBanco
        objAgencia.Agencia = _CodigoAgencia
        objAgencia.DigitoAgencia = _DigitoAgencia
        objAgencia.Praca = _Praca
        objAgencia.Cidade = _Cidade
        objAgencia.Estado = _Estado

        objAgencia.SalvarSql(Sqls)
    End Sub
#End Region

End Class