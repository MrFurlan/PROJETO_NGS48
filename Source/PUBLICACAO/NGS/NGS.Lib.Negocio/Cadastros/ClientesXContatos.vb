Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXContato
    Inherits List(Of ClienteXContato)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pcliente As Cliente)
        _Cliente = pcliente
        Dim sql As String
        sql = "Select cc.Cliente_Id, cc.Endereco_Id, cc.NomeContato, " & vbCrLf & _
              "       isnull(cc.Funcao,'') as Funcao, isnull(cc.Telefone,'') as Telefone, " & vbCrLf & _
              "       isnull(cc.Email,'') as Email, isnull(cc.Observacoes,'') as Observacoes, " & vbCrLf & _
              "       ISNULL(cc.Banco, 0) AS Banco, ISNULL(Bancos.Descricao, '') As NomeBanco, ISNULL(cc.Agencia, '') AS Agencia, " & vbCrLf & _
              "       ISNULL(cc.DigitoAgencia,'') DigitoAgencia, ISNULL(cc.Conta,'') AS Conta , ISNULL(cc.DigitoConta, '') AS DigitoConta " & vbCrLf & _
              "  from ClientexContato cc" & vbCrLf & _
              "  LEFT JOIN Bancos " & vbCrLf & _
              "     ON cc.Banco = Bancos.Banco_Id " & vbCrLf & _
              " Where Cliente_Id  ='" & pcliente.Codigo & "'" & vbCrLf & _
              "   and Endereco_Id = " & pcliente.CodigoEndereco

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "Contato")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Cont As New ClienteXContato(pcliente)
            Cont.NomeContato = row("NomeContato")
            Cont.Funcao = row("Funcao")
            Cont.Telefone = row("Telefone")
            Cont.email = row("Email")
            Cont.Observacao = row("Observacoes")
            Cont.CodigoBanco = row("Banco")
            Cont.NomeBanco = row("NomeBanco")
            Cont.CodigoAgencia = row("Agencia")
            Cont.DigitoAgencia = row("DigitoAgencia")
            Cont.CodigoContaCorrente = row("Conta")
            Cont.DigitoConta = row("DigitoConta")
            Me.Add(Cont)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Cliente As Cliente
#End Region

#Region "Property"
    Public Property Cliente() As Cliente
        Get
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each CxC As ClienteXContato In Me
            If _Cliente.IUD = "D" Or _Cliente.IUD = "I" Then CxC.IUD = _Cliente.IUD
            If CxC.IUD <> "" Then
                CxC.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXContato
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub

    Public Sub New(ByVal cliente_Id As String, ByVal endCliente_Id As Integer, ByVal nome As String)
        Dim sql As String
        sql = "Select cc.Cliente_Id, cc.Endereco_Id, cc.NomeContato, " & vbCrLf & _
              "       isnull(cc.Funcao,'') as Funcao, isnull(cc.Telefone,'') as Telefone, " & vbCrLf & _
              "       isnull(cc.Email,'') as Email, isnull(cc.Observacoes,'') as Observacoes, " & vbCrLf & _
              "       ISNULL(cc.Banco, 0) AS Banco, ISNULL(Bancos.Descricao, '') As NomeBanco, ISNULL(cc.Agencia, '') AS Agencia, " & vbCrLf & _
              "       ISNULL(cc.DigitoAgencia,'') DigitoAgencia, ISNULL(cc.Conta,'') AS Conta , ISNULL(cc.DigitoConta, '') AS DigitoConta " & vbCrLf & _
              "  from ClientexContato cc" & vbCrLf & _
              "  LEFT JOIN Bancos " & vbCrLf & _
              "     ON cc.Banco = Bancos.Banco_Id " & vbCrLf & _
              " Where Cliente_Id  ='" & cliente_Id & "'" & vbCrLf & _
              "   and Endereco_Id = " & endCliente_Id & vbCrLf & _
              "   and nome = '" & nome & "'"

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "Contato")
        With ds.Tables(0).Rows(0)
            Me.NomeContato = ("NomeContato")
            Me.Funcao = ("Funcao")
            Me.Telefone = ("Telefone")
            Me.email = ("Email")
            Me.Observacao = ("Observacoes")
            Me.CodigoBanco = ("Banco")
            Me.NomeBanco = ("NomeBanco")
            Me.CodigoAgencia = ("Agencia")
            Me.DigitoAgencia = ("DigitoAgencia")
            Me.CodigoContaCorrente = ("Conta")
            Me.DigitoConta = ("DigitoConta")
        End With
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Cliente As Cliente
    Private _NomeContato As String
    Private _Funcao As String
    Private _Telefone As String
    Private _email As String
    Private _Observacao As String
    Private _CodigoBanco As Integer
    Private _NomeBanco As String
    Private _CodigoAgencia As Integer
    Private _DigitoAgencia As String
    Private _CodigoContaCorrente As String
    Private _DigitoConta As String
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

    Public Property NomeContato() As String
        Get
            Return _NomeContato
        End Get
        Set(ByVal value As String)
            _NomeContato = value
        End Set
    End Property

    Public Property Funcao() As String
        Get
            Return _Funcao
        End Get
        Set(ByVal value As String)
            _Funcao = value
        End Set
    End Property

    Public Property Telefone() As String
        Get
            Return _Telefone
        End Get
        Set(ByVal value As String)
            _Telefone = value
        End Set
    End Property

    Public Property email() As String
        Get
            Return _email
        End Get
        Set(ByVal value As String)
            _email = value
        End Set
    End Property

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
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

    Public Property CodigoContaCorrente() As String
        Get
            Return _CodigoContaCorrente
        End Get
        Set(ByVal value As String)
            _CodigoContaCorrente = value
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
                sql = "Insert into ClienteXContato(Cliente_Id, Endereco_Id, NomeContato, Funcao, Telefone, Email, Observacoes, Banco, Agencia, DigitoAgencia, Conta, DigitoConta)" & vbCrLf & _
                      " values('" & _Cliente.Codigo & "'," & _Cliente.CodigoEndereco & ",'" & _NomeContato & "','" & _Funcao & "','" & _Telefone & "','" & _email & "','" & _Observacao & "'," & _CodigoBanco & "," & _CodigoAgencia & ",'" & _DigitoAgencia & "','" & _CodigoContaCorrente & "','" & _DigitoConta & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update ClienteXContato set " & vbCrLf & _
                      "    Funcao         = '" & _Funcao & "'" & vbCrLf & _
                      "   ,Telefone       = '" & _Telefone & "'" & vbCrLf & _
                      "   ,Email          = '" & _email & "'" & vbCrLf & _
                      "   ,Observacoes    = '" & _Observacao & "'" & vbCrLf & _
                      "   ,Banco          =  " & _CodigoBanco & vbCrLf & _
                      "   ,Agencia        =  " & _CodigoAgencia & "" & vbCrLf & _
                      "   ,DigitoAgencia  = '" & _DigitoAgencia & "'" & vbCrLf & _
                      "   ,Conta          = '" & _CodigoContaCorrente & "'" & vbCrLf & _
                      "   ,DigitoConta    = '" & _DigitoConta & "'" & vbCrLf & _
                      " Where Cliente_Id  = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and Endereco_Id =  " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and NomeContato = '" & _NomeContato & "'"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXContato " & vbCrLf & _
                      " Where Cliente_Id  = '" & _Cliente.Codigo & "'" & vbCrLf & _
                      "   and Endereco_Id =  " & _Cliente.CodigoEndereco & vbCrLf & _
                      "   and NomeContato = '" & _NomeContato & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class