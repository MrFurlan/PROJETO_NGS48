Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTransferenciaFinanceira
    Inherits List(Of TransferenciaFinanceira)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresaDebito As Cliente, ByVal pEmpresaCredito As Cliente)
        Dim str As String
        str = "SELECT EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, ContaContabil, ClienteContabil, EndClienteContabil, DebitoCredito, EmiteAviso" & vbCrLf & _
              "  FROM TransferenciasFinanceiras" & vbCrLf & _
              " Where EmpresaDebito   ='" & pEmpresaDebito.Codigo & "'" & vbCrLf & _
              "   and EnderecoDebito  = " & pEmpresaDebito.CodigoEndereco & vbCrLf & _
              "   and EmpresaCredito  ='" & pEmpresaCredito.Codigo & "'" & vbCrLf & _
              "   and EnderecoCredito = " & pEmpresaCredito.CodigoEndereco
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(str, "TransFinanceira")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Trans As New TransferenciaFinanceira
            Trans.CodigoEmpresaDebito = row("EmpresaDebito")
            Trans.EndEmpresaDebito = row("EnderecoDebito")
            Trans.CodigoEmpresaCredito = row("EmpresaCredito")
            Trans.EndEmpresaCredito = row("EnderecoCredito")
            Trans.CodigoEmpresaContabil = row("EmpresaContabil")
            Trans.EndEmpresaContabil = row("EnderecoContabil")
            Trans.ContaContabil = row("ContaContabil")
            Trans.CodigoClienteContabil = row("ClienteContabil")
            Trans.EndClienteContabil = row("EndClienteContabil")
            Trans.DebitoCredito = row("DebitoCredito")
            Trans.EmiteAviso = row("EmiteAviso") = "S"
            Me.Add(Trans)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TransferenciaFinanceira
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoEmpresaDebito As String = ""
    Private _EndEmpresaDebito As Integer
    Private _EmpresaDebito As Cliente

    Private _CodigoEmpresaCredito As String = ""
    Private _EndEmpresaCredito As Integer
    Private _EmpresaCredito As Cliente

    Private _CodigoEmpresaContabil As String = ""
    Private _EndEmpresaContabil As Integer
    Private _EmpresaContabil As Cliente

    Private _ContaContabil As String = ""

    Private _CodigoClienteContabil As String = ""
    Private _EndClienteContabil As Integer
    Private _ClienteContabil As Cliente

    Private _DebitoCredito As String = ""
    Private _EmiteAviso As Boolean
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


    Public Property CodigoEmpresaDebito() As String
        Get
            Return _CodigoEmpresaDebito
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaDebito = value
            _EmpresaDebito = Nothing
        End Set
    End Property

    Public Property EndEmpresaDebito() As Integer
        Get
            Return _EndEmpresaDebito
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaDebito = value
            _EmpresaDebito = Nothing
        End Set
    End Property

    Public ReadOnly Property EmpresaDebito() As Cliente
        Get
            If _CodigoEmpresaDebito.Length > 0 Then _EmpresaDebito = New Cliente(_CodigoEmpresaDebito, _EndEmpresaDebito)
            Return _EmpresaDebito
        End Get
    End Property


    Public Property CodigoEmpresaCredito() As String
        Get
            Return _CodigoEmpresaCredito
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaCredito = value
            _EmpresaCredito = Nothing
        End Set
    End Property

    Public Property EndEmpresaCredito() As Integer
        Get
            Return _EndEmpresaCredito
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaCredito = value
            _EmpresaCredito = Nothing
        End Set
    End Property

    Public ReadOnly Property EmpresaCredito() As Cliente
        Get
            If _EmpresaCredito Is Nothing And _CodigoEmpresaCredito.Length > 0 Then _EmpresaCredito = New Cliente(_CodigoEmpresaCredito, _EndEmpresaCredito)
            Return _EmpresaCredito
        End Get
    End Property


    Public Property CodigoEmpresaContabil() As String
        Get
            Return _CodigoEmpresaContabil
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaContabil = value
            _EmpresaContabil = Nothing
        End Set
    End Property

    Public Property EndEmpresaContabil() As Integer
        Get
            Return _EndEmpresaContabil
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaContabil = value
            _EmpresaContabil = Nothing
        End Set
    End Property

    Public ReadOnly Property EmpresaContabil() As Cliente
        Get
            If _EmpresaContabil Is Nothing And _CodigoEmpresaContabil.Length > 0 Then _EmpresaContabil = New Cliente(_CodigoEmpresaContabil, _EndEmpresaContabil)
            Return _EmpresaContabil
        End Get
    End Property


    Public Property ContaContabil() As String
        Get
            Return _ContaContabil
        End Get
        Set(ByVal value As String)
            _ContaContabil = value
        End Set
    End Property

    Public Property CodigoClienteContabil() As String
        Get
            Return _CodigoClienteContabil
        End Get
        Set(ByVal value As String)
            _CodigoClienteContabil = value
            _ClienteContabil = Nothing
        End Set
    End Property

    Public Property EndClienteContabil() As Integer
        Get
            Return _EndClienteContabil
        End Get
        Set(ByVal value As Integer)
            _EndClienteContabil = value
            _ClienteContabil = Nothing
        End Set
    End Property

    Public ReadOnly Property ClienteContabil() As Cliente
        Get
            If _ClienteContabil Is Nothing And _CodigoClienteContabil.Length > 0 Then _ClienteContabil = New Cliente(_CodigoClienteContabil, _EndClienteContabil)
            Return _ClienteContabil
        End Get
    End Property


    Public Property DebitoCredito() As String
        Get
            Return _DebitoCredito
        End Get
        Set(ByVal value As String)
            _DebitoCredito = value
        End Set
    End Property

    Public Property EmiteAviso() As Boolean
        Get
            Return _EmiteAviso
        End Get
        Set(ByVal value As Boolean)
            _EmiteAviso = value
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
                Sql = "Insert Into TransferenciasFinanceiras(EmpresaDebito, EnderecoDebito, EmpresaCredito, EnderecoCredito, EmpresaContabil, EnderecoContabil, ContaContabil, ClienteContabil, EndClienteContabil, DebitoCredito, EmiteAviso)" & vbCrLf & _
                      " values ('" & _CodigoEmpresaDebito & "'," & _EndEmpresaDebito & ",'" & _CodigoEmpresaCredito & "'," & _EndEmpresaCredito & ",'" & _CodigoEmpresaContabil & "'," & _EndEmpresaContabil & ",'" & _ContaContabil & "','" & _CodigoClienteContabil & "'," & _EndClienteContabil & ",'" & _DebitoCredito & "'," & CByte(_EmiteAviso) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE TransferenciasFinanceiras SET" & vbCrLf & _
                      "    DebitoCredito ='" & _DebitoCredito & "'" & vbCrLf & _
                      "   ,EmiteAviso    = " & CByte(_EmiteAviso) & vbCrLf & _
                      "  WHERE EmpresaDebito      ='" & _CodigoEmpresaDebito & "'" & vbCrLf & _
                      "    And EnderecoDebito     = " & _EndEmpresaDebito & vbCrLf & _
                      "    And EmpresaCredito     ='" & _CodigoEmpresaCredito & "'" & vbCrLf & _
                      "    And EnderecoCredito    = " & _EndEmpresaCredito & vbCrLf & _
                      "    And EmpresaContabil    ='" & _CodigoEmpresaContabil & "'" & vbCrLf & _
                      "    And EnderecoContabil   = " & _EndEmpresaContabil & vbCrLf & _
                      "    And ContaContabil      ='" & _ContaContabil & "'" & vbCrLf & _
                      "    And ClienteContabil    ='" & _CodigoClienteContabil & "'" & vbCrLf & _
                      "    And EndClienteContabil = " & _EndClienteContabil
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE TransferenciasFinanceiras" & vbCrLf & _
                      "  WHERE EmpresaDebito      ='" & _CodigoEmpresaDebito & "'" & vbCrLf & _
                      "    And EnderecoDebito     = " & _EndEmpresaDebito & vbCrLf & _
                      "    And EmpresaCredito     ='" & _CodigoEmpresaCredito & "'" & vbCrLf & _
                      "    And EnderecoCredito    = " & _EndEmpresaCredito & vbCrLf & _
                      "    And EmpresaContabil    ='" & _CodigoEmpresaContabil & "'" & vbCrLf & _
                      "    And EnderecoContabil   = " & _EndEmpresaContabil & vbCrLf & _
                      "    And ContaContabil      ='" & _ContaContabil & "'" & vbCrLf & _
                      "    And ClienteContabil    ='" & _CodigoClienteContabil & "'" & vbCrLf & _
                      "    And EndClienteContabil = " & _EndClienteContabil
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class