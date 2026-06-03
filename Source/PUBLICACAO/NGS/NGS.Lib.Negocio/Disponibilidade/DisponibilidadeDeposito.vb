Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListDisponibilidadeDeposito
    Inherits List(Of DisponibilidadeDeposito)

#Region "Construtor"
    Public Sub New(ByVal pDisponibilidade As Disponibilidade)
        _Disponibilidade = pDisponibilidade
        Dim sql As String = ""
        sql = "Select Disponibilidade_Id, Deposito_Id, EndDeposito_Id, SaldoInicialFiscal, SaldoInicialFisico" & vbCrLf & _
              "  from SaldoInicialDisponibilidadeDeposito" & vbCrLf & _
              " Where Disponibilidade_id = " & pDisponibilidade.CodigoDisponibilidade

        Dim banco As New AcessaBanco
        Dim ds As DataSet
        ds = banco.ConsultaDataSet(sql, "Deposito")

        For Each row In ds.Tables(0).Rows
            Dim Prd As New DisponibilidadeDeposito(pDisponibilidade)
            Prd.CodigoDeposito = row("Deposito_Id")
            Prd.EndDeposito = row("EndDeposito_Id")
            Prd.SaldoInicialFiscal = row("SaldoInicialFiscal")
            Prd.SaldoinicialFisico = row("SaldoInicialFisico")
            Me.Add(Prd)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Disponibilidade As Disponibilidade
#End Region

#Region "Property"
    Public ReadOnly Property Disponibilidade As Disponibilidade
        Get
            Return _Disponibilidade
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Item As DisponibilidadeDeposito In Me
            If Disponibilidade.IUD = "D" Or Disponibilidade.IUD = "I" Then Item.IUD = Disponibilidade.IUD
            If Item.IUD <> "" Then
                Item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region
End Class

<Serializable()> _
Public Class DisponibilidadeDeposito
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pDisponibilidade As Disponibilidade)
        _Disponibilidade = pDisponibilidade
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Disponibilidade As Disponibilidade
    Private _CodigoDeposito As String
    Private _EndDeposito As Integer
    Private _Deposito As Cliente
    Private _SaldoInicialFiscal As Decimal
    Private _SaldoinicialFisico As Decimal
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Disponibilidade As Disponibilidade
        Get
            Return _Disponibilidade
        End Get
    End Property

    Public Property CodigoDeposito As String
        Get
            Return _CodigoDeposito
        End Get
        Set(ByVal value As String)
            _CodigoDeposito = value
            _Deposito = Nothing
        End Set
    End Property

    Public Property EndDeposito As Integer
        Get
            Return _EndDeposito
        End Get
        Set(ByVal value As Integer)
            _EndDeposito = value
            _Deposito = Nothing
        End Set
    End Property

    Public Property Deposito As Cliente
        Get
            If _Deposito Is Nothing AndAlso Me.CodigoDeposito.Length > 0 Then _Deposito = New Cliente(Me.CodigoDeposito, Me.EndDeposito)
            Return _Deposito
        End Get
        Set(ByVal value As Cliente)
            _Deposito = value
        End Set
    End Property

    Public ReadOnly Property NomeDeposito As String
        Get
            If Me.Deposito Is Nothing Then Return ""
            Return Me.Deposito.Nome & "...." & Me.Deposito.Cidade & "-" & Me.Deposito.CodigoEstado
        End Get
    End Property

    Public Property SaldoInicialFiscal As Decimal
        Get
            Return _SaldoInicialFiscal
        End Get
        Set(ByVal value As Decimal)
            _SaldoInicialFiscal = value
        End Set
    End Property

    Public Property SaldoinicialFisico As Decimal
        Get
            Return _SaldoinicialFisico
        End Get
        Set(ByVal value As Decimal)
            _SaldoinicialFisico = value
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
        Dim sql As String = ""
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into SaldoInicialDisponibilidadeDeposito(Disponibilidade_Id, Deposito_Id, EndDeposito_Id, SaldoInicialFiscal, SaldoInicialFisico) " & vbCrLf & _
                      " Values(" & Disponibilidade.CodigoDisponibilidade & ",'" & Me.CodigoDeposito & "'," & Me.EndDeposito & "," & Str(Me.SaldoInicialFiscal) & "," & Str(Me.SaldoinicialFisico) & ")"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete SaldoInicialDisponibilidadeDeposito" & vbCrLf & _
                      "  Where Disponibilidade_Id = " & Disponibilidade.CodigoDisponibilidade & vbCrLf & _
                      "    And Deposito_Id        ='" & Me.CodigoDeposito & "'" & vbCrLf & _
                      "    And EndDeposito_Id     = " & Me.EndDeposito
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
