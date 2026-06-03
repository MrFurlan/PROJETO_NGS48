Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPlanoDeCusto
    Inherits List(Of PlanoDeCusto)

#Region "Contrutor"
    Public Sub New()
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Codigo_Id, Descricao, Totalizador, SinalPeso, SinalValor, DebitoMercadoria, CreditoMercadoria, HistoricoMercadoria," & vbCrLf & _
              "       DebitoFrete, CreditoFrete, HistoricoFrete, isnull(Desdobramento,'False') as Desdobramento, isnull(SaldoInicial,'') as SaldoInicial, isnull(FaseDoTotalizador,0) as FaseDoTotalizador, Rateio" & vbCrLf & _
              "  FROM PlanoDeCustos "

        ds = Banco.ConsultaDataSet(sql, "PlanoDeCusto")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PC As New PlanoDeCusto
            PC.Codigo = row("Codigo_Id")
            PC.Descricao = row("Descricao")
            PC.CodigoTotalizador = row("Totalizador")
            PC.SinalPeso = row("SinalPeso")
            PC.SinalValor = row("SinalValor")
            PC.CodigoDebitoMercadoria = row("DebitoMercadoria")
            PC.CodigoCreditoMercadoria = row("CreditoMercadoria")
            PC.CodigoHistoricoMercadoria = row("HistoricoMercadoria")
            PC.CodigoDebitoFrete = row("DebitoFrete")
            PC.CodigoCreditoFrete = row("CreditoFrete")
            PC.CodigoHistoricoFrete = row("HistoricoFrete")
            PC.Desdobramento = row("Desdobramento") = "True"
            PC.SaldoInicial = row("SaldoInicial") = "True"
            PC.FaseDoTotalizador = row("FaseDoTotalizador")
            PC.Rateio = row("Rateio")
            Me.Add(PC)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class PlanoDeCusto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal CodigoPlanoDeCusto As Integer)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Codigo_Id, Descricao, Totalizador, SinalPeso, SinalValor, DebitoMercadoria, CreditoMercadoria, HistoricoMercadoria," & vbCrLf & _
              "       DebitoFrete, CreditoFrete, HistoricoFrete, isnull(Desdobramento,'False') as Desdobramento, isnull(SaldoInicial,'') as SaldoInicial, FaseDoTotalizador, Rateio" & vbCrLf & _
              "  FROM PlanoDeCustos " & vbCrLf & _
              " Where Codigo_Id = " & CodigoPlanoDeCusto

        ds = Banco.ConsultaDataSet(sql, "PlanoDeCusto")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = row("Codigo_Id")
        _Descricao = row("Descricao")
        _CodigoTotalizador = row("Totalizador")
        _SinalPeso = row("SinalPeso")
        _SinalValor = row("SinalValor")
        _CodigoDebitoMercadoria = row("DebitoMercadoria")
        _CodigoCreditoMercadoria = row("CreditoMercadoria")
        _CodigoHistoricoMercadoria = row("HistoricoMercadoria")
        _CodigoDebitoFrete = row("DebitoFrete")
        _CodigoCreditoFrete = row("CreditoFrete")
        _CodigoHistoricoFrete = row("HistoricoFrete")
        _Desdobramento = row("Desdobramento") = "True"
        _SaldoInicial = row("SaldoInicial") = "True"
        _FaseDoTotalizador = row("FaseDoTotalizador")
        _Rateio = row("Rateio")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _CodigoTotalizador As Integer
    Private _Totalizador As PlanoDeCusto
    Private _SinalPeso As String
    Private _SinalValor As String
    Private _CodigoDebitoMercadoria As String = ""
    Private _DebitoMercadoria As PlanoDeConta
    Private _CodigoCreditoMercadoria As String = ""
    Private _CreditoMercadoria As PlanoDeConta
    Private _CodigoHistoricoMercadoria As Integer
    Private _HistoricoMercadoria As Historico
    Private _CodigoDebitoFrete As String = ""
    Private _DebitoFrete As PlanoDeConta
    Private _CodigoCreditoFrete As String = ""
    Private _CreditoFrete As PlanoDeConta
    Private _CodigoHistoricoFrete As Integer
    Private _HistoricoFrete As Historico
    Private _Desdobramento As Boolean
    Private _SaldoInicial As Boolean
    Private _FaseDoTotalizador As Integer
    Private _Rateio As Boolean
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property CodigoTotalizador() As Integer
        Get
            Return _CodigoTotalizador
        End Get
        Set(ByVal value As Integer)
            _CodigoTotalizador = value
        End Set
    End Property

    Public Property Totalizador() As PlanoDeCusto
        Get
            If _Totalizador Is Nothing And _CodigoTotalizador > 0 Then _Totalizador = New PlanoDeCusto(_CodigoTotalizador)
            Return _Totalizador
        End Get
        Set(ByVal value As PlanoDeCusto)
            _Totalizador = value
        End Set
    End Property

    Public Property SinalPeso() As String
        Get
            Return _SinalPeso
        End Get
        Set(ByVal value As String)
            _SinalPeso = value
        End Set
    End Property

    Public Property SinalValor() As String
        Get
            Return _SinalValor
        End Get
        Set(ByVal value As String)
            _SinalValor = value
        End Set
    End Property

    Public Property CodigoDebitoMercadoria() As String
        Get
            Return _CodigoDebitoMercadoria
        End Get
        Set(ByVal value As String)
            _CodigoDebitoMercadoria = value
        End Set
    End Property

    Public Property DebitoMercadoria() As PlanoDeConta
        Get
            If _DebitoMercadoria Is Nothing And _CodigoDebitoMercadoria.Length > 0 Then _DebitoMercadoria = New PlanoDeConta("", 0, _CodigoDebitoMercadoria)
            Return _DebitoMercadoria
        End Get
        Set(ByVal value As PlanoDeConta)
            _DebitoMercadoria = value
        End Set
    End Property

    Public Property CodigoCreditoMercadoria() As String
        Get
            Return _CodigoCreditoMercadoria
        End Get
        Set(ByVal value As String)
            _CodigoCreditoMercadoria = value
        End Set
    End Property

    Public Property CreditoMercadoria() As PlanoDeConta
        Get
            If _CreditoMercadoria Is Nothing And _CodigoCreditoMercadoria.Length > 0 Then _CreditoMercadoria = New PlanoDeConta("", 0, _CodigoCreditoMercadoria)
            Return _CreditoMercadoria
        End Get
        Set(ByVal value As PlanoDeConta)
            _CreditoMercadoria = value
        End Set
    End Property

    Public Property CodigoHistoricoMercadoria() As Integer
        Get
            Return _CodigoHistoricoMercadoria
        End Get
        Set(ByVal value As Integer)
            _CodigoHistoricoMercadoria = value
        End Set
    End Property

    Public Property HistoricoMercadoria() As Historico
        Get
            If _HistoricoMercadoria Is Nothing And _CodigoHistoricoMercadoria > 0 Then _HistoricoMercadoria = New Historico(_CodigoHistoricoMercadoria)
            Return _HistoricoMercadoria
        End Get
        Set(ByVal value As Historico)
            _HistoricoMercadoria = value
        End Set
    End Property

    Public Property CodigoDebitoFrete() As String
        Get
            Return _CodigoDebitoFrete
        End Get
        Set(ByVal value As String)
            _CodigoDebitoFrete = value
        End Set
    End Property

    Public Property DebitoFrete() As PlanoDeConta
        Get
            If _DebitoFrete Is Nothing And _CodigoDebitoFrete.Length > 0 Then _DebitoFrete = New PlanoDeConta("", 0, _CodigoDebitoFrete)
            Return _DebitoFrete
        End Get
        Set(ByVal value As PlanoDeConta)
            _DebitoFrete = value
        End Set
    End Property

    Public Property CodigoCreditoFrete() As String
        Get
            Return _CodigoCreditoFrete
        End Get
        Set(ByVal value As String)
            _CodigoCreditoFrete = value
        End Set
    End Property

    Public Property CreditoFrete() As PlanoDeConta
        Get
            If _CreditoFrete Is Nothing And _CodigoCreditoFrete.Length > 0 Then _CreditoFrete = New PlanoDeConta("", 0, _CodigoCreditoFrete)
            Return _CreditoFrete
        End Get
        Set(ByVal value As PlanoDeConta)
            _CreditoFrete = value
        End Set
    End Property

    Public Property CodigoHistoricoFrete() As Integer
        Get
            Return _CodigoHistoricoFrete
        End Get
        Set(ByVal value As Integer)
            _CodigoHistoricoFrete = value
        End Set
    End Property

    Public Property HistoricoFrete() As Historico
        Get
            If _HistoricoFrete Is Nothing And _CodigoHistoricoFrete > 0 Then _HistoricoFrete = New Historico(_CodigoHistoricoFrete)
            Return _HistoricoFrete
        End Get
        Set(ByVal value As Historico)
            _HistoricoFrete = value
        End Set
    End Property

    Public Property Desdobramento() As Boolean
        Get
            Return _Desdobramento
        End Get
        Set(ByVal value As Boolean)
            _Desdobramento = value
        End Set
    End Property

    Public Property SaldoInicial() As Boolean
        Get
            Return _SaldoInicial
        End Get
        Set(ByVal value As Boolean)
            _SaldoInicial = value
        End Set
    End Property

    Public Property FaseDoTotalizador() As Integer
        Get
            Return _FaseDoTotalizador
        End Get
        Set(ByVal value As Integer)
            _FaseDoTotalizador = value
        End Set
    End Property

    Public Property Rateio() As Boolean
        Get
            Return _Rateio
        End Get
        Set(ByVal value As Boolean)
            _Rateio = value
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
                Sql = " INSERT INTO PlanoDeCustos(Codigo_Id, Descricao, Totalizador, SinalPeso, SinalValor, DebitoMercadoria, CreditoMercadoria, HistoricoMercadoria, DebitoFrete, CreditoFrete, HistoricoFrete, Desdobramento, SaldoInicial, FaseDoTotalizador, Rateio) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "'," & _CodigoTotalizador & ",'" & _SinalPeso & "','" & _SinalValor & "','" & _CodigoDebitoMercadoria & "','" & _CodigoCreditoMercadoria & "'," & _CodigoHistoricoMercadoria & ",'" & _CodigoDebitoFrete & "','" & _CodigoCreditoFrete & "'," & _CodigoHistoricoFrete & ",'" & IIf(_Desdobramento, "True", "False") & "','" & IIf(_SaldoInicial, "True", "False") & "'," & _FaseDoTotalizador & "," & CByte(_Rateio) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update PlanoDeCustos set" & vbCrLf & _
                      "  Descricao           ='" & _Descricao & "'" & vbCrLf & _
                      " ,Totalizador         = " & _CodigoTotalizador & vbCrLf & _
                      " ,SinalPeso           ='" & _SinalPeso & "'" & vbCrLf & _
                      " ,SinalValor          ='" & _SinalValor & "'" & vbCrLf & _
                      " ,DebitoMercadoria    ='" & _CodigoDebitoMercadoria & "'" & vbCrLf & _
                      " ,CreditoMercadoria   ='" & _CodigoCreditoMercadoria & "'" & vbCrLf & _
                      " ,HistoricoMercadoria = " & _CodigoHistoricoMercadoria & vbCrLf & _
                      " ,DebitoFrete         ='" & _CodigoDebitoFrete & "'" & vbCrLf & _
                      " ,CreditoFrete        ='" & _CodigoDebitoFrete & "'" & vbCrLf & _
                      " ,HistoricoFrete      = " & _CodigoHistoricoFrete & vbCrLf & _
                      " ,Desdobramento       ='" & IIf(_Desdobramento, "True", "False") & "'" & vbCrLf & _
                      " ,SaldoInicial        ='" & IIf(_SaldoInicial, "True", "False") & "'" & vbCrLf & _
                      " ,FaseDoTotalizador   = " & _FaseDoTotalizador & vbCrLf & _
                      " ,Rateio              = " & CByte(_Rateio) & vbCrLf & _
                      " Where Codigo_Id = " & _Codigo
            Case "D"
                Sql = " DELETE PlanoDeCustos" & vbCrLf & _
                      "  WHERE Codigo_Id      ='" & _Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class