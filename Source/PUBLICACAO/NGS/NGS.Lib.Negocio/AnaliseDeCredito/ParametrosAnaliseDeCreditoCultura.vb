Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListParametrosAnaliseDeCreditoCultura
    Inherits List(Of ParametrosAnaliseDeCreditoCultura)

#Region "Contrutor"
    Public Sub New(ByVal pParametrosAnaliseDeCredito As ParametrosAnaliseDeCredito)
        _Parametros = pParametrosAnaliseDeCredito

        Dim sql As String
        sql = "SELECT Ano_Id, DefinicaoAno_Id, Safra_id, Cultura_Id," & vbCrLf & _
              "       Produtividade, PrecoSaco, CustoTotalHA, CustoPortifolioHa" & vbCrLf & _
              "  FROM AnaliseDeCreditoParametroCultura" & vbCrLf & _
              " Where Ano_Id          =" & pParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   and DefinicaoAno_Id =" & pParametrosAnaliseDeCredito.DefinicaoAno & vbCrLf & _
              " order by DefinicaoAno_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoCultura")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PamCult As New ParametrosAnaliseDeCreditoCultura(pParametrosAnaliseDeCredito)
            PamCult.CodigoSafra = row("Safra_id")
            PamCult.CodigoCultura = row("Cultura_Id")
            PamCult.Produtividade = row("Produtividade")
            PamCult.PrecoSaco = row("PrecoSaco")
            PamCult.CustoTotalHa = row("CustoTotalHA")
            PamCult.CustoPortifolioHa = row("CustoPortifolioHa")
            Me.Add(PamCult)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Parametros As ParametrosAnaliseDeCredito
#End Region

#Region "Property"
    Public ReadOnly Property Parametros() As ParametrosAnaliseDeCredito
        Get
            Return _Parametros
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each cult As ParametrosAnaliseDeCreditoCultura In Me
            If Parametros.IUD = "D" Or Parametros.IUD = "I" Then cult.IUD = Parametros.IUD
            If cult.IUD <> "" Then
                cult.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ParametrosAnaliseDeCreditoCultura
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pParametrosAnaliseDeCredito As ParametrosAnaliseDeCredito)
        _Parametros = pParametrosAnaliseDeCredito
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Parametros As ParametrosAnaliseDeCredito
    Private _CodigoSafra As String = ""
    Private _Safra As Safra
    Private _CodigoCultura As Integer
    Private _Cultura As Cultura
    Private _Produtividade As Decimal
    Private _PrecoSaco As Decimal
    Private _CustoTotalHa As Decimal
    Private _CustoPortifolioHa As Decimal
    Private _CulturaPortifolios As ListCulturaPortifolio
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

    Public ReadOnly Property Parametros() As ParametrosAnaliseDeCredito
        Get
            Return _Parametros
        End Get
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
            _Safra = Nothing
        End Set
    End Property

    Public ReadOnly Property Safra() As Safra
        Get
            If _Safra Is Nothing And Me.CodigoSafra.Length > 0 Then _Safra = New Safra(Me.CodigoSafra)
            Return _Safra
        End Get
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
            _Cultura = Nothing
        End Set
    End Property

    Public ReadOnly Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And _CodigoCultura > 0 Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura
        End Get
    End Property

    Public ReadOnly Property NomeCultura() As String
        Get
            If Not Cultura Is Nothing Then Return Cultura.Descricao
            Return ""
        End Get
    End Property

    Public Property Produtividade() As Decimal
        Get
            Return _Produtividade
        End Get
        Set(ByVal value As Decimal)
            _Produtividade = value
        End Set
    End Property

    Public Property PrecoSaco() As Decimal
        Get
            Return _PrecoSaco
        End Get
        Set(ByVal value As Decimal)
            _PrecoSaco = value
        End Set
    End Property

    Public ReadOnly Property Receita() As Decimal
        Get
            Return _Produtividade * _PrecoSaco
        End Get
    End Property

    Public Property CustoTotalHa() As Decimal
        Get
            Return _CustoTotalHa
        End Get
        Set(ByVal value As Decimal)
            _CustoTotalHa = value
        End Set
    End Property

    Public Property CustoPortifolioHa() As Decimal
        Get
            If CulturaPortifolios IsNot Nothing AndAlso CulturaPortifolios.Count > 0 Then
                _CustoPortifolioHa = CulturaPortifolios.CustoTotalPortifolio
            End If
            Return _CustoPortifolioHa
        End Get
        Set(ByVal value As Decimal)
            _CustoPortifolioHa = value
        End Set
    End Property

    '*****************************************
    Public Property CulturaPortifolios() As ListCulturaPortifolio
        Get
            If _CulturaPortifolios Is Nothing Then _CulturaPortifolios = New ListCulturaPortifolio(Me)
            Return _CulturaPortifolios
        End Get
        Set(ByVal value As ListCulturaPortifolio)
            _CulturaPortifolios = value
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
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = " INSERT INTO AnaliseDeCreditoParametroCultura (Ano_Id, DefinicaoAno_Id,Safra_id, Cultura_Id, Produtividade, PrecoSaco, CustoTotalHA, CustoPortifolioHa)" & vbCrLf & _
                      " VALUES ('" & Parametros.Ano & "'," & Parametros.DefinicaoAno & ",'" & Me.CodigoSafra & "','" & Me.CodigoCultura & "'," & Str(Me.Produtividade) & "," & Str(Me.PrecoSaco) & "," & Str(Me.CustoTotalHa) & "," & Str(Me.CustoPortifolioHa) & ")"
                Sqls.Add(sql)
            Case "U"
                sql = " Update AnaliseDeCreditoParametroCultura Set" & vbCrLf & _
                      "     Produtividade           =" & Str(Me.Produtividade) & vbCrLf & _
                      "    ,PrecoSaco               =" & Str(Me.PrecoSaco) & vbCrLf & _
                      "    ,CustoTotalHA            =" & Str(Me.CustoTotalHa) & vbCrLf & _
                      "    ,CustoPortifolioHa       =" & Str(Me.CustoPortifolioHa) & vbCrLf & _
                      " Where Ano_Id            = " & Parametros.Ano & vbCrLf & _
                      "   and DefinicaoAno_Id = " & Parametros.DefinicaoAno & vbCrLf & _
                      "   and Safra_Id             ='" & Me.CodigoSafra & "'" & vbCrLf & _
                      "   and Cultura_Id        ='" & Me.CodigoCultura & "'"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete AnaliseDeCreditoParametroCultura " & vbCrLf & _
                      " Where Ano_Id            = " & Parametros.Ano & vbCrLf & _
                      "   and DefinicaoSafra_Id = " & Parametros.DefinicaoAno & vbCrLf & _
                      "   and Safra             ='" & Me.CodigoSafra & "'" & vbCrLf & _
                      "   and Cultura_Id        ='" & Me.CodigoCultura & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
