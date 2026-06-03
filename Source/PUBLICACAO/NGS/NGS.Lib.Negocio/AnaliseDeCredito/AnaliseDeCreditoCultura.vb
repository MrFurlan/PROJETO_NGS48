Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAnaliseDeCreditoCultura
    Inherits List(Of AnaliseDeCreditoCultura)

#Region "Construtor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
        Dim sql As String = ""
        sql = "SELECT Analise_Id, Ano_Id, DefinicaoAno_Id, Safra_Id, Cultura_Id, AreaPlantio, RiscoCultura, CustoPortifolio, Producao, ReceitaCultura, CustoCultura" & vbCrLf & _
              "  FROM AnaliseDeCreditoCultura" & vbCrLf & _
              " Where Analise_Id = " & pAnaliseDeCredito.CodigoAnalise
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoCultura")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim AxC As New AnaliseDeCreditoCultura(pAnaliseDeCredito)
            AxC.CodigoSafra = row("Safra_Id")
            AxC.CodigoCultura = row("Cultura_Id")
            AxC.AreaPlantio = row("AreaPlantio")
            AxC.RiscoCultura = row("RiscoCultura")
            'AxC.CustoPortifolio = row("CustoPortifolio")
            'AxC.Producao = row("Producao")
            'AxC.ReceitaCultura = row("ReceitaCultura")
            'AxC.CustoCultura = row("CustoCultura")
            Me.Add(AxC)
        Next
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

    Public ReadOnly Property CoefRedutorRiscoCultura() As Decimal
        Get
            Dim crrc As Decimal = Decimal.Zero
            For Each row As AnaliseDeCreditoCultura In Me
                If row.PercCultura > 0 Then crrc += (row.PercCultura / 100) * row.RiscoCultura
            Next
            Select Case crrc
                Case Is > 2.5 : Return AnaliseDeCredito.ParametrosAnaliseDeCredito.PercReducaoRiscoAlto
                Case 1.6 To 2.5 : Return AnaliseDeCredito.ParametrosAnaliseDeCredito.PercReducaoRiscoMedio
                Case Is < 1.5 : Return AnaliseDeCredito.ParametrosAnaliseDeCredito.PercReducaoRiscoBaixo
            End Select
            Return crrc
        End Get
    End Property

    Public ReadOnly Property TotalReceitaCultura() As Decimal
        Get
            Dim RC As Decimal = 0
            For Each row As AnaliseDeCreditoCultura In Me
                RC += row.ReceitaCultura
            Next
            Return RC
        End Get
    End Property

    Public ReadOnly Property TotalCustoCultura() As Decimal
        Get
            Dim DC As Decimal = 0
            For Each row As AnaliseDeCreditoCultura In Me
                DC += row.CustoCultura
            Next
            Return DC
        End Get
    End Property

    Public ReadOnly Property TotalCustoPortifolio() As Decimal
        Get
            Dim CP As Decimal = 0
            For Each row As AnaliseDeCreditoCultura In Me
                CP += row.CustoPortifolio
            Next
            Return CP
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each cult As AnaliseDeCreditoCultura In Me
            If AnaliseDeCredito.IUD = "D" Or AnaliseDeCredito.IUD = "I" Then cult.IUD = AnaliseDeCredito.IUD
            If cult.IUD <> "" And cult.PercCultura > 0 Then
                cult.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class AnaliseDeCreditoCultura
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _AnaliseDeCredito As AnaliseDeCredito
    Private _ParametrosCultura As ParametrosAnaliseDeCreditoCultura
    Private _CodigoSafra As String = ""

    Private _CodigoCultura As Integer
    Private _Cultura As Cultura
    Private _NomeCultura As String = ""


    Private _AreaPlantio As Decimal
    Private _PercCultura As Decimal
    Private _RiscoCultura As Integer = 1

    Private _CustoPortifolio As Decimal
    Private _Producao As Decimal
    Private _ReceitaAgricola As Decimal
    Private _CustoProducao As Decimal
    Private _CustoArrendamento As Decimal
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

    Public ReadOnly Property ParametrosCultura() As ParametrosAnaliseDeCreditoCultura
        Get
            If _ParametrosCultura Is Nothing Then
                For Each PCult As ParametrosAnaliseDeCreditoCultura In AnaliseDeCredito.ParametrosAnaliseDeCredito.Culturas
                    If PCult.CodigoCultura = Me.CodigoCultura And PCult.CodigoSafra = Me.CodigoSafra Then
                        _ParametrosCultura = PCult
                    End If
                Next
                If _ParametrosCultura Is Nothing Then _ParametrosCultura = New ParametrosAnaliseDeCreditoCultura(Nothing)
            End If
            Return _ParametrosCultura
        End Get
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
            _Cultura = Nothing
            _NomeCultura = ""
        End Set
    End Property

    Public ReadOnly Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And Me.CodigoCultura > 0 Then _Cultura = New Cultura(Me.CodigoCultura)
            Return _Cultura
        End Get
    End Property

    Public ReadOnly Property NomeCultura() As String
        Get
            If Me.CodigoCultura > 0 And _NomeCultura.Length = 0 Then _NomeCultura = Cultura.Descricao
            Return _NomeCultura
        End Get
    End Property

    Public Property AreaPlantio() As Decimal
        Get
            Return _AreaPlantio
        End Get
        Set(ByVal value As Decimal)
            _AreaPlantio = value
        End Set
    End Property

    Public Property PercCultura() As Decimal
        Get
            Return _PercCultura
        End Get
        Set(ByVal value As Decimal)
            _PercCultura = value
        End Set
    End Property

    Public Property RiscoCultura() As Integer
        Get
            Return _RiscoCultura
        End Get
        Set(ByVal value As Integer)
            _RiscoCultura = value
        End Set
    End Property

    Public ReadOnly Property CustoPortifolio() As Decimal
        Get
            If ParametrosCultura.CodigoCultura > 0 Then
                Return AreaPlantio * ParametrosCultura.CustoPortifolioHa
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Producao() As Decimal
        Get
            If ParametrosCultura.CodigoCultura > 0 Then
                Return AreaPlantio * ParametrosCultura.Produtividade
            Else
                Return 0
            End If

        End Get
    End Property

    Public ReadOnly Property ReceitaCultura() As Decimal
        Get
            If ParametrosCultura.CodigoCultura > 0 Then
                Return AreaPlantio * ParametrosCultura.Produtividade * ParametrosCultura.PrecoSaco
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property CustoCultura() As Decimal
        Get
            If ParametrosCultura.CodigoCultura > 0 Then
                Return AreaPlantio * ParametrosCultura.CustoTotalHa
            Else
                Return 0
            End If
        End Get
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
                strSQL &= " Insert Into AnaliseDeCreditoCultura" & vbCrLf & _
                          " (Analise_Id, Ano_Id, DefinicaoAno_Id, Safra_Id, Cultura_Id, AreaPlantio, RiscoCultura, CustoPortifolio, Producao, ReceitaCultura, CustoCultura)" & vbCrLf & _
                          " Values (" & AnaliseDeCredito.CodigoAnalise & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & "," & vbCrLf & _
                          "'" & Me.CodigoSafra & "'," & Me.CodigoCultura & "," & Str(Me.AreaPlantio) & "," & Me.RiscoCultura & "," & Str(Me.CustoPortifolio) & "," & Str(Me.Producao) & "," & Str(Me.ReceitaCultura) & "," & Str(Me.CustoCultura) & ")"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = " Update AnaliseDeCreditoCultura set " & vbCrLf & _
                         "     AreaPlantio       = " & Str(Me.AreaPlantio) & vbCrLf & _
                         "    ,RiscoCultura      = " & Me.RiscoCultura & vbCrLf & _
                         "    ,CustoPortifolio   = " & Str(Me.CustoPortifolio) & vbCrLf & _
                         "    ,Producao          = " & Str(Me.Producao) & vbCrLf & _
                         "    ,ReceitaCultura    = " & Str(Me.ReceitaCultura) & vbCrLf & _
                         "    ,CustoCultura       = " & Str(Me.CustoCultura) & vbCrLf & _
                         "  Where Analise_Id      = " & AnaliseDeCredito.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & _
                         "    and Safra_Id        ='" & Me.CodigoSafra & "'" & vbCrLf & _
                         "    And Cultura_Id      = " & Me.CodigoCultura
                Sqls.Add(strSQL)
            Case "D"
                strSQL = " Delete AnaliseDeCreditoCultura" & vbCrLf & _
                         "  Where Analise_Id      = " & AnaliseDeCredito.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & _
                         "    and Safra_Id        ='" & Me.CodigoSafra & "'" & vbCrLf & _
                         "    And Cultura_Id      = " & Me.CodigoCultura
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
