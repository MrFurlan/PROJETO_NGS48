Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListParametrosAnaliseDeCredito
    Inherits List(Of ParametrosAnaliseDeCredito)

    Public Sub New(ByVal pAno As Integer)
        Dim sql As String
        sql = "SELECT Ano_Id, DefinicaoAno_Id, DataDefinicao," & vbCrLf & _
              "       PercContasAbertas, PerguntasComportamentais," & vbCrLf & _
              "       PercLimiteCreditoA, PercLimiteCreditoB, PercLimiteCreditoC," & vbCrLf & _
              "       PercReducaoRiscoAlto, PercReducaoRiscoMedio, PercReducaoRiscoBaixo," & vbCrLf & _
              "       CustoArrendamentoHa, CotacaoDolar" & vbCrLf & _
              "  FROM AnaliseDeCreditoParametro" & vbCrLf & _
              " Where Ano_Id = '" & pAno & "'" & vbCrLf & _
              " order by DefinicaoAno_Id desc"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCredito")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Pam As New ParametrosAnaliseDeCredito
            Pam.Ano = row("Ano_Id")
            Pam.DefinicaoAno = row("DefinicaoAno_Id")
            Pam.DataDefinicao = row("DataDefinicao")
            Pam.PercContasAbertas = row("PercContasAbertas")
            Pam.PerguntasComportamentais = row("PerguntasComportamentais")
            Pam.PercLimiteCreditoA = row("PercLimiteCreditoA")
            Pam.PercLimiteCreditoB = row("PercLimiteCreditoB")
            Pam.PercLimiteCreditoC = row("PercLimiteCreditoC")
            Pam.PercReducaoRiscoAlto = row("PercReducaoRiscoAlto")
            Pam.PercReducaoRiscoMedio = row("PercReducaoRiscoMedio")
            Pam.PercReducaoRiscoBaixo = row("PercReducaoRiscoBaixo")
            Pam.CustoArrendamentoHa = row("CustoArrendamentoHa")
            Pam.CotacaoDolar = row("CotacaoDolar")
            Me.Add(Pam)
        Next
    End Sub

End Class

<Serializable()> _
Public Class ParametrosAnaliseDeCredito
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pAno As Integer, Optional ByVal pDefinicao As Integer = -1) 'se nao for informado carrega a ultima Definicao
        Dim sql As String
        sql = "SELECT Ano_Id, DefinicaoAno_Id, DataDefinicao, PercContasAbertas, PerguntasComportamentais," & vbCrLf & _
              "       PercLimiteCreditoA, PercLimiteCreditoB, PercLimiteCreditoC," & vbCrLf & _
              "       PercReducaoRiscoAlto, PercReducaoRiscoMedio, PercReducaoRiscoBaixo, CustoArrendamentoHa, CotacaoDolar" & vbCrLf & _
              "  FROM AnaliseDeCreditoParametro" & vbCrLf & _
              " Where Ano_id = " & pAno

        If pDefinicao > -1 Then
            sql &= " and DefinicaoAno_Id = " & pDefinicao
        Else
            sql &= " and DefinicaoAno_Id = (Select max(DefinicaoAno_Id) from AnaliseDeCreditoParametro where Ano_id = " & pAno & ")"
        End If

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoParametro")

        For Each row As DataRow In ds.Tables(0).Rows
            Me.Ano = row("Ano_Id")
            Me.DefinicaoAno = row("DefinicaoAno_Id")
            Me.DataDefinicao = row("DataDefinicao")
            Me.PercContasAbertas = row("PercContasAbertas")
            Me.PerguntasComportamentais = row("PerguntasComportamentais")
            Me.PercLimiteCreditoA = row("PercLimiteCreditoA")
            Me.PercLimiteCreditoB = row("PercLimiteCreditoB")
            Me.PercLimiteCreditoC = row("PercLimiteCreditoC")
            Me.PercReducaoRiscoAlto = row("PercReducaoRiscoAlto")
            Me.PercReducaoRiscoMedio = row("PercReducaoRiscoMedio")
            Me.PercReducaoRiscoBaixo = row("PercReducaoRiscoBaixo")
            Me.CustoArrendamentoHa = row("CustoArrendamentoHa")
            Me.CotacaoDolar = row("CotacaoDolar")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Ano As Integer
    Private _DefinicaoAno As Integer
    Private _DataDefinicao As Date
    Private _PercContasAbertas As Decimal
    Private _PerguntasComportamentais As Boolean = True
    Private _PercLimiteCreditoA As Decimal
    Private _PercLimiteCreditoB As Decimal
    Private _PercLimiteCreditoC As Decimal
    Private _PercReducaoRiscoAlto As Decimal
    Private _PercReducaoRiscoMedio As Decimal
    Private _PercReducaoRiscoBaixo As Decimal
    Private _CustoArrendamentoHa As Decimal
    Private _CotacaoDolar As Decimal
    Private _Culturas As ListParametrosAnaliseDeCreditoCultura
    Private _Perguntas As ListParametrosAnaliseDeCreditoPergunta
#End Region

#Region "property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property DefinicaoAno() As Integer
        Get
            Return _DefinicaoAno
        End Get
        Set(ByVal value As Integer)
            _DefinicaoAno = value
        End Set
    End Property

    Public Property DataDefinicao() As Date
        Get
            Return _DataDefinicao
        End Get
        Set(ByVal value As Date)
            _DataDefinicao = value
        End Set
    End Property

    Public Property PercContasAbertas() As Decimal
        Get
            Return _PercContasAbertas
        End Get
        Set(ByVal value As Decimal)
            _PercContasAbertas = value
        End Set
    End Property

    Public Property PerguntasComportamentais() As Boolean
        Get
            Return _PerguntasComportamentais
        End Get
        Set(ByVal value As Boolean)
            _PerguntasComportamentais = value
        End Set
    End Property

    Public Property PercLimiteCreditoA() As Decimal
        Get
            Return _PercLimiteCreditoA
        End Get
        Set(ByVal value As Decimal)
            _PercLimiteCreditoA = value
        End Set
    End Property

    Public Property PercLimiteCreditoB() As Decimal
        Get
            Return _PercLimiteCreditoB
        End Get
        Set(ByVal value As Decimal)
            _PercLimiteCreditoB = value
        End Set
    End Property

    Public Property PercLimiteCreditoC() As Decimal
        Get
            Return _PercLimiteCreditoC
        End Get
        Set(ByVal value As Decimal)
            _PercLimiteCreditoC = value
        End Set
    End Property

    Public Property PercReducaoRiscoAlto() As Decimal
        Get
            Return _PercReducaoRiscoAlto
        End Get
        Set(ByVal value As Decimal)
            _PercReducaoRiscoAlto = value
        End Set
    End Property

    Public Property PercReducaoRiscoMedio() As Decimal
        Get
            Return _PercReducaoRiscoMedio
        End Get
        Set(ByVal value As Decimal)
            _PercReducaoRiscoMedio = value
        End Set
    End Property

    Public Property PercReducaoRiscoBaixo() As Decimal
        Get
            Return _PercReducaoRiscoBaixo
        End Get
        Set(ByVal value As Decimal)
            _PercReducaoRiscoBaixo = value
        End Set
    End Property

    Public Property CustoArrendamentoHa() As Decimal
        Get
            Return _CustoArrendamentoHa
        End Get
        Set(ByVal value As Decimal)
            _CustoArrendamentoHa = value
        End Set
    End Property

    Public Property CotacaoDolar() As Decimal
        Get
            Return _CotacaoDolar
        End Get
        Set(ByVal value As Decimal)
            _CotacaoDolar = value
        End Set
    End Property

    '*****************************************
    Public Property Culturas() As ListParametrosAnaliseDeCreditoCultura
        Get
            If _Culturas Is Nothing Then _Culturas = New ListParametrosAnaliseDeCreditoCultura(Me)
            Return _Culturas
        End Get
        Set(ByVal value As ListParametrosAnaliseDeCreditoCultura)
            _Culturas = value
        End Set
    End Property

    Public Property Perguntas() As ListParametrosAnaliseDeCreditoPergunta
        Get
            If _Perguntas Is Nothing Then _Perguntas = New ListParametrosAnaliseDeCreditoPergunta(Me)
            Return _Perguntas
        End Get
        Set(ByVal value As ListParametrosAnaliseDeCreditoPergunta)
            _Perguntas = value
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
                strSQL &= " Insert Into AnaliseDeCreditoParametro (Ano_Id, DefinicaoAno_Id, DataDefinicao, PercContasAbertas, PerguntasComportamentais,PercLimiteCreditoA, PercLimiteCreditoB, PercLimiteCreditoC, PercReducaoRiscoAlto, PercReducaoRiscoMedio, PercReducaoRiscoBaixo, CustoArrendamentoHa, CotacaoDolar)" & vbCrLf & _
                          " Values (" & Me.Ano & "," & Me.DefinicaoAno & ",'" & Me.DataDefinicao.ToString("yyyy-MM-dd") & "'," & Str(Me.PercContasAbertas) & "," & IIf(Me.PerguntasComportamentais, "1", "0") & "," & Str(Me.PercLimiteCreditoA) & "," & Str(Me.PercLimiteCreditoB) & "," & Str(Me.PercLimiteCreditoC) & "," & Str(Me.PercReducaoRiscoAlto) & "," & Str(Me.PercReducaoRiscoMedio) & "," & Str(Me.PercReducaoRiscoBaixo) & "," & Str(Me.CustoArrendamentoHa) & "," & Str(Me.CotacaoDolar) & ")"
                Sqls.Add(strSQL)
                Culturas.SalvarSql(Sqls)
                Perguntas.SalvarSql(Sqls)
            Case "U"
                strSQL = " Update AnaliseDeCreditoParametro set " & vbCrLf & _
                         "      DataDefinicao            ='" & Me.DataDefinicao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                         "     ,PercContasAbertas        = " & Str(Me.PercContasAbertas) & vbCrLf & _
                         "     ,PerguntasComportamentais = " & IIf(Me.PerguntasComportamentais, "1", "0") & vbCrLf & _
                         "     ,PercLimiteCreditoA       = " & Str(Me.PercLimiteCreditoA) & vbCrLf & _
                         "     ,PercLimiteCreditoB       = " & Str(Me.PercLimiteCreditoB) & vbCrLf & _
                         "     ,PercLimiteCreditoC       = " & Str(Me.PercLimiteCreditoC) & vbCrLf & _
                         "     ,PercReducaoRiscoAlto     = " & Str(Me.PercReducaoRiscoAlto) & vbCrLf & _
                         "     ,PercReducaoRiscoMedio    = " & Str(Me.PercReducaoRiscoMedio) & vbCrLf & _
                         "     ,PercReducaoRiscoBaixo    = " & Str(Me.PercReducaoRiscoBaixo) & vbCrLf & _
                         "     ,CustoArrendamentoHa      = " & Str(Me.CustoArrendamentoHa) & vbCrLf & _
                         "     ,CotacaoDolar             = " & Str(Me.CotacaoDolar) & vbCrLf & _
                         "  Where Ano_Id          = " & Me.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.DefinicaoAno
                Sqls.Add(strSQL)
                Culturas.SalvarSql(Sqls)
                Perguntas.SalvarSql(Sqls)
            Case "D"
                Perguntas.SalvarSql(Sqls)
                Culturas.SalvarSql(Sqls)
                strSQL = " Delete AnaliseDeCreditoParametro" & vbCrLf & _
                         "  Where Ano_Id          = " & Me.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Me.DefinicaoAno
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class


