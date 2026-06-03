Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAnaliseDeCreditoPergunta
    Inherits List(Of AnaliseDeCreditoPergunta)

#Region "Construtor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
        If pAnaliseDeCredito.IUD = "U" Then
            Dim ds As DataSet
            Dim sql As String = ""
            sql = "SELECT ACP.Analise_Id, ACP.Ano_Id, ACP.DefinicaoAno_Id, ACP.Pergunta_Id, ACPP.Descricao, ACPP.PercPeso, ACP.Resposta" & vbCrLf & _
                  "  FROM AnaliseDeCreditoPergunta ACP" & vbCrLf & _
                  " INNER JOIN AnaliseDeCreditoParametroPergunta ACPP" & vbCrLf & _
                  "    ON ACP.Ano_Id          = ACPP.Ano_Id " & vbCrLf & _
                  "   AND ACP.DefinicaoAno_Id = ACPP.DefinicaoAno_Id" & vbCrLf & _
                  "   AND ACP.Pergunta_Id     = ACPP.Pergunta_Id" & vbCrLf & _
                  " Where ACP.Analise_Id      = " & pAnaliseDeCredito.CodigoAnalise & vbCrLf & _
                  "   and ACP.Ano_Id          = " & pAnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                  "   and ACP.DefinicaoAno_Id = " & pAnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno

            Dim Banco As New AcessaBanco
            ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoPergunta")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim AnP As New AnaliseDeCreditoPergunta(pAnaliseDeCredito)
                AnP.CodigoPergunta = row("Pergunta_Id")
                AnP.DescPergunta = row("Descricao")
                AnP.PercPeso = row("PercPeso")
                AnP.Resposta = row("Resposta")
                Me.Add(AnP)
            Next
        Else
            For Each ACPP As ParametrosAnaliseDeCreditoPergunta In pAnaliseDeCredito.ParametrosAnaliseDeCredito.Perguntas
                Dim AnP As New AnaliseDeCreditoPergunta(pAnaliseDeCredito)
                AnP.CodigoPergunta = ACPP.CodigoPergunta
                AnP.DescPergunta = ACPP.Descricao
                AnP.PercPeso = ACPP.PercPeso
                AnP.Resposta = 3
                Me.Add(AnP)
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

    Public ReadOnly Property RatingCreditoCliente() As String
        Get
            Dim TotalPontos As Decimal = 0
            For Each row As AnaliseDeCreditoPergunta In Me
                TotalPontos += row.PesoResposta
            Next
            Select Case TotalPontos
                Case Is > 2.5
                    AnaliseDeCredito.LimiteCredito = "A"
                    Return "A"
                Case 2.2 To 2.5
                    AnaliseDeCredito.LimiteCredito = "B"
                    Return "B"
                Case Is < 2.2
                    AnaliseDeCredito.LimiteCredito = "C"
                    Return "C"
            End Select
            Return TotalPontos
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Perg As AnaliseDeCreditoPergunta In Me
            If AnaliseDeCredito.IUD = "D" Or AnaliseDeCredito.IUD = "I" Then Perg.IUD = AnaliseDeCredito.IUD
            If Perg.IUD <> "" Then
                Perg.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class AnaliseDeCreditoPergunta
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pAnaliseDeCredito As AnaliseDeCredito)
        _AnaliseDeCredito = pAnaliseDeCredito
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _AnaliseDeCredito As AnaliseDeCredito
    Private _CodigoPergunta As Integer
    Private _DescPergunta As String
    Private _PercPeso As Decimal
    Private _Resposta As Integer
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

    Public Property CodigoPergunta() As Integer
        Get
            Return _CodigoPergunta
        End Get
        Set(ByVal value As Integer)
            _CodigoPergunta = value
        End Set
    End Property

    Public Property DescPergunta() As String
        Get
            Return _DescPergunta
        End Get
        Set(ByVal value As String)
            _DescPergunta = value
        End Set
    End Property

    Public Property PercPeso() As Decimal
        Get
            Return _PercPeso
        End Get
        Set(ByVal value As Decimal)
            _PercPeso = value
        End Set
    End Property

    Public Property Resposta() As Integer
        Get
            Return _Resposta
        End Get
        Set(ByVal value As Integer)
            _Resposta = value
        End Set
    End Property

    Public ReadOnly Property PesoResposta() As Decimal
        Get
            Return (PercPeso / 100) * Resposta
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
                strSQL &= " Insert Into AnaliseDeCreditoPergunta(Analise_Id, Ano_Id, DefinicaoAno_Id, Pergunta_Id, Resposta)" & vbCrLf & _
                          " Values (" & AnaliseDeCredito.CodigoAnalise & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & "," & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & "," & Me.CodigoPergunta & "," & Me.Resposta & ")"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = " Update AnaliseDeCreditoPergunta set " & vbCrLf & _
                         "    Resposta   =" & Me.Resposta & vbCrLf & _
                         "  Where Analise_Id      = " & AnaliseDeCredito.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & _
                         "    and Pergunta_Id     = " & Me.CodigoPergunta
                Sqls.Add(strSQL)
            Case "D"
                strSQL = " Delete AnaliseDeCreditoPergunta" & vbCrLf & _
                         "  Where Analise_Id      = " & AnaliseDeCredito.CodigoAnalise & vbCrLf & _
                         "    and Ano_Id          = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & AnaliseDeCredito.ParametrosAnaliseDeCredito.DefinicaoAno & _
                         "    and Pergunta_Id     = " & Me.CodigoPergunta
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
