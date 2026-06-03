Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListParametrosAnaliseDeCreditoPergunta
    Inherits List(Of ParametrosAnaliseDeCreditoPergunta)

#Region "Contrutor"
    Public Sub New(ByVal pParametrosAnaliseDeCredito As ParametrosAnaliseDeCredito)
        _Parametros = pParametrosAnaliseDeCredito

        Dim sql As String
        sql = "SELECT Ano_Id, DefinicaoAno_Id, Pergunta_Id, Descricao, PercPeso" & vbCrLf & _
              "  FROM AnaliseDeCreditoParametroPergunta" & vbCrLf & _
              " Where Ano_Id          =" & pParametrosAnaliseDeCredito.Ano & vbCrLf & _
              "   and DefinicaoAno_Id =" & pParametrosAnaliseDeCredito.DefinicaoAno & vbCrLf & _
              " order by DefinicaoAno_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "AnaliseDeCreditoPergunta")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim PamPerg As New ParametrosAnaliseDeCreditoPergunta(pParametrosAnaliseDeCredito)
            PamPerg.CodigoPergunta = row("Pergunta_Id")
            PamPerg.Descricao = row("Descricao")
            PamPerg.PercPeso = row("PercPeso")
            Me.Add(PamPerg)
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
        For Each Perg As ParametrosAnaliseDeCreditoPergunta In Me
            If Parametros.IUD = "D" Or Parametros.IUD = "I" Then Perg.IUD = Parametros.IUD
            If Perg.IUD <> "" Then
                Perg.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class ParametrosAnaliseDeCreditoPergunta
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pParametroAnaliseCredito As ParametrosAnaliseDeCredito)
        _Parametros = pParametroAnaliseCredito
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Parametros As ParametrosAnaliseDeCredito
    Private _CodigoPergunta As Integer
    Private _Descricao As String = ""
    Private _PercPeso As Decimal
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

    Public Property CodigoPergunta() As Integer
        Get
            Return _CodigoPergunta
        End Get
        Set(ByVal value As Integer)
            _CodigoPergunta = value
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

    Public Property PercPeso() As Decimal
        Get
            Return _PercPeso
        End Get
        Set(ByVal value As Decimal)
            If value > 100 Or value < 0 Then value = 100
            _PercPeso = value
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
                strSQL &= " Insert Into AnaliseDeCreditoParametroPergunta (Ano_Id, DefinicaoAno_Id, Pergunta_Id, Descricao, PercPeso)" & vbCrLf & _
                          " Values (" & Parametros.Ano & "," & Parametros.DefinicaoAno & "," & Me.CodigoPergunta & ",'" & Me.Descricao & "'," & Str(Me.PercPeso) & ")"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = " Update AnaliseDeCreditoParametroPergunta set " & vbCrLf & _
                         "    Descricao   ='" & Me.Descricao & "'" & vbCrLf & _
                         "   ,PercPeso    = " & Str(Me.PercPeso) & vbCrLf & _
                         "  Where Ano_Id          = " & Parametros.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Parametros.DefinicaoAno & _
                         "    and Pergunta_Id     = " & Me.CodigoPergunta
                Sqls.Add(strSQL)
            Case "D"
                strSQL = " Delete AnaliseDeCreditoParametroPergunta" & vbCrLf & _
                         "  Where Ano_Id          = " & Parametros.Ano & vbCrLf & _
                         "    and DefinicaoAno_Id = " & Parametros.DefinicaoAno & _
                         "    and Pergunta_Id     = " & Me.CodigoPergunta
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class