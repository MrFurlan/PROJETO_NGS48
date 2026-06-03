Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClassificacaoDeProduto
    Inherits List(Of ClassificacaoDeProduto)

    Public Sub New()

    End Sub

End Class

<Serializable()> _
Public Class ClassificacaoDeProduto
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Classificacao As Classificacao
    Private _Produto As Produto
    Private _Analise As Analise
    Private _Sequencia As Integer
    Private _FaixaInicial As Double
    Private _FaixaFinal As Double
    Private _Indice As Double
    Private _Tolerancia As Double
    Private _FaixaValida As String
    Private _Observacoes As String
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

    Public Property Classificacao() As Classificacao
        Get
            If _Classificacao Is Nothing Then
                _Classificacao = New Classificacao()
            End If

            Return _Classificacao
        End Get
        Set(ByVal value As Classificacao)
            _Classificacao = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing Then
                _Produto = New Produto()
            End If

            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property Analise() As Analise
        Get
            If _Analise Is Nothing Then
                _Analise = New Analise()
            End If

            Return _Analise
        End Get
        Set(ByVal value As Analise)
            _Analise = value
        End Set
    End Property

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property FaixaInicial() As Double
        Get
            Return _FaixaInicial
        End Get
        Set(ByVal value As Double)
            _FaixaInicial = value
        End Set
    End Property

    Public Property FaixaFinal() As Double
        Get
            Return _FaixaFinal
        End Get
        Set(ByVal value As Double)
            _FaixaFinal = value
        End Set
    End Property

    Public Property Indide() As Double
        Get
            Return _Indice
        End Get
        Set(ByVal value As Double)
            _Indice = value
        End Set
    End Property

    Public Property Tolerancia As Double
        Get
            Return _Tolerancia
        End Get
        Set(ByVal value As Double)
            _Tolerancia = value
        End Set
    End Property

    Public Property FaixaValida() As String
        Get
            Return _FaixaValida
        End Get
        Set(ByVal value As String)
            _FaixaValida = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into Classificacoes(Tabela_Id, " &
                            "Produto_Id, " &
                            "Analise_Id, " &
                            "Sequencia_Id, " &
                            "FaixaInicial, " &
                            "FaixaFinal, " &
                            "Indice, " &
                            "Tolerancia, " &
                            "FaixaValida, " &
                            "Observacoes)" &
                      " Values (" &
                            _Classificacao.Codigo & "," &
                            _Produto.Codigo & "," &
                            _Analise.Codigo & ", " &
                            _Sequencia & "," &
                             _FaixaInicial.ToString.Replace(",", ".") & "," &
                             _FaixaFinal.ToString.Replace(",", ".") & "," &
                             _Indice.ToString.Replace(",", ".") & "," &
                             _Tolerancia.ToString.Replace(",", ".") & ",'" &
                            _FaixaValida & "','" +
                            _Observacoes + "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update Classificacoes set " &
                            "FaixaInicial = " & _FaixaInicial.ToString.Replace(",", ".") &
                            ",FaixaFinal = " & _FaixaFinal.ToString.Replace(",", ".") &
                            ",Indice = " & _Indice.ToString.Replace(",", ".") &
                            ",Tolerancia = " & _Tolerancia.ToString.Replace(",", ".") &
                            ",FaixaValida = '" & _FaixaValida &
                            "',Observacoes = '" & _Observacoes &
                      "' Where" &
                            " Tabela_Id = " & _Classificacao.Codigo &
                            " AND Produto_Id = " & _Produto.Codigo &
                            " AND Analise_Id = " & Analise.Codigo &
                            " AND Sequencia_Id = " & _Sequencia
                Sqls.Add(sql)
            Case "D"
                sql = " DELETE Classificacoes" &
                      " WHERE" &
                            " Tabela_Id = " & _Classificacao.Codigo &
                            " AND Produto_Id = " & _Produto.Codigo &
                            " AND Analise_Id = " & Analise.Codigo &
                            " AND Sequencia_Id = " & _Sequencia
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
