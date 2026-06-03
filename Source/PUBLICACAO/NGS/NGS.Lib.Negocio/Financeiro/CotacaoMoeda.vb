Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCotacao
    Inherits List(Of Cotacao)

    Sub New()

    End Sub

End Class

<Serializable()> _
Public Class Cotacao
    Implements IBaseEntity

#Region "Construtores"
    Sub New(ByVal pIndexador As Integer, ByVal pData As Date)
        Dim sql As String
        Dim ds As New DataSet
        Dim dr As DataRow
        sql = "select Indexador_Id, Data_Id, indice, isnull(Realizado,'P') as Realizado, UsuarioLiberacao, UsuarioLiberacaoData, UsuarioAlteracao, UsuarioAlteracaoData" & vbCrLf & _
              "  from Cotacoes " & vbCrLf & _
              " where Indexador_Id = " & pIndexador & " and Data_Id = '" & pData.ToString("yyyy-MM-dd") & "'" & vbCrLf

        ds = banco.ConsultaDataSet(sql, "Cotacoes")

        For Each dr In ds.Tables(0).Rows
            _CodigoIndexador = dr("Indexador_Id")
            _Data = dr("Data_Id")
            _Indice = dr("indice")
            _Realizado = dr("Realizado") = "R"
            If Not IsDBNull(dr("UsuarioLiberacao")) Then _UsuarioLiberacao = dr("UsuarioLiberacao")
            If Not IsDBNull(dr("UsuarioLiberacaoData")) Then _UsuarioLiberacaoData = dr("UsuarioLiberacaoData")
            If Not IsDBNull(dr("UsuarioAlteracao")) Then _UsuarioAlteracao = dr("UsuarioAlteracao")
            If Not IsDBNull(dr("UsuarioAlteracaoData")) Then _UsuarioAlteracaoData = dr("UsuarioAlteracaoData")
        Next
    End Sub
#End Region

#Region "Fields"
    Private banco As New AcessaBanco
    Private _IUD As String = ""
    Private _CodigoIndexador As Integer
    Private _Data As DateTime
    Private _Indice As Decimal
    Private _Realizado As Boolean
    Private _UsuarioLiberacao As String
    Private _UsuarioLiberacaoData As DateTime
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime
#End Region

#Region "Propriedades"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoIndexador() As Integer
        Get
            Return _CodigoIndexador
        End Get
        Set(ByVal value As Integer)
            _CodigoIndexador = value
        End Set
    End Property

    Public Property Data() As DateTime
        Get
            Return _Data
        End Get
        Set(ByVal value As DateTime)
            _Data = value
        End Set
    End Property

    Public Property Indice() As Decimal
        Get
            Return _Indice
        End Get
        Set(ByVal value As Decimal)
            _Indice = value
        End Set
    End Property

    Public Property Realizado() As Boolean
        Get
            Return _Realizado
        End Get
        Set(ByVal value As Boolean)
            _Realizado = value
        End Set
    End Property

    Public Property UsuarioLiberacao() As String
        Get
            Return _UsuarioLiberacao
        End Get
        Set(ByVal value As String)
            _UsuarioLiberacao = value
        End Set
    End Property

    Public Property UsuarioLiberacaoData() As DateTime
        Get
            Return _UsuarioLiberacaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioLiberacaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public ReadOnly Property Liberado() As Boolean
        Get
            If _UsuarioLiberacao Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

#End Region

#Region "Methods"

    Public Function chkLancamentos() As Boolean
        Dim sql As String = "Select IndiceFixado From Pedidos " & vbCrLf & _
                            "Where Indexador     = " & Me.CodigoIndexador & vbCrLf & _
                            "  And (IndiceFixado = " & Str(Me.Indice) & vbCrLf & _
                            "   Or DataPedido    = '" & Me.Data.ToString("yyyy-MM-dd") & "') " & vbCrLf & _
                            "" & vbCrLf & _
                            "Select Nota_Id From NotasFiscais " & vbCrLf & _
                            "Where Movimento = '" & Me.Data.ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                            "" & vbCrLf & _
                            "Select Registro_Id From ContasAPagar " & vbCrLf & _
                            "Where Prorrogacao = '" & Me.Data.ToString("yyyy-MM-dd") & "' " & vbCrLf & _
                            "" & vbCrLf & _
                            "Select Registro_Id From ContasAReceber " & vbCrLf & _
                            "Where Prorrogacao = '" & Me.Data.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                            "Select Titulo_Id From Titulos " & vbCrLf & _
                            "Where Reprogramacao = '" & Me.Data.ToString("yyyy-MM-dd") & "'"

        Dim ds As New DataSet
        ds = banco.ConsultaDataSet(sql, "CotacaoXItens")

        If Not ds Is Nothing AndAlso Not ds.Tables(0) Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then Return True
        If Not ds Is Nothing AndAlso Not ds.Tables(1) Is Nothing AndAlso ds.Tables(1).Rows.Count > 0 Then Return True
        If Not ds Is Nothing AndAlso Not ds.Tables(2) Is Nothing AndAlso ds.Tables(2).Rows.Count > 0 Then Return True
        If Not ds Is Nothing AndAlso Not ds.Tables(3) Is Nothing AndAlso ds.Tables(3).Rows.Count > 0 Then Return True
        If Not ds Is Nothing AndAlso Not ds.Tables(4) Is Nothing AndAlso ds.Tables(4).Rows.Count > 0 Then Return True

        Return False
    End Function

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case IUD
            Case "I"
                sql = "Update Cotacoes set " & vbCrLf & _
                      "    Indice               = " & Str(_Indice) & vbCrLf & _
                      "   ,Realizado            ='" & IIf(_Realizado, "R", "P") & "'" & vbCrLf & _
                      "   ,UsuarioLiberacao     ='" & _UsuarioLiberacao & "'" & vbCrLf & _
                      "   ,UsuarioLiberacaoData ='" & _UsuarioLiberacaoData.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      " Where Indexador_Id = " & _CodigoIndexador & vbCrLf & _
                      "   and Data_Id      ='" & _Data.ToString("yyyy-MM-dd") & "'"
                Sqls.Add(sql)
            Case "U"
                sql = "Update Cotacoes set " & vbCrLf & _
                      "    Indice               = " & Str(_Indice) & vbCrLf & _
                      "   ,Realizado            ='" & IIf(_Realizado, "R", "P") & "'" & vbCrLf & _
                      "   ,UsuarioAlteracao     ='" & _UsuarioAlteracao & "'" & vbCrLf & _
                      "   ,UsuarioAlteracaoData ='" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      " Where Indexador_Id = " & _CodigoIndexador & vbCrLf & _
                      "   and Data_Id      ='" & _Data.ToString("yyyy-MM-dd") & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
