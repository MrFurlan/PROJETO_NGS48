Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCarteiraDoTitulo
    Inherits List(Of CarteiraDoTitulo)

#Region "Contrutor"
    Public Sub New()
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT C.Carteira_Id AS Codigo, C.Descricao, isnull(C.Banco,0) AS Banco, isnull(C.FluxoDeCaixa,0) AS FluxoDeCaixa, isnull(C.EmiteDuplicata,0) as EmiteDuplicata, " & vbCrLf & _
              "case isnull(C.Banco,0) when 0 " & vbCrLf & _
              "     then '' " & vbCrLf & _
              "   else " & vbCrLf & _
              "     (convert(varchar,C.Banco) + '-' + B.Descricao) " & vbCrLf & _
              "   end AS DescricaoBanco " & vbCrLf & _
              "  FROM Carteira AS C " & vbCrLf & _
              "       LEFT JOIN Bancos AS B" & vbCrLf & _
              "              ON B.Banco_id = C.Banco " & vbCrLf & _
              " ORDER BY C.Carteira_Id "

        ds = Banco.ConsultaDataSet(sql, "Carteira")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim C As New CarteiraDoTitulo
            C.Codigo = row("Codigo")
            C.Descricao = row("Descricao")
            C.CodigoDoBanco = row("Banco")
            C.FluxoDeCaixa = row("FluxoDeCaixa")
            C.EmiteDuplicata = row("EmiteDuplicata")
            C.DescricaoBanco = IIf(IsDBNull(row("DescricaoBanco")), "", row("DescricaoBanco"))
            Me.Add(C)
        Next
    End Sub
#End Region

    Private Sub car(obj As Object)

    End Sub

End Class

<Serializable()> _
Public Class CarteiraDoTitulo
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT C.Carteira_Id , C.Descricao, isnull(C.Banco,0) AS Banco, isnull(C.FluxoDeCaixa,0) AS FluxoDeCaixa, isnull(C.EmiteDuplicata,0) as EmiteDuplicata, " & vbCrLf & _
              "       case isnull(C.Banco,0)" & vbCrLf & _
              "         when 0 " & vbCrLf & _
              "           then '' " & vbCrLf & _
              "           else (convert(varchar,C.Banco) + '-' + B.Descricao) " & vbCrLf & _
              "       end AS DescricaoBanco " & vbCrLf & _
              "  FROM Carteira AS C " & vbCrLf & _
              "  LEFT JOIN Bancos AS B" & vbCrLf & _
              "    ON B.Banco_id  = C.Banco " & vbCrLf & _
              " Where Carteira_Id = " & Codigo

        ds = Banco.ConsultaDataSet(sql, "CarteiraDoTitulo")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = row("Carteira_Id")
        _Descricao = row("Descricao")
        _CodigoDoBanco = row("Banco")
        _FluxoDeCaixa = row("FluxoDeCaixa")
        _EmiteDuplicata = row("EmiteDuplicata")
        _DescricaoBanco = row("DescricaoBanco")
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _CodigoDoBanco As Integer
    Private _Banco As Banco
    Private _DescricaoBanco As String
    Private _FluxoDeCaixa As Boolean
    Private _EmiteDuplicata As Boolean
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

    Public Property CodigoDoBanco() As Integer
        Get
            Return _CodigoDoBanco
        End Get
        Set(ByVal value As Integer)
            _CodigoDoBanco = value
        End Set
    End Property

    Public Property Banco() As Banco
        Get
            If _Banco Is Nothing And _CodigoDoBanco > 0 Then _Banco = New Banco(_CodigoDoBanco)
            Return _Banco
        End Get
        Set(ByVal value As Banco)
            _Banco = value
        End Set
    End Property

    Public Property DescricaoBanco() As String
        Get
            Return _DescricaoBanco
        End Get
        Set(ByVal value As String)
            _DescricaoBanco = value
        End Set
    End Property

    Public Property FluxoDeCaixa() As Boolean
        Get
            Return _FluxoDeCaixa
        End Get
        Set(ByVal value As Boolean)
            _FluxoDeCaixa = value
        End Set
    End Property

    Public Property EmiteDuplicata As Boolean
        Get
            Return _EmiteDuplicata
        End Get
        Set(ByVal value As Boolean)
            _EmiteDuplicata = value
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
                Sql = " INSERT INTO Carteira(Carteira_Id, Descricao, Banco, FluxoDeCaixa, EmiteDuplicata) " & vbCrLf & _
                      " VALUES (" & Me.Codigo & ",'" & Me.Descricao & "'," & Me.CodigoDoBanco & "," & IIf(Me.FluxoDeCaixa, 1, 0) & "," & IIf(Me.EmiteDuplicata, 1, 0) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update Carteira set" & vbCrLf & _
                      "  Descricao      ='" & Me.Descricao & "'" & vbCrLf & _
                      " ,Banco          = " & Me.CodigoDoBanco & vbCrLf & _
                      " ,FluxoDeCaixa   = " & IIf(Me.FluxoDeCaixa, 1, 0) & "" & vbCrLf & _
                      " ,EmiteDuplicata = " & IIf(Me.EmiteDuplicata, 1, 0) & "" & vbCrLf & _
                      " Where Carteira_Id = " & Me.Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Carteira" & vbCrLf & _
                      "  WHERE Carteira_Id      ='" & _Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class