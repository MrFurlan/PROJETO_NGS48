Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListPraga
    Inherits List(Of Praga)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal OrderBY As String = "")
        If CarregarDados Then
            Dim Banco As New AcessaBanco
            Dim ds As DataSet
            Dim sql As String = "Select Praga_Id, Categoria, isnull(NomeComum,'') as NomeComum, isnull(NomeCientifico,'') as NomeCientifico, isnull(Familia,'') as Familia, isnull(DadosAdicionais,'') as DadosAdicionais, isnull(CodigoIndea,'') as CodigoIndea" & vbCrLf & _
                                "  from Praga"
            If OrderBY.Length > 0 Then sql &= " Order By " & OrderBY

            ds = Banco.ConsultaDataSet(sql, "Praga")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Pr As New Praga

                Pr.Codigo = row("Praga_Id")
                Pr.NomeComum = row("NomeComum")
                Pr.NomeCientifico = row("NomeCientifico")
                Pr.CodigoCategoria = row("Categoria")
                Pr.FamiliaPraga = row("Familia")
                Pr.DadosAdicionais = row("DadosAdicionais")
                Pr.CodigoIndea = row("CodigoIndea")

                Me.Add(Pr)
            Next

        End If
    End Sub

End Class

<Serializable()> _
Public Class Praga

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Praga_Id, Categoria, isnull(NomeComum,'') as NomeComum, isnull(NomeCientifico,'') as NomeCientifico, isnull(Familia,'') as Familia, isnull(DadosAdicionais,'') as DadosAdicionais, isnull(CodigoIndea,'') as CodigoIndea" & vbCrLf & _
                            "  from Praga where Praga_Id =" & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Praga")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Praga_Id")
            _NomeComum = ds.Tables(0).Rows(0)("NomeComum")
            _NomeCientifico = ds.Tables(0).Rows(0)("NomeCientifico")
            _CodigoCategoria = ds.Tables(0).Rows(0)("Categoria")
            _FamiliaPraga = ds.Tables(0).Rows(0)("Familia")
            _DadosAdicionais = ds.Tables(0).Rows(0)("DadosAdicionais")
            _CodigoIndea = ds.Tables(0).Rows(0)("CodigoIndea")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Categoria As CategoriaPraga
    Private _CodigoCategoria As Integer
    Private _NomeComum As String
    Private _NomeCientifico As String
    Private _FamiliaPraga As String
    Private _DadosAdicionais As String
    Private _CodigoIndea As String
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

    Public Property NomeComum() As String
        Get
            Return _NomeComum
        End Get
        Set(ByVal value As String)
            _NomeComum = value
        End Set
    End Property

    Public Property NomeCientifico() As String
        Get
            Return _NomeCientifico
        End Get
        Set(ByVal value As String)
            _NomeCientifico = value
        End Set
    End Property

    Public Property CodigoCategoria() As Integer
        Get
            Return _CodigoCategoria
        End Get
        Set(ByVal value As Integer)
            _CodigoCategoria = value
        End Set
    End Property

    Public ReadOnly Property Categoria() As CategoriaPraga
        Get
            If _CodigoCategoria > 0 Then _Categoria = New CategoriaPraga(_CodigoCategoria)
            Return _Categoria
        End Get
    End Property

    Public Property FamiliaPraga() As String
        Get
            Return _FamiliaPraga
        End Get
        Set(ByVal value As String)
            _FamiliaPraga = value
        End Set
    End Property

    Public Property DadosAdicionais() As String
        Get
            Return _DadosAdicionais
        End Get
        Set(ByVal value As String)
            _DadosAdicionais = value
        End Set
    End Property

    Public Property CodigoIndea() As String
        Get
            Return _CodigoIndea
        End Get
        Set(ByVal value As String)
            _CodigoIndea = value
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
                Sql = " INSERT INTO Praga (Praga_Id, CategoriaPraga, NomeComumPraga, NomeCientificoPraga, FamiliaPraga, DadosAdicionaisPraga, CodigoIndeaPraga) " & vbCrLf & _
                      " VALUES (" & _Codigo & "," & _CodigoCategoria & ",'" & _NomeComum & "','" & _NomeCientifico & "','" & FamiliaPraga & "','" & DadosAdicionais & "','" & _CodigoIndea & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Praga SET" & vbCrLf & _
                      "    NomeComumPraga       = " & _NomeComum & "'" & vbCrLf & _
                      "   ,NomeCientificoPraga  ='" & _NomeCientifico & "'" & vbCrLf & _
                      "   ,CategoriaPraga       = " & _CodigoCategoria & vbCrLf & _
                      "   ,FamiliaPraga         ='" & _FamiliaPraga & "'" & vbCrLf & _
                      "   ,DadosAdicionaisPraga ='" & _CodigoIndea & "'" & vbCrLf & _
                      "  WHERE Praga_Id ='" & _Codigo & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Praga" & vbCrLf & _
                      "  WHERE Praga_Id     ='" & _Codigo & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
