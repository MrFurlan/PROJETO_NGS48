Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListDosagem
    Inherits List(Of Dosagem)

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCulturaPragaFito As Integer, Optional ByVal pCodigoFito As Integer = 0, Optional ByVal pCodigoCultura As Integer = 0, Optional ByVal pcodigoPraga As Integer = 0)
        If pCodigoCulturaPragaFito = 0 And (pCodigoFito = 0 Or pCodigoCultura = 0 Or pcodigoPraga = 0) Then Exit Sub

        Dim sql As String
        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        sql = "SELECT CPFD.CulturaPragaFito_Id, CPFD.Dosagem_Id, CPFD.Solo, CPFD.DosagemMinima, CPFD.DosagemMaxima, CPFD.UnidadeDeMedida, isnull(CPFD.VazaoTerrestre,'') as VazaoTerrestre," & vbCrLf & _
              "       isnull(CPFD.VazaoAerea,'') as VazaoAerea, isnull(CPFD.IntervaloDeSeguranca,'') as IntervaloDeSeguranca,  CPFD.DosagemRecomendada" & vbCrLf & _
              "  FROM CulturaxPragaxFitoxDosagem CPFD" & vbCrLf & _
              " inner join CulturaxPragaxFito CPF" & vbCrLf & _
              "    on CPFD.CulturaPragaFito_Id = CPF.CulturaPragaFito_Id" & vbCrLf
        If pCodigoCulturaPragaFito > 0 Then
            sql &= " Where CPFD.CulturaPragaFito_Id = " & pCodigoCulturaPragaFito
        Else
            sql &= " Where CPF.Fito    = " & pCodigoFito & vbCrLf & _
                   "   and CPF.Cultura = " & pCodigoCultura & vbCrLf & _
                   "   and CPF.Praga   = " & pcodigoPraga
        End If

        ds = Banco.ConsultaDataSet(sql, "Dosagem")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Dos As New Dosagem
            Dos.CodigoCulturaPragaFito = row("CulturaPragaFito_Id")
            Dos.CodigoDosagem = row("Dosagem_Id")
            Dos.CodigoSolo = row("Solo")
            Dos.DosagemMinima = row("DosagemMinima")
            Dos.DosagemMaxima = row("DosagemMaxima")
            Dos.UnidadeDeMedida = row("UnidadeDeMedida")
            Dos.VazaoTerrestre = row("VazaoTerrestre")
            Dos.VazaoAerea = row("VazaoAerea")
            Dos.IntervaloDeSeguranca = row("IntervaloDeSeguranca")
            Dos.DosagemRecomendada = row("DosagemRecomendada")
            Me.Add(Dos)
        Next

    End Sub
#End Region

End Class

<Serializable()> _
Public Class Dosagem

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCulturaPragaFito As String, ByVal pDosagem As String)
        Dim sql As String
        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        sql = "SELECT     CulturaPragaFito_Id, Dosagem_Id, Solo, DosagemMinima, DosagemMaxima, UnidadeDeMedida, VazaoTerrestre, " & vbCrLf & _
              "           VazaoAerea, IntervaloDeSeguranca, DosagemRecomendada " & vbCrLf & _
              "  FROM     CulturaxPragaxFitoxDosagem " & vbCrLf & _
              "  WHERE    CulturaPragaFito_Id = '" & pCulturaPragaFito & "' AND " & vbCrLf & _
              "           Dosagem_Id = " & pDosagem & ""

        ds = Banco.ConsultaDataSet(sql, "Dosagem")

        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoCulturaPragaFito = ds.Tables(0).Rows(0).Item("CulturaPragaFito_Id")
            _CodigoDosagem = ds.Tables(0).Rows(0).Item("Dosagem_Id")
            _CodigoSolo = ds.Tables(0).Rows(0).Item("Solo")
            _DosagemMinima = ds.Tables(0).Rows(0).Item("DosagemMinima")
            _DosagemMaxima = ds.Tables(0).Rows(0).Item("DosagemMaxima")
            _UnidadeDeMedida = ds.Tables(0).Rows(0).Item("UnidadeDeMedida")
            _VazaoTerrestre = ds.Tables(0).Rows(0).Item("VazaoTerrestre")
            _VazaoAerea = ds.Tables(0).Rows(0).Item("VazaoAerea")
            _IntervaloDeSeguranca = ds.Tables(0).Rows(0).Item("IntervaloDeSeguranca")
            _DosagemRecomendada = ds.Tables(0).Rows(0).Item("DosagemRecomendada")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoCulturaPragaFito As Integer
    Private _CodigoDosagem As Integer
    Private _Solo As Solo
    Private _CodigoSolo As Integer
    Private _DosagemMinima As Decimal
    Private _DosagemMaxima As Decimal
    Private _UnidadeDeMedida As String
    Private _VazaoTerrestre As String
    Private _VazaoAerea As String
    Private _IntervaloDeSeguranca As String
    Private _DosagemRecomendada As Decimal
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

    Public Property CodigoCulturaPragaFito() As Integer
        Get
            Return _CodigoCulturaPragaFito
        End Get
        Set(ByVal value As Integer)
            _CodigoCulturaPragaFito = value
        End Set
    End Property

    Public Property CodigoDosagem() As Integer
        Get
            Return _CodigoDosagem
        End Get
        Set(ByVal value As Integer)
            _CodigoDosagem = value
        End Set
    End Property

    Public ReadOnly Property Solo() As Solo
        Get
            If _Solo Is Nothing And _CodigoSolo > 0 Then _Solo = New Solo(_CodigoSolo)
            Return _Solo
        End Get
    End Property

    Public ReadOnly Property NomeSolo() As String
        Get
            If Solo Is Nothing Then
                Return ""
            Else
                Return Solo.Descricao
            End If
        End Get
    End Property

    Public Property CodigoSolo() As Integer
        Get
            Return _CodigoSolo
        End Get
        Set(ByVal value As Integer)
            _CodigoSolo = value
        End Set
    End Property

    Public Property DosagemMinima() As Decimal
        Get
            Return _DosagemMinima
        End Get
        Set(ByVal value As Decimal)
            _DosagemMinima = value
        End Set
    End Property

    Public Property DosagemMaxima() As Decimal
        Get
            Return _DosagemMaxima
        End Get
        Set(ByVal value As Decimal)
            _DosagemMaxima = value
        End Set
    End Property

    Public Property UnidadeDeMedida() As String
        Get
            Return _UnidadeDeMedida
        End Get
        Set(ByVal value As String)
            _UnidadeDeMedida = value
        End Set
    End Property

    Public Property VazaoTerrestre() As String
        Get
            Return _VazaoTerrestre
        End Get
        Set(ByVal value As String)
            _VazaoTerrestre = value
        End Set
    End Property

    Public Property VazaoAerea() As String
        Get
            Return _VazaoAerea
        End Get
        Set(ByVal value As String)
            _VazaoAerea = value
        End Set
    End Property

    Public Property IntervaloDeSeguranca() As String
        Get
            Return _IntervaloDeSeguranca
        End Get
        Set(ByVal value As String)
            _IntervaloDeSeguranca = value
        End Set
    End Property

    Public Property DosagemRecomendada() As Decimal
        Get
            Return _DosagemRecomendada
        End Get
        Set(ByVal value As Decimal)
            _DosagemRecomendada = value
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
                Sql = " INSERT INTO CulturaxPragaxFitoxDosagem (CulturaPragaFito_Id, Dosagem_Id, Solo, DosagemMinima, DosagemMaxima, UnidadeDeMedida, VazaoTerrestre, VazaoAerea, IntervaloDeSeguranca, DosagemRecomendada)" & vbCrLf & _
                      " VALUES (" & _CodigoCulturaPragaFito & "," & _CodigoDosagem & "," & _CodigoSolo & "," & Str(_DosagemMinima) & "," & Str(_DosagemMaxima) & ",'" & _UnidadeDeMedida & "','" & _VazaoTerrestre & "','" & _VazaoAerea & "','" & _IntervaloDeSeguranca & "'," & Str(_DosagemRecomendada) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE CulturaxPragaxFitoxDosagem SET" & vbCrLf & _
                      "   CulturaPragaFito_Id  = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "  ,Dosagem_Id           = " & _CodigoDosagem & vbCrLf & _
                      "  ,Solo                 = " & _CodigoSolo & vbCrLf & _
                      "  ,DosagemMinima        = " & Str(_DosagemMinima) & vbCrLf & _
                      "  ,DosagemMaxima        = " & Str(_DosagemMaxima) & vbCrLf & _
                      "  ,UnidadeDeMedida      ='" & _UnidadeDeMedida & "'" & vbCrLf & _
                      "  ,VazaoTerrestre       ='" & _VazaoTerrestre & "'" & vbCrLf & _
                      "  ,VazaoAerea           ='" & _VazaoAerea & "'" & vbCrLf & _
                      "  ,IntervaloDeSeguranca ='" & _IntervaloDeSeguranca & "'" & vbCrLf & _
                      "  ,DosagemRecomendada   = " & Str(_DosagemRecomendada) & "'" & vbCrLf & _
                      "  WHERE CulturaPragaFito_Id = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "    and Dosagem_Id          = " & _CodigoDosagem
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE CulturaxPragaxFitoxDosagem" & vbCrLf & _
                      "  WHERE CulturaPragaFito_Id = " & _CodigoCulturaPragaFito & vbCrLf & _
                      "    and Dosagem_Id          = " & _CodigoDosagem
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class