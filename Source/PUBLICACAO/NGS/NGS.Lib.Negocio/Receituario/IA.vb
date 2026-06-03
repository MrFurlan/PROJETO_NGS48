Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListIA
    Inherits List(Of IA)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False, Optional ByVal OrderBy As String = "")
        If CarregarDados Then
            Dim sql As String = "Select IA_Id, EstadoFisicoIA, GrupoQuimico, NomeComum, NomeQuimico, Solubilidade, PesoMolecular, PontoFusao, PressaoVapor From IA "
            If OrderBy.Length > 0 Then
                sql &= "Order By " & OrderBy
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "IA")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Ing As New IA
                Ing.CodigoIA = row("IA_Id")
                Ing.CodigoEstadoFisico = row("EstadoFisicoIA")
                Ing.CodigoGrupoQuimico = row("GrupoQuimico")
                Ing.NomeComum = row("NomeComum")
                Ing.NomeQuimico = row("NomeQuimico")
                Ing.Solubilidade = row("Solubilidade")
                Ing.PesoMolecular = row("PesoMolecular")
                Ing.PontoFusao = row("PontoFusao")
                Ing.PressaoVapor = row("PressaoVapor")
                Me.Add(Ing)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class IA

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoIA As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String

        Sql = "Select IA_Id, EstadoFisicoIA, GrupoQuimico, NomeComum, NomeQuimico, Solubilidade, PesoMolecular, PontoFusao, PressaoVapor From IA Where IA_id = " & pCodigoIA

        ds = Banco.ConsultaDataSet(Sql, "IA")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoIA = row("IA_Id")
            _CodigoEstadoFisico = row("EstadoFisicoIA")
            _CodigoGrupoQuimico = row("GrupoQuimico")
            _NomeComum = row("NomeComum")
            _NomeQuimico = row("NomeQuimico")
            _Solubilidade = row("Solubilidade")
            _PesoMolecular = row("PesoMolecular")
            _PontoFusao = row("PontoFusao")
            _PressaoVapor = row("PressaoVapor")
        End If
    End Sub

    Sub Numerador()
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim Sql As String

        Sql = "Select isnull(max(IA_Id),0) + 1 AS Numerador From IA"

        ds = Banco.ConsultaDataSet(Sql, "IA")

        If ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _CodigoIA = row("Numerador")
        End If
    End Sub
#End Region

#Region "Fieds"
    Private _IUD As String
    Private _CodigoIA As Integer
    Private _EstadoFisico As EstadoFisicoIA
    Private _CodigoEstadoFisico As Integer
    Private _GrupoQuimico As GrupoQuimico
    Private _CodigoGrupoQuimico As Integer
    Private _NomeComum As String
    Private _NomeQuimico As String
    Private _Solubilidade As String
    Private _PesoMolecular As String
    Private _PontoFusao As String
    Private _PressaoVapor As String
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

    Public Property CodigoIA() As Integer
        Get
            Return _CodigoIA
        End Get
        Set(ByVal value As Integer)
            _CodigoIA = value
        End Set
    End Property

    Public ReadOnly Property EstadoFisico() As EstadoFisicoIA
        Get
            If _EstadoFisico Is Nothing And _CodigoEstadoFisico > 0 Then _EstadoFisico = New EstadoFisicoIA(_CodigoEstadoFisico)
            Return _EstadoFisico
        End Get
    End Property

    Public Property CodigoEstadoFisico() As Integer
        Get
            Return _CodigoEstadoFisico
        End Get
        Set(ByVal value As Integer)
            _CodigoEstadoFisico = value
            _EstadoFisico = Nothing
        End Set
    End Property

    Public ReadOnly Property GrupoQuimico() As GrupoQuimico
        Get
            If _GrupoQuimico Is Nothing And _CodigoGrupoQuimico > 0 Then _GrupoQuimico = New GrupoQuimico(_CodigoGrupoQuimico)
            Return _GrupoQuimico
        End Get
    End Property

    Public Property CodigoGrupoQuimico() As Integer
        Get
            Return _CodigoGrupoQuimico
        End Get
        Set(ByVal value As Integer)
            _CodigoGrupoQuimico = value
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

    Public Property NomeQuimico() As String
        Get
            Return _NomeQuimico
        End Get
        Set(ByVal value As String)
            _NomeQuimico = value
        End Set
    End Property

    Public Property Solubilidade() As String
        Get
            Return _Solubilidade
        End Get
        Set(ByVal value As String)
            _Solubilidade = value
        End Set
    End Property

    Public Property PesoMolecular() As String
        Get
            Return _PesoMolecular
        End Get
        Set(ByVal value As String)
            _PesoMolecular = value
        End Set
    End Property

    Public Property PontoFusao() As String
        Get
            Return _PontoFusao
        End Get
        Set(ByVal value As String)
            _PontoFusao = value
        End Set
    End Property

    Public Property PressaoVapor() As String
        Get
            Return _PressaoVapor
        End Get
        Set(ByVal value As String)
            _PressaoVapor = value
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
                Sql = " INSERT INTO IA(IA_Id, EstadoFisicoIA, GrupoQuimico, NomeComum, NomeQuimico, Solubilidade, PesoMolecular, PontoFusao, PressaoVapor) " & vbCrLf & _
                      " VALUES (" & _CodigoIA & "," & _CodigoEstadoFisico & "," & _CodigoGrupoQuimico & ",'" & _NomeComum & "','" & _NomeQuimico & "','" & _Solubilidade & "','" & _PesoMolecular & "','" & _PontoFusao & "','" & PressaoVapor & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE IA SET" & vbCrLf & _
                      "   EstadoFisicoIA = " & _CodigoEstadoFisico & vbCrLf & _
                      "  ,GrupoQuimico   = " & _CodigoGrupoQuimico & vbCrLf & _
                      "  ,NomeComum      ='" & _NomeComum & "'" & vbCrLf & _
                      "  ,NomeQuimico    ='" & _NomeQuimico & "'" & vbCrLf & _
                      "  ,Solubilidade   ='" & _Solubilidade & "'" & vbCrLf & _
                      "  ,PesoMolecular  ='" & _PesoMolecular & "'" & vbCrLf & _
                      "  ,PontoFusao     ='" & _PontoFusao & "'" & vbCrLf & _
                      "  ,PressaoVapor   ='" & _PressaoVapor & "'" & vbCrLf & _
                      "  WHERE IA_Id     = " & _CodigoIA
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE IA" & vbCrLf & _
                      "  WHERE IA_Id =" & _CodigoIA
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class