Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListIngredienteAtivo
    Inherits List(Of IngredienteAtivo)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean)
        If CarregarDados Then
            Dim banco As New AcessaBanco
            Dim ds As New DataSet
            Dim sql As String = "Select IA_Id, EstadoFisicoIA, GrupoQuimico, NomeComumIA, NomeQuimicoIA, SolubilidadeIA, " & vbCrLf & _
                                "PesoMolecularIA, PontoFusaoIA, PressaoVaporIA FROM IA Order By IA_Id "

            ds = banco.ConsultaDataSet(sql, "IngredienteAtivo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim ia As New IngredienteAtivo


            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class IngredienteAtivo

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim banco As New AcessaBanco
        Dim ds As New DataSet
        Dim sql As String = "Select IA_Id, EstadoFisicoIA, GrupoQuimico, NomeComumIA, NomeQuimicoIA, SolubilidadeIA, " & vbCrLf & _
                            "PesoMolecularIA, PontoFusaoIA, PressaoVaporIA FROM IA " & vbCrLf & _
                            "WHERE IA_Id = " & pCodigo & " Order By IA_Id "
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _CodigoEstadoFisicoIA As Integer
    Private _EstadoFisicoIA As EstadoFisicoIA
    Private _CodigoGrupoQuimico As Integer
    Private _GrupoQuimico As GrupoQuimico
    Private _NomeComumIA As String
    Private _NomeQuimicoIA As String
    Private _SolubilidadeIA As String
    Private _PesoMolecularIA As String
    Private _PontoFusaoIA As String
    Private _PressaoVaporIA As String
#End Region

#Region "Property"
    Public Property UID() As String
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

    Public Property CodigoEstadoFisicoIA() As Integer
        Get
            Return _CodigoEstadoFisicoIA
        End Get
        Set(ByVal value As Integer)
            _CodigoEstadoFisicoIA = value
        End Set
    End Property

    Public ReadOnly Property EstadoFisicoIA() As EstadoFisicoIA
        Get
            If _CodigoEstadoFisicoIA > 0 Then _EstadoFisicoIA = New EstadoFisicoIA(_CodigoEstadoFisicoIA)
            Return _EstadoFisicoIA
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

    Public ReadOnly Property GrupoQuimico() As GrupoQuimico
        Get
            If _CodigoGrupoQuimico > 0 Then _GrupoQuimico = New GrupoQuimico(_CodigoGrupoQuimico)
            Return _GrupoQuimico
        End Get
    End Property

    Public Property NomeComumIA() As String
        Get
            Return _NomeComumIA
        End Get
        Set(ByVal value As String)
            _NomeComumIA = value
        End Set
    End Property

    Public Property NomeQuimicoIA() As String
        Get
            Return _NomeQuimicoIA
        End Get
        Set(ByVal value As String)
            _NomeQuimicoIA = value
        End Set
    End Property

    Public Property SolubilidadeIA() As String
        Get
            Return _SolubilidadeIA
        End Get
        Set(ByVal value As String)
            _SolubilidadeIA = value
        End Set
    End Property

    Public Property PesoMolecularIA() As String
        Get
            Return _PesoMolecularIA
        End Get
        Set(ByVal value As String)
            _PesoMolecularIA = value
        End Set
    End Property

    Public Property PontoFusaoIA() As String
        Get
            Return _PontoFusaoIA
        End Get
        Set(ByVal value As String)
            _PontoFusaoIA = value
        End Set
    End Property

    Public Property PressaoVaporIA() As String
        Get
            Return _PressaoVaporIA
        End Get
        Set(ByVal value As String)
            _PressaoVaporIA = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If UID Is Nothing Then

        End If
        Return True

    End Function
#End Region

End Class