Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListVisitaMotivo
    Inherits List(Of VisitaMotivo)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pNumeroVisita As Integer)
        Dim Sql As String = ""
        Sql = "SELECT NumeroVisita_Id NumeroVisita, Motivo_Id as Motivo," & vbCrLf & _
              "		  LTRIM(VM.Observacao) as Observacao" & vbCrLf & _
              "  from VisitaMotivo VM" & vbCrLf & _
              " WHERE VM.NumeroVisita_Id    = " & pNumeroVisita


        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(Sql, "Visitas")

        For Each row In ds.Tables(0).Rows
            Dim vis As New VisitaMotivo(pNumeroVisita)
            'vis.IUD = row("IUD")
            vis.NumeroVisita = row("NumeroVisita")
            vis.CodigoMotivo = row("Motivo")
            vis.Observacao = row("Observacao")
            Me.Add(vis)
        Next
    End Sub
#Region "Methods"
    Public Function Salvar() As Boolean
        Dim sqls As New ArrayList
        SalvarSql(sqls)

        Dim Banco As New AcessaBanco
        If Banco.GravaBanco(sqls) Then
            For Each row In Me
                row.IUD = ""
            Next
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As VisitaMotivo In Me
            item.SalvarSql(Sqls)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class VisitaMotivo
    Implements IBaseEntity

#Region "Construtor"


    Public Sub New(ByVal pNumeroVisita As Integer)
        Me.NumeroVisita = pNumeroVisita
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""

    '**** Ligação com o CRMXVisita *****************'
    Private _NumeroVisita As Integer

    Private _CodigoMotivo As Integer

    Private _Observacao As String = ""

    Private _Motivo As Motivo

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

    Public Property NumeroVisita()
        Get
            Return _NumeroVisita
        End Get
        Set(ByVal value)
            _NumeroVisita = value
        End Set
    End Property

    Public Property CodigoMotivo As Integer
        Get
            Return _CodigoMotivo
        End Get
        Set(ByVal value As Integer)
            _CodigoMotivo = value
        End Set
    End Property
    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property

    Public Property Motivo As Motivo
        Get
            If _Motivo Is Nothing And Me.CodigoMotivo > 0 Then
                _Motivo = New Motivo(CodigoMotivo)
            End If
            Return _Motivo
        End Get
        Set(ByVal value As Motivo)
            _Motivo = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD Is Nothing Then Return True
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
                strSQL = "insert into VisitaMotivo (NumeroVisita_Id, Motivo_Id, Observacao)" & vbCrLf & _
                         " values(" & Me.NumeroVisita & "," & Me.CodigoMotivo & ", '" & Me.Observacao & "')"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = "update VisitaMotivo " & vbCrLf & _
                         "   set Observacao      = '" & Me.Observacao & "' " & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                         "   and Motivo_Id       = " & Me.CodigoMotivo
                Sqls.Add(strSQL)
            Case "D"
                strSQL = "delete from VisitaMotivo " & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                         "   and Motivo_Id       = " & Me.CodigoMotivo
                Sqls.Add(strSQL)
        End Select
    End Sub

#End Region

End Class
