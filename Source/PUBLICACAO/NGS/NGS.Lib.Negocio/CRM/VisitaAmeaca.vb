Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListVisitaAmeaca
    Inherits List(Of VisitaAmeaca)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pNumeroVisita As Integer)
        Dim Sql As String = ""
        Sql = "SELECT NumeroVisita_Id NumeroVisita, Ameaca_Id as Ameaca, EndAmeaca_Id  as EndAmeaca," & vbCrLf & _
              "		  LTRIM(VM.Observacao) as Observacao" & vbCrLf & _
              "  from VisitaAmeaca VM" & vbCrLf & _
              " WHERE VM.NumeroVisita_Id    = " & pNumeroVisita


        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(Sql, "VisitaAmeaca")

        For Each row In ds.Tables(0).Rows
            Dim vis As New VisitaAmeaca(pNumeroVisita)
            'vis.IUD = row("IUD")
            vis.NumeroVisita = row("NumeroVisita")
            vis.CodigoAmeaca = row("Ameaca")
            vis.EndAmeaca = row("EndAmeaca")
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
        For Each item As VisitaAmeaca In Me
            item.SalvarSql(Sqls)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class VisitaAmeaca
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

    Private _CodigoAmeaca As String = ""
    Private _EndAmeaca As Integer
    Private _Ameaca As Cliente

    Private _Observacao As String = ""

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

    Public Property NumeroVisita As Integer
        Get
            Return _NumeroVisita
        End Get
        Set(ByVal value As Integer)
            _NumeroVisita = value
        End Set
    End Property

    Public Property CodigoAmeaca As String
        Get
            Return _CodigoAmeaca
        End Get
        Set(ByVal value As String)
            _CodigoAmeaca = value
        End Set
    End Property

    Public Property EndAmeaca As Integer
        Get
            Return _EndAmeaca
        End Get
        Set(ByVal value As Integer)
            _EndAmeaca = value
        End Set
    End Property



    Public Property Ameaca() As Cliente
        Get
            If _Ameaca Is Nothing And Me.CodigoAmeaca.Length > 0 Then
                _Ameaca = New Cliente(Me.CodigoAmeaca, Me.EndAmeaca)
            End If
            Return _Ameaca
        End Get
        Set(ByVal value As Cliente)
            _Ameaca = value
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
                strSQL = "insert into VisitaAmeaca (NumeroVisita_Id, Ameaca_Id, EndAmeaca_Id, Observacao)" & vbCrLf & _
                         " values(" & Me.NumeroVisita & ",'" & Me.CodigoAmeaca & "', " & Me.EndAmeaca & ", '" & Me.Observacao & "')"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = "update VisitaAmeaca " & vbCrLf & _
                         "   set Observacao      = '" & Me.Observacao & "' " & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                         "   and Ameaca_Id         = '" & Me.CodigoAmeaca & "' " & vbCrLf & _
                         "   and EndAmeaca_Id       = " & Me.EndAmeaca
                Sqls.Add(strSQL)
            Case "D"
                strSQL = "delete from  VisitaAmeaca " & vbCrLf & _
                        " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                        "   and Ameaca_Id         = '" & Me.CodigoAmeaca & "' " & vbCrLf & _
                        "   and EndAmeaca_Id       = " & Me.EndAmeaca
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
