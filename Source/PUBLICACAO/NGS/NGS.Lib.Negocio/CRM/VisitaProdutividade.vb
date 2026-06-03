Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListVisitaProdutividade
    Inherits List(Of VisitaProdutividade)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pNumeroVisita As Integer)
        Dim Sql As String = ""
        Sql = "SELECT NumeroVisita_Id NumeroVisita, Safra_Id as CodigoSafra, " & vbCrLf & _
              "		  Cultura_Id CodigoCultura, Produtividade" & vbCrLf & _
              "  from VisitaProdutividade VM" & vbCrLf & _
              " WHERE VM.NumeroVisita_Id    = " & pNumeroVisita


        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(Sql, "VisitaProdutividade")

        For Each row In ds.Tables(0).Rows
            Dim vis As New VisitaProdutividade(pNumeroVisita)
            'vis.IUD = row("IUD")
            vis.NumeroVisita = row("NumeroVisita")
            vis.CodigoSafra = row("CodigoSafra")
            vis.CodigoCultura = row("CodigoCultura")
            vis.Produtividade = row("Produtividade")
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
        For Each item As VisitaProdutividade In Me
            item.SalvarSql(Sqls)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class VisitaProdutividade
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

    Private _CodigoSafra As String

    Private _CodigoCultura As Integer
    Private _Cultura As Cultura

    Private _Produtividade As Decimal


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

    Public Property CodigoSafra As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
        End Set
    End Property
    Public Property CodigoCultura As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
        End Set
    End Property


    Public Property Produtividade As Decimal
        Get
            Return _Produtividade
        End Get
        Set(ByVal value As Decimal)
            _Produtividade = value
        End Set
    End Property


    Public Property Cultura As Cultura
        Get
            If _Cultura Is Nothing And Me.CodigoCultura > 0 Then
                _Cultura = New Cultura(CodigoCultura)
            End If
            Return _Cultura
        End Get
        Set(ByVal value As Cultura)
            _Cultura = value
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
                strSQL = "insert into VisitaProdutividade (NumeroVisita_Id, Safra_Id, Cultura_Id, Produtividade)" & vbCrLf & _
                         " values(" & Me.NumeroVisita & ",'" & Me.CodigoSafra & "', " & Me.CodigoCultura & "," & Str(Me.Produtividade) & ")"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = "update VisitaProdutividade " & vbCrLf & _
                         "   set Produtividade   = " & Str(Me.Produtividade) & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                         "   and Safra_Id        ='" & Me.CodigoSafra & "'" & vbCrLf & _
                         "   and Cultura_Id      = " & Me.CodigoCultura
                Sqls.Add(strSQL)
            Case "D"
                strSQL = "delete from VisitaProdutividade " & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita & vbCrLf & _
                         "   and Safra_Id        ='" & Me.CodigoSafra & "'" & vbCrLf & _
                         "   and Cultura_Id      = " & Me.CodigoCultura
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
