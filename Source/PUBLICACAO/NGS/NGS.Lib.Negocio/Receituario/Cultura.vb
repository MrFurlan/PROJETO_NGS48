Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCultura
    Inherits List(Of Cultura)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal OrderBY As String = "")
        If CarregarDados Then
            Dim sql As String = "Select Cultura_Id, Descricao, isnull(NomeCientifico,'') as NomeCientifico, isnull(FamiliaCultura,'') as FamiliaCultura,  isnull(RegistroCultura,0) as RegistroCultura, isnull(CodigoIndeaMT,0) as CodigoIndeaMT from Cultura "
            If OrderBY.Length > 0 Then sql &= " Order By " & OrderBY

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Cultura")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cl As New Cultura
                Cl.Codigo = row("Cultura_Id")
                Cl.Descricao = row("Descricao")
                Cl.NomeCientifico = row("NomeCientifico")
                Cl.FamiliaCultura = row("FamiliaCultura")
                Cl.RegistroCultura = row("RegistroCultura")
                Cl.CodigoIndeaMT = row("CodigoIndeaMT")
                Me.Add(Cl)
            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class Cultura

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Cultura_Id, Descricao, isnull(NomeCientifico,'') as NomeCientifico, isnull(FamiliaCultura,'') as FamiliaCultura,  isnull(RegistroCultura,0) as RegistroCultura, isnull(CodigoIndeaMT,0) as CodigoIndeaMT from Cultura where Cultura_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Cultura")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Cultura_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
            _NomeCientifico = ds.Tables(0).Rows(0)("NomeCientifico")
            _FamiliaCultura = ds.Tables(0).Rows(0)("FamiliaCultura")
            _RegistroCultura = ds.Tables(0).Rows(0)("RegistroCultura")
            _CodigoIndeaMT = ds.Tables(0).Rows(0)("CodigoIndeaMT")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer = 0
    Private _Descricao As String = ""
    Private _NomeCientifico As String = ""
    Private _FamiliaCultura As String = ""
    Private _RegistroCultura As Integer = 0
    Private _CodigoIndeaMT As Integer = 0
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

    Public Property NomeCientifico() As String
        Get
            Return _NomeCientifico
        End Get
        Set(ByVal value As String)
            _NomeCientifico = value
        End Set
    End Property

    Public Property FamiliaCultura() As String
        Get
            Return _FamiliaCultura
        End Get
        Set(ByVal value As String)
            _FamiliaCultura = value
        End Set
    End Property

    Public Property RegistroCultura() As Integer
        Get
            Return _RegistroCultura
        End Get
        Set(ByVal value As Integer)
            _RegistroCultura = value
        End Set
    End Property

    Public Property CodigoIndeaMT() As Integer
        Get
            Return _CodigoIndeaMT
        End Get
        Set(ByVal value As Integer)
            _CodigoIndeaMT = value
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
                Sql = " INSERT INTO Cultura(Cultura_Id, Descricao, NomeCientifico, FamiliaCultura, RegistroCultura, CodigoIndeaMT) " & vbCrLf & _
                      " VALUES (" & _Codigo & ",'" & _Descricao & "','" & _NomeCientifico & "','" & _FamiliaCultura & "'," & _RegistroCultura & "," & _CodigoIndeaMT & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Cultura SET" & vbCrLf & _
                      "    Descricao    ='" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Cultura_Id =" & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Cultura" & vbCrLf & _
                      "  WHERE Cultura_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
