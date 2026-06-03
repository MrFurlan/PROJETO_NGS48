Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMarca
    Inherits List(Of Marca)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        If CarregarDados Then
            Dim sql As String = "Select Marca_Id, Descricao from Marca"
            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Marca")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Mc As New Marca
                Mc.Codigo = row("Marca_Id")
                Mc.Descricao = row("Descricao")
                Me.Add(Mc)
            Next

        End If
    End Sub

End Class

<Serializable()> _
Public Class Marca
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Marca_Id, Descricao from Marca where Marca_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Marca")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Marca_Id")
            _Descricao = ds.Tables(0).Rows(0)("Descricao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Descricao As String
#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            If value = "D" Then
                If LiberaExclusaoDaMarca() Then
                    _IUD = value
                End If
            Else
                _IUD = value
            End If
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


    Public Function LiberaExclusaoDaMarca() As Boolean
        Dim sql As String = String.Empty
        Dim Banco As New AcessaBanco

        sql = "SELECT 1 " & vbCrLf & _
            "    FROM Produtos " & vbCrLf & _
            "   WHERE Marca = " & Me.Codigo
        If Banco.ConsultaDataSet(sql, "MarcaProduto").Tables(0).Rows.Count > 0 Then
            Throw New Exception("Esta Marca já está sendo usada em pelo menos um produto.")
            Return False
        Else
            Return True
        End If
    End Function
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
                Sql = " INSERT INTO Marca(Descricao) " & vbCrLf & _
                      " VALUES ('" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Marca SET" & vbCrLf & _
                      "    Descricao    ='" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Marca_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Marca" & vbCrLf & _
                      "  WHERE Marca_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class