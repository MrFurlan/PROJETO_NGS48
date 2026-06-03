Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListSeguimento
    Inherits List(Of Seguimento)

    Public Sub New(Optional ByVal CarregarDados As Boolean = False)
        If CarregarDados Then
            Dim sql As String = "Select Seguimento_Id, Descricao from Seguimentos"
            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "Seguimento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Mc As New Seguimento
                Mc.Codigo = row("Seguimento_Id")
                Mc.Descricao = row("Descricao")
                Me.Add(Mc)
            Next

        End If
    End Sub

End Class

<Serializable()>
Public Class Seguimento
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Seguimento_Id, Descricao from Seguimentos where Seguimento_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "Seguimento")
        If ds.Tables(0).Rows.Count > 0 Then
            _Codigo = ds.Tables(0).Rows(0)("Seguimento_Id")
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
                If LiberaExclusaoDaSeguimento() Then
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


    Public Function LiberaExclusaoDaSeguimento() As Boolean
        Dim sql As String = String.Empty
        Dim Banco As New AcessaBanco

        sql = "SELECT 1 " & vbCrLf &
            "    FROM Produtos " & vbCrLf &
            "   WHERE Seguimento = " & Me.Codigo
        If Banco.ConsultaDataSet(sql, "SeguimentoProduto").Tables(0).Rows.Count > 0 Then
            Throw New Exception("Esta Seguimento já está sendo usada em pelo menos um produto.")
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
                Sql = " INSERT INTO Seguimentos(Descricao) " & vbCrLf &
                      " VALUES ('" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Seguimentos SET" & vbCrLf &
                      "    Descricao    ='" & _Descricao & "'" & vbCrLf &
                      "  WHERE Seguimento_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Seguimentos" & vbCrLf &
                      "  WHERE Seguimento_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class