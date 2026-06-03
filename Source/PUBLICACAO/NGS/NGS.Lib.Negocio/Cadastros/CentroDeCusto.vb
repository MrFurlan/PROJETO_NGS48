Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCentroDeCusto
    Inherits List(Of CentroDeCusto)

    Public Sub New(Optional ByVal CarregarCentroDeCusto As Boolean = False, Optional ByVal Where As String = "")
        If CarregarCentroDeCusto Then
            Dim objBanco As New AcessaBanco
            Dim strSQL As String
            strSQL = "SELECT CentroDeCusto_Id, Descricao, Ativo " & vbCrLf & _
                     "  FROM CentrosDeCustos " & vbCrLf
            If Where.Length > 0 Then
                strSQL &= "Where " & Where
            End If

            strSQL &= " ORDER BY CentroDeCusto_Id"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "CentrosDeCustos")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim i As New CentroDeCusto
                i.Codigo = row("CentroDeCusto_Id").ToString()
                i.Descricao = row("Descricao").ToString()
                i.Ativo = row("Ativo")
                Add(i)
            Next
        End If
    End Sub
End Class

<Serializable()> _
Public Class CentroDeCusto
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _IUD As String
    Private _Codigo As String
    Private _Descricao As String
    Private _Ativo As Boolean

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer, Optional ByVal SomenteAtivos As Boolean = True)
        Selecionar(Codigo, SomenteAtivos)
    End Sub

#End Region

#Region "Propriedades"

    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
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

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
        End Set
    End Property

#End Region

#Region "Métodos"

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
                Sql = " INSERT INTO CentrosDeCustos (CentroDeCusto_Id, Descricao, Ativo) " & vbCrLf & _
                      " VALUES ('" & Codigo & "','" & Descricao & "','" & Ativo.ToString() & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE CentrosDeCustos SET" & vbCrLf & _
                      "    Descricao     = '" & Descricao & "'," & vbCrLf & _
                      "    Ativo     = '" & Ativo.ToString() & "'" & vbCrLf & _
                      "  WHERE CentroDecusto_Id = '" & Codigo & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE CentrosDeCustos" & vbCrLf & _
                      "  WHERE CentroDeCusto_Id = '" & Codigo & "'"
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As Integer, ByVal SomenteAtivos As Boolean) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim Sql As String = " SELECT CentroDeCusto_Id AS Codigo, Descricao, Ativo " & vbCrLf & _
                                   "   FROM CentrosDeCustos " & vbCrLf & _
                                   "  WHERE CentroDeCusto_Id = " & Codigo & vbCrLf
            If SomenteAtivos Then
                Sql &= "    AND Ativo = 1"
            End If

            Dim dsCentrosDeCustos As DataSet = objBanco.ConsultaDataSet(Sql, "CentrosDeCustos")

            With dsCentrosDeCustos.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Codigo").ToString()
                    Me.Descricao = .Item(0)("Descricao").ToString()
                    Me.Ativo = .Item(0)("Ativo")

                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

    Public Function JaUtilizado() As Boolean
        Dim objBanco As New AcessaBanco()
        Try
            Dim Sql As String = " SELECT TOP 1 1 " & vbCrLf & _
                                "   FROM (" & vbCrLf & _
                                "         SELECT CentroDeCusto" & vbCrLf & _
                                "           FROM NotasFiscaisXEncargos " & vbCrLf & _
                                "          UNION ALL" & vbCrLf & _
                                "         SELECT CentroDeCusto_Id" & vbCrLf & _
                                "           FROM NotasFiscaisXItensXRateio" & vbCrLf & _
                                "          UNION ALL" & vbCrLf & _
                                "         SELECT Custo  " & vbCrLf & _
                                "           FROM Razao " & vbCrLf & _
                                "         ) AS X " & vbCrLf & _
                                "   WHERE CentroDeCusto = '" & Codigo & "'" & vbCrLf

            If objBanco.ConsultaDataSet(Sql, "CentrosDeCustos").Tables(0).Rows.Count > 0 Then
                Return True
            Else : Return False
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function
#End Region

End Class