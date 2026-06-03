Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()>
Public Class ListCnpjSefaz
    Inherits List(Of CnpjSefaz)

    Public Sub New(Optional ByVal CarregarCnpjSefaz As Boolean = False, Optional ByVal Where As String = "")
        If CarregarCnpjSefaz Then
            Dim objBanco As New AcessaBanco
            Dim strSQL As String
            strSQL = "SELECT Cnpj_Id, Estado, Cidade " & vbCrLf &
                     "  FROM CnpjsSefaz " & vbCrLf

            If Where.Length > 0 Then
                strSQL &= "Where " & Where
            End If

            strSQL &= " ORDER BY Cnpj_Id"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "CnpjsSefaz")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim i As New CnpjSefaz
                i.Cnpj = row("Cnpj_Id")
                i.Estado = row("Estado")
                i.Cidade = row("Cidade")
                Add(i)
            Next
        End If
    End Sub

    Public Shared Widening Operator CType(v As ListCnpjSefaz) As DataSet
        Throw New NotImplementedException()
    End Operator
End Class

<Serializable()>
Public Class CnpjSefaz
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _IUD As String
    Private _Cnpj As String
    Private _Estado As String
    Private _Cidade As String

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Selecionar(Codigo)
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

    Public Property Cnpj() As String
        Get
            Return _Cnpj
        End Get
        Set(ByVal value As String)
            _Cnpj = value
        End Set
    End Property

    Public Property Estado() As String
        Get
            Return _Estado
        End Get
        Set(ByVal value As String)
            _Estado = value
        End Set
    End Property

    Public Property Cidade() As String
        Get
            Return _Cidade
        End Get
        Set(ByVal value As String)
            _Cidade = value
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
                Sql = " INSERT INTO CnpjsSefaz (Cnpj_Id, Estado, Cidade) " & vbCrLf &
                      " VALUES ('" & Cnpj & "','" & Estado & "','" & Cidade & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE CnpjsSefaz SET" & vbCrLf &
                      "    Estado = '" & Estado & "'," & vbCrLf &
                      "    Cidade = '" & Cidade & "'" & vbCrLf &
                      "  WHERE Cnpj_Id = '" & Cnpj & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE CnpjsSefaz" & vbCrLf &
                      "  WHERE Cnpj_Id = '" & Cnpj & "'"
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As Integer) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim Sql As String = " SELECT Cnpj_Id AS Codigo, Estado, Cidade " & vbCrLf &
                                   "   FROM CnpjsSefaz " & vbCrLf &
                                   "  WHERE Sefaz_Id = '" & Codigo & "'" & vbCrLf

            Dim dsCentrosDeCustos As DataSet = objBanco.ConsultaDataSet(Sql, "CnpjsSefaz")

            With dsCentrosDeCustos.Tables(0).Rows
                If .Count > 0 Then
                    Me.Cnpj = .Item(0)("Codigo")
                    Me.Estado = .Item(0)("Estado")
                    Me.Cidade = .Item(0)("Cidade")

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

#End Region

End Class