Imports NGS.Lib.Negocio

<Serializable()> _
Public Class ListTipoDeCertidao
    Inherits List(Of TipoDeCertidao)
    Implements IBaseEntity

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Carregar As Boolean)
        Dim objBanco As New AcessaBanco
        Dim Sql As String

        Sql = "SELECT Tipo_id, Descricao " & vbCrLf & _
              "  FROM TipoDeCertidao " & vbCrLf & _
              " ORDER BY Tipo_id " & vbCrLf

        Dim ds As DataSet = objBanco.ConsultaDataSet(Sql, "TipoDeCertidao")

        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Tc As New TipoDeCertidao
                Tc.Codigo = row("Tipo_id")
                Tc.Descricao = row("Descricao")
                Add(Tc)
            Next
        End If
    End Sub
#End Region

End Class

<Serializable()> _
Public Class TipoDeCertidao

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As String)
        Selecionar(Codigo)
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
#End Region

#Region "Properties"
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
                Sql = " INSERT INTO TipoDeCertidao (Descricao) " & vbCrLf & _
                      " VALUES ('" & _Descricao & "')"
                Sqls.Add(Sql)
            Case "U"

                Sql = " UPDATE TipoDeCertidao " & vbCrLf & _
                      "    SET Descricao = '" & _Descricao & "'" & vbCrLf & _
                      "  WHERE Tipo_Id =  " & _Codigo & ";"
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE TipoDeCertidao" & vbCrLf & _
                      "  WHERE Tipo_Id = " & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim Sql As String = "SELECT Tipo_Id, Descricao" & _
                                   "  FROM TipoDeCertidao " & _
                                   " WHERE Tipo_Id = " & Codigo

            Dim ds As DataSet = objBanco.ConsultaDataSet(Sql, "TipoDeCertidao")

            With ds.Tables(0).Rows
                If .Count > 0 Then
                    Me.Codigo = .Item(0)("Tipo_Id")
                    Me.Descricao = .Item(0)("Descricao").ToString()
                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Throw New Exception
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region
End Class