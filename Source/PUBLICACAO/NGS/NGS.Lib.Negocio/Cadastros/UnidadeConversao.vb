<Serializable()> _
Public Class ListUnidadeConversao
    Inherits List(Of UnidadeConversao)

#Region "Construtor"
    Public Sub New()
        Dim Banco As New AcessaBanco

        Try
            Dim Sql As String = "SELECT UnidadeOrigem_Id, UnidadeDestino_Id, Fator " & vbCrLf & _
                                "  FROM UnidadeConversao" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr As DataRow In ds.Tables(0).Rows
                    Dim obj As New UnidadeConversao
                    obj.CodigoUnidadeOrigem = dr("UnidadeOrigem_Id")
                    obj.CodigoUnidadeDestino = dr("UnidadeDestino_Id")
                    obj.Fator = dr("Fator")
                    Me.Add(obj)
                Next
            End If
        Catch ex As Exception
            Throw New Exception(ex.Message)
        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

End Class

<Serializable()> _
Public Class UnidadeConversao
    Implements IBaseEntity


#Region "Constructor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal UnidadeOrigem As String, ByVal UnidadeDestino As String)
        Dim Banco As New AcessaBanco
        Try
            Dim sql As String = String.Empty
            sql = "SELECT UnidadeOrigem_Id, UnidadeDestino_Id, Fator" & vbCrLf & _
                  "  FROM UnidadeConversao " & vbCrLf & _
                  " WHERE UnidadeOrigem_Id = '" & UnidadeOrigem & "'" & vbCrLf & _
                  "   AND UnidadeDestino_Id = '" & UnidadeDestino & "'" & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "UnidadeConversao")

            If ds.Tables(0).Rows.Count > 0 Then
                Dim dr As DataRow = ds.Tables(0).Rows(0)
                Me.CodigoUnidadeOrigem = dr("UnidadeOrigem_Id")
                Me.CodigoUnidadeDestino = dr("UnidadeDestino_Id")
                Me.Fator = dr("Fator")
            End If
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region


#Region "Fields"
    Public Property Erro As Exception
    Public Property IUD As String

    Private _CodigoUnidadeOrigem As String
    Public Property CodigoUnidadeOrigem() As String
        Get
            Return _CodigoUnidadeOrigem
        End Get
        Set(value As String)
            _CodigoUnidadeOrigem = value
            UnidadeOrigem = Nothing
        End Set
    End Property

    Private _UnidadeOrigem As UnidadeDeMedida
    Public Property UnidadeOrigem() As UnidadeDeMedida
        Get
            If _UnidadeDestino Is Nothing AndAlso Not String.IsNullOrWhiteSpace(CodigoUnidadeOrigem) Then UnidadeOrigem = New UnidadeDeMedida(Me.CodigoUnidadeOrigem)
            Return _UnidadeOrigem
        End Get
        Set(ByVal value As UnidadeDeMedida)
            _UnidadeOrigem = value
        End Set
    End Property

    Private _CodigoUnidadeDestino As String
    Public Property CodigoUnidadeDestino() As String
        Get
            Return _CodigoUnidadeDestino
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDestino = value
            UnidadeDestino = Nothing
        End Set
    End Property

    Private _UnidadeDestino As UnidadeDeMedida
    Public Property UnidadeDestino() As UnidadeDeMedida
        Get
            If _UnidadeDestino Is Nothing AndAlso Not String.IsNullOrWhiteSpace(CodigoUnidadeDestino) Then UnidadeDestino = New UnidadeDeMedida(CodigoUnidadeDestino)
            Return _UnidadeDestino
        End Get
        Set(ByVal value As UnidadeDeMedida)
            _UnidadeDestino = value
        End Set
    End Property

    Public Property Fator As Decimal
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
                Sql = " INSERT INTO UnidadeConversao(UnidadeOrigem_Id, UnidadeDestino_Id, Fator) " & vbCrLf & _
                      " VALUES ('" & Me.CodigoUnidadeOrigem & "','" & Me._CodigoUnidadeDestino & "'," & Str(Me.Fator) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE UnidadeConversao " & vbCrLf & _
                      "    SET Fator = " & Str(Me.Fator) & vbCrLf & _
                      "  WHERE UnidadeOrigem_Id    ='" & Me.CodigoUnidadeOrigem & "'" & vbCrLf & _
                      "    AND UnidadeDestino_Id    ='" & Me.CodigoUnidadeDestino & "'" & vbCrLf

                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE UnidadeConversao" & vbCrLf & _
                      "  WHERE UnidadeOrigem_Id    ='" & Me.CodigoUnidadeOrigem & "'" & vbCrLf & _
                      "    AND UnidadeDestino_Id    ='" & Me.CodigoUnidadeDestino & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region


End Class
