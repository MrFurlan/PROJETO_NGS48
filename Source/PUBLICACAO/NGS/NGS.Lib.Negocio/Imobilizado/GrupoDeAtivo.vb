Public Class GrupoDeAtivo

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = "Select Grupo_Id, Descricao, PercentualDepreciacao from GruposDeAtivos where Categoria_Id = " & pCodigo

        ds = Banco.ConsultaDataSet(sql, "GruposDeAtivos")
        If ds.Tables(0).Rows.Count > 0 Then
            With ds.Tables(0).Rows(0)
                Me.CodigoGrupo = ("Grupo_Id")
                Me.Descricao = ("Descricao")
                Me.PercentualDepreciacao = ("PercentualDepreciacao")
            End With

            'Me.CodigoGrupo = ds.Tables(0).Rows(0)("Grupo_Id")
            'Me.Descricao = ds.Tables(0).Rows(0)("Descricao")
            'Me.PercentualDepreciacao = ds.Tables(0).Rows(0)("PercentualDepreciacao")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoGrupo As String
    Private _Descricao As String
    Private _PercentualDepreciacao As String
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

    Public Property CodigoGrupo() As String
        Get
            Return _CodigoGrupo
        End Get
        Set(ByVal value As String)
            _CodigoGrupo = value
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

    Public Property PercentualDepreciacao() As String
        Get
            Return _PercentualDepreciacao
        End Get
        Set(ByVal value As String)
            _PercentualDepreciacao = value
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
                Sql = String.Format(" INSERT INTO GruposDeAtivos(Grupo_Id, Descricao, PercentualDepreciacao) VALUES ('{0}','{1}','{2}')", Me.CodigoGrupo, Me.Descricao, Str(Me.PercentualDepreciacao))
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE GruposDeAtivos SET" & vbCrLf & _
                      "    Descricao             = '" & Me.Descricao & "'" & vbCrLf & _
                      "    PercentualDepreciacao = '" & Str(Me.PercentualDepreciacao) & "'" & vbCrLf & _
                      "  WHERE Grupo_Id = '" & Me.CodigoGrupo & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE GruposDeAtivos" & vbCrLf & _
                      "  WHERE Grupo_Id = " & Me.CodigoGrupo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
