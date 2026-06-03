Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Fil

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal nomeArquivo As String)
        Dim sql As String = "SELECT * FROM nfe.fil WHERE nomearquivoped = '" & nomeArquivo & "'"
        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "fil")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            Me.Codigo = row("codped")
            Me.NomeArquivo = row("nomearquivoped")
            Me.Texto = row("texto")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Texto As String
    Private _NomeArquivo As String
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

    Public Property Texto() As String
        Get
            Return _Texto
        End Get
        Set(ByVal value As String)
            _Texto = value
        End Set
    End Property

    Public Property NomeArquivo() As String
        Get
            Return _NomeArquivo
        End Get
        Set(ByVal value As String)
            _NomeArquivo = value
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
                Sql = "DELETE FROM NFE.RESP WHERE nomearquivoresp = '" & String.Format("resp-{0}", _NomeArquivo) & "'"
                Sqls.Add(Sql)

                Sql = "INSERT INTO NFE.FIL" & vbCrLf & _
                      "     (TEXTO, NOMEARQUIVOPED) " & vbCrLf & _
                      " VALUES " & vbCrLf & _
                      "     ('" & _Texto & "', '" & _NomeArquivo & "') "
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE NFE.FIL" & vbCrLf & _
                      "     SET TEXTO           = '" & _Texto & "'," & vbCrLf & _
                      "         NOMEARQUIVOPED  = '" & _NomeArquivo & "'" & vbCrLf & _
                      "     WHERE CODPED = " & _Codigo
                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE NFE.FIL" & vbCrLf & _
                      "     WHERE CODPED = " & _Codigo
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
