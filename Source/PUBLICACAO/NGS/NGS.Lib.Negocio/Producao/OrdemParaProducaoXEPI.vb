Public Class ListOrdemParaProducaoXEPI
    Inherits List(Of OrdemParaProducaoXEPI)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, EPI_Id " & vbCrLf & _
                     "From OrdemDeProducaoXEPI " & vbCrLf & _
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXEPI")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPEpi As New OrdemParaProducaoXEPI(Me.OrdemParaProducao)

                oPEpi.CodigoEPI = row("EPI_Id")
                oPEpi.EPI = New EPI(row("EPI_Id"))

                Me.Add(oPEpi)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXEPI In Me
            If item.IUD = "I" OrElse item.IUD = "D" Then item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OrdemParaProducao As OrdemParaProducao
#End Region

#Region "Property"
    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class OrdemParaProducaoXEPI
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoEPI As Integer
    Private _EPI As EPI

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

    Public Property OrdemParaProducao() As OrdemParaProducao
        Get
            Return _OrdemParaProducao
        End Get
        Set(ByVal value As OrdemParaProducao)
            _OrdemParaProducao = value
        End Set
    End Property

    Public Property CodigoEPI As Integer
        Get
            Return _CodigoEPI
        End Get
        Set(value As Integer)
            _CodigoEPI = value
        End Set
    End Property

    Public Property EPI As EPI
        Get
            If _EPI Is Nothing And _CodigoEPI > 0 Then _EPI = New EPI(_CodigoEPI)
            Return _EPI
        End Get
        Set(value As EPI)
            _EPI = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal OrdemParaProducao As OrdemParaProducao)
        Me.OrdemParaProducao = OrdemParaProducao
    End Sub
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"
                strSql = "INSERT INTO OrdemDeProducaoXEPI(Empresa_Id, EndEmpresa_Id, Ordem_Id, EPI_Id)" & vbCrLf & _
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & "," & Me.CodigoEPI & ")"

            Case "D"
                strSql = "Delete OrdemDeProducaoXEPI " & vbCrLf & _
                         " Where Empresa_Id             = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id          = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf & _
                         "   and Ordem_Id               = " & OrdemParaProducao.Codigo & vbCrLf & _
                         "   and EPI_Id           = " & Me.CodigoEPI
        End Select

        Sqls.Add(strSql)
    End Sub
#End Region

End Class
