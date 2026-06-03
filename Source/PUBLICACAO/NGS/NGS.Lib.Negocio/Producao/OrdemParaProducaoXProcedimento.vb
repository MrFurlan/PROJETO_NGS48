Public Class ListOrdemParaProducaoXProcedimento
    Inherits List(Of OrdemParaProducaoXProcedimento)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objOrdemDeProducao As OrdemParaProducao)

        Me.OrdemParaProducao = objOrdemDeProducao

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Ordem_Id, Procedimento_Id " & vbCrLf & _
                     "From OrdemDeProducaoXProcedimento " & vbCrLf & _
                     "Where Empresa_id    = '" & Me.OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                     "  and EndEmpresa_Id = " & Me.OrdemParaProducao.EnderecoEmpresa & vbCrLf

            If Me.OrdemParaProducao.Codigo > 0 Then strSQL &= "  and Ordem_Id      = " & Me.OrdemParaProducao.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "OrdemParaProducaoXProcedimento")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim oPProcedimento As New OrdemParaProducaoXProcedimento(Me.OrdemParaProducao)

                oPProcedimento.CodigoProcedimento = row("Procedimento_Id")
                oPProcedimento.Procedimento = New ProcedimentoParaProducao(row("Procedimento_Id"))

                Me.Add(oPProcedimento)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As OrdemParaProducaoXProcedimento In Me
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
Public Class OrdemParaProducaoXProcedimento
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _OrdemParaProducao As OrdemParaProducao

    Private _CodigoProcedimento As Integer
    Private _Procedimento As ProcedimentoParaProducao

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

    Public Property CodigoProcedimento As Integer
        Get
            Return _CodigoProcedimento
        End Get
        Set(value As Integer)
            _CodigoProcedimento = value
        End Set
    End Property

    Public Property Procedimento As ProcedimentoParaProducao
        Get
            If _Procedimento Is Nothing And _CodigoProcedimento > 0 Then _Procedimento = New ProcedimentoParaProducao(_CodigoProcedimento)
            Return _Procedimento
        End Get
        Set(value As ProcedimentoParaProducao)
            _Procedimento = value
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
                strSql = "INSERT INTO OrdemDeProducaoXProcedimento(Empresa_Id, EndEmpresa_Id, Ordem_Id, Procedimento_Id)" & vbCrLf & _
                         "Values('" & OrdemParaProducao.CodigoEmpresa & "'," & OrdemParaProducao.EnderecoEmpresa & "," & OrdemParaProducao.Codigo & "," & Me.CodigoProcedimento & ")"

            Case "D"
                strSql = "Delete OrdemDeProducaoXProcedimento " & vbCrLf & _
                         " Where Empresa_Id             = '" & OrdemParaProducao.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id          = " & OrdemParaProducao.EnderecoEmpresa & vbCrLf & _
                         "   and Ordem_Id               = " & OrdemParaProducao.Codigo & vbCrLf & _
                         "   and Procedimento_Id           = " & Me.CodigoProcedimento
        End Select

        Sqls.Add(strSql)
    End Sub
#End Region

End Class
