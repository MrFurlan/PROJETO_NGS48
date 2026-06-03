

Public Class ListPlanoDeContaXReferencialBacen
    Inherits List(Of PlanoDeContaXReferencialBacen)
    Implements IBaseEntity

#Region "Construtors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Conta As String)
        Dim Banco As New AcessaBanco

        Dim sql As String = "select ID as Codigo, Conta, Referencial, isnull(Grupo, '') as Grupo, isnull(Produto, '') as Produto, ISNULL(CodigodeCustoECf, 0) AS CodigodeCustoECf, isnull(CodigoDeCustoECD, '') as CodigoDeCustoECD  " & vbCrLf & _
                            "  from PlanoDeContasXReferencialBacen                                                                     " & vbCrLf & _
                            " where Conta = '" & Conta & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PlanoDeContasXReferencialBacen")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim obj As New [Lib].Negocio.PlanoDeContaXReferencialBacen()
            obj.Codigo = row("Codigo")
            obj.Conta = row("Conta")
            obj.Referencial = row("Referencial")
            obj.CodigoGrupo = row("Grupo")
            obj.CodigoProduto = row("Produto")
            obj.CodigodeCustoECf = row("CodigodeCustoECf")
            obj.CodigodeCustoECD = row("CodigoDeCustoECD")

            Me.Add(obj)
        Next
    End Sub

#End Region

End Class

Public Class PlanoDeContaXReferencialBacen

#Region "Construtors"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim Banco As New AcessaBanco

        Dim sql As String = "select ID as Codigo, Conta, Referencial, isnull(Grupo, '') as Grupo, isnull(Produto, '') as Produto, ISNULL(CodigodeCustoECf, 0) AS CodigodeCustoECf, isnull(CodigoDeCustoECD, '') as CodigoDeCustoECD " & vbCrLf & _
                            "  from PlanoDeContasXReferencialBacen" & vbCrLf & _
                            " where Codigo = '" & Codigo & "'" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "PlanoDeContasXReferencialBacen")

        For Each row As DataRow In ds.Tables(0).Rows
            Me.Codigo = row("Codigo")
            Me.Conta = row("Conta")
            Me.Referencial = row("Referencial")
            Me.CodigoGrupo = row("Grupo")
            Me.CodigoProduto = row("Produto")
            Me.CodigodeCustoECf = row("CodigodeCustoECf")
            Me.CodigodeCustoECD = row("CodigoDeCustoECD")
        Next
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _Conta As String
    Private _Referencial As String
    Private _CodigoGrupo As String
    Private _Grupo As GrupoProduto
    Private _DescricaoGrupo As String
    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigodeCustoECf As Integer
    Private _CodigodeCustoECD As String
#End Region

#Region "Properts"
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

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing AndAlso Not String.IsNullOrWhiteSpace(_CodigoEmpresa) Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property Conta() As String
        Get
            Return _Conta
        End Get
        Set(ByVal value As String)
            _Conta = value
        End Set
    End Property

    Public Property Referencial() As String
        Get
            Return _Referencial
        End Get
        Set(ByVal value As String)
            _Referencial = value
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

    Public Property Grupo() As GrupoProduto
        Get
            If _Grupo Is Nothing AndAlso Not String.IsNullOrWhiteSpace(_CodigoGrupo) Then _Grupo = New GrupoProduto(_CodigoGrupo)
            Return _Grupo
        End Get
        Set(ByVal value As GrupoProduto)
            _Grupo = value
        End Set
    End Property

    Public ReadOnly Property DescricaoProduto() As String
        Get
            Dim str As String = ""
            If Not String.IsNullOrWhiteSpace(_CodigoProduto) AndAlso Produto IsNot Nothing Then
                str = _CodigoProduto & " - " & _Produto.Descricao
            End If
            Return str
        End Get
    End Property

    Public ReadOnly Property DescricaoGrupo() As String
        Get
            Dim str As String = ""
            If Not String.IsNullOrWhiteSpace(_CodigoGrupo) Then
                str = _CodigoGrupo & " - " & Me.Grupo.Descricao
            End If
            Return str
        End Get
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing AndAlso Not String.IsNullOrWhiteSpace(_CodigoProduto) Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigodeCustoECf() As Integer
        Get
            Return _CodigodeCustoECf
        End Get
        Set(ByVal value As Integer)
            _CodigodeCustoECf = value
        End Set
    End Property

    Public Property CodigodeCustoECD() As String
        Get
            Return _CodigodeCustoECD
        End Get
        Set(ByVal value As String)
            _CodigodeCustoECD = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList) As ArrayList
        Dim sql As String
        Dim ObjBanco As New AcessaBanco

        Select Case Me.IUD
            Case "I"
                sql = "  Merge PlanoDeContasXReferencialBacen as Target" & vbCrLf & _
                      "  Using (Select '99999999999999' as Empresa, 0 as EndEmpresa, '" & Me.Conta & "' as Conta, '" & Me.Referencial & "' as Referencial, " & _
                                 IIf(Me.CodigoGrupo = "NULL", "NULL", "'" & Me.CodigoGrupo & "'") & " as Grupo,  " & IIf(Me.CodigoProduto = "NULL", "NULL", "'" & Me.CodigoProduto & "'") & " as Produto, " & _
                                 IIf(Me.CodigodeCustoECf = 0, "NULL", Me.CodigodeCustoECf) & " as CodigodeCustoECF, " & IIf(Me.CodigodeCustoECD = 0, "NULL", "'" & Me.CodigodeCustoECD & "'") & " as CodigodeCustoECD) as Source" & vbCrLf & _
                      "     on Target.Empresa     = Source.Empresa" & vbCrLf & _
                      "    and Target.EndEmpresa  = Source.EndEmpresa" & vbCrLf & _
                      "    and Target.Conta       = Source.Conta" & vbCrLf & _
                      "    and Target.Referencial = Source.Referencial" & vbCrLf & _
                      "    and Target.Grupo       = Source.Grupo" & vbCrLf & _
                      "    and Target.Produto     = Source.Produto" & vbCrLf & _
                      "   WHEN MATCHED THEN" & vbCrLf & _
                      "     UPDATE set Target.CodigodeCustoECF = Source.CodigodeCustoECF, Target.CodigodeCustoECD = Source.CodigodeCustoECD" & vbCrLf & _
                      "   WHEN NOT MATCHED BY TARGET THEN" & vbCrLf & _
                      "     Insert (Empresa, EndEmpresa, Conta, Referencial, Grupo, Produto, CodigodeCustoECf, CodigodeCustoECD)" & vbCrLf & _
                      "       Values (Source.Empresa, Source.EndEmpresa, Source.Conta, Source.Referencial, Source.Grupo, Source.Produto, Source.CodigodeCustoECf, CodigodeCustoECD);" & vbCrLf

                sqls.Add(sql)
            Case "U"
                sql = "UPDATE PlanoDeContasXReferencialBacen" & vbCrLf & _
                      "set CodigodeCustoECF = " & Me.CodigodeCustoECf & ", CodigodeCustoECD = '" & Me.CodigodeCustoECD & "'" & vbCrLf & _
                      "WHERE ID = " & Me.Codigo & vbCrLf
                sqls.Add(sql)
            Case "D"
                sql = "DELETE FROM PlanoDeContasXReferencialBacen" & vbCrLf & _
                      "WHERE ID = " & Me.Codigo & vbCrLf
                sqls.Add(sql)
        End Select

        Return sqls
    End Function

#End Region

End Class
