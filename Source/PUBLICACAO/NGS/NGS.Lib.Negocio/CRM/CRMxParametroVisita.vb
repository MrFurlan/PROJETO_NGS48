Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis


<Serializable()> _
Public Class ListCRMxParametroVisita
    Inherits List(Of CRMxParametroVisita)

    Public Sub New()

    End Sub

    Public Sub New(ByVal pAno As Integer, Optional ByVal pEmpresa As String = "")
        Dim Sql As String = ""
        Sql = "select case when cxv.MinimoVisitas is null then 'I' else '' end as IUD," & vbCrLf & _
              "       Cxc.Empresa_Id as CodigoEmpresa, " & vbCrLf & _
                "	  Cxc.EndEmpresa_Id as EndEmpresa," & vbCrLf & _
                "	  Cxc.Ano_Id as Ano," & vbCrLf & _
                "	  Cxc.Consolidado_Id as Consolidado," & vbCrLf & _
                "	  isnull(R.Cliente_Id,'NENHUM') as codigoRepresentante," & vbCrLf & _
                "     isnull(Cr.EndRepresentante_Id,0) as EndRepresentante," & vbCrLf & _
                "     ISNULL(R.nome, 'NENHUM') NomeRepresentante, " & vbCrLf & _
                "     cxc.Classificacao," & vbCrLf & _
                "     CXC.TipoClienteQualitativo," & vbCrLf & _
                "      Cxc.Cliente_Id as CodigoCliente, " & vbCrLf & _
                "       c.Nome as NomeCliente, " & vbCrLf & _
                "       c.Endereco_id as EndCliente," & vbCrLf & _
                "       c.Complemento as Fazenda, " & vbCrLf & _
                "       ISNULL(cxv.MinimoVisitas, 0) as MinimoVisita" & vbCrLf & _
                "  from CRMParametroXCliente Cxc" & vbCrLf & _
                " inner join Clientes C " & vbCrLf & _
                "	 on  Cxc.Cliente_Id = C.Cliente_Id " & vbCrLf & _
                "  left join ClienteXRepresentante Cr " & vbCrLf & _
                "	 on cR.Cliente_Id    = C.Cliente_Id " & vbCrLf & _
                "   and Cr.EndCliente_Id = C.Endereco_Id" & vbCrLf & _
                "   and Cr.Principal     = 1" & vbCrLf & _
                "  left join Clientes R " & vbCrLf & _
                "	on  R.Cliente_Id  = Cr.Representante_Id" & vbCrLf & _
                "   and R.Endereco_id = Cr.EndRepresentante_id" & vbCrLf & _
                "  left Join CRMxParametroVisita CxV" & vbCrLf & _
                "    on CxV.Empresa_Id      = Cxc.Empresa_Id" & vbCrLf & _
                "   and CxV.EndEmpresa_Id   = Cxc.EndEmpresa_Id" & vbCrLf & _
                "   and CxV.Ano_Id          = Cxc.Ano_Id" & vbCrLf & _
                "   and CxV.Consolidado_Id  = Cxc.Consolidado_Id  " & vbCrLf & _
                "   and CxV.Cliente_Id      = C.Cliente_Id " & vbCrLf & _
                "   and CxV.EndCliente_Id   = C.Endereco_Id " & vbCrLf & _
                " where TipoCliente in ('C','P','E') " & vbCrLf & _
                "   AND cxc.Ano_Id     = " & pAno & vbCrLf
        If Not String.IsNullOrWhiteSpace(pEmpresa) Then
            Sql &= "   and cxc.Empresa_Id ='" & pEmpresa & "'" & vbCrLf
        End If
        Sql &= "   and cxc.Consolidado_Id = 0 " & vbCrLf & _
                " order by ordem" & vbCrLf
        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(Sql, "Visitas")

        For Each row In ds.Tables(0).Rows
            Dim vis As New CRMxParametroVisita
            vis.IUD = row("IUD")
            vis.CodigoEmpresa = row("CodigoEmpresa")
            vis.EndEmpresa = row("EndEmpresa")
            vis.Ano = row("Ano")
            vis.Consolidado = row("Consolidado")
            vis.CodigoRepresentante = row("CodigoRepresentante")
            vis.EndRepresentante = row("EndRepresentante")
            vis.NomeRepresentante = row("NomeRepresentante")
            vis.Classificacao = row("Classificacao")
            vis.Qualitativo = row("TipoClienteQualitativo")
            vis.CodigoCliente = row("CodigoCliente")
            vis.EndCliente = row("EndCliente")
            vis.NomeCliente = row("NomeCliente")
            vis.Fazenda = row("Fazenda")
            vis.MinimoVisita = row("MinimoVisita")

            Me.Add(vis)
        Next
    End Sub

    Public ReadOnly Property Representantes As Object
        Get
            Return (From p In Me Order By p.NomeRepresentante Select p.CodigoRepresentante, p.EndRepresentante, p.NomeRepresentante).Distinct()
        End Get
    End Property

    Public ReadOnly Property ClientesPorRepresentante(ByVal CodigoRepresentante As String) As Object
        Get
            Return (From c In Me Where c.CodigoRepresentante = CodigoRepresentante Order By c.Classificacao, c.NomeCliente Select id = c.CodigoCliente & "-" & c.EndCliente.ToString, c.NomeCliente, c.Fazenda, c.Classificacao, c.Qualitativo, c.MinimoVisita)
        End Get
    End Property


#Region "Methods"
    Public Function Salvar() As Boolean
        Dim sqls As New ArrayList
        SalvarSql(sqls)

        Dim Banco As New AcessaBanco
        If Banco.GravaBanco(sqls) Then
            For Each row In Me
                row.IUD = ""
            Next
            Return True
        Else
            Return False
        End If
    End Function


    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As CRMxParametroVisita In Me
            item.SalvarSql(Sqls)
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class CRMxParametroVisita
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(Optional ByVal pListParametrosVisita As ListCRMxParametroVisita = Nothing)
        _ParametrosVisita = pListParametrosVisita
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    '*****************************************
    Private _ParametrosVisita As ListCRMxParametroVisita
    '**** Ligação com o CRM *****************'
    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Ano As Integer
    Private _Consolidado As Integer

    '**** Representante *****************'
    Private _CodigoRepresentante As String = ""
    Private _NomeRepresentante As String = ""
    Private _EndRepresentante As String = ""
    Private _Representante As Cliente

    '**** Cliente *****************'
    Private _Cliente As Cliente
    Private _CodigoCliente As String = ""
    Private _NomeCliente As String = ""
    Private _EndCliente As Integer
    Private _Fazenda As String = ""


    '********* Exibicao ***********
    Private _Classificacao As String = ""
    Private _Qualitativo As String = ""

    '**** Número mínimo de visitas *****************'
    Private _MinimoVisita As Integer

    '********* Visitas ***************
    Private _Visitas As ListVisita

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

    Public Property ParametrosVisita As ListCRMxParametroVisita
        Get
            If _ParametrosVisita Is Nothing And Me.CodigoEmpresa.Length > 0 Then
                _ParametrosVisita = New ListCRMxParametroVisita(Me.Ano, Me.CodigoEmpresa)
            End If
            Return _ParametrosVisita
        End Get
        Set(ByVal value As ListCRMxParametroVisita)
            _ParametrosVisita = value
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
    Public Property EndEmpresa() As String
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As String)
            _EndEmpresa = value
        End Set
    End Property
    Public Property Ano() As String
        Get
            Return _Ano
        End Get
        Set(ByVal value As String)
            _Ano = value
        End Set
    End Property
    Public Property Consolidado As Integer
        Get
            Return _Consolidado
        End Get
        Set(ByVal value As Integer)
            _Consolidado = value
        End Set
    End Property

    Public Property CodigoRepresentante() As String
        Get
            Return _CodigoRepresentante
        End Get
        Set(ByVal value As String)
            _CodigoRepresentante = value
            _Representante = Nothing
            _NomeRepresentante = ""
        End Set
    End Property
    Public Property EndRepresentante As String
        Get
            Return _EndRepresentante
        End Get
        Set(ByVal value As String)
            _EndRepresentante = value
            _Representante = Nothing
            _NomeRepresentante = ""
        End Set
    End Property
    Public Property Representante() As Cliente
        Get
            If _Representante Is Nothing And Me.CodigoRepresentante.Length > 0 Then _Representante = New Cliente(Me.CodigoRepresentante, Me.EndRepresentante)
            Return _Representante
        End Get
        Set(ByVal value As Cliente)
            _Representante = value
        End Set
    End Property
    Public Property NomeRepresentante As String
        Get
            If _NomeRepresentante = "" Then
                If Me.Representante Is Nothing Then
                    Return ""
                Else
                    Return Me.Representante.Nome
                End If
            Else
                Return _NomeRepresentante
            End If
        End Get

        Set(ByVal value As String)
            _NomeRepresentante = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property
    Public Property NomeCliente() As String
        Get
            Return _NomeCliente

        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property
    Public Property Fazenda() As String
        Get
            Return _Fazenda

        End Get
        Set(ByVal value As String)
            _Fazenda = value
        End Set
    End Property
    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
        End Set
    End Property
    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And Me.CodigoCliente.Length > 0 Then _Cliente = New Cliente(Me.CodigoCliente, Me.EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property
    Public Property Classificacao As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
        End Set
    End Property
    Public Property Qualitativo As String
        Get
            Select Case _Qualitativo
                Case "R"
                    Return "Relacionamento"
                Case "N"
                    Return "Negócio"
                Case "P"
                    Return "Preço"
                Case Else
                    Return ""
            End Select
        End Get
        Set(ByVal value As String)
            _Qualitativo = value
        End Set
    End Property
    Public Property MinimoVisita() As Integer
        Get
            Return _MinimoVisita
        End Get
        Set(ByVal value As Integer)
            _MinimoVisita = value
        End Set
    End Property

    'Public Property Visitas As ListVisita
    '    Get
    '        If _Visitas Is Nothing Then _Visitas = New ListVisita(Me, Me.CodigoCliente, Me.EndCliente)
    '        Return _Visitas

    '    End Get
    '    Set(ByVal value As ListVisita)
    '        _Visitas = value
    '    End Set
    'End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD Is Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = ""
        Select Case Me.IUD
            Case "I"
                strSQL = "insert into CrmxParametroVisita (Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id, Cliente_Id, EndCliente_Id, MinimoVisitas) " & vbCrLf & _
                         " values('" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & " ," & Me.Ano & "," & IIf(Me.Consolidado, "1", "0") & ", '" & Me.CodigoCliente & "'," & Me.EndCliente & "," & Me.MinimoVisita & " )"
                Sqls.Add(strSQL)
            Case "U"
                strSQL = "update CrmxParametroVisita " & vbCrLf & _
                         "   set MinimoVisitas = " & Me.MinimoVisita & vbCrLf & _
                         " where Empresa_Id     ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id  = " & Me.EndEmpresa & " " & vbCrLf & _
                         "   and Ano_id         = " & Me.Ano & vbCrLf & _
                         "   and Consolidado_id = " & IIf(Me.Consolidado, "1", "0") & vbCrLf & _
                         "   and Cliente_Id     = '" & Me.CodigoCliente & "'" & vbCrLf & _
                         "   and EndCliente_Id  = " & Me.EndCliente & ""
                Sqls.Add(strSQL)
            Case "D"
                strSQL = ""
                Sqls.Add(strSQL)
        End Select

    End Sub
#End Region

End Class
