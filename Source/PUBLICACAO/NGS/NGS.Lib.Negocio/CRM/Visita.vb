Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListVisita
    Inherits List(Of Visita)

#Region "Construtor"
    Public Sub New()

    End Sub

    '********  VISITAS AVULSAS / FORA DO CRM  **********
    Public Sub New(ByVal pDataInicial As DateTime, ByVal pDataFinal As DateTime, Optional ByVal pRegiao As String = "-1", Optional ByVal pMicroRegiao As Integer = 0)
        Dim Sql As String = ""

        If pRegiao = "" Then
            pRegiao = "-1"
        End If

        Sql = "SELECT CxV.NumeroVisita_Id as NumeroVisita," & vbCrLf & _
              "       CxV.Cliente as CodigoCliente," & vbCrLf & _
              "       CxV.EndCliente," & vbCrLf & _
              "       CxV.Representante as CodigoRepresentante," & vbCrLf & _
              "       CxV.EndRepresentante," & vbCrLf & _
              "       CxV.Data," & vbCrLf & _
              "       CxV.KmInicial," & vbCrLf & _
              "       CxV.KmFinal, " & vbCrLf & _
              "       Case Cpc.TipoCliente " & vbCrLf & _
              "          when 'C' then 'S'" & vbCrLf & _
              "          when 'P' then 'S'" & vbCrLf & _
              "          when 'E' then 'S'" & vbCrLf & _
              "          else 'N'" & vbCrLf & _
              "       End CRMCliente" & vbCrLf & _
              "  FROM Visita AS CxV" & vbCrLf & _
              "  LEFT JOIN CRMParametroXCliente CpC" & vbCrLf & _
              "    ON Cxv.Cliente = Cpc.Cliente_Id " & vbCrLf & _
              " INNER JOIN Clientes C" & vbCrLf & _
              "    on Cxv.Cliente    = C.Cliente_Id" & vbCrLf & _
              "   and CxV.EndCliente = C.Endereco_Id" & vbCrLf & _
              " where CxV.Data between '" & pDataInicial.ToString("yyyy-MM-dd") & "' and '" & pDataFinal.ToString("yyyy-MM-dd") & "'" & vbCrLf
        If pRegiao > -1 Then Sql &= "   and C.Regiao          = " & pRegiao & vbCrLf
        If pMicroRegiao > 0 Then Sql &= "   and c.MicroRegiao     = " & pMicroRegiao

        CarregaDadosNaLista(Sql)

    End Sub

    '********  CRM  **********
    'Public Sub New(ByVal pParametro As CRMxParametroVisita, ByVal pCliente As String, ByVal pEndCliente As Integer)
    '    Dim Sql As String = ""

    '    Sql = "SELECT CxV.NumeroVisita_Id AS NumeroVisita," & vbCrLf & _
    '          "       CxP.Empresa_Id AS CodigoEmpresa," & vbCrLf & _
    '          "       CxP.EndEmpresa_Id AS EndEmpresa," & vbCrLf & _
    '          "       CxP.Ano_Id AS Ano," & vbCrLf & _
    '          "       CxP.Consolidado_Id AS Consolidado," & vbCrLf & _
    '          "       CxP.Cliente_Id AS CodigoCliente," & vbCrLf & _
    '          "       CxP.EndCliente_Id AS EndCliente," & vbCrLf & _
    '          "       CxV.Representante AS CodigoRepresentante," & vbCrLf & _
    '          "       CxV.EndRepresentante," & vbCrLf & _
    '          "       CxV.Data," & vbCrLf & _
    '          "       CxV.KmInicial," & vbCrLf & _
    '          "       CxV.KmFinal" & vbCrLf & _
    '          "  FROM CRMXParametroVisita AS CxP" & vbCrLf & _
    '          " INNER JOIN CRMParametrosXVisitas AS PxV" & vbCrLf & _
    '          "    ON CxP.Empresa_Id     = PxV.Empresa_Id" & vbCrLf & _
    '          "   AND CxP.EndEmpresa_Id  = PxV.EndEmpresa_Id" & vbCrLf & _
    '          "   AND CxP.Ano_Id         = PxV.Ano_Id" & vbCrLf & _
    '          "   AND CxP.Consolidado_Id = PxV.Consolidado_Id" & vbCrLf & _
    '          "   AND CxP.Cliente_Id     = PxV.Cliente_Id" & vbCrLf & _
    '          "   AND CxP.EndCliente_Id  = PxV.EndCliente_Id" & vbCrLf & _
    '          " INNER JOIN Visita AS CxV " & vbCrLf & _
    '          "    ON CxV.NumeroVisita_Id = PxV.NumeroVisita_Id" & vbCrLf & _
    '          " WHERE Cxp.Empresa_Id     ='" & pParametro.CodigoEmpresa & "'" & vbCrLf & _
    '          "   AND Cxp.EndEmpresa_Id  = " & pParametro.EndEmpresa & vbCrLf & _
    '          "   AND CxP.Ano_Id         = " & pParametro.Ano & vbCrLf & _
    '          "   and CxP.Consolidado_Id = " & IIf(pParametro.Consolidado, "1", "0") & vbCrLf & _
    '          "   AND Cxp.Cliente_Id     ='" & pCliente & "'" & vbCrLf & _
    '          "   AND Cxp.EndCliente_Id  = " & pEndCliente & vbCrLf

    '    CarregaDadosNaLista(Sql, pParametro)

    'End Sub


    '********  CRM  **********

    'Public Sub New(ByVal pCRM As CRM, Optional ByVal pDataInicial As DateTime = Nothing, Optional ByVal pDataFinal As DateTime = Nothing)
    Public Sub New(ByVal pCRM As CRM)
        Dim Sql As String = ""

        Sql = "SELECT CxV.NumeroVisita_Id AS NumeroVisita," & vbCrLf & _
              "       CxP.Empresa_Id AS CodigoEmpresa," & vbCrLf & _
              "       CxP.EndEmpresa_Id AS EndEmpresa," & vbCrLf & _
              "       CxP.Ano_Id AS Ano," & vbCrLf & _
              "       CxP.Consolidado_Id AS Consolidado," & vbCrLf & _
              "       CxP.Cliente_Id AS CodigoCliente," & vbCrLf & _
              "       CxP.EndCliente_Id AS EndCliente," & vbCrLf & _
              "       CxV.Representante AS CodigoRepresentante," & vbCrLf & _
              "       CxV.EndRepresentante," & vbCrLf & _
              "       CxV.Data," & vbCrLf & _
              "       CxV.KmInicial," & vbCrLf & _
              "       CxV.KmFinal," & vbCrLf & _
              "       '' as CRMCliente " & vbCrLf & _
              "  FROM CRMXParametroVisita AS CxP" & vbCrLf & _
              " INNER JOIN Visita AS CxV " & vbCrLf & _
              "   ON CxP.Cliente_Id     = CxV.Cliente" & vbCrLf & _
              "   AND CxP.EndCliente_Id  = CxV.EndCliente" & vbCrLf & _
              " WHERE CxP.Empresa_Id     ='" & pCRM.CodigoEmpresa & "'" & vbCrLf & _
              "   AND CxP.EndEmpresa_Id  = " & pCRM.EndEmpresa & vbCrLf & _
              "   AND CxP.Ano_Id         = " & pCRM.Ano & vbCrLf & _
              "   and CxP.Consolidado_Id = " & IIf(pCRM.Consolidado, "1", "0") & vbCrLf
        'If IsDate(pDataInicial) And IsDate(pDataFinal) Then
        '    Sql = Sql + " and CxV.Data between '" & pDataInicial.ToString("yyyy-MM-dd") & "' and '" & pDataFinal.ToString("yyyy-MM-dd") & "'" & vbCrLf
        'End If

        CarregaDadosNaLista(Sql)

    End Sub

#End Region

#Region "Methods"
    Private Sub CarregaDadosNaLista(ByVal Sql As String, Optional ByVal pParametro As CRMxParametroVisita = Nothing)
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(Sql, "Visitas")

        If ds.Tables.Count > 0 Then
            For Each row In ds.Tables(0).Rows
                Dim vis As New Visita(pParametro)
                vis.NumeroVisita = row("NumeroVisita")
                vis.CodigoRepresentante = row("CodigoRepresentante")
                vis.EndRepresentante = row("EndRepresentante")
                vis.CodigoCliente = row("CodigoCliente")
                vis.EndCliente = row("EndCliente")
                vis.Data = row("Data")
                vis.KmInicial = row("KmInicial")
                vis.KmFinal = row("KmFinal")
                vis.CRMCliente = row("CRMCliente") = "S"
                Me.Add(vis)
            Next
        End If
    End Sub

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
        For Each item As Visita In Me
            item.SalvarSql(Sqls)
        Next
    End Sub

    Public Function ListRepresentantes() As Object
        Return (From p In Me Order By p.NomeRepresentante Select p.CodigoRepresentante, p.EndRepresentante, p.NomeRepresentante).Distinct()
    End Function

    Public Function ListVisitasRealizadas(ByVal pCliente As String, ByVal pEndCliente As Integer) As Object
        Return (From p In Me Where p.CodigoCliente = pCliente And p.EndCliente = pEndCliente Order By p.NomeCliente Select p)
    End Function

    Public Function ListClientes(Optional ByVal pCodigoRepresentante As String = "", Optional ByVal pCRMCliente As String = "") As Object
        If Not String.IsNullOrWhiteSpace(pCodigoRepresentante) And Not String.IsNullOrWhiteSpace(pCRMCliente) Then
            Return (From c In Me Where c.CodigoRepresentante = pCodigoRepresentante And c.CRMCliente = IIf(pCRMCliente = "S", True, False) Order By c.NomeRepresentante, c.NomeCliente Select id = c.CodigoCliente & "-" & c.EndCliente.ToString, NomeCliente = c.Cliente.Nome, Fazenda = c.Cliente.Complemento, c.CodigoRepresentante, c.EndRepresentante, c.NomeRepresentante).Distinct
        ElseIf Not String.IsNullOrWhiteSpace(pCodigoRepresentante) Then
            Return (From c In Me Where c.CodigoRepresentante = pCodigoRepresentante Order By c.NomeRepresentante, c.NomeCliente Select id = c.CodigoCliente & "-" & c.EndCliente.ToString, NomeCliente = c.Cliente.Nome, Fazenda = c.Cliente.Complemento, c.CodigoRepresentante, c.EndRepresentante, c.NomeRepresentante).Distinct
        ElseIf Not String.IsNullOrWhiteSpace(pCRMCliente) Then
            Return (From c In Me Where c.CRMCliente = IIf(pCRMCliente = "S", True, False) Order By c.NomeRepresentante, c.NomeCliente Select id = c.CodigoCliente & "-" & c.EndCliente.ToString, NomeCliente = c.Cliente.Nome, Fazenda = c.Cliente.Complemento, c.CodigoRepresentante, c.EndRepresentante, c.NomeRepresentante).Distinct
        Else
            Return (From c In Me Order By c.NomeRepresentante, c.NomeCliente Select id = c.CodigoCliente & "-" & c.EndCliente.ToString, NomeCliente = c.Cliente.Nome, Fazenda = c.Cliente.Complemento, c.CodigoRepresentante, c.EndRepresentante, c.NomeRepresentante).Distinct
        End If
    End Function
#End Region

End Class


<Serializable()> _
Public Class Visita
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pParametro As CRMxParametroVisita)
        _Parametro = pParametro
        'Me.ListVisitaMotivo = New ListVisitaMotivo(Me.NumeroVisita)
        'Me.ListVisitaAmeaca = New ListVisitaAmeaca(Me.NumeroVisita)
        'Me.ListVisitaProdutividade = New ListVisitaProdutividade(Me.NumeroVisita)
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""

    Private _NumeroVisita As Integer

    '**** Ligação com o CRM *****************'
    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    'Private _Ano As Integer
    'Private _Consolidado As Integer

    '**** Representante *****************'
    Private _Representante As Cliente
    Private _CodigoRepresentante As String = ""
    Private _EndRepresentante As String = ""
    Private _NomeRepresentante As String = ""

    '**** Cliente *****************'
    Private _Cliente As Cliente
    Private _CodigoCliente As String = ""
    Private _EndCliente As Integer
    Private _NomeCliente As String = ""

    Private _CRMCliente As Boolean

    '****  *****************'
    Private _Data As Date
    Private _KmInicial As Integer
    Private _KmFinal As Integer

    '********* Motivo ***************
    Private _VisitaMotivo As ListVisitaMotivo
    Private _Motivos As String = ""

    '********* Ameaca ***************
    Private _VisitaAmeaca As ListVisitaAmeaca

    '********* Produtividade ***************
    Private _VisitaProdutividade As ListVisitaProdutividade

    '********* Parametro Visita ************
    Private _Parametro As CRMxParametroVisita

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

    Public Property NumeroVisita()
        Get
            Return _NumeroVisita
        End Get
        Set(ByVal value)
            _NumeroVisita = value
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
    'Public Property Ano() As Integer
    '    Get
    '        Return _Ano
    '    End Get
    '    Set(ByVal value As Integer)
    '        _Ano = value
    '    End Set
    'End Property
    'Public Property Consolidado As Integer
    '    Get
    '        Return _Consolidado
    '    End Get
    '    Set(ByVal value As Integer)
    '        _Consolidado = value
    '    End Set
    'End Property

    Public Property Representante() As Cliente
        Get
            If _Representante Is Nothing And Me.CodigoRepresentante.Length > 0 Then _Representante = New Cliente(Me.CodigoRepresentante, Me.EndRepresentante)
            Return _Representante
        End Get
        Set(ByVal value As Cliente)
            _Representante = value
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

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And Me.CodigoCliente.Length > 0 Then _Cliente = New Cliente(Me.CodigoCliente, Me.EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
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
    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property CRMCliente() As Boolean
        Get
            Return _CRMCliente
        End Get
        Set(ByVal value As Boolean)
            _CRMCliente = value
        End Set
    End Property


    Public Property Data As Date
        Get
            Return _Data
        End Get
        Set(ByVal value As Date)
            _Data = value
        End Set
    End Property
    Public Property KmInicial As Integer
        Get
            Return _KmInicial
        End Get
        Set(ByVal value As Integer)
            _KmInicial = value
        End Set
    End Property
    Public Property KmFinal As Integer
        Get
            Return _KmFinal
        End Get
        Set(ByVal value As Integer)
            _KmFinal = value
        End Set
    End Property

    Public Property ListVisitaMotivo As ListVisitaMotivo
        Get
            If _VisitaMotivo Is Nothing Then _VisitaMotivo = New ListVisitaMotivo(Me.NumeroVisita)
            Return _VisitaMotivo

        End Get
        Set(ByVal value As ListVisitaMotivo)
            _VisitaMotivo = value
        End Set
    End Property

    Public Property ListVisitaAmeaca As ListVisitaAmeaca
        Get
            If _VisitaAmeaca Is Nothing Then _VisitaAmeaca = New ListVisitaAmeaca(Me.NumeroVisita)
            Return _VisitaAmeaca

        End Get
        Set(ByVal value As ListVisitaAmeaca)
            _VisitaAmeaca = value
        End Set
    End Property

    Public Property ListVisitaProdutividade As ListVisitaProdutividade
        Get
            If _VisitaProdutividade Is Nothing Then _VisitaProdutividade = New ListVisitaProdutividade(Me.NumeroVisita)
            Return _VisitaProdutividade

        End Get
        Set(ByVal value As ListVisitaProdutividade)
            _VisitaProdutividade = value
        End Set
    End Property

    Public ReadOnly Property KmRodado As Integer
        Get
            Dim Rodado = _KmFinal - _KmInicial
            If Rodado > 0 Then
                Return Rodado
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property Finalidade As String
        Get
            If String.IsNullOrWhiteSpace(_Motivos) Then
                For Each objVisitaMotivo In ListVisitaMotivo
                    _Motivos &= IIf(String.IsNullOrWhiteSpace(_Motivos), "", " - ") & objVisitaMotivo.Motivo.Descricao
                Next
            End If
            Return _Motivos
        End Get
    End Property

    Public Property Parametro As CRMxParametroVisita
        Get
            Return _Parametro
        End Get
        Set(ByVal value As CRMxParametroVisita)
            _Parametro = value
        End Set
    End Property

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
                Dim Banco As New AcessaBanco
                'Dim N As New Numerador(Me.CodigoEmpresa, Me.EndEmpresa, 4)
                'Me.NumeroVisita = N.Sequencia + 1
                Dim tSql As String = "SELECT MAX(NumeroVisita_Id) NumeroVisita FROM Visita"
                Dim ds As DataSet = Banco.ConsultaDataSet(tSql, "Visita")
                Me.NumeroVisita = ds.Tables(0).Rows(0)("NumeroVisita") + 1

                strSQL = "INSERT INTO Visita (NumeroVisita_Id, Cliente, EndCliente, Representante, EndRepresentante,  Data, KmInicial, KmFinal) " & vbCrLf & _
                         " VALUES(" & Me.NumeroVisita & ",'" & Me.CodigoCliente & "'," & Me.EndCliente & ", '" & Me.CodigoRepresentante & "'," &
                                      Me.EndRepresentante & ",'" & Me.Data.ToString("yyyy-MM-dd") & "'," & Me.KmInicial & ", " & Me.KmFinal & ")"
                Sqls.Add(strSQL)

                'Dim objParametroVisita As Object = Me.Parametro.ParametrosVisita.Find(Function(s) s.CodigoCliente = Me.CodigoCliente And s.EndCliente = Me.EndCliente)
                'If Not objParametroVisita Is Nothing Then
                '    strSQL = " INSERT INTO CRMParametrosXVisitas (Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id, Cliente_Id, EndCliente_Id, NumeroVisita_Id)" & vbCrLf & _
                '             " VALUES('" & objParametroVisita.CodigoEmpresa & "'," & objParametroVisita.EndEmpresa & "," & objParametroVisita.Ano & "," & IIf(objParametroVisita.Consolidado, "1", "0") & ",'" & Me.CodigoCliente & "'," & Me.EndCliente & "," & Me.NumeroVisita & ") "
                '    Sqls.Add(strSQL)
                'End If

                'Sqls.Add(N.IncrementarNumeradorSql)
            Case "U"
                strSQL = "update Visita " & vbCrLf & _
                         "   set Data            = '" & Me.Data.ToString("yyyy-MM-dd") & "', " & vbCrLf & _
                         "       KmInicial       = " & Me.KmInicial & ", " & vbCrLf & _
                         "       KmFinal         = " & Me.KmFinal & vbCrLf & _
                         " where NumeroVisita_Id =" & Me.NumeroVisita
                Sqls.Add(strSQL)
            Case "D"
                strSQL = ""
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class
