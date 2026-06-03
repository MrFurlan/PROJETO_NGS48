Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCRM
    Inherits List(Of CRM)

    Public Sub New(ByVal pAno As Integer)
        Dim sql As String = ""
        sql = "SELECT Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id," & vbCrLf & _
              "       CRMPercentualCorte, CRMValorCorte," & vbCrLf & _
              "       MBVPeso, MBVPercentualMenor, MBVPercentualMaior, MBVPontuacaoMenor," & vbCrLf & _
              "       MBVPontuacaoEntre, MBVPontuacaoMaior," & vbCrLf & _
              "       SOCPeso, SOCPesoVenda, SOCPesoCompra," & vbCrLf & _
              "       SOCPercentualMenor, SOCPercentualMaior," & vbCrLf & _
              "       SOCPontuacaoMenor, SOCPontuacaoEntre, SOCPontuacaoMaior," & vbCrLf & _
              "       RatingPeso, RatingPontuacaoC, RatingPontuacaoB, RatingPontuacaoA," & vbCrLf & _
              "       Definicao, Mercado, ProdutosVenda, ProdutosCompra," & vbCrLf & _
              "       ConsideraCompra, EmpresaCompra, EndEmpresaCompra " & vbCrLf & _
              "  FROM CRMParametro" & vbCrLf & _
              " Where Ano_Id         = " & pAno & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CRMParametro")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        For Each row In ds.Tables(0).Rows
            Dim objCRM As New CRM()
            objCRM.CodigoEmpresa = row("Empresa_Id")
            objCRM.EndEmpresa = row("EndEmpresa_Id")
            objCRM.CodigoEmpresaCompra = row("EmpresaCompra")
            objCRM.EndEmpresaCompra = row("EndEmpresaCompra")

            objCRM.Ano = row("Ano_Id")
            objCRM.Consolidado = row("Consolidado_Id")
            objCRM.CRMPercentualCorte = row("CRMPercentualCorte")
            objCRM.CRMValorCorte = row("CRMValorCorte")
            objCRM.MBVPeso = row("MBVPeso")
            objCRM.MBVPercentualMenor = row("MBVPercentualMenor")
            objCRM.MBVPercentualMaior = row("MBVPercentualMaior")
            objCRM.SOCPeso = row("SOCPeso")
            objCRM.SOCPesoVenda = row("SOCPesoVenda")
            objCRM.SOCPesoCompra = row("SOCPesoCompra")
            objCRM.SOCPercentualMenor = row("SOCPercentualMenor")
            objCRM.SOCPercentualMaior = row("SOCPercentualMaior")
            objCRM.RatingPeso = row("RatingPeso")
            objCRM.Definicao = row("Definicao")
            objCRM.Mercado = row("Mercado")
            objCRM.ConsideraCompra = row("ConsideraCompra")
            objCRM.ProdutosVenda = row("ProdutosVenda")
            objCRM.ProdutosCompra = row("ProdutosCompra")
            Me.Add(objCRM)
        Next
    End Sub

    'Public ReadOnly Property getCRMNaoConsolidado As Object
    '    Get
    '        Return (From p In Me Where p.Consolidado = False Select p.Ano, p.CodigoEmpresa, p.EndEmpresa, p.NomeEmpresa, p.EmpresaCompra, p.Consolidado)
    '    End Get
    'End Property

End Class

<Serializable()> _
Public Class CRM
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pAno As Integer, ByVal pConsolidado As Boolean)
        Dim sql As String = ""
        sql = "SELECT Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id," & vbCrLf & _
              "       CRMPercentualCorte, CRMValorCorte," & vbCrLf & _
              "       MBVPeso, MBVPercentualMenor, MBVPercentualMaior, MBVPontuacaoMenor," & vbCrLf & _
              "       MBVPontuacaoEntre, MBVPontuacaoMaior," & vbCrLf & _
              "       SOCPeso, SOCPesoVenda, SOCPesoCompra," & vbCrLf & _
              "       SOCPercentualMenor, SOCPercentualMaior," & vbCrLf & _
              "       SOCPontuacaoMenor, SOCPontuacaoEntre, SOCPontuacaoMaior," & vbCrLf & _
              "       RatingPeso, RatingPontuacaoC, RatingPontuacaoB, RatingPontuacaoA," & vbCrLf & _
              "       Definicao, Mercado, ProdutosVenda, ProdutosCompra," & vbCrLf & _
              "       ConsideraCompra, EmpresaCompra, EndEmpresaCompra " & vbCrLf & _
              "  FROM CRMParametro" & vbCrLf & _
              " Where Empresa_Id     ='" & pCodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id  = " & pEndEmpresa & vbCrLf & _
              "   and Ano_Id         = " & pAno & vbCrLf & _
              "   and Consolidado_Id = " & IIf(pConsolidado, "1", "0")
        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "CRMParametro")

        If ds.Tables(0).Rows.Count = 0 Then
            Me.CodigoEmpresa = pCodigoEmpresa
            Me.EndEmpresa = pEndEmpresa
            Me.Ano = pAno
            Me.Consolidado = pConsolidado
            Exit Sub
        End If

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.IUD = "U"
        Me.CodigoEmpresa = row("Empresa_Id")
        Me.EndEmpresa = row("EndEmpresa_Id")

        Me.CodigoEmpresaCompra = row("EmpresaCompra")
        Me.EndEmpresaCompra = row("EndEmpresaCompra")

        Me.Ano = row("Ano_Id")
        Me.Consolidado = row("Consolidado_Id")
        Me.CRMPercentualCorte = row("CRMPercentualCorte")
        Me.CRMValorCorte = row("CRMValorCorte")
        Me.MBVPeso = row("MBVPeso")
        Me.MBVPercentualMenor = row("MBVPercentualMenor")
        Me.MBVPercentualMaior = row("MBVPercentualMaior")
        Me.SOCPeso = row("SOCPeso")
        Me.SOCPesoVenda = row("SOCPesoVenda")
        Me.SOCPesoCompra = row("SOCPesoCompra")
        Me.SOCPercentualMenor = row("SOCPercentualMenor")
        Me.SOCPercentualMaior = row("SOCPercentualMaior")
        Me.RatingPeso = row("RatingPeso")

        Me.Definicao = row("Definicao")
        Me.Mercado = row("Mercado")
        Me.ConsideraCompra = row("ConsideraCompra")
        Me.ProdutosVenda = row("ProdutosVenda")
        Me.ProdutosCompra = row("ProdutosCompra")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _Consolidado As Boolean
    Private _Ano As Integer

    Private _CodigoEmpresaCompra As String = ""
    Private _EndEmpresaCompra As Integer
    Private _EmpresaCompra As Cliente

    Private _CRMPercentualCorte As Decimal = 80
    Private _CRMValorCorte As Decimal = 0

    Private _MBVPeso As Decimal = 34
    Private _MBVPercentualMenor As Decimal = 10
    Private _MBVPercentualMaior As Decimal = 15

    Private _SOCPeso As Decimal = 33
    Private _SOCPesoVenda As Decimal = 70
    Private _SOCPesoCompra As Decimal = 30
    Private _SOCPercentualMenor As Decimal = 50
    Private _SOCPercentualMaior As Decimal = 70

    Private _RatingPeso As Decimal = 33

    Private _Definicao As Integer
    Private _Mercado As Char
    Private _ConsideraCompra As Boolean
    Private _ProdutosVenda As String = ""
    Private _ProdutosCompra As String = ""

    '*********** CRM Clientes ****************
    Private _CrmClientes As ListCRMxCliente
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
    '*********************************************************************************
    '*************    Empresa Venda / Consolidado Venda e Compra    ******************
    '*********************************************************************************
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public ReadOnly Property NomeEmpresa As String
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EndEmpresa)
            Return _Empresa.Nome & " " & _Empresa.Cidade & "-" & _Empresa.CodigoEstado
        End Get
    End Property

    '****************************************************
    '*************    Empresa Compra   ******************
    '****************************************************
    Public Property CodigoEmpresaCompra() As String
        Get
            Return _CodigoEmpresaCompra
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaCompra = value
            _EmpresaCompra = Nothing
        End Set
    End Property

    Public Property EndEmpresaCompra() As Integer
        Get
            Return _EndEmpresaCompra
        End Get
        Set(ByVal value As Integer)
            _EndEmpresaCompra = value
            _EmpresaCompra = Nothing
        End Set
    End Property

    Public ReadOnly Property NomeEmpresaCompra As String
        Get
            If _EmpresaCompra Is Nothing And _CodigoEmpresaCompra.Length > 0 Then _EmpresaCompra = New Cliente(Me.CodigoEmpresaCompra, Me.EndEmpresaCompra)
            Return _EmpresaCompra.Nome
        End Get
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property Consolidado() As Boolean
        Get
            Return _Consolidado
        End Get
        Set(ByVal value As Boolean)
            _Consolidado = value
        End Set
    End Property

    '**********************************************
    '***************    CRM    ********************
    '**********************************************
    Public Property CRMPercentualCorte() As Decimal
        Get
            Return _CRMPercentualCorte
        End Get
        Set(ByVal value As Decimal)
            _CRMPercentualCorte = value
        End Set
    End Property

    Public Property CRMValorCorte() As Decimal
        Get
            Return _CRMValorCorte
        End Get
        Set(ByVal value As Decimal)
            _CRMValorCorte = value
        End Set
    End Property

    '**********************************************
    '*********** Margem Bruta Venda ***************
    '**********************************************
    Public Property MBVPeso() As Decimal
        Get
            Return _MBVPeso
        End Get
        Set(ByVal value As Decimal)
            _MBVPeso = value
        End Set
    End Property

    Public Property MBVPercentualMenor() As Decimal
        Get
            Return _MBVPercentualMenor
        End Get
        Set(ByVal value As Decimal)
            _MBVPercentualMenor = value
        End Set
    End Property

    Public ReadOnly Property MBVPercentualEntreDe() As Decimal
        Get
            Return _MBVPercentualMenor
        End Get
    End Property

    Public ReadOnly Property MBVPercentualEntreAte() As Decimal
        Get
            Return _MBVPercentualMaior
        End Get
    End Property

    Public Property MBVPercentualMaior() As Decimal
        Get
            Return _MBVPercentualMaior
        End Get
        Set(ByVal value As Decimal)
            _MBVPercentualMaior = value
        End Set
    End Property

    Public ReadOnly Property MBVPontuacaoMenor() As Decimal
        Get
            Return _MBVPeso * 9 / 100
        End Get
    End Property

    Public ReadOnly Property MBVPontuacaoEntre() As Decimal
        Get
            Return _MBVPeso * 18 / 100
        End Get
    End Property

    Public ReadOnly Property MBVPontuacaoMaior() As Decimal
        Get
            Return _MBVPeso * 27 / 100
        End Get
    End Property

    '**********************************************
    '************* Participacao SOC ***************
    '**********************************************
    Public Property SOCPeso() As Decimal
        Get
            Return _SOCPeso
        End Get
        Set(ByVal value As Decimal)
            _SOCPeso = value
        End Set
    End Property

    Public Property SOCPesoVenda() As Decimal
        Get
            Return _SOCPesoVenda
        End Get
        Set(ByVal value As Decimal)
            _SOCPesoVenda = value
        End Set
    End Property

    Public Property SOCPesoCompra() As Decimal
        Get
            Return _SOCPesoCompra
        End Get
        Set(ByVal value As Decimal)
            _SOCPesoCompra = value
        End Set
    End Property

    Public Property SOCPercentualMenor() As Decimal
        Get
            Return _SOCPercentualMenor
        End Get
        Set(ByVal value As Decimal)
            _SOCPercentualMenor = value
        End Set
    End Property

    Public ReadOnly Property SOCPercentualEntreDe() As Decimal
        Get
            Return _SOCPercentualMenor
        End Get
    End Property

    Public ReadOnly Property SOCPercentualEntreAte() As Decimal
        Get
            Return _SOCPercentualMaior
        End Get
    End Property

    Public Property SOCPercentualMaior() As Decimal
        Get
            Return _SOCPercentualMaior
        End Get
        Set(ByVal value As Decimal)
            _SOCPercentualMaior = value
        End Set
    End Property

    Public ReadOnly Property SOCPontuacaoMenor() As Decimal
        Get
            Return _SOCPeso * 9 / 100
        End Get
    End Property

    Public ReadOnly Property SOCPontuacaoEntre() As Decimal
        Get
            Return _SOCPeso * 18 / 100
        End Get
    End Property

    Public ReadOnly Property SOCPontuacaoMaior() As Decimal
        Get
            Return _SOCPeso * 27 / 100
        End Get
    End Property

    '**********************************************
    '*************** Rating Credito ***************
    '**********************************************
    Public Property RatingPeso() As Decimal
        Get
            Return _RatingPeso
        End Get
        Set(ByVal value As Decimal)
            _RatingPeso = value
        End Set
    End Property

    Public ReadOnly Property RatingPontuacaoC() As Decimal
        Get
            Return _RatingPeso * 9 / 100
        End Get
    End Property

    Public ReadOnly Property RatingPontuacaoB() As Decimal
        Get
            Return _RatingPeso * 18 / 100
        End Get
    End Property

    Public ReadOnly Property RatingPontuacaoA() As Decimal
        Get
            Return _RatingPeso * 27 / 100
        End Get
    End Property

    '*******************************************************
    '*************** Parametros P/ CURVA ABC ***************
    '*******************************************************
    Public Property Definicao() As Integer
        Get
            Return _Definicao
        End Get
        Set(ByVal value As Integer)
            _Definicao = value
        End Set
    End Property

    Public Property Mercado() As Char
        Get
            Return _Mercado
        End Get
        Set(ByVal value As Char)
            _Mercado = value
        End Set
    End Property

    Public Property ConsideraCompra() As Boolean
        Get
            Return _ConsideraCompra
        End Get
        Set(ByVal value As Boolean)
            _ConsideraCompra = value
        End Set
    End Property

    Public Property ProdutosVenda() As String
        Get
            Return _ProdutosVenda
        End Get
        Set(ByVal value As String)
            _ProdutosVenda = value
        End Set
    End Property

    Public Property ProdutosCompra() As String
        Get
            Return _ProdutosCompra
        End Get
        Set(ByVal value As String)
            _ProdutosCompra = value
        End Set
    End Property

    '*********** CRM Clientes ****************
    Public Property CrmClientes() As ListCRMxCliente
        Get
            If _CrmClientes Is Nothing Then _CrmClientes = New ListCRMxCliente(Me)
            Return _CrmClientes
        End Get
        Set(ByVal value As ListCRMxCliente)
            _CrmClientes = value
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
                strSQL &= "insert into CRMParametro(Empresa_Id, EndEmpresa_Id, Ano_Id, Consolidado_Id, " & vbCrLf & _
                          "                         CRMPercentualCorte, CRMValorCorte," & vbCrLf & _
                          "                         MBVPeso, MBVPercentualMenor, MBVPercentualMaior," & vbCrLf & _
                          "                         MBVPontuacaoMenor, MBVPontuacaoEntre, MBVPontuacaoMaior," & vbCrLf & _
                          "                         SOCPeso, SOCPesoVenda, SOCPesoCompra, SOCPercentualMenor, SOCPercentualMaior," & vbCrLf & _
                          "                         SOCPontuacaoMenor, SOCPontuacaoEntre, SOCPontuacaoMaior," & vbCrLf & _
                          "                         RatingPeso, RatingPontuacaoC, RatingPontuacaoB, RatingPontuacaoA," & vbCrLf & _
                          "                         Definicao, Mercado, ConsideraCompra, ProdutosVenda, ProdutosCompra, EmpresaCompra, EndEmpresaCompra)" & vbCrLf & _
                          "  values('" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & "," & Me.Ano & "," & IIf(Me.Consolidado, "1", "0") & vbCrLf & _
                          "," & Str(Me.CRMPercentualCorte) & "," & Str(Me.CRMValorCorte) & vbCrLf & _
                          "," & Str(Me.MBVPeso) & "," & Str(Me.MBVPercentualMenor) & "," & Str(Me.MBVPercentualMaior) & vbCrLf & _
                          "," & Str(Me.MBVPontuacaoMenor) & "," & Str(Me.MBVPontuacaoEntre) & "," & Str(Me.MBVPontuacaoMaior) & vbCrLf & _
                          "," & Str(Me.SOCPeso) & "," & Str(Me.SOCPesoVenda) & "," & Str(Me.SOCPesoCompra) & "," & Str(Me.SOCPercentualMenor) & "," & Str(Me.SOCPercentualMaior) & vbCrLf & _
                          "," & Str(Me.SOCPontuacaoMenor) & "," & Str(Me.SOCPontuacaoEntre) & "," & Str(Me.SOCPontuacaoMaior) & vbCrLf & _
                          "," & Str(Me.RatingPeso) & "," & Str(Me.RatingPontuacaoC) & "," & Str(Me.RatingPontuacaoB) & "," & Str(Me.RatingPontuacaoA) & vbCrLf & _
                          "," & Me.Definicao & ",'" & Me.Mercado & "'," & IIf(Me.ConsideraCompra, 1, 0) & ",'" & Me.ProdutosVenda & "','" & Me.ProdutosCompra & "','" & Me.CodigoEmpresaCompra & "'," & EndEmpresaCompra & ")"
                Sqls.Add(strSQL)
                CrmClientes.SalvarSql(Sqls)
            Case "U"
                strSQL = " Update CRMParametro set " & vbCrLf & _
                         "    CRMPercentualCorte = " & Str(Me.CRMPercentualCorte) & vbCrLf & _
                         "   ,CRMValorCorte      = " & Str(Me.CRMValorCorte) & vbCrLf & _
                         "   ,MBVPeso            = " & Str(Me.MBVPeso) & vbCrLf & _
                         "   ,MBVPercentualMenor = " & Str(Me.MBVPercentualMenor) & vbCrLf & _
                         "   ,MBVPercentualMaior = " & Str(Me.MBVPercentualMaior) & vbCrLf & _
                         "   ,MBVPontuacaoMenor  = " & Str(Me.MBVPontuacaoMenor) & vbCrLf & _
                         "   ,MBVPontuacaoEntre  = " & Str(Me.MBVPontuacaoEntre) & vbCrLf & _
                         "   ,MBVPontuacaoMaior  = " & Str(Me.MBVPontuacaoMaior) & vbCrLf & _
                         "   ,SOCPeso            = " & Str(Me.SOCPeso) & vbCrLf & _
                         "   ,SOCPesoVenda       = " & Str(Me.SOCPesoVenda) & vbCrLf & _
                         "   ,SOCPesoCompra      = " & Str(Me.SOCPesoCompra) & vbCrLf & _
                         "   ,SOCPercentualMenor = " & Str(Me.SOCPercentualMenor) & vbCrLf & _
                         "   ,SOCPercentualMaior = " & Str(Me.SOCPercentualMaior) & vbCrLf & _
                         "   ,SOCPontuacaoMenor  = " & Str(Me.SOCPontuacaoMenor) & vbCrLf & _
                         "   ,SOCPontuacaoEntre  = " & Str(Me.SOCPontuacaoEntre) & vbCrLf & _
                         "   ,SOCPontuacaoMaior  = " & Str(Me.SOCPontuacaoMaior) & vbCrLf & _
                         "   ,RatingPeso         = " & Str(Me.RatingPeso) & vbCrLf & _
                         "   ,RatingPontuacaoC   = " & Str(Me.RatingPontuacaoC) & vbCrLf & _
                         "   ,RatingPontuacaoB   = " & Str(Me.RatingPontuacaoB) & vbCrLf & _
                         "   ,RatingPontuacaoA   = " & Str(Me.RatingPontuacaoA) & vbCrLf & _
                         "   ,Definicao          = " & Me.Definicao & vbCrLf & _
                         "   ,Mercado            ='" & Me.Mercado & "'" & vbCrLf & _
                         "   ,ConsideraCompra    = " & IIf(Me.ConsideraCompra, 1, 0) & vbCrLf & _
                         "   ,EmpresaCompra      ='" & Me.CodigoEmpresaCompra & "'" & vbCrLf & _
                         "   ,EndEmpresaCompra   = " & Me.EndEmpresaCompra & vbCrLf & _
                         "   ,ProdutosVenda      ='" & Me.ProdutosVenda & "'" & vbCrLf & _
                         "   ,ProdutosCompra     ='" & Me.ProdutosCompra & "'" & vbCrLf & _
                         " Where Empresa_Id     ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   and EndEmpresa_Id  = " & Me.EndEmpresa & vbCrLf & _
                         "   and Ano_Id         = " & Me.Ano & vbCrLf & _
                         "   and Consolidado_Id = " & IIf(Me.Consolidado, "1", "0") & vbCrLf
                Sqls.Add(strSQL)
                CrmClientes.SalvarSql(Sqls)
            Case "D"
                CrmClientes.SalvarSql(Sqls)
                strSQL = " Delete CRMParametro" & vbCrLf & _
                         "  Where Empresa_Id     ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "    and EndEmpresa_Id  = " & Me.EndEmpresa & vbCrLf & _
                         "    and Ano_Id         = " & Me.Ano & vbCrLf & _
                         "    and Consolidado_Id = " & IIf(Me.Consolidado, "1", "0") & vbCrLf
                Sqls.Add(strSQL)
        End Select
    End Sub
#End Region

End Class

