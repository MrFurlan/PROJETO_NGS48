Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListEncargo
    Inherits List(Of Encargo)

#Region "Construtor"
    Public Sub New(Optional ByVal Carregar As Boolean = False, Optional ByVal pWhere As String = "")
        If Not Carregar Then Exit Sub

        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Encargo_id, Descricao, isnull(ContaDebito,'') as ContaDebito,  isnull(TipoDePessoa,0) as TipoDePessoa, isnull(ContaCredito,'') as ContaCredito, BaseCalculo, Aliquota," & vbCrLf & _
              "       OperacaoXEncargo, Etapa, Atualizacao, isnull(GravaCentroDeCusto,0) as GravaCentroDeCusto, isnull(ImprimirNFE,0) as ImprimirNFE, isnull(VerificaEmpresa,0) as VerificaEmpresa," & vbCrLf & _
              "       isnull(Operador,'') as Operador, ISNULL(EncargoAgrupador,'') AS EncargoAgrupador, ValorOuPeso, isnull(PodeSofreRetencao,0) as PodeSofreRetencao, isnull(TipoDePessoaRetencao,0) as TipoDePessoaRetencao " & vbCrLf & _
              "  FROM Encargos" & vbCrLf & _
              IIf(pWhere.Length > 0, "Where " & pWhere, pWhere)

        Try
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Encargo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim Enc As New Encargo()
                Enc.Codigo = row("Encargo_id")
                Enc.Descricao = row("Descricao")
                Enc.TipoPessoa = row("TipoDePessoa")
                Enc.ContaDebito = row("ContaDebito")
                Enc.ContaCredito = row("ContaCredito")
                Enc.BaseCalculo = row("BaseCalculo")
                Enc.Aliquota = row("Aliquota")
                Enc.OperacaoXEncargo = row("OperacaoXEncargo") = "S"
                Enc.PodeSofreRetencao = row("PodeSofreRetencao")
                Enc.TipoPessoaRetencao = row("TipoDePessoaRetencao")
                Enc.GravaCentroDeCusto = row("GravaCentroDeCusto")
                Enc.ImprimirNFE = row("ImprimirNFE")
                Enc.Operador = row("Operador")
                If IsDBNull(row("ValorOuPeso")) Then
                    Enc.ValorOuPeso = -1
                Else
                    Enc.ValorOuPeso = Convert.ToInt16(row("ValorOuPeso"))
                End If
                Enc.EncargoAgrupador = row("EncargoAgrupador")
                Enc.VerificaEmpresa = row("VerificaEmpresa")

                If Enc.VerificaEmpresa Then Me.DescEncargoEmpresa &= IIf(Me.DescEncargoEmpresa.Length = 0, "", ", ") & Enc.Codigo

                Me.Add(Enc)
            Next
        Catch ex As Exception
            Throw ex
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _DescEncargoEmpresa As String = ""
#End Region

#Region "Property"
    Public Property DescEncargoEmpresa As String
        Get
            Return _DescEncargoEmpresa
        End Get
        Set(value As String)
            _DescEncargoEmpresa = value
        End Set
    End Property
#End Region

End Class

Public Class Encargo
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEncargo As String)
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT Encargo_id, Descricao, isnull(TipoDePessoa,0) as TipoDePessoa, isnull(ContaDebito,'') as ContaDebito,  isnull(ContaCredito,'') as ContaCredito, BaseCalculo, Aliquota," & vbCrLf & _
              "       OperacaoXEncargo, Etapa, Atualizacao, isnull(GravaCentroDeCusto,0) as GravaCentroDeCusto, isnull(ImprimirNFE,0) as ImprimirNFE, isnull(VerificaEmpresa,0) as VerificaEmpresa," & vbCrLf & _
              "       isnull(Operador,'') as Operador, ISNULL(EncargoAgrupador,'') AS EncargoAgrupador, ValorOuPeso, isnull(PodeSofreRetencao,0) as PodeSofreRetencao, isnull(TipoDePessoaRetencao,0) as TipoDePessoaRetencao" & vbCrLf & _
              "  FROM Encargos" & vbCrLf & _
              " Where Encargo_Id = '" & pEncargo & "'"
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Encargo")

        For Each row As DataRow In ds.Tables(0).Rows
            Me.IUD = "U"
            Me.Codigo = row("Encargo_id")
            Me.Descricao = row("Descricao")
            Me.TipoPessoa = row("TipoDePessoa")
            Me.ContaDebito = row("ContaDebito")
            Me.ContaCredito = row("ContaCredito")
            Me.BaseCalculo = row("BaseCalculo")
            Me.Aliquota = row("Aliquota")
            Me.OperacaoXEncargo = row("OperacaoXEncargo") = "S"
            Me.PodeSofreRetencao = row("PodeSofreRetencao")
            Me.TipoPessoaRetencao = row("TipoDePessoaRetencao")
            Me.Atualizacao = row("Atualizacao")
            Me.Etapa = row("Etapa")
            Me.GravaCentroDeCusto = row("GravaCentroDeCusto")
            Me.ImprimirNFE = row("ImprimirNFE")
            Me.Operador = row("Operador")
            If IsDBNull(row("ValorOuPeso")) Then
                Me.ValorOuPeso = -1
            Else
                Me.ValorOuPeso = Convert.ToInt16(row("ValorOuPeso"))
            End If
            Me.EncargoAgrupador = row("EncargoAgrupador")
            Me.VerificaEmpresa = row("VerificaEmpresa")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As String
    Private _Descricao As String
    Private _TipoPessoa As eTipoPessoa
    Private _ContaDebito As String
    Private _ContaCredito As String
    Private _BaseCalculo As Decimal
    Private _Aliquota As Decimal
    Private _Agrupar As Boolean
    Private _OperacaoXEncargo As Boolean
    Private _PodeSofreRetencao As Boolean
    Private _TipoPessoaRetencao As eTipoPessoa
    Private _Atualizacao As Boolean
    Private _Etapa As Integer
    Private _GravaCentroDeCusto As Boolean
    Private _ImprimirNFE As Boolean
    Private _Operador As String
    Private _EncargoAgrupador As String
    Private _ValorOuPeso As eValorPeso
    Private _VerificaEmpresa As Boolean
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public Property Codigo() As String
        Get
            Return _Codigo
        End Get
        Set(ByVal value As String)
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

    Public Property TipoPessoa As eTipoPessoa
        Get
            Return _TipoPessoa
        End Get
        Set(value As eTipoPessoa)
            _TipoPessoa = value
        End Set
    End Property

    Public Property ContaDebito() As String
        Get
            Return _ContaDebito
        End Get
        Set(ByVal value As String)
            _ContaDebito = value
        End Set
    End Property

    Public Property ContaCredito() As String
        Get
            Return _ContaCredito
        End Get
        Set(ByVal value As String)
            _ContaCredito = value
        End Set
    End Property

    Public Property BaseCalculo() As Decimal
        Get
            Return _BaseCalculo
        End Get
        Set(ByVal value As Decimal)
            _BaseCalculo = value
        End Set
    End Property

    Public Property Aliquota() As Decimal
        Get
            Return _Aliquota
        End Get
        Set(ByVal value As Decimal)
            _Aliquota = value
        End Set
    End Property

    Public Property Agrupar() As Boolean
        Get
            Return _Agrupar
        End Get
        Set(ByVal value As Boolean)
            _Agrupar = value
        End Set
    End Property

    Public Property OperacaoXEncargo() As Boolean
        Get
            Return _OperacaoXEncargo
        End Get
        Set(ByVal value As Boolean)
            _OperacaoXEncargo = value
        End Set
    End Property

    Public Property PodeSofreRetencao As Boolean
        Get
            Return _PodeSofreRetencao
        End Get
        Set(value As Boolean)
            _PodeSofreRetencao = value
        End Set
    End Property

    Public Property TipoPessoaRetencao As eTipoPessoa
        Get
            Return _TipoPessoaRetencao
        End Get
        Set(value As eTipoPessoa)
            _TipoPessoaRetencao = value
        End Set
    End Property

    Public Property Atualizacao() As Boolean
        Get
            Return _Atualizacao
        End Get
        Set(ByVal value As Boolean)
            _Atualizacao = value
        End Set
    End Property

    Public Property Etapa() As Integer
        Get
            Return _Etapa
        End Get
        Set(ByVal value As Integer)
            _Etapa = value
        End Set
    End Property

    Public Property GravaCentroDeCusto() As Boolean
        Get
            Return _GravaCentroDeCusto
        End Get
        Set(ByVal value As Boolean)
            _GravaCentroDeCusto = value
        End Set
    End Property

    Public Property ImprimirNFE() As Boolean
        Get
            Return _ImprimirNFE
        End Get
        Set(ByVal value As Boolean)
            _ImprimirNFE = value
        End Set
    End Property

    Public Property Operador As String
        Get
            Return _Operador
        End Get
        Set(value As String)
            _Operador = value
        End Set
    End Property

    Public Property EncargoAgrupador As String
        Get
            Return _EncargoAgrupador
        End Get
        Set(value As String)
            _EncargoAgrupador = value
        End Set
    End Property

    Public Property ValorOuPeso As eValorPeso
        Get
            Return _ValorOuPeso
        End Get
        Set(value As eValorPeso)
            _ValorOuPeso = value
        End Set
    End Property

    Public Property VerificaEmpresa As Boolean
        Get
            Return _VerificaEmpresa
        End Get
        Set(value As Boolean)
            _VerificaEmpresa = value
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
                Sql = " INSERT INTO Encargos(Encargo_id, Descricao, TipoDePessoa, ContaDebito, ContaCredito, BaseCalculo, Aliquota, OperacaoXEncargo, Etapa, Atualizacao, GravaCentroDeCusto, ImprimirNFE, EncargoAgrupador, ValorOuPeso, Operador, PodeSofreRetencao, TipoDePessoaRetencao,VerificaEmpresa) " & vbCrLf & _
                      " VALUES ('" & Me.Codigo & "','" & Me.Descricao & "'," & Me.TipoPessoa & ",'" & Me.ContaDebito & "','" & Me.ContaCredito & "'," & Str(Me.BaseCalculo) & "," & Str(Me.Aliquota) & ",'" & IIf(Me.OperacaoXEncargo, "S", "N") & "'," & Me.Etapa & "," & IIf(Me.Atualizacao, 1, 0) & "," & IIf(Me.GravaCentroDeCusto, 1, 0) & "," & IIf(Me.ImprimirNFE, 1, 0) & ",'" & Me.EncargoAgrupador & "'," & IIf(Me.ValorOuPeso = eValorPeso.Nenhum, "NULL", Me.ValorOuPeso) & ",'" & Me.Operador & "','" & Me.PodeSofreRetencao.ToString & "'," & Me.TipoPessoaRetencao & ",'" & Me.VerificaEmpresa.ToString & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Encargos SET" & vbCrLf & _
                      "   Descricao            ='" & Me.Descricao & "'" & vbCrLf & _
                      "  ,TipoDePessoa         = " & Me.TipoPessoa & vbCrLf & _
                      "  ,ContaDebito          ='" & Me.ContaDebito & "'" & vbCrLf & _
                      "  ,ContaCredito         ='" & Me.ContaCredito & "'" & vbCrLf & _
                      "  ,BaseCalculo          = " & Str(Me.BaseCalculo) & vbCrLf & _
                      "  ,Aliquota             = " & Str(Me.Aliquota) & vbCrLf & _
                      "  ,OperacaoXEncargo     ='" & IIf(Me.OperacaoXEncargo, "S", "N") & "'" & vbCrLf & _
                      "  ,Etapa                = " & Me.Etapa & vbCrLf & _
                      "  ,Atualizacao          = " & IIf(Me.Atualizacao, 1, 0) & vbCrLf & _
                      "  ,GravaCentroDeCusto   = " & IIf(Me.GravaCentroDeCusto, 1, 0) & vbCrLf & _
                      "  ,ImprimirNFE          = " & IIf(Me.ImprimirNFE, 1, 0) & vbCrLf & _
                      "  ,EncargoAgrupador     ='" & Me.EncargoAgrupador & "'" & vbCrLf & _
                      "  ,ValorOuPeso          = " & IIf(Me.ValorOuPeso = eValorPeso.Nenhum, "NULL", Me.ValorOuPeso) & vbCrLf & _
                      "  ,Operador             ='" & Me.Operador & "'" & vbCrLf & _
                      "  ,PodeSofreRetencao    ='" & Me.PodeSofreRetencao.ToString & "'" & vbCrLf & _
                      "  ,TipoDePessoaRetencao = " & TipoPessoaRetencao & vbCrLf & _
                      "  ,VerificaEmpresa      ='" & Me.VerificaEmpresa.ToString & "'" & _
                      "  WHERE Encargo_id ='" & Me.Codigo & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Encargos" & vbCrLf & _
                      "  WHERE Encargo_id ='" & Me.Codigo & "'" & vbCrLf

                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class