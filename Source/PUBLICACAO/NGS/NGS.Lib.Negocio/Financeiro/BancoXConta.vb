Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListBancosXContas
    Inherits List(Of BancosXContas)

    Public Sub New()

    End Sub

    Public Sub New(ByVal CarregarDados As Boolean, Optional ByVal Empresa As String = "")
        If CarregarDados Then
            Dim sql As String = " Select BxC.Empresa_Id, BxC.EndEmpresa_Id, BxC.Banco_Id, Bancos.Descricao AS NomeBanco, BxC.Agencia_Id, " & vbCrLf & _
                                " BxC.DigitoAgencia_Id, BxC.Conta_Id, BxC.DigitoConta_Id, BxC.ContaContabil , BxC.Observacoes,  isnull(BxC.TipoConta,'C') AS TipoConta, " & vbCrLf & _
                                " isnull(BxC.UsuarioInclusao,'') AS UsuarioInclusao, isnull(BxC.UsuarioInclusaoData, CURRENT_TIMESTAMP) AS UsuarioInclusaoData, " & vbCrLf & _
                                " isnull(BxC.UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(BxC.UsuarioAlteracaoData, CURRENT_TIMESTAMP) AS UsuarioAlteracaoData, " & vbCrLf & _
                                " isnull(BxC.NumCheque, 0) as NumCheque, isnull(BxC.NumChequeInicial, 0) as NumChequeInicial, isnull(BxC.NumChequeFinal, 0) as NumChequeFinal, " & vbCrLf & _
                                " isnull(BXC.FluxoDeCaixa,0) as FluxoDeCaixa, isnull(BXC.LimiteBancario, 0) as LimiteBancario, isnull(BxC.Ativo, 0) as Ativo" & vbCrLf & _
                                " From BancosXContas BxC " & vbCrLf & _
                                " INNER JOIN Bancos " & vbCrLf & _
                                "    ON BxC.Banco_id = Bancos.Banco_Id "

            If Empresa.Length > 0 Then
                sql &= "Where BxC.Empresa_id = '" & Empresa & "'"
            End If

            Dim Banco As New AcessaBanco
            Dim ds As DataSet

            ds = Banco.ConsultaDataSet(sql, "BancosXContas")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim BxC As New BancosXContas
                BxC.CodigoEmpresa = row("Empresa_Id")
                BxC.EndEmpresa = row("EndEmpresa_Id")
                BxC.CodigoBanco = row("Banco_Id")
                BxC.NomeBanco = row("NomeBanco")
                BxC.Agencia = row("Agencia_Id")
                BxC.DigitoAgencia = row("DigitoAgencia_Id")
                BxC.Conta = row("Conta_Id")
                BxC.DigitoConta = row("DigitoConta_Id")
                BxC.CodigoContaContabil = row("ContaContabil")
                BxC.Observacoes = row("Observacoes")
                BxC.TipoConta = row("TipoConta")
                BxC.UsuarioInclusao = row("UsuarioInclusao")
                BxC.DataInclusao = row("UsuarioInclusaoData")
                BxC.UsuarioAlteracao = row("UsuarioAlteracao")
                BxC.DataAlteracao = row("UsuarioAlteracaoData")
                BxC.NumChequeAtual = row("NumCheque")
                BxC.NumChequeInicial = row("NumChequeInicial")
                BxC.NumChequeFinal = row("NumChequeFinal")
                BxC.FluxoDeCaixa = row("FluxoDeCaixa")
                BxC.LimiteBanc = row("LimiteBancario")
                BxC.Ativo = row("Ativo")
                Me.Add(BxC)
            Next
        End If
    End Sub
End Class

<Serializable()> _
Public Class BancosXContas
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal objBancoXConta As [Lib].Negocio.BancosXContas)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = " Select BxC.Empresa_Id, BxC.EndEmpresa_Id, BxC.Banco_Id, Bancos.Descricao AS NomeBanco, BxC.Agencia_Id, " & vbCrLf & _
        " BxC.DigitoAgencia_Id, BxC.Conta_Id, BxC.DigitoConta_Id, BxC.ContaContabil , BxC.Observacoes, isnull(BxC.TipoConta,'C') AS TipoConta,  " & vbCrLf & _
        " isnull(BxC.UsuarioInclusao,'') AS UsuarioInclusao, isnull(BxC.UsuarioInclusaoData, CURRENT_TIMESTAMP) AS UsuarioInclusaoData, " & vbCrLf & _
        " isnull(BxC.UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(BxC.UsuarioAlteracaoData, CURRENT_TIMESTAMP) AS UsuarioAlteracaoData, " & vbCrLf & _
        " isnull(BxC.NumCheque, 0) as NumCheque, isnull(BxC.NumChequeInicial, 0) as NumChequeInicial, isnull(BxC.NumChequeFinal, 0) as NumChequeFinal, " & vbCrLf & _
        " isnull(BxC.FluxoDeCaixa,0) as FluxoDeCaixa, isnull(BxC.LimiteBancario, 0) as LimiteBancario, isnull(BxC.Ativo, 0) as Ativo" & vbCrLf & _
        " From BancosXContas BxC " & vbCrLf & _
        " INNER JOIN Bancos " & vbCrLf & _
        "    ON BxC.Banco_id = Bancos.Banco_Id " & vbCrLf & _
        "  WHERE BxC.Empresa_Id      = '" & objBancoXConta.CodigoEmpresa & "'" & vbCrLf & _
        "    AND BxC.EndEmpresa_Id   = " & objBancoXConta.EndEmpresa & vbCrLf & _
        "    AND BxC.Banco_Id        = " & objBancoXConta.CodigoBanco & vbCrLf & _
        "    AND BxC.Agencia_Id      = '" & objBancoXConta.Agencia & "'" & vbCrLf & _
        "    AND BxC.DigitoAgencia_Id = '" & objBancoXConta.DigitoAgencia & "'" & vbCrLf & _
        "    AND BxC.Conta_Id         = '" & objBancoXConta.Conta & "'" & vbCrLf & _
        "    AND BxC.DigitoConta_Id   = '" & objBancoXConta.DigitoConta & "'"

        ds = Banco.ConsultaDataSet(sql, "BancosXContas")

        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoEmpresa = ds.Tables(0).Rows(0)("Empresa_Id")
            _EndEmpresa = ds.Tables(0).Rows(0)("EndEmpresa_Id")
            _CodigoBanco = ds.Tables(0).Rows(0)("Banco_Id")
            _NomeBanco = ds.Tables(0).Rows(0)("NomeBanco")
            _Agencia = ds.Tables(0).Rows(0)("Agencia_Id")
            _DigitoAgencia = ds.Tables(0).Rows(0)("DigitoAgencia_Id")
            _Conta = ds.Tables(0).Rows(0)("Conta_Id")
            _DigitoConta = ds.Tables(0).Rows(0)("DigitoConta_Id")
            _CodigoContaContabil = ds.Tables(0).Rows(0)("ContaContabil")
            _Observacoes = ds.Tables(0).Rows(0)("Observacoes")
            _TipoConta = ds.Tables(0).Rows(0)("TipoConta")
            _UsuarioInclusao = ds.Tables(0).Rows(0)("UsuarioInclusao")
            _UsuarioInclusaoData = ds.Tables(0).Rows(0)("UsuarioInclusaoData")
            _UsuarioAlteracao = ds.Tables(0).Rows(0)("UsuarioAlteracao")
            _UsuarioAlteracaoData = ds.Tables(0).Rows(0)("UsuarioAlteracaoData")
            _NumChequeAtual = ds.Tables(0).Rows(0)("NumCheque")
            _NumChequeInicial = ds.Tables(0).Rows(0)("NumChequeInicial")
            _NumChequeFinal = ds.Tables(0).Rows(0)("NumChequeFinal")
            _FluxoDeCaixa = ds.Tables(0).Rows(0)("FluxoDeCaixa")
            _LimiteBanc = ds.Tables(0).Rows(0)("LimiteBancario")
            _Ativo = ds.Tables(0).Rows(0)("Ativo")
        End If
    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pBanco As Integer, ByVal pAgencia As String, ByVal pDigitoAgencia As String, ByVal pConta As String, ByVal pDigitoConta As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = " Select BxC.Empresa_Id, BxC.EndEmpresa_Id, BxC.Banco_Id, Bancos.Descricao AS NomeBanco, BxC.Agencia_Id, " & vbCrLf & _
                            " BxC.DigitoAgencia_Id, BxC.Conta_Id, BxC.DigitoConta_Id, BxC.ContaContabil , BxC.Observacoes, isnull(BxC.TipoConta,'C') AS TipoConta,  " & vbCrLf & _
                            " isnull(BxC.UsuarioInclusao,'') AS UsuarioInclusao, isnull(BxC.UsuarioInclusaoData, CURRENT_TIMESTAMP) AS UsuarioInclusaoData, " & vbCrLf & _
                            " isnull(BxC.UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(BxC.UsuarioAlteracaoData, CURRENT_TIMESTAMP) AS UsuarioAlteracaoData, " & vbCrLf & _
                            " isnull(BxC.NumCheque, 0) as NumCheque, isnull(BxC.NumChequeInicial, 0) as NumChequeInicial, isnull(BxC.NumChequeFinal, 0) as NumChequeFinal," & vbCrLf & _
                            " isnull(BxC.FluxoDeCaixa,0) as FluxoDeCaixa, isnull(BxC.LimiteBancario, 0) as LimiteBancario, isnull(BxC.Ativo, 0) as Ativo" & vbCrLf & _
                            " From BancosXContas BxC " & vbCrLf & _
                            " INNER JOIN Bancos " & vbCrLf & _
                            "    ON BxC.Banco_id = Bancos.Banco_Id " & vbCrLf & _
                            "  WHERE BxC.Empresa_Id      = '" & pEmpresa & "'" & vbCrLf & _
                            "    AND BxC.EndEmpresa_Id   = " & pEndEmpresa & vbCrLf & _
                            "    AND BxC.Banco_Id        = " & pBanco & vbCrLf & _
                            "    AND BxC.Agencia_Id      = '" & pAgencia & "'" & vbCrLf & _
                            "    AND BxC.DigitoAgencia_Id = '" & pDigitoAgencia & "'" & vbCrLf & _
                            "    AND BxC.Conta_Id         = '" & pConta & "'" & vbCrLf & _
                            "    AND BxC.DigitoConta_Id   = '" & pDigitoConta & "'"

        ds = Banco.ConsultaDataSet(sql, "BancosXContas")

        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoEmpresa = ds.Tables(0).Rows(0)("Empresa_Id")
            _EndEmpresa = ds.Tables(0).Rows(0)("EndEmpresa_Id")
            _CodigoBanco = ds.Tables(0).Rows(0)("Banco_Id")
            _NomeBanco = ds.Tables(0).Rows(0)("NomeBanco")
            _Agencia = ds.Tables(0).Rows(0)("Agencia_Id")
            _DigitoAgencia = ds.Tables(0).Rows(0)("DigitoAgencia_Id")
            _Conta = ds.Tables(0).Rows(0)("Conta_Id")
            _DigitoConta = ds.Tables(0).Rows(0)("DigitoConta_Id")
            _CodigoContaContabil = ds.Tables(0).Rows(0)("ContaContabil")
            _Observacoes = ds.Tables(0).Rows(0)("Observacoes")
            _TipoConta = ds.Tables(0).Rows(0)("TipoConta")
            _UsuarioInclusao = ds.Tables(0).Rows(0)("UsuarioInclusao")
            _UsuarioInclusaoData = ds.Tables(0).Rows(0)("UsuarioInclusaoData")
            _UsuarioAlteracao = ds.Tables(0).Rows(0)("UsuarioAlteracao")
            _UsuarioAlteracaoData = ds.Tables(0).Rows(0)("UsuarioAlteracaoData")
            _NumChequeAtual = ds.Tables(0).Rows(0)("NumCheque")
            _NumChequeInicial = ds.Tables(0).Rows(0)("NumChequeInicial")
            _NumChequeFinal = ds.Tables(0).Rows(0)("NumChequeFinal")
            _FluxoDeCaixa = ds.Tables(0).Rows(0)("FluxoDeCaixa")
            _LimiteBanc = ds.Tables(0).Rows(0)("LimiteBancario")
            _Ativo = ds.Tables(0).Rows(0)("Ativo")
        End If
    End Sub

    Public Sub New(ByVal pBanco As Integer, ByVal pAgencia As String, ByVal pDigitoAgencia As String, ByVal pConta As String, ByVal pDigitoConta As String)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = " Select GE.Empresa_Id AS Unidade, BxC.Empresa_Id, BxC.EndEmpresa_Id, BxC.Banco_Id, Bancos.Descricao AS NomeBanco, BxC.Agencia_Id, " & vbCrLf &
                            " BxC.DigitoAgencia_Id, BxC.Conta_Id, BxC.DigitoConta_Id, BxC.ContaContabil , BxC.Observacoes, isnull(BxC.TipoConta,'C') AS TipoConta,  " & vbCrLf &
                            " isnull(BxC.UsuarioInclusao,'') AS UsuarioInclusao, isnull(BxC.UsuarioInclusaoData, CURRENT_TIMESTAMP) AS UsuarioInclusaoData, " & vbCrLf &
                            " isnull(BxC.UsuarioAlteracao,'') AS UsuarioAlteracao, isnull(BxC.UsuarioAlteracaoData, CURRENT_TIMESTAMP) AS UsuarioAlteracaoData, " & vbCrLf &
                            " isnull(BxC.FluxoDeCaixa,0) as FluxoDeCaixa, isnull(BxC.LimiteBancario, 0) as LimiteBancario " & vbCrLf &
                            " From BancosXContas BxC " & vbCrLf &
                            " INNER JOIN GruposXEmpresas AS GE " & vbCrLf &
                            " ON GE.Cliente_Id = BxC.Empresa_Id " & vbCrLf &
                            " AND GE.EndCliente_Id = BxC.EndEmpresa_Id " & vbCrLf &
                            " INNER JOIN Bancos " & vbCrLf &
                            "    ON BxC.Banco_id = Bancos.Banco_Id " & vbCrLf &
                            "  WHERE BxC.Banco_Id = " & pBanco & vbCrLf &
                            "    AND BxC.Agencia_Id      = " & pAgencia & vbCrLf &
                            "    AND BxC.DigitoAgencia_Id = '" & pDigitoAgencia & "'" & vbCrLf &
                            "    AND BxC.Conta_Id         = '" & pConta & "'" & vbCrLf &
                            "    AND BxC.DigitoConta_Id   = '" & pDigitoConta.Trim() & "'"

        ds = Banco.ConsultaDataSet(sql, "BancosXContas")

        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoEmpresa = ds.Tables(0).Rows(0)("Empresa_Id")
            _CodigoUnidadeDeNegocio = ds.Tables(0).Rows(0)("Unidade")
            _EndEmpresa = ds.Tables(0).Rows(0)("EndEmpresa_Id")
            _CodigoBanco = ds.Tables(0).Rows(0)("Banco_Id")
            _NomeBanco = ds.Tables(0).Rows(0)("NomeBanco")
            _Agencia = ds.Tables(0).Rows(0)("Agencia_Id")
            _DigitoAgencia = ds.Tables(0).Rows(0)("DigitoAgencia_Id")
            _Conta = ds.Tables(0).Rows(0)("Conta_Id")
            _DigitoConta = ds.Tables(0).Rows(0)("DigitoConta_Id")
            _CodigoContaContabil = ds.Tables(0).Rows(0)("ContaContabil")
            _Observacoes = ds.Tables(0).Rows(0)("Observacoes")
            _TipoConta = ds.Tables(0).Rows(0)("TipoConta")
            _UsuarioInclusao = ds.Tables(0).Rows(0)("UsuarioInclusao")
            _UsuarioInclusaoData = ds.Tables(0).Rows(0)("UsuarioInclusaoData")
            _UsuarioAlteracao = ds.Tables(0).Rows(0)("UsuarioAlteracao")
            _UsuarioAlteracaoData = ds.Tables(0).Rows(0)("UsuarioAlteracaoData")
            _FluxoDeCaixa = ds.Tables(0).Rows(0)("FluxoDeCaixa")
            _LimiteBanc = ds.Tables(0).Rows(0)("LimiteBancario")
        End If
    End Sub

    Public Sub New(ByVal pContaContabil As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = " Select GE.Empresa_Id AS Unidade," & vbCrLf & _
                            "        BxC.Empresa_Id," & vbCrLf & _
                            "        BxC.EndEmpresa_Id," & vbCrLf & _
                            "        BxC.Banco_Id," & vbCrLf & _
                            "        Bancos.Descricao AS NomeBanco," & vbCrLf & _
                            "        BxC.Agencia_Id, " & vbCrLf & _
                            "        BxC.DigitoAgencia_Id," & vbCrLf & _
                            "        BxC.Conta_Id," & vbCrLf & _
                            "        BxC.DigitoConta_Id," & vbCrLf & _
                            "        BxC.ContaContabil," & vbCrLf & _
                            "        BxC.Observacoes," & vbCrLf & _
                            "        isnull(BxC.TipoConta,'C') AS TipoConta,  " & vbCrLf & _
                            "        isnull(BxC.UsuarioInclusao,'') AS UsuarioInclusao," & vbCrLf & _
                            "        isnull(BxC.UsuarioInclusaoData, CURRENT_TIMESTAMP) AS UsuarioInclusaoData, " & vbCrLf & _
                            "        isnull(BxC.UsuarioAlteracao,'') AS UsuarioAlteracao," & vbCrLf & _
                            "        isnull(BxC.UsuarioAlteracaoData, CURRENT_TIMESTAMP) AS UsuarioAlteracaoData, " & vbCrLf & _
                            "        isnull(BxC.FluxoDeCaixa,0) as FluxoDeCaixa, isnull(BxC.LimiteBancario, 0) as LimiteBancario " & vbCrLf & _
                            "   From BancosXContas BxC " & vbCrLf & _
                            "  INNER JOIN GruposXEmpresas AS GE " & vbCrLf & _
                            "     ON GE.Cliente_Id = BxC.Empresa_Id " & vbCrLf & _
                            "    AND GE.EndCliente_Id = BxC.EndEmpresa_Id " & vbCrLf & _
                            "  INNER JOIN Bancos " & vbCrLf & _
                            "     ON BxC.Banco_id = Bancos.Banco_Id " & vbCrLf & _
                            "  WHERE BxC.ContaContabil ='" & pContaContabil & "'" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "BancosXContas")

        If ds.Tables(0).Rows.Count > 0 Then
            _CodigoEmpresa = ds.Tables(0).Rows(0)("Empresa_Id")
            _CodigoUnidadeDeNegocio = ds.Tables(0).Rows(0)("Unidade")
            _EndEmpresa = ds.Tables(0).Rows(0)("EndEmpresa_Id")
            _CodigoBanco = ds.Tables(0).Rows(0)("Banco_Id")
            _NomeBanco = ds.Tables(0).Rows(0)("NomeBanco")
            _Agencia = ds.Tables(0).Rows(0)("Agencia_Id")
            _DigitoAgencia = ds.Tables(0).Rows(0)("DigitoAgencia_Id")
            _Conta = ds.Tables(0).Rows(0)("Conta_Id")
            _DigitoConta = ds.Tables(0).Rows(0)("DigitoConta_Id")
            _CodigoContaContabil = ds.Tables(0).Rows(0)("ContaContabil")
            _Observacoes = ds.Tables(0).Rows(0)("Observacoes")
            _TipoConta = ds.Tables(0).Rows(0)("TipoConta")
            _UsuarioInclusao = ds.Tables(0).Rows(0)("UsuarioInclusao")
            _UsuarioInclusaoData = ds.Tables(0).Rows(0)("UsuarioInclusaoData")
            _UsuarioAlteracao = ds.Tables(0).Rows(0)("UsuarioAlteracao")
            _UsuarioAlteracaoData = ds.Tables(0).Rows(0)("UsuarioAlteracaoData")
            _FluxoDeCaixa = ds.Tables(0).Rows(0)("FluxoDeCaixa")
            _LimiteBanc = ds.Tables(0).Rows(0)("LimiteBancario")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoEmpresa As String = ""
    Private _CodigoUnidadeDeNegocio As String = String.Empty
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente
    Private _CodigoBanco As Integer
    Private _NomeBanco As String = ""
    Private _Banco As Banco
    Private _Agencia As String = ""
    Private _DigitoAgencia As String = ""
    Private _Conta As String = ""
    Private _DigitoConta As String = ""
    Private _CodigoContaContabil As String = ""
    Private _ContaContabil As PlanoDeConta
    Private _Observacoes As String = ""
    Private _TipoConta As String = ""
    Private _UsuarioInclusao As String = ""
    Private _UsuarioInclusaoData As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _UsuarioAlteracaoData As DateTime
    Private _NumChequeAtual As Integer
    Private _NumChequeInicial As Integer
    Private _NumChequeFinal As Integer
    Private _Ativo As Integer
    Private _LimiteBanc As Decimal
    Private _FluxoDeCaixa As Boolean
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

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property CodigoUnidadeDeNegocio() As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _CodigoUnidadeDeNegocio = value
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

    Public Property CodigoBanco() As Integer
        Get
            Return _CodigoBanco
        End Get
        Set(ByVal value As Integer)
            _CodigoBanco = value
        End Set
    End Property

    Public Property NomeBanco() As String
        Get
            Return _NomeBanco
        End Get
        Set(ByVal value As String)
            _NomeBanco = value
        End Set
    End Property

    Public Property Banco() As Banco
        Get
            If _CodigoBanco > 0 And _Banco Is Nothing Then _Banco = New Banco(_CodigoBanco)
            Return _Banco
        End Get
        Set(ByVal value As Banco)
            _Banco = value
        End Set
    End Property

    Public Property Agencia() As String
        Get
            Return _Agencia
        End Get
        Set(ByVal value As String)
            _Agencia = value
        End Set
    End Property

    Public Property DigitoAgencia() As String
        Get
            Return _DigitoAgencia
        End Get
        Set(ByVal value As String)
            _DigitoAgencia = value
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

    Public Property DigitoConta() As String
        Get
            Return _DigitoConta
        End Get
        Set(ByVal value As String)
            _DigitoConta = value
        End Set
    End Property

    Public Property CodigoContaContabil() As String
        Get
            Return _CodigoContaContabil
        End Get
        Set(ByVal value As String)
            _CodigoContaContabil = value
        End Set
    End Property

    Public Property ContaContabil() As PlanoDeConta
        Get
            If _CodigoContaContabil.Length > 0 And _ContaContabil Is Nothing Then _ContaContabil = New PlanoDeConta("", 0, _CodigoContaContabil)
            Return _ContaContabil
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaContabil = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property TipoConta() As String
        Get
            Return _TipoConta
        End Get
        Set(ByVal value As String)
            _TipoConta = value
        End Set
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property NumChequeAtual() As Integer
        Get
            Return _NumChequeAtual
        End Get
        Set(ByVal value As Integer)
            _NumChequeAtual = value
        End Set
    End Property

    Public Property NumChequeInicial() As Integer
        Get
            Return _NumChequeInicial
        End Get
        Set(ByVal value As Integer)
            _NumChequeInicial = value
        End Set
    End Property

    Public Property NumChequeFinal() As Integer
        Get
            Return _NumChequeFinal
        End Get
        Set(ByVal value As Integer)
            _NumChequeFinal = value
        End Set
    End Property

    Public Property Ativo() As Integer
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Integer)
            _Ativo = value
        End Set
    End Property

    Public Property LimiteBanc() As Decimal
        Get
            Return _LimiteBanc
        End Get
        Set(ByVal value As Decimal)
            _LimiteBanc = value
        End Set
    End Property

    Public Property FluxoDeCaixa() As Boolean
        Get
            Return _FluxoDeCaixa
        End Get
        Set(ByVal value As Boolean)
            _FluxoDeCaixa = value
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
                Sql = " INSERT INTO BancosXContas (Empresa_Id, EndEmpresa_Id, Banco_Id, Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, " & vbCrLf & _
                      "              ContaContabil, Observacoes, TipoConta, UsuarioInclusao, UsuarioInclusaoData, NumCheque, NumChequeInicial, NumChequeFinal, FluxoDeCaixa, LimiteBancario, Ativo) " & vbCrLf & _
                      " VALUES ('" & _CodigoEmpresa & "'," & _EndEmpresa & "," & _CodigoBanco & ",'" & _Agencia & "','" & _DigitoAgencia & "','" & _Conta & "','" & _DigitoConta & "'," & vbCrLf & _
                      "         '" & _CodigoContaContabil & "','" & _Observacoes & "','" & _TipoConta & "','" & _UsuarioInclusao & "','" & _UsuarioInclusaoData.ToSqlDate() & "', " & vbCrLf & _
                      "          " & _NumChequeAtual & "," & _NumChequeInicial & ", " & _NumChequeFinal & ", '" & _FluxoDeCaixa & "', " & Str(_LimiteBanc) & ", " & _Ativo & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE BancosXContas " & vbCrLf & _
                      "    SET ContaContabil        = '" & _CodigoContaContabil & "'," & vbCrLf & _
                      "        Observacoes          = '" & _Observacoes & "'," & vbCrLf & _
                      "        TipoConta            = '" & _TipoConta & "'," & vbCrLf & _
                      "        UsuarioAlteracao     = '" & _UsuarioAlteracao & "'," & vbCrLf & _
                      "        UsuarioAlteracaoData = '" & _UsuarioAlteracaoData.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "        NumCheque            = " & _NumChequeAtual & "," & vbCrLf & _
                      "        NumChequeInicial     = " & _NumChequeInicial & "," & vbCrLf & _
                      "        NumChequeFinal       = " & _NumChequeFinal & "," & vbCrLf & _
                      "        FluxoDeCaixa         = '" & _FluxoDeCaixa & "'," & vbCrLf & _
                      "        LimiteBancario       = " & Str(_LimiteBanc) & "," & vbCrLf & _
                      "        Ativo                = " & _Ativo & vbCrLf & _
                      "  WHERE Empresa_Id       = '" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    AND EndEmpresa_Id    = " & _EndEmpresa & vbCrLf & _
                      "    AND Banco_Id         = " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia_Id       = '" & _Agencia & "'" & vbCrLf & _
                      "    AND DigitoAgencia_Id = '" & _DigitoAgencia & "'" & vbCrLf & _
                      "    AND Conta_Id         = '" & _Conta & "'" & vbCrLf & _
                      "    AND DigitoConta_Id   = '" & _DigitoConta & "'" & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE BancosXContas " & vbCrLf & _
                      "  WHERE Empresa_Id       = '" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    AND EndEmpresa_Id    = " & _EndEmpresa & vbCrLf & _
                      "    AND Banco_Id         = " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia_Id       = '" & _Agencia & "'" & vbCrLf & _
                      "    AND DigitoAgencia_Id = '" & _DigitoAgencia & "'" & vbCrLf & _
                      "    AND Conta_Id         = '" & _Conta & "'" & vbCrLf & _
                      "    AND DigitoConta_Id   = '" & _DigitoConta & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class