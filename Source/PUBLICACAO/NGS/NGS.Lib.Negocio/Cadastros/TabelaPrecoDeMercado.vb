Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListTabelaPrecoDeMercado
    Inherits List(Of TabelaPrecoDeMercado)

    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal Deposito As String, ByVal Produto As String, Optional ByVal DataInicial As String = "", Optional ByVal DataFinal As String = "")
        Dim sql As String
        sql = "SELECT TPM.Empresa_Id, " & vbCrLf & _
              "       TPM.EndEmpresa_Id, " & vbCrLf & _
              "       TPM.Deposito_Id, " & vbCrLf & _
              "       TPM.EndDeposito_Id, " & vbCrLf & _
              "       TPM.Produto_Id, " & vbCrLf & _
              "       TPM.Data_Id, " & vbCrLf & _
              "       TPM.ValorOficial, " & vbCrLf & _
              "       TPM.ValorMoeda, " & vbCrLf & _
              "       TPM.BaseDeCalculo, " & vbCrLf & _
              "       Empresa.Reduzido as EmpresaReduzido,  " & vbCrLf & _
              "       Empresa.Nome as EmpresaNome,  " & vbCrLf & _
              "       Deposito.Reduzido AS DepositoReduzido,  " & vbCrLf & _
              "       Deposito.Nome AS DepositoNome, " & vbCrLf & _
              "       Produtos.Nome AS ProdutoNome " & vbCrLf & _
              "  FROM TabelaDePrecosDeMercado TPM " & vbCrLf & _
              " INNER JOIN Clientes AS Empresa" & vbCrLf & _
              "    ON TPM.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf & _
              "   AND TPM.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
              " INNER JOIN Clientes AS Deposito" & vbCrLf & _
              "    ON TPM.Deposito_Id    = Deposito.Cliente_Id  " & vbCrLf & _
              "   AND TPM.EndDeposito_Id = Deposito.Endereco_Id " & vbCrLf & _
              " INNER JOIN Produtos " & vbCrLf & _
              "    ON TPM.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
              " Where 1 = 1 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(Empresa) Then
            sql &= " and TPM.Empresa_id    ='" & Empresa.Split("-")(0) & "'" & vbCrLf & _
                   " and TPM.EndEmpresa_id = " & Empresa.Split("-")(1)
        End If

        If Not String.IsNullOrWhiteSpace(Deposito) Then
            sql &= " and TPM.Deposito_Id ='" & Deposito.Split("-")(0) & "'" & vbCrLf & _
                   " and TPM.EndDeposito_id = " & Deposito.Split("-")(1)
        End If

        If Not String.IsNullOrWhiteSpace(Produto) Then
            sql &= " and TPM.Produto_Id = '" & Produto & "'"
        End If

        If Not String.IsNullOrWhiteSpace(DataInicial) AndAlso Not String.IsNullOrWhiteSpace(DataFinal) Then
            sql &= " and TPM.Data_Id Between '" & CDate(DataInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(DataFinal).ToString("yyyy-MM-dd") & "' " & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim DsPrecoMercado As DataSet = Banco.ConsultaDataSet(sql, "TPM")

        For Each dr As DataRow In DsPrecoMercado.Tables(0).Rows
            Dim objPrecoMercado = New TabelaPrecoDeMercado
            objPrecoMercado.Empresa = dr("Empresa_Id")
            objPrecoMercado.EndEmpresa = dr("EndEmpresa_Id")
            objPrecoMercado.EmpresaReduzido = dr("EmpresaReduzido")
            objPrecoMercado.EmpresaNome = dr("Empresa_Id") & "-" & dr("EndEmpresa_Id") & "-" & dr("EmpresaNome")

            objPrecoMercado.Deposito = dr("Deposito_id")
            objPrecoMercado.EndDeposito = dr("EndDeposito_id")
            objPrecoMercado.DepositoReduzido = dr("DepositoReduzido")
            objPrecoMercado.DepositoNome = dr("Deposito_Id") & "-" & dr("EndDeposito_Id") & "-" & dr("DepositoNome")

            objPrecoMercado.Data_Id = dr("Data_Id")

            objPrecoMercado.CodigoProduto = dr("Produto_Id")
            objPrecoMercado.ProdutoNome = dr("Produto_Id") & " - " & dr("ProdutoNome")
            objPrecoMercado.ValorOficial = dr("ValorOficial")
            objPrecoMercado.ValorMoeda = dr("ValorMoeda")
            objPrecoMercado.BaseDeCalculo = dr("BaseDeCalculo")

            Me.Add(objPrecoMercado)
        Next
    End Sub
End Class

<Serializable()> _
Public Class TabelaPrecoDeMercado
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal Deposito As String, ByVal Produto As String, ByVal data As String)
        Dim sql As String
        sql = "SELECT TPM.Empresa_Id, " & vbCrLf & _
              "       TPM.EndEmpresa_Id, " & vbCrLf & _
              "       TPM.Deposito_Id, " & vbCrLf & _
              "       TPM.EndDeposito_Id, " & vbCrLf & _
              "       TPM.Produto_Id, " & vbCrLf & _
              "       TPM.Data_Id, " & vbCrLf & _
              "       TPM.ValorOficial, " & vbCrLf & _
              "       TPM.ValorMoeda, " & vbCrLf & _
              "       TPM.BaseDeCalculo, " & vbCrLf & _
              "       Empresa.Reduzido as EmpresaReduzido,  " & vbCrLf & _
              "       Empresa.Nome as EmpresaNome,  " & vbCrLf & _
              "       Deposito.Reduzido AS DepositoReduzido,  " & vbCrLf & _
              "       Deposito.Nome AS DepositoNome, " & vbCrLf & _
              "       Produtos.Nome AS ProdutoNome " & vbCrLf & _
              "  FROM TabelaDePrecosDeMercado TPM " & vbCrLf & _
              " INNER JOIN Clientes AS Empresa" & vbCrLf & _
              "    ON TPM.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf & _
              "   AND TPM.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
              " INNER JOIN Clientes AS Deposito" & vbCrLf & _
              "    ON TPM.Deposito_Id    = Deposito.Cliente_Id  " & vbCrLf & _
              "   AND TPM.EndDeposito_Id = Deposito.Endereco_Id " & vbCrLf & _
              " INNER JOIN Produtos " & vbCrLf & _
              "    ON TPM.Produto_Id = Produtos.Produto_Id " & vbCrLf & _
              " Where 1 = 1 " & vbCrLf

        If Not String.IsNullOrWhiteSpace(Empresa) Then
            sql &= " and TPM.Empresa_id    ='" & Empresa.Split("-")(0) & "'" & vbCrLf & _
                   " and TPM.EndEmpresa_id = " & Empresa.Split("-")(1)
        End If

        If Not String.IsNullOrWhiteSpace(Deposito) Then
            sql &= " and TPM.Deposito_Id ='" & Deposito.Split("-")(0) & "'" & vbCrLf & _
                   " and TPM.EndDeposito_id = " & Deposito.Split("-")(1)
        End If

        If Not String.IsNullOrWhiteSpace(Produto) Then
            sql &= " and TPM.Produto_Id = '" & Produto & "'"
        End If

        If Not String.IsNullOrWhiteSpace(data) Then
            sql &= " and TPM.Data_Id = '" & CDate(data).ToString("yyyy-MM-dd") & "'" & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        Dim DsPrecoMercado As DataSet = Banco.ConsultaDataSet(sql, "TPM")

        For Each row As DataRow In DsPrecoMercado.Tables(0).Rows
            Dim objPrecoMercado = New TabelaPrecoDeMercado
            _Empresa = row("Empresa_Id")
            _EndEmpresa = row("EndEmpresa_Id")
            _EmpresaReduzido = row("EmpresaReduzido")
            _EmpresaNome = row("Empresa_Id") & "-" & row("EndEmpresa_Id") & "-" & row("EmpresaNome")

            _Deposito = row("Deposito_id")
            _EndDeposito = row("EndDeposito_id")
            _DepositoReduzido = row("DepositoReduzido")
            _DepositoNome = row("Deposito_Id") & "-" & row("EndDeposito_Id") & "-" & row("DepositoNome")

            _Data_Id = row("Data_Id")

            _CodigoProduto = row("Produto_Id")
            _ProdutoNome = row("ProdutoNome")
            _ValorOficial = row("ValorOficial")
            _ValorMoeda = row("ValorMoeda")
            _BaseDeCalculo = row("BaseDeCalculo")
        Next
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String

    Private _Empresa As String = ""
    Private _EndEmpresa As Integer
    Private _EmpresaReduzido As String = ""
    Private _EmpresaNome As String

    Private _Deposito As String = ""
    Private _EndDeposito As Integer
    Private _DepositoReduzido As String
    Private _DepositoNome As String

    Private _CodigoProduto As String = ""
    Private _ProdutoNome As String
    Private _Produto As Produto
    Private _Data_Id As Date

    Private _ValorOficial As Decimal
    Private _ValorMoeda As Decimal
    Private _BaseDeCalculo As Integer

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property Empresa() As String
        Get
            Return _Empresa
        End Get
        Set(ByVal value As String)
            _Empresa = value
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

    Public Property EmpresaReduzido() As String
        Get
            Return _EmpresaReduzido
        End Get
        Set(ByVal value As String)
            _EmpresaReduzido = value
        End Set
    End Property

    Public Property EmpresaNome() As String
        Get
            Return _EmpresaNome
        End Get
        Set(ByVal value As String)
            _EmpresaNome = value
        End Set
    End Property

    Public Property Deposito() As String
        Get
            Return _Deposito
        End Get
        Set(ByVal value As String)
            _Deposito = value
        End Set
    End Property

    Public Property EndDeposito() As Integer
        Get
            Return _EndDeposito
        End Get
        Set(ByVal value As Integer)
            _EndDeposito = value
        End Set
    End Property

    Public Property DepositoReduzido() As String
        Get
            Return _DepositoReduzido
        End Get
        Set(ByVal value As String)
            _DepositoReduzido = value
        End Set
    End Property

    Public Property DepositoNome() As String
        Get
            Return _DepositoNome
        End Get
        Set(ByVal value As String)
            _DepositoNome = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property


    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public Property ProdutoNome() As String
        Get
            Return _ProdutoNome
        End Get
        Set(ByVal value As String)
            _ProdutoNome = value
        End Set
    End Property

    Public Property Data_Id() As Date
        Get
            Return _Data_Id
        End Get
        Set(ByVal value As Date)
            _Data_Id = value
        End Set
    End Property


    Public Property ValorMoeda() As Decimal
        Get
            Return _ValorMoeda
        End Get
        Set(ByVal value As Decimal)
            _ValorMoeda = value
        End Set
    End Property

    Public Property ValorOficial() As Decimal
        Get
            Return _ValorOficial
        End Get
        Set(ByVal value As Decimal)
            _ValorOficial = value
        End Set
    End Property

    Public Property BaseDeCalculo() As Integer
        Get
            Return _BaseDeCalculo
        End Get
        Set(ByVal value As Integer)
            _BaseDeCalculo = value
        End Set
    End Property


#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = Nothing
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String

        Select Case Me.IUD
            Case "I"
                sql = "	Insert into TabelaDePrecosDeMercado " & vbCrLf & _
                      "   (Empresa_Id, EndEmpresa_Id, Deposito_Id, EndDeposito_Id, Produto_Id, " & vbCrLf & _
                      "    Data_Id, ValorOficial, ValorMoeda, BaseDeCalculo)" & vbCrLf & _
                      " Values('" & _Empresa & "'," & _EndEmpresa & ",'" & _Deposito & "'," & _EndDeposito & ", '" & _CodigoProduto & "'," & vbCrLf & _
                      "'" & _Data_Id.ToString("yyyy-MM-dd") & "', " & Str(_ValorOficial) & ", " & Str(_ValorMoeda) & ", " & _BaseDeCalculo & "); " & vbCrLf
                Sqls.Add(sql)

            Case "U"
                sql = "	  update TabelaDePrecosDeMercado set " & vbCrLf & _
                      "      ValorOficial      =  " & Str(_ValorOficial) & vbCrLf & _
                      "      ,ValorMoeda       =  " & Str(_ValorMoeda) & vbCrLf & _
                      "      ,BaseDeCalculo    =  " & _BaseDeCalculo & "" & vbCrLf & _
                      "	  Where Empresa_Id     = '" & _Empresa & "'" & vbCrLf & _
                      "		and EndEmpresa_Id  =  " & _EndEmpresa & " " & vbCrLf & _
                      "		and Deposito_Id    = '" & _Deposito & "'" & vbCrLf & _
                      "		and EndDeposito_Id =  " & _EndDeposito & " " & vbCrLf & _
                      "		and Produto_Id     = '" & _CodigoProduto & "' " & vbCrLf & _
                      "		and Data_Id        = '" & _Data_Id.ToString("yyyy-MM-dd") & "';" & vbCrLf
                Sqls.Add(sql)

            Case "D"
                sql = " Delete TabelaDePrecosDeMercado" & vbCrLf & _
                      "	  Where Empresa_Id     = '" & _Empresa & "'" & vbCrLf & _
                      "		and EndEmpresa_Id  = " & _EndEmpresa & " " & vbCrLf & _
                      "		and Deposito_Id    = '" & _Deposito & "'" & vbCrLf & _
                      "		and EndDeposito_Id = " & _EndDeposito & " " & vbCrLf & _
                      "		and Produto_Id     = '" & _CodigoProduto & "' " & vbCrLf & _
                      "		and Data_Id        = '" & _Data_Id.ToString("yyyy-MM-dd") & "';" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

    Public Function JaExiste() As Boolean
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT 1 " & vbCrLf & _
              "  FROM TabelaDePrecosDeMercado " & vbCrLf & _
              "	  Where Empresa_Id       ='" & _Empresa & "'" & vbCrLf & _
              "		and EndEmpresa_Id    = " & _EndEmpresa & " " & vbCrLf & _
              "		and Deposito_Id      ='" & _Deposito & "'" & vbCrLf & _
              "		and EndDeposito_Id   = " & _EndDeposito & " " & vbCrLf & _
              "		and Produto_Id       = " & _CodigoProduto & " " & vbCrLf & _
              "		and Data_Id          = '" & _Data_Id.ToString("yyyy-MM-dd") & "'" & vbCrLf

        Dim dsExisteFat As DataSet = Banco.ConsultaDataSet(sql, "TabelasDePrecosDeMercado")

        If dsExisteFat.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

End Class