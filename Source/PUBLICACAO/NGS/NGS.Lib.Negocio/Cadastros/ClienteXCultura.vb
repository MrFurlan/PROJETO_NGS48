Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListClienteXCultura
    Inherits List(Of ClienteXCultura)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoCliente As String, ByVal pEnderecoCliente As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As New DataSet
        Dim sql As String

        sql = "SELECT Cliente_Id, EndCliente_Id, Ano_Id, Cultura_Id, AreaPlantada" & vbCrLf & _
              "  FROM ClienteXCultura" & vbCrLf & _
              " Where Cliente_ID    ='" & pCodigoCliente & "'" & vbCrLf & _
              "   AND EndCliente_Id = " & pEnderecoCliente

        ds = Banco.ConsultaDataSet(sql, "CC")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CC As New ClienteXCultura
            CC.CodigoCliente = row("Cliente_Id")
            CC.EnderecoCliente = row("EndCliente_Id")
            CC.Ano = row("Ano_Id")
            CC.CodigoCultura = row("Cultura_Id")
            CC.AreaPlantada = row("AreaPlantada")
            Me.Add(CC)
        Next

    End Sub
#End Region

End Class

<Serializable()> _
Public Class ClienteXCultura
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As Cliente)
        _Cliente = pCliente
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoCliente As String
    Private _EnderecoCliente As Integer
    Private _Cliente As Cliente
    Private _Ano As Integer
    Private _CodigoCultura As Integer
    Private _Cultura As Cultura
    Private _AreaPlantada As Double
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

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property Ano() As Integer
        Get
            Return _Ano
        End Get
        Set(ByVal value As Integer)
            _Ano = value
        End Set
    End Property

    Public Property CodigoCultura() As Integer
        Get
            Return _CodigoCultura
        End Get
        Set(ByVal value As Integer)
            _CodigoCultura = value
        End Set
    End Property

    Public Property Cultura() As Cultura
        Get
            If _Cultura Is Nothing And _CodigoCultura > -1 Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura
        End Get
        Set(ByVal value As Cultura)
            _Cultura = value
        End Set
    End Property

    Public ReadOnly Property NomeCultura() As String
        Get
            If _Cultura Is Nothing And _CodigoCultura > -1 Then _Cultura = New Cultura(_CodigoCultura)
            Return _Cultura.Descricao
        End Get
    End Property

    Public Property AreaPlantada() As Double
        Get
            Return _AreaPlantada
        End Get
        Set(ByVal value As Double)
            _AreaPlantada = value
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
                Sql = " INSERT INTO ClienteXCultura(Cliente_Id, EndCliente_Id, Ano_Id, Cultura_Id, AreaPlantada) " & vbCrLf & _
                      " VALUES ('" & _CodigoCliente & "'," & _EnderecoCliente & "," & _Ano & "," & _CodigoCultura & "," & Str(_AreaPlantada) & ")"

                Sqls.Add(Sql)

                Sql = "UPDATE Clientes "
                Sql &= "SET UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', "
                Sql &= "UsuarioAlteracaoData = '" & Date.Now.ToSqlDate() & "' "
                Sql &= "WHERE Cliente_Id   = '" & _CodigoCliente & "' "
                Sql &= "  AND Endereco_Id  = " & _EnderecoCliente

                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE ClienteXCultura SET" & vbCrLf & _
                      "    AreaPlantada = " & Str(_AreaPlantada) & vbCrLf & _
                      "  WHERE Cliente_Id    ='" & _CodigoCliente & "'" & vbCrLf & _
                      "    AND EndCliente_Id = " & _EnderecoCliente & vbCrLf & _
                      "    AND Ano_Id        = " & _Ano & vbCrLf & _
                      "    AND Cultura_Id    = " & _CodigoCultura
                Sqls.Add(Sql)

                Sql = "UPDATE Clientes "
                Sql &= "SET UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', "
                Sql &= "UsuarioAlteracaoData = '" & Date.Now.ToSqlDate() & "' "
                Sql &= "WHERE Cliente_Id  = '" & _CodigoCliente & "' "
                Sql &= "  AND Endereco_Id = " & _EnderecoCliente

                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE ClienteXCultura" & vbCrLf & _
                      "  WHERE Cliente_Id    ='" & _CodigoCliente & "'" & vbCrLf & _
                      "    AND EndCliente_Id = " & _EnderecoCliente & vbCrLf & _
                      "    AND Ano_Id        = " & _Ano & vbCrLf & _
                      "    AND Cultura_Id    = " & _CodigoCultura
                Sqls.Add(Sql)

                Sql = "UPDATE Clientes "
                Sql &= "SET UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', "
                Sql &= "UsuarioAlteracaoData = '" & Date.Now.ToSqlDate() & "' "
                Sql &= "WHERE Cliente_Id  = '" & _CodigoCliente & "' "
                Sql &= "  AND Endereco_Id = " & _EnderecoCliente

                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class