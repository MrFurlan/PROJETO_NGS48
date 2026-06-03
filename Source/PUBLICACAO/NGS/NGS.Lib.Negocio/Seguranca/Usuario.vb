Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Security.Cryptography
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

Public Class ListUsuario
    Inherits List(Of Usuario)

    Public Function Selecionar() As Boolean
        Dim db As New AcessaBanco()
        Try
            Dim strSQL As String = "SELECT Usuario_Id, NomeCompleto, Empresa, EndEmpresa, COALESCE(MenuDeAcesso, '') AS Menu, Email, " & vbCrLf &
                                   "       ISNULL(Cliente, '') AS Cliente, ISNULL(Endcliente, 0) AS EndCliente, isnull(Ativo,0) as Ativo, " & vbCrLf &
                                   "       isnull(ImprimirDanfe,0) as ImprimirDanfe, ISNULL(Usuarios.LiberaEmpresa, 0) as LiberaEmpresa, " & vbCrLf &
                                   "       ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, ISNULL(LiberaJanela, 0) as LiberaJanela," & vbCrLf &
                                   "       ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                                   "       ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard, ISNULL(SeguimentoDashboard, '') AS SeguimentoDashboard, ISNULL(ConsolidarDashboard, 0) AS ConsolidarDashboard " & vbCrLf &
                                   "  FROM Usuarios " & vbCrLf &
                                   " ORDER BY Usuario_Id"

            Dim dsUsuarios As DataSet = db.ConsultaDataSet(strSQL, "Usuarios")

            For Each drUsuario As DataRow In dsUsuarios.Tables(0).Rows
                Dim objUsuario As New Usuario()
                objUsuario.Usuario_Id = drUsuario("Usuario_Id").ToString().ToUpper()
                objUsuario.NomeCompleto = drUsuario("NomeCompleto").ToString()
                objUsuario.Empresa_Id = drUsuario("Empresa").ToString()
                objUsuario.EnderecoEmpresa = Convert.ToInt32(drUsuario("EndEmpresa"))
                objUsuario.MenuDeAcesso = drUsuario("Menu").ToString()
                objUsuario.Email = drUsuario("EMail").ToString()
                objUsuario.CodigoCliente = drUsuario("CodigoCliente").ToString()
                objUsuario.EndCliente = drUsuario("EndCliente").ToString()
                objUsuario.ImprimirDanfe = drUsuario("ImprimirDanfe")
                objUsuario.UsuarioAtivo = drUsuario("Ativo")
                objUsuario.LiberaEmpresa = drUsuario("LiberaEmpresa")
                objUsuario.TrocaEmpresa = drUsuario("TrocaEmpresa")
                objUsuario.LiberaJanela = drUsuario("LiberaJanela")
                objUsuario.DataInicioDashboard = drUsuario("DataInicioDashboard")
                objUsuario.DataFimDashboard = drUsuario("DataFimDashboard")
                objUsuario.SeguimentoDashboard = drUsuario("SeguimentoDashboard")
                Me.Add(objUsuario)
            Next

            Return True
        Catch ex As Exception
            Return False
        Finally
            db = Nothing
        End Try
    End Function

End Class

Public Class Usuario
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pUsuario As String)
        Dim db As New AcessaBanco
        Dim strSQL As String = "  SELECT Usuario_Id, NomeCompleto, Senha, Empresa, EndEmpresa, COALESCE(MenuDeAcesso, '') as Menu, " & vbCrLf &
                               "         isnull(Email,'') as Email, AcessoUnidade, AcessoEmpresa, AcessoEndEmpresa, isnull(Ativo,0) as Ativo, " & vbCrLf &
                               "         ISNULL(Cliente, '') AS Cliente, ISNULL(Endcliente, 0) AS EndCliente, isnull(ImprimirDanfe,0) as ImprimirDanfe, " & vbCrLf &
                               "         ISNULL(Usuarios.LiberaEmpresa, 0) as LiberaEmpresa, ISNULL(TrocaEmpresa, 0) as TrocaEmpresa, " & vbCrLf &
                               "         ISNULL(LiberaJanela, 0) as LiberaJanela, " & vbCrLf &
                               "         ISNULL(DataInicioDashboard, DATEADD(DAY, -30, GETDATE())) AS DataInicioDashboard," & vbCrLf &
                               "         ISNULL(DataFimDashboard, GETDATE()) AS DataFimDashboard" & vbCrLf &
                               "  FROM Usuarios " & vbCrLf &
                               "  WHERE Usuario_Id = '" & pUsuario & "'"

        Dim dsUsuarios As DataSet = db.ConsultaDataSet(strSQL, "Usuarios")
        If dsUsuarios.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = dsUsuarios.Tables(0).Rows(0)
            _Usuario_Id = row("Usuario_Id").ToString().ToUpper()
            _NomeCompleto = row("NomeCompleto")
            _Senha = row("Senha")
            _Empresa_Id = row("Empresa")
            _EndEmpresa = Convert.ToInt32(row("EndEmpresa"))
            _MenuDeAcesso = row("Menu")
            _Email = row("Email")
            _AcessoUnidade = row("AcessoUnidade")
            _AcessoEmpresa = row("AcessoEmpresa")
            _AcessoEnderecoEmpresa = row("AcessoEndEmpresa")
            _CodigoCliente = row("Cliente")
            _EndCliente = row("EndCliente")
            _ImprimirDanfe = row("ImprimirDanfe")
            _UsuarioAtivo = row("Ativo")
            _LiberaEmpresa = row("LiberaEmpresa")
            _TrocaEmpresa = row("TrocaEmpresa")
            _LiberaJanela = row("LiberaJanela")
            _DataInicioDashboard = row("DataInicioDashboard")
            _DataFimDashboard = row("DataFimDashboard")
        End If

    End Sub
#End Region

#Region "Fields"

    Private _IUD As String
    Private _Usuario_Id As String
    Private _NomeCompleto As String
    Private _Senha As String
    Private _Empresa_Id As String
    Private _EndEmpresa As Integer
    Private _MenuDeAcesso As String
    Private _Email As String
    Private _AcessoUnidade As String
    Private _AcessoEmpresa As String
    Private _AcessoEnderecoEmpresa As Integer
    Private _TrocarSenha As Boolean
    Private _Empresa As Cliente
    Private _CodigoCliente As String
    Private _EndCliente As Integer
    Private _Cliente As Cliente
    Private _ImprimirDanfe As Boolean = False
    Private _UsuarioAtivo As Boolean = False
    Private _LiberaEmpresa As Boolean = False
    Private _TrocaEmpresa As Boolean = False
    Private _LiberaJanela As Boolean = False
    Private _DataInicioDashboard As DateTime = Date.Now.AddDays(-30)
    Private _DataFimDashboard As DateTime = Date.Now
    Private _SeguimentoDashboard As String
    Private _EmpresasDashboard As String

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

    Public Property Usuario_Id() As String
        Get
            Return _Usuario_Id
        End Get
        Set(ByVal value As String)
            _Usuario_Id = value
        End Set
    End Property

    Public Property Empresa_Id() As String
        Get
            Return _Empresa_Id
        End Get
        Set(ByVal value As String)
            _Empresa_Id = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property MenuDeAcesso() As String
        Get
            Return _MenuDeAcesso
        End Get
        Set(ByVal value As String)
            _MenuDeAcesso = value
        End Set
    End Property

    Public Property Email() As String
        Get
            Return _Email
        End Get
        Set(ByVal value As String)
            _Email = value
        End Set
    End Property

    Public Property NomeCompleto() As String
        Get
            Return _NomeCompleto
        End Get
        Set(ByVal value As String)
            _NomeCompleto = value
        End Set
    End Property

    Public Property Senha() As String
        Get
            Return _Senha
        End Get
        Set(ByVal value As String)
            _Senha = value
        End Set
    End Property

    Public Property AcessoUnidade() As String
        Get
            Return _AcessoUnidade
        End Get
        Set(ByVal value As String)
            _AcessoUnidade = value
        End Set
    End Property

    Public Property AcessoEmpresa() As String
        Get
            Return _AcessoEmpresa
        End Get
        Set(ByVal value As String)
            _AcessoEmpresa = value
        End Set
    End Property

    Public Property AcessoEnderecoEmpresa() As Integer
        Get
            Return _AcessoEnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _AcessoEnderecoEmpresa = value
        End Set
    End Property

    Public Property TrocarSenha() As Boolean
        Get
            Return _TrocarSenha
        End Get
        Set(ByVal value As Boolean)
            _TrocarSenha = value
        End Set
    End Property

    Public ReadOnly Property Empresa As Cliente
        Get
            If _Empresa Is Nothing Then _Empresa = New Cliente(Empresa_Id, EnderecoEmpresa)
            Return _Empresa
        End Get
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property

    Public ReadOnly Property Cliente As Cliente
        Get
            If _Cliente Is Nothing Then _Cliente = New Cliente(CodigoCliente, EndCliente)
            Return _Cliente
        End Get
    End Property

    Public Property ImprimirDanfe As Boolean
        Get
            Return _ImprimirDanfe
        End Get
        Set(value As Boolean)
            _ImprimirDanfe = value
        End Set
    End Property

    Public Property UsuarioAtivo As Boolean
        Get
            Return _UsuarioAtivo
        End Get
        Set(value As Boolean)
            _UsuarioAtivo = value
        End Set
    End Property

    Public Property LiberaEmpresa As Boolean
        Get
            Return _LiberaEmpresa
        End Get
        Set(value As Boolean)
            _LiberaEmpresa = value
        End Set
    End Property

    Public Property TrocaEmpresa As Boolean
        Get
            Return _TrocaEmpresa
        End Get
        Set(value As Boolean)
            _TrocaEmpresa = value
        End Set
    End Property

    Public Property LiberaJanela As Boolean
        Get
            Return _LiberaJanela
        End Get
        Set(value As Boolean)
            _LiberaJanela = value
        End Set
    End Property

    Public Property DataInicioDashboard As DateTime
        Get
            Return _DataInicioDashboard
        End Get
        Set(value As DateTime)
            _DataInicioDashboard = value
        End Set
    End Property

    Public Property DataFimDashboard As DateTime
        Get
            Return _DataFimDashboard
        End Get
        Set(value As DateTime)
            _DataFimDashboard = value
        End Set
    End Property

    Public Property SeguimentoDashboard As String
        Get
            Return _SeguimentoDashboard
        End Get
        Set(value As String)
            _SeguimentoDashboard = value
        End Set
    End Property

    Public Property EmpresasDashboard As String
        Get
            Return _EmpresasDashboard
        End Get
        Set(value As String)
            _EmpresasDashboard = value
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
        End If
        Return False
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)

        Dim sql As String
        Select Case Me.IUD

            Case "I"
                sql = "INSERT INTO Usuarios" & vbCrLf &
                "(Usuario_Id" & vbCrLf &
                ",NomeCompleto " & vbCrLf &
                ",Senha " & vbCrLf &
                ",Empresa " & vbCrLf &
                ",EndEmpresa " & vbCrLf &
                ",MenuDeAcesso " & vbCrLf &
                ",Email " & vbCrLf &
                ",AcessoUnidade " & vbCrLf &
                ",AcessoEmpresa " & vbCrLf &
                ",AcessoEndEmpresa" & vbCrLf &
                ",ImprimirDanfe" & vbCrLf &
                ",LiberaEmpresa" & vbCrLf &
                ",TrocaEmpresa" & vbCrLf &
                ",LiberaJanela" & vbCrLf &
                ",Ativo) " & vbCrLf &
                "VALUES " & vbCrLf &
                "('" & Me.Usuario_Id & "' " & vbCrLf &
                ",'" & Me.NomeCompleto & "' " & vbCrLf &
                ",'" & Me.Senha & "' " & vbCrLf &
                ",'" & Me.Empresa_Id & "' " & vbCrLf &
                ",'" & Me.EnderecoEmpresa & "' " & vbCrLf &
                ",'" & Me.MenuDeAcesso & "' " & vbCrLf &
                ",'" & Me.Email & "' " & vbCrLf &
                ",'" & Me.AcessoUnidade & "' " & vbCrLf &
                ",'" & Me.AcessoEmpresa & "' " & vbCrLf &
                ",'" & Me.AcessoEnderecoEmpresa & "'" & vbCrLf &
                ",'" & Me.ImprimirDanfe.ToString & "'" & vbCrLf &
                ",'" & Me.LiberaEmpresa.ToString & "'" & vbCrLf &
                ",'" & Me.TrocaEmpresa.ToString & "'" & vbCrLf &
                ",'" & Me.LiberaJanela.ToString & "'" & vbCrLf &
                ",'" & Me.UsuarioAtivo.ToString & "')"
                Sqls.Add(sql)

            Case "U"
                sql = "UPDATE Usuarios SET " & vbCrLf &
                      "   NomeCompleto                      = '" & Me.NomeCompleto & "', " & vbCrLf &
                      "   Empresa                           = '" & Me.Empresa_Id & "', " & vbCrLf &
                      "   EndEmpresa                        = '" & Me.EnderecoEmpresa & "', " & vbCrLf &
                      "   MenuDeAcesso                      = '" & Me.MenuDeAcesso & "', " & vbCrLf &
                      "   Email                             = '" & Me.Email & "', " & vbCrLf &
                      "   AcessoUnidade                     = '" & Me.AcessoUnidade & "', " & vbCrLf &
                      "   AcessoEmpresa                     = '" & Me.AcessoEmpresa & "', " & vbCrLf &
                      "   AcessoEndEmpresa                  = '" & Me.AcessoEnderecoEmpresa & "'," & vbCrLf &
                      "   ImprimirDanfe                     = '" & Me.ImprimirDanfe & "'," & vbCrLf &
                      "   LiberaEmpresa                     = '" & Me.LiberaEmpresa & "'," & vbCrLf &
                      "   TrocaEmpresa                      = '" & Me.TrocaEmpresa & "'," & vbCrLf &
                      "   LiberaJanela                      = '" & Me.LiberaJanela & "'," & vbCrLf &
                      "   Ativo                             = '" & Me.UsuarioAtivo & "'" & vbCrLf &
                      " WHERE Usuario_Id                    = '" & Me.Usuario_Id & "'"
                Sqls.Add(sql)

            Case "U_DASHBOARD"
                sql = "UPDATE Usuarios SET " & vbCrLf &
                      "   DataInicioDashboard               = '" & Me.DataInicioDashboard.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "   DataFimDashboard                  = '" & Me.DataFimDashboard.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "   SeguimentoDashboard               = '" & Me.SeguimentoDashboard & "'," & vbCrLf &
                      "   EmpresasDashboard                 = '" & Me.EmpresasDashboard & "'" & vbCrLf &
                      " WHERE Usuario_Id                    = '" & Me.Usuario_Id & "'"
                Sqls.Add(sql)

            Case "D"
                sql = "DELETE FROM Usuarios WHERE Usuario_Id = '" & Me.Usuario_Id & "'"
                Sqls.Add(sql)
        End Select

    End Sub

#End Region

End Class
