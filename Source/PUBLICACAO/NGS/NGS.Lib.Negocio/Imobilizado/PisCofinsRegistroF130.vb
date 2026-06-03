Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()>
Public Class lstPisCofinsRegistrosF130
    Inherits List(Of PisCofinsRegistroF130)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()
        Dim objBanco As New AcessaBanco()
        Try
            Dim strSql As String = " SELECT Empresa_Id, Movimento_Id, Situacao, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, " & vbCrLf &
                                   " Mes_Ope_Aquis, Vl_Oper_Aquis, Parc_Oper_Nao_Bc_Cred, Vl_Bc_Cred, Ind_Nr_Parc, Cst_Pis, Vl_Bc_Pis, " & vbCrLf &
                                   " Aliq_Pis, Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob, " & vbCrLf &
                                   " UsuarioInclusao, UsuarioInclusaoData, ISNULL(UsuarioAlteracao, 0) as UsuarioAlteracao, ISNULL(UsuarioAlteracaoData, 0) as UsuarioAlteracaoData " & vbCrLf &
                                   "  FROM PisCofins_Reg_F130 " & vbCrLf &
                                   " WHERE Situacao = 1" & vbCrLf &
                                   " ORDER BY Movimento_Id DESC"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "PisCofins_Reg_F130")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim F130 As New PisCofinsRegistroF130

                F130.EmpresaId = row("Empresa_Id")
                F130.MovimentoId = row("Movimento_Id")
                F130.Situacao = row("Situacao")
                F130.NatBcCred = row("Nat_Bc_Cred")
                F130.IdentBemImob = row("Ident_Bem_Imob")
                F130.IndOrigCred = row("Ind_Orig_Cred")
                F130.IndUtilBemImob = row("Ind_Util_Bem_Imob")
                F130.MesOpeAquis = row("Mes_Ope_Aquis")
                F130.VlOperAquis = row("Vl_Oper_Aquis")
                F130.ParcOperNaoBcCred = row("Parc_Oper_Nao_Bc_Cred")
                F130.VlBcCred = row("Vl_Bc_Cred")
                F130.IndNrParc = row("Ind_Nr_Parc")
                F130.Cst = row("Cst_Pis")
                F130.AliqPis = row("Aliq_Pis")
                F130.VlPis = row("Vl_Pis")
                F130.AliqCofins = row("Aliq_Cofins")
                F130.VlCofins = row("Vl_Cofins")

                If Not IsDBNull(row("Cod_Cta")) Then
                    F130.CodCta = row("Cod_Cta")
                End If

                If Not IsDBNull(row("Cod_Ccus")) Then
                    F130.CodCcus = row("Cod_Ccus")
                End If

                If Not IsDBNull(row("Desc_Bem_Imob")) Then
                    F130.DescBemImob = row("Desc_Bem_Imob")
                End If

                If Not IsDBNull(row("UsuarioInclusao")) Then
                    F130.UsuarioInclusao = row("UsuarioInclusao")
                End If

                If Not IsDBNull(row("UsuarioInclusaoData")) Then
                    F130.UsuarioInclusaoData = row("UsuarioInclusaoData")
                End If

                If Not IsDBNull(row("UsuarioAlteracao")) Then
                    F130.UsuarioAlteracao = row("UsuarioAlteracao")
                End If

                If Not IsDBNull(row("UsuarioAlteracaoData")) Then
                    F130.UsuarioAlteracaoData = row("UsuarioAlteracaoData")
                End If

                Add(F130)
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub

    Public Sub New(ByVal empresa As String, ByVal endEmpresa As String, ByVal inicial As String, ByVal final As String)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSql As String = " SELECT Empresa_Id, Movimento_Id, Cod_Cta_Id, Situacao, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, " & vbCrLf &
                                   " Mes_Ope_Aquis, Vl_Oper_Aquis, Parc_Oper_Nao_Bc_Cred, Vl_Bc_Cred, Ind_Nr_Parc, Cst_Pis, Vl_Bc_Pis, " & vbCrLf &
                                   " Aliq_Pis, Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob, " & vbCrLf &
                                   " UsuarioInclusao, UsuarioInclusaoData, ISNULL(UsuarioAlteracao, 0) as UsuarioAlteracao, ISNULL(UsuarioAlteracaoData, 0) as UsuarioAlteracaoData " & vbCrLf &
                                   "  FROM PisCofins_Reg_F130 " & vbCrLf &
                                   " WHERE Situacao = 1" & vbCrLf &
                                   " AND Empresa_Id = '" & empresa & "'" & vbCrLf &
                                   " AND EndEmpresa_Id = " & endEmpresa & vbCrLf &
                                   " AND Movimento_Id BETWEEN '" & CDate(inicial).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                   " AND '" & CDate(final).ToString("yyyy-MM-dd") & "'" & vbCrLf &
                                   " ORDER BY Movimento_Id DESC"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "PisCofins_Reg_F130")
            For Each row As DataRow In ds.Tables(0).Rows
                Dim F130 As New PisCofinsRegistroF130

                F130.EmpresaId = row("Empresa_Id")
                F130.MovimentoId = row("Movimento_Id")
                F130.Situacao = row("Situacao")
                F130.NatBcCred = row("Nat_Bc_Cred")
                F130.IdentBemImob = row("Ident_Bem_Imob")
                F130.IndOrigCred = row("Ind_Orig_Cred")
                F130.IndUtilBemImob = row("Ind_Util_Bem_Imob")
                F130.MesOpeAquis = row("Mes_Ope_Aquis")
                F130.VlOperAquis = row("Vl_Oper_Aquis")
                F130.ParcOperNaoBcCred = row("Parc_Oper_Nao_Bc_Cred")
                F130.VlBcCred = row("Vl_Bc_Cred")
                F130.IndNrParc = row("Ind_Nr_Parc")
                F130.Cst = row("Cst_Pis")
                F130.AliqPis = row("Aliq_Pis")
                F130.VlPis = row("Vl_Pis")
                F130.AliqCofins = row("Aliq_Cofins")
                F130.VlCofins = row("Vl_Cofins")

                If Not IsDBNull(row("Cod_Cta")) Then
                    F130.CodCta = row("Cod_Cta")
                End If

                If Not IsDBNull(row("Cod_Ccus")) Then
                    F130.CodCcus = row("Cod_Ccus")
                End If

                If Not IsDBNull(row("Desc_Bem_Imob")) Then
                    F130.DescBemImob = row("Desc_Bem_Imob")
                End If

                F130.UsuarioInclusao = row("UsuarioInclusao")
                F130.UsuarioInclusaoData = row("UsuarioInclusaoData")
                F130.UsuarioAlteracao = row("UsuarioAlteracao")
                F130.UsuarioAlteracaoData = row("UsuarioAlteracaoData")

                Add(F130)
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

End Class


<Serializable()>
Public Class PisCofinsRegistroF130
    Implements IBaseEntity

#Region "Constructor"

    Public Sub New()

    End Sub

    Public Sub New(Optional ByVal Empresa As String = "", Optional ByVal EndEmpresa As Integer = 0, Optional ByVal Movimento As String = "", Optional ByVal Conta As String = "")
        Dim objBanco As New AcessaBanco
        Try
            Dim strSql As String = " SELECT Empresa_Id, Movimento_Id, Situacao, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, " & vbCrLf &
                                   " Mes_Ope_Aquis, Vl_Oper_Aquis, Parc_Oper_Nao_Bc_Cred, Vl_Bc_Cred, Ind_Nr_Parc, Cst_Pis, Vl_Bc_Pis, " & vbCrLf &
                                   " Aliq_Pis, Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob, " & vbCrLf &
                                   " UsuarioInclusao, UsuarioInclusaoData, ISNULL(UsuarioAlteracao, 0) as UsuarioAlteracao, ISNULL(UsuarioAlteracaoData, 0) as UsuarioAlteracaoData " & vbCrLf &
                                   "  FROM PisCofins_Reg_F130 " & vbCrLf &
                                   " WHERE Situacao = 1" & vbCrLf

            If Empresa IsNot Nothing Then
                strSql &= " AND Empresa_Id  = '" & Empresa & "'" & vbCrLf
            End If
            If Not EndEmpresa = 0 Then
                strSql &= " AND EndEmpresa_Id =  " & EndEmpresa & vbCrLf
            End If
            If Movimento IsNot Nothing Then
                strSql &= " AND Movimento_Id  = '" & CDate(Movimento).ToString("yyyy-MM-dd") & "'" & vbCrLf
            End If
            If Conta IsNot Nothing Then
                strSql &= " AND Cod_Cta_Id    = '" & Conta & "'" & vbCrLf
            End If
            strSql &= " ORDER BY Movimento_Id DESC"

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSql, "PisCofins_Reg_F130")
            For Each row As DataRow In ds.Tables(0).Rows
                _EmpresaId = row("Empresa_Id")
                _MovimentoId = CDate(row("Movimento_Id"))
                _Situacao = row("Situacao")
                _NatBcCred = row("Nat_Bc_Cred")
                _IdentBemImob = row("Ident_Bem_Imob")
                _IndOrigCred = row("Ind_Orig_Cred")
                _IndUtilBemImob = row("Ind_Util_Bem_Imob")
                _MesOpeAquis = row("Mes_Ope_Aquis")
                _VlOperAquis = row("Vl_Oper_Aquis")
                _ParcOperNaoBcCred = row("Parc_Oper_Nao_Bc_Cred")
                _VlBcCred = row("Vl_Bc_Cred")
                _IndNrParc = row("Ind_Nr_Parc")
                _Cst = row("Cst_Pis")
                _AliqPis = row("Aliq_Pis")
                _VlPis = row("Vl_Pis")
                _AliqCofins = row("Aliq_Cofins")
                _VlCofins = row("Vl_Cofins")

                If Not IsDBNull(row("Cod_Cta")) Then
                    _CodCta = row("Cod_Cta")
                End If

                If Not IsDBNull(row("Cod_Ccus")) Then
                    _CodCcus = row("Cod_Ccus")
                End If

                If Not IsDBNull(row("Desc_Bem_Imob")) Then
                    _DescBemImob = row("Desc_Bem_Imob")
                End If

                _UsuarioInclusao = row("UsuarioInclusao")
                _UsuarioInclusaoData = row("UsuarioInclusaoData")
                _UsuarioAlteracao = row("UsuarioAlteracao")
                _UsuarioAlteracaoData = row("UsuarioAlteracaoData")
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _EmpresaId As String
    Private _MovimentoId As Date
    Private _Situacao As Integer
    Private _NatBcCred As String
    Private _IdentBemImob As Integer
    Private _IndOrigCred As Integer
    Private _IndUtilBemImob As Integer
    Private _MesOpeAquis As String = "0"
    Private _ParcOperNaoBcCred As Integer
    Private _IndNrParc As Integer
    Private _VlOperAquis As String
    Private _VlBcCred As String
    Private _Cst As Integer
    Private _AliqPis As String
    Private _VlPis As String
    Private _AliqCofins As String
    Private _VlCofins As String
    Private _CodCta As String
    Private _CodCcus As String
    Private _DescBemImob As String
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As String
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As String
    Public Erro As Exception

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

    Public Property EmpresaId() As String
        Get
            Return _EmpresaId
        End Get
        Set(ByVal value As String)
            _EmpresaId = value
        End Set
    End Property

    Public Property MovimentoId() As Date
        Get
            Return _MovimentoId
        End Get
        Set(ByVal value As Date)
            _MovimentoId = value
        End Set
    End Property

    Public Property Situacao() As Integer
        Get
            Return _Situacao
        End Get
        Set(ByVal value As Integer)
            _Situacao = value
        End Set
    End Property

    Public Property NatBcCred() As String
        Get
            Return _NatBcCred
        End Get
        Set(ByVal value As String)
            _NatBcCred = value
        End Set
    End Property

    Public Property IdentBemImob() As Integer
        Get
            Return _IdentBemImob
        End Get
        Set(ByVal value As Integer)
            _IdentBemImob = value
        End Set
    End Property

    Public Property IndOrigCred() As Integer
        Get
            Return _IndOrigCred
        End Get
        Set(ByVal value As Integer)
            _IndOrigCred = value
        End Set
    End Property

    Public Property IndUtilBemImob() As Integer
        Get
            Return _IndUtilBemImob
        End Get
        Set(ByVal value As Integer)
            _IndUtilBemImob = value
        End Set
    End Property

    Public Property MesOpeAquis() As String
        Get
            Return _MesOpeAquis
        End Get
        Set(ByVal value As String)
            _MesOpeAquis = value
        End Set
    End Property

    Public Property VlOperAquis() As String
        Get
            Return _VlOperAquis
        End Get
        Set(ByVal value As String)
            _VlOperAquis = value
        End Set
    End Property

    Public Property ParcOperNaoBcCred() As Integer
        Get
            Return _ParcOperNaoBcCred
        End Get
        Set(ByVal value As Integer)
            _ParcOperNaoBcCred = value
        End Set
    End Property

    Public Property VlBcCred() As String
        Get
            Return _VlBcCred
        End Get
        Set(ByVal value As String)
            _VlBcCred = value
        End Set
    End Property

    Public Property IndNrParc() As Integer
        Get
            Return _IndNrParc
        End Get
        Set(ByVal value As Integer)
            _IndNrParc = value
        End Set
    End Property

    Public Property Cst() As Integer
        Get
            Return _Cst
        End Get
        Set(ByVal value As Integer)
            _Cst = value
        End Set
    End Property

    Public Property AliqPis() As String
        Get
            Return _AliqPis
        End Get
        Set(ByVal value As String)
            _AliqPis = value
        End Set
    End Property

    Public Property VlPis() As String
        Get
            Return _VlPis
        End Get
        Set(ByVal value As String)
            _VlPis = value
        End Set
    End Property

    Public Property AliqCofins() As String
        Get
            Return _AliqCofins
        End Get
        Set(ByVal value As String)
            _AliqCofins = value
        End Set
    End Property

    Public Property VlCofins() As String
        Get
            Return _VlCofins
        End Get
        Set(ByVal value As String)
            _VlCofins = value
        End Set
    End Property

    Public Property CodCta() As String
        Get
            Return _CodCta
        End Get
        Set(ByVal value As String)
            _CodCta = value
        End Set
    End Property

    Public Property CodCcus() As String
        Get
            Return _CodCcus
        End Get
        Set(ByVal value As String)
            _CodCcus = value
        End Set
    End Property

    Public Property DescBemImob() As String
        Get
            Return _DescBemImob
        End Get
        Set(ByVal value As String)
            _DescBemImob = value
        End Set
    End Property

    Public Property UsuarioInclusao As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property UsuarioInclusaoData As String
        Get
            Return _UsuarioInclusaoData
        End Get
        Set(value As String)
            _UsuarioInclusaoData = value
        End Set
    End Property

    Public Property UsuarioAlteracao As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property UsuarioAlteracaoData As String
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(value As String)
            _UsuarioAlteracaoData = value
        End Set
    End Property
#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim objBanco As New AcessaBanco()
        Dim Sqls As New ArrayList

        SalvarSql(Sqls)
        Return objBanco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO PisCofins_Reg_F130 " & vbCrLf &
                                  " (Empresa_Id, " & vbCrLf &
                                  " Movimento_Id, " & vbCrLf &
                                  " Situacao, " & vbCrLf &
                                  " Nat_Bc_Cred, " & vbCrLf &
                                  " Ident_Bem_Imob, " & vbCrLf &
                                  " Ind_Orig_Cred, " & vbCrLf &
                                  " Ind_Util_Bem_Imob, " & vbCrLf &
                                  " Mes_Ope_Aquis, " & vbCrLf &
                                  " Vl_Oper_Aquis, " & vbCrLf &
                                  " Parc_Oper_Nao_Bc_Cred," & vbCrLf &
                                  " Vl_Bc_Cred, " & vbCrLf &
                                  " Ind_Nr_Parc, " & vbCrLf &
                                  " Cst_Pis, " & vbCrLf &
                                  " Vl_Bc_Pis, " & vbCrLf &
                                  " Aliq_Pis, " & vbCrLf &
                                  " Vl_Pis, " & vbCrLf &
                                  " Cst_Cofins, " & vbCrLf &
                                  " Vl_Bc_Cofins, " & vbCrLf &
                                  " Aliq_Cofins, " & vbCrLf &
                                  " Vl_Cofins, " & vbCrLf &
                                  " Cod_Cta, " & vbCrLf &
                                  " Cod_Ccus, " & vbCrLf &
                                  " Desc_Bem_Imob, " & vbCrLf &
                                  " UsuarioInclusao, " & vbCrLf &
                                  " UsuarioInclusaoData)" & vbCrLf &
                           " VALUES ('" & _EmpresaId & "'," & vbCrLf &
                                    "'" & _MovimentoId.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                                          _Situacao & "," & vbCrLf &
                                    "'" & _NatBcCred & "'," & vbCrLf &
                                          _IdentBemImob & "," & vbCrLf &
                                          _IndOrigCred & "," & vbCrLf &
                                          _IndUtilBemImob & "," & vbCrLf &
                                    "'" & _MesOpeAquis & "'," & vbCrLf &
                                          _VlOperAquis & "," & vbCrLf &
                                          _ParcOperNaoBcCred & "," & vbCrLf &
                                          _VlBcCred & "," & vbCrLf &
                                          _IndNrParc & "," & vbCrLf &
                                          _Cst & "," & vbCrLf &
                                          _VlBcCred & "," & vbCrLf &
                                          _AliqPis & "," & vbCrLf &
                                          _VlPis & "," & vbCrLf &
                                          _Cst & "," & vbCrLf &
                                          _VlBcCred & "," & vbCrLf &
                                          _AliqCofins & "," & vbCrLf &
                                          _VlCofins & "," & vbCrLf &
                                    "'" & _CodCta & "'," & vbCrLf &
                                    "'" & _CodCcus & "'," & vbCrLf &
                                    "'" & _DescBemImob & "'," & vbCrLf &
                                    "'" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                                    "'" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "')"
                Sqls.Add(Sql)

            Case "U"
                Sql = " UPDATE PisCofins_Reg_F130 SET" & vbCrLf &
                      "    Nat_Bc_Cred           = '" & _NatBcCred & "'," & vbCrLf &
                      "    Ident_Bem_Imob        =  " & _IdentBemImob & "," & vbCrLf &
                      "    Ind_Orig_Cred         =  " & _IndOrigCred & "," & vbCrLf &
                      "    Ind_Util_Bem_Imob     =  " & _IndUtilBemImob & "," & vbCrLf &
                      "    Mes_Ope_Aquis         = '" & _MesOpeAquis & "'," & vbCrLf &
                      "    Vl_Oper_Aquis         = '" & _VlOperAquis & "'," & vbCrLf &
                      "    Parc_Oper_Nao_Bc_Cred = '" & _ParcOperNaoBcCred & "'," & vbCrLf &
                      "    Vl_Bc_Cred            = '" & _VlBcCred & "'," & vbCrLf &
                      "    Ind_Nr_Parc           =  " & _IndNrParc & "," & vbCrLf &
                      "    Cst_Pis               =  " & _Cst & "," & vbCrLf &
                      "    Vl_Bc_Pis             = '" & _VlBcCred & "'," & vbCrLf &
                      "    Aliq_Pis              = '" & _AliqPis & "'," & vbCrLf &
                      "    Vl_Pis                = '" & _VlPis & "'," & vbCrLf &
                      "    Cst_Cofins            =  " & _Cst & "," & vbCrLf &
                      "    Vl_Bc_Cofins          = '" & _VlBcCred & "'," & vbCrLf &
                      "    Aliq_Cofins           = '" & _AliqCofins & "'," & vbCrLf &
                      "    Vl_Cofins             = '" & _VlCofins & "'," & vbCrLf &
                      "    Cod_Cta               =  " & _CodCta & "," & vbCrLf &
                      "    Cod_Ccus              =  " & _CodCcus & "," & vbCrLf &
                      "    Desc_Bem_Imob         = '" & _DescBemImob & "'," & vbCrLf &
                      "    UsuarioAlteracao      = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                      "    UsuarioAlteracaoData  = '" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "'" & vbCrLf &
                      "  WHERE Empresa_Id        = '" & _EmpresaId & "'" & vbCrLf &
                      "    AND Movimento_Id      = '" & _MovimentoId.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "    AND Cod_Cta           = '" & _CodCta & "'"

                Sqls.Add(Sql)

            Case "D"
                Sql = " UPDATE PisCofins_Reg_F130 SET" & vbCrLf &
                      "  Situacao              = 3, " & vbCrLf &
                      "  UsuarioAlteracao      = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "  UsuarioAlteracaoData  = '" & Now().ToString("yyyy-MM-dd hh:mm:ss") & "'" & vbCrLf &
                      "  WHERE Empresa_Id      = '" & _EmpresaId & "'" & vbCrLf &
                      "    AND Movimento_Id    = '" & _MovimentoId.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "    AND Cod_Cta         = '" & _CodCta & "'" & vbCrLf &
                      "    AND Vl_Oper_Aquis   = " & _VlOperAquis

                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal MovimentoId As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, EndEmpresa_Id, Movimento_Id, Cod_Cta_Id, Nat_Bc_Cred, Ident_Bem_Imob, Ind_Orig_Cred, Ind_Util_Bem_Imob, " &
                                   "Mes_Ope_Aquis, Vl_Oper_Aquis, Parc_Oper_Nao_Bc_Cred, Vl_Bc_Cred, Ind_Nr_Parc, Cst_Pis, Vl_Bc_Pis, " &
                                   "Aliq_Pis, Vl_Pis, Cst_Cofins, Vl_Bc_Cofins, Aliq_Cofins, Vl_Cofins, Cod_Cta, Cod_Ccus, Desc_Bem_Imob " &
                                   "  FROM PisCofins_Reg_F130 " &
                                   " WHERE Movimento_Id = '" & MovimentoId & "'"

            Dim F130Imobilizado As DataSet = objBanco.ConsultaDataSet(strSQL, "PisCofins_Reg_F130")

            With F130Imobilizado.Tables(0).Rows
                If .Count > 0 Then
                    Me.EmpresaId = .Item(0)("Empresa_Id").ToString()
                    Me.MovimentoId = .Item(0)("Movimento_Id").ToString()
                    Me.NatBcCred = .Item(0)("Nat_Bc_Cred").ToString()
                    Me.IdentBemImob = .Item(0)("Ident_Bem_Imob").ToString()
                    Me.IndOrigCred = .Item(0)("Ind_Orig_Cred").ToString()
                    Me.IndUtilBemImob = .Item(0)("Ind_Util_Bem_Imob").ToString()
                    Me.MesOpeAquis = .Item(0)("Mes_Ope_Aquis").ToString()
                    Me.VlOperAquis = .Item(0)("Vl_Oper_Aquis").ToString()
                    Me.ParcOperNaoBcCred = .Item(0)("Parc_Oper_Nao_Bc_Cred").ToString()
                    Me.VlBcCred = .Item(0)("Vl_Bc_Cred").ToString()
                    Me.IndNrParc = .Item(0)("Ind_Nr_Parc").ToString()
                    Me.Cst = .Item(0)("Cst_Pis").ToString()
                    Me.AliqPis = .Item(0)("Aliq_Pis").ToString()
                    Me.VlPis = .Item(0)("Vl_Pis").ToString()
                    Me.AliqCofins = .Item(0)("Aliq_Cofins").ToString()
                    Me.VlCofins = .Item(0)("Vl_Cofins").ToString()
                    Me.CodCta = .Item(0)("Cod_Cta").ToString()
                    Me.CodCcus = .Item(0)("Cod_Ccus").ToString()
                    Me.DescBemImob = .Item(0)("Desc_Bem_Imob").ToString()

                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

#End Region

End Class