Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis


'*********************************************************************************************************
'*************************  LISTA DE OPERACAO X ESTADO  **************************************************
'*********************************************************************************************************
Public Class ListOperacaoXEstado
    Inherits List(Of OperacaoXEstado)
#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(Parametros As OperacaoXEstado)
        Dim sql As String

        sql = " Select OE.Codigo_Id, OE.Ativo, OE.InicioVigencia," & vbCrLf &
              "        OE.Empresa, OE.GrupoProduto, OE.Produto," & vbCrLf &
              "        OE.Operacao, OE.SubOperacao," & vbCrLf &
              "        OE.EstadoOrigem, OE.EstadoDestino," & vbCrLf &
              "        OE.GrupoFiscal, OE.CodigoFiscal," & vbCrLf &
              "        isnull(OE.STICMS,-1) as STICMS, isnull(OE.ObsICMS,-1) as ObsICMS," & vbCrLf &
              "        isnull(OE.STIPI,-1) as STIPI, " & vbCrLf &
              "        isnull(OE.STPISCOFINS,0) as STPISCOFINS, isnull(OE.ObsPISCOFINS,-1) as ObsPISCOFINS," & vbCrLf &
              "        isnull(OE.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento," & vbCrLf &
              "        OE.UsuarioInclusao, OE.UsuarioInclusaoData, isnull(OE.CodigoBeneficio,'') AS CodigoBeneficio" & vbCrLf &
              "   FROM OperacaoXEstado OE" & vbCrLf

        If Parametros.Codigo > 0 Then
            sql &= " Inner join OperacaoXEstado OE2" & vbCrLf &
                   "    on OE2.GrupoProduto  = OE.GrupoProduto" & vbCrLf &
                   "   and OE2.Produto       = OE.Produto" & vbCrLf &
                   "   and OE2.Operacao      = OE.Operacao" & vbCrLf &
                   "   and OE2.SubOperacao   = OE.SubOperacao" & vbCrLf &
                   "   and OE2.EstadoOrigem  = OE.EstadoOrigem" & vbCrLf &
                   "   and OE2.EstadoDestino = OE.EstadoDestino" & vbCrLf &
                   "   and OE2.Codigo_id     = " & Parametros.Codigo & vbCrLf

            sql &= "WHERE (OE.Empresa = '99999999' OR OE.Empresa = '" & Parametros.Empresa & "')" & vbCrLf

        Else
            sql &= " Where 1 = 1"


            If Parametros.Ativo = 1 Then
                sql &= "   And OE.Ativo =1" & vbCrLf
            End If

            sql &= "   And (OE.Empresa = '99999999' OR OE.Empresa = '" & Parametros.Empresa & "')" & vbCrLf

            If Parametros.CodigoGrupoProduto.Length > 0 Then
                sql &= "   And OE.GrupoProduto ='" & Parametros.CodigoGrupoProduto & "'" & vbCrLf

                If Parametros.CodigoProduto.Length > 0 Then
                    sql &= "   And OE.Produto ='" & Parametros.CodigoProduto & "'" & vbCrLf
                Else
                    sql &= "   And OE.Produto =''" & vbCrLf
                End If
            Else
                If Parametros.CodigoProduto.Length > 0 Then sql &= "   And OE.Produto ='" & Parametros.CodigoProduto & "'" & vbCrLf
            End If

            If Parametros.CodigoOperacao > 0 Then sql &= "   And OE.Operacao = " & Parametros.CodigoOperacao & vbCrLf
            If Parametros.CodigoSubOperacao > 0 Then sql &= "   And OE.SubOperacao = " & Parametros.CodigoSubOperacao & vbCrLf
            If Parametros.EstadoOrigem.Length > 0 Then sql &= "   And OE.EstadoOrigem ='" & Parametros.EstadoOrigem & "'" & vbCrLf

            If Parametros.RegiaoDestino.Length > 0 Then
                If Parametros.EstadoDestino.Length > 0 Then
                    sql &= "   And (OE.EstadoDestino ='" & Parametros.RegiaoDestino & "' OR OE.EstadoDestino ='" & Parametros.EstadoDestino & "')" & vbCrLf
                Else
                    sql &= "   And OE.EstadoDestino ='" & Parametros.RegiaoDestino & "'" & vbCrLf
                End If
            Else
                If Parametros.EstadoDestino.Length > 0 Then sql &= "   And OE.EstadoDestino ='" & Parametros.EstadoDestino & "'" & vbCrLf
            End If
        End If

        sql &= " Order by OE.InicioVigencia, OE.Codigo_Id"

        Dim objBanco As New AcessaBanco()
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "OpXUf")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            sql = " Select OE.Codigo_Id, OE.Ativo, OE.InicioVigencia," & vbCrLf &
            "        OE.Empresa, OE.GrupoProduto, OE.Produto," & vbCrLf &
            "        OE.Operacao, OE.SubOperacao," & vbCrLf &
            "        OE.EstadoOrigem, OE.EstadoDestino," & vbCrLf &
            "        OE.GrupoFiscal, OE.CodigoFiscal," & vbCrLf &
            "        isnull(OE.STICMS,-1) as STICMS, isnull(OE.ObsICMS,-1) as ObsICMS," & vbCrLf &
            "        isnull(OE.STIPI,-1) as STIPI, " & vbCrLf &
            "        isnull(OE.STPISCOFINS,0) as STPISCOFINS, isnull(OE.ObsPISCOFINS,-1) as ObsPISCOFINS," & vbCrLf &
            "        isnull(OE.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento," & vbCrLf &
            "        OE.UsuarioInclusao, OE.UsuarioInclusaoData, isnull(OE.CodigoBeneficio,'') AS CodigoBeneficio" & vbCrLf &
            "   FROM OperacaoXEstado OE" & vbCrLf

            If Parametros.Codigo > 0 Then
                sql &= " Inner join OperacaoXEstado OE2" & vbCrLf &
                       "    on OE2.GrupoProduto  = OE.GrupoProduto" & vbCrLf &
                       "   and OE2.Produto       = OE.Produto" & vbCrLf &
                       "   and OE2.Operacao      = OE.Operacao" & vbCrLf &
                       "   and OE2.SubOperacao   = OE.SubOperacao" & vbCrLf &
                       "   and OE2.EstadoOrigem  = OE.EstadoOrigem" & vbCrLf &
                       "   and OE2.EstadoDestino = OE.EstadoDestino" & vbCrLf &
                       "   and OE2.Codigo_id     = " & Parametros.Codigo & vbCrLf
            Else
                sql &= " Where 1 = 1"

                If Parametros.Ativo = 1 Then
                    sql &= "   And OE.Ativo =1" & vbCrLf
                End If

                sql &= "   And (OE.Empresa = '99999999' OR OE.Empresa = '" & Parametros.Empresa & "')" & vbCrLf

                If Parametros.CodigoGrupoProduto.Length > 0 Then
                    sql &= "   And OE.GrupoProduto ='" & Parametros.CodigoGrupoProduto & "'" & vbCrLf
                Else
                    If Parametros.CodigoProduto.Length > 0 Then sql &= "   And OE.Produto ='" & Parametros.CodigoProduto & "'" & vbCrLf
                End If

                If Parametros.CodigoOperacao > 0 Then sql &= "   And OE.Operacao = " & Parametros.CodigoOperacao & vbCrLf
                If Parametros.CodigoSubOperacao > 0 Then sql &= "   And OE.SubOperacao = " & Parametros.CodigoSubOperacao & vbCrLf
                If Parametros.EstadoOrigem.Length > 0 Then sql &= "   And OE.EstadoOrigem ='" & Parametros.EstadoOrigem & "'" & vbCrLf

                If Parametros.RegiaoDestino.Length > 0 Then
                    If Parametros.EstadoDestino.Length > 0 Then
                        sql &= "   And (OE.EstadoDestino ='" & Parametros.RegiaoDestino & "' OR OE.EstadoDestino ='" & Parametros.EstadoDestino & "')" & vbCrLf
                    Else
                        sql &= "   And OE.EstadoDestino ='" & Parametros.RegiaoDestino & "'" & vbCrLf
                    End If
                Else
                    If Parametros.EstadoDestino.Length > 0 Then sql &= "   And OE.EstadoDestino ='" & Parametros.EstadoDestino & "'" & vbCrLf
                End If
            End If

            sql &= " Order by OE.InicioVigencia, OE.Codigo_Id"


            ds = objBanco.ConsultaDataSet(sql, "OpXUf")
        End If

        For Each row As DataRow In ds.Tables("OpXUf").Rows
            Dim Operacao As New OperacaoXEstado()
            Operacao.Codigo = row("Codigo_id")
            Operacao.Ativo = row("Ativo")
            Operacao.InicioVigencia = row("InicioVigencia")
            Operacao.Empresa = row("Empresa")
            Operacao.CodigoGrupoProduto = row("GrupoProduto")
            Operacao.CodigoProduto = row("Produto")
            Operacao.CodigoOperacao = row("Operacao")
            Operacao.CodigoSubOperacao = row("SubOperacao")
            Operacao.EstadoOrigem = row("EstadoOrigem")
            Operacao.EstadoDestino = row("EstadoDestino")
            Operacao.CodigoGrupoFiscal = row("GrupoFiscal")
            Operacao.CodigoFiscal = row("CodigoFiscal")
            Operacao.CodigoSTICMS = row("STICMS")
            Operacao.CodigoObsICMS = row("ObsICMS")
            Operacao.CodigoSTIPI = row("STIPI")
            Operacao.CodigoSTPISCOFINS = row("STPISCOFINS")
            Operacao.CodigoObsPISCOFINS = row("ObsPISCOFINS")
            Operacao.UsuarioInclusao = row("UsuarioInclusao")
            Operacao.DataHoraInclusao = row("UsuarioInclusaoData")
            Operacao.CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
            Operacao.CodigoBeneficio = row("CodigoBeneficio")
            Me.Add(Operacao)
        Next
    End Sub
#End Region
#Region "Fields"

#End Region
#Region "Property"

#End Region
#Region "Methods"

#End Region
End Class

'*********************************************************************************************************
'*******************************  OPERACAO X ESTADO  *****************************************************
'*********************************************************************************************************
<Serializable()> _
Public Class OperacaoXEstado
#Region "Contrutor"
    Public Sub New(Optional ByVal CriaComEncargosObrigatorios As Boolean = False)
        If CriaComEncargosObrigatorios Then
            Me.Encargos.Add(New OperacaoXEstadoXEncargo(Me, "PRODUTO"))
            Me.Encargos.Add(New OperacaoXEstadoXEncargo(Me, "LIQUIDO"))
        End If
    End Sub

    Public Sub New(Parametros As OperacaoXEstado, Optional SomenteAsVigentes As Boolean = False, Optional CodigoFiscal As Integer = 0)
        Selecionar(Parametros, SomenteAsVigentes, CodigoFiscal)
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Codigo As Integer
    Private _Ativo As Boolean = True
    Private _InicioVigencia As Date

    Private _Empresa As String = ""

    Private _CodigoGrupoProduto As String = ""
    Private _CodigoProduto As String = ""
    Private _CodigoOperacao As Integer
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao

    Private _EstadoOrigem As String = ""
    Private _EstadoDestino As String = ""
    Private _RegiaoDestino As String = ""
    Private _CodigoGrupoFiscal As Integer
    Private _CodigoFiscal As Integer
    Private _CodigoSTICMS As Integer = -1
    Private _CodigoObsICMS As Integer
    Private _CodigoSTIPI As Integer = -1
    Private _CodigoSTPISCOFINS As Integer = 0
    'Private _CodigoSTPIBSCBS As Integer = 0
    Private _CodigoObsPISCOFINS As Integer
    Private _UsuarioInclusao As String
    Private _DataHoraInclusao As Date
    '******** Listas *********************
    Private _Encargos As ListOperacaoXEstadoXEncargo

    '***** Usados nas Cargas *************
    Private _Retencao As Boolean = True
    Private _TipoDePessoa As eTipoPessoa = eTipoPessoa.Nenhuma
    Private _EmpresaObriga As Boolean

    '***** Usado na NFE -  CBENEF "Código de Benefício Fiscal na UF aplicado ao item"
    Private _CodigoBeneficio As String = ""
    Private _BeneficioICMS As AjustesDaApuracaoIcms

    'Código da Natureza de Redimento, utilizado no SpedReinf
    Private _CodigoNaturezaDeRendimento As Integer

    Private _Cliente As [Lib].Negocio.Cliente

    Private _CodigoSTIBSCBS As Integer = -1
    Private _CodigoClassificacaoIBSCBS As Integer = -1
    Private _UsarCalculadoraDeImposto As Boolean
    Private _ReducaoIBS_Perc As Decimal
    Private _ReducaoCBS_Perc As Decimal

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

    Public Property Codigo As Integer
        Get
            Return _Codigo
        End Get
        Set(value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Ativo As Boolean
        Get
            Return _Ativo
        End Get
        Set(value As Boolean)
            _Ativo = value
        End Set
    End Property

    Public Property InicioVigencia As Date
        Get
            Return _InicioVigencia
        End Get
        Set(value As Date)
            _InicioVigencia = value
        End Set
    End Property

    Public Property Empresa As String
        Get
            Return _Empresa
        End Get
        Set(value As String)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoGrupoProduto As String
        Get
            Return _CodigoGrupoProduto
        End Get
        Set(value As String)
            _CodigoGrupoProduto = value
        End Set
    End Property

    Public Property CodigoProduto As String
        Get
            Return _CodigoProduto
        End Get
        Set(value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property CodigoOperacao As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
        End Set
    End Property

    Public ReadOnly Property SubOperacao() As SubOperacao
        Get
            If _SubOperacao Is Nothing And _CodigoSubOperacao > 0 Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)

            Return _SubOperacao
        End Get
    End Property

    Public Property EstadoOrigem As String
        Get
            Return _EstadoOrigem
        End Get
        Set(value As String)
            _EstadoOrigem = value
        End Set
    End Property

    Public Property EstadoDestino As String
        Get
            Return _EstadoDestino
        End Get
        Set(value As String)
            _EstadoDestino = value
        End Set
    End Property

    Public Property RegiaoDestino As String
        Get
            Return _RegiaoDestino
        End Get
        Set(value As String)
            _RegiaoDestino = value
        End Set
    End Property

    Public Property CodigoGrupoFiscal As Integer
        Get
            Return _CodigoGrupoFiscal
        End Get
        Set(value As Integer)
            _CodigoGrupoFiscal = value
        End Set
    End Property

    Public Property CodigoFiscal As Integer
        Get
            Return _CodigoFiscal
        End Get
        Set(value As Integer)
            _CodigoFiscal = value
        End Set
    End Property

    Public Property CodigoSTICMS As Integer
        Get
            Return _CodigoSTICMS
        End Get
        Set(value As Integer)
            _CodigoSTICMS = value
        End Set
    End Property

    Public Property CodigoObsICMS As Integer
        Get
            Return _CodigoObsICMS
        End Get
        Set(value As Integer)
            _CodigoObsICMS = value
        End Set
    End Property

    Public Property CodigoSTIPI As Integer
        Get
            Return _CodigoSTIPI
        End Get
        Set(value As Integer)
            _CodigoSTIPI = value
        End Set
    End Property

    Public Property CodigoSTPISCOFINS As Integer
        Get
            Return _CodigoSTPISCOFINS
        End Get
        Set(value As Integer)
            _CodigoSTPISCOFINS = value
        End Set
    End Property

    Public Property CodigoObsPISCOFINS As Integer
        Get
            Return _CodigoObsPISCOFINS
        End Get
        Set(value As Integer)
            _CodigoObsPISCOFINS = value
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

    Public Property DataHoraInclusao As Date
        Get
            Return _DataHoraInclusao
        End Get
        Set(value As Date)
            _DataHoraInclusao = value
        End Set
    End Property


    '**************** LISTAS **************************
    Public Property Encargos As ListOperacaoXEstadoXEncargo
        Get
            If _Encargos Is Nothing Then _Encargos = New ListOperacaoXEstadoXEncargo(Me)
            Return _Encargos
        End Get
        Set(value As ListOperacaoXEstadoXEncargo)
            _Encargos = value
        End Set
    End Property

    'Usados nas cargas
    Public Property TipoDePessoa As eTipoPessoa
        Get
            Return _TipoDePessoa
        End Get
        Set(value As eTipoPessoa)
            _TipoDePessoa = value
        End Set
    End Property

    Public Property Retencao As Boolean
        Get
            Return _Retencao
        End Get
        Set(value As Boolean)
            _Retencao = value
        End Set
    End Property

    Public Property EmpresaObriga As Boolean
        Get
            Return _EmpresaObriga
        End Get
        Set(value As Boolean)
            _EmpresaObriga = value
        End Set
    End Property

    Public Property CodigoBeneficio As String
        Get
            Return _CodigoBeneficio
        End Get
        Set(value As String)
            _CodigoBeneficio = value
        End Set
    End Property

    Public ReadOnly Property BeneficioICMS() As AjustesDaApuracaoIcms
        Get
            If _CodigoBeneficio.Length > 0 Then _BeneficioICMS = New AjustesDaApuracaoIcms(Me.CodigoBeneficio)

            Return _BeneficioICMS
        End Get
    End Property

    Public Property CodigoNaturezaDeRendimento As Integer
        Get
            Return _CodigoNaturezaDeRendimento
        End Get
        Set(value As Integer)
            _CodigoNaturezaDeRendimento = value
        End Set
    End Property

    Public Property Cliente As [Lib].Negocio.Cliente
        Get
            Return _Cliente
        End Get
        Set(value As [Lib].Negocio.Cliente)
            _Cliente = value
        End Set
    End Property

    'Public Property CodigoSTPIBSCBS As Integer
    '    Get
    '        Return _CodigoSTPIBSCBS
    '    End Get
    '    Set(value As Integer)
    '        _CodigoSTPIBSCBS = value
    '    End Set
    'End Property

    Public Property CodigoSTIBSCBS As Integer
        Get
            Return _CodigoSTIBSCBS
        End Get
        Set(value As Integer)
            _CodigoSTIBSCBS = value
        End Set
    End Property

    Public Property CodigoClassificacaoIBSCBS As Integer
        Get
            Return _CodigoClassificacaoIBSCBS
        End Get
        Set(value As Integer)
            _CodigoClassificacaoIBSCBS = value
        End Set
    End Property

    Public Property UsarCalculadoraDeImposto As Boolean
        Get
            Return _UsarCalculadoraDeImposto
        End Get
        Set(value As Boolean)
            _UsarCalculadoraDeImposto = value
        End Set
    End Property

    Public Property ReducaoIBS_Perc As Decimal
        Get
            Return _ReducaoIBS_Perc
        End Get
        Set(value As Decimal)
            _ReducaoIBS_Perc = value
        End Set
    End Property

    Public Property ReducaoCBS_Perc As Decimal
        Get
            Return _ReducaoCBS_Perc
        End Get
        Set(value As Decimal)
            _ReducaoCBS_Perc = value
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sqls2 As New ArrayList
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "Declare" & vbCrLf &
                      " @Cod integer" & vbCrLf &
                      "INSERT INTO OperacaoXEstado (Ativo, InicioVigencia, Empresa, GrupoProduto, Produto, Operacao, SubOperacao, EstadoOrigem, EstadoDestino," & vbCrLf &
                      "        GrupoFiscal, CodigoFiscal, STICMS, ObsICMS, STIPI, STPISCOFINS, ObsPISCOFINS, UsarCalculadoraDeImposto, STIBSCBS, ClassificacaoIBSCBS, ReducaoIBS_Perc, ReducaoCBS_Perc, UsuarioInclusao, UsuarioInclusaoData, CodigoNaturezaDeRendimento, CodigoBeneficio) " & vbCrLf &
                      "VALUES (" & IIf(Me.Ativo, "1", "0") & ",'" & Me.InicioVigencia.ToString("yyyy-MM-dd") & "','" & Me.Empresa & "','" & Me.CodigoGrupoProduto & "','" & Me.CodigoProduto & "'," & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & ",'" & Me.EstadoOrigem & "','" & Me.EstadoDestino & "'," & vbCrLf &
                      Me.CodigoGrupoFiscal & "," & Me.CodigoFiscal & "," & Me.CodigoSTICMS & "," & Me.CodigoObsICMS & "," & Me.CodigoSTIPI & "," & Me.CodigoSTPISCOFINS & "," & Me.CodigoObsPISCOFINS & "," & IIf(Me.UsarCalculadoraDeImposto, "1", "0") & "," & Me.CodigoSTIBSCBS & "," & Me.CodigoClassificacaoIBSCBS & "," & Str(Me.ReducaoIBS_Perc) & "," & Str(Me.ReducaoCBS_Perc) & ",'" & UsuarioServidor.Usuario.Usuario_Id & "','" & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") & "'," & Me.CodigoNaturezaDeRendimento & ",'" & Me.CodigoBeneficio & "');" & vbCrLf &
                      " set @cod =  SCOPE_IDENTITY()" & vbCrLf &
                      " Update OperacaoXEstado set" & vbCrLf &
                      "    Ativo = 0" & vbCrLf &
                      "  Where Empresa        ='" & Me.Empresa & "'" & vbCrLf &
                      "    and GrupoProduto   ='" & Me.CodigoGrupoProduto & "'" & vbCrLf &
                      "    and Produto        ='" & Me.CodigoProduto & "'" & vbCrLf &
                      "    and Operacao       = " & Me.CodigoOperacao & vbCrLf &
                      "    and SubOperacao    = " & Me.CodigoSubOperacao & vbCrLf &
                      "    and EstadoOrigem   ='" & Me.EstadoOrigem & "'" & vbCrLf &
                      "    and EstadoDestino  ='" & Me.EstadoDestino & "'" & vbCrLf &
                      "    and InicioVigencia <'" & Me.InicioVigencia.ToString("yyyy-MM-dd") & "'"
                sqls2.Add(sql)

                SalvarTabelasRelacionadasSql(sqls2)

            Case "U"
                sql = "UPDATE OperacaoXEstado SET" & vbCrLf &
                      "   Ativo =" & IIf(Me.Ativo, "1", "0") & vbCrLf &
                      " Where Codigo_Id = " & Me.Codigo & ";" & vbCrLf
                sqls2.Add(sql)

                SalvarTabelasRelacionadasSql(sqls2)
            Case "D"
                SalvarTabelasRelacionadasSql(sqls2)

                sql = "DELETE OperacaoXEstado " &
                      " Where Codigo_Id = " & Me.Codigo & ";" & vbCrLf
                sqls2.Add(sql)
        End Select

        sql = String.Join(" ", sqls2.ToArray)
        Sqls.Add(sql)
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        'Encargos Da Operacao
        If Not Encargos Is Nothing Then Encargos.SalvarSql(Sqls)
    End Sub

    Private Sub Selecionar(Parametros As OperacaoXEstado, Optional SomenteAsVigentes As Boolean = False, Optional CodigoFiscal As Integer = 0)
        Dim sql As String = ""
        Dim Sql4 As String = ""
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet

        sql = " Select Codigo_Id, Ativo, InicioVigencia," & vbCrLf &
              "        Empresa, GrupoProduto, Produto," & vbCrLf &
              "        Operacao, SubOperacao," & vbCrLf &
              "        EstadoOrigem, EstadoDestino," & vbCrLf &
              "        GrupoFiscal, CodigoFiscal," & vbCrLf &
              "        isnull(STICMS,-1) as STICMS,  isnull(ObsICMS,-1) as ObsICMS," & vbCrLf &
              "        isnull(STIPI,-1) as STIPI," & vbCrLf &
              "        isnull(STPISCOFINS,0) as STPISCOFINS, isnull(ObsPISCOFINS,-1) as ObsPISCOFINS," & vbCrLf &
              IIf(SomenteAsVigentes, "", "isnull(CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento,") & vbCrLf &
              "        UsuarioInclusao, UsuarioInclusaoData, isnull(CodigoBeneficio,'') AS CodigoBeneficio," & vbCrLf &
              "        ISNULL(UsarCalculadoraDeImposto, 0) AS UsarCalculadoraDeImposto, ISNULL(STIBSCBS, 0) AS STIBSCBS, ISNULL(ClassificacaoIBSCBS, 0) AS ClassificacaoIBSCBS, " & vbCrLf &
              "        ISNULL(ReducaoIBS_Perc, 0) AS ReducaoIBS_Perc, ISNULL(ReducaoCBS_Perc, 0) AS ReducaoCBS_Perc " & vbCrLf &
              "   FROM " & IIf(SomenteAsVigentes, "VW_OperacoesVigentes", "OperacaoXEstado") & vbCrLf

        If SomenteAsVigentes Then
            sql &= "   Where (Empresa = '99999999' OR Empresa ='" & Parametros.Empresa & "')" & vbCrLf &
                   "     and GrupoProduto  ='" & Parametros.CodigoGrupoProduto & "'" & vbCrLf &
                   "     and Produto       ='" & Parametros.CodigoProduto & "'" & vbCrLf &
                   "     and Operacao      = " & Parametros.CodigoOperacao & vbCrLf &
                   "     and SubOperacao   = " & Parametros.CodigoSubOperacao & vbCrLf &
                   "     and EstadoOrigem  ='" & Parametros.EstadoOrigem & "'" & vbCrLf

            If CodigoFiscal > 0 Then
                sql &= "     and CodigoFiscal = " & CodigoFiscal & vbCrLf
            Else
                sql &= "     and EstadoDestino ='" & Parametros.EstadoDestino & "'" & vbCrLf
            End If

            ds = objBanco.ConsultaDataSet(sql, "OpXUf")
        Else
            '0 - C/Produto - Estado
            '1 - C/Produto - Regiao do Estado
            '2 - S/Produto - Estado
            '3 - S/Produto - Regiao do Estado
            For x As Integer = 0 To 3 Step 1
                If Parametros.Codigo > 0 Then
                    Sql4 = "  where Codigo_id  =" & Parametros.Codigo & vbCrLf
                Else
                    Sql4 = "  where Codigo_id = (Select max(Codigo_id)" & vbCrLf &
                           "                       from OperacaoXEstado " & vbCrLf &
                           "                      Where Ativo          = 1" & vbCrLf &
                           "                        and Empresa        ='" & Parametros.Empresa & "'" & vbCrLf &
                           "                        and GrupoProduto   ='" & Parametros.CodigoGrupoProduto & "'" & vbCrLf &
                           "                        and Produto        = " & IIf(x = 2 Or x = 3, "''", "'" & Parametros.CodigoProduto & "'") & vbCrLf &
                           "                        and Operacao       = " & Parametros.CodigoOperacao & vbCrLf &
                           "                        and SubOperacao    = " & Parametros.CodigoSubOperacao & vbCrLf &
                           "                        and EstadoOrigem   ='" & Parametros.EstadoOrigem & "'" & vbCrLf

                    If CodigoFiscal > 0 Then
                        Sql4 &= "                        and CodigoFiscal   = " & CodigoFiscal & vbCrLf
                    Else
                        Sql4 &= "                        and EstadoDestino  = " & IIf(x = 1 Or x = 3, "(Select Regiao from Estados where Estado_id = '" & Parametros.EstadoDestino & "')", "'" & Parametros.EstadoDestino & "'") & vbCrLf
                    End If

                    Sql4 &= "                        and InicioVigencia = (SELECT max(InicioVigencia)" & vbCrLf &
                           "						                        FROM OperacaoXEstado" & vbCrLf &
                           "                                               Where Ativo           = 1" & vbCrLf &
                           "                                                 and (Empresa = '99999999' OR Empresa ='" & Parametros.Empresa & "')" & vbCrLf &
                           "                                                 and GrupoProduto    ='" & Parametros.CodigoGrupoProduto & "'" & vbCrLf &
                           "                                                 and Produto         = " & IIf(x = 2 Or x = 3, "''", "'" & Parametros.CodigoProduto & "'") & vbCrLf &
                           "                                                 and Operacao        = " & Parametros.CodigoOperacao & vbCrLf &
                           "                                                 and SubOperacao     = " & Parametros.CodigoSubOperacao & vbCrLf &
                           "                                                 and EstadoOrigem    ='" & Parametros.EstadoOrigem & "'" & vbCrLf

                    If CodigoFiscal > 0 Then
                        Sql4 &= "                                                 and CodigoFiscal    = " & IIf(x = 1 Or x = 3, "(Select Regiao from Estados where Estado_id = '" & Parametros.EstadoDestino & "')", "'" & Parametros.EstadoDestino & "'") & vbCrLf
                    Else
                        Sql4 &= "                                                 and EstadoDestino   = " & IIf(x = 1 Or x = 3, "(Select Regiao from Estados where Estado_id = '" & Parametros.EstadoDestino & "')", "'" & Parametros.EstadoDestino & "'") & vbCrLf
                    End If

                    If IsDate(Parametros.InicioVigencia) AndAlso Parametros.InicioVigencia.Year > 2000 Then
                        Sql4 &= "                                            and InicioVigencia <='" & Parametros.InicioVigencia.ToString("yyyy-MM-dd") & "'" & vbCrLf
                    End If
                    Sql4 &= "                    ))"
                End If

                ds = objBanco.ConsultaDataSet(sql & Sql4, "OpXUf")
                If ds.Tables(0).Rows.Count > 0 Then Exit For
            Next
        End If

        For Each row As DataRow In ds.Tables("OpXUf").Rows
            Me.Codigo = row("Codigo_id")
            Me.Ativo = row("Ativo")
            Me.InicioVigencia = row("InicioVigencia")
            Me.Empresa = row("Empresa")
            Me.CodigoGrupoProduto = row("GrupoProduto")
            Me.CodigoProduto = row("Produto")
            Me.CodigoOperacao = row("Operacao")
            Me.CodigoSubOperacao = row("SubOperacao")
            Me.EstadoOrigem = row("EstadoOrigem")
            Me.EstadoDestino = row("EstadoDestino")
            Me.CodigoGrupoFiscal = row("GrupoFiscal")
            Me.CodigoFiscal = row("CodigoFiscal")
            Me.CodigoSTICMS = row("STICMS")
            Me.CodigoObsICMS = row("ObsICMS")
            Me.CodigoSTIPI = row("STIPI")
            Me.CodigoSTPISCOFINS = row("STPISCOFINS")
            Me.CodigoObsPISCOFINS = row("ObsPISCOFINS")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.DataHoraInclusao = row("UsuarioInclusaoData")

            Me.TipoDePessoa = Parametros.TipoDePessoa
            Me.Retencao = Parametros.Retencao
            Me.EmpresaObriga = Parametros.EmpresaObriga
            If Not SomenteAsVigentes Then
                Me.CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
            End If
            Me.CodigoBeneficio = row("CodigoBeneficio")

            Me.UsarCalculadoraDeImposto = row("UsarCalculadoraDeImposto")
            Me.CodigoSTIBSCBS = row("STIBSCBS")
            Me.CodigoClassificacaoIBSCBS = row("ClassificacaoIBSCBS")
            Me.ReducaoIBS_Perc = row("ReducaoIBS_Perc")
            Me.ReducaoCBS_Perc = row("ReducaoCBS_Perc")

        Next
    End Sub

    Public Sub ConsultarOperacaoNotaDeTerceiro(Parametros As OperacaoXEstado, Optional SomenteAsVigentes As Boolean = False)

        Dim sql As String = ""
        Dim objBanco As New AcessaBanco()
        Dim ds As New DataSet

        sql = "     SELECT Codigo_Id, Ativo, InicioVigencia,
                        Empresa, GrupoProduto, Produto,
                        Operacao, SubOperacao,
                        EstadoOrigem, EstadoDestino,
                        GrupoFiscal, CodigoFiscal,
                        isnull(STICMS,-1) as STICMS,  isnull(ObsICMS,-1) as ObsICMS,
                        isnull(STIPI,-1) as STIPI,
                        isnull(STPISCOFINS,0) as STPISCOFINS, isnull(ObsPISCOFINS,-1) as ObsPISCOFINS,
                        isnull(CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento,
                        UsuarioInclusao, UsuarioInclusaoData, isnull(CodigoBeneficio,'') AS CodigoBeneficio
                    FROM OperacaoXEstado
                    WHERE Operacao          = 70
                        AND SubOperacao     = 99
                        AND Produto         = '903010007'; "

        ds = objBanco.ConsultaDataSet(sql, "OperacaoXEstado")

        For Each row As DataRow In ds.Tables("OperacaoXEstado").Rows

            Me.Codigo = row("Codigo_id")
            Me.Ativo = row("Ativo")
            Me.InicioVigencia = row("InicioVigencia")
            Me.Empresa = row("Empresa")
            Me.CodigoGrupoProduto = row("GrupoProduto")
            Me.CodigoProduto = row("Produto")
            Me.CodigoOperacao = row("Operacao")
            Me.CodigoSubOperacao = row("SubOperacao")
            Me.EstadoOrigem = row("EstadoOrigem")
            Me.EstadoDestino = row("EstadoDestino")
            Me.CodigoGrupoFiscal = row("GrupoFiscal")
            Me.CodigoFiscal = row("CodigoFiscal")
            Me.CodigoSTICMS = row("STICMS")
            Me.CodigoObsICMS = row("ObsICMS")
            Me.CodigoSTIPI = row("STIPI")
            Me.CodigoSTPISCOFINS = row("STPISCOFINS")
            Me.CodigoObsPISCOFINS = row("ObsPISCOFINS")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.DataHoraInclusao = row("UsuarioInclusaoData")

            Me.TipoDePessoa = Parametros.TipoDePessoa
            Me.Retencao = Parametros.Retencao
            Me.EmpresaObriga = Parametros.EmpresaObriga
            If Not SomenteAsVigentes Then
                Me.CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
            End If
            Me.CodigoBeneficio = row("CodigoBeneficio")

        Next

    End Sub

#End Region

End Class
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************
'*****************************************************************************************************************************************

'*********************************************************************************************************
'*************************  LISTA DE OPERACAO X ESTADO X ENCARGO **************************************
'*********************************************************************************************************
Public Class ListOperacaoXEstadoXEncargo
    Inherits List(Of OperacaoXEstadoXEncargo)

#Region "Contrutor"
    Public Sub New(OxE As OperacaoXEstado)
        _OperacaoEstado = OxE
        Dim sql As String
        sql = "SELECT OEE.Encargo_Id, OEE.Sinal, ISNULL(OEE.DebitaConta,'') DebitaConta, ISNULL(OEE.CreditaConta,'') CreditaConta, OEE.AliquotaBase, OEE.Aliquota, isnull(OEE.AliquotaExibicao,OEE.Aliquota) as AliquotaExibicao, isnull(OEE.AliquotaLimite,100) as AliquotaLimite," & vbCrLf &
              "       isnull(E.PodeSofreRetencao,0) as PodeSofreRetencao,  isnull(E.TipoDePessoaRetencao,0) as TipoDePessoaRetencao, isnull(OEE.ObservacaoTributaria, 0) AS ObservacaoTributaria, ISNULL(OE.UsarCalculadoraDeImposto, 0) AS UsarCalculadoraDeImposto, ISNULL(OE.STIBSCBS, 0) AS STIBSCBS, ISNULL(OE.ClassificacaoIBSCBS, 0) AS ClassificacaoIBSCBS, ISNULL(OE.ReducaoIBS_Perc, 0) AS ReducaoIBS_Perc, ISNULL(OE.ReducaoCBS_Perc, 0) AS ReducaoCBS_Perc " & vbCrLf &
              "  FROM OperacaoXEstadoXEncargo OEE" & vbCrLf &
              " Inner join OperacaoXEstado OE" & vbCrLf &
              "    on OEE.Codigo_id = OE.Codigo_id" & vbCrLf &
              " Inner Join Operacoes OP" & vbCrLf &
              "    on OE.Operacao = OP.Operacao_id" & vbCrLf &
              " Inner join Encargos E" & vbCrLf &
              "    on E.Encargo_id = OEE.Encargo_Id" & vbCrLf &
              " Where OEE.Codigo_id =" & OxE.Codigo & vbCrLf &
              "   and OEE.Codigo_id > 0" & vbCrLf

        If OxE.TipoDePessoa = eTipoPessoa.Fisica Then

            'Pedido BAXI -  5629 - 5632 estava com problemas na geração dos encargos PIS e COFINS, 
            If Left(OxE.Empresa, 8) = "40938762" AndAlso OxE.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                sql &= "   and (isnull(E.TipoDePessoa,3) in (1,2,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
            Else
                sql &= "   and (isnull(E.TipoDePessoa,3) in (1,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
            End If

        ElseIf OxE.TipoDePessoa = eTipoPessoa.Juridica Then
            If Not OxE.Cliente Is Nothing AndAlso OxE.Cliente.Codigo > 0 AndAlso OxE.Cliente.CodigoCategoria = 3 Then
                ''SE CATEGORIA DE CLIENTE FOR PRODUTOR EMPRESA = 3 DEVE VIR TODOS OS ENCARGOS - FURLAN - 22/04/2024
                'sql &= "   and (isnull(E.TipoDePessoa,3) in (1,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
                'Tirei fora o "and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" e vamos acompanhar - Furlan 03/02/2025

                If OxE.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                    sql &= "   and (isnull(E.TipoDePessoa,3) in (1,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
                Else
                    sql &= "   and (isnull(E.TipoDePessoa,3) in (1,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
                End If
            ElseIf OxE.EstadoOrigem = "MT" AndAlso OxE.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso Not OxE.SubOperacao.Devolucao Then
                ''SE EMPRESA MT E FOR SAIDA DEVE VIR ENCARGO FETHAB - FURLAN - 02/05/2024
                sql &= "   and (isnull(E.TipoDePessoa,3) in (1,2,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & " and OEE.Encargo_Id not in('FUNRURAL','SENAR'))" & vbCrLf
            Else
                sql &= "   and (isnull(E.TipoDePessoa,3) in (2,3) " & IIf(OxE.EmpresaObriga, "or (E.VerificaEmpresa = 1 and OP.Classe in ('" & eClassesOperacoes.DOACAO.ToString & "','" & eClassesOperacoes.FISCAL.ToString & "','" & eClassesOperacoes.TRANSFERENCIAS.ToString & "','" & eClassesOperacoes.VENDAS.ToString & "'))", "") & ")" & vbCrLf
            End If
        End If

        sql &= " Order by Case OEE.Encargo_id when 'PRODUTO' then 1 when 'LIQUIDO' then 3 else 2 end, OEE.Encargo_id"

        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "OPxEn")

        For Each row In ds.Tables("OPxEn").Rows

            OxE.UsarCalculadoraDeImposto = row("UsarCalculadoraDeImposto")
            OxE.CodigoSTIBSCBS = row("STIBSCBS")
            OxE.CodigoClassificacaoIBSCBS = row("ClassificacaoIBSCBS")
            OxE.ReducaoIBS_Perc = row("ReducaoIBS_Perc")
            OxE.ReducaoCBS_Perc = row("ReducaoCBS_Perc")

            Dim Encargo As New OperacaoXEstadoXEncargo(OxE)
            Encargo.CodigoEncargo = row("Encargo_Id")
            Encargo.Sinal = row("Sinal")
            Encargo.CodigoDebitaConta = row("DebitaConta")
            Encargo.CodigoCreditaConta = row("CreditaConta")
            Encargo.AliquotaBase = row("AliquotaBase")
            Encargo.Aliquota = row("Aliquota")
            Encargo.AliquotaExibicao = row("AliquotaExibicao")
            Encargo.AliquotaLimite = row("AliquotaLimite")
            Encargo.ObservacaoTributaria = row("ObservacaoTributaria")

            If row("PodeSofreRetencao") Then
                Select Case row("TipoDePessoaRetencao")
                    Case 1 'Fisica
                        Me.DescRetencaoPF &= IIf(Me.DescRetencaoPF.Length > 0, ",", "") + Encargo.CodigoEncargo
                    Case 2 'Juridica
                        Me.DescRetencaoPJ &= IIf(Me.DescRetencaoPJ.Length > 0, ",", "") + Encargo.CodigoEncargo
                    Case 3 'Ambas
                        Me.DescRetencaoPF &= IIf(Me.DescRetencaoPF.Length > 0, ",", "") + Encargo.CodigoEncargo
                        Me.DescRetencaoPJ &= IIf(Me.DescRetencaoPJ.Length > 0, ",", "") + Encargo.CodigoEncargo
                End Select

            End If

            If Not row("PodeSofreRetencao") Then
                Me.Add(Encargo)
            ElseIf (row("TipoDePessoaRetencao") = OxE.TipoDePessoa Or row("TipoDePessoaRetencao") = 3) Then
                If OxE.Retencao Then
                    Me.Add(Encargo)
                End If
            Else
                Me.Add(Encargo)
            End If
        Next
    End Sub
#End Region

#Region "Fields"
    Private _OperacaoEstado As OperacaoXEstado
    Private _DescRetencaoPF As String = ""
    Private _DescRetencaoPJ As String = ""
#End Region

#Region "Property"
    Public ReadOnly Property OperacaoEstado As OperacaoXEstado
        Get
            Return _OperacaoEstado
        End Get
    End Property

    Public Property DescRetencaoPF As String
        Get
            Return _DescRetencaoPF
        End Get
        Set(value As String)
            _DescRetencaoPF = value
        End Set
    End Property

    Public Property DescRetencaoPJ As String
        Get
            Return _DescRetencaoPJ
        End Get
        Set(value As String)
            _DescRetencaoPJ = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Sqls.Count = 0 Then Return True

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each Encargo As OperacaoXEstadoXEncargo In Me
            If Me.OperacaoEstado.IUD = "D" Or Me.OperacaoEstado.IUD = "I" Then Encargo.IUD = Me.OperacaoEstado.IUD
            If Encargo.IUD <> "" Then
                Encargo.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'*********************************************************************************************************
'*******************************  OPERACAO X ESTADO X ENCARGO ********************************************
'*********************************************************************************************************
<Serializable()> _
Public Class OperacaoXEstadoXEncargo
#Region "Contrutor"

    Public Sub New()

    End Sub

    Public Sub New(OxE As OperacaoXEstado, Optional ByVal pEncargo As String = Nothing)
        _OperacaoEstado = OxE
        If Not String.IsNullOrEmpty(pEncargo) Then
            Me.CodigoEncargo = pEncargo
            Me.CodigoDebitaConta = Me.Encargo.ContaDebito
            Me.CodigoCreditaConta = Me.Encargo.ContaCredito
            Me.AliquotaBase = Me.Encargo.BaseCalculo
            Me.Aliquota = Me.Encargo.Aliquota
            Me.AliquotaExibicao = Me.Encargo.Aliquota
        End If
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _OperacaoEstado As OperacaoXEstado
    Private _CodigoEncargo As String
    Private _Encargo As Encargo
    Private _Sinal As String
    Private _CodigoDebitaConta As String
    Private _DebitaConta As PlanoDeConta
    Private _CodigoCreditaConta As String
    Private _CreditaConta As PlanoDeConta
    Private _AliquotaBase As Decimal
    Private _Aliquota As Decimal
    Private _AliquotaExibicao As Decimal
    Private _AliquotaLimite As Decimal
    Private _ObservacaoTributaria As Integer


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

    Public ReadOnly Property OperacaoEstado As OperacaoXEstado
        Get
            Return _OperacaoEstado
        End Get
    End Property

    Public Property CodigoEncargo As String
        Get
            Return _CodigoEncargo
        End Get
        Set(value As String)
            _CodigoEncargo = value
        End Set
    End Property

    Public Property Encargo As Encargo
        Get
            If _Encargo Is Nothing And Not String.IsNullOrEmpty(Me.CodigoEncargo) Then _Encargo = New Encargo(Me.CodigoEncargo)
            Return _Encargo
        End Get
        Set(value As Encargo)
            _Encargo = value
        End Set
    End Property

    Public Property Sinal As String
        Get
            Return _Sinal
        End Get
        Set(value As String)
            _Sinal = value
        End Set
    End Property

    Public Property CodigoDebitaConta As String
        Get
            Return _CodigoDebitaConta
        End Get
        Set(value As String)
            _CodigoDebitaConta = value
        End Set
    End Property

    Public Property CodigoCreditaConta As String
        Get
            Return _CodigoCreditaConta
        End Get
        Set(value As String)
            _CodigoCreditaConta = value
        End Set
    End Property

    Public Property DebitaConta() As PlanoDeConta
        Get
            If _DebitaConta Is Nothing And _CodigoDebitaConta.Length > 0 Then _DebitaConta = New PlanoDeConta("", 0, _CodigoDebitaConta)
            Return _DebitaConta
        End Get
        Set(ByVal value As PlanoDeConta)
            _DebitaConta = value
        End Set
    End Property

    Public Property CreditaConta() As PlanoDeConta
        Get
            If _CreditaConta Is Nothing And _CodigoCreditaConta.Length > 0 Then _CreditaConta = New PlanoDeConta("", 0, _CodigoCreditaConta)
            Return _CreditaConta
        End Get
        Set(ByVal value As PlanoDeConta)
            _CreditaConta = value
        End Set
    End Property

    Public Property AliquotaBase As Decimal
        Get
            Return _AliquotaBase
        End Get
        Set(value As Decimal)
            _AliquotaBase = value
        End Set
    End Property

    Public Property Aliquota As Decimal
        Get
            Return _Aliquota
        End Get
        Set(value As Decimal)
            _Aliquota = value
        End Set
    End Property

    Public Property AliquotaExibicao As Decimal
        Get
            Return _AliquotaExibicao
        End Get
        Set(value As Decimal)
            _AliquotaExibicao = value
        End Set
    End Property

    Public Property AliquotaLimite As Decimal
        Get
            Return _AliquotaLimite
        End Get
        Set(value As Decimal)
            _AliquotaLimite = value
        End Set
    End Property

    Public Property ObservacaoTributaria As Integer
        Get
            Return _ObservacaoTributaria
        End Get
        Set(value As Integer)
            _ObservacaoTributaria = value
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
        Dim sql As String = ""

        Select Case _IUD
            Case "I"
                sql = "INSERT INTO OperacaoXEstadoXEncargo (Codigo_Id, Encargo_Id, Sinal, DebitaConta, CreditaConta, AliquotaBase, Aliquota, AliquotaExibicao, AliquotaLimite, ObservacaoTributaria) " & vbCrLf &
                      "VALUES (@cod,'" & Me.CodigoEncargo & "','" & Me.Sinal & "','" & IIf(Me.CodigoEncargo.Equals("LIQUIDO"), "", Me.CodigoDebitaConta) & "'," & vbCrLf &
                      "'" & IIf(Me.CodigoEncargo.Equals("LIQUIDO"), "", Me.CodigoCreditaConta) & "'," & Str(Me.AliquotaBase) & "," & Str(Me.Aliquota) & "," & Str(Me.AliquotaExibicao) & "," & Str(AliquotaLimite) & ", " & ObservacaoTributaria & ");" & vbCrLf
                Sqls.Add(sql)
            Case "U"
                sql = "UPDATE OperacaoXEstadoXEncargo SET" & vbCrLf & _
                      "    Sinal                ='" & Me.Sinal & "'" & vbCrLf & _
                      "   ,DebitaConta          ='" & IIf(Me.CodigoEncargo.Equals("LIQUIDO"), "", Me.CodigoDebitaConta) & "'" & vbCrLf & _
                      "   ,CreditaConta         ='" & IIf(Me.CodigoEncargo.Equals("LIQUIDO"), "", Me.CodigoCreditaConta) & "'" & vbCrLf & _
                      "   ,AliquotaBase         = " & Str(Me.AliquotaBase) & vbCrLf & _
                      "   ,Aliquota             = " & Str(Me.Aliquota) & vbCrLf & _
                      "   ,AliquotaExibicao     = " & Str(Me.AliquotaExibicao) & vbCrLf & _
                      "   ,AliquotaLimite       = " & Str(Me.AliquotaLimite) & vbCrLf & _
                      "   ,ObservacaoTributaria = " & ObservacaoTributaria & vbCrLf & _
                      " Where Codigo_Id  = " & Me.OperacaoEstado.Codigo & vbCrLf & _
                      "   and Encargo_id ='" & Me.CodigoEncargo & "'" & ";" & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = "DELETE OperacaoXEstadoXEncargo" & _
                      " Where Codigo_Id  = " & Me.OperacaoEstado.Codigo & vbCrLf & _
                      "   and Encargo_id ='" & Me.CodigoEncargo & "'" & ";" & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub
#End Region
End Class