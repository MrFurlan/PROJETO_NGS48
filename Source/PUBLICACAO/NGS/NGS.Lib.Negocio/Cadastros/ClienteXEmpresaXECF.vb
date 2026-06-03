Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ClienteXEmpresaXECF
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(cxe As ClienteXEmpresa)
        _Empresa = cxe

        Dim sql As String
        sql = "SELECT IND_SIT_INI_PER, SIT_ESPECIAL, TIP_ECF, OPT_REFIS, OPT_PAES, FORMA_TRIB, FORMA_APUR, COD_QUALIF_PJ, FORMA_TRIB_PER, MES_BAL_RED, TIP_ESC_PRE," & vbCrLf & _
              "       OPT_EXT_RTT, DIF_FCONT, IND_ALIQ_CSLL, ISNULL(IND_ALIQ_CSLL_I,0) IND_ALIQ_CSLL_I , IND_QTE_SCP, IND_ADM_FUN_CLU, IND_PART_CONS, IND_OP_EXT, IND_OP_VINC, IND_PJ_ENQUAD, IND_PART_EXT, IND_ATIV_RURAL, IND_LUC_EXP," & vbCrLf & _
              "       IND_RED_ISEN, IND_FIN, IND_DOA_ELEIT, IND_PART_COLIG, IND_VEND_EXP, IND_REC_EXT, IND_ATIV_EXT, IND_COM_EXP, IND_PGTO_EXT, IND_E_COM_TI, IND_ROY_REC, IND_ROY_PAG," & vbCrLf & _
              "       IND_REND_SERV, IND_PGTO_REM, IND_INOV_TEC, IND_CAP_INF, IND_PJ_HAB, IND_POLO_AM, IND_ZON_EXP, IND_AREA_COM, TIP_ENT, FORMA_APUR_I, APUR_CSLL" & vbCrLf & _
              "  FROM ClienteXEmpresaXECF" & vbCrLf & _
              " Where Left(empresa_id,8) ='" & Left(cxe.Empresa_id, 8) & "'"
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "ECF")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.IND_SIT_INI_PER = row("IND_SIT_INI_PER")
        Me.SIT_ESPECIAL = row("SIT_ESPECIAL")
        Me.TIP_ECF = row("TIP_ECF")
        Me.OPT_REFIS = row("OPT_REFIS")
        Me.OPT_PAES = row("OPT_PAES")
        Me.FORMA_TRIB = row("FORMA_TRIB")
        Me.FORMA_APUR = row("FORMA_APUR")
        Me.COD_QUALIF_PJ = row("COD_QUALIF_PJ")
        Me.FORMA_TRIB_PER = row("FORMA_TRIB_PER")
        Me.MES_BAL_RED = row("MES_BAL_RED")
        Me.TIP_ESC_PRE = row("TIP_ESC_PRE")

        Me.TIP_ENT = row("TIP_ENT")
        Me.FORMA_APUR_I = row("FORMA_APUR_I")
        Me.APUR_CSLL = row("APUR_CSLL")

        Me.OPT_EXT_RTT = row("OPT_EXT_RTT")
        Me.DIF_FCONT = row("DIF_FCONT")
        Me.IND_ALIQ_CSLL = row("IND_ALIQ_CSLL")
        Me.IND_ALIQ_CSLL_I = row("IND_ALIQ_CSLL_I")
        Me.IND_QTE_SCP = row("IND_QTE_SCP")
        Me.IND_ADM_FUN_CLU = row("IND_ADM_FUN_CLU")
        Me.IND_PART_CONS = row("IND_PART_CONS")
        Me.IND_OP_EXT = row("IND_OP_EXT")
        Me.IND_OP_VINC = row("IND_OP_VINC")
        Me.IND_PJ_ENQUAD = row("IND_PJ_ENQUAD")
        Me.IND_PART_EXT = row("IND_PART_EXT")
        Me.IND_ATIV_RURAL = row("IND_ATIV_RURAL")
        Me.IND_LUC_EXP = row("IND_LUC_EXP")
        Me.IND_RED_ISEN = row("IND_RED_ISEN")
        Me.IND_FIN = row("IND_FIN")
        Me.IND_DOA_ELEIT = row("IND_DOA_ELEIT")
        Me.IND_PART_COLIG = row("IND_PART_COLIG")
        Me.IND_VEND_EXP = row("IND_VEND_EXP")
        Me.IND_REC_EXT = row("IND_REC_EXT")
        Me.IND_ATIV_EXT = row("IND_ATIV_EXT")
        Me.IND_COM_EXP = row("IND_COM_EXP")
        Me.IND_PGTO_EXT = row("IND_PGTO_EXT")
        Me.IND_E_COM_TI = row("IND_E_COM_TI")
        Me.IND_ROY_REC = row("IND_ROY_REC")
        Me.IND_ROY_PAG = row("IND_ROY_PAG")
        Me.IND_REND_SERV = row("IND_REND_SERV")
        Me.IND_PGTO_REM = row("IND_PGTO_REM")
        Me.IND_INOV_TEC = row("IND_INOV_TEC")
        Me.IND_CAP_INF = row("IND_CAP_INF")
        Me.IND_PJ_HAB = row("IND_PJ_HAB")
        Me.IND_POLO_AM = row("IND_POLO_AM")
        Me.IND_ZON_EXP = row("IND_ZON_EXP")
        Me.IND_AREA_COM = row("IND_AREA_COM")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Empresa As ClienteXEmpresa
    Private _IND_SIT_INI_PER As Integer
    Private _SIT_ESPECIAL As Integer
    Private _TIP_ECF As Integer
    Private _OPT_REFIS As Boolean
    Private _OPT_PAES As Boolean
    Private _FORMA_TRIB As Integer
    Private _FORMA_APUR As String = " "
    Private _COD_QUALIF_PJ As Integer
    Private _FORMA_TRIB_PER As String = " "
    Private _MES_BAL_RED As String = " "
    Private _TIP_ESC_PRE As String = " "

    Private _TIP_ENT As Integer
    Private _FORMA_APUR_I As String = " "
    Private _APUR_CSLL As String = " "

    Private _OPT_EXT_RTT As Boolean
    Private _DIF_FCONT As Boolean
    Private _IND_ALIQ_CSLL As Boolean
    Private _IND_ALIQ_CSLL_I As Integer
    Private _IND_QTE_SCP As Integer
    Private _IND_ADM_FUN_CLU As Boolean
    Private _IND_PART_CONS As Boolean
    Private _IND_OP_EXT As Boolean
    Private _IND_OP_VINC As Boolean
    Private _IND_PJ_ENQUAD As Boolean
    Private _IND_PART_EXT As Boolean
    Private _IND_ATIV_RURAL As Boolean
    Private _IND_LUC_EXP As Boolean
    Private _IND_RED_ISEN As Boolean
    Private _IND_FIN As Boolean
    Private _IND_DOA_ELEIT As Boolean
    Private _IND_PART_COLIG As Boolean
    Private _IND_VEND_EXP As Boolean
    Private _IND_REC_EXT As Boolean
    Private _IND_ATIV_EXT As Boolean
    Private _IND_COM_EXP As Boolean
    Private _IND_PGTO_EXT As Boolean
    Private _IND_E_COM_TI As Boolean
    Private _IND_ROY_REC As Boolean
    Private _IND_ROY_PAG As Boolean
    Private _IND_REND_SERV As Boolean
    Private _IND_PGTO_REM As Boolean
    Private _IND_INOV_TEC As Boolean
    Private _IND_CAP_INF As Boolean
    Private _IND_PJ_HAB As Boolean
    Private _IND_POLO_AM As Boolean
    Private _IND_ZON_EXP As Boolean
    Private _IND_AREA_COM As Boolean
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
    Public ReadOnly Property Empresa As ClienteXEmpresa
        Get
            Return _Empresa
        End Get
    End Property
    Public Property IND_SIT_INI_PER As Integer
        Get
            Return _IND_SIT_INI_PER
        End Get
        Set(value As Integer)
            _IND_SIT_INI_PER = value
        End Set
    End Property
    Public Property SIT_ESPECIAL As Integer
        Get
            Return _SIT_ESPECIAL
        End Get
        Set(value As Integer)
            _SIT_ESPECIAL = value
        End Set
    End Property
    Public Property TIP_ECF As Integer
        Get
            Return _TIP_ECF
        End Get
        Set(value As Integer)
            _TIP_ECF = value
        End Set
    End Property
    Public Property OPT_REFIS As Boolean
        Get
            Return _OPT_REFIS
        End Get
        Set(value As Boolean)
            _OPT_REFIS = value
        End Set
    End Property
    Public Property OPT_PAES As Boolean
        Get
            Return _OPT_PAES
        End Get
        Set(value As Boolean)
            _OPT_PAES = value
        End Set
    End Property
    Public Property FORMA_TRIB As Integer
        Get
            Return _FORMA_TRIB
        End Get
        Set(value As Integer)
            _FORMA_TRIB = value
        End Set
    End Property
    Public Property FORMA_APUR As String
        Get
            Return _FORMA_APUR
        End Get
        Set(value As String)
            _FORMA_APUR = value
        End Set
    End Property
    Public Property COD_QUALIF_PJ As Integer
        Get
            Return _COD_QUALIF_PJ
        End Get
        Set(value As Integer)
            _COD_QUALIF_PJ = value
        End Set
    End Property
    Public Property FORMA_TRIB_PER As String
        Get
            Return _FORMA_TRIB_PER
        End Get
        Set(value As String)
            _FORMA_TRIB_PER = value
        End Set
    End Property
    Public Property MES_BAL_RED As String
        Get
            Return _MES_BAL_RED
        End Get
        Set(value As String)
            _MES_BAL_RED = value
        End Set
    End Property
    Public Property TIP_ESC_PRE As String
        Get
            Return _TIP_ESC_PRE
        End Get
        Set(value As String)
            _TIP_ESC_PRE = value
        End Set
    End Property

    Public Property TIP_ENT As Integer
        Get
            Return _TIP_ENT
        End Get
        Set(value As Integer)
            _TIP_ENT = value
        End Set
    End Property

    Public Property FORMA_APUR_I As String
        Get
            Return _FORMA_APUR_I
        End Get
        Set(value As String)
            _FORMA_APUR_I = value
        End Set
    End Property

    Public Property APUR_CSLL As String
        Get
            Return _APUR_CSLL
        End Get
        Set(value As String)
            _APUR_CSLL = value
        End Set
    End Property

    Public Property OPT_EXT_RTT As Boolean
        Get
            Return _OPT_EXT_RTT
        End Get
        Set(value As Boolean)
            _OPT_EXT_RTT = value
        End Set
    End Property
    Public Property DIF_FCONT As Boolean
        Get
            Return _DIF_FCONT
        End Get
        Set(value As Boolean)
            _DIF_FCONT = value
        End Set
    End Property
    Public Property IND_ALIQ_CSLL As Boolean
        Get
            Return _IND_ALIQ_CSLL
        End Get
        Set(value As Boolean)
            _IND_ALIQ_CSLL = value
        End Set
    End Property

    Public Property IND_ALIQ_CSLL_I As Integer
        Get
            Return _IND_ALIQ_CSLL_I
        End Get
        Set(ByVal value As Integer)
            _IND_ALIQ_CSLL_I = value
        End Set
    End Property


    Public Property IND_QTE_SCP As Integer
        Get
            Return _IND_QTE_SCP
        End Get
        Set(value As Integer)
            _IND_QTE_SCP = value
        End Set
    End Property
    Public Property IND_ADM_FUN_CLU As Boolean
        Get
            Return _IND_ADM_FUN_CLU
        End Get
        Set(value As Boolean)
            _IND_ADM_FUN_CLU = value
        End Set
    End Property
    Public Property IND_PART_CONS As Boolean
        Get
            Return _IND_PART_CONS
        End Get
        Set(value As Boolean)
            _IND_PART_CONS = value
        End Set
    End Property
    Public Property IND_OP_EXT As Boolean
        Get
            Return _IND_OP_EXT
        End Get
        Set(value As Boolean)
            _IND_OP_EXT = value
        End Set
    End Property
    Public Property IND_OP_VINC As Boolean
        Get
            Return _IND_OP_VINC
        End Get
        Set(value As Boolean)
            _IND_OP_VINC = value
        End Set
    End Property
    Public Property IND_PJ_ENQUAD As Boolean
        Get
            Return _IND_PJ_ENQUAD
        End Get
        Set(value As Boolean)
            _IND_PJ_ENQUAD = value
        End Set
    End Property
    Public Property IND_PART_EXT As Boolean
        Get
            Return _IND_PART_EXT
        End Get
        Set(value As Boolean)
            _IND_PART_EXT = value
        End Set
    End Property
    Public Property IND_ATIV_RURAL As Boolean
        Get
            Return _IND_ATIV_RURAL
        End Get
        Set(value As Boolean)
            _IND_ATIV_RURAL = value
        End Set
    End Property
    Public Property IND_LUC_EXP As Boolean
        Get
            Return _IND_LUC_EXP
        End Get
        Set(value As Boolean)
            _IND_LUC_EXP = value
        End Set
    End Property
    Public Property IND_RED_ISEN As Boolean
        Get
            Return _IND_RED_ISEN
        End Get
        Set(value As Boolean)
            _IND_RED_ISEN = value
        End Set
    End Property
    Public Property IND_FIN As Boolean
        Get
            Return _IND_FIN
        End Get
        Set(value As Boolean)
            _IND_FIN = value
        End Set
    End Property
    Public Property IND_DOA_ELEIT As Boolean
        Get
            Return _IND_DOA_ELEIT
        End Get
        Set(value As Boolean)
            _IND_DOA_ELEIT = value
        End Set
    End Property
    Public Property IND_PART_COLIG As Boolean
        Get
            Return _IND_PART_COLIG
        End Get
        Set(value As Boolean)
            _IND_PART_COLIG = value
        End Set
    End Property
    Public Property IND_VEND_EXP As Boolean
        Get
            Return _IND_VEND_EXP
        End Get
        Set(value As Boolean)
            _IND_VEND_EXP = value
        End Set
    End Property
    Public Property IND_REC_EXT As Boolean
        Get
            Return _IND_REC_EXT
        End Get
        Set(value As Boolean)
            _IND_REC_EXT = value
        End Set
    End Property
    Public Property IND_ATIV_EXT As Boolean
        Get
            Return _IND_ATIV_EXT
        End Get
        Set(value As Boolean)
            _IND_ATIV_EXT = value
        End Set
    End Property
    Public Property IND_COM_EXP As Boolean
        Get
            Return _IND_COM_EXP
        End Get
        Set(value As Boolean)
            _IND_COM_EXP = value
        End Set
    End Property
    Public Property IND_PGTO_EXT As Boolean
        Get
            Return _IND_PGTO_EXT
        End Get
        Set(value As Boolean)
            _IND_PGTO_EXT = value
        End Set
    End Property
    Public Property IND_E_COM_TI As Boolean
        Get
            Return _IND_E_COM_TI
        End Get
        Set(value As Boolean)
            _IND_E_COM_TI = value
        End Set
    End Property
    Public Property IND_ROY_REC As Boolean
        Get
            Return _IND_ROY_REC
        End Get
        Set(value As Boolean)
            _IND_ROY_REC = value
        End Set
    End Property
    Public Property IND_ROY_PAG As Boolean
        Get
            Return _IND_ROY_PAG
        End Get
        Set(value As Boolean)
            _IND_ROY_PAG = value
        End Set
    End Property
    Public Property IND_REND_SERV As Boolean
        Get
            Return _IND_REND_SERV
        End Get
        Set(value As Boolean)
            _IND_REND_SERV = value
        End Set
    End Property
    Public Property IND_PGTO_REM As Boolean
        Get
            Return _IND_PGTO_REM
        End Get
        Set(value As Boolean)
            _IND_PGTO_REM = value
        End Set
    End Property
    Public Property IND_INOV_TEC As Boolean
        Get
            Return _IND_INOV_TEC
        End Get
        Set(value As Boolean)
            _IND_INOV_TEC = value
        End Set
    End Property
    Public Property IND_CAP_INF As Boolean
        Get
            Return _IND_CAP_INF
        End Get
        Set(value As Boolean)
            _IND_CAP_INF = value
        End Set
    End Property
    Public Property IND_PJ_HAB As Boolean
        Get
            Return _IND_PJ_HAB
        End Get
        Set(value As Boolean)
            _IND_PJ_HAB = value
        End Set
    End Property
    Public Property IND_POLO_AM As Boolean
        Get
            Return _IND_POLO_AM
        End Get
        Set(value As Boolean)
            _IND_POLO_AM = value
        End Set
    End Property
    Public Property IND_ZON_EXP As Boolean
        Get
            Return _IND_ZON_EXP
        End Get
        Set(value As Boolean)
            _IND_ZON_EXP = value
        End Set
    End Property
    Public Property IND_AREA_COM As Boolean
        Get
            Return _IND_AREA_COM
        End Get
        Set(value As Boolean)
            _IND_AREA_COM = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        If Empresa.Matriz = "N" Then Exit Sub
        Dim sql As String
        Select Case Me.IUD
            Case "I", "U"
                sql = "MERGE ClienteXEmpresaXECF AS Target" & vbCrLf & _
                      "USING (SELECT '" & Empresa.Empresa_id & "' as Empresa_id) AS Source" & vbCrLf & _
                      "   ON left(Target.Empresa_id,8)  = Left(Source.Empresa_id,8)" & vbCrLf & _
                      "WHEN MATCHED THEN" & vbCrLf & _
                      " UPDATE SET" & vbCrLf & _
                      "   IND_SIT_INI_PER = " & Me.IND_SIT_INI_PER & vbCrLf & _
                      "  ,SIT_ESPECIAL    = " & Me.SIT_ESPECIAL & vbCrLf & _
                      "  ,TIP_ECF         = " & Me.TIP_ECF & vbCrLf & _
                      "  ,OPT_REFIS       ='" & Me.OPT_REFIS.ToString & "'" & vbCrLf & _
                      "  ,OPT_PAES        ='" & Me.OPT_PAES.ToString & "'" & vbCrLf & _
                      "  ,FORMA_TRIB      = " & Me.FORMA_TRIB & vbCrLf & _
                      "  ,FORMA_APUR      ='" & Me.FORMA_APUR & "'" & vbCrLf & _
                      "  ,COD_QUALIF_PJ   = " & Me.COD_QUALIF_PJ & vbCrLf & _
                      "  ,FORMA_TRIB_PER  ='" & Me.FORMA_TRIB_PER & "'" & vbCrLf & _
                      "  ,MES_BAL_RED     ='" & Me.MES_BAL_RED & "'" & vbCrLf & _
                      "  ,TIP_ESC_PRE     ='" & Me.TIP_ESC_PRE & "'" & vbCrLf & _
                      "  ,TIP_ENT         = " & Me.TIP_ENT & vbCrLf & _
                      "  ,FORMA_APUR_I    ='" & Me.FORMA_APUR_I & "'" & vbCrLf & _
                      "  ,APUR_CSLL       ='" & Me.APUR_CSLL & "'" & vbCrLf & _
                      "  ,OPT_EXT_RTT     ='" & Me.OPT_EXT_RTT & "'" & vbCrLf & _
                      "  ,DIF_FCONT       ='" & Me.DIF_FCONT & "'" & vbCrLf & _
                      "  ,IND_ALIQ_CSLL   ='" & Me.IND_ALIQ_CSLL & "'" & vbCrLf & _
                      "  ,IND_ALIQ_CSLL_I   ='" & Me.IND_ALIQ_CSLL_I & "'" & vbCrLf & _
                      "  ,IND_QTE_SCP     ='" & Me.IND_QTE_SCP & "'" & vbCrLf & _
                      "  ,IND_ADM_FUN_CLU ='" & Me.IND_ADM_FUN_CLU & "'" & vbCrLf & _
                      "  ,IND_PART_CONS   ='" & Me.IND_PART_CONS & "'" & vbCrLf & _
                      "  ,IND_OP_EXT      ='" & Me.IND_OP_EXT & "'" & vbCrLf & _
                      "  ,IND_OP_VINC     ='" & Me.IND_OP_VINC & "'" & vbCrLf & _
                      "  ,IND_PJ_ENQUAD   ='" & Me.IND_PJ_ENQUAD & "'" & vbCrLf & _
                      "  ,IND_PART_EXT    ='" & Me.IND_PART_EXT & "'" & vbCrLf & _
                      "  ,IND_ATIV_RURAL  ='" & Me.IND_ATIV_RURAL & "'" & vbCrLf & _
                      "  ,IND_LUC_EXP     ='" & Me.IND_LUC_EXP & "'" & vbCrLf & _
                      "  ,IND_RED_ISEN    ='" & Me.IND_RED_ISEN & "'" & vbCrLf & _
                      "  ,IND_FIN         ='" & Me.IND_FIN & "'" & vbCrLf & _
                      "  ,IND_DOA_ELEIT   ='" & Me.IND_DOA_ELEIT & "'" & vbCrLf & _
                      "  ,IND_PART_COLIG  ='" & Me.IND_PART_COLIG & "'" & vbCrLf & _
                      "  ,IND_VEND_EXP    ='" & Me.IND_VEND_EXP & "'" & vbCrLf & _
                      "  ,IND_REC_EXT     ='" & Me.IND_REC_EXT & "'" & vbCrLf & _
                      "  ,IND_ATIV_EXT    ='" & Me.IND_ATIV_EXT & "'" & vbCrLf & _
                      "  ,IND_COM_EXP     ='" & Me.IND_COM_EXP & "'" & vbCrLf & _
                      "  ,IND_PGTO_EXT    ='" & Me.IND_PGTO_EXT & "'" & vbCrLf & _
                      "  ,IND_E_COM_TI    ='" & Me.IND_E_COM_TI & "'" & vbCrLf & _
                      "  ,IND_ROY_REC     ='" & Me.IND_ROY_REC & "'" & vbCrLf & _
                      "  ,IND_ROY_PAG     ='" & Me.IND_ROY_PAG & "'" & vbCrLf & _
                      "  ,IND_REND_SERV   ='" & Me.IND_REND_SERV & "'" & vbCrLf & _
                      "  ,IND_PGTO_REM    ='" & Me.IND_PGTO_REM & "'" & vbCrLf & _
                      "  ,IND_INOV_TEC    ='" & Me.IND_INOV_TEC & "'" & vbCrLf & _
                      "  ,IND_CAP_INF     ='" & Me.IND_CAP_INF & "'" & vbCrLf & _
                      "  ,IND_PJ_HAB      ='" & Me.IND_PJ_HAB & "'" & vbCrLf & _
                      "  ,IND_POLO_AM     ='" & Me.IND_POLO_AM & "'" & vbCrLf & _
                      "  ,IND_ZON_EXP     ='" & Me.IND_ZON_EXP & "'" & vbCrLf & _
                      "  ,IND_AREA_COM    ='" & Me.IND_AREA_COM & "'" & vbCrLf & _
                      "WHEN NOT MATCHED BY TARGET THEN" & vbCrLf & _
                      " INSERT(Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                      "        IND_SIT_INI_PER, SIT_ESPECIAL, TIP_ECF, OPT_REFIS, OPT_PAES, FORMA_TRIB, FORMA_APUR, COD_QUALIF_PJ, FORMA_TRIB_PER, MES_BAL_RED, TIP_ESC_PRE, TIP_ENT, FORMA_APUR_I, APUR_CSLL," & vbCrLf & _
                      "        OPT_EXT_RTT, DIF_FCONT, IND_ALIQ_CSLL, IND_ALIQ_CSLL_I, IND_QTE_SCP, IND_ADM_FUN_CLU, IND_PART_CONS, IND_OP_EXT, IND_OP_VINC, IND_PJ_ENQUAD, IND_PART_EXT, IND_ATIV_RURAL, IND_LUC_EXP," & vbCrLf & _
                      "        IND_RED_ISEN, IND_FIN, IND_DOA_ELEIT, IND_PART_COLIG, IND_VEND_EXP, IND_REC_EXT, IND_ATIV_EXT, IND_COM_EXP, IND_PGTO_EXT, IND_E_COM_TI, IND_ROY_REC, IND_ROY_PAG," & vbCrLf & _
                      "        IND_REND_SERV, IND_PGTO_REM, IND_INOV_TEC, IND_CAP_INF, IND_PJ_HAB, IND_POLO_AM, IND_ZON_EXP, IND_AREA_COM) " & vbCrLf & _
                      " VALUES " & vbCrLf & _
                      "('" & Empresa.Empresa_id & "'," & Empresa.EndEmpresa_Id & vbCrLf & _
                      ", " & Me.IND_SIT_INI_PER & vbCrLf & _
                      ", " & Me.SIT_ESPECIAL & vbCrLf & _
                      ", " & Me.TIP_ECF & vbCrLf & _
                      ",'" & Me.OPT_REFIS.ToString & "'" & vbCrLf & _
                      ",'" & Me.OPT_PAES.ToString & "'" & vbCrLf & _
                      ", " & Me.FORMA_TRIB & vbCrLf & _
                      ",'" & Me.FORMA_APUR & "'" & vbCrLf & _
                      ", " & Me.COD_QUALIF_PJ & vbCrLf & _
                      ",'" & Me.FORMA_TRIB_PER & "'" & vbCrLf & _
                      ",'" & Me.MES_BAL_RED & "'" & vbCrLf & _
                      ",'" & Me.TIP_ESC_PRE & "'" & vbCrLf & _
                      ", " & Me.TIP_ENT & vbCrLf & _
                      ",'" & Me.FORMA_APUR_I & "'" & vbCrLf & _
                      ",'" & Me.APUR_CSLL & "'" & vbCrLf & _
                      ",'" & Me.OPT_EXT_RTT.ToString & "'" & vbCrLf & _
                      ",'" & Me.DIF_FCONT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ALIQ_CSLL.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ALIQ_CSLL_I.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_QTE_SCP.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ADM_FUN_CLU.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PART_CONS.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_OP_EXT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_OP_VINC.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PJ_ENQUAD.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PART_EXT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ATIV_RURAL.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_LUC_EXP.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_RED_ISEN.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_FIN.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_DOA_ELEIT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PART_COLIG.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_VEND_EXP.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_REC_EXT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ATIV_EXT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_COM_EXP.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PGTO_EXT.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_E_COM_TI.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ROY_REC.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ROY_PAG.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_REND_SERV.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PGTO_REM.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_INOV_TEC.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_CAP_INF.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_PJ_HAB.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_POLO_AM.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_ZON_EXP.ToString & "'" & vbCrLf & _
                      ",'" & Me.IND_AREA_COM.ToString & "');"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete ClienteXEmpresaXECF" & vbCrLf & _
                      " where Empresa_Id    = " & Empresa.Empresa_id & vbCrLf & _
                      "   and EndEmpresa_id = " & Empresa.EndEmpresa_Id & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

#End Region

End Class
