Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

'**************************************************************************************************************************
'*****************************  Lista de Empresas com o Balanço  **********************************************************
'**************************************************************************************************************************
<Serializable()> _
Public Class ListBalancoAuditadoMes
    Inherits List(Of BalancoAuditadoMes)

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Methods"
    Public Sub Carregar(Ano As Integer)
        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        Dim sql As String = ""

        sql = "Select CE.Empresa_id, " & vbCrLf & _
              "       CE.EndEmpresa_Id," & vbCrLf & _
              "       " & Ano & " as Ano," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,1)  = 1   then bl.auditado  else 0 end),0) Janeiro," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,2)  = 2   then bl.auditado  else 0 end),0) Fevereiro," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,3)  = 3   then bl.auditado  else 0 end),0) Marco," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,4)  = 4   then bl.auditado  else 0 end),0) Abril," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,5)  = 5   then bl.auditado  else 0 end),0) Maio," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,6)  = 6   then bl.auditado  else 0 end),0) Junho," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,7)  = 7   then bl.auditado  else 0 end),0) Julho," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,8)  = 8   then bl.auditado  else 0 end),0) Agosto," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,9)  = 9   then bl.auditado  else 0 end),0) Setembro," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,10) = 10  then bl.auditado  else 0 end),0) Outubro," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,11) = 11  then bl.auditado  else 0 end),0) Novembro," & vbCrLf & _
              "       isnull(max(case when isnull(bl.mes_id,12) = 12  then bl.auditado  else 0 end),0) Dezembro" & vbCrLf & _
              "  from ClientesxEmpresas CE " & vbCrLf & _
              "  left join BalancoAuditadoMes bl" & vbCrLf & _
              "    on bl.Empresa_id    = Ce.Empresa_id" & vbCrLf & _
              "   and bl.EndEmpresa_id = Ce.EndEmpresa_id" & vbCrLf & _
              "   and bl.ano_id        = " & Ano & vbCrLf & _
              "  full join(select 1 Mes union select 2 union select 3 union select 4  union select 5  union select 6 union " & vbCrLf & _
              "            select 7     union select 8 union select 9 union select 10 union select 11 union select 12 " & vbCrLf & _
              "            ) Meses" & vbCrLf & _
              "   on 1 = 1" & vbCrLf & _
              " group by CE.Empresa_id," & vbCrLf & _
              "          CE.EndEmpresa_Id" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Fechamento")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim BL As New BalancoAuditadoMes
            BL.CodigoEmpresa = row("Empresa_id")
            BL.EndEmpresa = row("EndEmpresa_Id")
            BL.Ano = row("Ano")

            BL.JaneiroAuditado = row("Janeiro")
            BL.FevereiroAuditado = row("Fevereiro")
            BL.MarcoAuditado = row("Marco")
            BL.AbrilAuditado = row("Abril")
            BL.MaioAuditado = row("Maio")
            BL.JunhoAuditado = row("Junho")
            BL.JulhoAuditado = row("Julho")
            BL.AgostoAuditado = row("Agosto")
            BL.SetembroAuditado = row("Setembro")
            BL.OutrubroAuditado = row("Outubro")
            BL.NovembroAuditado = row("Novembro")
            BL.DezembroAuditado = row("Dezembro")
            Add(BL)
        Next
    End Sub

    Function Auditado(pEmpresa As String, pEndEmpresa As Integer, pMes As Integer, Optional pConsolidade As Boolean = False) As Boolean
        If pConsolidade Then
            Return Me.Where(Function(s) s.Auditado(pMes) = True).Count > 0
        Else
            Return Me.Where(Function(s) s.CodigoEmpresa = pEmpresa And s.EndEmpresa = pEndEmpresa And s.Auditado(pMes) = True).Count > 0
        End If
    End Function
#End Region

End Class


'***************************************************************************************************************************
'*****************************  Classe Para Bloqueio dos meses Auditados no Balanco  ***************************************
'***************************************************************************************************************************
<Serializable()> _
Public Class BalancoAuditadoMes
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub
#End Region

#Region "Fields"
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _EmpresaBL As NGS.Lib.Negocio.Cliente
    Private _Ano As Integer
    Private _Mes As Integer
    Private _JaneiroAuditado As Boolean = False
    Private _FevereiroAuditado As Boolean = False
    Private _MarcoAuditado As Boolean = False
    Private _AbrilAuditado As Boolean = False
    Private _MaioAuditado As Boolean = False
    Private _JunhoAuditado As Boolean = False
    Private _JulhoAuditado As Boolean = False
    Private _AgostoAuditado As Boolean = False
    Private _SetembroAuditado As Boolean = False
    Private _OutrubroAuditado As Boolean = False
    Private _NovembroAuditado As Boolean = False
    Private _DezembroAuditado As Boolean = False
#End Region

#Region "Property"
    Public Property CodigoEmpresa As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EndEmpresa As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public ReadOnly Property EmpresaBL As NGS.Lib.Negocio.Cliente
        Get
            If _EmpresaBL Is Nothing AndAlso Not String.IsNullOrWhiteSpace(Me.CodigoEmpresa) Then _EmpresaBL = New NGS.Lib.Negocio.Cliente(Me.CodigoEmpresa, Me.EndEmpresa)
            Return _EmpresaBL
        End Get
    End Property

    Public ReadOnly Property DescricaoEmpresa
        Get
            Return Funcoes.FormatarCpfCnpj(Me.CodigoEmpresa) + " : " + EmpresaBL.Nome + " / " + EmpresaBL.Cidade + "-" + EmpresaBL.CodigoEstado
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

    Public Property JaneiroAuditado As Boolean
        Get
            Return _JaneiroAuditado
        End Get
        Set(value As Boolean)
            _JaneiroAuditado = value
        End Set
    End Property

    Public Property FevereiroAuditado As Boolean
        Get
            Return _FevereiroAuditado
        End Get
        Set(value As Boolean)
            _FevereiroAuditado = value
        End Set
    End Property

    Public Property MarcoAuditado As Boolean
        Get
            Return _MarcoAuditado
        End Get
        Set(value As Boolean)
            _MarcoAuditado = value
        End Set
    End Property

    Public Property AbrilAuditado As Boolean
        Get
            Return _AbrilAuditado
        End Get
        Set(value As Boolean)
            _AbrilAuditado = value
        End Set
    End Property

    Public Property MaioAuditado As Boolean
        Get
            Return _MaioAuditado
        End Get
        Set(value As Boolean)
            _MaioAuditado = value
        End Set
    End Property

    Public Property JunhoAuditado As Boolean
        Get
            Return _JunhoAuditado
        End Get
        Set(value As Boolean)
            _JunhoAuditado = value
        End Set
    End Property

    Public Property JulhoAuditado As Boolean
        Get
            Return _JulhoAuditado
        End Get
        Set(value As Boolean)
            _JulhoAuditado = value
        End Set
    End Property

    Public Property AgostoAuditado As Boolean
        Get
            Return _AgostoAuditado
        End Get
        Set(value As Boolean)
            _AgostoAuditado = value
        End Set
    End Property

    Public Property SetembroAuditado As Boolean
        Get
            Return _SetembroAuditado
        End Get
        Set(value As Boolean)
            _SetembroAuditado = value
        End Set
    End Property

    Public Property OutrubroAuditado As Boolean
        Get
            Return _OutrubroAuditado
        End Get
        Set(value As Boolean)
            _OutrubroAuditado = value
        End Set
    End Property

    Public Property NovembroAuditado As Boolean
        Get
            Return _NovembroAuditado
        End Get
        Set(value As Boolean)
            _NovembroAuditado = value
        End Set
    End Property

    Public Property DezembroAuditado As Boolean
        Get
            Return _DezembroAuditado
        End Get
        Set(value As Boolean)
            _DezembroAuditado = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar(pMes As Integer) As Boolean
        Dim sql As String
        sql = "MERGE BalancoAuditadoMes AS Target" & vbCrLf & _
              "USING (SELECT '" & Me.CodigoEmpresa & "' as Empresa_id, " & Me.EndEmpresa & " as EndEmpresa_id, " & Me.Ano & " as Ano_id, " & pMes & " as Mes_Id, " & IIf(Not Auditado(pMes), 1, 0) & " as Auditado) AS Source" & vbCrLf & _
              "   ON Target.Empresa_id    = Source.Empresa_id" & vbCrLf & _
              "  AND Target.EndEmpresa_id = Source.EndEmpresa_id" & vbCrLf & _
              "  AND Target.Ano_id        = Source.Ano_id" & vbCrLf & _
              "  AND Target.Mes_id        = Source.Mes_id" & vbCrLf & _
              "WHEN MATCHED THEN" & vbCrLf & _
              "    UPDATE SET Target.Auditado = Source.Auditado" & vbCrLf & _
              "WHEN NOT MATCHED BY TARGET THEN" & vbCrLf & _
              "    INSERT (Empresa_id, EndEmpresa_Id, Ano_id, Mes_Id, Auditado)" & vbCrLf & _
              "    VALUES (Source.Empresa_id, Source.EndEmpresa_Id, Source.Ano_id,Source.Mes_Id, Source.Auditado);" & vbCrLf

        If Not Auditado(pMes) Then
            sql &= "Update AcessosXMovimento set" & vbCrLf & _
                   "   Situacao = 'BLOQUEADO'" & vbCrLf & _
                   " where Empresa_id          ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                   "   and EndEmpresa_id       = " & Me.EndEmpresa & vbCrLf & _
                   "   and year(movimento_id)  = " & Me.Ano & vbCrLf & _
                   "   and Month(movimento_id) = " & pMes & vbCrLf & _
                   "   and Situacao = 'LIBERADO';" & vbCrLf

            sql &= "Delete AcessosXProcessos" & vbCrLf & _
                   " Where Empresa_id          ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                   "   and EndEmpresa_id       = " & Me.EndEmpresa & vbCrLf & _
                   "   and year(movimento_id)  = " & Me.Ano & vbCrLf & _
                   "   and month(movimento_id) = " & pMes & vbCrLf

            sql &= "Delete AcessosXUsuarios" & vbCrLf & _
                   " Where Empresa_id          ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                   "   and EndEmpresa_id       = " & Me.EndEmpresa & vbCrLf & _
                   "   and year(movimento_id)  = " & Me.Ano & vbCrLf & _
                   "   and month(movimento_id) = " & pMes & vbCrLf
        End If


        Dim Banco As New AcessaBanco
        If Banco.GravaBanco(sql) Then
            Select Case pMes
                Case 1 : JaneiroAuditado = Not JaneiroAuditado
                Case 2 : FevereiroAuditado = Not FevereiroAuditado
                Case 3 : MarcoAuditado = Not MarcoAuditado
                Case 4 : AbrilAuditado = Not AbrilAuditado
                Case 5 : MaioAuditado = Not MaioAuditado
                Case 6 : JunhoAuditado = Not JunhoAuditado
                Case 7 : JulhoAuditado = Not JulhoAuditado
                Case 8 : AgostoAuditado = Not AgostoAuditado
                Case 9 : SetembroAuditado = Not SetembroAuditado
                Case 10 : OutrubroAuditado = Not OutrubroAuditado
                Case 11 : NovembroAuditado = Not NovembroAuditado
                Case 12 : DezembroAuditado = Not DezembroAuditado
            End Select
            Return True
        Else
            Return False
        End If
    End Function

    Public Function Auditado(pmes As Integer)
        Select Case pmes
            Case 1 : Return JaneiroAuditado
            Case 2 : Return FevereiroAuditado
            Case 3 : Return MarcoAuditado
            Case 4 : Return AbrilAuditado
            Case 5 : Return MaioAuditado
            Case 6 : Return JunhoAuditado
            Case 7 : Return JulhoAuditado
            Case 8 : Return AgostoAuditado
            Case 9 : Return SetembroAuditado
            Case 10 : Return OutrubroAuditado
            Case 11 : Return NovembroAuditado
            Case 12 : Return DezembroAuditado
            Case Else : Return ""
        End Select
    End Function
#End Region

End Class








