Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListEncargosPlanoDeContas
    Inherits List(Of EncargosPlanoDeContas)
    Implements IBaseEntity

#Region "Fields"
    Private _Conta As PlanoDeConta
#End Region

#Region "Property"
    Public Property Conta() As PlanoDeConta
        Get
            Return _Conta
        End Get
        Set(ByVal value As PlanoDeConta)
            _Conta = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pConta As PlanoDeConta, Optional ByVal RecPag As String = "")
        Conta = pConta
        Dim Banco As New AcessaBanco
        Dim sql As String
        Dim ds As DataSet

        sql = "SELECT EPC.Empresa_Id, EPC.EndEmpresa_Id, EPC.Conta_Id, EPC.ContaEncargo_Id" & vbCrLf & _
              "  FROM EncargosPlanoDeContas EPC" & vbCrLf & _
              " Inner Join PlanoDeContas PC" & vbCrLf & _
              "    on PC.Empresa_Id    = EPC.Empresa_Id" & vbCrLf & _
              "   and PC.EndEmpresa_Id = EPC.EndEmpresa_Id" & vbCrLf & _
              "   and PC.Conta_Id      = EPC.ContaEncargo_Id" & vbCrLf & _
              " Where EPC.Empresa_id    ='" & pConta.CodigoEmpresa & "'" & vbCrLf & _
              "   and EPC.EndEmpresa_id = " & pConta.EnderecoEmpresa & vbCrLf & _
              "   and EPC.Conta_Id      ='" & pConta.Conta & "'" & vbCrLf
        If RecPag = "R" Then sql &= " and isnull(PC.Receber,'') <> ''"
        If RecPag = "P" Then sql &= " and isnull(PC.Pagar,'') <> ''"

        sql &= " Order by case" & vbCrLf & _
               "            When EPC.Conta_Id = EPC.ContaEncargo_Id" & vbCrLf & _
               "              Then 1 " & vbCrLf & _
               "              Else 2 " & vbCrLf & _
               "          End, EPC.ContaEncargo_Id"

        ds = Banco.ConsultaDataSet(sql, "EncargosPlanoDeContas")
        For Each row As DataRow In ds.Tables(0).Rows
            Dim Enc As New EncargosPlanoDeContas(Conta)
            Enc.CodigoContaEncargo = row("ContaEncargo_Id")
            Dim objCta As New PlanoDeConta("", 0, Enc.CodigoContaEncargo)
            Enc.TituloEncargo = objCta.Titulo
            Me.Add(Enc)
        Next
    End Sub
#End Region

#Region "Methods"
    Public Function Salvar(ByRef Sqls As ArrayList, Optional ByVal UsaNumerador As Boolean = True) As Boolean
        Dim Banco As New AcessaBanco
        Sqls.Clear()
        Me.SalvarSql(Sqls)

        If Sqls.Count = 0 OrElse Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each enc As EncargosPlanoDeContas In Me
            If Conta.IUD = "D" Or Conta.IUD = "I" Then enc.IUD = Conta.IUD
            If enc.IUD <> "" Then
                enc.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class EncargosPlanoDeContas
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New(ByVal pConta As PlanoDeConta)
        Conta = pConta
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Conta As PlanoDeConta
    Private _CodigoContaEncargo As String
    Private _TituloEncargo As String
    Private _ContaEncargo As PlanoDeConta
    Private _Selecionado As Boolean = False
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

    Public Property Conta() As PlanoDeConta
        Get
            Return _Conta
        End Get
        Set(ByVal value As PlanoDeConta)
            _Conta = value
        End Set
    End Property

    Public Property CodigoContaEncargo() As String
        Get
            Return _CodigoContaEncargo
        End Get
        Set(ByVal value As String)
            _CodigoContaEncargo = value
            _ContaEncargo = Nothing
        End Set
    End Property

    Public Property TituloEncargo() As String
        Get
            Return _TituloEncargo
        End Get
        Set(ByVal value As String)
            _TituloEncargo = value
        End Set
    End Property

    Public Property ContaEncargo() As PlanoDeConta
        Get
            If _ContaEncargo Is Nothing And _CodigoContaEncargo.Length > 0 Then _ContaEncargo = New PlanoDeConta(Conta.CodigoEmpresa, Conta.EnderecoEmpresa, _CodigoContaEncargo)
            Return _ContaEncargo
        End Get
        Set(ByVal value As PlanoDeConta)
            _ContaEncargo = value
        End Set
    End Property

    Public Property Selecionado() As Boolean
        Get
            Return _Selecionado
        End Get
        Set(ByVal value As Boolean)
            _Selecionado = value
        End Set
    End Property
#End Region

#Region "Fields"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList
        SalvarSql(sqls)
        Return Banco.GravaBanco(sqls)
    End Function

    Public Function SalvarSql(ByRef sqls As ArrayList) As ArrayList
        Dim strSQL As String
        Dim ObjBanco As New AcessaBanco

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO EncargosPlanoDeContas(Empresa_Id, EndEmpresa_Id, Conta_Id, ContaEncargo_Id)" & vbCrLf & _
                         " VALUES ('" & Conta.CodigoEmpresa & "'," & Conta.EnderecoEmpresa & ",'" & Conta.Conta & "','" & Me.CodigoContaEncargo & "')"
                sqls.Add(strSQL)
            Case "D"
                strSQL = "Delete EncargosPlanoDeContas " & vbCrLf & _
                         " WHERE Empresa_Id          ='" & Conta.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id       = " & Conta.EnderecoEmpresa & vbCrLf & _
                         "   AND Conta_Id            ='" & Conta.Conta & "'" & vbCrLf & _
                         "   AND ContaEncargo_Id     ='" & Me.CodigoContaEncargo & "'" & vbCrLf
                sqls.Add(strSQL)
        End Select

        Return sqls
    End Function
#End Region

End Class