Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCertidaoNegativa
    Inherits List(Of CertidaoNegativa)

    Public Sub New()

    End Sub

    Public Sub UltimasInformadas(Optional ByVal pCertidaoNegativa As CertidaoNegativa = Nothing, Optional ByVal pDataEmissaoInicial As String = "", Optional ByVal pDataEmissaoFinal As String = "", Optional ByVal pDataValidadeInicial As String = "", Optional ByVal pDataValidadeFinal As String = "", Optional ByVal IE As String = "", Optional ByVal pEstadoDoCliente As String = "", Optional ByVal pConsolidada As Boolean = True, Optional ByVal pEmpresa As String = "")
        Dim sql As String
        sql = "Select CN.Cliente_id, " & vbCrLf & _
              "       (Select Top 1 C.Nome from Clientes C where C.Cliente_id = CN.Cliente_Id) as Nome," & vbCrLf & _
              "       CN.Numero_id, " & vbCrLf & _
              "       CN.DataEmissao, " & vbCrLf & _
              "       CN.DataValidade, " & vbCrLf & _
              "       CN.CodigoAutenticidade," & vbCrLf & _
              "       CN.ModeloCertidao" & vbCrLf & _
              "  from CertidaoNegativa CN" & vbCrLf

        If pConsolidada Then
            sql &= " Inner Join (" & vbCrLf & _
                   "			    select case" & vbCrLf & _
                   "                       when Len(Cliente_id) = 11 then Cliente_id" & vbCrLf & _
                   "                       When Len(Cliente_id) = 14 then left(Cliente_id,8) " & vbCrLf & _
                   "                     end as Cliente," & vbCrLf & _
                   "                     Max(DataValidade) as Data" & vbCrLf & _
                   "			      from CertidaoNegativa" & vbCrLf & _
                   "			     group by case" & vbCrLf & _
                   "				 		    when Len(Cliente_id) = 11 then Cliente_id" & vbCrLf & _
                   "						    When Len(Cliente_id) = 14 then left(Cliente_id,8) " & vbCrLf & _
                   "						  end" & vbCrLf & _
                   "			   ) sb" & vbCrLf & _
                   "    on case" & vbCrLf & _
                   "         When len(CN.Cliente_id) = 11 then CN.Cliente_id" & vbCrLf & _
                   "         when Len(CN.Cliente_id) = 14 then left(CN.Cliente_id,8) " & vbCrLf & _
                   "       end = sb.Cliente" & vbCrLf & _
                   "   and CN.DataValidade = sb.Data" & vbCrLf & _
                   " Where 1 = 1"
        Else
            sql &= " Inner Join (" & vbCrLf & _
                   "			    select Cliente_id as Cliente," & vbCrLf & _
                   "                       Max(DataValidade) as Data" & vbCrLf & _
                   "			      from CertidaoNegativa" & vbCrLf & _
                   "			     group by Cliente_id" & vbCrLf & _
                   "			 ) sb" & vbCrLf & _
                   "    on CN.Cliente_id   = sb.Cliente" & vbCrLf & _
                   "   and CN.DataValidade = sb.Data" & vbCrLf & _
                   " Where 1 = 1"
        End If

        If IE.Length > 0 Then
            sql &= " and CN.cliente_Id in (select Cliente_id from Clientes where inscricao like '%" & IE & "%')"
        End If

        If pEmpresa = "S" Then
            sql &= " and CN.cliente_Id in (Select Empresa_id from clientesXEmpresas)"
        End If

        If Not pCertidaoNegativa Is Nothing Then
            If pCertidaoNegativa.CodigoCliente <> "" Then sql &= " And " & IIf(pCertidaoNegativa.CodigoCliente.Length = 11, " CN.Cliente_id ='" & pCertidaoNegativa.CodigoCliente & "'", "Left(CN.Cliente_id,8) = '" & pCertidaoNegativa.CodigoCliente.Substring(0, 8) & "'")
            If pCertidaoNegativa.CodigoAutenticidade <> "" Then sql &= " And CN.CodigoAutenticidade = '" & pCertidaoNegativa.CodigoAutenticidade & "'"
            If pCertidaoNegativa.Numero <> "" Then sql &= " And CN.Numero_Id ='" & pCertidaoNegativa.Numero & "'"
        End If

        If Not pDataEmissaoInicial = "" Then
            sql &= " And Cn.DataEmissao between '" & CDate(pDataEmissaoInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(pDataEmissaoFinal).ToString("yyyy-MM-dd") & "'"
        End If

        If Not pDataValidadeInicial = "" Then
            sql &= " And Cn.DataValidade between '" & CDate(pDataValidadeInicial).ToString("yyyy-MM-dd") & "' and '" & CDate(pDataValidadeFinal).ToString("yyyy-MM-dd") & "'"
        End If

        If pEstadoDoCliente.Length > 0 Then
            sql &= " and exists(select 1 from clientes cl where cl.cliente_id = CN.Cliente_id and cl.estado in " & pEstadoDoCliente & ")"
        End If

        sql &= " Order By Nome "

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "CN")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CN As New CertidaoNegativa
            CN.CodigoCliente = row("Cliente_id")
            CN.NomeCliente = row("Nome")
            CN.Numero = row("Numero_id")
            CN.DataEmissao = row("DataEmissao")
            CN.DataValidade = row("DataValidade")
            CN.CodigoAutenticidade = row("CodigoAutenticidade")
            CN.CodigoModeloCertidao = row("ModeloCertidao")
            Me.Add(CN)
        Next

    End Sub

    Public Sub Historico(ByVal pCodigoCliente As String, ByVal Consolidado As Boolean)
        Dim sql As String
        sql = "Select CN.Cliente_id, " & vbCrLf & _
              "       (Select Top 1 C.Nome from Clientes C where C.Cliente_id = CN.Cliente_Id) as Nome," & vbCrLf & _
              "       CN.Numero_id, " & vbCrLf & _
              "       CN.DataEmissao, " & vbCrLf & _
              "       CN.DataValidade, " & vbCrLf & _
              "       CN.CodigoAutenticidade," & vbCrLf & _
              "       CN.ModeloCertidao" & vbCrLf & _
              "  from CertidaoNegativa CN" & vbCrLf

        If Consolidado Then
            sql &= " Where " & IIf(pCodigoCliente.Length = 11, "CN.Cliente_Id", "Left(CN.Cliente_Id,8)") & " = '" & IIf(pCodigoCliente.Length = 11, pCodigoCliente, pCodigoCliente.Substring(0, 8)) & "'" & vbCrLf
        Else
            sql &= " Where CN.Cliente_Id = '" & pCodigoCliente & "'" & vbCrLf
        End If

        sql &= " Order by CN.DataEmissao desc"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "CN")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CN As New CertidaoNegativa
            CN.CodigoCliente = row("Cliente_id")
            CN.NomeCliente = row("Nome")
            CN.Numero = row("Numero_id")
            CN.DataEmissao = row("DataEmissao")
            CN.DataValidade = row("DataValidade")
            CN.CodigoAutenticidade = row("CodigoAutenticidade")
            CN.CodigoModeloCertidao = row("ModeloCertidao")
            Me.Add(CN)
        Next
    End Sub

End Class

<Serializable()> _
Public Class CertidaoNegativa
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCliente As String, ByVal pConsolidado As Boolean)
        Dim sql As String
        sql = "Select CN.Cliente_id, " & vbCrLf & _
              "       (Select Top 1 C.Nome from Clientes C where C.Cliente_id = CN.Cliente_Id) as Nome," & vbCrLf & _
              "       CN.Numero_id, " & vbCrLf & _
              "       CN.DataEmissao, " & vbCrLf & _
              "       CN.DataValidade, " & vbCrLf & _
              "       CN.CodigoAutenticidade," & vbCrLf & _
              "       CN.ModeloCertidao" & vbCrLf & _
              "  from CertidaoNegativa CN" & vbCrLf
        If pConsolidado Then
            sql &= " Inner Join (" & vbCrLf & _
                   "			    select case" & vbCrLf & _
                   "                       when Len(Cliente_id) = 11 then Cliente_id" & vbCrLf & _
                   "                       When Len(Cliente_id) = 14 then left(Cliente_id,8) " & vbCrLf & _
                   "                     end as Cliente," & vbCrLf & _
                   "                     Max(DataValidade) as Data" & vbCrLf & _
                   "			      from CertidaoNegativa" & vbCrLf & _
                   "			     group by case" & vbCrLf & _
                   "				 		    when Len(Cliente_id) = 11 then Cliente_id" & vbCrLf & _
                   "						    When Len(Cliente_id) = 14 then left(Cliente_id,8) " & vbCrLf & _
                   "						  end" & vbCrLf & _
                   "			   ) sb" & vbCrLf & _
                   "    on case" & vbCrLf & _
                   "         When len(CN.Cliente_id) = 11 then CN.Cliente_id" & vbCrLf & _
                   "         when Len(CN.Cliente_id) = 14 then left(CN.Cliente_id,8) " & vbCrLf & _
                   "       end = sb.Cliente" & vbCrLf & _
                   "   and CN.DataValidade = sb.Data" & vbCrLf & _
                   " Where " & IIf(pCliente.Length = 14, " left(CN.Cliente_id,8) = '" & pCliente.Substring(0, 8), " CN.Cliente_Id = '" & pCliente) & "'"
        Else
            sql &= " Inner Join (" & vbCrLf & _
                   "			    select Cliente_id as Cliente," & vbCrLf & _
                   "                       Max(DataValidade) as Data" & vbCrLf & _
                   "			      from CertidaoNegativa" & vbCrLf & _
                   "			     group by Cliente_id" & vbCrLf & _
                   "			   ) sb" & vbCrLf & _
                   "    on CN.Cliente_id   = sb.Cliente" & vbCrLf & _
                   "   and CN.DataValidade = sb.Data" & vbCrLf & _
                   " Where CN.Cliente_id = '" & pCliente & "'"
        End If


        Dim Banco As New AcessaBanco
        Dim ds As DataSet
        ds = Banco.ConsultaDataSet(sql, "CN")

        For Each row As DataRow In ds.Tables(0).Rows
            _CodigoCliente = row("Cliente_id")
            _NomeCliente = row("Nome")
            _Numero = row("Numero_id")
            _DataEmissao = row("DataEmissao")
            _DataValidade = row("DataValidade")
            _CodigoAutenticidade = row("CodigoAutenticidade")
            _CodigoModeloCertidao = row("ModeloCertidao")
        Next
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _CodigoCliente As String = ""
    Private _NomeCliente As String = ""
    Private _Numero As String = ""
    Private _DataEmissao As Date
    Private _DataValidade As Date
    Private _CodigoAutenticidade As String
    Private _CodigoModeloCertidao As Integer
    Private _ModeloCertidao As ModeloCertidao
    Private _DescricaoModelo As String = ""
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

    Public Property NomeCliente() As String
        Get
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property

    Public Property Numero() As String
        Get
            Return _Numero
        End Get
        Set(ByVal value As String)
            _Numero = value
        End Set
    End Property

    Public Property DataEmissao() As Date
        Get
            Return _DataEmissao
        End Get
        Set(ByVal value As Date)
            _DataEmissao = value
        End Set
    End Property

    Public Property DataValidade() As Date
        Get
            Return _DataValidade
        End Get
        Set(ByVal value As Date)
            _DataValidade = value
        End Set
    End Property

    Public Property CodigoAutenticidade() As String
        Get
            Return _CodigoAutenticidade
        End Get
        Set(ByVal value As String)
            _CodigoAutenticidade = value
        End Set
    End Property

    Public Property CodigoModeloCertidao() As Integer
        Get
            Return _CodigoModeloCertidao
        End Get
        Set(ByVal value As Integer)
            _CodigoModeloCertidao = value
            _ModeloCertidao = Nothing
            _DescricaoModelo = ""
        End Set
    End Property

    Public Property ModeloCertidao() As ModeloCertidao
        Get
            If _ModeloCertidao Is Nothing Then
                _ModeloCertidao = New ModeloCertidao(_CodigoModeloCertidao)
                _DescricaoModelo = _ModeloCertidao.Descricao
            End If

            Return _ModeloCertidao
        End Get
        Set(ByVal value As ModeloCertidao)
            _ModeloCertidao = value
        End Set
    End Property

    Public Property DescricaoModelo()
        Get
            If _DescricaoModelo = "" And Not ModeloCertidao Is Nothing Then _DescricaoModelo = ModeloCertidao.Descricao
            Return _DescricaoModelo
        End Get
        Set(ByVal value)
            _DescricaoModelo = value
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
                Sql = "Insert Into CertidaoNegativa(Cliente_id, Numero_id, DataEmissao, DataValidade, CodigoAutenticidade, ModeloCertidao)" & vbCrLf & _
                      " values('" & _CodigoCliente & "','" & _Numero & "','" & _DataEmissao.ToString("yyyy-MM-dd") & "','" & _DataValidade.ToString("yyyy-MM-dd") & "','" & _CodigoAutenticidade & "'," & Str(_CodigoModeloCertidao) & ")"
                Sqls.Add(Sql)
            Case "U"
                Sql = "Update CertidaoNegativa Set  " & _
                      "  DataEmissao         ='" & _DataEmissao.ToString("yyyy-MM-dd") & "'" & _
                      ", DataValidade        ='" & _DataValidade.ToString("yyyy-MM-dd") & "'" & _
                      ", CodigoAutenticidade ='" & _CodigoAutenticidade & "'" & _
                      ", ModeloCertidao      = " & Str(_CodigoModeloCertidao) & _
                      " Where Cliente_id ='" & _CodigoCliente & "'" & _
                      "   and Numero_id  ='" & _Numero & "'"
                Sqls.Add(Sql)
            Case "D"
                Sql = "Delete CertidaoNegativa " & vbCrLf & _
                      " Where Cliente_id ='" & _CodigoCliente & "'" & _
                      "   and Numero_id  ='" & _Numero & "'"
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class