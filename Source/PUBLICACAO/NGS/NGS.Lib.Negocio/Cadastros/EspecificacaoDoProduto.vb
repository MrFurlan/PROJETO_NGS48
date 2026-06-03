Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

<Serializable()> _
Public Class ListEspecificacaoDoProduto
    Inherits List(Of EspecificacaoDoProduto)

#Region "Construtor"
    Public Sub New()
        Dim sql As String = "Select Codigo_Id, Descricao, isnull(Ativo,0) as Ativo from EspecificacaoDoProduto"
        Dim banco As New AcessaBanco
        Dim ds As DataSet

        ds = banco.ConsultaDataSet(sql, "TD")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim tb As New EspecificacaoDoProduto
            tb.Codigo = row("Codigo_Id")
            tb.Descricao = row("Descricao")
            tb.Ativo = row("Ativo")
            Me.Add(tb)
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class EspecificacaoDoProduto
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigo As Integer)
        Dim sql As String = "Select Codigo_Id, Descricao, isnull(Ativo,0) as Ativo from EspecificacaoDoProduto where Codigo_id = " & pCodigo
        Dim db As New AcessaBanco
        Dim ds As DataSet

        ds = db.ConsultaDataSet(sql, "EspecificacaoDoProduto")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim row As DataRow = ds.Tables(0).Rows(0)
            _Codigo = row("Codigo_Id")
            _Descricao = row("Descricao")
            _Ativo = row("Ativo")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _Ativo As Boolean = False

    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _UsuarioCancelamento As String = ""
    Private _DataCancelamento As DateTime

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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property Ativo() As Boolean
        Get
            Return _Ativo
        End Get
        Set(ByVal value As Boolean)
            _Ativo = value
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
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
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
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime
        Get
            Return _DataCancelamento
        End Get
        Set(ByVal value As DateTime)
            _DataCancelamento = value
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
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into EspecificacaoDoProduto(Descricao, Ativo, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                      " Values ('" & _Descricao & "','" & _Ativo.ToString() & "','" & HttpContext.Current.Session("ssNomeUsuario") & "', getdate())"
                Sqls.Add(sql)
            Case "U"
                sql = " Update EspecificacaoDoProduto set " & vbCrLf & _
                      "        Descricao ='" & _Descricao & "'" & vbCrLf & _
                      "       ,Ativo ='" & _Ativo.ToString() & "'" & vbCrLf & _
                      "	      ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "	      ,UsuarioAlteracaoData = getdate() " & vbCrLf & _
                      "	Where Codigo_Id =" & _Codigo
                Sqls.Add(sql)
            Case "D"
                sql = " Update EspecificacaoDoProduto  " & vbCrLf & _
                      "    Set Ativo                    ='" & _Ativo.ToString() & "'" & vbCrLf & _
                      "	      ,UsuarioCancelamento     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "	     ,UsuarioCancelamentoData = getdate() " & vbCrLf & _
                      "	Where Codigo_Id =" & _Codigo
                Sqls.Add(sql)

                sql = " Update ProdutoXEspecificacao  " & vbCrLf & _
                      "    Set Ativo                    ='" & _Ativo.ToString() & "'" & vbCrLf & _
                      "	      ,UsuarioCancelamento     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                      "       ,UsuarioCancelamentoData = getdate() " & vbCrLf & _
                      "	Where CodigoEspecificacao_Id =" & _Codigo
                Sqls.Add(sql)
        End Select

    End Sub
#End Region

End Class