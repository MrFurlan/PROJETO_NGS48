Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable>
Public Class ListNaviosXInvoice
    Inherits List(Of NavioXInvoice)
    Implements IBaseEntity

#Region "Constructor"
    Public Sub New()

    End Sub

    Public Sub New(Optional ByVal Carregar As Boolean = False, Optional ByVal pWhere As String = "")
        If Carregar Then
            Dim objBanco As New AcessaBanco
            Dim sql As String = " SELECT NXI.Codigo_Id, NXI.Navio_Id, NV.Descricao, NXI.Observacao, NXI.DataDeChegada, NXI.Ativo, NXI.UsuarioInclusao, NXI.UsuarioInclusaoData, " & vbCrLf &
                                " ISNULL(NXI.UsuarioAlteracao, 0) as UsuarioAlteracao, ISNULL(NXI.UsuarioAlteracaoData, 0) as UsuarioAlteracaoData, NXI.Pais " & vbCrLf &
                                " FROM NaviosXInvoice AS NXI " & vbCrLf &
                                " INNER JOIN Navios AS NV " & vbCrLf &
                                " ON NXI.Navio_Id = NV.Codigo_Id " & vbCrLf &
                                " INNER JOIN NavioXInvoiceXProduto AS NVP " & vbCrLf &
                                " ON NVP.Codigo_Id = NXI.Codigo_Id " & vbCrLf

            If Not pWhere.Trim.Length = 0 Then
                sql &= pWhere
            End If
            sql &= "  ORDER BY NXI.Codigo_Id "

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "Navios")

            If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then

                For Each Row As DataRow In ds.Tables(0).Rows

                    Dim objNavioXInvoice As New NavioXInvoice

                    objNavioXInvoice.Codigo = Row("Codigo_Id")
                    objNavioXInvoice.Navio_Id = Row("Navio_Id")
                    objNavioXInvoice.Descricao = Row("Descricao")
                    objNavioXInvoice.DataDeChegada = Row("DataDeChegada")
                    objNavioXInvoice.Observacao = Row("Observacao")
                    objNavioXInvoice.Ativo = Row("Ativo")
                    objNavioXInvoice.UsuarioInclusao = Row("UsuarioInclusao")
                    objNavioXInvoice.UsuarioInclusaoData = Row("UsuarioInclusaoData")
                    objNavioXInvoice.UsuarioAlteracao = Row("UsuarioAlteracao")
                    objNavioXInvoice.UsuarioAlteracaoData = Row("UsuarioAlteracaoData")
                    objNavioXInvoice.CodigoPais = Row("Pais")

                    Me.Add(objNavioXInvoice)

                Next

            End If

        End If
    End Sub
#End Region

End Class

<Serializable>
Public Class NavioXInvoice
    Implements IBaseEntity

#Region "Constructor"

    Public Sub New()
    End Sub

    Public Sub New(Codigo As Integer)

        Dim Banco As New AcessaBanco
        Dim sql As String = " SELECT NXI.Codigo_Id, NXI.Navio_Id, NV.Descricao, " & vbCrLf &
                                "NXI.DataDeChegada, NXI.Observacao, NXI.Ativo, " & vbCrLf &
                                "NXI.UsuarioInclusao, NXI.UsuarioInclusaoData," & vbCrLf &
                                "ISNULL(NXI.UsuarioAlteracao, '') UsuarioAlteracao, ISNULL(NXI.UsuarioAlteracaoData, '') UsuarioAlteracaoData, NXI.Pais " & vbCrLf &
                            " FROM NaviosXInvoice AS NXI " & vbCrLf &
                            " INNER JOIN Navios AS NV " & vbCrLf &
                            " ON NXI.Navio_Id = NV.Codigo_Id " & vbCrLf &
                            " WHERE NXI.Codigo_Id = " & Codigo

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Navios")
        If ds.Tables.Count = 0 OrElse ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim Row As DataRow = ds.Tables(0).Rows(0)

        _Codigo = Row("Codigo_Id")
        _Navio_Id = Row("Navio_Id")
        _Descricao = Row("Descricao")
        _DataDeChegada = Row("DataDeChegada")
        _Observacao = Row("Observacao")
        _Ativo = Row("Ativo")
        _UsuarioInclusao = Row("UsuarioInclusao")
        _UsuarioInclusaoData = Row("UsuarioInclusaoData")
        _UsuarioAlteracao = Row("UsuarioAlteracao")
        _UsuarioAlteracaoData = Row("UsuarioAlteracaoData")
        _CodigoPais = Row("Pais")

        _NavioXInvoiceXProduto = New NavioXInvoiceXProduto(Me)

    End Sub
#End Region

#Region "Fields"

    Private _IUD As String
    Private _Codigo As String
    Private _Navio_Id As String
    Private _Descricao As String
    Private _DataDeChegada As DateTime
    Private _Observacao As String
    Private _Ativo As Boolean
    Private _UsuarioInclusao As String
    Private _UsuarioInclusaoData As String
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As String
    Private _CodigoPais As Integer
    Private _Pais As Pais

    Private _NavioXInvoiceXProduto As NavioXInvoiceXProduto

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

    Public Property Codigo As String
        Get
            Return _Codigo
        End Get
        Set(value As String)
            _Codigo = value
        End Set
    End Property

    Public Property Navio_Id As String
        Get
            Return _Navio_Id
        End Get
        Set(value As String)
            _Navio_Id = value
        End Set
    End Property

    Public Property Descricao As String
        Get
            Return _Descricao
        End Get
        Set(value As String)
            _Descricao = value
        End Set
    End Property

    Public Property DataDeChegada As DateTime
        Get
            Return _DataDeChegada
        End Get
        Set(value As DateTime)
            _DataDeChegada = value
        End Set
    End Property

    Public Property Observacao As String
        Get
            Return _Observacao
        End Get
        Set(value As String)
            _Observacao = value
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

    Public Property NavioXInvoiceXProduto() As NavioXInvoiceXProduto
        Get
            If _NavioXInvoiceXProduto Is Nothing Then _NavioXInvoiceXProduto = New NavioXInvoiceXProduto(Me)

            Return _NavioXInvoiceXProduto
        End Get
        Set(ByVal value As NavioXInvoiceXProduto)
            _NavioXInvoiceXProduto = value
        End Set
    End Property

    Public Property CodigoPais As Integer
        Get
            Return _CodigoPais
        End Get
        Set(value As Integer)
            _CodigoPais = value
        End Set
    End Property


    Public Property Pais As Pais
        Get
            If _Pais Is Nothing AndAlso _CodigoPais > 0 Then _Pais = New Pais(_CodigoPais)

            Return _Pais
        End Get
        Set(value As Pais)
            _Pais = value
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

        Select Case Me.IUD
            Case "I"
                sql = "Declare" & vbCrLf &
                      " @Cod integer" & vbCrLf &
                      "INSERT INTO NaviosXInvoice (Navio_Id, DataDeChegada, Observacao, Ativo, UsuarioInclusao, UsuarioInclusaoData, Pais)" & vbCrLf &
                      "Values (" & _Navio_Id & "," & vbCrLf &
                      "'" & _DataDeChegada.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "'" & _Observacao & "'," & vbCrLf &
                      "'" & IIf(_Ativo, 1, 0) & "'," & vbCrLf &
                      "'" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "'" & Now().ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "" & CodigoPais & ");" & vbCrLf &
                      " set @cod =  SCOPE_IDENTITY();" & vbCrLf &
                      "INSERT INTO NavioXInvoiceXProduto(Codigo_Id, Produto_Id, Quantidade)" & vbCrLf &
                                                "Values(@cod,'" & Me.NavioXInvoiceXProduto.CodigoProduto & "'," & Me.NavioXInvoiceXProduto.Quantidade.ToString().Replace(",", ".") & ");"
                Sqls.Add(sql)

            Case "U"
                sql = " UPDATE NaviosXInvoice " & vbCrLf &
                      "    SET Observacao = '" & _Observacao & "'," & vbCrLf &
                      "        DataDeChegada = '" & _DataDeChegada.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                      "        Pais          = " & _CodigoPais & "," & vbCrLf &
                      "     UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "     UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      " WHERE Codigo_Id = " & _Codigo & ""
                Sqls.Add(sql)

                salvarRelacionados(Sqls)

            Case "D"

                'salvarRelacionados(Sqls)

                sql = " UPDATE NaviosXInvoice " & vbCrLf &
                      "    SET Ativo = 0, UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                      "        UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                      "  WHERE Codigo_Id = " & _Codigo & ""
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub salvarRelacionados(ByRef sqls As ArrayList)

        If Me.NavioXInvoiceXProduto.CodigoProduto.Length > 0 Then
            Me.NavioXInvoiceXProduto.IUD = "U"
            Me.NavioXInvoiceXProduto.SalvarSql(sqls)
        End If

    End Sub

#End Region

End Class