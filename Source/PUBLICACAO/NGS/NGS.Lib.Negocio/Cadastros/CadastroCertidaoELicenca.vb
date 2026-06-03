Imports NGS.Lib.Negocio

<Serializable()> _
Public Class ListCadastroCertidaoELicenca
    Inherits List(Of CadastroCertidaoELicenca)
    Implements IBaseEntity

#Region "Builders"
    Public Sub New()

    End Sub
    Public Sub New(ByVal Empresa_Id As String, ByVal Tipo As Integer)
        Dim objBanco As New AcessaBanco
        Dim sql As String
        sql = " Select cl.Empresa_Id, tc.Tipo_Id as Tipo, " & vbCrLf & _
              "        cl.DataEmissao, cl.DataVencimento, cl.Email, cl.AvisoDias, cl.Observacao " & vbCrLf & _
              "   from CertidaoLicenca cl                                                       " & vbCrLf & _
              "  Inner Join TipoDeCertidao tc                                                   " & vbCrLf & _
              "     on tc.Tipo_Id = cl.Tipo_Id                                                  " & vbCrLf & _
              "  where 1=1" & vbCrLf

        If Not String.IsNullOrWhiteSpace(Empresa_Id) Then
            sql &= "And cl.Empresa_Id = '" & Empresa_Id & "'" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(Tipo) Then
            sql &= "And tc.Tipo_Id  = " & Tipo & vbCrLf
        End If

        sql &= "  ORDER BY cl.Empresa_Id "

        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "CertidaoLicenca")

        If (ds IsNot Nothing AndAlso ds.Tables.Count > 0) Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim Cert As New CadastroCertidaoELicenca
                Cert.CodigoEmpresa = row("Empresa_Id")
                Cert.CodigoTipo = row("Tipo")
                Cert.DataEmissao = row("DataEmissao")
                Cert.DataVencimento = row("DataVencimento")
                Cert.Email = row("Email")
                Cert.AvisoDias = row("AvisoDias")
                Cert.Observacao = row("Observacao")
                Add(Cert)
            Next
        End If
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CadastroCertidaoELicenca

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As String)
        Selecionar(Codigo)
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoEmpresa As String
    Private _DescEmp As String
    Private _DataEmissao As Date
    Private _DataVencimento As Date
    Private _Email As String
    Private _AvisoDias As Integer
    Private _Observacao As String
    Private _Tipo As TipoDeCertidao
    Private _CodigoTipo As Integer
    Private _DescricaoTipo As String
#End Region

#Region "Properties"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property
    Public Property DescEmp() As String
        Get
            Return _DescEmp
        End Get
        Set(ByVal value As String)
            _DescEmp = value
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
    Public Property DataVencimento() As Date
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As Date)
            _DataVencimento = value
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
    Public Property AvisoDias() As Integer
        Get
            Return _AvisoDias
        End Get
        Set(ByVal value As Integer)
            _AvisoDias = value
        End Set
    End Property
    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property
    Public Property Tipo() As TipoDeCertidao
        Get
            If _Tipo Is Nothing AndAlso _CodigoTipo <> 0 Then _Tipo = New TipoDeCertidao(_CodigoTipo)
            Return _Tipo
        End Get
        Set(ByVal value As TipoDeCertidao)
            _Tipo = value
        End Set
    End Property
    Public Property CodigoTipo() As Integer
        Get
            If _Tipo Is Nothing Then
                _Tipo = New TipoDeCertidao(_CodigoTipo)
            End If

            Return _CodigoTipo
        End Get
        Set(ByVal value As Integer)
            _CodigoTipo = value
        End Set
    End Property
    Public ReadOnly Property DescricaoTipo() As String
        Get
            Dim desc As String = ""
            If Me.Tipo IsNot Nothing Then desc = Me.Tipo.Codigo & " - " & Me.Tipo.Descricao
            Return desc
        End Get
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
                Sql = " INSERT INTO CertidaoLicenca (Empresa_Id, Tipo_Id, DataEmissao, DataVencimento, Email, AvisoDias, Observacao) " & vbCrLf & _
                       " VALUES ('" & Me.CodigoEmpresa & "' , " & Me.CodigoTipo & " , '" & Me.DataEmissao.ToSqlDate() & "' , '" & Me.DataVencimento.ToSqlDate() & "' , '" & Me.Email & "' , " & Me.AvisoDias & " , '" & Me.Observacao & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE CertidaoLicenca SET" & vbCrLf & _
                      "    Empresa_Id         = '" & Me.CodigoEmpresa & "', " & vbCrLf & _
                      "    Tipo_Id            =  " & Me.CodigoTipo & ", " & vbCrLf & _
                      "    DataEmissao        = '" & Me.DataEmissao.ToSqlDate() & "', " & vbCrLf & _
                      "    DataVencimento     = '" & Me.DataVencimento.ToSqlDate() & "', " & vbCrLf & _
                      "    Email              = '" & Me.Email & "', " & vbCrLf & _
                      "    AvisoDias          =  " & Me.AvisoDias & ", " & vbCrLf & _
                      "    Observacao         = '" & Me.Observacao & "'" & vbCrLf & _
                      "  WHERE Empresa_Id     = '" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "  AND Tipo_Id          =  " & Me.CodigoTipo
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE CertidaoLicenca" & vbCrLf & _
                      "  WHERE Empresa_Id = '" & Me.CodigoEmpresa & "'" & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Function Selecionar(ByVal Codigo As String) As Boolean
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String = "SELECT Empresa_Id, DataEmissao, DataVencimento, Email, AvisoDias, Observacao" & _
                                   "  FROM CertidaoLicenca " & _
                                   " WHERE Empresa_Id = " & Codigo

            Dim ds As DataSet = objBanco.ConsultaDataSet(strSQL, "CertidaoLicenca")

            With ds.Tables(0).Rows
                If .Count > 0 Then
                    Me.CodigoEmpresa = .Item(0)("Empresa_Id").ToString()
                    Me.DataEmissao = .Item(0)("DataEmissao")
                    Me.DataVencimento = .Item(0)("DataVencimento")
                    Me.Email = .Item(0)("Email").ToString()
                    Me.AvisoDias = .Item(0)("AvisoDias")
                    Me.Observacao = .Item(0)("Observacao").ToString()
                    Return True
                Else : Return False
                End If
            End With
        Catch ex As Exception
            Throw New Exception
        Finally
            objBanco = Nothing
        End Try
    End Function
#End Region

End Class