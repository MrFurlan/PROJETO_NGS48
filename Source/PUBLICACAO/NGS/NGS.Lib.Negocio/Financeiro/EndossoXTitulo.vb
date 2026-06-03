Imports System.Web

Public Class ListEndossoXTitulo
    Inherits List(Of EndossoXTitulo)
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal objEndosso As Endosso)

        Me.Endosso = objEndosso

        Dim db As New AcessaBanco()

        Try
            Dim strSQL As String

            strSQL = "Select Empresa_Id, EndEmpresa_Id, Codigo_Id, Titulo_Id " & vbCrLf & _
                     "From EndossoXTitulo " & vbCrLf & _
                     "Where Empresa_id    = '" & Me.Endosso.CodigoEmpresa & "'" & vbCrLf & _
                     "  and EndEmpresa_Id = " & Me.Endosso.EnderecoEmpresa & vbCrLf

            If Me.Endosso.Codigo > 0 Then strSQL &= "  and Codigo_Id      = " & Me.Endosso.Codigo & vbCrLf

            Dim ds As DataSet = db.ConsultaDataSet(strSQL, "EndossoXTitulo")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim eXt As New EndossoXTitulo(Me.Endosso)

                eXt.CodigoTitulo = row("Titulo_Id")

                Me.Add(eXt)
            Next

        Catch ex As Exception
            Throw ex
        Finally
            db = Nothing
        End Try
    End Sub

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each item As EndossoXTitulo In Me
            item.IUD = Me.Endosso.IUD

            item.SalvarSql(Sqls)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _Endosso As Endosso
#End Region

#Region "Property"
    Public Property Endosso() As Endosso
        Get
            Return _Endosso
        End Get
        Set(ByVal value As Endosso)
            _Endosso = value
        End Set
    End Property

#End Region

End Class

<Serializable()> _
Public Class EndossoXTitulo
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String = ""

    Private _Endosso As Endosso

    Private _CodigoTitulo As String
    Private _Titulo As Titulo

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

    Public Property Endosso() As Endosso
        Get
            Return _Endosso
        End Get
        Set(ByVal value As Endosso)
            _Endosso = value
        End Set
    End Property

    Public Property CodigoTitulo() As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(ByVal value As Integer)
            _CodigoTitulo = value
        End Set
    End Property

    Public Property Titulo() As Titulo
        Get
            If _Titulo Is Nothing And Me.CodigoTitulo > 0 Then _Titulo = New Titulo(Me.CodigoTitulo)
            Return _Titulo
        End Get
        Set(ByVal value As Titulo)
            _Titulo = value
        End Set
    End Property

#End Region

#Region "Construtor"
    Public Sub New()
    End Sub

    Public Sub New(ByVal Endosso As Endosso)
        Me.Endosso = Endosso
    End Sub
#End Region

#Region "Methods"

    Public Function Salvar() As Boolean
        Dim objBanco As New AcessaBanco()
        Dim sqls As New ArrayList

        SalvarSql(sqls)
        Return objBanco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSql As String = ""

        Select Case Me.IUD
            Case "I"

                strSql = "INSERT INTO EndossoXTitulo(Empresa_Id, EndEmpresa_Id, Codigo_Id, Titulo_Id)" & vbCrLf & _
                         "Values('" & Endosso.CodigoEmpresa & "'," & Endosso.EnderecoEmpresa & "," & Endosso.Codigo & "," & Me.CodigoTitulo & ")"

                Sqls.Add(strSql)

                'Grava Bloqueio no Título
                Titulo.IUD = "U"
                Titulo.CodigoSituacao = eSituacao.EndossoTitulo
                Dim obs As String = Titulo.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Endosso alterou a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Endosso alterou a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                Titulo.ObservacoesControleInterno = obs

                Titulo.SalvarSql(Sqls)

            Case "D"

                'Libera Bloqueio do Título
                Titulo.IUD = "U"
                Titulo.CodigoSituacao = eSituacao.Normal
                Dim obs As String = Titulo.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Endosso excluído, alterado a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Endosso excluído, alterado a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                Titulo.ObservacoesControleInterno = obs

                Titulo.SalvarSql(Sqls)

                'Exclui TítuloXEndosso
                strSql = "Delete EndossoXTitulo " & vbCrLf & _
                           " Where Empresa_Id    = '" & Endosso.CodigoEmpresa & "'" & vbCrLf & _
                           "   and EndEmpresa_Id = " & Endosso.EnderecoEmpresa & vbCrLf & _
                           "   and Codigo_Id     = " & Endosso.Codigo

                Sqls.Add(strSql)


            Case "F"

                'Libera Bloqueio do Título
                Titulo.IUD = "U"
                Titulo.CodigoSituacao = eSituacao.Normal
                Dim obs As String = Titulo.ObservacoesControleInterno
                If Not String.IsNullOrWhiteSpace(obs) AndAlso obs.Length > 0 Then
                    obs = obs & ". Endosso finalizado, alterado a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                Else
                    obs = "Endosso finalizado, alterado a Situação do Título para " & Titulo.CodigoSituacao.ToString() & "-" & Titulo.Situacao.Descricao & " em " & Format(Now, "dd/MM/yyyy HH:mm:ss") & " feita por " & HttpContext.Current.Session("ssNomeUsuario")
                End If

                Titulo.ObservacoesControleInterno = obs

                Titulo.SalvarSql(Sqls)

        End Select

    End Sub

#End Region

End Class
