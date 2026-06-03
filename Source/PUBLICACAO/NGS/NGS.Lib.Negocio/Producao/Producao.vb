<Serializable()> _
Public Class lstProducao
    Inherits List(Of Producao)

#Region "Builders"

    Public Sub New()
        Dim sql As String = " SELECT Producao_id, UnidadeDeNegocio, Empresa, EndEmpresa, " & vbCrLf & _
                            "        Deposito, EndDeposito, Movimento, Etapa, Safra, Observacao" & vbCrLf & _
                            "   FROM NewProducao                                                                " & vbCrLf

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ProducaoXItem")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                Dim obj As New Producao()
                With obj
                    .CodigoProducao = row("Producao_id")
                    .UnidadeDeNegocio = row("UnidadeDeNegocio")
                    .CodigoEmpresa = row("Empresa")
                    .EndEmpresa = row("EndEmpresa")
                    .CodigoDeposito = row("Deposito")
                    .EndDeposito = row("EndDeposito")
                    .Movimento = row("Movimento")
                    .Etapa = row("Etapa")
                    .Safra = row("Safra")
                    .Observacao = row("Observacao")
                End With
                Me.Add(obj)
            Next

        End If
    End Sub

#End Region

End Class



<Serializable()> _
Public Class Producao

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal producao As Integer)
        Dim sql As String = " SELECT Producao_id, UnidadeDeNegocio, Empresa, EndEmpresa,         " & vbCrLf & _
                            "        Deposito, EndDeposito, Movimento, Etapa, Safra, Observacao  " & vbCrLf & _
                            "   FROM NewProducao                                                 " & vbCrLf

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ProducaoXItem")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

            For Each row As DataRow In ds.Tables(0).Rows
                With Me
                    .CodigoProducao = row("Producao_id")
                    .UnidadeDeNegocio = row("UnidadeDeNegocio")
                    .CodigoEmpresa = row("Empresa")
                    .EndEmpresa = row("EndEmpresa")
                    .CodigoDeposito = row("Deposito")
                    .EndDeposito = row("EndDeposito")
                    .Movimento = row("Movimento")
                    .Etapa = row("Etapa")
                    .Safra = row("Safra")
                    .Observacao = row("Observacao")
                End With
            Next
        End If
    End Sub


#End Region

#Region "fields"
    Private _IUD As String
    Private _CodigoProducao As Integer
    Private _UnidadeDeNegocio As String
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _CodigoDeposito As String
    Private _EndDeposito As String
    Private _Movimento As DateTime
    Private _Etapa As Integer
    Private _Safra As String
    Private _Observacao As String
    Private _lstProducaoXItem As lstProducaoXItem
    Public Property lstProducaoXItem() As lstProducaoXItem
        Get
            If _lstProducaoXItem Is Nothing AndAlso _CodigoProducao > 0 Then _lstProducaoXItem = New [Lib].Negocio.lstProducaoXItem(Me.CodigoProducao)
            Return _lstProducaoXItem
        End Get
        Set(ByVal value As lstProducaoXItem)
            _lstProducaoXItem = value
        End Set
    End Property
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

    Public Property CodigoProducao() As Integer
        Get
            Return _CodigoProducao
        End Get
        Set(ByVal value As Integer)
            _CodigoProducao = value
        End Set
    End Property

    Public Property UnidadeDeNegocio() As String
        Get
            Return _UnidadeDeNegocio
        End Get
        Set(ByVal value As String)
            _UnidadeDeNegocio = value
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

    Public Property EndEmpresa() As String
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As String)
            _EndEmpresa = value
        End Set
    End Property

    Public Property CodigoDeposito() As String
        Get
            Return _CodigoDeposito
        End Get
        Set(ByVal value As String)
            _CodigoDeposito = value
        End Set
    End Property

    Public Property EndDeposito() As String
        Get
            Return _EndDeposito
        End Get
        Set(ByVal value As String)
            _EndDeposito = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property Etapa() As Integer
        Get
            Return _Etapa
        End Get
        Set(ByVal value As Integer)
            _Etapa = value
        End Set
    End Property

    Public Property Safra() As String
        Get
            Return _Safra
        End Get
        Set(ByVal value As String)
            _Safra = value
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
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim db As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return db.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""
        Select Case _IUD
            Case "I"
                sql = "INSERT INTO NewProducao (UnidadeDeNegocio, Empresa, EndEmpresa, Deposito, EndDeposito," & vbCrLf & _
                    "                           Movimento, Etapa, Safra, Observacao, UsuarioInclusao, UsuarioInclusaoData)" & vbCrLf & _
                    String.Format("       VALUES ('{0}', '{1}', {2}, '{3}', {4}, '{5}', {6}, '{7}', '{8}', '{9}', '{10}');", _
                                  Me.UnidadeDeNegocio, Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoDeposito, Me.EndDeposito, Me.Movimento, _
                                  Me.Etapa, Me.Safra, Me.Observacao, DateTime.Now, UsuarioServidor.NomeUsuario) & vbCrLf
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = String.Format("UPDATE NewProducao SET" & vbCrLf & _
                              "             UnidadeDeNegocio = {0}, Empresa = {1}, EndEmpresa = {2}, Deposito = {3}, EndDeposito = {4}," & vbCrLf & _
                              "             Movimento = {5}, Etapa = {6}, Safra = {7}, Observacao = {8}" & vbCrLf & _
                              "       WHERE Producao_Id = {9};", Me.UnidadeDeNegocio, Me.CodigoEmpresa, Me.EndEmpresa, Me.CodigoDeposito, Me.EndDeposito, Me.Movimento, _
                                                                 Me.Etapa, Me.Safra, Me.Observacao)

                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                sql = "DELETE NewProducao " & vbCrLf & _
                      "  WHERE  Producao_Id = " & Me.CodigoProducao

                SalvarTabelasRelacionadasSql(Sqls)
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Me.lstProducaoXItem IsNot Nothing AndAlso Me.lstProducaoXItem.Count > 0 Then
            For Each obj As ProducaoXItem In Me.lstProducaoXItem
                If Me.IUD <> "U" Then
                    obj.IUD = Me.IUD
                ElseIf obj.IUD = "" Then
                    obj.IUD = "U"
                End If
                obj.SalvarSql(Sqls)
            Next
        End If
    End Sub
#End Region

End Class
