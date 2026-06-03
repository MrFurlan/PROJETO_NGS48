<Serializable()> _
Public Class lstProducaoXItem
    Inherits List(Of ProducaoXItem)

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal producao As Integer)
        Dim sql As String = " SELECT pxi.Producao_id, pxi.Produto_id, pxi.Operacao, pxi.SubOperacao, pxi.Lote, pxi.Classificacao, " & vbCrLf & _
                            "        pxi.Embalagem, pxi.TipoDeEmbalagem, pxi.CapacidadeEmbalagem, pxi.QtdeFiscal, pxi.QtdeFisica  " & vbCrLf & _
                            "   FROM NewProducaoxItem AS pxi                                                                      " & vbCrLf & _
                            "  Inner Join NewProducao p                                                                           " & vbCrLf & _
                            "     on p.Producao_Id = pxi.Producao_Id                                                              " & vbCrLf & _
                            "  where pxi.Producao_id = " & producao & ";                                                          " & vbCrLf

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ProducaoXItem")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Dim obj As New ProducaoXItem()
            For Each row As DataRow In ds.Tables(0).Rows
                With obj
                    .CodigoProducao = row("Producao_id")
                    .CodigoProduto = row("Produto_id")
                    .CodigoOperacao = row("Operacao")
                    .CodigoSubOperacao = row("SubOperacao")
                    .Lote = row("Lote")
                    .Classificacao = row("Classificacao")
                    .Embalagem = row("Embalagem")
                    .TipoDeEmbalagem = row("TipoDeEmbalagem")
                    .CapacidadeDeEmbalagem = row("CapacidadeEmbalagem")
                    .QtdeFiscal = row("QtdeFiscal")
                    .QtdeFisica = row("QtdeFisica")
                End With
                Me.Add(obj)
            Next
        End If
    End Sub


#End Region

End Class

<Serializable()> _
Public Class ProducaoXItem

#Region "Builders"

    Public Sub New()

    End Sub

    Public Sub New(ByVal producao As Integer, ByVal produto As String)
        Dim sql As String = " SELECT pxi.Producao_id, pxi.Produto_id, pxi.Operacao, pxi.SubOperacao, pxi.Lote, pxi.Classificacao, " & vbCrLf & _
                            "        pxi.Embalagem, pxi.TipoDeEmbalagem, pxi.CapacidadeEmbalagem, pxi.QtdeFiscal, pxi.QtdeFisica  " & vbCrLf & _
                            "   FROM NewProducaoxItem AS pxi                                                                      " & vbCrLf & _
                            "  Inner Join NewProducao p                                                                           " & vbCrLf & _
                            "     on p.Producao_Id = pxi.Producao_Id                                                              " & vbCrLf & _
                            "  where pxi.Producao_id = " & producao & ";                                                          " & vbCrLf

        Dim ds As DataSet = New AcessaBanco().ConsultaDataSet(sql, "ProducaoXItem")

        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In ds.Tables(0).Rows
                With Me
                    .CodigoProducao = row("Producao_id")
                    .CodigoProduto = row("Produto_id")
                    .CodigoOperacao = row("Operacao")
                    .CodigoSubOperacao = row("SubOperacao")
                    .Lote = row("Lote")
                    .Classificacao = row("Classificacao")
                    .Embalagem = row("Embalagem")
                    .TipoDeEmbalagem = row("TipoDeEmbalagem")
                    .CapacidadeDeEmbalagem = row("CapacidadeEmbalagem")
                    .QtdeFiscal = row("QtdeFiscal")
                    .QtdeFisica = row("QtdeFisica")
                End With
            Next
        End If
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoProducao As Integer
    Private _CodigoProduto As String
    Private _CodigoOperacao As String
    Private _CodigoSubOperacao As String
    Private _Lote As String
    Private _Classificacao As String
    Private _Embalagem As Integer
    Private _TipoDeEmbalagem As String
    Private _CapacidadeDeEmbalagem As Integer
    Private _QtdeFisica As Decimal
    Private _QtdeFiscal As Decimal
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

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property CodigoOperacao() As String
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As String)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao() As String
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As String)
            _CodigoSubOperacao = value
        End Set
    End Property

    Public Property Lote() As String
        Get
            Return _Lote
        End Get
        Set(ByVal value As String)
            _Lote = value
        End Set
    End Property

    Public Property Classificacao() As String
        Get
            Return _Classificacao
        End Get
        Set(ByVal value As String)
            _Classificacao = value
        End Set
    End Property

    Public Property Embalagem() As Integer
        Get
            Return _Embalagem
        End Get
        Set(ByVal value As Integer)
            _Embalagem = value
        End Set
    End Property

    Public Property TipoDeEmbalagem() As String
        Get
            Return _TipoDeEmbalagem
        End Get
        Set(ByVal value As String)
            _TipoDeEmbalagem = value
        End Set
    End Property

    Public Property CapacidadeDeEmbalagem() As Integer
        Get
            Return _CapacidadeDeEmbalagem
        End Get
        Set(ByVal value As Integer)
            _CapacidadeDeEmbalagem = value
        End Set
    End Property

    Public Property QtdeFisica() As Decimal
        Get
            Return _QtdeFisica
        End Get
        Set(ByVal value As Decimal)
            _QtdeFisica = value
        End Set
    End Property

    Public Property QtdeFiscal() As Decimal
        Get
            Return _QtdeFiscal
        End Get
        Set(ByVal value As Decimal)
            _QtdeFiscal = value
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
                sql = "INSERT INTO NewProducaoXItem (Producao_id, Produto_id, Operacao, SubOperacao, Lote, Classificacao, " & vbCrLf & _
                      "                              Embalagem, TipoDeEmbalagem, CapacidadeEmbalagem, QtdeFiscal, QtdeFisica)" & vbCrLf & _
                    String.Format("          VALUES ('{0}', '{1}', {2}, '{3}', {4}, '{5}', {6}, '{7}', '{8}', '{9}', '{10}');", _
                                                     Me.CodigoProducao, Me.CodigoProduto, Me.CodigoOperacao, Me.CodigoSubOperacao, Me.Lote, Me.Classificacao, _
                                                     Me.Embalagem, Me.TipoDeEmbalagem, Me.CapacidadeDeEmbalagem, Me.QtdeFiscal, Me.QtdeFisica)
                Sqls.Add(sql)
                'SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = String.Format("UPDATE NewProducaoXItem SET" & vbCrLf & _
                              "             Producao_id = {0}, Produto_id = '{1}', Operacao = {2}, SubOperacao = {3}, Lote = '{4}', Classificacao = '{5}', " & vbCrLf & _
                              "             Embalagem = {6}, TipoDeEmbalagem = '{7}', CapacidadeEmbalagem = '{8}', QtdeFiscal = '{9}', QtdeFisica = '{10}' " & vbCrLf & _
                              "       WHERE Producao_Id = {0}", _
                                            Me.CodigoProducao, Me.CodigoProduto, Me.CodigoOperacao, Me.CodigoSubOperacao, Me.Lote, Me.Classificacao, _
                                            Me.Embalagem, Me.TipoDeEmbalagem, Me.CapacidadeDeEmbalagem, Me.QtdeFiscal, Me.QtdeFisica)
                Sqls.Add(sql)
                'SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                sql = " DELETE NewProducaoXItem " & vbCrLf & _
                      "  WHERE Producao_Id = " & Me.CodigoProducao

                'SalvarTabelasRelacionadasSql(Sqls)
                Sqls.Add(sql)
        End Select
    End Sub

    'Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
    '    If Me.lstProducaoXItem IsNot Nothing AndAlso Me.lstProducaoXItem.Count > 0 Then
    '        For Each obj As ProducaoXItem In Me.lstProducaoXItem
    '            If Me.IUD <> "U" Then
    '                obj.IUD = Me.IUD
    '            ElseIf obj.IUD = "" Then
    '                obj.IUD = "U"
    '            End If
    '            obj.SalvarSql(Sqls)
    '        Next
    '    End If
    'End Sub
#End Region

End Class
