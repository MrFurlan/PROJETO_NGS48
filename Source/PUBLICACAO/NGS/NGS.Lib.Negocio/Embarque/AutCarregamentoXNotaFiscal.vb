Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListAutCarregamentoXNotaFiscal
    Inherits List(Of AutCarregamentoXNotaFiscal)

    Public Sub New()
    End Sub

    Public Sub New(ByVal objAutCarregamento As AutCarregamento)
        Dim sql As String = ""
        sql &= "select * from AutCarregamentoXNotaFiscal where Carregamento_Id = '" & objAutCarregamento.Carregamento_Id & "'"

        Dim db As New AcessaBanco
        Dim ds As DataSet = db.ConsultaDataSet(sql, "AutCarregamentoXNotaFiscal")

        For Each row In ds.Tables(0).Rows
            Dim obj As New AutCarregamentoXNotaFiscal
            obj.Carregamento_Id = row("Carregamento_Id")
            obj.Empresa_Id = row("Empresa_Id")
            obj.EndEmpresa_Id = row("EndEmpresa_Id")
            obj.Cliente_Id = row("Cliente_Id")
            obj.EndCliente_Id = row("EndCliente_Id")
            obj.EntradaSaida_Id = IIf(row("EntradaSaida_Id") = "E", [Lib].Negocio.eEntradaSaida.Entrada, [Lib].Negocio.eEntradaSaida.Saida)
            obj.Serie_Id = row("Serie_Id")
            obj.Nota_Id = row("Nota_Id")
            obj.CFOP_Id = row("CFOP_Id")
            obj.Sequencia_Id = row("Sequencia_Id")
            obj.Produto_Id = row("Produto_Id")
            obj.QuantidadeOrigem = row("QuantidadeOrigem")
            obj.QuantidadeDestino = row("QuantidadeDestino")
            Me.Add(obj)
        Next
    End Sub

End Class

<Serializable()> _
Public Class AutCarregamentoXNotaFiscal
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String
    Private _ParentAutCarregamentoxItens As AutCarregamentoXItens
    Private _Carregamento_Id As String
    Private _Empresa_Id As String
    Private _EndEmpresa_Id As String
    Private _Cliente_Id As String
    Private _EndCliente_Id As String
    Private _EntradaSaida_Id As [Lib].Negocio.eEntradaSaida
    Private _Serie_Id As String
    Private _Nota_Id As String
    Private _CFOP_Id As String
    Private _Sequencia_Id As String
    Private _Produto_Id As String
    Private _QuantidadeOrigem As String
    Private _QuantidadeDestino As String
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

    Public Property Carregamento_Id() As String
        Get
            Return _Carregamento_Id
        End Get
        Set(ByVal value As String)
            _Carregamento_Id = value
        End Set
    End Property

    Public Property Empresa_Id() As String
        Get
            Return _Empresa_Id
        End Get
        Set(ByVal value As String)
            _Empresa_Id = value
        End Set
    End Property

    Public Property EndEmpresa_Id() As String
        Get
            Return _EndEmpresa_Id
        End Get
        Set(ByVal value As String)
            _EndEmpresa_Id = value
        End Set
    End Property

    Public Property Cliente_Id() As String
        Get
            Return _Cliente_Id
        End Get
        Set(ByVal value As String)
            _Cliente_Id = value
        End Set
    End Property

    Public Property EndCliente_Id() As String
        Get
            Return _EndCliente_Id
        End Get
        Set(ByVal value As String)
            _EndCliente_Id = value
        End Set
    End Property

    Public Property EntradaSaida_Id() As [Lib].Negocio.eEntradaSaida
        Get
            Return _EntradaSaida_Id
        End Get
        Set(ByVal value As [Lib].Negocio.eEntradaSaida)
            _EntradaSaida_Id = value
        End Set
    End Property

    Public Property Serie_Id() As String
        Get
            Return _Serie_Id
        End Get
        Set(ByVal value As String)
            _Serie_Id = value
        End Set
    End Property

    Public Property Nota_Id() As String
        Get
            Return _Nota_Id
        End Get
        Set(ByVal value As String)
            _Nota_Id = value
        End Set
    End Property

    Public Property CFOP_Id() As String
        Get
            Return _CFOP_Id
        End Get
        Set(ByVal value As String)
            _CFOP_Id = value
        End Set
    End Property

    Public Property Sequencia_Id() As String
        Get
            Return _Sequencia_Id
        End Get
        Set(ByVal value As String)
            _Sequencia_Id = value
        End Set
    End Property

    Public Property Produto_Id() As String
        Get
            Return _Produto_Id
        End Get
        Set(ByVal value As String)
            _Produto_Id = value
        End Set
    End Property

    Public Property QuantidadeOrigem() As String
        Get
            Return _QuantidadeOrigem
        End Get
        Set(ByVal value As String)
            _QuantidadeOrigem = value
        End Set
    End Property

    Public Property QuantidadeDestino() As String
        Get
            Return _QuantidadeDestino
        End Get
        Set(ByVal value As String)
            _QuantidadeDestino = value
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

        If db.GravaBanco(Sqls) Then
            IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        Dim sql As String = ""
        Select Case _IUD
            Case "I"
                sql = "INSERT INTO AutCarregamentoXNotaFiscal (Carregamento_Id, Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, CFOP_Id, Sequencia_Id, Produto_Id, QuantidadeOrigem, QuantidadeDestino) " & vbCrLf & _
                      "VALUES (" & Me.Carregamento_Id & ", '" & Me.Empresa_Id & "'," & Me.EndEmpresa_Id & ",'" & Me.Cliente_Id & "', " & Me.EndCliente_Id & ", '" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "','" & Me.Serie_Id & "', " & Me.Nota_Id & ", " & Me.CFOP_Id & ", " & Me.Sequencia_Id & ", '" & Me.Produto_Id & "', " & Str(Me.QuantidadeOrigem) & ", " & Str(Me.QuantidadeDestino) & ")"
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "UPDATE AutCarregamentoXNotaFiscal SET " & _
                "         QuantidadeOrigem = '" & Me.QuantidadeOrigem & "'" & vbCrLf & _
                "       , QuantidadeDestino = '" & Me.QuantidadeDestino & "'" & vbCrLf & _
                "   WHERE 1=1 " & vbCrLf & _
                "   AND Empresa_Id = '" & Me.Empresa_Id & "'" & vbCrLf & _
                "   AND EndEmpresa_Id = " & Me.EndEmpresa_Id & vbCrLf & _
                "   AND Cliente_Id = '" & Me.Cliente_Id & "'" & vbCrLf & _
                "   AND EndCliente_Id = " & Me.EndCliente_Id & vbCrLf & _
                "   AND EntradaSaida_Id = '" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                "   AND Serie_Id = '" & Me.Serie_Id & "'" & vbCrLf & _
                "   AND Nota_Id = " & Me.Nota_Id & vbCrLf & _
                "   AND CFOP_Id = " & Me.CFOP_Id & vbCrLf & _
                "   AND Sequencia_Id = " & Me.Sequencia_Id & vbCrLf & _
                "   AND Produto_Id = '" & Str(Me.Produto_Id) & "'" & vbCrLf & _
                "   AND Carregamento_Id = " & Str(Me.Carregamento_Id)
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                sql = "DELETE AutCarregamentoXNotaFiscal " & _
                      "   WHERE 1=1 " & vbCrLf & _
                      "   AND Empresa_Id = '" & Me.Empresa_Id & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Me.EndEmpresa_Id & vbCrLf & _
                      "   AND Cliente_Id = '" & Me.Cliente_Id & "'" & vbCrLf & _
                      "   AND EndCliente_Id = " & Me.EndCliente_Id & vbCrLf & _
                      "   AND EntradaSaida_Id = '" & IIf(Me.EntradaSaida_Id = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   AND Serie_Id = '" & Me.Serie_Id & "'" & vbCrLf & _
                      "   AND Nota_Id = " & Me.Nota_Id & vbCrLf & _
                      "   AND CFOP_Id = " & Me.CFOP_Id & vbCrLf & _
                      "   AND Sequencia_Id = " & Me.Sequencia_Id & vbCrLf & _
                      "   AND Produto_Id = '" & Me.Produto_Id & "'" & vbCrLf & _
                      "   AND Carregamento_Id = " & Me.Carregamento_Id
                SalvarTabelasRelacionadasSql(Sqls)
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        'If Me.Transportadores IsNot Nothing AndAlso Me.Transportadores.Count > 0 Then
        '    For Each item As EmbarqueTransportador In Me.Transportadores
        '        item.IUD = Me.IUD
        '        item.SalvarSql(Sqls)
        '    Next
        'End If
    End Sub
#End Region

End Class
