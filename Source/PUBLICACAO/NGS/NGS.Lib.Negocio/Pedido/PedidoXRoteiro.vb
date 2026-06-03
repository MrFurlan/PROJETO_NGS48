
'***********************************************************************************************************
'*********************************  Lista de Roteiros do Pedido  ******************************************
'***********************************************************************************************************
Public Class ListPedidoXRoteiro
    Inherits List(Of PedidoXRoteiro)

#Region "Construtors"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        Pedido = pPedido
        If pPedido.Codigo > 0 Then

            Dim banco As New AcessaBanco

            Try
                Dim sql As String = ""
                sql = " SELECT Empresa_Id, EndEmpresa_Id, Pedido_Id, TipoOrigem_Id, Origem_Id, EndOrigem_Id, " & vbCrLf & _
                      "        TipoDestino_Id, Destino_Id, EndDestino_Id, ViaDeTransporte_Id, " & vbCrLf & _
                      "        ISNULL(QuantidadeOriginal,0) AS QuantidadeOriginal, ISNULL(QuantidadeAtual,0) AS QuantidadeAtual, " & vbCrLf & _
                      "        ISNULL(Valor, 0) AS Valor, ISNULL(Log,'') Log " & vbCrLf & _
                      "   FROM PedidoXRoteiro " & vbCrLf & _
                      "  WHERE Empresa_Id    ='" & Me.Pedido.CodigoEmpresa & "' " & _
                      "    AND EndEmpresa_Id = " & Me.Pedido.EnderecoEmpresa & _
                      "    AND Pedido_Id     = " & Me.Pedido.Codigo

                Dim ds As DataSet = banco.ConsultaDataSet(sql, "ListaDePedidoXRoteiro")

                For Each row As DataRow In ds.Tables(0).Rows
                    Dim Roteiro As New PedidoXRoteiro(Pedido)
                    Roteiro.TipoOrigem = row("TipoOrigem_Id")
                    Roteiro.CodigoOrigem = row("Origem_Id")
                    Roteiro.EnderecoOrigem = row("EndOrigem_Id")

                    Roteiro.TipoDestino = row("TipoDestino_Id")
                    Roteiro.CodigoDestino = row("Destino_id")
                    Roteiro.EnderecoDestino = row("EndDestino_Id")

                    Roteiro.CodigoViaDeTransporte = row("ViaDeTransporte_Id")
                    Roteiro.QuantidadeOriginal = row("QuantidadeOriginal")
                    Roteiro.QuantidadeAtual = row("QuantidadeAtual")
                    Roteiro.Valor = row("Valor")
                    Roteiro.Log = row("Log")
                    Me.Add(Roteiro)

                Next
            Catch ex As Exception
            Finally
                banco = Nothing
            End Try
        End If
    End Sub
#End Region

#Region "Fields"
    Private _Pedido As Negocio.Pedido
    Private _TabelaViaDeTransporte As Negocio.ViaDeTransportes
#End Region

#Region "Properties"
    Public Property Pedido() As Negocio.Pedido
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Negocio.Pedido)
            _Pedido = value
        End Set
    End Property

    Public ReadOnly Property TabelaViaDeTransporte As Negocio.ViaDeTransportes
        Get
            If _TabelaViaDeTransporte Is Nothing Then _TabelaViaDeTransporte = New ViaDeTransportes(True)
            Return _TabelaViaDeTransporte
        End Get
    End Property


    Public ReadOnly Property TotalQtdDEOD As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "DE" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "DE" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property

    Public ReadOnly Property TotalQtdLEOD As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "LE" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "LE" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property


    Public ReadOnly Property TotalQtdDETR As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "DE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "DE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property


    Public ReadOnly Property TotalQtdLETR As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "LE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "LE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property


    Public ReadOnly Property TotalQtdTROD As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "TR" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "TR" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property

    Public ReadOnly Property TotalQtdEmbarque As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where (a.TipoOrigem = "OD" And a.TipoDestino = "DE") Or (a.TipoOrigem = "OD" And a.TipoDestino = "LE") Or (a.TipoOrigem = "OD" And a.TipoDestino = "TR") Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "DE" Or a.TipoOrigem = "LE" Select a.QuantidadeAtual).Sum
            End If
        End Get
    End Property

    Public ReadOnly Property TotalQtdDestino As Decimal
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where (a.TipoOrigem = "OD" And a.TipoDestino = "DE") Or (a.TipoOrigem = "OD" And a.TipoDestino = "LE") Or (a.TipoOrigem = "TR" And a.TipoDestino = "LE") Or (a.TipoOrigem = "TR" And a.TipoDestino = "DE") Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where (a.TipoOrigem = "DE" And a.TipoDestino = "OD") Or (a.TipoOrigem = "LE" And a.TipoDestino = "OD") Or (a.TipoOrigem = "TR" And a.TipoDestino = "OD") Select a.QuantidadeAtual).Sum
            End If
        End Get
    End Property

    Public ReadOnly Property TotalQtd As Decimal
        Get

            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "DE" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoDestino = "LE" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoDestino = "DE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoDestino = "LE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum

            Else
                Return (From a In Me Where a.TipoOrigem = "DE" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoOrigem = "LE" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoOrigem = "DE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum + _
                (From a In Me Where a.TipoOrigem = "LE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum
            End If

        End Get
    End Property

    Public ReadOnly Property VerificaQuantidadeTransbordo() As Integer
        Get
            If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                Return (From a In Me Where a.TipoDestino = "DE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum + _
                   (From a In Me Where a.TipoDestino = "LE" And a.TipoOrigem = "TR" Select a.QuantidadeAtual).Sum - _
                   (From a In Me Where a.TipoDestino = "TR" And a.TipoOrigem = "OD" Select a.QuantidadeAtual).Sum
            Else
                Return (From a In Me Where a.TipoOrigem = "DE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum + _
                   (From a In Me Where a.TipoOrigem = "LE" And a.TipoDestino = "TR" Select a.QuantidadeAtual).Sum - _
                   (From a In Me Where a.TipoOrigem = "TR" And a.TipoDestino = "OD" Select a.QuantidadeAtual).Sum
            End If
        End Get
    End Property

#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each Roteiro As PedidoXRoteiro In Me
            If Pedido.IUD = "I" Or Pedido.IUD = "D" Then Roteiro.IUD = Pedido.IUD
            Roteiro.SalvarSql(Sqls)
        Next
    End Sub
#End Region


End Class

'***********************************************************************************************************
'**************************************   Roteiros do Pedido   ********************************************
'***********************************************************************************************************
Public Class PedidoXRoteiro


#Region "Construtors"
    Public Sub New(ByVal pPedido As Negocio.Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _Pedido As Negocio.Pedido

    Private _TipoOrigem As String
    Private _CodigoOrigem As String
    Private _EnderecoOrigem As Integer
    Private _Origem As Negocio.Cliente

    Private _TipoDestino As String
    Private _CodigoDestino As String
    Private _EnderecoDestino As Integer
    Private _Destino As Negocio.Cliente

    Private _ViaDeTransporte As Negocio.ViaDeTransporte
    Private _TabelaViaDeTransporte As Negocio.ViaDeTransportes
    Private _CodigoViaDeTransporte As Integer
    Private _QuantidadeOriginal As Decimal
    Private _QuantidadeAtual As Decimal
    Private _Valor As Decimal
    Private _Log As String
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

    Public ReadOnly Property Pedido As Negocio.Pedido
        Get
            Return _Pedido
        End Get
    End Property

    '************************** ORIGEM *************************'
    Public Property TipoOrigem() As String
        Get
            Return _TipoOrigem
        End Get
        Set(ByVal value As String)
            _TipoOrigem = value
        End Set
    End Property

    Public Property CodigoOrigem() As String
        Get
            Return _CodigoOrigem
        End Get
        Set(ByVal value As String)
            _CodigoOrigem = value
            _Origem = Nothing
        End Set
    End Property

    Public Property EnderecoOrigem() As Integer
        Get
            Return _EnderecoOrigem
        End Get
        Set(ByVal value As Integer)
            _EnderecoOrigem = value
        End Set
    End Property

    Public Property Origem() As Cliente
        Get
            If _Origem Is Nothing Then _Origem = New Cliente(Me.CodigoOrigem, Me.EnderecoOrigem)
            Return _Origem
        End Get
        Set(ByVal value As Cliente)
            _Origem = value
        End Set
    End Property

    '************************ DESTINO *************************'
    Public Property TipoDestino() As String
        Get
            Return _TipoDestino
        End Get
        Set(ByVal value As String)
            _TipoDestino = value
        End Set
    End Property

    Public Property CodigoDestino As String
        Get
            Return _CodigoDestino
        End Get
        Set(ByVal value As String)
            _CodigoDestino = value
            _Destino = Nothing
        End Set
    End Property

    Public Property EnderecoDestino() As Integer
        Get
            Return _EnderecoDestino
        End Get
        Set(ByVal value As Integer)
            _EnderecoDestino = value
        End Set
    End Property

    Public Property Destino() As Cliente
        Get
            If _Destino Is Nothing Then _Destino = New Cliente(Me.CodigoDestino, Me.EnderecoDestino)
            Return _Destino
        End Get
        Set(ByVal value As Cliente)
            _Destino = value
        End Set
    End Property

    '*********************** OUTRAS INFORMAÇÕES ******************'
    Public Property ViaDeTransporte() As Negocio.ViaDeTransporte
        Get
            If _ViaDeTransporte Is Nothing And CodigoViaDeTransporte > 0 Then _ViaDeTransporte = New Negocio.ViaDeTransporte(CodigoViaDeTransporte)
            Return _ViaDeTransporte
        End Get
        Set(ByVal value As Negocio.ViaDeTransporte)
            _ViaDeTransporte = value
        End Set
    End Property

    Public Property CodigoViaDeTransporte() As Integer
        Get
            Return _CodigoViaDeTransporte
        End Get
        Set(ByVal value As Integer)
            _CodigoViaDeTransporte = value
            _ViaDeTransporte = Nothing
        End Set
    End Property

    Public ReadOnly Property TabelaViaDeTransporte As Negocio.ViaDeTransportes
        Get
            If _TabelaViaDeTransporte Is Nothing Then _TabelaViaDeTransporte = New ViaDeTransportes(True)
            Return _TabelaViaDeTransporte
        End Get
    End Property


    Public Property QuantidadeOriginal As Decimal
        Get
            Return _QuantidadeOriginal
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeOriginal = value
        End Set
    End Property

    Public Property QuantidadeAtual As Decimal
        Get
            Return _QuantidadeAtual
        End Get
        Set(ByVal value As Decimal)
            _QuantidadeAtual = value
        End Set
    End Property

    Public Property Valor() As Decimal
        Get
            Return _Valor
        End Get
        Set(ByVal value As Decimal)
            _Valor = value
        End Set
    End Property


    Public ReadOnly Property ValorPorTonelada() As Decimal
        Get
            Return Me.Valor * 1000
        End Get
    End Property

    Public Property Log() As String
        Get
            Return _Log
        End Get
        Set(ByVal value As String)
            _Log = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I", "U"
                Sql = ";MERGE PedidoXRoteiro as Dest" & vbCrLf & _
                      " USING (Select '" & Me.Pedido.CodigoEmpresa & "' as Empresa_Id," & Me.Pedido.EnderecoEmpresa & " as EndEmpresa_Id," & Me.Pedido.Codigo & " as Pedido_Id, " & vbCrLf & _
                      "               '" & Me.TipoOrigem & "' as TipoOrigem_Id, '" & Me.CodigoOrigem & "' as Origem_Id," & Me.EnderecoOrigem & " as EndOrigem_Id, " & vbCrLf & _
                      "               '" & Me.TipoDestino & "' as TipoDestino_Id, '" & Me.CodigoDestino & "' as Destino_Id," & Me.EnderecoDestino & " as EndDestino_Id, " & vbCrLf & _
                      "               '" & Me.CodigoViaDeTransporte & "' as ViaDeTransporte_Id ) AS Ori" & vbCrLf & _
                      "    ON Dest.Empresa_Id     = Ori.Empresa_Id" & vbCrLf & _
                      "   AND Dest.EndEmpresa_Id  = Ori.EndEmpresa_Id" & vbCrLf & _
                      "   AND Dest.Pedido_Id      = Ori.Pedido_Id" & vbCrLf & _
                      "   AND Dest.TipoOrigem_Id  = Ori.TipoOrigem_Id" & vbCrLf & _
                      "   AND Dest.Origem_Id      = Ori.Origem_Id" & vbCrLf & _
                      "   AND Dest.EndOrigem_Id   = Ori.EndOrigem_Id" & vbCrLf & _
                      "   AND Dest.TipoDestino_Id = Ori.TipoDestino_Id" & vbCrLf & _
                      "   AND Dest.Destino_Id     = Ori.Destino_Id" & vbCrLf & _
                      "   AND Dest.EndDestino_Id  = Ori.EndDestino_Id" & vbCrLf & _
                      "  WHEN NOT MATCHED " & vbCrLf & _
                      "       THEN INSERT (Empresa_Id, EndEmpresa_Id, Pedido_Id," & vbCrLf & _
                      "                    TipoOrigem_Id, Origem_Id, EndOrigem_Id, " & vbCrLf & _
                      "                    TipoDestino_Id, Destino_Id, EndDestino_Id, " & vbCrLf & _
                      "                    ViaDeTransporte_Id, QuantidadeOriginal, QuantidadeAtual, Valor, Log) " & vbCrLf & _
                      "            VALUES ('" & Me.Pedido.CodigoEmpresa & "'," & vbCrLf & _
                                             Me.Pedido.EnderecoEmpresa & ", " & vbCrLf & _
                                             Me.Pedido.Codigo & vbCrLf & _
                                             ",'" & Me.TipoOrigem & "','" & Me.CodigoOrigem & "'," & Me.EnderecoOrigem & vbCrLf & _
                                             ",'" & Me.TipoDestino & "', '" & Me.CodigoDestino & "'," & Me.EnderecoDestino & "," & vbCrLf & _
                                             Me.CodigoViaDeTransporte & "," & vbCrLf & _
                                             Str(Me.QuantidadeOriginal) & "," & vbCrLf & _
                                             Str(Me.QuantidadeAtual) & "," & vbCrLf & _
                                             Str(Me.Valor) & vbCrLf & _
                                             ",'" & Me.Log & "')" & vbCrLf & _
                    "   WHEN MATCHED " & vbCrLf & _
                    "        THEN UPDATE  " & vbCrLf & _
                    "                SET QuantidadeOriginal = " & Str(Me.QuantidadeOriginal) & "," & vbCrLf & _
                    "                    QuantidadeAtual = " & Str(Me.QuantidadeAtual) & "," & vbCrLf & _
                    "                    Valor =" & Str(Me.Valor) & "," & vbCrLf & _
                    "                    Log =  '" & Me.Log & "'; " & vbCrLf

                Sqls.Add(Sql)
            Case "D"
                Sql = "DELETE PedidoXRoteiro " & vbCrLf
                CompoeWhere(Sql)
                Sqls.Add(Sql)
        End Select
    End Sub

    Private Sub CompoeWhere(ByRef sql As String)
        sql &= " WHERE Empresa_Id = '" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf & _
            "   AND EndEmpresa_Id =" & Me.Pedido.EnderecoEmpresa & vbCrLf & _
            "   AND Pedido_Id =" & Me.Pedido.Codigo & "" & vbCrLf & _
            "   AND TipoOrigem_Id = '" & Me.TipoOrigem & "'" & vbCrLf & _
            "   AND Origem_Id = '" & Me.CodigoOrigem & "'" & vbCrLf & _
            "   AND EndOrigem_Id = " & Me.EnderecoOrigem & vbCrLf & _
            "   AND TipoDestino_Id = '" & Me.TipoDestino & "'" & vbCrLf & _
            "   AND Destino_Id = '" & Me.CodigoDestino & "'" & vbCrLf & _
            "   AND EndDestino_Id = " & Me.EnderecoDestino & vbCrLf & _
            "   AND ViaDeTransporte_Id = " & Me.CodigoViaDeTransporte & vbCrLf
    End Sub

#End Region
End Class
