<Serializable()> _
Public Class ListBancosXContasXCancelamentoCheque
    Inherits List(Of BancosXContasXCancelamentoCheque)

    Public Sub New()

    End Sub

    Public Sub New(ByVal ObjBXC As [Lib].Negocio.BancosXContas, Optional ByVal NumCheque As Integer = 0)
        Dim sql As String = "   SELECT  BXCXC.Empresa_Id, BXCXC.EndEmpresa_Id, BXCXC.Banco_Id, B.Descricao AS NomeBanco, BXCXC.Agencia_Id, BXCXC.DigitoAgencia_Id, BXCXC.Conta_Id, BXCXC.DigitoConta_Id,      " & vbCrLf & _
                            "   		BXCXC.NumCheque_Id, BXCXC.DataCancelamento, BXCXC.Pedido, BXCXC.Titulo, BXCXC.Observacao, BXCXC.UsuarioCancelamento, BXCXC.TipoCancelamento " & vbCrLf & _
                            "   FROM    BancosXContasXCancelamentoCheque BXCXC                                                                  " & vbCrLf & _
                            "       INNER JOIN Bancos B" & vbCrLf & _
                            "           ON B.Banco_Id = BXCXC.BANCO_ID" & vbCrLf & _
                            "   WHERE	BXCXC.Empresa_Id			= '" & ObjBXC.CodigoEmpresa & "'" & vbCrLf & _
                            "   	AND	BXCXC.EndEmpresa_Id		=  " & ObjBXC.EndEmpresa & vbCrLf & _
                            "   	AND	BXCXC.Banco_Id			= " & ObjBXC.CodigoBanco & vbCrLf & _
                            "   	AND	BXCXC.Agencia_Id			= '" & ObjBXC.Agencia & "'" & vbCrLf & _
                            "   	AND	BXCXC.DigitoAgencia_Id	= '" & ObjBXC.DigitoAgencia & "'" & vbCrLf & _
                            "   	AND	BXCXC.Conta_Id			= '" & ObjBXC.Conta & "'" & vbCrLf & _
                            "   	AND	BXCXC.DigitoConta_Id		= '" & ObjBXC.DigitoConta & "'" & vbCrLf
        If NumCheque <> 0 Then
            sql &= "AND BXCXC.numCheque_Id > " & NumCheque & vbCrLf
        End If

        sql &= "   order by  BXCXC.numCheque_Id desc" & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet

        ds = Banco.ConsultaDataSet(sql, "BancosXContas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim BxCxC As New BancosXContasXCancelamentoCheque
            BxCxC.CodigoEmpresa = row("Empresa_Id")
            BxCxC.EndEmpresa = row("EndEmpresa_Id")
            BxCxC.CodigoBanco = row("Banco_Id")
            BxCxC.NomeBanco = row("NomeBanco")
            BxCxC.CodigoAgencia = row("Agencia_Id")
            BxCxC.CodigoDigitoAgencia = row("DigitoAgencia_Id")
            BxCxC.CodigoConta = row("Conta_Id")
            BxCxC.CodigoDigitoConta = row("DigitoConta_Id")
            BxCxC.NumCheque = row("NumCheque_Id")
            BxCxC.DataCancelamento = row("DataCancelamento")
            BxCxC.Pedido = row("Pedido")
            BxCxC.Titulo = row("Titulo")
            BxCxC.Observacao = row("Observacao")
            BxCxC.UsuarioCancelamento = row("UsuarioCancelamento")
            BxCxC.TipoCancelamento = row("TipoCancelamento")
            Me.Add(BxCxC)
        Next
    End Sub
End Class

<Serializable()> _
Public Class BancosXContasXCancelamentoCheque
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _CodigoBanco As Integer
    Private _NomeBanco As String
    Private _CodigoAgencia As String
    Private _CodigoDigitoAgencia As String
    Private _CodigoConta As String
    Private _CodigoDigitoConta As String
    Private _NumCheque As Integer
    Private _DataCancelamento As DateTime
    Private _UsuarioCancelamento As String
    Private _TipoCancelamento As String
    Private _Pedido As Integer
    Private _Titulo As Integer
    Private _Observacao As String
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

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property CodigoBanco() As Integer
        Get
            Return _CodigoBanco
        End Get
        Set(ByVal value As Integer)
            _CodigoBanco = value
        End Set
    End Property

    Public Property NomeBanco() As String
        Get
            Return _NomeBanco
        End Get
        Set(ByVal value As String)
            _NomeBanco = value
        End Set
    End Property

    Public Property CodigoAgencia() As String
        Get
            Return _CodigoAgencia
        End Get
        Set(ByVal value As String)
            _CodigoAgencia = value
        End Set
    End Property

    Public Property CodigoDigitoAgencia() As String
        Get
            Return _CodigoDigitoAgencia
        End Get
        Set(ByVal value As String)
            _CodigoDigitoAgencia = value
        End Set
    End Property

    Public Property CodigoConta() As String
        Get
            Return _CodigoConta
        End Get
        Set(ByVal value As String)
            _CodigoConta = value
        End Set
    End Property

    Public Property CodigoDigitoConta() As String
        Get
            Return _CodigoDigitoConta
        End Get
        Set(ByVal value As String)
            _CodigoDigitoConta = value
        End Set
    End Property

    Public Property NumCheque() As Integer
        Get
            Return _NumCheque
        End Get
        Set(ByVal value As Integer)
            _NumCheque = value
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

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property TipoCancelamento() As String
        Get
            Return _TipoCancelamento
        End Get
        Set(ByVal value As String)
            _TipoCancelamento = value
        End Set
    End Property

    Public Property Pedido() As Integer
        Get
            Return _Pedido
        End Get
        Set(ByVal value As Integer)
            _Pedido = value
        End Set
    End Property

    Public Property Titulo() As Integer
        Get
            Return _Titulo
        End Get
        Set(ByVal value As Integer)
            _Titulo = value
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

#Region "Metodo"
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
                Sql = " INSERT INTO BancosXContasXCancelamentoCheque (Empresa_id, EndEmpresa_id, Banco_id, Agencia_id, DigitoAgencia_id, Conta_id, DigitoConta_id, NumCheque_id, DataCancelamento, Pedido, Titulo, Observacao, UsuarioCancelamento, TipoCancelamento) " & vbCrLf & _
                      " VALUES ('" & _CodigoEmpresa & "'," & _EndEmpresa & "," & _CodigoBanco & ",'" & _CodigoAgencia & "','" & _CodigoDigitoAgencia & "','" & _CodigoConta & "'," & vbCrLf & _
                      "         '" & _CodigoDigitoConta & "'," & _NumCheque & ",'" & _DataCancelamento.ToString("yyyy-MM-dd") & "'," & _Pedido & ", " & vbCrLf & _
                      "          " & _Titulo & ", '" & _Observacao & "', '" & _UsuarioCancelamento & "', '" & TipoCancelamento & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE BancosXContasXCancelamentoCheque " & vbCrLf & _
                      "    SET DataCancelamento     = '" & _DataCancelamento.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "        Pedido               =  " & _Pedido & "," & vbCrLf & _
                      "        Titulo               =  " & _Titulo & "," & vbCrLf & _
                      "        Observacao           = '" & _Observacao & "'," & vbCrLf & _
                      "        UsuarioCancelamento  = '" & _UsuarioCancelamento & "'," & vbCrLf & _
                      "        TipoCancelamento     = '" & _TipoCancelamento & "'," & vbCrLf & _
                      "  WHERE Empresa_id           = '" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    AND EndEmpresa_id        =  " & _EndEmpresa & vbCrLf & _
                      "    AND Banco_id             =  " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia_id           = '" & _CodigoAgencia & "'" & vbCrLf & _
                      "    AND DigitoAgencia _id    = '" & _CodigoDigitoAgencia & "'" & vbCrLf & _
                      "    AND Conta_id             = '" & _CodigoConta & "'" & vbCrLf & _
                      "    AND DigitoConta_id       = '" & _CodigoDigitoConta & "'" & vbCrLf & _
                      "    AND NumCheque_id         =  " & _NumCheque & "" & vbCrLf

                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE BancosXContasXCancelamentoCheque " & vbCrLf & _
                     "  WHERE  Empresa       _id       = '" & _CodigoEmpresa & "'" & vbCrLf & _
                      "    AND EndEmpresa    _id       =  " & _EndEmpresa & vbCrLf & _
                      "    AND Banco         _id       =  " & _CodigoBanco & vbCrLf & _
                      "    AND Agencia       _id       = '" & _CodigoAgencia & "'" & vbCrLf & _
                      "    AND DigitoAgencia _id       = '" & _CodigoDigitoAgencia & "'" & vbCrLf & _
                      "    AND Conta         _id       = '" & _CodigoConta & "'" & vbCrLf & _
                      "    AND DigitoConta   _id       = '" & _CodigoDigitoConta & "'" & vbCrLf & _
                      "    AND NumCheque     _id       =  " & _NumCheque & " " & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class
