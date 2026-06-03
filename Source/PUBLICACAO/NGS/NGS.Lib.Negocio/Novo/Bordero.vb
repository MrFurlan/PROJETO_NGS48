Imports System.Data
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListBordero
        Inherits List(Of Bordero)

#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(pCodigoEmpresa As String, pEndEmpresa As Integer, pWhere As String)
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, CarteiraDoTitulo, TituloBordero, DataEnvio, situacao" & vbCrLf & _
                  "  From Bordero" & vbCrLf

            If pWhere.Length > 0 Then sql &= " Where " & pWhere

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Bordero")

            For Each row In ds.Tables(0).Rows
                Dim B As New Bordero()
                B.CodigoEmpresa = row("Empresa_Id")
                B.EnderecoEmpresa = row("EndEmpresa_Id")
                B.CodigoBordero = row("Bordero_Id")
                B.CodigoCarteiraDoTitulo = row("CarteiraDoTitulo")
                B.CodigoTituloBordero = row("TituloBordero")
                B.DataEnvio = row("DataEnvio")
                B.CodigoSituacao = row("Situacao")
                Me.Add(B)
            Next
        End Sub
#End Region

    End Class

    <Serializable()> _
    Public Class Bordero
#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(pCodigoEmpresa As String, pEndEmpresa As Integer, pCodigoBordero As Integer)
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, CarteiraDoTitulo, TituloBordero, DataEnvio, situacao" & vbCrLf & _
                  "  From Bordero" & vbCrLf & _
                  " Where Bordero_Id = " & pCodigoBordero
            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Bordero")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row = ds.Tables(0).Rows(0)
            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.CodigoBordero = row("Bordero_Id")
            Me.CodigoCarteiraDoTitulo = row("CarteiraDoTitulo")
            Me.CodigoTituloBordero = row("TituloBordero")
            Me.DataEnvio = row("DataEnvio")
            Me.CodigoSituacao = row("Situacao")

        End Sub

        Public Sub New(pCodigoTitulo As Integer)
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, CarteiraDoTitulo, TituloBordero, DataEnvio, situacao" & vbCrLf & _
                  "  From Bordero" & vbCrLf & _
                  " Where TituloBordero = " & pCodigoTitulo
            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Bordero")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row = ds.Tables(0).Rows(0)
            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.CodigoBordero = row("Bordero_Id")
            Me.CodigoCarteiraDoTitulo = row("CarteiraDoTitulo")
            Me.CodigoTituloBordero = row("TituloBordero")
            Me.DataEnvio = row("DataEnvio")
            Me.CodigoSituacao = row("Situacao")

        End Sub
#End Region

#Region "Fields"
        Private _IUD As String = ""
        Private _CodigoEmpresa As String = ""
        Private _EndEmpresa As Integer
        Private _Empresa As Cliente
        Private _CodigoBordero As Integer
        Private _CodigoCarteiraDoTitulo As String = ""
        Private _CarteiraDoTitulo As CarteiraDoTitulo 'Carteira do titulo
        Private _CodigoTituloBordero As Integer
        Private _TituloDoBordero As Novo.TituloNovo
        Private _DataEnvio As Date
        Private _CodigoSituacao As Integer
        Private _JurosTaxa As Decimal
        Private _TitulosDoBordero As Novo.ListBorderoXTitulo
        Private _TitulosRecomprados As Novo.ListBorderoXTitulo
        Private _Contrato As String = ""
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
                _Empresa = Nothing
            End Set
        End Property
        Public Property EnderecoEmpresa() As Integer
            Get
                Return _EndEmpresa
            End Get
            Set(ByVal value As Integer)
                _EndEmpresa = value
                _Empresa = Nothing
            End Set
        End Property
        Public Property Empresa() As Cliente
            Get
                If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
                Return _Empresa
            End Get
            Set(ByVal value As Cliente)
                _Empresa = value
            End Set
        End Property
        Public Property CodigoBordero As Integer
            Get
                Return _CodigoBordero
            End Get
            Set(value As Integer)
                _CodigoBordero = value
            End Set
        End Property
        Public Property CodigoCarteiraDoTitulo() As String
            Get
                Return _CodigoCarteiraDoTitulo
            End Get
            Set(ByVal value As String)
                _CodigoCarteiraDoTitulo = value
                _CarteiraDoTitulo = Nothing
            End Set
        End Property
        Public Property CarteiraDoTitulo() As CarteiraDoTitulo
            Get
                If _CarteiraDoTitulo Is Nothing And _CodigoCarteiraDoTitulo > 0 Then _CarteiraDoTitulo = New CarteiraDoTitulo(_CodigoCarteiraDoTitulo)
                Return _CarteiraDoTitulo
            End Get
            Set(ByVal value As CarteiraDoTitulo)
                _CarteiraDoTitulo = value
            End Set
        End Property
        Public Property CodigoTituloBordero() As Integer
            Get
                Return _CodigoTituloBordero
            End Get
            Set(ByVal value As Integer)
                _CodigoTituloBordero = value
            End Set
        End Property
        Public Property TituloDoBordero As Novo.TituloNovo
            Get
                If _TituloDoBordero Is Nothing And _CodigoTituloBordero > 0 Then _TituloDoBordero = New Novo.TituloNovo(_CodigoTituloBordero)
                Return _TituloDoBordero
            End Get
            Set(value As Novo.TituloNovo)
                _TituloDoBordero = value
            End Set
        End Property
        Public Property DataEnvio() As Date
            Get
                Return _DataEnvio
            End Get
            Set(ByVal value As Date)
                _DataEnvio = value
            End Set
        End Property
        Public Property JurosTaxa As Decimal
            Get
                Return _JurosTaxa
            End Get
            Set(value As Decimal)
                _JurosTaxa = value
            End Set
        End Property
        Public Property CodigoSituacao As Integer
            Get
                Return _CodigoSituacao
            End Get
            Set(value As Integer)
                _CodigoSituacao = value
            End Set
        End Property
        Public Property TitulosDoBordero As Novo.ListBorderoXTitulo
            Get
                If _TitulosDoBordero Is Nothing And _CodigoBordero > 0 Then
                    _TitulosDoBordero = New ListBorderoXTitulo(Me)
                ElseIf _TitulosDoBordero Is Nothing And _CodigoBordero = 0 Then
                    _TitulosDoBordero = New ListBorderoXTitulo()
                    _TitulosDoBordero.Bordero = Me
                End If

                Return _TitulosDoBordero
            End Get
            Set(value As Novo.ListBorderoXTitulo)
                _TitulosDoBordero = value
            End Set
        End Property
        Public Property TitulosRecomprados As Novo.ListBorderoXTitulo
            Get
                If Me.CodigoBordero = 0 Then
                    _TitulosRecomprados = New ListBorderoXTitulo()
                    _TitulosRecomprados.Bordero = Me
                Else
                    _TitulosRecomprados = New ListBorderoXTitulo(Me.CodigoBordero, True)
                End If
                Return _TitulosRecomprados
            End Get
            Set(value As Novo.ListBorderoXTitulo)
                _TitulosRecomprados = value
            End Set
        End Property
        Public Property Contrato() As String
            Get
                Return _Contrato
            End Get
            Set(ByVal value As String)
                _Contrato = value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList
            SalvarSql(sqls)
            Return Banco.GravaBanco(sqls)
        End Function

        Public Sub SalvarSql(ByRef sqls As ArrayList)
            Dim strSQL As String
            Dim ObjBanco As New AcessaBanco

            Select Case Me.IUD
                Case "I"
                    Dim numBordero As New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 4)
                    _CodigoBordero = numBordero.Sequencia + 1
                    sqls.Add(numBordero.IncrementarNumeradorSql())

                    strSQL = "INSERT INTO Bordero(Empresa_Id, EndEmpresa_Id, Bordero_Id, CarteiraDoTitulo, TituloBordero, DataEnvio, Situacao, JurosTaxa, Contrato) " & vbCrLf & _
                             " VALUES ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.CodigoBordero & "," & Me.CodigoCarteiraDoTitulo & "," & Me.CodigoTituloBordero & ",'" & Me.DataEnvio.ToString("yyyy-MM-dd") & "'," & Me.CodigoSituacao & "," & Str(Math.Round(Me.JurosTaxa, 2)) & ",'" & Me.Contrato & "')"
                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "U"
                    strSQL = " UPDATE Bordero Set " & vbCrLf & _
                             "   CarteiraDoTitulo      = " & Me.CodigoCarteiraDoTitulo & vbCrLf & _
                             "  ,TituloBordero         = " & Me.CodigoTituloBordero & vbCrLf & _
                             "  ,DataEnvio             ='" & Me.DataEnvio.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                             "  ,Situacao              = " & Me.CodigoSituacao & vbCrLf & _
                             "  ,JurosTaxa             = " & Me.JurosTaxa & vbCrLf & _
                             "  ,Contrato              ='" & Me.Contrato & "'" & vbCrLf & _
                             " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                             "   And EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                             "   And Bordero_id    = " & Me.CodigoBordero & vbCrLf
                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "D"
                    SalvarTabelasRelacionadasSql(sqls)
                    strSQL = "Update Bordero set " & vbCrLf & _
                             "  Situacao = 2 " & vbCrLf & _
                             " WHERE Bordero_id = " & Me.CodigoBordero
                    sqls.Add(strSQL)
                Case Else
                    SalvarTabelasRelacionadasSql(sqls)
            End Select
        End Sub

        Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)

            For Each BxT In Me.TitulosDoBordero
                Select Case Me.IUD
                    Case "I"
                        BxT.IUD = "I"
                        BxT.CodigoEmpresa = Me.CodigoEmpresa
                        BxT.EnderecoEmpresa = Me.EnderecoEmpresa
                        BxT.CodigoBordero = Me.CodigoBordero
                        BxT.Bordero = Me
                        BxT.SalvarSql(Sqls)
                    Case "U"
                        BxT.IUD = "U"
                        BxT.SalvarSql(Sqls)
                    Case "D"
                        BxT.IUD = "D"
                        BxT.SalvarSql(Sqls)
                End Select

            Next

        End Sub

        Function ImprimirBordero() As DataSet
            Dim Banco As New AcessaBanco
            Dim ds As DataSet
            Dim sql As String

            sql = " SELECT  " & vbCrLf & _
                 " T.Empresa, " & vbCrLf & _
                 " EndEmpresa, " & vbCrLf & _
                 " C.Nome, " & vbCrLf & _
                 " C.Endereco, " & vbCrLf & _
                 " C.Numero, " & vbCrLf & _
                 " C.Complemento, " & vbCrLf & _
                 " C.Cidade, " & vbCrLf & _
                 " C.Estado, " & vbCrLf & _
                 " B.DataEnvio, " & vbCrLf & _
                 " B.JurosTaxa, " & vbCrLf & _
                 " TC.Conta_Id, " & vbCrLf & _
                 " PDC.Titulo, " & vbCrLf & _
                 " TC.ValorOficial, " & vbCrLf & _
                 " TC.ValorMoeda, " & vbCrLf & _
                 " FCT.Titulo AS TituloBordero, " & vbCrLf & _
                 " BCO.Titulo AS TituloBanco, " & vbCrLf & _
                    " TP.Descricao " & vbCrLf & _
                    " FROM Bordero B " & vbCrLf & _
                    " INNER JOIN Titulos T  " & vbCrLf & _
                 " ON B.TituloBordero = T.Titulo_Id " & vbCrLf & _
                    " INNER JOIN TitulosxContaContabil TC  " & vbCrLf & _
                 " ON T.Titulo_Id = TC.Titulo_Id " & vbCrLf & _
                    " INNER JOIN Clientes C " & vbCrLf & _
                 " ON T.Empresa = C.Cliente_Id " & vbCrLf & _
                 " AND T.EndEmpresa = C.Endereco_Id " & vbCrLf & _
                    " INNER JOIN PlanoDeContas AS PDC " & vbCrLf & _
                 " ON TC.Conta_Id = PDC.Conta_Id " & vbCrLf & _
                    " INNER JOIN PlanoDeContas AS FCT " & vbCrLf & _
                 " ON T.ContaContabilCliFor = FCT.Conta_Id " & vbCrLf & _
                    " INNER JOIN PlanoDeContas AS BCO " & vbCrLf & _
                 " ON T.ContaContabilRecPag = BCO.Conta_Id " & vbCrLf & _
                    " INNER JOIN TiposDePagamentos TP  " & vbCrLf & _
                 " ON T.TipoPagto = TP.TipoDePagamento_Id " & vbCrLf & _
                    " WHERE B.TituloBordero = 564659 " & vbCrLf

            ds = Banco.ConsultaDataSet(sql, "Bordero")

            sql = " SELECT  " & vbCrLf & _
                    " '402020115' AS Conta_Id, " & vbCrLf & _
                    " T.Titulo_Id, " & vbCrLf & _
                    " T.Reprogramacao, " & vbCrLf & _
                    " T.Empresa, " & vbCrLf & _
                    " T.EndEmpresa, " & vbCrLf & _
                    "C.Nome, " & vbCrLf & _
                    " (SELECT " & vbCrLf & _
                    " 	SUM(case " & vbCrLf & _
                    "	  when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag in ('C','P') and Tc.DC_Id = 'C' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                    "	  when Tc.Conta_Id  = Tp.ContaContabilRecPag and Tp.RecPag = 'R'        and Tc.DC_Id = 'D' then case when M.Classificacao ='O' then Tc.ValorOficial else Tc.ValorMoeda end " & vbCrLf & _
                    "	  else 0 " & vbCrLf & _
                    "        End " & vbCrLf & _
                    "	) " & vbCrLf & _
                    "	FROM Titulos Tp " & vbCrLf & _
                    "	inner Join TitulosxContaContabil Tc " & vbCrLf & _
                    "	on Tc.Titulo_Id = Tp.Titulo_Id " & vbCrLf & _
                    "	INNER JOIN Moedas M " & vbCrLf & _
                    "	on M.Moeda_id = Tp.Moeda " & vbCrLf & _
                    "        WHERE(TP.Titulo_Id = T.Titulo_Id) " & vbCrLf & _
                    "	Group by Tp.Titulo_Id) AS ValorLiquido, " & vbCrLf & _
                    "	'' AS Float, " & vbCrLf & _
                    "	'' AS Prazo, " & vbCrLf & _
                    "	'' AS Desagio, " & vbCrLf & _
                    "	'' AS Liquido " & vbCrLf & _
                    " FROM Bordero B " & vbCrLf & _
                    " INNER JOIN BorderoXTitulo BT " & vbCrLf & _
                    "	ON B.Bordero_Id = BT.Bordero_Id " & vbCrLf & _
                    " INNER JOIN Titulos T " & vbCrLf & _
                    "	ON BT.Titulo_Id = T.Titulo_Id " & vbCrLf & _
                    " INNER JOIN Clientes C " & vbCrLf & _
                    "	ON T.CliFor = C.Cliente_Id " & vbCrLf & _
                    "	AND T.EnderecoCliFor = C.Endereco_Id " & vbCrLf & _
                    "            WHERE B.TituloBordero = 564659  " & vbCrLf

            ds.Merge(Banco.ConsultaDataSet(sql, "Titulo"))

            For Each row As DataRow In ds.Tables("Titulo").Rows
                row("Float") = "4"
                row("Prazo") = Convert.ToDateTime(row("Reprogramacao")).Subtract(Convert.ToDateTime(ds.Tables("Bordero").Rows(0).ItemArray(8).ToString)).Days
                row("Desagio") = (row("ValorLiquido") * (Convert.ToDecimal(ds.Tables("Bordero").Rows(0).ItemArray(9).ToString) / 100) * (row("Prazo") / 30))
                row("Liquido") = row("ValorLiquido") - row("Desagio")
            Next

            Return ds
        End Function
#End Region

    End Class
End Namespace
