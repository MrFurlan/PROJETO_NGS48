Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Configuration

<Serializable()> _
Public Class ListFaturaDeFrete
    Inherits List(Of FaturaDeFrete)
    Implements IBaseEntity

    Public ReadOnly Property FinanceiroNovo As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
        End Get
    End Property

    Public Sub New()

    End Sub

    Public Sub New(ByVal nf As NotaFiscal, ByVal lst As List(Of FaturaDeFreteXItens))
        If lst IsNot Nothing AndAlso lst.Count > 0 Then
            Dim sql As String = ""
            Dim db As New AcessaBanco
            Dim ids As String = lst.Select(Function(s) s.CodigoFatura).Aggregate(Function(a, b) Convert.ToString(a) & ", " & Convert.ToString(b))

            sql &= " SELECT Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_id, Movimento, ValorDaFatura " & vbCrLf &
                   "   FROM FaturasDeFretes " &
                    " WHERE Empresa_Id       ='" & nf.CodigoEmpresa & "' " &
                    "   AND EndEmpresa_Id    ='" & nf.EnderecoEmpresa & "' " &
                    "   AND Conveniado_Id    ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, nf.CodigoCliente, nf.CodigoTomador) & "' " &
                    "   AND EndConveniado_Id ='" & IIf(nf.EntradaSaida = eEntradaSaida.Entrada, nf.EnderecoCliente, nf.EnderecoTomador) & "' " &
                    "   AND Fatura_Id in (" & ids & ") "

            Dim ds As DataSet = db.ConsultaDataSet(sql, "FaturasDeFretes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim ff = New FaturaDeFrete
                ff.CodigoEmpresa = row("Empresa_Id")
                ff.EnderecoEmpresa = row("EndEmpresa_Id")
                ff.CodigoConveniado = row("Conveniado_id")
                ff.EnderecoConveniado = row("EndConveniado_id")
                ff.CodigoFatura = row("Fatura_Id")
                ff.Movimento = row("Movimento")
                ff.ValorDaFatura = row("ValorDaFatura")
                Me.Add(ff)
            Next
        End If
    End Sub

    Public Sub New(ByVal fatura As FaturaDeFrete, ByVal usarPeriodo As Boolean, ByVal clienteFatura As Cliente, Optional ByVal dataInicial As String = "", Optional ByVal dataFinal As String = "", Optional ByVal provisao As eProvisao = 0)
        Dim sql As String
        Dim db As New AcessaBanco

        sql = "SELECT ISNULL(FXI.Nota_Id, 0) AS Nota_Id, FF.Empresa_Id, FF.EndEmpresa_Id, FF.Conveniado_Id, FF.EndConveniado_Id, " & vbCrLf &
              "       FF.Fatura_Id, FF.Movimento, FF.ValorDaFatura, " & vbCrLf &
              "       Empresa.Nome AS EmpresaNome, Empresa.Reduzido AS EmpresaReduzido, " & vbCrLf &
              "       Conveniado.Nome AS ConveniadoNome, Conveniado.Reduzido AS ConveniadoReduzido, ISNULL(FXI.Encargo_Id, '') AS Encargo_Id,  sbt.PagarReceber, ISNULL(FF.RegistroMestre, 0) AS RegistroMestre, ISNULL(FF.LancamentoManual, 0) AS LancamentoManual " & vbCrLf &
              " FROM FaturasDeFretes AS FF " & vbCrLf &
              " LEFT JOIN FaturasDeFretesXItens FXI " & vbCrLf &
              "    ON FXI.Empresa_Id    = ff.Empresa_Id " & vbCrLf &
              "   AND FXI.EndEmpresa_Id = ff.EndEmpresa_Id " & vbCrLf &
              "   AND FXI.Fatura_Id     = ff.Fatura_Id " & vbCrLf &
              " INNER JOIN Clientes AS Empresa " & vbCrLf &
              "    ON FF.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf &
              "   AND FF.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
              " INNER JOIN Clientes AS Conveniado " & vbCrLf &
              "    ON FF.Conveniado_Id    = Conveniado.Cliente_Id " & vbCrLf &
              "   AND FF.EndConveniado_Id = Conveniado.Endereco_Id " & vbCrLf

        sql &= " INNER JOIN (" & vbCrLf &
               "                SELECT distinct fxt.Empresa_Id, fxt.EndEmpresa_Id, fxt.Conveniado_Id, fxt.EndConveniado_Id, fxt.Fatura_Id, 'P' AS PagarReceber" & vbCrLf &
               "	            FROM FaturaDeFreteXTitulo fxt" & vbCrLf &
               "                INNER JOIN " & IIf(FinanceiroNovo, "Titulos", "ContasAPagar") & " AS CP " & vbCrLf &
               "				   ON fxt.Titulo_Id = CP.Registro_Id " & vbCrLf &
               "                WHERE 1 = 1" & vbCrLf

        If usarPeriodo Then
            If (Not String.IsNullOrWhiteSpace(dataInicial) AndAlso Not String.IsNullOrWhiteSpace(dataFinal)) Then
                sql &= "				  and CP.Prorrogacao Between '" & dataInicial.ToSqlDate() & "' and '" & dataFinal.ToSqlDate() & "' " & vbCrLf
            End If
        End If

        If provisao > 0 Then
            sql &= "				  AND CP.Provisao =  " & provisao & vbCrLf
        End If

        sql &= " UNION " & vbCrLf &
                   "                SELECT distinct fxt.Empresa_Id, fxt.EndEmpresa_Id, fxt.Conveniado_Id, fxt.EndConveniado_Id, fxt.Fatura_Id, 'R' AS PagarReceber" & vbCrLf &
                   "	            FROM FaturaDeFreteXTitulo fxt" & vbCrLf &
                   "                INNER JOIN " & IIf(FinanceiroNovo, "Titulos", "ContasAReceber") & " AS CR " & vbCrLf &
                   "				   ON fxt.Titulo_Id = CR.Registro_Id " & vbCrLf &
                   "                WHERE 1 = 1" & vbCrLf

        If usarPeriodo Then
            If (Not String.IsNullOrWhiteSpace(dataInicial) AndAlso Not String.IsNullOrWhiteSpace(dataFinal)) Then
                sql &= "				  and CR.Prorrogacao Between '" & dataInicial.ToSqlDate() & "' and '" & dataFinal.ToSqlDate() & "' " & vbCrLf
            End If
        End If

        If provisao > 0 Then
            sql &= "				  AND CR.Provisao =  " & provisao & vbCrLf
        End If

        If provisao > 0 Then
            sql &= "            ) sbt" & vbCrLf &
                       "	 on sbt.Empresa_Id       = FF.Empresa_Id" & vbCrLf &
                       "	and sbt.EndEmpresa_Id    = FF.EndEmpresa_Id" & vbCrLf &
                       "	and sbt.Conveniado_Id    = FF.Conveniado_Id" & vbCrLf &
                       "	and sbt.EndConveniado_Id = FF.EndConveniado_Id" & vbCrLf &
                       "	and sbt.Fatura_Id        = FF.Fatura_Id" & vbCrLf
        End If

        sql &= " WHERE 1 = 1 " & vbCrLf

        If Not fatura.CodigoFatura Is Nothing AndAlso Not String.IsNullOrWhiteSpace(fatura.CodigoFatura) Then
            sql &= "    AND (FF.Fatura_Id = '" & fatura.CodigoFatura & "')" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(fatura.CodigoEmpresa) Then
            sql &= "    AND FF.Empresa_Id    = '" & Funcoes.EliminarCaracteresEspeciais(fatura.CodigoEmpresa) & "'" & vbCrLf &
                   "    AND FF.EndEmpresa_Id = " & fatura.EnderecoEmpresa & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(fatura.CodigoConveniado) Then
            sql &= "    AND FF.Conveniado_Id = '" & Funcoes.EliminarCaracteresEspeciais(fatura.CodigoConveniado) & "'" & vbCrLf &
                   "    AND FF.EndConveniado_id = " & fatura.EnderecoConveniado & vbCrLf
        End If

        If Not clienteFatura Is Nothing AndAlso Not String.IsNullOrWhiteSpace(clienteFatura.Codigo) Then
            sql &= "    AND FXI.Cliente_Id = '" & Funcoes.EliminarCaracteresEspeciais(clienteFatura.Codigo) & "'" & vbCrLf &
                   "    AND FXI.EndCliente_Id = " & clienteFatura.CodigoEndereco & vbCrLf
        End If

        sql &= "  ORDER BY FF.Movimento DESC, FF.Fatura_Id, FXI.Nota_Id DESC, FXI.Encargo_Id "

        Dim ds As DataSet = db.ConsultaDataSet(sql, "FaturasDeFretes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing And ds.Tables.Count > 0 Then

            For Each row As DataRow In ds.Tables(0).Rows

                Dim ff = New FaturaDeFrete
                ff.CodigoNota = row("Nota_Id")
                ff.CodigoEmpresa = row("Empresa_Id")
                ff.EnderecoEmpresa = row("EndEmpresa_Id")
                ff.EmpresaReduzido = row("EmpresaReduzido")
                ff.EmpresaNome = row("EmpresaNome")
                ff.CodigoConveniado = row("Conveniado_id")
                ff.EnderecoConveniado = row("EndConveniado_id")
                ff.ConveniadoReduzido = row("ConveniadoReduzido")
                ff.ConveniadoNome = row("ConveniadoNome")
                ff.CodigoFatura = row("Fatura_Id")
                ff.Movimento = row("Movimento")
                ff.ValorDaFatura = row("ValorDaFatura")
                ff.Encargo = row("Encargo_Id")
                ff.PagarReceber = row("PagarReceber")
                ff.RegistroMestre = row("RegistroMestre")
                ff.LancamentoManual = row("LancamentoManual")

                Me.Add(ff)

            Next
        End If
    End Sub

End Class

<Serializable()> _
Public Class FaturaDeFrete
    Implements IBaseEntity

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal PFrete As FaturaDeFrete)
        Dim Sql As String
        Sql = " SELECT FF.Empresa_Id, FF.EndEmpresa_Id, FF.Conveniado_Id, FF.EndConveniado_Id, " & vbCrLf & _
              "        FF.Fatura_Id, FF.Movimento, FF.ValorDaFatura, " & vbCrLf & _
              "        Empresa.Nome AS EmpresaNome, Empresa.Reduzido AS EmpresaReduzido, " & vbCrLf & _
              "        conveniado.Nome AS ConveniadoNome, conveniado.Reduzido AS ConveniadoReduzido" & vbCrLf & _
              "   FROM FaturasDeFretes AS FF " & vbCrLf & _
              "  INNER JOIN Clientes AS Empresa " & vbCrLf & _
              "     ON FF.Empresa_Id    = Empresa.Cliente_Id " & vbCrLf & _
              "    AND FF.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf & _
              "  INNER JOIN Clientes AS conveniado " & vbCrLf & _
              "     ON FF.Conveniado_Id    = conveniado.Cliente_Id " & vbCrLf & _
              "    AND FF.EndConveniado_Id = conveniado.Endereco_Id " & vbCrLf & _
              "  WHERE FF.Empresa_id       ='" & Funcoes.EliminarCaracteresEspeciais(PFrete.CodigoEmpresa) & "'" & vbCrLf & _
              "    AND FF.EndEmpresa_id    = " & PFrete.EnderecoEmpresa & vbCrLf & _
              "    AND FF.Fatura_Id        ='" & PFrete.CodigoFatura & "'" & vbCrLf & _
              "    AND FF.Conveniado_Id    ='" & PFrete.CodigoConveniado & "'" & vbCrLf & _
              "    AND FF.EndConveniado_Id = " & PFrete.EnderecoConveniado & vbCrLf


        Dim Banco As New AcessaBanco
        Dim DsFaturas As DataSet = Banco.ConsultaDataSet(Sql, "FF")

        For Each dr As DataRow In DsFaturas.Tables(0).Rows
            Me.CodigoEmpresa = dr("Empresa_Id")
            Me.EnderecoEmpresa = dr("EndEmpresa_Id")
            Me.EmpresaReduzido = dr("EmpresaReduzido")
            Me.EmpresaNome = dr("EmpresaNome")
            Me.CodigoConveniado = dr("Conveniado_id")
            Me.EnderecoConveniado = dr("EndConveniado_id")
            Me.ConveniadoReduzido = dr("ConveniadoReduzido")
            Me.ConveniadoNome = dr("ConveniadoNome")
            Me.CodigoFatura = dr("Fatura_Id")
            Me.Movimento = dr("Movimento")
            Me.ValorDaFatura = dr("ValorDaFatura")
        Next
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String

    Private _CodigoNota As String
    Private _CodigoEmpresa As String
    Private _EndEmpresa As Integer
    Private _EmpresaReduzido As String
    Private _EmpresaNome As String

    Private _CodigoConveniado As String
    Private _ConveniadoEnd As Integer
    Private _ConveniadoReduzido As String
    Private _ConveniadoNome As String
    Private _Conveniado As Negocio.Cliente

    Private _CodigoFatura As String
    Private _Movimento As Date

    Private _TituloFatura As Titulo
    Private _TituloNovoFatura As Novo.TituloNovo
    Private _Vencimento As Date
    Private _Titulos As Novo.ListTituloNovo

    Private _ValorDaFatura As Decimal
    Private _ValorLancado As Decimal


    Private _Itens As ListFaturasDeFretesXItens
    Private _ListTituloFatura As ListFaturaDeFretexTitulo
    Private _Encargo As String
    Private _PagarReceber As String
    Private _RegistroMestre As Integer
    Private _LancamentoManual As Integer

#End Region

#Region "Propriedades"

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoNota() As String
        Get
            Return _CodigoNota
        End Get
        Set(ByVal value As String)
            _CodigoNota = value
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

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
        End Set
    End Property

    Public Property EmpresaReduzido() As String
        Get
            Return _EmpresaReduzido
        End Get
        Set(ByVal value As String)
            _EmpresaReduzido = value
        End Set
    End Property

    Public Property EmpresaNome() As String
        Get
            Return _EmpresaNome
        End Get
        Set(ByVal value As String)
            _EmpresaNome = value
        End Set
    End Property

    Public Property CodigoConveniado() As String
        Get
            Return _CodigoConveniado
        End Get
        Set(ByVal value As String)
            _CodigoConveniado = value
        End Set
    End Property

    Public Property EnderecoConveniado() As Integer
        Get
            Return _ConveniadoEnd
        End Get
        Set(ByVal value As Integer)
            _ConveniadoEnd = value
        End Set
    End Property

    Public Property ConveniadoReduzido() As String
        Get
            Return _ConveniadoReduzido
        End Get
        Set(ByVal value As String)
            _ConveniadoReduzido = value
        End Set
    End Property

    Public Property ConveniadoNome() As String
        Get
            Return _ConveniadoNome
        End Get
        Set(ByVal value As String)
            _ConveniadoNome = value
        End Set
    End Property

    Public Property Conveniado() As Negocio.Cliente
        Get
            If _Conveniado Is Nothing And Me.CodigoConveniado.Trim.Length > 0 Then _Conveniado = New Negocio.Cliente(Me.CodigoConveniado, Me.EnderecoConveniado)
            Return _Conveniado
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Conveniado = value
        End Set
    End Property

    Public Property CodigoFatura() As String
        Get
            Return _CodigoFatura
        End Get
        Set(ByVal value As String)
            _CodigoFatura = value
        End Set
    End Property

    Public Property Movimento() As Date
        Get
            Return _Movimento
        End Get
        Set(ByVal value As Date)
            _Movimento = value
        End Set
    End Property

    Public Property ValorDaFatura() As Decimal
        Get
            Return _ValorDaFatura
        End Get
        Set(ByVal value As Decimal)
            _ValorDaFatura = value
        End Set
    End Property

    Public ReadOnly Property ValorLancadoFatura() As Decimal
        Get
            Dim _Vlr As Decimal = 0
            If Itens Is Nothing OrElse Itens.Count = 0 Then
                Return 0
            Else
                For Each row As FaturaDeFreteXItens In _Itens
                    _Vlr += row.ValorLancadoNota
                Next
                Return _Vlr
            End If
        End Get
    End Property

    Public ReadOnly Property ValorSaldoFatura() As Decimal
        Get
            Return ValorDaFatura - ValorLancadoFatura
        End Get
    End Property


    '------------- Alterado --------------------*
    Public ReadOnly Property FinanceiroNovo As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
        End Get
    End Property

    Public Property ListTituloFatura As ListFaturaDeFretexTitulo
        Get
            If _ListTituloFatura Is Nothing Then _ListTituloFatura = New ListFaturaDeFretexTitulo(Me)
            Return _ListTituloFatura
        End Get
        Set(value As ListFaturaDeFretexTitulo)
            _ListTituloFatura = value
        End Set
    End Property

    Public ReadOnly Property Vencimento() As Date
        Get
            If ListTituloFatura Is Nothing OrElse ListTituloFatura.Count = 0 Then
                Return Date.Now
            Else
                If FinanceiroNovo Then
                    Return ListTituloFatura.Select(Function(s) s.TituloNovo.Reprogramacao).Min
                Else
                    Return ListTituloFatura.Select(Function(s) s.Titulo.Prorrogacao).Min
                End If
            End If
        End Get
    End Property

    '------------- fim -------------------------*

    Public Property Itens() As ListFaturasDeFretesXItens
        Get
            If _Itens Is Nothing Then _Itens = New ListFaturasDeFretesXItens(Me)
            Return _Itens
        End Get
        Set(ByVal value As ListFaturasDeFretesXItens)
            _Itens = value
        End Set
    End Property

    Public Property Encargo() As String
        Get
            Return _Encargo
        End Get
        Set(ByVal value As String)
            _Encargo = value
        End Set
    End Property

    Public Property PagarReceber As String
        Get
            Return _PagarReceber
        End Get
        Set(value As String)
            _PagarReceber = value
        End Set
    End Property

    Public Property RegistroMestre As Integer
        Get
            Return _RegistroMestre
        End Get
        Set(value As Integer)
            _RegistroMestre = value
        End Set
    End Property

    Public Property LancamentoManual As Integer
        Get
            Return _LancamentoManual
        End Get
        Set(value As Integer)
            _LancamentoManual = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = Nothing
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String

        Select Case Me.IUD
            Case "I"
                strSQL = "INSERT INTO FaturasDeFretes (Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, Movimento, ValorDaFatura, LancamentoManual)" & vbCrLf &
                         "VALUES ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoConveniado & "'," & Me.EnderecoConveniado & "," & Me.CodigoFatura & "," & "'" & CDate(Me.Movimento).ToString("yyyy-MM-dd") & "'," & Str(Me.ValorDaFatura) & ", " & Me.LancamentoManual & "); "
                Sqls.Add(strSQL)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                strSQL = "UPDATE FaturasDeFretes " & vbCrLf & _
                         "   SET Movimento ='" & CDate(Me.Movimento).ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                         "       ValorDaFatura  = " & Str(Me.ValorDaFatura) & vbCrLf & _
                         " WHERE Empresa_Id       ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id    = " & Me.EnderecoEmpresa & " " & vbCrLf & _
                         "	 AND Conveniado_Id    ='" & Me.CodigoConveniado & "'" & vbCrLf & _
                         "	 AND EndConveniado_Id = " & Me.EnderecoConveniado & " " & vbCrLf & _
                         "	 AND Fatura_Id        = " & Me.CodigoFatura & ";" & vbCrLf
                Sqls.Add(strSQL)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                strSQL = "DELETE FaturasDeFretes" & vbCrLf & _
                         " WHERE Empresa_Id       ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id    = " & Me.EnderecoEmpresa & " " & vbCrLf & _
                         "	 AND Conveniado_Id    ='" & Me.CodigoConveniado & "'" & vbCrLf & _
                         "	 AND EndConveniado_Id = " & Me.EnderecoConveniado & " " & vbCrLf & _
                         "	 AND Fatura_Id        = " & Me.CodigoFatura & ";"
                Sqls.Add(strSQL)
            Case ""
                SalvarTabelasRelacionadasSql(Sqls)
        End Select
    End Sub

    Private Function SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not Me.Itens Is Nothing Then Me.Itens.SalvarSql(Sqls)
        If Not Me.ListTituloFatura Is Nothing Then Me.ListTituloFatura.SalvarSql(Sqls)
        Return Sqls
    End Function

    Public Function JaFaturada() As Boolean
        Dim Banco As New AcessaBanco
        Dim sql As String

        sql = "SELECT 1 " & vbCrLf &
              "  FROM FaturasDeFretes " & vbCrLf &
              " WHERE Conveniado_Id    = '" & Me.CodigoConveniado & "' " & vbCrLf &
              "   AND EndConveniado_Id = '" & Me.EnderecoConveniado & "' " & vbCrLf &
              "   AND Fatura_Id        = '" & Me.CodigoFatura & "' " & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "FaturasDeFretes")
        If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        End If
        Return False
    End Function

    Public Sub AtualizarVinculos(ByRef Sqls As ArrayList)

        Dim strSQL As String

        strSQL = "BEGIN TRY

                    /* ======= PARÂMETROS ======= */
                    DECLARE 
                        @Empresa_Id          VARCHAR(18) = '" & Me.CodigoEmpresa & "',
                        @EndEmpresa_Id       INT         = " & Me.EnderecoEmpresa & ",
                        @Conveniado_Id       VARCHAR(18) = '" & Me.CodigoConveniado & "',
                        @EndConveniado_Id    INT         = " & Me.EnderecoConveniado & ",
                        @FaturaMestre_Id     INT         = " & Me.CodigoFatura & ";

                    /* ======= 0) SANITY CHECK: Fatura Mestre existe? (ok manter conveniado aqui) ======= */
                    IF NOT EXISTS (
                        SELECT 1
                        FROM FaturasDeFretes FF
                        WHERE FF.Empresa_Id       = @Empresa_Id
                          AND FF.EndEmpresa_Id    = @EndEmpresa_Id
                          AND FF.Conveniado_Id    = @Conveniado_Id
                          AND FF.EndConveniado_Id = @EndConveniado_Id
                          AND FF.Fatura_Id        = @FaturaMestre_Id
                    )
                    BEGIN
                        RAISERROR('Fatura Mestre não encontrada com os parâmetros informados.', 16, 1);
                    END;

                    /* ======= 1) NOTAS da Fatura Mestre (SEM filtrar por Conveniado) ======= */
                    IF OBJECT_ID('tempdb..#NotasMestre') IS NOT NULL DROP TABLE #NotasMestre;
                    SELECT DISTINCT
                        Xi.Empresa_Id,
                        Xi.EndEmpresa_Id,
                        Xi.EntradaSaida_Id,
                        Xi.Serie_Id,
                        Xi.Nota_Id
                    INTO #NotasMestre
                    FROM FaturasDeFretesXItens Xi
                    WHERE Xi.Empresa_Id     = @Empresa_Id
                      AND Xi.EndEmpresa_Id  = @EndEmpresa_Id
                      AND Xi.Fatura_Id      = @FaturaMestre_Id;

                    /* ======= 2) FATURAS FILHAS por notas (independente do Conveniado) ======= */
                    IF OBJECT_ID('tempdb..#FaturasFilhas') IS NOT NULL DROP TABLE #FaturasFilhas;
                    SELECT DISTINCT
                        X.Empresa_Id,
                        X.EndEmpresa_Id,
                        X.Conveniado_Id,       -- pode ser diferente da mestre
                        X.EndConveniado_Id,
                        X.Fatura_Id
                    INTO #FaturasFilhas
                    FROM FaturasDeFretesXItens X
                    INNER JOIN #NotasMestre N
                        ON  N.Empresa_Id        = X.Empresa_Id
                        AND N.EndEmpresa_Id     = X.EndEmpresa_Id
                        AND N.Nota_Id           = X.Nota_Id
                        AND N.Serie_Id          = X.Serie_Id
                        AND N.EntradaSaida_Id   = X.EntradaSaida_Id
                    WHERE X.Fatura_Id <> @FaturaMestre_Id;

                    /* ======= 3) TÍTULO MESTRE (ligado à Fatura Mestre) ======= */
                    IF OBJECT_ID('tempdb..#TituloMestre') IS NOT NULL DROP TABLE #TituloMestre;

                    ;WITH TMs AS (
                        SELECT 
                            T.Empresa_Id,
                            T.EndEmpresa_Id,
                            T.Fatura_Id,
                            T.Titulo_Id,
                            Fonte = 'CAR'
                        FROM FaturaDeFreteXTitulo T
                        INNER JOIN ContasAReceber CAR
                            ON CAR.Registro_Id = T.Titulo_Id
                        WHERE T.Empresa_Id       = @Empresa_Id
                          AND T.EndEmpresa_Id    = @EndEmpresa_Id
                          AND T.Fatura_Id        = @FaturaMestre_Id

                        UNION ALL

                        SELECT 
                            T.Empresa_Id,
                            T.EndEmpresa_Id,
                            T.Fatura_Id,
                            T.Titulo_Id,
                            Fonte = 'CAP'
                        FROM FaturaDeFreteXTitulo T
                        INNER JOIN ContasAPagar CAP
                            ON CAP.Registro_Id = T.Titulo_Id
                        WHERE T.Empresa_Id       = @Empresa_Id
                          AND T.EndEmpresa_Id    = @EndEmpresa_Id
                          AND T.Fatura_Id        = @FaturaMestre_Id
                    )
                    SELECT TOP (1) * INTO #TituloMestre
                    FROM TMs
                    ORDER BY Fonte;  -- prioridade CAR; troque se quiser priorizar CAP

                    IF NOT EXISTS (SELECT 1 FROM #TituloMestre)
                        RAISERROR('Nenhum Título Mestre encontrado para a fatura mestre.', 16, 1);

                    DECLARE 
                        @TituloMestre_Id INT,
                        @FonteMestre     VARCHAR(3);
                    SELECT 
                        @TituloMestre_Id = Titulo_Id,
                        @FonteMestre     = Fonte
                    FROM #TituloMestre;

                    /* Baixa do Título Mestre (para replicar nos filhos) */
                    DECLARE @BaixaMestre DATE = NULL;
                    IF (@FonteMestre = 'CAR')
                        SELECT @BaixaMestre = CAR.Baixa FROM ContasAReceber CAR WHERE CAR.Registro_Id = @TituloMestre_Id;
                    ELSE
                        SELECT @BaixaMestre = CAP.Baixa FROM ContasAPagar CAP WHERE CAP.Registro_Id = @TituloMestre_Id;

                    /* Data de baixa a aplicar nos filhos: se nula no mestre, usa hoje */
                    DECLARE @BaixaAplicar DATE = ISNULL(@BaixaMestre, CONVERT(date, '" & CDate(Me.Vencimento).ToString("yyyy-MM-dd") & "'));

                    /* ======= 4) TÍTULOS FILHOS (ligados às Faturas Filhas) ======= */
                    IF OBJECT_ID('tempdb..#TitulosFilhos') IS NOT NULL DROP TABLE #TitulosFilhos;

                    ;WITH TF AS (
                        SELECT 
                            T.Empresa_Id,
                            T.EndEmpresa_Id,
                            T.Fatura_Id,
                            T.Titulo_Id,
                            Fonte = 'CAR'
                        FROM FaturaDeFreteXTitulo T
                        INNER JOIN #FaturasFilhas F
                           ON F.Empresa_Id = T.Empresa_Id
                          AND F.EndEmpresa_Id = T.EndEmpresa_Id
                          AND F.Fatura_Id = T.Fatura_Id
                        INNER JOIN ContasAReceber CAR
                           ON CAR.Registro_Id = T.Titulo_Id

                        UNION ALL

                        SELECT 
                            T.Empresa_Id,
                            T.EndEmpresa_Id,
                            T.Fatura_Id,
                            T.Titulo_Id,
                            Fonte = 'CAP'
                        FROM FaturaDeFreteXTitulo T
                        INNER JOIN #FaturasFilhas F
                           ON F.Empresa_Id = T.Empresa_Id
                          AND F.EndEmpresa_Id = T.EndEmpresa_Id
                          AND F.Fatura_Id = T.Fatura_Id
                        INNER JOIN ContasAPagar CAP
                           ON CAP.Registro_Id = T.Titulo_Id
                    )
                    SELECT * INTO #TitulosFilhos FROM TF;

                    /* ======= 4.1) Observação do Título Mestre (lista de filhos), sem STRING_AGG ======= */
                    DECLARE @QtdeFilhos INT;
                    SELECT @QtdeFilhos = COUNT(*) FROM #TitulosFilhos;

                    DECLARE @ObsLista NVARCHAR(MAX) = N'';

                    IF (@QtdeFilhos > 0)
                    BEGIN
                        SELECT @ObsLista =
                            STUFF((
                                SELECT N' - ' + CAST(TF.Titulo_Id AS NVARCHAR(30))
                                FROM #TitulosFilhos AS TF
                                ORDER BY TF.Titulo_Id
                                FOR XML PATH(''), TYPE
                            ).value('.', 'NVARCHAR(MAX)'), 1, 3, N'');
                    END;

                                        IF (@QtdeFilhos > 0)
                    BEGIN
                        SELECT @ObsLista =
                            STUFF((
                                SELECT N' - ' + CAST(TF.Titulo_Id AS NVARCHAR(30))
                                FROM #TitulosFilhos AS TF
                                ORDER BY TF.Titulo_Id
                                FOR XML PATH(''), TYPE
                            ).value('.', 'NVARCHAR(MAX)'), 1, 3, N'');
                    END;

                    DECLARE @ObservacaoMestre NVARCHAR(MAX) =
                    CASE 
                        WHEN @QtdeFilhos = 0 THEN N''  -- << sem filhos: deixa em branco
                        ELSE N'Agrupamento dos Registros - ' + @ObsLista
                    END;

                     /* ======= 6) ATUALIZAÇÕES ======= */

                    /* 6.1) Faturas FILHAS: setar RegistroMestre = FaturaMestre */
                    UPDATE FF
                       SET FF.RegistroMestre = @FaturaMestre_Id
                    FROM FaturasDeFretes FF
                    INNER JOIN #FaturasFilhas F
                       ON F.Empresa_Id = FF.Empresa_Id
                      AND F.EndEmpresa_Id = FF.EndEmpresa_Id
                      AND F.Fatura_Id = FF.Fatura_Id;

                    /* 6.2) TÍTULO MESTRE: Grupado = 'M' + Observacoes
                          - Se não houver filhos: limpa (em branco)
                          - Senão: sobrescreve na 1ª vez ou concatena nas próximas */
                    IF (@QtdeFilhos = 0)
                    BEGIN
                        IF (@FonteMestre = 'CAR')
                        BEGIN
                            UPDATE CAR
                               SET CAR.Grupado     = 'M',
                                   CAR.Observacoes = N''
                            FROM ContasAReceber CAR
                            WHERE CAR.Registro_Id = @TituloMestre_Id;
                        END
                        ELSE
                        BEGIN
                            UPDATE CAP
                               SET CAP.Grupado     = 'M',
                                   CAP.Observacoes = N''
                            FROM ContasAPagar CAP
                            WHERE CAP.Registro_Id = @TituloMestre_Id;
                        END
                    END
                    ELSE
                    BEGIN
                        IF (@FonteMestre = 'CAR')
                        BEGIN
                            UPDATE CAR
                               SET CAR.Grupado     = 'M',
                                   CAR.Observacoes =
                                       CASE
                                           WHEN ISNULL(LTRIM(RTRIM(CAR.Observacoes)),'') = ''
                                                OR CAR.Observacoes LIKE 'Agrupamento dos Registros%'
                                                THEN @ObservacaoMestre
                                           ELSE RTRIM(CAR.Observacoes) + ' | ' + @ObservacaoMestre
                                       END
                            FROM ContasAReceber CAR
                            WHERE CAR.Registro_Id = @TituloMestre_Id;
                        END
                        ELSE
                        BEGIN
                            UPDATE CAP
                               SET CAP.Grupado     = 'M',
                                   CAP.Observacoes =
                                       CASE
                                           WHEN ISNULL(LTRIM(RTRIM(CAP.Observacoes)),'') = ''
                                                OR CAP.Observacoes LIKE 'Agrupamento dos Registros%'
                                                THEN @ObservacaoMestre
                                           ELSE RTRIM(CAP.Observacoes) + ' | ' + @ObservacaoMestre
                                       END
                            FROM ContasAPagar CAP
                            WHERE CAP.Registro_Id = @TituloMestre_Id;
                        END
                    END

                    /* 6.3) TÍTULOS FILHOS:
                           Grupado='S', RegistroMestre = TituloMestre,
                           Situacao=1, Provisao=1, Baixa = @BaixaAplicar */
                    /* Contas A Receber (filhos) */
                    UPDATE CAR
                       SET CAR.Grupado        = 'S',
                           CAR.RegistroMestre = @TituloMestre_Id,
                           CAR.Situacao       = 1,
                           CAR.Provisao       = 1,
                           CAR.Baixa          = @BaixaAplicar
                    FROM ContasAReceber CAR
                    INNER JOIN #TitulosFilhos TF
                       ON TF.Titulo_Id = CAR.Registro_Id
                      AND TF.Fonte     = 'CAR';

                    /* Contas A Pagar (filhos) */
                    UPDATE CAP
                       SET CAP.Grupado        = 'S',
                           CAP.RegistroMestre = @TituloMestre_Id,
                           CAP.Situacao       = 1,
                           CAP.Provisao       = 1,
                           CAP.Baixa          = @BaixaAplicar
                    FROM ContasAPagar CAP
                    INNER JOIN #TitulosFilhos TF
                       ON TF.Titulo_Id = CAP.Registro_Id
                      AND TF.Fonte     = 'CAP';

                    /* 6.4) DESAGRUPAR TÍTULOS QUE NÃO SÃO MAIS FILHOS (foram removidos das faturas filhas)
                           - RegistroMestre = NULL
                           - Grupado (ou Agrupamento) = NULL
                           - Provisao = 3 (previsão)
                    */

                    /* Contas A Receber órfãos: apontam para o mestre, mas não estão em #TitulosFilhos */
                    UPDATE CAR
                       SET CAR.RegistroMestre = NULL,
                           CAR.Grupado        = NULL,   -- Se a coluna se chama Agrupamento, troque para CAR.Agrupamento = NULL
                           CAR.Provisao       = 3
                    FROM ContasAReceber CAR
                    WHERE CAR.RegistroMestre = @TituloMestre_Id
                      AND NOT EXISTS (
                            SELECT 1
                            FROM #TitulosFilhos TF
                            WHERE TF.Titulo_Id = CAR.Registro_Id
                              AND TF.Fonte = 'CAR'
                      );

                    /* Contas A Pagar órfãos */
                    UPDATE CAP
                       SET CAP.RegistroMestre = NULL,
                           CAP.Grupado        = NULL,   -- Se a coluna se chama Agrupamento, troque para CAP.Agrupamento = NULL
                           CAP.Provisao       = 3
                    FROM ContasAPagar CAP
                    WHERE CAP.RegistroMestre = @TituloMestre_Id
                      AND NOT EXISTS (
                            SELECT 1
                            FROM #TitulosFilhos TF
                            WHERE TF.Titulo_Id = CAP.Registro_Id
                              AND TF.Fonte = 'CAP'
                      );

                    /* 6.5) DESAGRUPAR FATURAS FILHAS QUE NÃO SÃO MAIS FILHAS
                           - Zera RegistroMestre nas FaturasDeFretes que ainda apontam para a mestre,
                             mas não estão em #FaturasFilhas (e não são a própria fatura mestre) */

                    UPDATE FF
                       SET FF.RegistroMestre = NULL
                    FROM FaturasDeFretes FF
                    WHERE FF.Empresa_Id      = @Empresa_Id
                      AND FF.EndEmpresa_Id   = @EndEmpresa_Id
                      AND FF.RegistroMestre  = @FaturaMestre_Id
                      AND FF.Fatura_Id      <> @FaturaMestre_Id
                      AND NOT EXISTS (
                            SELECT 1
                            FROM #FaturasFilhas F
                            WHERE F.Empresa_Id    = FF.Empresa_Id
                              AND F.EndEmpresa_Id = FF.EndEmpresa_Id
                              AND F.Fatura_Id     = FF.Fatura_Id
                      );

                    /* ======= 7) SAÍDA DE CONFERÊNCIA ======= */

                    PRINT 'Atualizações concluídas. Conferência:';

                    PRINT 'Faturas Filhas atualizadas (RegistroMestre = FaturaMestre):';
                    SELECT FF.Empresa_Id, FF.EndEmpresa_Id, FF.Conveniado_Id, FF.Fatura_Id, FF.RegistroMestre
                    FROM FaturasDeFretes FF
                    INNER JOIN #FaturasFilhas F
                       ON F.Empresa_Id = FF.Empresa_Id
                      AND F.EndEmpresa_Id = FF.EndEmpresa_Id
                      AND F.Fatura_Id = FF.Fatura_Id
                    ORDER BY FF.Fatura_Id;

                    PRINT 'Título Mestre após atualização:';
                    IF (@FonteMestre = 'CAR')
                        SELECT Registro_Id, Grupado, Observacoes, Baixa
                        FROM ContasAReceber
                        WHERE Registro_Id = @TituloMestre_Id;
                    ELSE
                        SELECT Registro_Id, Grupado, Observacoes, Baixa
                        FROM ContasAPagar
                        WHERE Registro_Id = @TituloMestre_Id;

                    PRINT 'Títulos Filhos após atualização (CAR):';
                    SELECT CAR.Registro_Id, CAR.Grupado, CAR.RegistroMestre, CAR.Situacao, CAR.Provisao, CAR.Baixa
                    FROM ContasAReceber CAR
                    INNER JOIN #TitulosFilhos TF ON TF.Titulo_Id = CAR.Registro_Id AND TF.Fonte='CAR'
                    ORDER BY CAR.Registro_Id;

                    PRINT 'Títulos Filhos após atualização (CAP):';
                    SELECT CAP.Registro_Id, CAP.Grupado, CAP.RegistroMestre, CAP.Situacao, CAP.Provisao, CAP.Baixa
                    FROM ContasAPagar CAP
                    INNER JOIN #TitulosFilhos TF ON TF.Titulo_Id = CAP.Registro_Id AND TF.Fonte='CAP'
                    ORDER BY CAP.Registro_Id;

                END TRY
                BEGIN CATCH
                    DECLARE @ErrMsg NVARCHAR(4000) = ERROR_MESSAGE(), @ErrSev INT = ERROR_SEVERITY(), @ErrState INT = ERROR_STATE();
                    RAISERROR(@ErrMsg, @ErrSev, @ErrState);
                END CATCH; "

        Sqls.Add(strSQL)

    End Sub

#End Region

End Class