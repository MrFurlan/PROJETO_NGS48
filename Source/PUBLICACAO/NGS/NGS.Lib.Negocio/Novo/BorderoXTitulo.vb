Imports System.Data
Imports NGS.Lib.Uteis

Namespace Novo
    <Serializable()> _
    Public Class ListBorderoXTitulo
        Inherits List(Of BorderoXTitulo)

#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(ByVal pBordero As Novo.Bordero)
            _Bordero = pBordero
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, Titulo_Id, isnull(TituloRecompra,0) as TituloRecompra, isnull(BorderoRecompra,0) as BorderoRecompra, isnull(Situacao,1) as Situacao" & vbCrLf & _
                  "  From BorderoXTitulo" & vbCrLf & _
                  " Where Empresa_Id    ='" & pBordero.CodigoEmpresa & "'" & vbCrLf & _
                  "   and EndEmpresa_Id = " & pBordero.EnderecoEmpresa & vbCrLf & _
                  "   and Bordero_Id    = " & pBordero.CodigoBordero

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Bordero")

            For Each row In ds.Tables(0).Rows
                Dim BxT As New BorderoXTitulo()
                BxT.CodigoEmpresa = row("Empresa_Id")
                BxT.EnderecoEmpresa = row("EndEmpresa_Id")
                BxT.CodigoBordero = row("Bordero_Id")
                BxT.CodigoTitulo = row("Titulo_Id")
                BxT.CodigoTituloRecompra = row("TituloRecompra")
                BxT.CodigoBorderoRecompra = row("BorderoRecompra")
                BxT.CodigoSituacao = row("Situacao")
                Me.Add(BxT)
            Next

        End Sub

        Public Sub New(ByVal pCodigoTituloOuBorderoSeRecompra As Integer, ByVal Recompra As Boolean)
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, Titulo_Id, TituloRecompra, isnull(BorderoRecompra,0) as BorderoRecompra, isnull(Situacao,1) as Situacao" & vbCrLf & _
                  "  From BorderoXTitulo" & vbCrLf

            If Recompra Then
                sql &= " Where BorderoRecompra = " & pCodigoTituloOuBorderoSeRecompra & vbCrLf
            Else
                sql &= " Where Titulo_Id       = " & pCodigoTituloOuBorderoSeRecompra & vbCrLf
            End If


            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "BorderoXTitulo")

            For Each row In ds.Tables(0).Rows
                Dim BxT As New BorderoXTitulo()
                BxT.CodigoEmpresa = row("Empresa_Id")
                BxT.EnderecoEmpresa = row("EndEmpresa_Id")
                BxT.CodigoBordero = row("Bordero_Id")
                BxT.CodigoTitulo = row("Titulo_Id")
                If Not row.IsNull("TituloRecompra") Then BxT.CodigoTituloRecompra = row("TituloRecompra")
                BxT.CodigoBorderoRecompra = row("BorderoRecompra")
                BxT.CodigoSituacao = row("Situacao")
                Me.Add(BxT)
            Next

        End Sub
#End Region

#Region "Fields"
        Private _Bordero As Novo.Bordero
#End Region

#Region "Property"
        Public Property Bordero As Novo.Bordero
            Get
                Return _Bordero
            End Get
            Set(ByVal value As Novo.Bordero)
                _Bordero = value
            End Set
        End Property
#End Region

#Region "Methods"
        Public Function Salvar() As Boolean
            Dim Banco As New AcessaBanco
            Dim sqls As New ArrayList

            sqls.Clear()
            Me.SalvarSQL(sqls)

            If sqls.Count = 0 OrElse Banco.GravaBanco(sqls) Then
                Return True
            Else
                Return False
            End If
        End Function

        Public Sub SalvarSQL(ByRef Sqls As ArrayList)
            For Each TxB In Me
                TxB.IUD = Bordero.IUD
                TxB.SalvarSql(Sqls)
            Next
        End Sub
#End Region

    End Class

    <Serializable()> _
    Public Class BorderoXTitulo

#Region "Construtor"
        Public Sub New()

        End Sub

        Public Sub New(ByVal pCodigoBordero As Integer, ByVal pCodigoTitulo As Integer)
            Dim sql As String = ""
            sql = "Select Empresa_Id, EndEmpresa_Id, Bordero_Id, Titulo_Id, TituloRecompra, BorderoRecompra, Situacao" & vbCrLf & _
                  "  From BorderoXTitulo" & vbCrLf & _
                  " Where Titulo_Id  = " & pCodigoTitulo & _
                  "   and Bordero_Id = " & pCodigoBordero

            Dim Banco As New AcessaBanco
            Dim ds As DataSet = Banco.ConsultaDataSet(sql, "BorderoXtitulo")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub

            Dim row = ds.Tables(0).Rows(0)
            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.CodigoBordero = row("Bordero_Id")
            Me.CodigoTitulo = row("Titulo_Id")
            Me.CodigoTituloRecompra = row("TituloRecompra")
            Me.CodigoBorderoRecompra = row("BorderoRecompra")
            Me.CodigoSituacao = row("Situacao")
        End Sub
#End Region

#Region "Fields"
        Private _IUD As String = ""
        Private _CodigoEmpresa As String = ""
        Private _EndEmpresa As Integer
        Private _Empresa As Cliente
        Private _CodigoBordero As Integer
        Private _Bordero As Novo.Bordero
        Private _CodigoTitulo As Integer
        Private _Titulo As Novo.TituloNovo
        Private _CodigoTituloRecompra As Integer
        Private _TituloRecompra As Novo.TituloNovo
        Private _CodigoBorderoRecompra As Integer
        Private _CodigoSituacao As Integer
#End Region

#Region "Property"
        'Teste Depois Apagar
        Public ReadOnly Property Codigo As Integer
            Get
                Return _CodigoTitulo
            End Get
        End Property
        '**********************
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
            Set(ByVal value As Integer)
                _CodigoBordero = value
                _Bordero = Nothing
            End Set
        End Property
        Public Property Bordero As Novo.Bordero
            Get
                If _Bordero Is Nothing And _CodigoBordero > 0 Then _Bordero = New Novo.Bordero(_CodigoEmpresa, _EndEmpresa, _CodigoBordero)
                Return _Bordero
            End Get
            Set(ByVal value As Novo.Bordero)
                _Bordero = value
                _CodigoBordero = value.CodigoBordero
            End Set
        End Property
        Public Property CodigoTitulo() As Integer
            Get
                Return _CodigoTitulo
            End Get
            Set(ByVal value As Integer)
                _CodigoTitulo = value
                _Titulo = Nothing
            End Set
        End Property
        Public Property Titulo As Novo.TituloNovo
            Get
                If _Titulo Is Nothing And _CodigoTitulo > 0 Then _Titulo = New Novo.TituloNovo(_CodigoTitulo)
                Return _Titulo
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _Titulo = value
            End Set
        End Property
        Public Property CodigoTituloRecompra() As Integer
            Get
                Return _CodigoTituloRecompra
            End Get
            Set(ByVal value As Integer)
                _CodigoTituloRecompra = value
                _TituloRecompra = Nothing
            End Set
        End Property
        Public Property TituloRecompra() As Novo.TituloNovo
            Get
                If _TituloRecompra Is Nothing And CodigoTituloRecompra > 0 Then _TituloRecompra = New Novo.TituloNovo(CodigoTituloRecompra)
                Return _TituloRecompra
            End Get
            Set(ByVal value As Novo.TituloNovo)
                _TituloRecompra = value
            End Set
        End Property
        Public Property CodigoBorderoRecompra As Integer
            Get
                Return _CodigoBorderoRecompra
            End Get
            Set(ByVal value As Integer)
                _CodigoBorderoRecompra = value
            End Set
        End Property
        Public Property CodigoSituacao As Integer
            Get
                Return _CodigoSituacao
            End Get
            Set(ByVal value As Integer)
                _CodigoSituacao = value
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
                    strSQL = " INSERT INTO BorderoXTitulo(Empresa_Id, EndEmpresa_Id, Bordero_Id, Titulo_Id, TituloRecompra, BorderoRecompra, Situacao) " & vbCrLf & _
                             " VALUES ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.CodigoBordero & "," & Me.CodigoTitulo & "," & IIf(Me.CodigoTituloRecompra > 0, Me.CodigoTituloRecompra, "NULL") & "," & IIf(Me.CodigoBorderoRecompra > 0, Me.CodigoBorderoRecompra, "NULL") & "," & Me.CodigoSituacao & ")"
                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "U"
                    strSQL = " UPDATE BorderoXTitulo Set " & vbCrLf & _
                             "    TituloRecompra  = " & Me.CodigoTituloRecompra & vbCrLf & _
                             "    ,BorderoRecompra = " & Me.CodigoBorderoRecompra & vbCrLf & _
                             "    ,Situacao        = " & Me.CodigoSituacao & vbCrLf & _
                             " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                             "   And EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                             "   And Bordero_id    = " & Me.CodigoBordero & vbCrLf & _
                             "   And Titulo_Id     = " & Me.CodigoTitulo
                    sqls.Add(strSQL)
                    SalvarTabelasRelacionadasSql(sqls)
                Case "D"
                    SalvarTabelasRelacionadasSql(sqls)
                    strSQL = "Update BorderoXTitulo Set" & vbCrLf & _
                             "   Situacao = 2" & vbCrLf & _
                             " WHERE Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                             "   And EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                             "   And Bordero_id    = " & Me.CodigoBordero & vbCrLf & _
                             "   And Titulo_Id     = " & Me.CodigoTitulo
                    sqls.Add(strSQL)
                Case Else
                    SalvarTabelasRelacionadasSql(sqls)
            End Select
        End Sub

        Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
            If Me.IUD = "I" Or Me.IUD = "U" Then
                If Me.Titulo.CodigoProvisao <> 1 Then
                    Me.Titulo.IUD = "U"
                    If Me.Bordero.CodigoTituloBordero > 0 Then
                        Me.Titulo.CodigoEmpresaRecPag = Me.Bordero.TituloDoBordero.CodigoEmpresaRecPag
                        Me.Titulo.EndEmpresaRecPag = Me.Bordero.TituloDoBordero.EndEmpresaRecPag
                        Me.Titulo.CodigoTipoPgto = Me.Bordero.TituloDoBordero.CodigoTipoPgto
                        Me.Titulo.CodigoContaContabilRecPag = Me.Bordero.TituloDoBordero.CodigoContaContabilRecPag
                        Me.Titulo.CodigoCarteiraDoTitulo = Me.Bordero.TituloDoBordero.CodigoCarteiraDoTitulo
                    End If
                    Me.Titulo.SalvarSql(Sqls)
                End If
            ElseIf Me.IUD = "D" Then
                If Me.Titulo.CodigoProvisao <> 1 Then
                    Me.Titulo.IUD = "U"
                    Me.Titulo.CodigoEmpresaRecPag = Me.Titulo.CodigoEmpresa
                    Me.Titulo.EndEmpresaRecPag = Me.Titulo.EnderecoEmpresa
                    Me.Titulo.CodigoTipoPgto = 1
                    Me.Titulo.CodigoContaContabilRecPag = Me.Titulo.Empresa.Empresa.CodigoContaGrupoBanco
                    Me.Titulo.CodigoCarteiraDoTitulo = 0
                    Me.Titulo.SalvarSql(Sqls)
                End If
            End If
        End Sub
#End Region

    End Class
End Namespace

