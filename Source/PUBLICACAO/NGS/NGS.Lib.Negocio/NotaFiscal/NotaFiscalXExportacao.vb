Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class NotaFiscalXExportacao

#Region "Contrutor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota

        Dim sql As String
        sql = " SELECT ISNULL(NrDespachoExp, '') AS NrDespachoExp, isnull(DataDespachoExp,getdate()) as DataDespachoExp, " & vbCrLf & _
              "        PaisDestino, Navio, DataAverbacao, isnull(DrawBack, '') as DrawBack, isnull(FaturaExportacao, '') as FaturaExportacao " & vbCrLf & _
              "   FROM NotaFiscalXExportacao" & vbCrLf & _
              "  Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
              "    and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
              "    and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
              "    and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
              "    and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "    and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
              "    and Nota_Id         = " & Me.NF.Codigo & vbCrLf & _
              "    and Nota_id > 0"
        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(sql, "DadosExportacao")

        If ds.Tables(0).Rows.Count = 0 Then
            Me.IUD = "N"
            Exit Sub
        End If


        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.IUD = "U"
        Me.NrDespachoExp = row("NrDespachoExp")
        Me.DataDespachoExp = row("DataDespachoExp")
        Me.CodigoPaisDestino = row("PaisDestino")
        Me.Navio = row("Navio")
        Me.DataAverbacao = IIf(IsDBNull(row("DataAverbacao")), Nothing, row("DataAverbacao"))
        Me.NumAtoConcessorio = row("DrawBack")
        Me.FaturaExportacao = row("FaturaExportacao")
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal

    Private _IUD As String = ""

    Private _NrDespachoExp As String = ""
    Private _DataDespachoExp As Date = Now

    Private _CodigoPaisDestino As Integer
    Private _Navio As String = ""

    Private _DataAverbacao As Date?
    Private _NumAtoConcessorio As String = ""
    Private _DtaRegAtoConcessorio As Date = Now
    Private _DtaValidAtoConcessorio As Date = Now

    Private _FaturaExportacao As String = ""
#End Region

#Region "Property"
    Public ReadOnly Property NF() As NotaFiscal
        Get
            Return _NF
        End Get
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property NrDespachoExp() As String
        Get
            Return _NrDespachoExp
        End Get
        Set(ByVal value As String)
            _NrDespachoExp = value
        End Set
    End Property

    Public Property DataDespachoExp() As Date
        Get
            Return _DataDespachoExp
        End Get
        Set(ByVal value As Date)
            _DataDespachoExp = value
        End Set
    End Property


    Public Property CodigoPaisDestino() As Integer
        Get
            Return _CodigoPaisDestino
        End Get
        Set(ByVal value As Integer)
            _CodigoPaisDestino = value
        End Set
    End Property

    Public Property Navio() As String
        Get
            Return _Navio
        End Get
        Set(ByVal value As String)
            _Navio = value
        End Set
    End Property

    Public Property DataAverbacao() As Date?
        Get
            Return _DataAverbacao
        End Get
        Set(ByVal value As Date?)
            _DataAverbacao = value
        End Set
    End Property

    Public Property NumAtoConcessorio() As String
        Get
            Return _NumAtoConcessorio
        End Get
        Set(ByVal value As String)
            _NumAtoConcessorio = value.ToUpper
        End Set
    End Property

    Public Property DtaRegAtoConcessorio() As Date
        Get
            Return _DtaRegAtoConcessorio
        End Get
        Set(ByVal value As Date)
            _DtaRegAtoConcessorio = value

        End Set
    End Property

    Public Property DtaValidAtoConcessorio() As Date
        Get
            Return _DtaValidAtoConcessorio
        End Get
        Set(ByVal value As Date)
            _DtaValidAtoConcessorio = value

        End Set
    End Property

    Public Property FaturaExportacao() As String
        Get
            Return _FaturaExportacao
        End Get
        Set(ByVal value As String)
            _FaturaExportacao = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = ""
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into NotaFiscalXExportacao(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id," & vbCrLf & _
                      "                                   NrDespachoExp, DataDespachoExp, " & vbCrLf & _
                      "                                   PaisDestino, Navio, DataAverbacao, DrawBack, FaturaExportacao) " & vbCrLf & _
                      " Values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ",'" & NF.CodigoCliente & "'," & NF.EnderecoCliente & ",'" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & NF.Serie & "'," & NF.Codigo & "," & vbCrLf & _
                      IIf(Me.NrDespachoExp.Trim.Length = 0, "NULL", "'" & Me.NrDespachoExp.Trim & "'") & "," & IIf(Me.NrDespachoExp.Trim.Length = 0, "NULL", "'" & Me.DataDespachoExp.ToString("yyyy-MM-dd") & "'") & ", " & vbCrLf & _
                      Me.CodigoPaisDestino & ",'" & Me.Navio & "',"
                If Me.DataAverbacao Is Nothing Then
                    sql &= "NULL"
                Else
                    sql &= "'" & CDate(Me.DataAverbacao).ToString("yyyy-MM-dd") & "'"
                End If

                sql &= ", '" & Me.NumAtoConcessorio & "', '" & Me.FaturaExportacao & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update NotaFiscalXExportacao set" & vbCrLf & _
                      "   NrDespachoExp       =" & IIf(Me.NrDespachoExp.Trim.Length = 0, "NULL", "'" & Me.NrDespachoExp.Trim & "'") & vbCrLf & _
                      "  ,DataDespachoExp     = " & IIf(_NrDespachoExp.Length = 0, "NULL", "'" & Me.DataDespachoExp.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,PaisDestino         = " & Me.CodigoPaisDestino & vbCrLf & _
                      "  ,Navio               ='" & Me.Navio & "'" & vbCrLf & _
                      "  ,DataAverbacao       ='" & CDate(Me.DataAverbacao).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "  ,DrawBack            ='" & Me.NumAtoConcessorio & "'" & vbCrLf & _
                      "  ,FaturaExportacao    ='" & Me.FaturaExportacao & "'" & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalXExportacao" & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
