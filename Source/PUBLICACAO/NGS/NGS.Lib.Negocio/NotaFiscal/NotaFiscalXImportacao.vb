Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class NotaFiscalXImportacao

#Region "Contrutor"
    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota

        Dim sql As String
        sql = " SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id,  " & vbCrLf & _
              " NumeroDeclaracaoImportacao,  isnull(DataDeclaracaoImportacao,getdate()) as DataDeclaracaoImportacao, isnull(RegistroDeImportacao,'') as RegistroDeImportacao,  isnull(DataRegistroDeImportacao,getdate()) as DataRegistroDeImportacao, " & vbCrLf & _
              " isnull(LocalEmbarqueImportacao,'') as LocalEmbarqueImportacao, isnull(EstadoEmbarqueImportacao,'') as EstadoEmbarqueImportacao,  isnull(DataEmbarqueImportacao,getdate()) as DataEmbarqueImportacao, " & vbCrLf & _
              " isnull(LocalDesembarqueImportacao,'') as LocalDesembarqueImportacao, isnull(EstadoDesembarqueImportacao,'') as EstadoDesembarqueImportacao,  isnull(DataDesembarqueImportacao,getdate()) as DataDesembarqueImportacao, " & vbCrLf & _
              " isnull(Fabricante,'') as Fabricante, EndFabricante, isnull(NumAtoConcessorio,'') as NumAtoConcessorio,  isnull(DtaRegAtoConcessorio,getdate()) as DtaRegAtoConcessorio,  isnull(DtaValidAtoConcessorio,getdate()) as DtaValidAtoConcessorio, isnull(NrFatura,'') as NrFatura, " & vbCrLf & _
              " isnull(ViaDeTransporte, 0) as ViaDeTransporte,  isnull(TipoDeImportacao, 0) AS TipoDeImportacao, isnull(ValorVAFRMM, 0) AS ValorVAFRMM " & vbCrLf & _
              "   FROM NotaFiscalXImportacao" & vbCrLf & _
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

        ds = Banco.ConsultaDataSet(sql, "DadosImportacao")

        If ds.Tables(0).Rows.Count = 0 Then
            Me.IUD = "N"
            Exit Sub
        End If


        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.IUD = "U"
        Me.NumeroDeclaracaoImportacao = row("NumeroDeclaracaoImportacao")
        Me.DataDeclaracaoImportacao = row("DataDeclaracaoImportacao")
        Me.RegistroDeImportacao = row("RegistroDeImportacao")
        Me.DataRegistroDeImportacao = row("DataRegistroDeImportacao")
        Me.LocalEmbarqueImportacao = row("LocalEmbarqueImportacao")
        Me.EstadoEmbarqueImportacao = row("EstadoEmbarqueImportacao")
        Me.DataEmbarqueImportacao = row("DataEmbarqueImportacao")
        Me.LocalDesembarqueImportacao = row("LocalDesembarqueImportacao")
        Me.EstadoDesembarqueImportacao = row("EstadoDesembarqueImportacao")
        Me.DataDesembarqueImportacao = row("DataDesembarqueImportacao")
        Me.CodigoFabricante = row("Fabricante")
        Me.EndFabricante = row("EndFabricante")
        Me.NumAtoConcessorio = row("NumAtoConcessorio")
        Me.DtaRegAtoConcessorio = row("DtaRegAtoConcessorio")
        Me.DtaValidAtoConcessorio = row("DtaValidAtoConcessorio")
        Me.NrFatura = row("NrFatura")
        Me.ViaDeTransporte = row("ViaDeTransporte")
        Me.TipoDeImportacao = row("TipoDeImportacao")
        Me.ValorVAFRMM = row("ValorVAFRMM")

    End Sub
#End Region

#Region "Fields"

    Private _NF As NotaFiscal

    Private _IUD As String = ""

    Private _NumeroDeclaracaoImportacao As String = ""
    Private _DataDeclaracaoImportacao As Date = Now

    Private _RegistroDeImportacao As String = ""
    Private _DataRegistroDeImportacao As Date = Now

    Private _LocalEmbarqueImportacao As String
    Private _EstadoEmbarqueImportacao As String
    Private _DataEmbarqueImportacao As DateTime

    '*****************************************
    Private _LocalDesembarqueImportacao As String
    Private _EstadoDesembarqueImportacao As String
    Private _DataDesembarqueImportacao As DateTime
    Private _CodigoFabricante As String
    Private _EndFabricante As Integer
    Private _Fabricante As Cliente

    '*****************************************

    Private _NumAtoConcessorio As String = ""
    Private _DtaRegAtoConcessorio As Date = Now
    Private _DtaValidAtoConcessorio As Date = Now

    Private _NrFatura As String = ""

    Private _ViaDeTransporte As Integer
    Private _TipoDeImportacao As Integer
    Private _ValorVAFRMM As Decimal
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

    Public Property NumeroDeclaracaoImportacao() As String
        Get
            Return _NumeroDeclaracaoImportacao
        End Get
        Set(ByVal value As String)
            _NumeroDeclaracaoImportacao = value
        End Set
    End Property

    Public Property DataDeclaracaoImportacao() As Date
        Get
            Return _DataDeclaracaoImportacao
        End Get
        Set(ByVal value As Date)
            _DataDeclaracaoImportacao = value
        End Set
    End Property

    Public Property RegistroDeImportacao() As String
        Get
            Return _RegistroDeImportacao
        End Get
        Set(ByVal value As String)
            _RegistroDeImportacao = value
        End Set
    End Property

    Public Property DataRegistroDeImportacao() As Date
        Get
            Return _DataRegistroDeImportacao
        End Get
        Set(ByVal value As Date)
            _DataRegistroDeImportacao = value
        End Set
    End Property

    Public Property LocalEmbarqueImportacao() As String
        Get
            Return _LocalEmbarqueImportacao
        End Get
        Set(ByVal value As String)
            _LocalEmbarqueImportacao = value
        End Set
    End Property

    Public Property EstadoEmbarqueImportacao() As String
        Get
            Return _EstadoEmbarqueImportacao
        End Get
        Set(ByVal value As String)
            _EstadoEmbarqueImportacao = value
        End Set
    End Property

    Public Property DataEmbarqueImportacao() As DateTime
        Get
            Return _DataEmbarqueImportacao
        End Get
        Set(ByVal value As DateTime)
            _DataEmbarqueImportacao = value
        End Set
    End Property

    Public Property LocalDesembarqueImportacao() As String
        Get
            Return _LocalDesembarqueImportacao
        End Get
        Set(ByVal value As String)
            _LocalDesembarqueImportacao = value
        End Set
    End Property

    Public Property EstadoDesembarqueImportacao() As String
        Get
            Return _EstadoDesembarqueImportacao
        End Get
        Set(ByVal value As String)
            _EstadoDesembarqueImportacao = value
        End Set
    End Property

    Public Property DataDesembarqueImportacao() As DateTime
        Get
            Return _DataDesembarqueImportacao
        End Get
        Set(ByVal value As DateTime)
            _DataDesembarqueImportacao = value
        End Set
    End Property

    Public Property CodigoFabricante() As String
        Get
            Return _CodigoFabricante
        End Get
        Set(ByVal value As String)
            _CodigoFabricante = value
            _Fabricante = Nothing
        End Set
    End Property

    Public Property EndFabricante() As Integer
        Get
            Return _EndFabricante
        End Get
        Set(ByVal value As Integer)
            _EndFabricante = value
            _Fabricante = Nothing
        End Set
    End Property

    Public Property Fabricante() As Cliente
        Get
            If _Fabricante Is Nothing And _CodigoFabricante.Length > 0 Then _Fabricante = New Cliente(_CodigoFabricante, _EndFabricante)
            Return _Fabricante
        End Get
        Set(ByVal value As Cliente)
            _Fabricante = value
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

    Public Property NrFatura() As String
        Get
            Return _NrFatura
        End Get
        Set(ByVal value As String)
            _NrFatura = value
        End Set
    End Property

    Public Property ViaDeTransporte() As Integer
        Get
            Return _ViaDeTransporte
        End Get
        Set(ByVal value As Integer)
            _ViaDeTransporte = value
        End Set
    End Property

    Public Property TipoDeImportacao() As Integer
        Get
            Return _TipoDeImportacao
        End Get
        Set(ByVal value As Integer)
            _TipoDeImportacao = value
        End Set
    End Property

    Public Property ValorVAFRMM() As Decimal
        Get
            Return _ValorVAFRMM
        End Get
        Set(ByVal value As Decimal)
            _ValorVAFRMM = value
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
                sql = " Insert Into NotaFiscalXImportacao(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id," & vbCrLf & _
                      " NumeroDeclaracaoImportacao, DataDeclaracaoImportacao, RegistroDeImportacao, DataRegistroDeImportacao, " & vbCrLf & _
                      " LocalEmbarqueImportacao, EstadoEmbarqueImportacao, DataEmbarqueImportacao, " & vbCrLf & _
                      " LocalDesembarqueImportacao, EstadoDesembarqueImportacao, DataDesembarqueImportacao, " & vbCrLf & _
                      " Fabricante, EndFabricante, NumAtoConcessorio, DtaRegAtoConcessorio, DtaValidAtoConcessorio, NrFatura, ViaDeTransporte, TipoDeImportacao, ValorVAFRMM) " & vbCrLf & _
                      " Values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ",'" & NF.CodigoCliente & "'," & NF.EnderecoCliente & ",'" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & NF.Serie & "'," & NF.Codigo & "," & vbCrLf & _
                      "'" & Me.NumeroDeclaracaoImportacao & "'," & IIf(Me.NumeroDeclaracaoImportacao.Trim.Length = 0, "NULL", "'" & Me.DataDeclaracaoImportacao.ToString("yyyy-MM-dd") & "'") & ", " & vbCrLf & _
                      "'" & Me.RegistroDeImportacao & "'," & IIf(Me.RegistroDeImportacao.Trim.Length = 0, "NULL", "'" & Me.DataRegistroDeImportacao.ToString("yyyy-MM-dd") & "'") & ", " & vbCrLf & _
                      "'" & _LocalEmbarqueImportacao & "','" & _EstadoEmbarqueImportacao & "'," & IIf(Me._LocalEmbarqueImportacao.Trim.Length = 0, "NULL", "'" & Me._DataEmbarqueImportacao.ToString("yyyy-MM-dd") & "'") & ", " & vbCrLf & _
                      "'" & _LocalDesembarqueImportacao & "','" & _EstadoDesembarqueImportacao & "'," & IIf(Me._LocalDesembarqueImportacao.Trim.Length = 0, "NULL", "'" & Me._DataDesembarqueImportacao.ToString("yyyy-MM-dd") & "'") & " , " & vbCrLf & _
                      "'" & _CodigoFabricante & "'," & _EndFabricante & ", '" & Me.NumAtoConcessorio & "'," & IIf(Me.NumAtoConcessorio.Trim.Length = 0, "NULL,NULL", "'" & Me.DtaRegAtoConcessorio.ToString("yyyy-MM-dd") & "', '" & Me.DtaValidAtoConcessorio.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      ",'" & Me.NrFatura & "'," & Me.ViaDeTransporte & "," & Me.TipoDeImportacao & "," & Str(Me.ValorVAFRMM) & ")"

                Sqls.Add(sql)
            Case "U"
                sql = "Update NotaFiscalXImportacao set" & vbCrLf & _
                      "   NumeroDeclaracaoImportacao  ='" & Me.NumeroDeclaracaoImportacao & "'" & vbCrLf & _
                      "  ,DataDeclaracaoImportacao    = " & IIf(NumeroDeclaracaoImportacao.Length = 0, "NULL", "'" & Me.DataDeclaracaoImportacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,RegistroDeImportacao        ='" & Me.RegistroDeImportacao & "'" & vbCrLf & _
                      "  ,DataRegistroDeImportacao    = " & IIf(RegistroDeImportacao.Length = 0, "NULL", "'" & Me.DataRegistroDeImportacao.ToString("yyyy-MM-dd") & "'") & vbCrLf & _
                      "  ,LocalEmbarqueImportacao     ='" & Me.LocalEmbarqueImportacao & "'" & vbCrLf & _
                      "  ,EstadoEmbarqueImportacao    ='" & Me.EstadoEmbarqueImportacao & "'" & vbCrLf & _
                      "  ,DataEmbarqueImportacao      =" & IIf(Me.LocalEmbarqueImportacao.Trim.Length = 0, "NULL", "'" & Me.DataEmbarqueImportacao.ToString("yyyy-MM-dd") & "'") & "" & vbCrLf & _
                      "  ,LocalDesembarqueImportacao  ='" & Me.LocalDesembarqueImportacao & "'" & vbCrLf & _
                      "  ,EstadoDesembarqueImportacao ='" & Me.EstadoDesembarqueImportacao & "'" & vbCrLf & _
                      "  ,DataDesembarqueImportacao   =" & IIf(Me.LocalDesembarqueImportacao.Trim.Length = 0, "NULL", "'" & Me.DataDesembarqueImportacao.ToString("yyyy-MM-dd") & "'") & "" & vbCrLf & _
                      "  ,Fabricante                  ='" & Me.CodigoFabricante & "'" & vbCrLf & _
                      "  ,EndFabricante               =" & _EndFabricante & " " & vbCrLf & _
                      "  ,NumAtoConcessorio           ='" & Me.NumAtoConcessorio & "'" & vbCrLf & _
                      "  ,DtaRegAtoConcessorio        = " & IIf(Me.NumAtoConcessorio.Trim.Length = 0, "NULL", "'" & Me.DtaRegAtoConcessorio.ToString("yyyy-MM-dd") & "' ") & "" & vbCrLf & _
                      "  ,DtaValidAtoConcessorio      = " & IIf(Me.NumAtoConcessorio.Trim.Length = 0, "NULL", "'" & Me.DtaValidAtoConcessorio.ToString("yyyy-MM-dd") & "' ") & "" & vbCrLf & _
                      "  ,NrFatura                    ='" & Me.NrFatura & "'" & vbCrLf & _
                      "  ,ViaDeTransporte             =" & Me.ViaDeTransporte & vbCrLf & _
                      "  ,TipoDeImportacao            =" & Me.TipoDeImportacao & vbCrLf & _
                      "  ,ValorVAFRMM                 =" & Str(Me.ValorVAFRMM) & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalXImportacao" & vbCrLf & _
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