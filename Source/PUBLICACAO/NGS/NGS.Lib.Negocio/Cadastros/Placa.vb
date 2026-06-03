Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class Placa
    Implements IBaseEntity

#Region "Variáveis locais"

    Private _Placa01 As String = ""
    Private _ViaDeTransporte As Integer
    Private _ViaDeTransporteDetalhes As ViaDeTransporte
    Private _TipoDeVeiculo As Integer
    Private _TipoDeVeiculoDetalhes As TipoDeVeiculo
    Private _Placa02 As String = ""
    Private _Placa03 As String = ""
    Private _Placa04 As String = ""
    Private _CidadePlaca01 As String = ""
    Private _EstadoPlaca01 As String = ""
    Private _EstadoPlaca01Detalhes As Estado
    Private _Habilitacao As String = ""
    Private _CpfMotorista As String = ""
    Private _EndCpfMotorista As Integer
    Private _Motorista As Cliente
    Private _Restricao As String = ""
    Private _Observacao As String = ""
    Private _CidadePlaca02 As String = ""
    Private _EstadoPlaca02 As String = ""
    Private _Estado02Detalhes As Estado
    Private _CidadePlaca03 As String = ""
    Private _EstadoPlaca03 As String = ""
    Private _Estado03Detalhes As Estado
    Private _CidadePlaca04 As String = ""
    Private _EstadoPlaca04 As String = ""
    Private _Estado04Detalhes As Estado
    Private _IUD As String
    Private _CodigoProprietario01 As String = ""
    Private _EndProprietario01 As Integer
    Private _CodigoProprietario02 As String = ""
    Private _EndProprietario02 As Integer
    Private _CodigoProprietario03 As String = ""
    Private _EndProprietario03 As Integer
    Private _CodigoProprietario04 As String = ""
    Private _EndProprietario04 As Integer
    Private _RNTRCPlaca01 As String = ""
    Private _RNTRCPlaca02 As String = ""
    Private _RNTRCPlaca03 As String = ""
    Private _RNTRCPlaca04 As String = ""
    Private _Proprietario01 As Cliente
    Private _Proprietario02 As Cliente
    Private _Proprietario03 As Cliente
    Private _Proprietario04 As Cliente
#End Region

#Region "Campos públicos"

    Public Erro As Exception

#End Region

#Region "Construtores"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Placa As String)
        Me.Placa01 = Placa
        Me.Selecionar()
    End Sub

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

    Public Property Placa01() As String
        Get
            Return _Placa01
        End Get
        Set(ByVal value As String)
            _Placa01 = value
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

    Public Property ViaDeTransporteDetalhes() As ViaDeTransporte
        Get
            If _ViaDeTransporteDetalhes Is Nothing AndAlso _ViaDeTransporte > 0 Then _ViaDeTransporteDetalhes = New ViaDeTransporte(_ViaDeTransporte)
            Return _ViaDeTransporteDetalhes
        End Get
        Set(ByVal value As ViaDeTransporte)
            _ViaDeTransporteDetalhes = value
        End Set
    End Property

    Public Property TipoDeVeiculo() As Integer
        Get
            Return _TipoDeVeiculo
        End Get
        Set(ByVal value As Integer)
            _TipoDeVeiculo = value
        End Set
    End Property

    Public Property TipoDeVeiculoDetalhes() As TipoDeVeiculo
        Get
            If _TipoDeVeiculoDetalhes Is Nothing AndAlso _TipoDeVeiculo > 0 Then _TipoDeVeiculoDetalhes = New TipoDeVeiculo(_TipoDeVeiculo)
            Return _TipoDeVeiculoDetalhes
        End Get
        Set(ByVal value As TipoDeVeiculo)
            _TipoDeVeiculoDetalhes = value
        End Set
    End Property

    Public Property Placa02() As String
        Get
            Return _Placa02
        End Get
        Set(ByVal value As String)
            _Placa02 = value
        End Set
    End Property

    Public Property Placa03() As String
        Get
            Return _Placa03
        End Get
        Set(ByVal value As String)
            _Placa03 = value
        End Set
    End Property

    Public Property Placa04() As String
        Get
            Return _Placa04
        End Get
        Set(ByVal value As String)
            _Placa04 = value
        End Set
    End Property

    Public Property CidadePlaca01() As String
        Get
            Return _CidadePlaca01
        End Get
        Set(ByVal value As String)
            _CidadePlaca01 = value
        End Set
    End Property

    Public Property EstadoPlaca01() As String
        Get
            Return _EstadoPlaca01
        End Get
        Set(ByVal value As String)
            _EstadoPlaca01 = value
        End Set
    End Property

    Public Property EstadoPlaca01Detalhes() As Estado
        Get
            If _EstadoPlaca01Detalhes Is Nothing AndAlso Not String.IsNullOrEmpty(_EstadoPlaca01) Then _EstadoPlaca01Detalhes = New Estado(_EstadoPlaca01)
            Return _EstadoPlaca01Detalhes
        End Get
        Set(ByVal value As Estado)
            _EstadoPlaca01Detalhes = value
        End Set
    End Property

    Public Property Habilitacao() As String
        Get
            Return _Habilitacao
        End Get
        Set(ByVal value As String)
            _Habilitacao = value
        End Set
    End Property

    Public Property CpfMotorista() As String
        Get
            Return _CpfMotorista
        End Get
        Set(ByVal value As String)
            _CpfMotorista = value
            _Motorista = Nothing
        End Set
    End Property

    Public Property EndCpfMotorista() As Integer
        Get
            Return _EndCpfMotorista
        End Get
        Set(ByVal value As Integer)
            _EndCpfMotorista = value
        End Set
    End Property

    Public Property Motorista() As Cliente
        Get
            If _Motorista Is Nothing AndAlso Not String.IsNullOrEmpty(Me.CpfMotorista) Then _Motorista = New Cliente(Me.CpfMotorista, Me.EndCpfMotorista)
            Return _Motorista
        End Get
        Set(ByVal value As Cliente)
            _Motorista = value
        End Set
    End Property

    Public Property Restricao() As String
        Get
            Return _Restricao
        End Get
        Set(ByVal value As String)
            _Restricao = value
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

    Public Property CidadePlaca02() As String
        Get
            Return _CidadePlaca02
        End Get
        Set(ByVal value As String)
            _CidadePlaca02 = value
        End Set
    End Property

    Public Property EstadoPlaca02() As String
        Get
            Return _EstadoPlaca02
        End Get
        Set(ByVal value As String)
            _EstadoPlaca02 = value
        End Set
    End Property

    Public Property Estado02Detalhes() As Estado
        Get
            If _Estado02Detalhes Is Nothing AndAlso Not String.IsNullOrEmpty(_EstadoPlaca02) Then _Estado02Detalhes = New Estado(_EstadoPlaca02)
            Return _Estado02Detalhes
        End Get
        Set(ByVal value As Estado)
            _Estado02Detalhes = value
        End Set
    End Property

    Public Property CidadePlaca03() As String
        Get
            Return _CidadePlaca03
        End Get
        Set(ByVal value As String)
            _CidadePlaca03 = value
        End Set
    End Property

    Public Property EstadoPlaca03() As String
        Get
            Return _EstadoPlaca03
        End Get
        Set(ByVal value As String)
            _EstadoPlaca03 = value
        End Set
    End Property

    Public Property Estado03Detalhes() As Estado
        Get
            If _Estado03Detalhes Is Nothing AndAlso Not String.IsNullOrEmpty(_EstadoPlaca03) Then _Estado03Detalhes = New Estado(_EstadoPlaca03)
            Return _Estado03Detalhes
        End Get
        Set(ByVal value As Estado)
            _Estado03Detalhes = value
        End Set
    End Property

    Public Property CidadePlaca04() As String
        Get
            Return _CidadePlaca04
        End Get
        Set(ByVal value As String)
            _CidadePlaca04 = value
        End Set
    End Property

    Public Property EstadoPlaca04() As String
        Get
            Return _EstadoPlaca04
        End Get
        Set(ByVal value As String)
            _EstadoPlaca04 = value
        End Set
    End Property

    Public Property Estado04Detalhes() As Estado
        Get
            If _Estado04Detalhes Is Nothing AndAlso Not String.IsNullOrEmpty(_EstadoPlaca04) Then _Estado04Detalhes = New Estado(_EstadoPlaca04)
            Return _Estado04Detalhes
        End Get
        Set(ByVal value As Estado)
            _Estado04Detalhes = value
        End Set
    End Property

    Public Property CodigoProprietario01() As String
        Get
            Return _CodigoProprietario01
        End Get
        Set(ByVal value As String)
            _CodigoProprietario01 = value
        End Set
    End Property

    Public Property EndProprietario01() As Integer
        Get
            Return _EndProprietario01
        End Get
        Set(ByVal value As Integer)
            _EndProprietario01 = value
        End Set
    End Property

    Public Property Proprietario01() As Cliente
        Get
            If _Proprietario01 Is Nothing AndAlso Not String.IsNullOrEmpty(Me.CodigoProprietario01) Then _Proprietario01 = New Cliente(Me.CodigoProprietario01, Me.EndProprietario01)
            Return _Proprietario01
        End Get
        Set(ByVal value As Cliente)
            _Proprietario01 = value
        End Set
    End Property

    Public Property CodigoProprietario02() As String
        Get
            Return _CodigoProprietario02
        End Get
        Set(ByVal value As String)
            _CodigoProprietario02 = value
        End Set
    End Property

    Public Property EndProprietario02() As Integer
        Get
            Return _EndProprietario02
        End Get
        Set(ByVal value As Integer)
            _EndProprietario02 = value
        End Set
    End Property

    Public Property Proprietario02() As Cliente
        Get
            If _Proprietario02 Is Nothing AndAlso Not String.IsNullOrEmpty(Me.CodigoProprietario02) Then _Proprietario02 = New Cliente(Me.CodigoProprietario02, Me.EndProprietario02)
            Return _Proprietario02
        End Get
        Set(ByVal value As Cliente)
            _Proprietario02 = value
        End Set
    End Property

    Public Property CodigoProprietario03() As String
        Get
            Return _CodigoProprietario03
        End Get
        Set(ByVal value As String)
            _CodigoProprietario03 = value
        End Set
    End Property

    Public Property EndProprietario03() As Integer
        Get
            Return _EndProprietario03
        End Get
        Set(ByVal value As Integer)
            _EndProprietario03 = value
        End Set
    End Property

    Public Property Proprietario03() As Cliente
        Get
            If _Proprietario03 Is Nothing AndAlso Not String.IsNullOrEmpty(Me.CodigoProprietario03) Then _Proprietario03 = New Cliente(Me.CodigoProprietario03, Me.EndProprietario03)
            Return _Proprietario03
        End Get
        Set(ByVal value As Cliente)
            _Proprietario03 = value
        End Set
    End Property

    Public Property CodigoProprietario04() As String
        Get
            Return _CodigoProprietario04
        End Get
        Set(ByVal value As String)
            _CodigoProprietario04 = value
        End Set
    End Property

    Public Property EndProprietario04() As Integer
        Get
            Return _EndProprietario04
        End Get
        Set(ByVal value As Integer)
            _EndProprietario04 = value
        End Set
    End Property

    Public Property Proprietario04() As Cliente
        Get
            If _Proprietario04 Is Nothing AndAlso Not String.IsNullOrEmpty(Me.CodigoProprietario04) Then _Proprietario04 = New Cliente(Me.CodigoProprietario04, Me.EndProprietario04)
            Return _Proprietario04
        End Get
        Set(ByVal value As Cliente)
            _Proprietario04 = value
        End Set
    End Property

    Public Property RNTRCPlaca01() As String
        Get
            Return _RNTRCPlaca01
        End Get
        Set(ByVal value As String)
            _RNTRCPlaca01 = value
        End Set
    End Property

    Public Property RNTRCPlaca02() As String
        Get
            Return _RNTRCPlaca02
        End Get
        Set(ByVal value As String)
            _RNTRCPlaca02 = value
        End Set
    End Property

    Public Property RNTRCPlaca03() As String
        Get
            Return _RNTRCPlaca03
        End Get
        Set(ByVal value As String)
            _RNTRCPlaca03 = value
        End Set
    End Property

    Public Property RNTRCPlaca04() As String
        Get
            Return _RNTRCPlaca04
        End Get
        Set(ByVal value As String)
            _RNTRCPlaca04 = value
        End Set
    End Property

#End Region

#Region "Métodos"

    Public Function Selecionar() As Boolean
        Dim db As New AcessaBanco()
        Try
            Dim sql As String = ""
            sql = "SELECT   Placa_Id AS Placa, ISNULL(ViaTransporte_Id, 0) AS ViaDeTransporte, " & vbCrLf & _
                  "         ISNULL(TipoVeiculo_Id, 0) AS TipoDeVeiculo, Placa01, Placa02, Placa03, CidadePlaca, " & vbCrLf & _
                  "         EstadoPlaca, Habilitacao, CpfMotorista, ISNULL(EndCpfMotorista, 0) AS EndCpfMotorista, " & vbCrLf & _
                  "         isnull(Restricao, 'N') as Restricao, isnull(Observacao, '') as Observacao, CidadePlaca01, EstadoPlaca01, CidadePlaca02, " & vbCrLf & _
                  "         EstadoPlaca02, CidadePlaca03, EstadoPlaca03, ISNULL(RNTRCPlaca,'') AS RNTRCPlaca, ISNULL(RNTRCPlaca01,'') AS RNTRCPlaca01, " & vbCrLf & _
                  "         ISNULL(RNTRCPlaca02,'') AS RNTRCPlaca02, ISNULL(RNTRCPlaca03,'') AS RNTRCPlaca03, " & vbCrLf & _
                  "         isnull(Proprietario, '') AS Proprietario, isnull(EndProprietario,0) AS EndProprietario, " & vbCrLf & _
                  "         isnull(Proprietario01, '') AS Proprietario1, isnull(EndProprietario01,0) AS EndProprietario1, " & vbCrLf & _
                  "         isnull(Proprietario02, '') AS Proprietario2, isnull(EndProprietario02,0) AS EndProprietario2, " & vbCrLf & _
                  "         isnull(Proprietario03, '') AS Proprietario3, isnull(EndProprietario03,0) AS EndProprietario3 " & vbCrLf & _
                  " FROM  Placas " & vbCrLf & _
                  "     WHERE Placa_Id = '" & Me.Placa01 & "'" & vbCrLf

            Me.IUD = "I"
            Dim ds As DataSet = db.ConsultaDataSet(sql, "Placas")
            If ds.Tables(0).Rows.Count > 0 Then
                Dim row As DataRow = ds.Tables(0).Rows(0)
                Me.IUD = "U"
                Me.Placa01 = row("Placa").ToString()
                Me.ViaDeTransporte = Convert.ToInt32(row("ViaDeTransporte"))
                Me.TipoDeVeiculo = Convert.ToInt32(row("TipoDeVeiculo"))
                Me.Placa02 = row("Placa01").ToString()
                Me.Placa03 = row("Placa02").ToString()
                Me.Placa04 = row("Placa03").ToString()
                Me.CidadePlaca01 = row("CidadePlaca").ToString()
                Me.EstadoPlaca01 = row("EstadoPlaca").ToString()
                Me.Habilitacao = row("Habilitacao").ToString()
                Me.CpfMotorista = row("CpfMotorista").ToString()
                Me.EndCpfMotorista = Convert.ToInt32(row("EndCpfMotorista"))
                Me.Restricao = row("Restricao").ToString()
                Me.Observacao = row("Observacao").ToString()
                Me.CidadePlaca02 = row("CidadePlaca01").ToString()
                Me.EstadoPlaca02 = row("EstadoPlaca01").ToString()
                Me.CidadePlaca03 = row("CidadePlaca02").ToString()
                Me.EstadoPlaca03 = row("EstadoPlaca02").ToString()
                Me.CidadePlaca04 = row("CidadePlaca03").ToString()
                Me.EstadoPlaca04 = row("EstadoPlaca03").ToString()
                Me.RNTRCPlaca01 = row("RNTRCPlaca")
                Me.RNTRCPlaca02 = row("RNTRCPlaca01")
                Me.RNTRCPlaca03 = row("RNTRCPlaca02")
                Me.RNTRCPlaca04 = row("RNTRCPlaca03")
                Me.CodigoProprietario01 = row("Proprietario")
                Me.EndProprietario01 = row("EndProprietario")
                Me.CodigoProprietario02 = row("Proprietario1")
                Me.EndProprietario02 = row("EndProprietario1")
                Me.CodigoProprietario03 = row("Proprietario2")
                Me.EndProprietario03 = row("EndProprietario2")
                Me.CodigoProprietario04 = row("Proprietario3")
                Me.EndProprietario04 = row("EndProprietario3")
                Return True
            Else
                Return False
            End If
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            db = Nothing
        End Try
    End Function

    Public Function TemLaudo() As Boolean
        Dim db As New AcessaBanco()

        Dim sql As String = ""
        sql = "SELECT   Placa From Pesagem " & vbCrLf & _
              "WHERE Placa = '" & Me.Placa01 & "'"

        Dim ds As DataSet = db.ConsultaDataSet(sql, "Laudo")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function TemNota() As Boolean
        Dim db As New AcessaBanco()

        Dim sql As String = ""
        sql = "SELECT   Placa From notasfiscaisXtransportadores " & vbCrLf & _
              "WHERE Placa = '" & Me.Placa01 & "'"

        Dim ds As DataSet = db.ConsultaDataSet(sql, "NotaFiscal")

        If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return True
        Else
            Return False
        End If
    End Function


    Public Function Salvar() As Boolean
        Dim db As New AcessaBanco()
        Dim Sqls As New ArrayList
        SalvarSql(Sqls)
        Return db.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String = String.Empty
        Select Case IUD
            Case "I"
                sql = "INSERT INTO  Placas (Placa_Id, ViaTransporte_Id, TipoVeiculo_Id, Placa01, Placa02, Placa03, CidadePlaca, EstadoPlaca, Habilitacao, " & vbCrLf & _
                      "                     CpfMotorista, EndCpfMotorista, Restricao, Observacao, CidadePlaca01, EstadoPlaca01, CidadePlaca02, " & vbCrLf & _
                      "                     EstadoPlaca02, CidadePlaca03, EstadoPlaca03, RNTRCPlaca, RNTRCPlaca01, RNTRCPlaca02, RNTRCPlaca03, Proprietario, EndProprietario, " & vbCrLf & _
                      "                     Proprietario01, EndProprietario01, Proprietario02, EndProprietario02, Proprietario03, EndProprietario03) " & vbCrLf & _
                      "             VALUES ('" & Me.Placa01 & "', " & Me.ViaDeTransporte & ", " & Me.TipoDeVeiculo & ", '" & Me.Placa02 & "', '" & Me.Placa03 & "', '" & Me.Placa04 & "', " & vbCrLf & _
                      "                     '" & Me.CidadePlaca01 & "', '" & Me.EstadoPlaca01 & "', '" & Me.Habilitacao & "', '" & Me.CpfMotorista & "', " & Me.EndCpfMotorista & ", '" & Me.Restricao & "', " & vbCrLf & _
                      "                     '" & Me.Observacao & "', '" & Me.CidadePlaca02 & "', '" & Me.EstadoPlaca02 & "', '" & Me.CidadePlaca03 & "', '" & Me.EstadoPlaca03 & "', '" & Me.CidadePlaca04 & "', " & vbCrLf & _
                      "                     '" & Me.EstadoPlaca04 & "', '" & Me.RNTRCPlaca01 & "', '" & Me.RNTRCPlaca02 & "', '" & Me.RNTRCPlaca03 & "', '" & Me.RNTRCPlaca04 & "', " & vbCrLf & _
                      "                     '" & Me.CodigoProprietario01 & "', " & Me.EndProprietario01 & ", '" & Me.CodigoProprietario02 & "', " & Me.EndProprietario02 & ", " & vbCrLf & _
                      "                     '" & Me.CodigoProprietario03 & "', " & Me.EndProprietario03 & ", '" & Me.CodigoProprietario04 & "', " & Me.EndProprietario04 & ") " & vbCrLf
            Case "U"
                sql = " UPDATE Placas SET ViaTransporte_Id    =  " & Me.ViaDeTransporte & ", " & vbCrLf & _
                      "                   TipoVeiculo_Id      =  " & Me.TipoDeVeiculo & ", " & vbCrLf & _
                      "                   Placa01             = '" & Me.Placa02 & "', " & vbCrLf & _
                      "                   Placa02             = '" & Me.Placa03 & "', " & vbCrLf & _
                      "                   Placa03             = '" & Me.Placa04 & "', " & vbCrLf & _
                      "                   CidadePlaca         = '" & Me.CidadePlaca01 & "', " & vbCrLf & _
                      "                   EstadoPlaca         = '" & Me.EstadoPlaca01 & "', " & vbCrLf & _
                      "                   Habilitacao         = '" & Me.Habilitacao & "', " & vbCrLf & _
                      "                   CpfMotorista        = '" & Me.CpfMotorista & "', " & vbCrLf & _
                      "                   EndCpfMotorista     =  " & Me.EndCpfMotorista & ", " & vbCrLf & _
                      "                   Restricao           = '" & Me.Restricao & "', " & vbCrLf & _
                      "                   Observacao          = '" & Me.Observacao & "', " & vbCrLf & _
                      "                   CidadePlaca01       = '" & Me.CidadePlaca02 & "', " & vbCrLf & _
                      "                   EstadoPlaca01       = '" & Me.EstadoPlaca02 & "', " & vbCrLf & _
                      "                   CidadePlaca02       = '" & Me.CidadePlaca03 & "', " & vbCrLf & _
                      "                   EstadoPlaca02       = '" & Me.EstadoPlaca03 & "', " & vbCrLf & _
                      "                   CidadePlaca03       = '" & Me.CidadePlaca04 & "', " & vbCrLf & _
                      "                   EstadoPlaca03       = '" & Me.EstadoPlaca04 & "', " & vbCrLf & _
                      "                   RNTRCPlaca          = '" & Me.RNTRCPlaca01 & "', " & vbCrLf & _
                      "                   RNTRCPlaca01        = '" & Me.RNTRCPlaca02 & "', " & vbCrLf & _
                      "                   RNTRCPlaca02        = '" & Me.RNTRCPlaca03 & "', " & vbCrLf & _
                      "                   RNTRCPlaca03        = '" & Me.RNTRCPlaca04 & "', " & vbCrLf & _
                      "                   Proprietario        = '" & Me.CodigoProprietario01 & "', " & vbCrLf & _
                      "                   EndProprietario     =  " & Me.EndProprietario01 & ", " & vbCrLf & _
                      "                   Proprietario01      = '" & Me.CodigoProprietario02 & "', " & vbCrLf & _
                      "                   EndProprietario01   =  " & Me.EndProprietario02 & ", " & vbCrLf & _
                      "                   Proprietario02      = '" & Me.CodigoProprietario03 & "', " & vbCrLf & _
                      "                   EndProprietario02   =  " & Me.EndProprietario03 & ", " & vbCrLf & _
                      "                   Proprietario03      = '" & Me.CodigoProprietario04 & "', " & vbCrLf & _
                      "                   EndProprietario03   =  " & Me.EndProprietario04 & vbCrLf & _
                      "         WHERE Placa_Id = '" & Me.Placa01 & "'" & vbCrLf
            Case "D"
                sql = "DELETE FROM Placas " & vbCrLf & _
                      "  WHERE Placa_Id = '" & Me.Placa01 & "'" & vbCrLf
        End Select
        Sqls.Add(sql)
        GravaHistorico(Sqls)
    End Sub

    Private Sub GravaHistorico(ByRef Sqls As ArrayList)
        Dim sql As String = String.Empty & vbCrLf

        sql = " INSERT INTO PlacasXHistorico" & vbCrLf & _
         "                    (Placa_Id, Data_Id, ViaTransporte_Id, TipoVeiculo_Id, Motorista, EndMotorista, Habilitacao, Proprietario, EndProprietario, RNTCPlaca, Observacao, EstadoPlaca, " & vbCrLf & _
         "                     CidadePlaca, Restricao, Placa01, ProprietarioPlaca01, EndProprietarioPlaca01, RNTRCPlaca01, EstadoPlaca01, CidadePlaca01, Placa02, ProprietarioPlaca02,       " & vbCrLf & _
         "                     EndProprietarioPlaca02, RNTRCPlaca02, EstadoPlaca02, CidadePlaca02, Placa03, ProprietarioPlaca03, EndProprietarioPlaca03, RNTRCPlaca03, EstadoPlaca03,        " & vbCrLf & _
         "                     CidadePlaca03, UsuarioAlteracao, IUD)                                                                                                                                          " & vbCrLf & _
         "            VALUES ('" & Me.Placa01 & "', GETDATE(), " & Me.ViaDeTransporte & ", " & Me.TipoDeVeiculo & ", '" & Me.CpfMotorista & "', " & Me.EndCpfMotorista & ", '" & Me.Habilitacao & "', " & vbCrLf & _
         "                    '" & Me.CodigoProprietario01 & "'," & Me.EndProprietario01 & ", '" & Me.RNTRCPlaca01 & "', '" & Me.Observacao & "', '" & Me.EstadoPlaca01 & "', '" & Me.CidadePlaca01 & "', " & vbCrLf & _
         "                    '" & Me.Restricao & "', '" & Me.Placa02 & "', '" & Me.CodigoProprietario02 & "', " & Me.EndProprietario02 & ", '" & Me.RNTRCPlaca02 & "', '" & Me.EstadoPlaca02 & "', " & vbCrLf & _
         "                    '" & Me.CidadePlaca02 & "', '" & Me.Placa03 & "', '" & Me.CodigoProprietario03 & "', " & Me.EndProprietario03 & ", '" & Me.RNTRCPlaca03 & "', '" & Me.EstadoPlaca03 & "', " & vbCrLf & _
         "                    '" & Me.CidadePlaca03 & "', '" & Me.Placa04 & "', '" & Me.CodigoProprietario04 & "', " & Me.EndProprietario04 & ", '" & Me.RNTRCPlaca04 & "', '" & Me.EstadoPlaca04 & "', " & vbCrLf & _
         "                    '" & Me.CidadePlaca04 & "', '" & UsuarioServidor.NomeUsuario & "', '" & Me.IUD & "')" & vbCrLf

        Sqls.Add(sql)
    End Sub
#End Region

End Class