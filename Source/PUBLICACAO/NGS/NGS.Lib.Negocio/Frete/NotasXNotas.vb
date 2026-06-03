Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotasXNotas
    Inherits List(Of NotasXNotas)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal PNotaOrigem As NotaFiscal, Optional ByVal SomenteNotasDeOrigem As Boolean = True, Optional ByVal BuscarPorOrigem As Boolean = False)
        If SomenteNotasDeOrigem Then
            CargaNormal(PNotaOrigem, BuscarPorOrigem)
        Else
            CargaNotaOriUN(PNotaOrigem)
        End If
    End Sub

    Private Sub CargaNormal(ByVal PNotaOrigem As NotaFiscal, ByVal BuscarPorOrigem As Boolean)
        Dim Sql As String
        Dim ds As New DataSet
        Dim dr As DataRow

        Sql = "SELECT NxN.Empresa_Id,NxN.EndEmpresa_Id,NxN.Cliente_Id,NxN.EndCliente_Id,NxN.EntradaSaida_Id,NxN.Serie_Id,NxN.Nota_Id, " & vbCrLf & _
              "       NxN.OrigemEmpresa_Id,NxN.OrigemEndEmpresa_Id,NxN.OrigemCliente_Id,NxN.OrigemEndCliente_Id,NxN.OrigemEntradaSaida_Id,NxN.OrigemSerie_Id, " & vbCrLf & _
              "       NxN.OrigemNota_Id,Sum(NFxI.PesoFiscal) as PesoFiscal,SUM(isnull(R.PesoBruto,0)) AS PesoBrutoRomaneio " & vbCrLf & _
              "  FROM NotasXNotas NxN " & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
              "    ON NxN.OrigemEmpresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
              "   AND NxN.OrigemEndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
              "   AND NxN.OrigemCliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
              "   AND NxN.OrigemSerie_Id        = NFxI.Serie_Id " & vbCrLf & _
              "   AND NxN.OrigemNota_Id         = NFxI.Nota_Id " & vbCrLf & _
              "  LEFT JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
              "    ON NxN.OrigemEmpresa_Id      = nfxr.Empresa_Id " & vbCrLf & _
              "   AND NxN.OrigemEndEmpresa_Id   = nfxr.EndEmpresa_Id " & vbCrLf & _
              "   AND NxN.OrigemCliente_Id      = nfxr.Cliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEndCliente_Id   = nfxr.EndCliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEntradaSaida_Id = nfxr.EntradaSaida_Id " & vbCrLf & _
              "   AND NxN.OrigemSerie_Id        = nfxr.Serie_Id " & vbCrLf & _
              "   AND NxN.OrigemNota_Id         = nfxr.Nota_Id " & vbCrLf & _
              "  LEFT JOIN  Romaneios R " & vbCrLf & _
              "    ON nfxr.Empresa_Id    = R.Empresa_Id " & vbCrLf & _
              "   AND nfxr.EndEmpresa_Id = R.EndEmpresa_Id " & vbCrLf & _
              "   AND nfxr.Romaneio_Id   = R.Romaneio_Id " & vbCrLf

        If Not BuscarPorOrigem Then
            Sql &= " WHERE NxN.Nota_Id         ='" & PNotaOrigem.Codigo & "' " & vbCrLf & _
                   "   AND NxN.Empresa_Id      ='" & PNotaOrigem.CodigoEmpresa & "'" & vbCrLf & _
                   "   AND NxN.EndEmpresa_Id   = " & PNotaOrigem.EnderecoEmpresa & vbCrLf & _
                   "   AND NxN.Cliente_Id      ='" & PNotaOrigem.CodigoCliente & "'" & vbCrLf & _
                   "   AND NxN.EndCliente_Id   ='" & PNotaOrigem.EnderecoCliente & "' " & vbCrLf & _
                   "   AND NxN.EntradaSaida_Id ='" & IIf(PNotaOrigem.EntradaSaida = 1, "S", "E") & "'" & vbCrLf & _
                   "   AND NxN.Serie_Id        ='" & PNotaOrigem.Serie & "'" & vbCrLf
        Else
            Sql &= " WHERE NxN.OrigemNota_Id         ='" & PNotaOrigem.Codigo & "' " & vbCrLf & _
                   "   AND NxN.OrigemEmpresa_Id      ='" & PNotaOrigem.CodigoEmpresa & "'" & vbCrLf & _
                   "   AND NxN.OrigemEndEmpresa_Id   = " & PNotaOrigem.EnderecoEmpresa & vbCrLf & _
                   "   AND NxN.OrigemCliente_Id      ='" & PNotaOrigem.CodigoCliente & "'" & vbCrLf & _
                   "   AND NxN.OrigemEndCliente_Id   ='" & PNotaOrigem.EnderecoCliente & "' " & vbCrLf & _
                   "   AND NxN.OrigemEntradaSaida_Id ='" & IIf(PNotaOrigem.EntradaSaida = 1, "S", "E") & "'" & vbCrLf & _
                   "   AND NxN.OrigemSerie_Id        ='" & PNotaOrigem.Serie & "'" & vbCrLf
        End If
        Sql &= " GROUP BY NxN.Empresa_Id,NxN.EndEmpresa_Id,NxN.Cliente_Id,NxN.EndCliente_Id,NxN.EntradaSaida_Id,NxN.Serie_Id,NxN.Nota_Id, " & vbCrLf & _
               "       NxN.OrigemEmpresa_Id,NxN.OrigemEndEmpresa_Id,NxN.OrigemCliente_Id,NxN.OrigemEndCliente_Id,NxN.OrigemEntradaSaida_Id,NxN.OrigemSerie_Id," & vbCrLf & _
               "       NxN.OrigemNota_Id " & vbCrLf

        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(Sql, "NotasXNotas")

        For Each dr In ds.Tables(0).Rows
            Dim NxN As New NotasXNotas
            NxN.EmpresaCnpj = dr("Empresa_Id")
            NxN.EndEmpresa = dr("EndEmpresa_Id")
            NxN.ClienteCnpj = dr("Cliente_Id")
            NxN.EndCliente = dr("EndCliente_Id")
            NxN.EntradaSaida = dr("EntradaSaida_Id")
            NxN.Serie = dr("Serie_Id")
            NxN.NumeroNota = dr("Nota_Id")
            NxN.OrigemEmpresaCnpj = dr("OrigemEmpresa_Id")
            NxN.OrigemEndEmpresa = dr("OrigemEndEmpresa_Id")
            NxN.OrigemClienteCnpj = dr("OrigemCliente_Id")
            NxN.OrigemEndCliente = dr("OrigemEndCliente_Id")
            NxN.OrigemEntradaSaida = dr("OrigemEntradaSaida_Id")
            NxN.OrigemSerie = dr("OrigemSerie_Id")
            NxN.OrigemNota = dr("OrigemNota_Id")
            NxN.PesoFiscal = dr("PesoFiscal")
            NxN.PesoBrutoRomaneio = dr("PesoBrutoRomaneio")
            Me.Add(NxN)
        Next
    End Sub

    Private Sub CargaNotaOriUN(ByVal PNotaOrigem As NotaFiscal)
        Dim Sql As String
        Dim ds As New DataSet
        Dim dr As DataRow
        Sql = "SELECT NxN.Empresa_Id,NxN.EndEmpresa_Id,NxN.Cliente_Id,NxN.EndCliente_Id,NxN.EntradaSaida_Id,NxN.Serie_Id,NxN.Nota_Id, " & vbCrLf & _
              "       NxN.OrigemEmpresa_Id,NxN.OrigemEndEmpresa_Id,NxN.OrigemCliente_Id,NxN.OrigemEndCliente_Id,NxN.OrigemEntradaSaida_Id,NxN.OrigemSerie_Id," & vbCrLf & _
              "       NxN.OrigemNota_Id,Sum(NFxI.PesoFiscal) as PesoFiscal, SUM(isnull(R.PesoBruto,0)) AS PesoFiscalRomaneio " & vbCrLf & _
              "  FROM NotasFiscais NF " & vbCrLf & _
              " INNER JOIN NotasXNotas NxN " & vbCrLf & _
              "    ON NF.Empresa_Id      = NxN.Empresa_Id " & vbCrLf & _
              "   AND NF.EndEmpresa_Id   = NxN.EndEmpresa_Id " & vbCrLf & _
              "   AND NF.Cliente_Id      = NxN.Cliente_Id " & vbCrLf & _
              "   AND NF.EndCliente_Id   = NxN.EndCliente_Id " & vbCrLf & _
              "   AND NF.EntradaSaida_Id = NxN.EntradaSaida_Id " & vbCrLf & _
              "   AND NF.Serie_Id        = NxN.Serie_Id " & vbCrLf & _
              "   AND NF.Nota_Id         = NxN.Nota_Id " & vbCrLf & _
              " INNER JOIN NotasXNotas as UN " & vbCrLf & _
              "    ON NxN.OrigemEmpresa_Id      = UN.Empresa_Id " & vbCrLf & _
              "   AND NxN.OrigemEndEmpresa_Id   = UN.EndEmpresa_Id " & vbCrLf & _
              "   AND NxN.OrigemCliente_Id      = UN.Cliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEndCliente_Id   = UN.EndCliente_Id " & vbCrLf & _
              "   AND NxN.OrigemEntradaSaida_Id = UN.EntradaSaida_Id " & vbCrLf & _
              "   AND NxN.OrigemSerie_Id        = UN.Serie_Id " & vbCrLf & _
              "   AND NxN.OrigemNota_Id         = UN.Nota_Id " & vbCrLf & _
              " INNER JOIN NotasFiscaisXItens NFxI " & vbCrLf & _
              "    ON UN.OrigemEmpresa_Id      = NFxI.Empresa_Id " & vbCrLf & _
              "   AND UN.OrigemEndEmpresa_Id   = NFxI.EndEmpresa_Id " & vbCrLf & _
              "   AND UN.OrigemCliente_Id      = NFxI.Cliente_Id " & vbCrLf & _
              "   AND UN.OrigemEndCliente_Id   = NFxI.EndCliente_Id " & vbCrLf & _
              "   AND UN.OrigemEntradaSaida_Id = NFxI.EntradaSaida_Id " & vbCrLf & _
              "   AND UN.OrigemSerie_Id        = NFxI.Serie_Id " & vbCrLf & _
              "   AND UN.OrigemNota_Id         = NFxI.Nota_Id " & vbCrLf & _
              "  LEFT JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf & _
              "    ON UN.OrigemEmpresa_Id      = nfxr.Empresa_Id " & vbCrLf & _
              "   AND UN.OrigemEndEmpresa_Id   = nfxr.EndEmpresa_Id " & vbCrLf & _
              "   AND UN.OrigemCliente_Id      = nfxr.Cliente_Id " & vbCrLf & _
              "   AND UN.OrigemEndCliente_Id   = nfxr.EndCliente_Id " & vbCrLf & _
              "   AND UN.OrigemEntradaSaida_Id = nfxr.EntradaSaida_Id " & vbCrLf & _
              "   AND UN.OrigemSerie_Id        = nfxr.Serie_Id " & vbCrLf & _
              "   AND UN.OrigemNota_Id         = nfxr.Nota_Id " & vbCrLf & _
              "  LEFT JOIN  Romaneios R " & vbCrLf & _
              "    ON nfxr.Empresa_Id    = R.Empresa_Id " & vbCrLf & _
              "   AND nfxr.EndEmpresa_Id = R.EndEmpresa_Id " & vbCrLf & _
              "   AND nfxr.Romaneio_Id   = R.Romaneio_Id " & vbCrLf & _
              " WHERE NxN.Nota_Id         ='" & PNotaOrigem.Codigo & "' " & vbCrLf & _
              "   AND NxN.Empresa_Id      ='" & PNotaOrigem.CodigoEmpresa & "'" & vbCrLf & _
              "   AND NxN.EndEmpresa_Id   = " & PNotaOrigem.EnderecoEmpresa & vbCrLf & _
              "   AND NxN.Cliente_Id      ='" & PNotaOrigem.CodigoCliente & "'" & vbCrLf & _
              "   AND NxN.EndCliente_Id   ='" & PNotaOrigem.EnderecoCliente & "' " & vbCrLf & _
              "   AND NxN.EntradaSaida_Id ='" & IIf(PNotaOrigem.EntradaSaida = 1, "S", "E") & "'" & vbCrLf & _
              "   AND NxN.Serie_Id        ='" & PNotaOrigem.Serie & "'" & vbCrLf & _
              " GROUP BY NxN.Empresa_Id,NxN.EndEmpresa_Id,NxN.Cliente_Id,NxN.EndCliente_Id,NxN.EntradaSaida_Id,NxN.Serie_Id,NxN.Nota_Id, " & vbCrLf & _
              "       NxN.OrigemEmpresa_Id,NxN.OrigemEndEmpresa_Id,NxN.OrigemCliente_Id,NxN.OrigemEndCliente_Id,NxN.OrigemEntradaSaida_Id,NxN.OrigemSerie_Id," & vbCrLf & _
              "       NxN.OrigemNota_Id " & vbCrLf

        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(Sql, "NotasXNotas")

        For Each dr In ds.Tables(0).Rows
            Dim NxN As New NotasXNotas
            'NxN.OD = dr("OD")
            NxN.EmpresaCnpj = dr("Empresa_Id")
            NxN.EndEmpresa = dr("EndEmpresa_Id")
            NxN.ClienteCnpj = dr("Cliente_Id")
            NxN.EndCliente = dr("EndCliente_Id")
            NxN.EntradaSaida = dr("EntradaSaida_Id")
            NxN.Serie = dr("Serie_Id")
            NxN.NumeroNota = dr("Nota_Id")
            NxN.OrigemEmpresaCnpj = dr("OrigemEmpresa_Id")
            NxN.OrigemEndEmpresa = dr("OrigemEndEmpresa_Id")
            NxN.OrigemClienteCnpj = dr("OrigemCliente_Id")
            NxN.OrigemEndCliente = dr("OrigemEndCliente_Id")
            NxN.OrigemEntradaSaida = dr("OrigemEntradaSaida_Id")
            NxN.OrigemSerie = dr("OrigemSerie_Id")
            NxN.OrigemNota = dr("OrigemNota_Id")
            NxN.PesoFiscal = dr("PesoFiscal")
            NxN.PesoBrutoRomaneio = dr("PesoFiscalRomaneio")
            Me.Add(NxN)
        Next
    End Sub

    Public Function Salvar(ByVal pNF As NotaFiscal) As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        Me.SalvarSQL(pNF, Sqls)

        If Sqls.Count = 0 OrElse Banco.GravaBanco(Sqls) Then
            For Each x As NotasXNotas In Me
                x.IUD = ""
            Next
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSQL(ByVal pNF As NotaFiscal, ByRef Sqls As ArrayList)
        For Each Cl As NotasXNotas In Me
            If Not Cl.IUD = Nothing Then
                Cl.SalvarSql(pNF, Sqls)
            End If
        Next
    End Sub

#End Region

End Class

<Serializable()> _
Public Class NotasXNotas
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(ByVal PNotaOrigem As NotaFiscal, Optional ByVal Origem As Boolean = False)
        Dim Sql As String
        Dim ds As New DataSet
        Dim dr As DataRow
        Sql = " SELECT TOP 1 NxN.Empresa_Id, NxN.EndEmpresa_Id, NxN.Cliente_Id, NxN.EndCliente_Id, NxN.EntradaSaida_Id, NxN.Serie_Id, NxN.Nota_Id, " & vbCrLf & _
              "        NxN.OrigemEmpresa_Id, NxN.OrigemEndEmpresa_Id, NxN.OrigemCliente_Id, " & vbCrLf & _
              "        NxN.OrigemEndCliente_Id, NxN.OrigemEntradaSaida_Id, NxN.OrigemSerie_Id, NxN.OrigemNota_Id " & vbCrLf & _
              "   FROM NotasXNotas NxN " & vbCrLf & _
              "  WHERE " & vbCrLf
        If Origem Then
            Sql &= "        OrigemEmpresa_Id      ='" & PNotaOrigem.CodigoEmpresa & "' " & vbCrLf & _
                   "    AND OrigemEndEmpresa_Id   =" & PNotaOrigem.EnderecoEmpresa & " " & vbCrLf & _
                   "    AND OrigemCliente_Id      ='" & PNotaOrigem.CodigoCliente & "' " & vbCrLf & _
                   "    AND OrigemEndCliente_Id   =" & PNotaOrigem.EnderecoCliente & " " & vbCrLf & _
                   "    AND OrigemEntradaSaida_Id ='" & IIf(PNotaOrigem.EntradaSaida = 1, "S", "E") & "' " & vbCrLf & _
                   "    AND OrigemSerie_Id        ='" & PNotaOrigem.Serie & "' " & vbCrLf & _
                   "    AND OrigemNota_Id         =" & PNotaOrigem.Codigo & "" & vbCrLf
        Else
            Sql &= "        Empresa_Id      ='" & PNotaOrigem.CodigoEmpresa & "' " & vbCrLf & _
                   "    AND EndEmpresa_Id   =" & PNotaOrigem.EnderecoEmpresa & " " & vbCrLf & _
                   "    AND Cliente_Id      ='" & PNotaOrigem.CodigoCliente & "' " & vbCrLf & _
                   "    AND EndCliente_Id   =" & PNotaOrigem.EnderecoCliente & " " & vbCrLf & _
                   "    AND EntradaSaida_Id ='" & IIf(PNotaOrigem.EntradaSaida = 1, "S", "E") & "' " & vbCrLf & _
                   "    AND Serie_Id        ='" & PNotaOrigem.Serie & "' " & vbCrLf & _
                   "    AND Nota_Id         =" & PNotaOrigem.Codigo & "" & vbCrLf
        End If

        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(Sql, "NotasXNotas")
        For Each dr In ds.Tables(0).Rows
            Me.EmpresaCnpj = dr("Empresa_Id")
            Me.EndEmpresa = dr("EndEmpresa_Id")
            Me.ClienteCnpj = dr("Cliente_Id")
            Me.EndCliente = dr("EndCliente_Id")
            Me.EntradaSaida = dr("EntradaSaida_Id")
            Me.Serie = dr("Serie_Id")
            Me.NumeroNota = dr("Nota_Id")
            Me.OrigemEmpresaCnpj = dr("OrigemEmpresa_Id")
            Me.OrigemEndEmpresa = dr("OrigemEndEmpresa_Id")
            Me.OrigemClienteCnpj = dr("OrigemCliente_Id")
            Me.OrigemEndCliente = dr("OrigemEndCliente_Id")
            Me.OrigemEntradaSaida = dr("OrigemEntradaSaida_Id")
            Me.OrigemSerie = dr("OrigemSerie_Id")
            Me.OrigemNota = dr("OrigemNota_Id")
            Dim recupera As New ListNotasXNotas
            recupera = New ListNotasXNotas(PNotaOrigem, IIf(_OrigemSerie <> "UN", True, False))
            Dim i As Integer
            For i = 0 To recupera.Count - 1
                Me.PesoBrutoRomaneio += recupera.Item(i).PesoBrutoRomaneio
                Me.PesoFiscal += recupera.Item(i).PesoFiscal
            Next
        Next

    End Sub

    Public Sub New()

    End Sub

#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _OD As String = ""
    Private _ReduzidoEmpresa As String
    Private _NomeEmpresa As String
    Private _ReduzidoCliente As String
    Private _NomeCliente As String

    'Nota Fiscal
    Private _EmpresaCnpj As String
    Private _EndEmpresa As Integer
    Private _ClienteCnpj As String
    Private _EndCliente As Integer
    Private _EntradaSaida As String
    Private _Serie As String
    Private _NumeroNota As String
    Private _NotaFiscal As NotaFiscal

    'Origem Da Nota
    Private _OrigemEmpresaCnpj As String
    Private _OrigemEndEmpresa As Integer
    Private _OrigemClienteCnpj As String
    Private _OrigemEndCliente As Integer
    Private _OrigemEntradaSaida As String
    Private _OrigemSerie As String
    Private _OrigemNota As String

    Private _PesoFiscal As Decimal
    Private _PesoBrutoRomaneio As Decimal

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

    Public Property OD() As String
        Get
            Return _OD
        End Get
        Set(ByVal value As String)
            _OD = value
        End Set
    End Property

    Public Property ReduzidoEmpresa() As String
        Get
            Return _ReduzidoEmpresa
        End Get
        Set(ByVal value As String)
            _ReduzidoEmpresa = value
        End Set
    End Property

    Public Property NomeEmpresa() As String
        Get
            Return _NomeEmpresa
        End Get
        Set(ByVal value As String)
            _NomeEmpresa = value
        End Set
    End Property

    Public Property ReduzidoCliente() As String
        Get
            Return _ReduzidoCliente
        End Get
        Set(ByVal value As String)
            _ReduzidoCliente = value
        End Set
    End Property

    Public Property NomeCliente() As String
        Get
            Return _NomeCliente
        End Get
        Set(ByVal value As String)
            _NomeCliente = value
        End Set
    End Property

    'Nota Fiscal
    Public Property EmpresaCnpj() As String
        Get
            Return _EmpresaCnpj
        End Get
        Set(ByVal value As String)
            _EmpresaCnpj = value
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

    Public Property ClienteCnpj() As String
        Get
            Return _ClienteCnpj
        End Get
        Set(ByVal value As String)
            _ClienteCnpj = value
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
        End Set
    End Property

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value
        End Set
    End Property

    Public Property NumeroNota() As String
        Get
            Return _NumeroNota
        End Get
        Set(ByVal value As String)
            _NumeroNota = value
        End Set
    End Property

    Public Function NotaFiscal() As NotaFiscal
        Dim Nota As New NotaFiscal()
        If Not String.IsNullOrWhiteSpace(Me.EmpresaCnpj) _
          AndAlso Not String.IsNullOrWhiteSpace(Me.ClienteCnpj) _
          AndAlso Me.NumeroNota > 0 Then
            Nota.CodigoEmpresa = Me.EmpresaCnpj
            Nota.EnderecoEmpresa = Me.EndEmpresa
            Nota.CodigoCliente = Me.ClienteCnpj
            Nota.EnderecoCliente = Me.EndCliente
            Nota.EntradaSaida = Me.EntradaSaida.Equals("1")
            Nota.Serie = Me.Serie
            Nota.Codigo = Me.NumeroNota
        End If
        Return New NotaFiscal(Nota)
    End Function

    'Origem Da Nota
    Public Property OrigemEmpresaCnpj() As String
        Get
            Return _OrigemEmpresaCnpj
        End Get
        Set(ByVal value As String)
            _OrigemEmpresaCnpj = value
        End Set
    End Property

    Public Property OrigemEndEmpresa() As Integer
        Get
            Return _OrigemEndEmpresa
        End Get
        Set(ByVal value As Integer)
            _OrigemEndEmpresa = value
        End Set
    End Property

    Public Property OrigemClienteCnpj() As String
        Get
            Return _OrigemClienteCnpj
        End Get
        Set(ByVal value As String)
            _OrigemClienteCnpj = value
        End Set
    End Property

    Public Property OrigemEndCliente() As Integer
        Get
            Return _OrigemEndCliente
        End Get
        Set(ByVal value As Integer)
            _OrigemEndCliente = value
        End Set
    End Property

    Public Property OrigemEntradaSaida() As String
        Get
            Return _OrigemEntradaSaida
        End Get
        Set(ByVal value As String)
            _OrigemEntradaSaida = value
        End Set
    End Property

    Public Property OrigemSerie() As String
        Get
            Return _OrigemSerie
        End Get
        Set(ByVal value As String)
            _OrigemSerie = value
        End Set
    End Property

    Public Property OrigemNota() As String
        Get
            Return _OrigemNota
        End Get
        Set(ByVal value As String)
            _OrigemNota = value
        End Set
    End Property

    Public Property PesoFiscal() As Decimal
        Get
            Return _PesoFiscal
        End Get
        Set(ByVal value As Decimal)
            _PesoFiscal = value
        End Set
    End Property

    Public Property PesoBrutoRomaneio() As Decimal
        Get
            Return _PesoBrutoRomaneio
        End Get
        Set(ByVal value As Decimal)
            _PesoBrutoRomaneio = value
        End Set
    End Property
#End Region

#Region "Métodos"

    Public Function RNotaFrete() As NotaFiscal
        If Me.EmpresaCnpj Is Nothing Then Return Nothing 'se nao a relacionamento entre notas consultadas

        Dim NotaConsulta As New NotaFiscal()
        NotaConsulta.CodigoEmpresa = Me.EmpresaCnpj
        NotaConsulta.EnderecoEmpresa = Me.EndEmpresa
        NotaConsulta.CodigoCliente = Me.ClienteCnpj
        NotaConsulta.EnderecoCliente = Me.EndCliente
        NotaConsulta.EntradaSaida = Me.EntradaSaida.Equals("1")
        NotaConsulta.Serie = Me.Serie
        NotaConsulta.Codigo = Me.NumeroNota

        Return New NotaFiscal(NotaConsulta)

    End Function

    Public Function RNotaOrigem() As NotaFiscal
        Dim NotaConsultaOrigem As New NotaFiscal()
        NotaConsultaOrigem.CodigoEmpresa = Me.OrigemEmpresaCnpj
        NotaConsultaOrigem.EnderecoEmpresa = Me.OrigemEndEmpresa
        NotaConsultaOrigem.CodigoCliente = Me.OrigemClienteCnpj
        NotaConsultaOrigem.EnderecoCliente = Me.OrigemEndCliente
        NotaConsultaOrigem.EntradaSaida = Me.OrigemEntradaSaida.Equals("1")
        NotaConsultaOrigem.Serie = Me.OrigemSerie
        NotaConsultaOrigem.Codigo = Me.OrigemNota

        Return New NotaFiscal(NotaConsultaOrigem)

    End Function

    Public Function Salvar(ByVal pNF As NotaFiscal) As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        Me.SalvarSql(pNF, Sqls)

        If Sqls.Count = 0 Then Return True

        If Banco.GravaBanco(Sqls) Then
            Me.IUD = ""
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByVal NF As NotaFiscal, ByRef Sqls As ArrayList)
        Dim Sql As String
        Select Case Me.IUD

            'Monta o sql de acordo com o tipo da nota
            Case Me.IUD = "I"
                Sql = " INSERT INTO NotasXNotas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, " & vbCrLf & _
                      "                         OrigemPais_Id, OrigemEmpresa_Id, OrigemEndEmpresa_Id, OrigemCliente_Id, OrigemEndCliente_Id, " & vbCrLf & _
                      "                         OrigemEntradaSaida_Id, OrigemSerie_Id, OrigemNota_Id)" & vbCrLf & _
                      "'" & Me.EmpresaCnpj & "'," & Me.EndEmpresa & ",'" & Me.ClienteCnpj & "'," & Me.EndCliente & ",'" & Me.EntradaSaida & "','" & Me.Serie & "'," & Me.NumeroNota & vbCrLf

            Case Me.IUD = "D"
                Sql = "DELETE NotasXNotas" & vbCrLf & _
                      " WHERE Empresa_Id            ='" & NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id         = " & NF.EnderecoEmpresa & " " & vbCrLf & _
                      "   AND Cliente_Id            ='" & NF.CodigoCliente & "'" & vbCrLf & _
                      "   AND EndCliente_Id         = " & NF.EnderecoCliente & vbCrLf & _
                      "   AND EntradaSaida_Id       ='" & NF.EntradaSaida & "'" & vbCrLf & _
                      "   AND Serie_Id              ='" & NF.Serie & "'" & vbCrLf & _
                      "   AND Nota_Id               = " & NF.Codigo & vbCrLf & _
                      "   AND OrigemEmpresa_Id      ='" & Me.EmpresaCnpj & "'" & vbCrLf & _
                      "   AND OrigemEndEmpresa_Id   = " & Me.EndEmpresa & vbCrLf & _
                      "   AND OrigemCliente_Id      ='" & Me.ClienteCnpj & "'" & vbCrLf & _
                      "   AND OrigemEndCliente_Id   = " & Me.EndCliente & vbCrLf & _
                      "   AND OrigemEntradaSaida_Id ='" & Me.EntradaSaida & "'" & vbCrLf & _
                      "   AND OrigemSerie_Id        ='" & Me.Serie & "'" & vbCrLf & _
                      "   AND OrigemNota_Id         = " & Me.NumeroNota & vbCrLf
                If Sql <> "" Then Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class