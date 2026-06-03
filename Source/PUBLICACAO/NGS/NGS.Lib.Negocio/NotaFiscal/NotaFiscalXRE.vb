Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXRE
    Inherits List(Of NotaFiscalXRE)

#Region "Contrutor"
    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota
        Dim sql As String
        sql = " SELECT RegistroDeExportacao_Id, isnull(DataRegistroDeExportacao,getdate()) as DataRegistroDeExportacao, UFProdutor" & vbCrLf & _
              "   FROM NotaFiscalXRe" & vbCrLf & _
              "  Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
              "    and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
              "    and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
              "    and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
              "    and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "    and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
              "    and Nota_Id         = " & Me.NF.Codigo & vbCrLf & _
              "    And Nota_id > 0"
        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(sql, "DadosRE")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        For Each row As DataRow In ds.Tables(0).Rows
            Dim RE As New NotaFiscalXRE(pNota)
            RE.IUD = "U"
            RE.RegistroDeExportacao = row("RegistroDeExportacao_Id")
            RE.DataRegistroDeExportacao = row("DataRegistroDeExportacao")
            RE.UfProdutor = row("UFProdutor")
            Me.Add(RE)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal
#End Region

#Region "Property"
    Public ReadOnly Property NF() As NotaFiscal
        Get
            Return _NF
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
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

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As NotaFiscalXRE In Me
            If NF.IUD = "D" Or NF.IUD = "I" Then item.IUD = NF.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class NotaFiscalXRE

#Region "Contrutor"
    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal

    Private _IUD As String = ""

    Private _RegistroDeExportacao As String = ""
    Private _DataRegistroDeExportacao As Date = Now
    Private _UfProdutor As String = ""
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

    Public Property RegistroDeExportacao() As String
        Get
            Return _RegistroDeExportacao
        End Get
        Set(ByVal value As String)
            _RegistroDeExportacao = value
        End Set
    End Property

    Public Property DataRegistroDeExportacao() As Date
        Get
            Return _DataRegistroDeExportacao
        End Get
        Set(ByVal value As Date)
            _DataRegistroDeExportacao = value
        End Set
    End Property

    Public Property UfProdutor() As String
        Get
            Return _UfProdutor
        End Get
        Set(ByVal value As String)
            _UfProdutor = value
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
                sql = " Insert Into NotaFiscalXRE(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id," & vbCrLf & _
                      "                           RegistroDeExportacao_Id, DataRegistroDeExportacao, UFProdutor) " & vbCrLf & _
                      " Values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ",'" & NF.CodigoCliente & "'," & NF.EnderecoCliente & ",'" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & NF.Serie & "'," & NF.Codigo & "," & vbCrLf & _
                      "'" & Me.RegistroDeExportacao & "','" & Me.DataRegistroDeExportacao.ToString("yyyy-MM-dd") & "','" & Me.UfProdutor & "')"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalXRE" & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf & _
                      "   and RegistroDeExportacao_Id ='" & Me.RegistroDeExportacao & "'" & vbCrLf & _
                      "   and UFProdutor ='" & Me.UfProdutor & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
