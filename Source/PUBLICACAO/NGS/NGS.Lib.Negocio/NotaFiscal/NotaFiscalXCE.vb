Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Web
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListNotaFiscalXCE
    Inherits List(Of NotaFiscalXCE)

#Region "Contrutor"
    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota
        Dim sql As String
        sql = " SELECT ConhecimentoDeEmbarque_Id, isnull(DataConhecimento,getdate()) as DataConhecimento, TipoConhecimento" & vbCrLf & _
              "   FROM NotaFiscalXCE" & vbCrLf & _
              "  Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
              "    and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
              "    and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
              "    and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
              "    and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
              "    and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
              "    and Nota_Id         = " & Me.NF.Codigo & vbCrLf & _
              "    and Nota_Id         > 0"
        Dim ds As DataSet
        Dim Banco As New AcessaBanco

        ds = Banco.ConsultaDataSet(sql, "DadosRE")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        For Each row As DataRow In ds.Tables(0).Rows
            Dim CE As New NotaFiscalXCE(pNota)
            CE.IUD = "U"
            CE.ConhecimentoDeEmbarque = row("ConhecimentoDeEmbarque_Id")
            CE.DataConhecimento = row("DataConhecimento")
            CE.TipoConhecimento = row("TipoConhecimento")
            Me.Add(CE)
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
        For Each item As NotaFiscalXCE In Me
            If NF.IUD = "D" Or NF.IUD = "I" Then item.IUD = NF.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class NotaFiscalXCE
#Region "Contrutor"
    Public Sub New(ByVal pNota As NotaFiscal)
        _NF = pNota
    End Sub
#End Region

#Region "Fields"
    Private _NF As NotaFiscal

    Private _IUD As String = ""

    Private _ConhecimentoDeEmbarque As String = ""
    Private _DataConhecimento As Date = Now
    Private _TipoConhecimento As String = ""
    Private _DescTipoConhecimento As String = ""
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

    Public Property ConhecimentoDeEmbarque() As String
        Get
            Return _ConhecimentoDeEmbarque
        End Get
        Set(ByVal value As String)
            _ConhecimentoDeEmbarque = value
        End Set
    End Property

    Public Property DataConhecimento() As Date
        Get
            Return _DataConhecimento
        End Get
        Set(ByVal value As Date)
            _DataConhecimento = value
        End Set
    End Property

    Public Property TipoConhecimento() As String
        Get
            Return _TipoConhecimento
        End Get
        Set(ByVal value As String)
            _TipoConhecimento = value
        End Set
    End Property

    Public ReadOnly Property DescTipoConhecimento() As String
        Get
            Select Case _TipoConhecimento
                Case "01" : Return "AWB"
                Case "02" : Return "MAWB"
                Case "03" : Return "HAWB"
                Case "04" : Return "COMAT"
                Case "06" : Return "R. EXPRESSAS"
                Case "07" : Return "ETIQ. REXPRESSAS"
                Case "08" : Return "HR. EXPRESSAS"
                Case "09" : Return "AV7"
                Case "10" : Return "BL"
                Case "11" : Return "MBL"
                Case "12" : Return "HBL"
                Case "13" : Return "CRT"
                Case "14" : Return "DSIC"
                Case "16" : Return "COMAT BL"
                Case "17" : Return "RWB"
                Case "18" : Return "HRWB"
                Case "19" : Return "TIF/DTA"
                Case "20" : Return "CP2"
                Case "91" : Return "NÂO IATA"
                Case "92" : Return "MNAO IATA"
                Case "93" : Return "HNAO IATA"
                Case "99" : Return "OUTROS"
            End Select
            Return String.Empty
        End Get
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
                sql = " Insert Into NotaFiscalXCE(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id," & vbCrLf & _
                      "                           ConhecimentoDeEmbarque_Id, DataConhecimento, TipoConhecimento) " & vbCrLf & _
                      " Values('" & NF.CodigoEmpresa & "'," & NF.EnderecoEmpresa & ",'" & NF.CodigoCliente & "'," & NF.EnderecoCliente & ",'" & IIf(NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "','" & NF.Serie & "'," & NF.Codigo & "," & vbCrLf & _
                      "'" & Me.ConhecimentoDeEmbarque & "','" & Me.DataConhecimento.ToString("yyyy-MM-dd") & "','" & Me.TipoConhecimento & "')"

                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotaFiscalXCE" & vbCrLf & _
                      " Where Empresa_Id      ='" & Me.NF.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id   = " & Me.NF.EnderecoEmpresa & vbCrLf & _
                      "   and Cliente_Id      ='" & Me.NF.CodigoCliente & "'" & vbCrLf & _
                      "   and EndCliente_Id   = " & Me.NF.EnderecoCliente & vbCrLf & _
                      "   and EntradaSaida_Id ='" & IIf(Me.NF.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "'" & vbCrLf & _
                      "   and Serie_Id        ='" & Me.NF.Serie & "'" & vbCrLf & _
                      "   and Nota_Id         = " & Me.NF.Codigo & vbCrLf & _
                      "   and ConhecimentoDeEmbarque_Id ='" & Me.ConhecimentoDeEmbarque & "'"
                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class