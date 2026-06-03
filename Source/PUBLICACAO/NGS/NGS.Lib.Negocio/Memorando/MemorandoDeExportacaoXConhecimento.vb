Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMemorandoDeExportacaoXConhecimento
    Inherits List(Of MemorandoDeExportacaoXConhecimento)

#Region "Fields"
    Private _Memorando As MemorandoDeExportacao
#End Region

#Region "Property"
    Public Property Memorando() As MemorandoDeExportacao
        Get
            Return _Memorando
        End Get
        Set(ByVal value As MemorandoDeExportacao)
            _Memorando = value
        End Set
    End Property
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Mem As MemorandoDeExportacao)
        _Memorando = Mem

        Dim Sql As String
        Sql = "SELECT EmpresaMemorando_Id, EndEmpresaMemorando_Id, Memorando_Id, " & vbCrLf & _
              "ConhecimentoDeEmbarque_Id, DataConhecimento, TipoConhecimento " & vbCrLf & _
              "  FROM MemorandoDeExportacaoXConhecimentoDeEmbarque" & vbCrLf & _
              " Where EmpresaMemorando_Id    ='" & Mem.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and EndEmpresaMemorando_Id = " & Mem.EnderecoEmpresaMemorando & vbCrLf & _
              "   and Memorando_Id           = '" & Mem.CodigoMemorando & "' " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Conhecimentos")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim MMxCON As New MemorandoDeExportacaoXConhecimento(Mem)
            MMxCON.CodigoEmpresaMem = row("EmpresaMemorando_Id")
            MMxCON.EnderecoEmpresaMem = row("EndEmpresaMemorando_Id")
            MMxCON.CodigoMemorando = row("Memorando_Id")
            MMxCON.NumConhecimentoDeEmbarque = row("ConhecimentoDeEmbarque_Id")
            MMxCON.DataConhecimento = row("DataConhecimento")
            MMxCON.TipoConhecimento = row("TipoConhecimento")

            Me.Add(MMxCON)
        Next

        Banco = Nothing
    End Sub
#End Region

#Region "Methods"


    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As MemorandoDeExportacaoXConhecimento In Me
            If Memorando.IUD = "D" Or Memorando.IUD = "I" Then item.IUD = Memorando.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

#End Region

End Class


<Serializable()> _
Public Class MemorandoDeExportacaoXConhecimento

#Region "Contrutor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Mem As MemorandoDeExportacao)
        _Memorando = Mem
    End Sub

#End Region

#Region "Fields"
    Private _IUD As String
    Private _Memorando As MemorandoDeExportacao
    Private _CodigoEmpresaMem As String
    Private _EnderecoEmpresaMem As Integer
    Private _CodigoMemorando As String = ""
    Private _NumConhecimentoDeEmbarque As String
    Private _DataConhecimento As Date = Now
    Private _TipoConhecimento As String = ""
    Private _DescTipoConhecimento As String = ""


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

    Public Property Memorando() As MemorandoDeExportacao
        Get
            Return _Memorando
        End Get
        Set(ByVal value As MemorandoDeExportacao)
            _Memorando = value
        End Set
    End Property

    Public Property CodigoEmpresaMem() As String
        Get
            Return _CodigoEmpresaMem
        End Get
        Set(ByVal value As String)
            _CodigoEmpresaMem = value

        End Set
    End Property

    Public Property EnderecoEmpresaMem() As Integer
        Get
            Return _EnderecoEmpresaMem
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresaMem = value

        End Set
    End Property


    Public Property CodigoMemorando() As String
        Get
            Return _CodigoMemorando
        End Get
        Set(ByVal value As String)
            _CodigoMemorando = value

        End Set
    End Property

    Public Property NumConhecimentoDeEmbarque() As String
        Get
            Return _NumConhecimentoDeEmbarque
        End Get
        Set(ByVal value As String)
            _NumConhecimentoDeEmbarque = value

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
                sql = " Insert Into MemorandoDeExportacaoXConhecimentoDeEmbarque(EmpresaMemorando_Id, EndEmpresaMemorando_Id, " & vbCrLf & _
                      " Memorando_Id, ConhecimentoDeEmbarque_Id, DataConhecimento, TipoConhecimento)" & vbCrLf & _
                      " Values('" & Memorando.CodigoEmpresaMemorando & "'," & Memorando.EnderecoEmpresaMemorando & ",'" & Memorando.CodigoMemorando & "', " & vbCrLf & _
                      "        '" & _NumConhecimentoDeEmbarque & "','" & _DataConhecimento.ToString("yyyy-MM-dd") & "', '" & _TipoConhecimento & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update MemorandoDeExportacaoXConhecimentoDeEmbarque set" & vbCrLf & _
                      "   DataConhecimento        = '" & _DataConhecimento.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "   TipoConhecimento        = '" & _TipoConhecimento & "'" & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and ConhecimentoDeEmbarque_Id='" & _NumConhecimentoDeEmbarque & "'"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete MemorandoDeExportacaoXConhecimentoDeEmbarque" & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and ConhecimentoDeEmbarque_Id='" & _NumConhecimentoDeEmbarque & "'"

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
