Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListMemorandoDeExportacaoXRegistroDeExportacao
    Inherits List(Of MemorandoDeExportacaoXRegistroDeExportacao)

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
              "RegistroDeExportacao_Id, DataExportacao, UfProdutor " & vbCrLf & _
              "  FROM MemorandoDeExportacaoXRegistroDeExportacao" & vbCrLf & _
              " Where EmpresaMemorando_Id    ='" & Mem.CodigoEmpresaMemorando & "'" & vbCrLf & _
              "   and EndEmpresaMemorando_Id = " & Mem.EnderecoEmpresaMemorando & vbCrLf & _
              "   and Memorando_Id           = '" & Mem.CodigoMemorando & "' " & vbCrLf

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Notas")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim MMxEX As New MemorandoDeExportacaoXRegistroDeExportacao(Mem)
            MMxEX.CodigoEmpresaMem = row("EmpresaMemorando_Id")
            MMxEX.EnderecoEmpresaMem = row("EndEmpresaMemorando_Id")
            MMxEX.CodigoMemorando = row("Memorando_Id")
            MMxEX.CodRegistroDeExportacao = row("RegistroDeExportacao_Id")
            MMxEX.DataRegExportacao = row("DataExportacao")
            MMxEX.UfProdutor = row("UfProdutor")

            Me.Add(MMxEX)
        Next

        Banco = Nothing
    End Sub
#End Region

#Region "Methods"


    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As MemorandoDeExportacaoXRegistroDeExportacao In Me
            If Memorando.IUD = "D" Or Memorando.IUD = "I" Then item.IUD = Memorando.IUD
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub

#End Region
End Class


<Serializable()> _
Public Class MemorandoDeExportacaoXRegistroDeExportacao

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
    Private _CodRegistroDeExportacao As String
    Private _DataRegExportacao As Date = Now
    Private _UfProdutor As String = ""


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

    Public Property CodRegistroDeExportacao() As String
        Get
            Return _CodRegistroDeExportacao
        End Get
        Set(ByVal value As String)
            _CodRegistroDeExportacao = value

        End Set
    End Property

    Public Property DataRegExportacao() As Date
        Get
            Return _DataRegExportacao
        End Get
        Set(ByVal value As Date)
            _DataRegExportacao = value

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
                sql = " Insert Into MemorandoDeExportacaoXRegistroDeExportacao(EmpresaMemorando_Id, EndEmpresaMemorando_Id, " & vbCrLf & _
                      " Memorando_Id, RegistroDeExportacao_Id, DataExportacao, UfProdutor)" & vbCrLf & _
                      " Values('" & Memorando.CodigoEmpresaMemorando & "'," & Memorando.EnderecoEmpresaMemorando & ",'" & Memorando.CodigoMemorando & "', " & vbCrLf & _
                      "        '" & _CodRegistroDeExportacao & "','" & _DataRegExportacao.ToString("yyyy-MM-dd") & "', '" & _UfProdutor & "')"
                Sqls.Add(sql)
            Case "U"
                sql = "Update MemorandoDeExportacaoXRegistroDeExportacao set" & vbCrLf & _
                      "   DataExportacao        = '" & _DataRegExportacao.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "   UfProdutor        = '" & _UfProdutor & "'" & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and RegistroDeExportacao_Id='" & _CodRegistroDeExportacao & "'"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete MemorandoDeExportacaoXRegistroDeExportacao" & vbCrLf & _
                      " Where EmpresaMemorando_Id    ='" & Memorando.CodigoEmpresaMemorando & "'" & vbCrLf & _
                      "   and EndEmpresaMemorando_Id = " & Memorando.EnderecoEmpresaMemorando & vbCrLf & _
                      "   and Memorando_Id           = '" & Memorando.CodigoMemorando & "'" & vbCrLf & _
                      "   and RegistroDeExportacao_Id='" & _CodRegistroDeExportacao & "'"

                Sqls.Add(sql)
        End Select
    End Sub
#End Region

End Class
