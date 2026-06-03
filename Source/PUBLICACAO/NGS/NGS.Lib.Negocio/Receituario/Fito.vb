Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListaFito
    Inherits List(Of Fito)

    Public Sub New(Optional ByVal pCodigoProduto As String = "", Optional ByVal pCodigoFito As Integer = 0, Optional ByVal pNome As String = "", Optional ByVal OrderBY As String = "", Optional ByVal pCodigoIndea As String = "")
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String
        Dim strWhereAnd As String = "WHERE"
        sql = "SELECT F.Fito_Id, F.ClasseTox, F.ClasseAmbiental, F.FormulacaoFito, F.IA, F.ClasseRisco, F.NomeComercial, F.NomeTecnico, " & vbCrLf & _
              "       isnull(F.FormulaBruta,'') as FormulaBruta, isnull(F.UFRestricao,'') as UFRestricao, isnull(F.ConcentracaoING,'') as ConcentracaoING, " & vbCrLf & _
              "       F.Inflamavel, F.RegistroMA, F.RegistroONU, isnull(F.ModoAplicacao,'') as ModoAplicacao,  isnull(F.InstrucoesUSO,'') as InstrucoesUSO, isnull(F.DescarteEmbalagem,'') as DescarteEmbalagem, " & vbCrLf & _
              "       isnull(F.PrimeirosSocorros,'') as PrimeirosSocorros, F.MeioAmbiente, isnull(F.Incompatibilidade,'') as Incompatibilidade, isnull(F.CodigoIndeaMT,'') as CodigoIndeaMT, " & vbCrLf & _
              "       isnull(F.NomeArquivoFrenteFE,'') as NomeArquivoFrenteFE, isnull(F.NomeArquivoVersoFE,'') as NomeArquivoVersoFE,  " & vbCrLf & _
              "       isnull(F.NomeEmbarque,'') as NomeEmbarque,  isnull(F.SubClasseRisco,'') as SubClasseRisco" & vbCrLf & _
              "  FROM Fito F" & vbCrLf & _
              "  Left Join ProdutoXFito PF" & vbCrLf & _
              "    on F.Fito_Id    = PF.Fito_Id " & vbCrLf
        If pCodigoProduto.Length > 0 Then
            sql &= strWhereAnd & " PF.Produto_Id ='" & pCodigoProduto & "' " & vbCrLf
            strWhereAnd = "AND"
        End If

        If pCodigoFito > 0 Then
            sql &= strWhereAnd & " F.Fito_id = " & pCodigoFito & " " & vbCrLf
            strWhereAnd = "AND"
        End If

        If pNome.Length > 0 Then
            sql &= strWhereAnd & " F.NomeComercial Like '" & pNome & "%'" & vbCrLf
        End If

        If pCodigoIndea.Length > 0 Then
            sql &= strWhereAnd & " F.CodigoIndeaMT Like '%" & pCodigoIndea & "%'" & vbCrLf
            strWhereAnd = "AND"
        End If



        If OrderBY.Length > 0 Then sql &= " Order By " & OrderBY

        ds = Banco.ConsultaDataSet(sql, "Fito")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim FT As New Fito
            FT.CodigoFito = row("Fito_Id")
            FT.CodigoClasseTox = row("ClasseTox")
            FT.CodigoClasseAmbiental = row("ClasseAmbiental")
            FT.CodigoIA = row("IA")
            FT.CodigoClasseDeRisco = row("ClasseRisco")
            FT.NomeComercial = row("NomeComercial")
            FT.NomeTecnico = row("NomeTecnico")
            FT.FormulaBruta = row("FormulaBruta")
            FT.UFRestricao = row("UFRestricao")
            FT.ConcentracaoING = row("ConcentracaoING")
            FT.Inflamavel = row("Inflamavel")
            FT.RegistroMA = row("RegistroMA")
            FT.RegistroONU = row("RegistroONU")
            FT.ModoAplicacao = row("ModoAplicacao")
            FT.InstrucoesUso = row("InstrucoesUSO")
            FT.DescarteEmbalagem = row("DescarteEmbalagem")
            FT.PrimeirosSocorros = row("PrimeirosSocorros")
            FT.MeioAmbiente = row("MeioAmbiente")
            FT.Incompatibilidade = row("Incompatibilidade")
            FT.CodigoIndeaMT = row("CodigoIndeaMT")
            FT.NomeArquivoFrenteFE = row("NomeArquivoFrenteFE")
            FT.NomeArquivoVersoFE = row("NomeArquivoVersoFE")
            FT.NomeEmbarque = row("NomeEmbarque")
            FT.SubClasseRisco = row("SubClasseRisco")
            Me.Add(FT)
        Next

    End Sub

End Class

<Serializable()> _
Public Class Fito

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoProduto As String, ByVal pCodigoFito As Integer)
        If pCodigoProduto.Length = 0 And pCodigoFito = 0 Then Exit Sub
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = "SELECT F.Fito_Id, F.ClasseTox, F.ClasseAmbiental, F.FormulacaoFito, F.IA, F.ClasseRisco, F.NomeComercial, F.NomeTecnico, " & vbCrLf & _
              "       isnull(F.FormulaBruta,'') as FormulaBruta, isnull(F.UFRestricao,'') as UFRestricao, isnull(F.ConcentracaoING,'') as ConcentracaoING, " & vbCrLf & _
              "       F.Inflamavel, F.RegistroMA, F.RegistroONU, isnull(F.ModoAplicacao,'') as ModoAplicacao,  isnull(F.InstrucoesUSO,'') as InstrucoesUSO, isnull(F.DescarteEmbalagem,'') as DescarteEmbalagem, " & vbCrLf & _
              "       isnull(F.PrimeirosSocorros,'') as PrimeirosSocorros, F.MeioAmbiente, isnull(F.Incompatibilidade,'') as Incompatibilidade, isnull(F.CodigoIndeaMT,'') as CodigoIndeaMT, " & vbCrLf & _
              "       isnull(F.NomeArquivoFrenteFE,'') as NomeArquivoFrenteFE, isnull(F.NomeArquivoVersoFE,'') as NomeArquivoVersoFE,  " & vbCrLf & _
              "       isnull(F.NomeEmbarque,'') as NomeEmbarque,  isnull(F.SubClasseRisco,'') as SubClasseRisco" & vbCrLf & _
              "  FROM Fito F" & vbCrLf & _
              "  Left Join ProdutoXFito PF" & vbCrLf & _
              "    on F.Fito_Id    = PF.Fito_Id" & vbCrLf
        If pCodigoProduto.Length > 0 Then
            sql &= "  Where PF.Produto_Id ='" & pCodigoProduto & "'" & vbCrLf
        Else
            sql &= "  Where  F.Fito_id = " & pCodigoFito & "" & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "Fito")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoFito = row("Fito_Id")
        _CodigoClasseTox = row("ClasseTox")
        _CodigoClasseAmbiental = row("ClasseAmbiental")
        _CodigoFormulacaoFito = row("FormulacaoFito")
        _CodigoIA = row("IA")
        _CodigoClasseDeRisco = row("ClasseRisco")
        _NomeComercial = row("NomeComercial")
        _NomeTecnico = row("NomeTecnico")
        _FormulaBruta = row("FormulaBruta")
        _UFRestricao = row("UFRestricao")
        _ConcentracaoING = row("ConcentracaoING")
        _Inflamavel = row("Inflamavel")
        _Corrosivo = row("Inflamavel")
        _RegistroMA = row("RegistroMA")
        _RegistroONU = row("RegistroONU")
        _ModoAplicacao = row("ModoAplicacao")
        _InstrucoesUso = row("InstrucoesUSO")
        _DescarteEmbalagem = row("DescarteEmbalagem")
        _PrimeirosSocorros = row("PrimeirosSocorros")
        _MeioAmbiente = row("MeioAmbiente")
        _Incompatibilidade = row("Incompatibilidade")
        _CodigoIndeaMT = row("CodigoIndeaMT")
        _NomeArquivoFrenteFE = row("NomeArquivoFrenteFE")
        _NomeArquivoVersoFE = row("NomeArquivoVersoFE")
        _NomeEmbarque = row("NomeEmbarque")
        _SubClasseRisco = row("SubClasseRisco")
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoFito As Integer
    Private _ClasseTox As ClasseToxicologica
    Private _CodigoClasseTox As Integer
    Private _ClasseAmbiental As ClasseAmbiental
    Private _CodigoClasseAmbiental As Integer
    Private _FormulacaoFito As FormulacaoFito
    Private _CodigoFormulacaoFito As Integer
    Private _IA As IA
    Private _CodigoIA As Integer
    Private _ClasseDeRisco As ClasseDeRisco
    Private _CodigoClasseDeRisco As Integer
    Private _NomeComercial As String
    Private _NomeTecnico As String
    Private _FormulaBruta As String
    Private _UFRestricao As String
    Private _ConcentracaoING As String
    Private _Inflamavel As String
    Private _Corrosivo As String
    Private _RegistroMA As String
    Private _RegistroONU As String
    Private _ModoAplicacao As String
    Private _InstrucoesUso As String
    Private _DescarteEmbalagem As String
    Private _PrimeirosSocorros As String
    Private _MeioAmbiente As String
    Private _Incompatibilidade As String
    Private _CodigoIndeaMT As String
    Private _NomeArquivoFrenteFE As String
    Private _NomeArquivoVersoFE As String
    Private _NomeEmbarque As String
    Private _SubClasseRisco As String
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

    Public Property CodigoFito() As Integer
        Get
            Return _CodigoFito
        End Get
        Set(ByVal value As Integer)
            _CodigoFito = value
        End Set
    End Property

    Public ReadOnly Property ClasseTox() As ClasseToxicologica
        Get
            If _ClasseTox Is Nothing And _CodigoClasseTox > 0 Then _ClasseTox = New ClasseToxicologica(_CodigoClasseTox)
            Return _ClasseTox
        End Get
    End Property

    Public Property CodigoClasseTox() As Integer
        Get
            Return _CodigoClasseTox
        End Get
        Set(ByVal value As Integer)
            _CodigoClasseTox = value
            _ClasseTox = Nothing
        End Set
    End Property

    Public ReadOnly Property ClasseAmbiental() As ClasseAmbiental
        Get
            If _ClasseAmbiental Is Nothing And _CodigoClasseAmbiental > 0 Then _ClasseAmbiental = New ClasseAmbiental(_CodigoClasseAmbiental)
            Return _ClasseAmbiental
        End Get
    End Property

    Public Property CodigoClasseAmbiental() As Integer
        Get
            Return _CodigoClasseAmbiental
        End Get
        Set(ByVal value As Integer)
            _CodigoClasseAmbiental = value
            _ClasseAmbiental = Nothing
        End Set
    End Property

    Public ReadOnly Property FormulacaoFito() As FormulacaoFito
        Get
            If _FormulacaoFito Is Nothing And _CodigoFormulacaoFito > 0 Then _FormulacaoFito = New FormulacaoFito(_CodigoFormulacaoFito)
            Return _FormulacaoFito
        End Get
    End Property

    Public Property CodigoFormulacaoFito() As Integer
        Get
            Return _CodigoFormulacaoFito
        End Get
        Set(ByVal value As Integer)
            _CodigoFormulacaoFito = value
            _FormulacaoFito = Nothing
        End Set
    End Property

    Public ReadOnly Property IA() As IA
        Get
            If _IA Is Nothing And _CodigoIA > 0 Then _IA = New IA(_CodigoIA)
            Return _IA
        End Get
    End Property

    Public Property CodigoIA() As Integer
        Get
            Return _CodigoIA
        End Get
        Set(ByVal value As Integer)
            _CodigoIA = value
            _IA = Nothing
        End Set
    End Property

    Public ReadOnly Property ClasseDeRisco() As ClasseDeRisco
        Get
            If _ClasseDeRisco Is Nothing And _CodigoClasseDeRisco > 0 Then _ClasseDeRisco = New ClasseDeRisco(_CodigoClasseDeRisco)
            Return _ClasseDeRisco
        End Get
    End Property

    Public Property CodigoClasseDeRisco() As Integer
        Get
            Return _CodigoClasseDeRisco
        End Get
        Set(ByVal value As Integer)
            _CodigoClasseDeRisco = value
            _ClasseDeRisco = Nothing
        End Set
    End Property

    Public Property NomeComercial() As String
        Get
            Return _NomeComercial
        End Get
        Set(ByVal value As String)
            _NomeComercial = value
        End Set
    End Property

    Public Property NomeTecnico() As String
        Get
            Return _NomeTecnico
        End Get
        Set(ByVal value As String)
            _NomeTecnico = value
        End Set
    End Property

    Public Property FormulaBruta() As String
        Get
            Return _FormulaBruta
        End Get
        Set(ByVal value As String)
            _FormulaBruta = value
        End Set
    End Property

    Public Property UFRestricao() As String
        Get
            Return _UFRestricao
        End Get
        Set(ByVal value As String)
            _UFRestricao = value
        End Set
    End Property

    Public Property ConcentracaoING() As String
        Get
            Return _ConcentracaoING
        End Get
        Set(ByVal value As String)
            _ConcentracaoING = value
        End Set
    End Property

    Public Property Inflamavel() As String
        Get
            Return _Inflamavel
        End Get
        Set(ByVal value As String)
            _Inflamavel = value
        End Set
    End Property

    Public Property Corrosivo() As String
        Get
            Return _Corrosivo
        End Get
        Set(ByVal value As String)
            _Corrosivo = value
        End Set
    End Property

    Public Property RegistroMA() As String
        Get
            Return _RegistroMA
        End Get
        Set(ByVal value As String)
            _RegistroMA = value
        End Set
    End Property

    Public Property RegistroONU() As String
        Get
            Return _RegistroONU
        End Get
        Set(ByVal value As String)
            _RegistroONU = value
        End Set
    End Property

    Public Property ModoAplicacao() As String
        Get
            Return _ModoAplicacao
        End Get
        Set(ByVal value As String)
            _ModoAplicacao = value
        End Set
    End Property

    Public Property InstrucoesUso() As String
        Get
            Return _InstrucoesUso
        End Get
        Set(ByVal value As String)
            _InstrucoesUso = value
        End Set
    End Property

    Public Property DescarteEmbalagem() As String
        Get
            Return _DescarteEmbalagem
        End Get
        Set(ByVal value As String)
            _DescarteEmbalagem = value
        End Set
    End Property

    Public Property PrimeirosSocorros() As String
        Get
            Return _PrimeirosSocorros
        End Get
        Set(ByVal value As String)
            _PrimeirosSocorros = value
        End Set
    End Property

    Public Property MeioAmbiente() As String
        Get
            Return _MeioAmbiente
        End Get
        Set(ByVal value As String)
            _MeioAmbiente = value
        End Set
    End Property

    Public Property Incompatibilidade() As String
        Get
            Return _Incompatibilidade
        End Get
        Set(ByVal value As String)
            _Incompatibilidade = value
        End Set
    End Property

    Public Property CodigoIndeaMT() As String
        Get
            Return _CodigoIndeaMT
        End Get
        Set(ByVal value As String)
            _CodigoIndeaMT = value
        End Set
    End Property

    Public Property NomeArquivoFrenteFE() As String
        Get
            Return _NomeArquivoFrenteFE
        End Get
        Set(ByVal value As String)
            _NomeArquivoFrenteFE = value
        End Set
    End Property

    Public Property NomeArquivoVersoFE() As String
        Get
            Return _NomeArquivoVersoFE
        End Get
        Set(ByVal value As String)
            _NomeArquivoVersoFE = value
        End Set
    End Property

    Public Property NomeEmbarque() As String
        Get
            Return _NomeEmbarque
        End Get
        Set(ByVal value As String)
            _NomeEmbarque = value
        End Set
    End Property

    Public Property SubClasseRisco() As String
        Get
            Return _SubClasseRisco
        End Get
        Set(ByVal value As String)
            _SubClasseRisco = value
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
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO FITO (Fito_Id, ClasseTox, ClasseAmbiental, FormulacaoFito, IA, ClasseRisco, NomeComercial, NomeTecnico, FormulaBruta, UFRestricao, ConcentracaoING, Inflamavel, " & vbCrLf & _
                      "                   Corrosivo, RegistroMA, RegistroONU, ModoAplicacao, InstrucoesUSO, DescarteEmbalagem, PrimeirosSocorros, MeioAmbiente, Incompatibilidade, CodigoIndeaMT, " & vbCrLf & _
                      "                   NomeArquivoFrenteFE, NomeArquivoVersoFE, NomeEmbarque, SubClasseRisco)" & vbCrLf & _
                      " VALUES (" & _CodigoFito & "," & _CodigoClasseTox & "," & _CodigoClasseAmbiental & "," & _CodigoFormulacaoFito & "," & _CodigoIA & "," & _CodigoClasseDeRisco & ",'" & _NomeComercial & "','" & _NomeTecnico & "','" & _FormulaBruta & "','" & UFRestricao & "','" & _ConcentracaoING & "','" & _Inflamavel & "'," & vbCrLf & _
                      "         '" & _Corrosivo & "','" & _RegistroMA & "','" & _RegistroONU & "','" & _ModoAplicacao & "','" & _InstrucoesUso & "','" & _DescarteEmbalagem & "','" & _PrimeirosSocorros & "','" & _MeioAmbiente & "','" & _Incompatibilidade & "','" & _CodigoIndeaMT & "'," & vbCrLf & _
                      "         '" & _NomeArquivoFrenteFE & "','" & _NomeArquivoVersoFE & "','" & NomeEmbarque & "','" & SubClasseRisco & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE FITO SET" & vbCrLf & _
                      "   ClasseTox         = " & _CodigoClasseTox & vbCrLf & _
                      "  ,ClasseAmbiental   = " & _CodigoClasseAmbiental & vbCrLf & _
                      "  ,FormulacaoFito    = " & _CodigoFormulacaoFito & vbCrLf & _
                      "  ,IA                = " & _CodigoIA & vbCrLf & _
                      "  ,ClasseRisco       = " & _CodigoClasseDeRisco & vbCrLf & _
                      "  ,NomeComercial     ='" & _NomeComercial & "'" & vbCrLf & _
                      "  ,NomeTecnico       ='" & _NomeTecnico & "'" & vbCrLf & _
                      "  ,FormulaBruta      ='" & _FormulaBruta & "'" & vbCrLf & _
                      "  ,UFRestricao       ='" & _UFRestricao & "'" & vbCrLf & _
                      "  ,ConcentracaoING   ='" & _ConcentracaoING & "'" & vbCrLf & _
                      "  ,Inflamavel        ='" & _Inflamavel & "'" & vbCrLf & _
                      "  ,Corrosivo         ='" & _Corrosivo & "'" & vbCrLf & _
                      "  ,RegistroMA        ='" & _RegistroMA & "'" & vbCrLf & _
                      "  ,RegistroONU       ='" & _RegistroONU & "'" & vbCrLf & _
                      "  ,ModoAplicacao     ='" & _ModoAplicacao & "'" & vbCrLf & _
                      "  ,InstrucoesUSO     ='" & _InstrucoesUso & "'" & vbCrLf & _
                      "  ,DescarteEmbalagem ='" & _DescarteEmbalagem & "'" & vbCrLf & _
                      "  ,PrimeirosSocorros ='" & _PrimeirosSocorros & "'" & vbCrLf & _
                      "  ,MeioAmbiente      ='" & _MeioAmbiente & "'" & vbCrLf & _
                      "  ,Incompatibilidade ='" & _Incompatibilidade & "'" & vbCrLf & _
                      "  ,CodigoIndeaMT     ='" & _CodigoIndeaMT & "'" & vbCrLf & _
                      "  ,NomeArquivoFrenteFE='" & _NomeArquivoFrenteFE & "'" & vbCrLf & _
                      "  ,NomeArquivoVersoFE='" & _NomeArquivoVersoFE & "'" & vbCrLf & _
                      "  ,NomeEmbarque      ='" & _NomeEmbarque & "'" & vbCrLf & _
                      "  ,SubClasseRisco    ='" & _SubClasseRisco & "'" & vbCrLf & _
                      "  WHERE Fito_Id =" & _CodigoFito & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE FITO" & vbCrLf & _
                      "  WHERE Fito_Id = " & _CodigoFito & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub

    Public Sub VerSequencia()
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String
        sql = "Select max(Fito_id) + 1 as Fito From Fito" & vbCrLf

        ds = Banco.ConsultaDataSet(sql, "Fito")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        _CodigoFito = row("Fito")
    End Sub
#End Region

End Class