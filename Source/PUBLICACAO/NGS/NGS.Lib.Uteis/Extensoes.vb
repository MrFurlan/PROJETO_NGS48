Imports System
Imports System.Linq
Imports System.Text
Imports System.IO
Imports System.Web
Imports System.Data
Imports System.Collections
Imports System.Globalization
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Reflection
Imports System.Windows.Forms
Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports System.Security.Cryptography
Imports System.Text.RegularExpressions
Imports System.Runtime.Serialization.Formatters.Binary
Imports System.Runtime.CompilerServices

Module Extensoes

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Concat(Of t)(lst As IEnumerable(Of t), valueFunction As Func(Of t, Object)) As String
        Return lst.Concat(valueFunction, ", ", Nothing, Nothing)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Concat(Of t)(lst As IEnumerable(Of t), valueFunction As Func(Of t, Object), separator As String) As String
        Return lst.Concat(valueFunction, separator, Nothing, Nothing)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Concat(Of t)(lst As IEnumerable(Of t), valueFunction As Func(Of t, Object), separator As String, format As String) As String
        Return lst.Concat(valueFunction, separator, format, Nothing)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Concat(Of t)(lst As IEnumerable(Of t), valueFunction As Func(Of t, Object), separator As String, format As String, defaultValue As String) As String
        If lst Is Nothing OrElse lst.Count() = 0 Then
            Return defaultValue
        End If

        Dim str As New StringBuilder()
        For Each obj As t In lst
            Dim value As Object = valueFunction(obj)
            Dim valuestr As String = Convert.ToString(value)

            If String.IsNullOrEmpty(valuestr) Then
                Continue For
            End If

            If String.IsNullOrEmpty(format) Then
                str.Append(valuestr + separator)
            Else
                str.Append(String.Format("{0:" & Convert.ToString(format) & "}", value) + separator)
            End If
        Next
        Return If(String.IsNullOrEmpty(str.ToString()), defaultValue, str.Remove(str.Length - separator.Length, separator.Length).ToString())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFirstThreeCharacters(ByVal str As String) As String
        If str.Length < 3 Then
            Return str
        Else
            Return str.Substring(0, 3)
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToInt32(ByVal obj As [Object]) As Int32
        If obj Is Nothing OrElse [String].IsNullOrEmpty(obj.ToString()) Then
            Return 0
        End If
        Try
            Return Convert.ToInt32(obj)
        Catch generatedExceptionName As Exception
            Return 0
        End Try
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Sub Times(ByVal value As Integer, ByVal action As Action(Of Integer))
        For i As Integer = 0 To value - 1
            action(i)
        Next
    End Sub

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFirstName(ByVal str As [String], ByVal maxChars As Int32) As [String]
        If [String].IsNullOrEmpty(str) Then
            Return str
        End If
        Dim name = str.Trim()
        Dim ind = name.IndexOf(" "c)
        If ind = -1 Then
            Return If((maxChars <> -1 AndAlso name.Length > maxChars), name.Remove(maxChars - 3).Trim() & "...", name.Trim())
        Else
            name = name.Substring(0, ind)
            If maxChars <> -1 AndAlso name.Length > maxChars Then
                Return name.Remove(maxChars - 3).Trim() & "..."
            End If
        End If
        Return name.Trim()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetLastName(ByVal str As [String], ByVal maxChars As Int32) As [String]
        If [String].IsNullOrEmpty(str) Then
            Return str
        End If
        Dim name = str.Trim()
        Dim ind = name.LastIndexOf(" "c)
        If ind = -1 Then
            Return If((maxChars <> -1 AndAlso name.Length > maxChars), name.Remove(maxChars - 3).Trim() & "...", name.Trim())
        Else
            name = name.Substring(ind + 1)
            If maxChars <> -1 AndAlso name.Length > maxChars Then
                Return name.Remove(maxChars - 3).Trim() & "..."
            End If
        End If
        Return name.Trim()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Except(Of t)(ByVal lst1 As IEnumerable(Of t), ByVal lst2 As IEnumerable(Of t), ByVal func As Func(Of t, [Object])) As IEnumerable(Of t)
        Return lst1.Where(Function(obj1) Not lst2.Any(Function(obj2)
                                                          Dim v1 = func(obj1)
                                                          Dim v2 = func(obj2)
                                                          If v1 Is Nothing AndAlso v2 Is Nothing Then
                                                              Return True
                                                          End If
                                                          If v1 Is Nothing OrElse v2 Is Nothing Then
                                                              Return False
                                                          End If
                                                          Return v1.Equals(v2)

                                                      End Function))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function EndsWith(ByVal obj As [String], ByVal ParamArray str As [String]()) As [Boolean]
        If [String].IsNullOrEmpty(obj) Then
            Return False
        End If
        For Each s As [String] In str
            If obj.EndsWith(s) Then
                Return True
            End If
        Next
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsMatch(ByVal str As [String], ByVal pattern As [String]) As [Boolean]
        If [String].IsNullOrEmpty(str) Then
            Return False
        End If
        Return Regex.IsMatch(str, pattern, RegexOptions.IgnoreCase)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFilename(ByVal str As [String]) As [String]
        If [String].IsNullOrEmpty(str) Then
            Return str
        End If

        Dim lst1 As [String] = " \/:*?""<>|"
        Dim lst2 As [String] = "__________"
        Return str.RemoveAccents().ReplaceAll(lst1.ToCharArray().ToList(), lst2.ToCharArray().ToList()).ToLower()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveAccents(ByVal str As [String]) As [String]
        If [String].IsNullOrEmpty(str) Then
            Return str
        End If

        Dim lst1 As [String] = "áéíóúàèìòùäëïöüãõâêîôûçÁÉÍÓÚÀÈÌÒÙÄËÏÖÜÃÕÂÊÎÔÛÇ"
        Dim lst2 As [String] = "aeiouaeiouaeiouaoaeioucAEIOUAEIOUAEIOUAOAEIOUC"
        Return str.ReplaceAll(lst1.ToCharArray().ToList(), lst2.ToCharArray().ToList())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveAspas(ByVal str As [String]) As [String]
        Return str.Replace("""", "").Replace("'", "")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ReplaceAll(ByVal str As [String], ByVal oldChars As List(Of [Char]), ByVal newChars As List(Of [Char])) As [String]
        If [String].IsNullOrEmpty(str) OrElse oldChars Is Nothing OrElse newChars Is Nothing Then
            Return str
        End If

        Dim builder As New StringBuilder(str)
        For Each c As [Char] In oldChars
            Dim valor = c
            builder.Replace(valor, newChars(oldChars.FindIndex(Function(cc) cc = valor)))
        Next

        Return builder.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveAll(Of t)(ByVal lst As IEnumerable(Of t), ByVal func As Func(Of t, [Boolean])) As IEnumerable(Of t)
        Dim newlst As New List(Of t)()
        For Each o As t In lst
            If Not func(o) Then
                newlst.Add(o)
            End If
        Next
        Return newlst
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function UnionAll(Of t)(ByVal lst As IEnumerable(Of t), ByVal second As IEnumerable(Of t)) As IEnumerable(Of t)
        Dim newlst = New List(Of t)()
        newlst.AddRange(lst)
        newlst.AddRange(second)
        Return newlst
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveAllExcept(Of t)(ByVal lst As IEnumerable(Of t), ByVal ParamArray array As t()) As IEnumerable(Of t)
        Dim newlst As New List(Of t)()
        For Each o As t In lst
            newlst.Add(o)
        Next
        For Each o As t In newlst
            If Not array.Contains(o) Then
                newlst.Remove(o)
            End If
        Next
        Return newlst
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToList(Of t)(ByVal lst As ICollection) As List(Of t)
        Return lst.Cast(Of t)().ToList()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Any(ByVal obj As [String], ByVal ParamArray str As [String]()) As [Boolean]
        If [String].IsNullOrEmpty(obj) Then
            Return False
        End If

        For Each s As [String] In str
            If obj.Contains(s) Then
                Return True
            End If
        Next
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function AllEquals(Of t)(ByVal lst As IEnumerable(Of t), ByVal func As Func(Of t, [Object])) As [Boolean]
        If lst Is Nothing OrElse lst.Count() = 0 Then
            Return True
        End If

        Dim value = func(lst.First())
        If value Is Nothing AndAlso lst.All(Function(obj) func(obj) Is Nothing) Then
            Return True
        End If
        Return lst.All(Function(obj) value.Equals(func(obj)))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Encode(ByVal obj As [String]) As [String]
        If [String].IsNullOrEmpty(obj) Then
            Return obj
        End If

        Dim abb As [String] = ""
        For Each c As [Char] In obj.ToCharArray()
            Dim value As Int32 = Convert.ToInt32(c)
            abb += If((value > 255), "#(#" & value & "#)#", Convert.ToString(c))
        Next
        Return abb
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Decode(ByVal obj As [String]) As [String]
        If [String].IsNullOrEmpty(obj) Then
            Return obj
        End If

        Dim pattern As [String] = "#\({1}#\d+#\){1}#"
        If Not obj.IsMatch(pattern) Then
            Return obj
        End If

        Dim abb As [String] = obj
        While abb.IsMatch(pattern)
            Dim ind1 As Int32 = abb.IndexOf("#(#")
            Dim ind2 As Int32 = abb.IndexOf("#)#")

            Dim valor As Int32 = Convert.ToInt32(abb.Substring(ind1 + 3, ind2 - (ind1 + 3)))
            abb = abb.Remove(ind1, ind2 - ind1 + 3)
            abb = abb.Insert(ind1, Convert.ToString(CType(CChar(CStr(valor)), [Char])))
        End While
        Return abb
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ContainsIgnoreCase(ByVal obj As [String], ByVal str As [String]) As [Boolean]
        If [String].IsNullOrEmpty(obj) OrElse [String].IsNullOrEmpty(str) Then
            Return False
        End If
        Return obj.ToLower().Contains(str.ToLower())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Clone(Of T)(ByVal obj As T) As T
        If obj Is Nothing Then
            Throw New NullReferenceException("Can't clone null objects.")
        End If

        Dim ms As New MemoryStream()
        Dim bf As New BinaryFormatter()
        bf.Serialize(ms, obj)
        ms.Position = 0
        Dim clone__1 As Object = bf.Deserialize(ms)
        ms.Close()
        Return DirectCast(clone__1, T)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ContainsAny(ByVal obj As [String], ByVal ParamArray lst As [String]()) As [Boolean]
        If obj Is Nothing OrElse lst Is Nothing OrElse lst.Length = 0 Then
            Return False
        End If
        For Each v As String In lst
            If obj.Contains(v) Then
                Return True
            End If
        Next
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ContainsAll(ByVal obj As IEnumerable(Of [String]), ByVal ParamArray lst As [String]()) As [Boolean]
        If obj Is Nothing OrElse obj.Count() = 0 OrElse lst Is Nothing OrElse lst.Length = 0 Then
            Return False
        End If
        For Each v As String In lst
            If Not obj.Contains(v) Then
                Return False
            End If
        Next
        Return True
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ContainsAny(ByVal obj As IEnumerable(Of [String]), ByVal ParamArray lst As [String]()) As [Boolean]
        If obj Is Nothing OrElse obj.Count() = 0 OrElse lst Is Nothing OrElse lst.Length = 0 Then
            Return False
        End If
        For Each v As String In lst
            If obj.Contains(v) Then
                Return True
            End If
        Next
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetDecimalPlaces(ByVal value As [Decimal]) As Int32
        Return Convert.ToDouble(value).GetDecimalPlaces()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetDecimalPlaces(ByVal value As System.Nullable(Of [Decimal])) As Int32
        Return If(value Is Nothing, 0, Convert.ToDouble(value.Value).GetDecimalPlaces())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFormatedNumber(ByVal value As System.Nullable(Of [Decimal]), ByVal decimals As Int32) As [String]
        Return If(value.HasValue, value.Value.ToString("###############0." & Convert.ToString(("0".Repeat(decimals)))), "")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFormatedNumber(ByVal value As [Decimal], ByVal decimals As Int32) As [String]
        Return value.ToString("###############0." & Convert.ToString(("0".Repeat(decimals))))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetDecimalPlaces(ByVal value As [Double]) As Int32
        Dim s = NumberFormatInfo.CurrentInfo.NumberDecimalSeparator
        Dim str = Convert.ToString(Convert.ToDecimal(value))
        Dim ind = str.LastIndexOf(s)
        If ind = -1 Then
            Return 0
        End If
        Return str.Length - ind - 1
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyChars(ByVal str As [String], ByVal charListToKeep As [String]) As [String]
        Dim newStr = New StringBuilder()
        str.ToCharArray().Where(Function(c) charListToKeep.Any(Function(cc) cc = c)).ToList().ForEach(Function(c) newStr.Append(c))
        Return newStr.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyNumbers(ByVal str As [String]) As [String]
        Return str.OnlyChars("0123456789")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyChars(ByVal str As [String]) As [String]
        Return str.OnlyChars("abcdefghijklmnopqrstuvxwyzABCDEFGHIJKLMNOPQRSTUVXWYZ")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function OnlyConsonants(ByVal str As [String]) As [String]
        Return str.OnlyChars("bcdfghjklmnpqrstvxwyzBCDFGHJKLMNPQRSTVXWYZ")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function WithoutAccents(ByVal str As [String]) As [String]
        Dim chars1 = "áéíóúçãõàèìòùäëïöüÁÉÍÓÚÇÃÕÀÈÌÒÙÄËÏÖÜ"
        Dim chars2 = "aeioucaoaeiouaeiouAEIOUCAOAEIOUAEIOU"
        For x As Int32 = 0 To chars1.Length - 1
            str = str.Replace(chars1(x), chars2(x))
        Next
        Return str
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function StartsWithAny(ByVal obj As [String], ByVal ParamArray str As [String]()) As [Boolean]
        If [String].IsNullOrEmpty(obj) OrElse str Is Nothing OrElse str.Length = 0 OrElse str.All(Function(s) [String].IsNullOrEmpty(s)) Then
            Return False
        End If
        For Each s As String In str
            If obj.StartsWith(s, StringComparison.InvariantCultureIgnoreCase) Then
                Return True
            End If
        Next
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ValidEmail(ByVal str As [String]) As [Boolean]
        Dim rgx As New Regex("^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$")
        Return rgx.IsMatch(str)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Format(ByVal str As [String], ByVal mask As [String]) As [String]
        Dim dado As New StringBuilder()
        For Each c As Char In str
            If [Char].IsNumber(c) Then
                dado.Append(c)
            End If
        Next

        Dim indMascara As Integer = mask.Length
        Dim indCampo As Integer = dado.Length

        While indCampo > 0 AndAlso indMascara > 0
            If mask(System.Threading.Interlocked.Decrement(indMascara)) = "#"c Then
                indCampo -= 1
            End If
        End While

        Dim saida As New StringBuilder()
        While indMascara < mask.Length
            saida.Append(If((mask(indMascara) = "#"c), dado(System.Math.Max(System.Threading.Interlocked.Increment(indCampo), indCampo - 1)), mask(indMascara)))
            indMascara += 1
        End While

        Return saida.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FormatPhone(ByVal obj As [String]) As [String]
        obj = obj.Replace("(", "").Replace(")", "").Replace("-", "")
        Dim r As String = String.Empty
        Dim pattern As String = "(\d{3})(\d{3})(\d{4})"
        r = Regex.Replace(obj, pattern, "($1) $2-$3")
        Return r
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FormatCPF(ByVal cpf As [String]) As [String]
        If [String].IsNullOrEmpty(cpf) Then
            Return String.Empty
        End If
        Dim valor = Convert.ToUInt64(cpf.OnlyNumbers())
        Return [String].Format("{0:000\.000\.000\-00}", valor)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FormatCNPJ(ByVal cnpj As [String]) As [String]
        If [String].IsNullOrEmpty(cnpj) Then
            Return String.Empty
        End If
        Dim valor = Convert.ToUInt64(cnpj.OnlyNumbers())
        Return [String].Format("{0:00\.000\.000\/0000\-00}", valor)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FormatRG(ByVal rg As [String]) As [String]
        Return [String].Format("{0:00\.000\.000\-0}", rg)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValidCPF(ByVal cpf As [String]) As [Boolean]
        Dim valor As String = cpf.Replace(".", "")
        valor = valor.Replace("-", "")

        If valor.Length <> 11 Then
            Return False
        End If

        Dim igual As Boolean = True
        Dim i As Integer = 1
        While i < 11 AndAlso igual
            If valor(i) <> valor(0) Then
                igual = False
            End If
            i += 1
        End While

        If igual OrElse valor = "12345678909" Then
            Return False
        End If

        Dim numeros As Integer() = New Integer(10) {}
        For ii As Integer = 0 To 10
            numeros(ii) = Integer.Parse(valor(ii).ToString())
        Next

        Dim soma As Integer = 0
        For iii As Integer = 0 To 8
            soma += (10 - iii) * numeros(iii)
        Next

        Dim resultado As Integer = soma Mod 11
        If resultado = 1 OrElse resultado = 0 Then
            If numeros(9) <> 0 Then
                Return False
            End If
        ElseIf numeros(9) <> 11 - resultado Then
            Return False
        End If

        soma = 0
        For iiii As Integer = 0 To 9
            soma += (11 - iiii) * numeros(iiii)
        Next
        resultado = soma Mod 11

        If resultado = 1 OrElse resultado = 0 Then
            If numeros(10) <> 0 Then
                Return False
            End If
        ElseIf numeros(10) <> 11 - resultado Then
            Return False
        End If

        Return True
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValidCNPJ(ByVal cnpj As [String]) As [Boolean]
        Dim vCNPJ As String = cnpj.Replace(".", "")
        vCNPJ = vCNPJ.Replace("/", "")
        vCNPJ = vCNPJ.Replace("-", "")

        Dim digitos As Integer(), soma As Integer(), resultado As Integer()
        Dim nrDig As Integer
        Dim ftmt As String
        Dim CNPJOk As Boolean()

        ftmt = "6543298765432"
        digitos = New Integer(13) {}
        soma = New Integer(1) {}
        soma(0) = 0
        soma(1) = 0
        resultado = New Integer(1) {}
        resultado(0) = 0
        resultado(1) = 0
        CNPJOk = New Boolean(1) {}
        CNPJOk(0) = False
        CNPJOk(1) = False

        Try
            For nrDig = 0 To 13
                digitos(nrDig) = Integer.Parse(vCNPJ.Substring(nrDig, 1))
                If nrDig <= 11 Then
                    soma(0) += (digitos(nrDig) * Integer.Parse(ftmt.Substring(nrDig + 1, 1)))
                End If
                If nrDig <= 12 Then
                    soma(1) += (digitos(nrDig) * Integer.Parse(ftmt.Substring(nrDig, 1)))
                End If
            Next

            For nrDig = 0 To 1
                resultado(nrDig) = (soma(nrDig) Mod 11)
                If (resultado(nrDig) = 0) OrElse (resultado(nrDig) = 1) Then
                    CNPJOk(nrDig) = (digitos(12 + nrDig) = 0)
                Else
                    CNPJOk(nrDig) = (digitos(12 + nrDig) = (11 - resultado(nrDig)))
                End If
            Next
            Return (CNPJOk(0) AndAlso CNPJOk(1))
        Catch
            Return False
        End Try
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValidCEP(ByVal cep As [String]) As [Boolean]
        If [String].IsNullOrEmpty(cep) Then
            Return False
        End If
        cep = cep.Replace(".", "").Replace("-", "")
        If cep.Length = 8 Then
            cep = cep.Substring(0, 5) & "-" & cep.Substring(5, 3)
        End If
        Return Regex.IsMatch(cep, ("[0-9]{5}-[0-9]{3}"))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsValidEmail(ByVal email As [String]) As [Boolean]
        If [String].IsNullOrEmpty(email) Then
            Return False
        End If
        Dim rg As New Regex("^[A-Za-z0-9](([_\.\-]?[a-zA-Z0-9]+)*)@([A-Za-z0-9]+)(([\.\-]?[a-zA-Z0-9]+)*)\.([A-Za-z]{2,})$")
        Return rg.IsMatch(email)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToList(ByVal str As [String]) As List(Of [Char])
        Dim lst = New List(Of [Char])()
        For x As Int32 = 0 To str.Length - 1
            lst.Add(str(x))
        Next
        Return lst
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetStringDiff(ByVal datetime2 As DateTime, ByVal datetime1 As DateTime) As [String]
        Dim diff = datetime2.Subtract(datetime1)
        Return [String].Format("{0:00}:{1:000}", diff.Seconds, diff.Milliseconds)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function Repeat(ByVal str As [String], ByVal n As Int32) As [String]
        Dim rs = ""
        For x As Int32 = 0 To n - 1
            rs += str
        Next
        Return rs
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function SHA256Encrypt(ByVal str As [String]) As [String]
        If str Is Nothing Then
            str = ""
        End If
        Dim bytes As [Byte]() = Encoding.[Default].GetBytes(str)
        Dim sha As New SHA256Managed()
        bytes = sha.ComputeHash(bytes)
        Dim rs As [String] = Convert.ToBase64String(bytes)
        sha.Clear()
        Return rs
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsCnpj(ByVal cnpj As [String]) As [Boolean]
        If Not [String].IsNullOrEmpty(cnpj) Then
            cnpj = cnpj.Trim()
        End If
        If Not cnpj.IsMatch("(^(\d{2,3}.\d{3}.\d{3}/\d{4}-\d{2})|(\d{14,15})$)") Then
            Return False
        End If

        cnpj = cnpj.Replace("/", "")
        cnpj = cnpj.Replace(".", "")
        cnpj = cnpj.Replace("-", "")

        Dim digits As [String]
        If Not (cnpj.Length >= 14 AndAlso cnpj.Length <= 15) OrElse cnpj = "0".Repeat(cnpj.Length) Then
            Return False
        End If

        digits = cnpj.Substring(cnpj.Length - 2)
        Return VerifyDigit(Convert.ToInt32(digits.Substring(0, 1)), cnpj.Substring(0, cnpj.Length - 2)) AndAlso VerifyDigit(Convert.ToInt32(digits.Substring(1, 1)), cnpj.Substring(0, cnpj.Length - 1))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Private Function VerifyDigit(ByVal digit As Int32, ByVal numbers As [String]) As [Boolean]
        Dim sum As Int32 = 0, i As Int32, result As Int32
        Dim pos As Int32 = numbers.Length - 7
        For i = 0 To numbers.Length - 1
            sum += Convert.ToInt32(numbers.Substring(i, 1)) * System.Math.Max(System.Threading.Interlocked.Decrement(pos), pos + 1)
            If pos < 2 Then
                pos = 9
            End If
        Next
        result = If(sum Mod 11 < 2, 0, 11 - (sum Mod 11))
        Return result = digit
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDateTime(ByVal obj As [Object]) As DateTime
        Return Convert.ToDateTime(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableDateTime(ByVal obj As [Object]) As System.Nullable(Of DateTime)
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToDateTime(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableInt32(ByVal obj As [Object]) As System.Nullable(Of Int32)
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToInt32(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableDecimal(ByVal obj As [Object]) As System.Nullable(Of [Decimal])
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToDecimal(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDecimal(ByVal obj As [Object]) As [Decimal]
        Return Convert.ToDecimal(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToNullableDouble(ByVal obj As [Object]) As System.Nullable(Of [Double])
        If obj Is Nothing Then
            Return Nothing
        End If
        Return Convert.ToDouble(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDouble(ByVal obj As [Object]) As [Double]
        Return Convert.ToDouble(obj)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToTimeSpan(ByVal obj As [Object]) As TimeSpan
        Return TimeSpan.Parse(obj.ToString())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToFormattedTimeSpan(ByVal obj As System.Nullable(Of TimeSpan)) As [String]
        Dim ts = obj.ToTimeSpan()
        Return New DateTime(ts.Ticks).ToString("HH:mm")
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function AddWeekDays(ByVal [date] As DateTime, ByVal days As Int32) As DateTime
        While days > 0
            [date] = [date].AddDays(1)
            If [date].DayOfWeek = DayOfWeek.Saturday Then
                [date] = [date].AddDays(2)
            End If
            If [date].DayOfWeek = DayOfWeek.Sunday Then
                [date] = [date].AddDays(1)
            End If
            days -= 1
        End While
        Return [date]
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetEnumDescription(ByVal value As [Enum]) As [String]
        Dim fi As FieldInfo = value.[GetType]().GetField(value.ToString())
        Dim attributes = DirectCast(fi.GetCustomAttributes(GetType(DescriptionAttribute), False), DescriptionAttribute())
        Return If((attributes.Length > 0), attributes(0).Description, value.ToString())
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function CastTo(Of t)(ByVal o As [Object], ByVal type As t) As t
        Return DirectCast(o, t)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsNumber(ByVal obj As [Object]) As [Boolean]
        If TypeOf obj Is Int16 OrElse TypeOf obj Is Int32 OrElse TypeOf obj Is Int64 Then
            Return True
        End If
        If TypeOf obj Is [Decimal] OrElse TypeOf obj Is [Single] OrElse TypeOf obj Is [Double] Then
            Return True
        End If
        If TypeOf obj Is [String] AndAlso Not [String].IsNullOrEmpty(TryCast(obj, [String])) AndAlso TryCast(obj, [String]).IsMatch("^\d+$") Then
            Return True
        End If
        Return False
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function IsUniqueException(ByVal ex As Exception) As [Boolean]
        Return ex.Message.ToUpper().ContainsIgnoreCase("UNICIDADE") OrElse (ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message.ToUpper().ContainsIgnoreCase("UNICIDADE")) OrElse ex.Message.ToUpper().ContainsIgnoreCase("UNIQUE") OrElse (ex.InnerException IsNot Nothing AndAlso ex.InnerException.Message.ToUpper().ContainsIgnoreCase("UNIQUE"))
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetFirstDay(ByVal dt As DateTime) As DateTime
        Dim today As DateTime = DateTime.Now
        Dim firstDay As DateTime = today.AddDays(-(today.Day - 1))
        Return firstDay
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function GetLastDay(ByVal dt As DateTime) As DateTime
        Dim today As DateTime = DateTime.Now
        today = today.AddMonths(1)
        Dim lastDay As DateTime = today.AddDays(-(today.Day))
        Return lastDay
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function CapitalizeWords(ByVal value As [String]) As [String]
        If value Is Nothing Then
            Throw New ArgumentNullException("value")
        End If

        If value.Length = 0 Then
            Return value
        End If

        Dim result As New StringBuilder(value)
        result(0) = Char.ToUpper(result(0))

        For i As Integer = 1 To result.Length - 1
            If Char.IsWhiteSpace(result(i - 1)) Then
                result(i) = Char.ToUpper(result(i))
            End If
        Next
        Return result.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function CapitalizeWordsCulture(ByVal value As [String]) As [String]
        Return CultureInfo.CurrentCulture.TextInfo.ToTitleCase(value)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToMD5(ByVal value As [String]) As [String]
        Dim md5__1 As MD5 = MD5.Create()

        Dim inputBytes As Byte() = System.Text.Encoding.ASCII.GetBytes(value)
        Dim hash As Byte() = md5__1.ComputeHash(inputBytes)

        Dim sb As New StringBuilder()

        For i As Integer = 0 To hash.Length - 1
            sb.Append(hash(i).ToString("X2"))
        Next

        Return sb.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveSpecialCharacters(ByVal str As [String]) As [String]
        Dim normalizedString = str

        Dim symbolTable = New Dictionary(Of Char, Char())()
        symbolTable.Add("a"c, New Char() {"à"c, "á"c, "ä"c, "â"c, "ã"c})
        symbolTable.Add("c"c, New Char() {"ç"c})
        symbolTable.Add("e"c, New Char() {"è"c, "é"c, "ë"c, "ê"c})
        symbolTable.Add("i"c, New Char() {"ì"c, "í"c, "ï"c, "î"c})
        symbolTable.Add("o"c, New Char() {"ò"c, "ó"c, "ö"c, "ô"c, "õ"c})
        symbolTable.Add("u"c, New Char() {"ù"c, "ú"c, "ü"c, "û"c})

        For Each key As Char In symbolTable.Keys
            For Each symbol As Char In symbolTable(key)
                normalizedString = normalizedString.Replace(symbol, key)
            Next
        Next

        normalizedString = Regex.Replace(str, "[^0-9a-zA-ZéúíóáÉÚÍÓÁèùìòàÈÙÌÒÀõãñÕÃÑêûîôâÊÛÎÔÂëÿüïöäËYÜÏÖÄçÇ\s]+?", String.Empty)
        Return normalizedString
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Private Function RemoveHtmlTags(ByVal str As [String]) As [String]
        Const pattern As String = "<(.|\n)*?>"
        Return Regex.Replace(str, pattern, String.Empty)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function RemoveInjection(ByVal str As [String]) As [String]
        Dim caracteresInvalidos As String() = {"select ", "drop ", "--", "insert ", "delete ", "xp_", _
         "'", "%", "update ", "group by ", "having ", "sum\(", _
         "count\(", "alter table", " – ", "–", " –", "– ", _
         "varchar", "declare", "cast\(", "exec\("}
        For i As Integer = 1 To caracteresInvalidos.Length - 1
            str = Regex.Replace(str, caracteresInvalidos(i), "", RegexOptions.IgnoreCase)
        Next
        Return str.Trim()
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDataTable(Of T)(ByVal iList As IList(Of T)) As DataTable
        Dim dataTable As New DataTable()
        Dim propertyDescriptorCollection As PropertyDescriptorCollection = TypeDescriptor.GetProperties(GetType(T))
        For i As Integer = 0 To propertyDescriptorCollection.Count - 1
            Dim propertyDescriptor As PropertyDescriptor = propertyDescriptorCollection(i)
            dataTable.Columns.Add(propertyDescriptor.Name, propertyDescriptor.PropertyType)
        Next
        Dim values As Object() = New Object(propertyDescriptorCollection.Count - 1) {}
        For Each iListItem As T In iList
            For i As Integer = 0 To values.Length - 1
                values(i) = propertyDescriptorCollection(i).GetValue(iListItem)
            Next
            dataTable.Rows.Add(values)
        Next
        Return dataTable
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Private Function ToArray(Of T)(ByVal dataTable As DataTable, ByVal tType As Type, ByVal tPropertiesInfo As PropertyInfo(), ByVal columnIndices As Integer()) As T()
        Dim dataRow As DataRow
        Dim tInstance As T

        Dim array As T() = New T(dataTable.Rows.Count - 1) {}
        For i As Integer = 0 To dataTable.Rows.Count - 1
            dataRow = dataTable.Rows(i)
            tInstance = DirectCast(Activator.CreateInstance(tType), T)
            For j As Integer = 0 To tPropertiesInfo.Count() - 1
                tPropertiesInfo(j).SetValue(tInstance, dataRow(columnIndices(j)), Nothing)
            Next
            array(i) = tInstance
        Next
        Return array
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function FirstDayOfMonthFromDateTime(ByVal dt As DateTime) As DateTime
        Return New DateTime(dt.Year, dt.Month, 1)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function LastDayOfMonthFromDateTime(ByVal dt As DateTime) As DateTime
        Dim firstDayOfTheMonth As New DateTime(dt.Year, dt.Month, 1)
        Return firstDayOfTheMonth.AddMonths(1).AddDays(-1)
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToStringNull(ByVal obj As Object) As String
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToString(obj)
        Else
            Return Nothing
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToInt(ByVal obj As Object) As Integer
        If obj IsNot Nothing Then
            Return Convert.ToInt32(obj)
        Else
            Return 0
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToIntNull(ByVal obj As Object) As System.Nullable(Of Integer)
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToInt32(obj)
        Else
            Return Nothing
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDateTimeNull(ByVal obj As Object) As System.Nullable(Of DateTime)
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToDateTime(obj)
        Else
            Return Nothing
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDoubleNull(ByVal obj As Object) As System.Nullable(Of Double)
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToDouble(obj)
        Else
            Return Nothing
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToDecimalNull(ByVal obj As Object) As System.Nullable(Of Decimal)
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToDecimal(obj)
        Else
            Return Nothing
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToLong(ByVal obj As Object) As Long
        If obj IsNot Nothing Then
            Return Convert.ToInt64(obj)
        Else
            Return 0
        End If
    End Function

    <System.Runtime.CompilerServices.Extension()> _
    Public Function ToLongNull(ByVal obj As Object) As System.Nullable(Of Long)
        If Not Convert.IsDBNull(obj) Then
            Return Convert.ToInt64(obj)
        Else
            Return Nothing
        End If
    End Function

    <Extension()> _
    Public Function ToIntSqkFK(ByVal pInt As Integer) As String
        If pInt = 0 Then
            Return "NULL"
        Else : Return pInt
        End If
    End Function
End Module