Imports System.Collections.Specialized
Imports System.Globalization
Imports System.Web

<System.Runtime.CompilerServices.Extension>
Public Module UriExtensions


    '''' <summary>
    '''' Converts the provided app-relative path into an absolute Url containing the 
    '''' full host name
    '''' </summary>
    '''' <param name="relativeUrl">App-Relative path</param>
    '''' <returns>Provided relativeUrl parameter as fully qualified Url</returns>
    '''' <example>~/path/to/foo to http://www.web.com/path/to/foo</example>
    '<System.Runtime.CompilerServices.Extension>
    'Public Function ToAbsoluteUrl(relativeUrl As String) As String
    '    If String.IsNullOrEmpty(relativeUrl) Then
    '        Return relativeUrl
    '    End If

    '    If HttpContext.Current Is Nothing Then
    '        Return relativeUrl
    '    End If

    '    If relativeUrl.StartsWith("/") Then
    '        relativeUrl = relativeUrl.Insert(0, "~")
    '    End If
    '    If Not relativeUrl.StartsWith("~/") Then
    '        relativeUrl = relativeUrl.Insert(0, "~/")
    '    End If

    '    Dim url as Uri = HttpContext.Current.Request.Url
    '    Dim port as String = If(url.Port <> 80, (":" + url.Port.ToString(CultureInfo.InvariantCulture)), [String].Empty)

    '    Return [String].Format("{0}://{1}{2}{3}", url.Scheme, url.Host, port, VirtualPathUtility.ToAbsolute(relativeUrl))
    'End Function


    <System.Runtime.CompilerServices.Extension>
    Public Function ModifyQueryString(baseUri As Uri, updates As NameValueCollection, removes As IEnumerable(Of String)) As String
        Dim query As NameValueCollection = HttpUtility.ParseQueryString(baseUri.Query)

        Dim url As String = baseUri.GetLeftPart(UriPartial.Path)

        If updates IsNot Nothing Then
            For Each key As String In updates.Keys
                query.[Set](key, updates(key))
            Next
        End If
        If removes IsNot Nothing Then
            For Each param As String In removes
                query.Remove(param)
            Next
        End If


        If query.HasKeys() Then
            Return String.Format("{0}?{1}", url, query.ToString())
        Else
            Return url
        End If
    End Function



    <System.Runtime.CompilerServices.Extension> _
    Public Function UpdateQueryParam(baseUri As Uri, param As String, value As Object) As String
        Dim collect = New NameValueCollection()
        collect.Add(param, value.ToString())
        Return baseUri.ModifyQueryString(collect, Nothing)
    End Function

    <System.Runtime.CompilerServices.Extension> _
    Public Function RemoveQueryParam(baseUri As Uri, param As String) As String
        Dim removes = New List(Of String)
        removes.Add(param)
        Return baseUri.ModifyQueryString(Nothing, removes)
    End Function

    <System.Runtime.CompilerServices.Extension> _
    Public Function UpdateQueryParams(baseUri As Uri, updates As NameValueCollection) As String
        Return baseUri.ModifyQueryString(updates, Nothing)
    End Function

    <System.Runtime.CompilerServices.Extension> _
    Public Function RemoveQueryParams(baseUri As Uri, removes As List(Of String)) As String
        Return baseUri.ModifyQueryString(Nothing, removes)
    End Function
End Module
