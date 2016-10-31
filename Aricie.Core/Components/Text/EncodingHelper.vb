Imports System.Globalization
Imports System.Text
Imports System.Web

Namespace Text
    ''' <summary>
    ''' Helper to retrieve and encoding based on the dedicated enumeration
    ''' </summary>
    Public Module EncodingHelper

        <System.Runtime.CompilerServices.Extension> _
        Public Function GetEncoding(objSimpleEncoding As SimpleEncoding) As Encoding
            Select Case objSimpleEncoding
                Case SimpleEncoding.Default
                    Return Encoding.Default
                Case SimpleEncoding.UTF8
                    Return Encoding.UTF8
                Case SimpleEncoding.ASCII
                    Return Encoding.ASCII
                Case SimpleEncoding.Unicode
                    Return Encoding.Unicode
                Case SimpleEncoding.UTF7
                    Return Encoding.UTF7
                Case SimpleEncoding.UTF32
                    Return Encoding.UTF32
                Case SimpleEncoding.BigEndianUnicode
                    Return Encoding.BigEndianUnicode
                Case Else
                    Return Encoding.Default
            End Select
        End Function

        <System.Runtime.CompilerServices.Extension> _
        Public Function HtmlEncode(source As String, method As HtmlEncodeMethod) As String
            Select Case method
                Case HtmlEncodeMethod.SecurityEscape
                    Return System.Security.SecurityElement.Escape(source)
                Case HtmlEncodeMethod.NumberEntities
                    Dim toReturn As New StringBuilder
                    For Each objChar As Char In source
                        toReturn.Append("&#")
                        toReturn.Append(AscW(objChar).ToString(CultureInfo.InvariantCulture))
                        toReturn.Append(";")
                    Next
                    Return toReturn.ToString()
            End Select
            Return HttpUtility.HtmlEncode(source)
        End Function

    End Module
End NameSpace