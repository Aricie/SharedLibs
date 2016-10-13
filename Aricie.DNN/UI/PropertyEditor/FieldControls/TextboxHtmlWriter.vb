Imports System.Web.UI

Public Class TextboxHtmlWriter
    Inherits FieldEditorHtmlWriter

    Public Sub New(parentControl As Control, writer As HtmlTextWriter, autoPostback As Boolean, passwordMode As Boolean)
        MyBase.New(parentControl, writer, autoPostback, passwordMode)
    End Sub

    Public Overrides Sub RenderBeginTag(ByVal objTagKey As System.Web.UI.HtmlTextWriterTag)
        If AutoPostBack Then
            If Not Me.IsAttributeDefined(HtmlTextWriterAttribute.Onchange) Then
                Dim onClick As String = GetStringPostBackRefrence()
                Me.AddAttribute(HtmlTextWriterAttribute.Onchange, onClick)
            End If
        End If
        BaseRenderBeginTag(objTagKey)
    End Sub
End Class
