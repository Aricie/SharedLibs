﻿Imports System.Web.UI
Imports System.Web.UI.WebControls
Imports Aricie.DNN.UI.Attributes
Imports Aricie.DNN.ComponentModel
Imports Aricie.ComponentModel
Imports DotNetNuke.UI.WebControls
Imports Aricie.DNN.Security.Trial
Imports Aricie.DNN.Services
Imports Aricie.Services

Namespace UI.WebControls.EditControls


    


    Public Class PropertyEditorEditControl
        Inherits AricieEditControl
        Implements INamingContainer

        Private _InnerEditor As PropertyEditorControl
        Private _width As Unit
        Private _labelWidth As Unit
        Private _editControlWidth As Unit
        Private _EnableExports As Nullable(Of Boolean)

        Private _ActionButton As ActionButtonInfo

        Private Shared _Surrogates As New Dictionary(Of Type, IDynamicSurrogate)

        Shared Sub New()
            InitSurrogates()
        End Sub


        Private Shared Sub InitSurrogates()
            _Surrogates.Add(GetType(Type), New DynamicSurrogate(Of Type, DotNetType) _
                            With {.ConvertTo = (Function(objSource)
                                                    Dim toReturn As New DotNetType(objSource)
                                                    toReturn.TypeSelector = TypeSelector.NewType
                                                    Return toReturn
                                                End Function), _
                                 .ConvertFrom = (Function(objSurrogate)
                                                     If objSurrogate IsNot Nothing Then
                                                         Return objSurrogate.GetDotNetType()
                                                     End If
                                                     Return Nothing
                                                 End Function)})
        End Sub
       

        Public ReadOnly Property InnerEditor() As PropertyEditorControl
            Get
                Return _InnerEditor
            End Get
        End Property


        Protected Overrides Sub OnLoad(ByVal e As EventArgs)
            Me.EnsureChildControls()
            MyBase.OnLoad(e)
        End Sub

        Protected Overrides Sub OnDataChanged(ByVal e As EventArgs)
            Me.OnDataChanged()
        End Sub

        Private Overloads Sub OnDataChanged()
            'Me.EnsureChildControls()
            Dim args As New PropertyEditorEventArgs(Me.Name)
            If _DynamicSurrogate IsNot Nothing Then
                Dim newVal As Object = _DynamicSurrogate.ConvertFromSurrogate(Me._InnerEditor.DataSource)
                SaveCurrentSurrogateValue()
                'If ReflectionHelper.AreEqual(newVal, Me.Value) Then
                '    Exit Sub
                'End If
                args.Value = newVal
            Else
                args.Value = Me._InnerEditor.DataSource
            End If

            args.OldValue = Me.OldValue
            args.Changed = ReflectionHelper.AreEqual(Me.Value, Me.OldValue)
            args.StringValue = Me.StringValue
            MyBase.OnValueChanged(args)
        End Sub


        Protected Overrides Sub OnAttributesChanged()
            If (CustomAttributes IsNot Nothing) Then
                For Each attribute As Attribute In CustomAttributes
                    If TypeOf attribute Is FieldStyleAttribute Then
                        _width = DirectCast(attribute, FieldStyleAttribute).Width
                        _labelWidth = DirectCast(attribute, FieldStyleAttribute).LabelWidth
                        _editControlWidth = DirectCast(attribute, FieldStyleAttribute).EditControlWidth
                    ElseIf TypeOf attribute Is EnableExportsAttribute Then
                        Me._EnableExports = DirectCast(attribute, EnableExportsAttribute).Enabled
                    ElseIf TypeOf attribute Is ActionButtonAttribute Then
                        Me._ActionButton = ActionButtonInfo.FromAttribute(DirectCast(attribute, ActionButtonAttribute))
                    End If
                Next
            End If
        End Sub

        Protected Overridable Function GetNewEditor() As PropertyEditorControl
            Return New AriciePropertyEditorControl()
        End Function

        Protected Overrides Sub CreateChildControls()

            Try
                For Each objSurrogatePair As KeyValuePair(Of Type, IDynamicSurrogate) In _Surrogates
                    If (Me.Value IsNot Nothing AndAlso objSurrogatePair.Key.IsInstanceOfType(Me.Value)) _
                        OrElse (Me.Value Is Nothing AndAlso objSurrogatePair.Key.IsAssignableFrom(Me.ParentAricieField.AricieEditorInfo.PropertyType)) Then
                        _DynamicSurrogate = objSurrogatePair.Value
                    End If
                Next
                If Me.Value IsNot Nothing OrElse _DynamicSurrogate IsNot Nothing Then
                    Me._InnerEditor = GetNewEditor()
                    If TypeOf Me._InnerEditor Is AriciePropertyEditorControl Then
                        Dim aEditor As AriciePropertyEditorControl = TryCast(Me._InnerEditor, AriciePropertyEditorControl)
                        If Me.ParentAricieField IsNot Nothing Then
                            If Me.ParentAricieField.IsHidden Then
                                aEditor.IsHidden = True
                            End If
                        End If
                        If Me.ParentAricieEditor IsNot Nothing Then
                            aEditor.DisableExports = ParentAricieEditor.DisableExports
                            aEditor.PropertyDepth = Me.ParentAricieEditor.PropertyDepth + 1
                            aEditor.EnabledOnDemandSections = ParentAricieEditor.EnabledOnDemandSections AndAlso Me.OndemandEnabled
                            aEditor.TrialStatus = ParentAricieEditor.TrialStatus
                        Else
                            aEditor.PropertyDepth = 0
                        End If
                        If Me._EnableExports.HasValue Then
                            aEditor.DisableExports = Not Me._EnableExports.Value
                        End If
                    End If

                    Me._InnerEditor.ID = "pe"
                    If Me.ParentAricieField IsNot Nothing _
                            AndAlso TypeOf (Me.ParentAricieField.Editor) Is CollectionEditControl _
                            AndAlso CType(Me.ParentAricieField.Editor, CollectionEditControl).DisplayStyle = CollectionDisplayStyle.Accordion Then
                        Me.Controls.Add(Me._InnerEditor)
                    Else
                        Dim strCssClass As String = "odd"
                        If _ActionButton IsNot Nothing AndAlso TypeOf Me._InnerEditor Is AriciePropertyEditorControl Then
                            DirectCast(Me._InnerEditor, AriciePropertyEditorControl).ActionButton = Me._ActionButton
                        End If



                        Dim ariciePropCt As AriciePropertyEditorControl = TryCast(_InnerEditor, AriciePropertyEditorControl)
                        If ariciePropCt IsNot Nothing Then
                            If (ariciePropCt IsNot Nothing) Then
                                If (ariciePropCt.PropertyDepth Mod 2 = 0) Then
                                    strCssClass = "even"
                                End If
                            End If
                        End If
                        Me._InnerEditor.CssClass = strCssClass
                        strCssClass = String.Empty
                        Dim subPE As Control = AddSubDiv(Me._InnerEditor, strCssClass)
                        Me.Controls.Add(subPE)
                    End If


                    If Not _width = Unit.Empty Then
                        Me._InnerEditor.Width = _width
                    Else
                        'quel est l'interet de ce code ? un div prend automatiquement 100% de son conteneur.
                        'Dim newValue As Double = Math.Round(Me.ParentEditor.Width.Value - 20)
                        'Select Case Me.ParentEditor.Width.Type
                        '    Case UnitType.Percentage
                        '        Me._InnerEditor.Width = Unit.Percentage(newValue)
                        '    Case UnitType.Point
                        '        Me._InnerEditor.Width = Unit.Point(Convert.ToInt32(newValue))
                        '    Case UnitType.Pixel
                        '        Me._InnerEditor.Width = Unit.Pixel(Convert.ToInt32(newValue))
                        '    Case Else
                        '        Me._InnerEditor.Width = Unit.Parse(Me.ParentEditor.Width.ToString.Replace( _
                        '                                           Me.ParentEditor.Width.Value.ToString( _
                        '                                           CultureInfo.InvariantCulture), _
                        '                                           Convert.ToInt32(newValue).ToString( _
                        '                                        CultureInfo.InvariantCulture)))
                        'End Select

                    End If
                    'If Me.ParentEditor.CssClass = "ItemEven" Then
                    '    Me._InnerEditor.CssClass = "ItemOdd"
                    'Else
                    '    Me._InnerEditor.CssClass = "ItemEven"
                    'End If

                    'Me._InnerEditor.CssClass = 

                    Me._InnerEditor.LabelWidth = CType(IIf(_labelWidth = Unit.Empty, Me.ParentEditor.LabelWidth, _labelWidth), Unit)
                    Me._InnerEditor.EditControlWidth = CType(IIf(_editControlWidth = Unit.Empty, Me.ParentEditor.EditControlWidth, _editControlWidth), Unit)
                    Me._InnerEditor.EnableClientValidation = Me.ParentEditor.EnableClientValidation
                    Me._InnerEditor.ErrorStyle.CssClass = Me.ParentEditor.ErrorStyle.CssClass
                    Me._InnerEditor.GroupHeaderStyle.CssClass = Me.ParentEditor.GroupHeaderStyle.CssClass
                    Me._InnerEditor.GroupHeaderIncludeRule = Me.ParentEditor.GroupHeaderIncludeRule
                    Me._InnerEditor.HelpStyle.CssClass = Me.ParentEditor.HelpStyle.CssClass
                    Me._InnerEditor.LabelStyle.CssClass = Me.ParentEditor.LabelStyle.CssClass
                    Me._InnerEditor.VisibilityStyle.CssClass = Me.ParentEditor.VisibilityStyle.CssClass
                    Me._InnerEditor.GroupByMode = Me.ParentEditor.GroupByMode
                    Me._InnerEditor.DisplayMode = Me.ParentEditor.DisplayMode
                    Me._InnerEditor.EditMode = Me.EditMode
                    Me._InnerEditor.LocalResourceFile = Me.LocalResourceFile
                    Me._InnerEditor.HelpDisplayMode = Me.ParentEditor.HelpDisplayMode
                    Me._InnerEditor.ShowRequired = Me.ParentEditor.ShowRequired
                    Me._InnerEditor.ShowVisibility = Me.ParentEditor.ShowVisibility
                    Me._InnerEditor.SortMode = Me.ParentEditor.SortMode



                    'AddHandler Me._InnerEditor.ItemCreated, AddressOf Me.InnerEditor_ItemCreatedEventHandler



                    'AddHandler _datalist.ItemDataBound, AddressOf DatalistItemDataBound
                    'AddHandler _datalist.ItemCommand, AddressOf DatalistItemCommand

                    'Me._InnerEditor.DataSource = Value
                    'Me._InnerEditor.DataBind()

                    Me.DataBind()

                    'FormHelper.AddSection(Me, Me._InnerEditor, Me.Name)
                End If
            Finally
                Me.ChildControlsCreated = True
            End Try

        End Sub

        'Private Sub InnerEditor_ItemCreatedEventHandler(ByVal sender As Object, ByVal e As PropertyEditorItemEventArgs)

        '    'Me.ParentField.
        '    Me.OnItemChanged(sender, New PropertyEditorEventArgs(e.Editor.Name))

        'End Sub

        Private _DynamicSurrogate As IDynamicSurrogate

        Public ReadOnly Property UrlStateKey() As String
            Get
                Return "CtlUrlState-" & Me.ClientID.GetHashCode.ToString
            End Get
        End Property


        Private _CurrentSurrogateValue As Object

        Public ReadOnly Property CurrentSurrogateValue As Object
            Get
                If _CurrentSurrogateValue Is Nothing Then
                    'todo: this does not work with dotnettype since it does not serialize its entire state. Find another trick (session state? ughh...)
                    Dim serializedSurrogate As String = DnnContext.Current.AdvancedClientVariable(Me, UrlStateKey)
                    If Not serializedSurrogate.IsNullOrEmpty() Then
                        _CurrentSurrogateValue = ReflectionHelper.Deserialize(Of Serializable(Of Object))(serializedSurrogate).Value
                    Else
                        _CurrentSurrogateValue = _DynamicSurrogate.ConvertToSurrogate(Me.Value)
                    End If
                End If

                Return _CurrentSurrogateValue
            End Get
        End Property

        


        Public Overrides Sub DataBind()
            Dim objValue As Object
            If _DynamicSurrogate IsNot Nothing Then
                objValue = CurrentSurrogateValue
            Else
                objValue = Me.Value
            End If
            If objValue IsNot Nothing Then

                If Not ParentAricieEditor.BoundEntities.Contains(objValue) Then
                    If TypeOf _InnerEditor Is AriciePropertyEditorControl AndAlso DirectCast(_InnerEditor, AriciePropertyEditorControl).PropertyDepth < 10 Then
                        Me._InnerEditor.DataSource = objValue
                        Me._InnerEditor.DataBind()
                        'AddHandler _InnerEditor.ItemCreated, New EditorCreatedEventHandler(AddressOf Me.EditorItemCreated)
                        'MyBase.DataBind()
                    End If
                    If Me.Value IsNot Nothing Then
                        Me.ParentAricieEditor.BoundEntities.Add(objValue)
                    End If
                Else
                    Me.ParentAricieEditor.DisplayLocalizedMessage("CircularReferenceHidden.Message", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.YellowWarning)
                End If
            Else
                Me.ParentAricieEditor.DisplayLocalizedMessage("NullEntityHidden.Message", DotNetNuke.UI.Skins.Controls.ModuleMessage.ModuleMessageType.YellowWarning)
            End If
        End Sub


        Public Overrides Function LoadPostData(ByVal postDataKey As String, ByVal postCollection As System.Collections.Specialized.NameValueCollection) As Boolean
            Return _DynamicSurrogate IsNot Nothing
            'Return False
        End Function


        Protected Overrides Sub OnPreRender(ByVal e As EventArgs)
            If Me._DynamicSurrogate IsNot Nothing Then

                Me.OnDataChanged()
            End If
            MyBase.OnPreRender(e)

            If (Page IsNot Nothing) AndAlso Me.EditMode = PropertyEditorMode.Edit Then
                Me.Page.RegisterRequiresPostBack(Me)
            End If
        End Sub
        Protected Overrides Sub Render(ByVal writer As HtmlTextWriter)
            RenderChildren(writer)
        End Sub

        Private Sub SaveCurrentSurrogateValue()
            _CurrentSurrogateValue = Me._InnerEditor.DataSource
            Dim objSerialized As New Serializable(Of Object)(_CurrentSurrogateValue)
            DnnContext.Current.AdvancedClientVariable(Me, UrlStateKey) = ReflectionHelper.Serialize(objSerialized).Beautify()
        End Sub
        Public Overrides Sub EnforceTrialMode(ByVal mode As TrialPropertyMode)
            MyBase.EnforceTrialMode(mode)
            For Each subField As AricieFieldEditorControl In Me.InnerEditor.Fields
                Dim collectionEditControl = TryCast(subField.Editor, CollectionEditControl)
                If (collectionEditControl IsNot Nothing) Then
                    collectionEditControl.EnforceTrialMode(mode)
                End If
            Next
        End Sub
    End Class

End Namespace
