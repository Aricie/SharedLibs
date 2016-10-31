


Namespace ComponentModel

    
    Public Class GenericEventArgs(Of T)
        Inherits EventArgs


        Private _Items As New Dictionary(Of Integer, T)

        Public Sub New()
            'Me.Item = ReflectionHelper.CreateObject(Of T)()
        End Sub

        Public Sub New(ByVal item As T)
            Me.Item = item
        End Sub

        Public Sub New(ByVal item1 As T, ByVal item2 As T)
            Me.New(item1)
            Me.Items(1) = item2
        End Sub

        Public Sub New(ByVal listItems As IList(Of T))
            Me.New(listItems(0))
            For index As Integer = 1 To listItems.Count - 1
                Me.Items(index) = listItems(index)
            Next
        End Sub


        Public Property Item() As T
            Get
                Return Me.Items(0)
            End Get
            Set(ByVal value As T)
                Me.Items(0) = value
            End Set
        End Property

        Public Property Items(ByVal index As Integer) As T
            Get
                If Me._Items.ContainsKey(index) Then
                    Return _Items(index)
                End If
                Return Nothing
            End Get
            Set(ByVal value As T)
                _Items(index) = value
            End Set
        End Property


    End Class


    ''' <summary>
    ''' G�n�ral purpose Eventargs with
    ''' </summary>
    
    Public Class ChangedEventArgs
        Inherits EventArgs

        Public Sub New()

        End Sub

        Public Sub New(ByVal oldValue As Object, ByVal newValue As Object)
            MyBase.New()
            Me._OldValue = oldValue
            Me._NewValue = newValue
        End Sub

        Private _OldValue As Object
        Private _NewValue As Object



        Public Property OldValue() As Object
            Get
                Return Me._OldValue
            End Get
            Set(ByVal value As Object)
                Me._OldValue = value
            End Set
        End Property

        Public Property NewValue() As Object
            Get
                Return Me._NewValue
            End Get
            Set(ByVal value As Object)
                Me._NewValue = value
            End Set
        End Property
    End Class

    
    Public Class ChangedEventArgs(Of T)
        Inherits GenericEventArgs(Of T)

        Public Sub New()
            MyBase.New()
        End Sub

        Public Sub New(ByVal oldValue As T, ByVal newValue As T)
            MyBase.New(oldValue, newValue)
        End Sub
        Public Sub New(ByVal oldValue As T, ByVal newValue As T, boolChangedAllowed As Boolean)
            MyBase.New(oldValue, newValue)
            Me._ChangedAllowed = boolChangedAllowed
        End Sub

        Private _ChangedAllowed As Boolean = True

        Public ReadOnly Property ChangedAllowed As Boolean
            Get
                Return _ChangedAllowed
            End Get
        End Property

        Public Property ChangedInEvent As Boolean


        Public ReadOnly Property OldValue() As T
            Get
                Return Me.Items(0)
            End Get
        End Property

        Public Property NewValue() As T
            Get
                Return Me.Items(1)
            End Get
            Set(ByVal value As T)
                If Me.ChangedAllowed Then
                    Me.ChangedInEvent = True
                    Me.Items(1) = value
                Else
                    Throw New ApplicationException("New Value not allowed for modification in Event args")
                End If
            End Set
        End Property
    End Class

End Namespace



