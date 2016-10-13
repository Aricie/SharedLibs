
Imports System.ComponentModel
Imports Aricie.DNN.UI.Attributes
Imports System.Threading

Namespace Services.Workers
    ''' <summary>
    ''' Information class for task queue
    ''' </summary>
    ''' <remarks></remarks>
    
    Public Class TaskQueueInfo

        Public Sub New()

        End Sub



        Public Sub New(ByVal nbThread As Integer, ByVal isBackground As Boolean, ByVal initialWaitTime As TimeSpan, ByVal wakeUpTime As TimeSpan, ByVal taskWaitTime As TimeSpan)
            Me._NbThreads = nbThread
            Me._IsBackground = isBackground
            Me._InitialWaitTime = New STimeSpan(initialWaitTime)
            Me._WakeUpWaitTime = New STimeSpan(wakeUpTime)
            Me._TaksWaitTime = New STimeSpan(taskWaitTime)
        End Sub

         <DefaultValue(False)> _
        <ExtendedCategory("MainSettings")> _
        Public Property Synchronous As Boolean


        ''' <summary>
        ''' Number of threads that will run the task
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("MainSettings")> _
        <ConditionalVisible("Synchronous", True)> _
         <DefaultValue(1)> _
        Public Property NbThreads() As Integer = 1


        ''' <summary>
        ''' Time to wait for task
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("MainSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property TaksWaitTime() As New STimeSpan

        ''' <summary>
        ''' Initial waiting time for task
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property InitialWaitTime() As New STimeSpan

        ''' <summary>
        ''' Waiting time before wake up
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property WakeUpWaitTime() As New STimeSpan





        ''' <summary>
        ''' Task in the background
        ''' </summary>
        ''' <value></value>
        ''' <returns></returns>
        ''' <remarks></remarks>
        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
         <DefaultValue(True)> _
        Public Property IsBackground() As Boolean = True

        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property IdleTimeout As New STimeSpan(TimeSpan.FromMilliseconds(100))

        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
         <DefaultValue(ApartmentState.Unknown)> _
        Public Property ApartmentState As ApartmentState = ApartmentState.Unknown

         <DefaultValue(ThreadPriority.Lowest)> _
        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property ThreadPriority() As ThreadPriority = ThreadPriority.Lowest

         <DefaultValue(False)> _
        <ExtendedCategory("AdvancedSettings")> _
        <ConditionalVisible("Synchronous", True)> _
        Public Property EnablePerformanceCounters As Boolean

         <DefaultValue("Aricie-STP")> _
        <ConditionalVisible("Synchronous", True)> _
        <ConditionalVisible("EnablePerformanceCounters", False, True)> _
        <ExtendedCategory("AdvancedSettings")> _
        Public Property PerformanceCounterInstanceName As String = "Aricie-STP"


    End Class
End Namespace