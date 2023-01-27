' Developer Express Code Central Example:
' How to implement progress indicator for slow RichEditControl operations
' 
' This example demonstrates how to utilize the IProgressIndicationService service
' interface of the RichEditControl. The ProgressBar control on the form indicates
' how the mail merge operation proceeds.
' 
' You can find sample updates and versions for different programming languages here:
' http://www.devexpress.com/example=E3075
Imports System
Imports DevExpress.Services
Imports DevExpress.Xpf.Editors
Imports System.Windows.Threading

Namespace ProgressIndicator

'#Region "#myprogressindicator"
    Friend Class MyProgressIndicatorService
        Implements IProgressIndicationService

        Private _Indicator As ProgressBarEdit

        Public Property Indicator As ProgressBarEdit
            Get
                Return _Indicator
            End Get

            Set(ByVal value As ProgressBarEdit)
                _Indicator = value
            End Set
        End Property

        Public Sub New(ByVal provider As IServiceProvider, ByVal indicator As ProgressBarEdit)
            _Indicator = indicator
        End Sub

'#Region "IProgressIndicationService Members"
        Private Sub Begin(ByVal displayName As String, ByVal minProgress As Integer, ByVal maxProgress As Integer, ByVal currentProgress As Integer) Implements IProgressIndicationService.Begin
            _Indicator.Minimum = minProgress
            _Indicator.Maximum = maxProgress
            _Indicator.EditValue = currentProgress
            _Indicator.Visibility = System.Windows.Visibility.Visible
            Refresh()
        End Sub

        Private Sub [End]() Implements IProgressIndicationService.End
            _Indicator.Visibility = System.Windows.Visibility.Collapsed
            Refresh()
        End Sub

        Private Sub SetProgress(ByVal currentProgress As Integer) Implements IProgressIndicationService.SetProgress
            _Indicator.EditValue = currentProgress
            Refresh()
        End Sub

'#End Region
        Private Sub Refresh()
            Dim emptyDelegate As Action = Sub()
            End Sub
            _Indicator.Dispatcher.Invoke(DispatcherPriority.Render, emptyDelegate)
        End Sub
    End Class
'#End Region  ' #myprogressindicator
End Namespace
