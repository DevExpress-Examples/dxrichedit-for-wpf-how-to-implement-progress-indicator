Imports System
Imports System.Windows
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit
'#Region "#usings"
Imports DevExpress.Services

'#End Region  ' #usings
Namespace ProgressIndicator

    ''' <summary>
    ''' Interaction logic for MainWindow.xaml
    ''' </summary>
    Public Partial Class MainWindow
        Inherits Window

        Public Sub New()
            Me.InitializeComponent()
        End Sub

        Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Me.richEditControl1.ApplyTemplate()
            Me.richEditControl1.LoadDocument("Docs\invitation.docx")
            Me.richEditControl1.Options.MailMerge.DataSource = New SampleData()
        End Sub

        Private Sub btnMailMerge_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
            Dim myMergeOptions As MailMergeOptions = Me.richEditControl1.Document.CreateMailMergeOptions()
            myMergeOptions.MergeMode = MergeMode.NewSection
            Me.richEditControl1.Document.MailMerge(myMergeOptions, Me.richEditControl2.Document)
            Me.tabControl.SelectedIndex = 1
        End Sub

        Private Sub richEditControl1_MailMergeStarted(ByVal sender As Object, ByVal e As MailMergeStartedEventArgs)
'#Region "#servicesubst"
            Me.richEditControl1.ReplaceService(Of IProgressIndicationService)(New MyProgressIndicatorService(Me.richEditControl1, Me.progressBarControl1))
'#End Region  ' #servicesubst
        End Sub

        Private Sub richEditControl1_MailMergeFinished(ByVal sender As Object, ByVal e As MailMergeFinishedEventArgs)
            Me.richEditControl1.RemoveService(GetType(IProgressIndicationService))
        End Sub

        Private Sub richEditControl1_MailMergeRecordStarted(ByVal sender As Object, ByVal e As MailMergeRecordStartedEventArgs)
            ' Imitating slow data fetching
            Threading.Thread.Sleep(100)
        End Sub

        Private Sub richEditControl1_MailMergeRecordFinished(ByVal sender As Object, ByVal e As MailMergeRecordFinishedEventArgs)
            e.RecordDocument.AppendDocumentContent("Docs\bungalow.docx", DocumentFormat.OpenXml)
        End Sub

        Private Sub tabControl_SelectionChanged(ByVal sender As Object, ByVal e As DevExpress.Xpf.Core.TabControlSelectionChangedEventArgs)
            Select Case Me.tabControl.SelectedIndex
                Case 0
                    Me.btnMailMerge.IsEnabled = True
                Case 1
                    Me.btnMailMerge.IsEnabled = False
            End Select
        End Sub
    End Class
End Namespace
