Imports System
Imports System.Windows
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

#Region "#usings"
Imports DevExpress.Services
#End Region ' #usings

Namespace ProgressIndicator
	''' <summary>
	''' Interaction logic for MainWindow.xaml
	''' </summary>
	Partial Public Class MainWindow
		Inherits Window

		Public Sub New()
			InitializeComponent()
		End Sub

		Private Sub richEditControl1_Loaded(ByVal sender As Object, ByVal e As RoutedEventArgs)
			richEditControl1.ApplyTemplate()
			richEditControl1.LoadDocument("Docs\invitation.docx")
			richEditControl1.Options.MailMerge.DataSource = New SampleData()
		End Sub

		Private Sub btnMailMerge_Click(ByVal sender As Object, ByVal e As RoutedEventArgs)
			Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
			myMergeOptions.MergeMode = MergeMode.NewSection
			richEditControl1.Document.MailMerge(myMergeOptions, richEditControl2.Document)
			tabControl.SelectedIndex = 1
		End Sub

		Private Sub richEditControl1_MailMergeStarted(ByVal sender As Object, ByVal e As MailMergeStartedEventArgs)
'			#Region "#servicesubst"
		   richEditControl1.ReplaceService(Of IProgressIndicationService) (New MyProgressIndicatorService(richEditControl1, Me.progressBarControl1))
'			#End Region ' #servicesubst
		End Sub

		Private Sub richEditControl1_MailMergeFinished(ByVal sender As Object, ByVal e As MailMergeFinishedEventArgs)
			richEditControl1.RemoveService(GetType(IProgressIndicationService))
		End Sub

		Private Sub richEditControl1_MailMergeRecordStarted(ByVal sender As Object, ByVal e As MailMergeRecordStartedEventArgs)
			' Imitating slow data fetching
			System.Threading.Thread.Sleep(100)
		End Sub

		Private Sub richEditControl1_MailMergeRecordFinished(ByVal sender As Object, ByVal e As MailMergeRecordFinishedEventArgs)
			e.RecordDocument.AppendDocumentContent("Docs\bungalow.docx", DocumentFormat.OpenXml)
		End Sub

		Private Sub tabControl_SelectionChanged(ByVal sender As Object, ByVal e As DevExpress.Xpf.Core.TabControlSelectionChangedEventArgs)
			Select Case tabControl.SelectedIndex
				Case 0
					Me.btnMailMerge.IsEnabled = True
				Case 1
					Me.btnMailMerge.IsEnabled = False
			End Select
		End Sub
	End Class
End Namespace
