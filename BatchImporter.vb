Option Strict On

Module BatchImporter
    <STAThread>
    Sub Main()
        Dim SEApp As SolidEdgeFramework.Application = Nothing
        Dim SEDoc As SolidEdgeFramework.SolidEdgeDocument = Nothing
        Dim ProgramSettings As New Dictionary(Of String, String)
		Dim Filenames As List(Of String)
		Dim Filename As String
		Dim NewFilename As String
		Dim TemplateFilename As String
		Dim ImportFileExtension As String
		Dim ImportDirectory As String
		Dim ExportDirectory As String
		Dim TemplateExtension As String
		Dim IOErrorsMax As Integer
		Dim IOErrors As Integer = 0
		Dim SEWasRunning As Boolean = False

		OleMessageFilter.Register()

		ProgramSettings = GetProgramSettings()
		If ProgramSettings Is Nothing Then Exit Sub

		TemplateFilename = ProgramSettings("TemplateFilename")
		ImportFileExtension = ProgramSettings("ImportFileExtension")
		ImportDirectory = ProgramSettings("ImportDirectory")
		ExportDirectory = ProgramSettings("ExportDirectory")
		IOErrorsMax = CInt(ProgramSettings("IOErrorsMax"))

		If Not FileIO.FileSystem.DirectoryExists(ExportDirectory) Then
			FileIO.FileSystem.CreateDirectory(ExportDirectory)
		End If

		TemplateExtension = System.IO.Path.GetExtension(TemplateFilename)  ' C:\project\part.par -> .par

		Filenames = GetFilenames(ProgramSettings("ImportFileExtension"), ProgramSettings("ImportDirectory"))
		If Filenames Is Nothing Then Exit Sub


		Console.WriteLine("Connecting to Solid Edge")

		Try
			SEApp = CType(GetObject(, "SolidEdge.Application"), SolidEdgeFramework.Application)
			SEWasRunning = True
		Catch ex As Exception
			SEApp = CType(CreateObject("SolidEdge.Application"), SolidEdgeFramework.Application)
			SEWasRunning = False
		End Try

		SEApp.Visible = True
		SEApp.DisplayAlerts = False
		SEApp.WindowState = 2  'Maximizes Solid Edge
		SEApp.Activate()

		For Each Filename In Filenames

			If IOErrors > 3 Then
				MsgBox(String.Format("Number of file IO errors {0} exceed maximum of {1}.  Exiting...", IOErrors, IOErrorsMax), vbOKOnly)
				Exit For
			End If

			Console.WriteLine(String.Format("Opening {0}", Filename))
			NewFilename = ""

			Try
				SEDoc = DirectCast(SEApp.Documents.OpenWithTemplate(Filename, ProgramSettings("TemplateFilename")), SolidEdgeFramework.SolidEdgeDocument)
				SEApp.DoIdle()
				SEDoc.Activate()
			Catch ex As Exception
				Console.WriteLine(String.Format("Unable to open {0}", Filename))
				IOErrors += 1
				Continue For
			End Try

			NewFilename = System.IO.Path.GetFileName(Filename) ' C:\project\part.stp -> part.stp
			NewFilename = String.Format("{0}\{1}", ExportDirectory, NewFilename) ' part.stp -> .\OutDir\part.stp
			NewFilename = System.IO.Path.ChangeExtension(NewFilename, TemplateExtension) ' .\OutDir\part.stp -> .\OutDir\part.psm
			Console.WriteLine(String.Format("Saving {0}", NewFilename))

			Try
				SEDoc.SaveAs(NewFilename)
				SEApp.DoIdle()
				SEDoc.Close()
				SEApp.DoIdle()
			Catch ex As Exception
				Console.WriteLine(String.Format("Unable to save {0}", NewFilename))
				IOErrors += 1
				Continue For
			End Try

		Next

		If Not SEWasRunning Then
			SEApp.Quit()
		End If

		OleMessageFilter.Revoke()


	End Sub

	Private Function GetFilenames(ImportFileExtension As String, ImportDirectory As String) As List(Of String)
		Dim FoundFiles As New List(Of String)
		Dim ActiveFileExtensionsList As New List(Of String)
		ActiveFileExtensionsList.Add(ImportFileExtension)

		If FileIO.FileSystem.DirectoryExists(ImportDirectory) Then
			Try
				FoundFiles = FileIO.FileSystem.GetFiles(
					ImportDirectory,
					FileIO.SearchOption.SearchTopLevelOnly,
					ActiveFileExtensionsList.ToArray).ToList

			Catch ex As Exception
				Dim s As String = "An error occurred searching for files."
				s = String.Format("{0}{1}{2}", s, vbCrLf, ex.ToString)
				MsgBox(s, vbOKOnly)
				FoundFiles = Nothing
			End Try
		Else
			MsgBox(String.Format("Directory '{0}' not found", ImportDirectory), vbOKOnly)
			FoundFiles = Nothing
		End If

		If FoundFiles IsNot Nothing AndAlso FoundFiles.Count = 0 Then
			MsgBox(String.Format("No '{0}' files found", ImportFileExtension), vbOKOnly)
			FoundFiles = Nothing
		End If

		Return FoundFiles
	End Function

	Private Function GetProgramSettings() As Dictionary(Of String, String)
		Dim ProgramSettings As New Dictionary(Of String, String)
		Dim Settings As List(Of String) = Nothing
		Dim Key As String
		Dim Value As String
		Dim ProgramSettingsFilename As String
		Dim StartupPath As String = System.AppDomain.CurrentDomain.BaseDirectory
		Dim tf As Boolean
		Dim RequiredKeys As List(Of String) = {"TemplateFilename", "ImportFileExtension", "ImportDirectory", "ExportDirectory", "IOErrorsMax"}.ToList

		ProgramSettingsFilename = String.Format("{0}program_settings.txt", StartupPath)

		Try
			Settings = IO.File.ReadAllLines(ProgramSettingsFilename).ToList

			For Each KVPair As String In Settings

				Dim s As String = KVPair.Trim()

				tf = s = ""
				tf = tf OrElse s(0) = "'"
				tf = tf OrElse Not s.Contains("=")

				If tf Then Continue For

				Key = s.Split("="c)(0)
				Value = s.Split("="c)(1)

				ProgramSettings(Key.Trim()) = Value.Trim()
			Next

		Catch ex As Exception
			MsgBox(String.Format("Problem reading {0}", ProgramSettingsFilename), vbOKOnly)
			Return Nothing
		End Try

		Dim s1 As String = ""
		For Each s As String In RequiredKeys
			If Not ProgramSettings.Keys.Contains(s) Then
				s1 = String.Format("    {0}{1}{2}", s1, s, vbCrLf)
			End If
		Next

		If Not s1 = "" Then
			s1 = String.Format("The following variable names not found in program_settings.txt{0}{1}", vbCrLf, s1)
			MsgBox(s1, vbOKOnly)
			Return Nothing
		End If

		Return ProgramSettings

	End Function

End Module
