Imports System.IO
Imports System.IO.Compression
Imports Microsoft.VisualBasic.FileIO
Imports Word = Microsoft.Office.Interop.Word
Imports DocumentFormat.OpenXml
Imports DocumentFormat.OpenXml.Packaging
Imports DocumentFormat.OpenXml.Wordprocessing

Public Structure TableInfo
    Public TableStartPosition As Long
    Public TableEndPosition As Long
    Public TableRtfText As String
End Structure


Public Class Form1

    Public RTFConverter As RichTextBox
    Public Property wordApp As Word.Application
    Public timerCount As Integer

    Const baseFolder = "CLConverter"
    Const baseInputFolder = baseFolder + "\Input"
    Const baseOutputFolder = baseFolder + "\Output"
    Private Sub ChooseZip_Btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChooseZip_Btn.Click

        Me.Message_Lbl.Text = "Ready"

        Using cfdialog As New OpenFileDialog

            cfdialog.Title = "Choose zip archive"
            cfdialog.InitialDirectory = "C:\"
            cfdialog.Filter = "All files (*.*)|*.*|Zip files (*.zip)|*.zip"
            cfdialog.FilterIndex = 2
            cfdialog.RestoreDirectory = True
            cfdialog.Multiselect = True

            ' Do nothing if user does not click OK when choosing file
            If cfdialog.ShowDialog() <> DialogResult.OK Then Return

            Dim chosenFile As String

            ' And populate the listview with file names
            For Each chosenFile In cfdialog.FileNames
                Me.FilesList_Lview.Items.Add(chosenFile)
            Next chosenFile

        End Using

    End Sub

    Private Sub ClearList_Lbl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearList_Lbl.Click
        Me.FilesList_Lview.Clear()
        Me.Message_Lbl.Text = "Ready"
        Call ClearInOutFolders()
    End Sub

    Private Sub FilesList_Lview_DragDrop(sender As Object, e As DragEventArgs) Handles FilesList_Lview.DragDrop

        Me.Message_Lbl.Text = "Ready"

        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)

        For Each chosenFile In files

            If Directory.Exists(chosenFile) Then

                Dim containedFiles As String()

                containedFiles = Directory.GetFiles(chosenFile)

                For Each file In containedFiles

                    If Path.GetExtension(file).ToLower = ".zip" Then
                        Me.FilesList_Lview.Items.Add(file)
                    End If

                Next file

            Else
                Me.FilesList_Lview.Items.Add(chosenFile)
            End If

        Next chosenFile

    End Sub




    Private Sub FilesList_Lview_DragEnter(ByVal sender As Object, ByVal e As System.Windows.Forms.DragEventArgs) Handles FilesList_Lview.DragEnter
        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.Copy
        End If
    End Sub

    Private Sub FilesList_Lview_KeyUp(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles FilesList_Lview.KeyUp
        ' Use delete key to remove one entry in archives list (in ListView control)
        If Convert.ToInt32(e.KeyCode) = CInt(Keys.Delete) Then
            Me.FilesList_Lview.SelectedItems(0).Remove()
        End If
    End Sub

    Private Sub CheckCreateWorkingFolders()

        Dim hostFolder As String

        hostFolder = My.Settings.Default_App_WorkingFolder

        Me.Message_Lbl.Text = "Initializing..."
        Me.Refresh()

        ' If first usage, create base working directory in MyDocuments
        If Not Directory.Exists(Path.Combine(hostFolder, baseFolder)) Then
            Directory.CreateDirectory(Path.Combine(hostFolder, baseFolder))
        End If

        ' and same for input folder
        If Not Directory.Exists(Path.Combine(hostFolder, baseInputFolder)) Then
            Directory.CreateDirectory(Path.Combine(hostFolder, baseInputFolder))
        End If

        ' and output folder
        If Not Directory.Exists(Path.Combine(hostFolder, baseOutputFolder)) Then
            Directory.CreateDirectory(Path.Combine(hostFolder, baseOutputFolder))
        End If


    End Sub
    Private Sub Extract_AllRtfs()

        Dim zipArchive As ListViewItem

        Me.Message_Lbl.Text = "Extracting rtf files from archive(s)..."
        Me.Refresh()

        Dim hostFolder As String = Path.Combine(My.Settings.Default_App_WorkingFolder, baseInputFolder)

        ' Iterate all zips chosen and process them
        For Each zipArchive In Me.FilesList_Lview.Items

            Dim zipTempFolder As String
            zipTempFolder = Path.Combine(hostFolder, Path.GetFileNameWithoutExtension(zipArchive.Text.ToString))

            Using arch = ZipFile.Open(zipArchive.Text.ToString, ZipArchiveMode.Read)

                If Directory.Exists(zipTempFolder) Then
                    ' Clean before trying to extract
                    If Not Directory.GetFiles(zipTempFolder).Length = 0 Then
                        Call ClearInOutFolders()
                    End If
                End If

                If Not Directory.Exists(zipTempFolder) Then Directory.CreateDirectory(zipTempFolder)

                For Each zipEntry As ZipArchiveEntry In arch.Entries

                    If zipEntry.Name.EndsWith(".rtf", StringComparison.OrdinalIgnoreCase) Then
                        zipEntry.ExtractToFile(Path.Combine(zipTempFolder.ToString, zipEntry.Name))
                    End If

                Next

            End Using

            ' for one archive, increment by 10, for several, split and increment by equiv to their portion of desired increment - 10
            Me.ProgressBar1.Increment(10 / Me.FilesList_Lview.Items.Count)

            Application.DoEvents()

        Next

    End Sub

    Function Convert_RtfsToDocs_InAllFolders(InputFolders As String()) As String
        ' Return value: -1 for succes, 0 or 0,2,5 for input folders array numbers 0 & 2 & 5 being failed archive atempts of conversion

        Dim convertResult As String = ""

        'Me.Message_Lbl.Text = "Launching hidden Word app..."
        'Dim wordApp As New Word.Application

        Me.Message_Lbl.Text = "Converting all rtfs in docx format..."

        Dim archOutFolder As String

        ' process all folders which represent dearchived zip archives, open and convert to docx all files
        For Each inArchive In InputFolders

            ' Archive output folder, in app's output folder - a folder named as each archive
            archOutFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder, Path.GetFileNameWithoutExtension(inArchive))

            ' and create it for first usage in each archive
            If Not Directory.Exists(archOutFolder) Then Directory.CreateDirectory(archOutFolder)

            'Dim resultArchiveName As String

            ' Safety measure
            'resultArchiveName = Path.GetFileNameWithoutExtension(inArchive)

            Dim archFiles As String()
            archFiles = Directory.GetFiles(inArchive)

            Dim count As Integer = -1

            ' Process all files in folder (convert to docx)
            For Each archFile In archFiles

                count += 1

                Dim tmpDoc As Word.Document

                'Dim rtfConv As VariantType
                'rtfConv = wordApp.FileConverters("Rich Text Format (RTF)").OpenFormat

                tmpDoc = wordApp.Documents.Open(archFile.ToString, ConfirmConversions:=False, AddToRecentFiles:=False, Visible:=False)

                ' no more processing and converting if we couldnt open it
                If IsNothing(tmpDoc) Then
                    convertResult = IIf(convertResult = "", count, convertResult & "," & count)
                Else

                    ' print current file on message bar
                    Me.Message_Lbl.Text = Split(Me.Message_Lbl.Text, "...")(0) & "..." & " (" & tmpDoc.Name & ")"

                    Dim newFileName As String
                    ' rename files acc to provided scheme (11 chars then a dot and language code and new extension)
                    newFileName = Mid(Path.GetFileNameWithoutExtension(archFile.ToString), 1, 11) & "." & Mid(Path.GetFileNameWithoutExtension(archFile.ToString), 12, 2)
                    newFileName = Path.Combine(archOutFolder, newFileName & ".docx")

                    'If InStr(Path.GetFileNameWithoutExtension(archFile.ToString).ToLower, "en") > 0 Then resultArchiveName = Path.GetFileNameWithoutExtension(archFile.ToString) & ".zip"

                    tmpDoc.SaveAs2(newFileName, FileFormat:=Word.WdSaveFormat.wdFormatXMLDocument)
                    tmpDoc.Close()

                    Me.ProgressBar1.Increment((75 / InputFolders.Count) / archFiles.Count)

                End If

                Application.DoEvents()

            Next

        Next

        'Try
        '    wordApp.Quit()
        '    wordApp = Nothing
        'Catch
        '    MsgBox("Error " & Err.Number & " occured in " & Err.Source & " with description " & Err.Description)
        '    Err.Clear()
        'End Try

        If convertResult <> "" Then
            Return convertResult
        Else
            ' having reached so far, it's all right.
            Return "-1"
        End If

    End Function


    Function Rezip_All_Folders(InputFolders As String()) As Boolean

        Me.Message_Lbl.Text = "Rezipping resulting docx's..."

        Dim hostFolder As String = Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder)


        For Each inputFolder As String In InputFolders

            Dim resultArchiveName As String

            If Dir(inputFolder & "\*EN*.docx") <> "" Then
                resultArchiveName = Path.GetFileNameWithoutExtension(Dir(inputFolder & "\*EN*.docx")) & ".zip"
            Else
                resultArchiveName = Path.GetFileNameWithoutExtension(Dir(inputFolder & "\*.docx")) & ".zip"
            End If

            Dim archOutFilename As String = Path.Combine(hostFolder, resultArchiveName)

            ' 
            If File.Exists(archOutFilename) Then
                If MsgBox("Archive " & resultArchiveName & vbCr & vbCr & " already exists! OVERWRITE?", MsgBoxStyle.OkCancel, "Pre-existing archive") = vbOK Then
                    File.Delete(archOutFilename)
                    ZipFile.CreateFromDirectory(inputFolder, archOutFilename)
                Else
                    Me.Message_Lbl.Text = "Aborted"
                    Application.Exit()
                End If
            Else
                ZipFile.CreateFromDirectory(inputFolder, archOutFilename)
            End If

        Next inputFolder

        Return True

    End Function


    Private Sub Process_Btn_Click(sender As Object, e As EventArgs) Handles Process_Btn.Click

        Me.FilesList_Lview.UseWaitCursor = True

        ' uncover the progress bar
        Me.FilesList_Lview.Height = Me.FilesList_Lview.Height - 12

        Call CheckCreateWorkingFolders()

        Dim hostFolder As String

        hostFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseInputFolder)

        Me.ProgressBar1.Increment(5)

        ' extract all rtf files from archive(s) in listview control
        Call Extract_AllRtfs()

        ' Correct for integer division leftover accidents, would cause the progress bar not to be at desired value (15 ie)
        If Me.ProgressBar1.Value < 15 Then Me.ProgressBar1.Value = 15
        Me.Message_Lbl.Text = "Extracting rtf files done"


        Dim inputArchivesDirs As String()

        inputArchivesDirs = Directory.GetDirectories(hostFolder)


        Dim archivesProcessResult As String

        archivesProcessResult = Convert_RtfsToDocs_InAllFolders(inputArchivesDirs)
        'archivesProcessResult = Convert_RtfsToDocs_wRTFBox_InAllFolders(inputArchivesDirs)


        '' HIGHLY EXPERIMENTAL, try file conversion with Open XML SDK 2.5 !
        'archivesProcessResult = Convert_RtfsToDocs_InAllFolders_OXML(inputArchivesDirs)

        If Not archivesProcessResult = "-1" Then     ' Success
            MsgBox("Processing archive(s) " & archivesProcessResult & " failed!", vbOKOnly + MsgBoxStyle.Information)
        End If


        hostFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder)


        'Dim embedConvertResult As String

        'embedConvertResult = Open_andSave_DocxFiles_inFolders(Directory.GetDirectories(hostFolder))

        'If Not embedConvertResult = "-1" Then
        '    MsgBox("Processing embedded docxs " & embedConvertResult & " failed!", vbOKOnly + MsgBoxStyle.Information)
        'End If


        Call Rezip_All_Folders(Directory.GetDirectories(hostFolder))

        ' Launch explorer with output folder if user chose it
        If My.Settings.App_Open_OutputFolder Then Process.Start("explorer.exe", hostFolder)

        ' and also clean folders of temp files if user so desires
        If My.Settings.App_AutoClean_Folders Then
            Me.Message_Lbl.Text = "Auto-Cleaning app folders..."
            Call ClearInOutFolders()
        End If


        If Me.ProgressBar1.Value <> 100 Then Me.ProgressBar1.Value = 100

        Me.FilesList_Lview.UseWaitCursor = False

        Me.ProgressBar1.Value = 0

        ' cover the progress bar again... not needed
        Me.FilesList_Lview.Height = Me.FilesList_Lview.Height + 12

        Me.Message_Lbl.Text = "All done"


        'Dim ddocx As Novacode.DocX = Novacode.DocX.Load()
        'ddocx.c()
        'OpenXmlPowerTools.DocumentAssembler.()

    End Sub

    Function Convert_RtfsToDocs_wRTFBox_InAllFolders(InputFolders As String()) As String

        Return ""

    End Function

    Private Function Open_andSave_DocxFiles_inFolders(InputFolders As String()) As String
        ' Return -1 for all files opened and saved without errors.
        ' Return 0,4,5 for instance for files numbers 1, 3 and 4 - counted from 0, as having returned errors
        ' on open and or save!

        Me.Message_Lbl.Text = "Converting embedded rtfs..."

        Dim returnResults As String = ""

        For Each inputFolder In InputFolders

            Dim inputFiles As String()

            inputFiles = Directory.GetFiles(inputFolder)

            Dim countFile As Integer = -1   ' We want it 0 on first file

            For Each inputFile In inputFiles

                countFile += 1

                If Not Open_andSave_DocxFile(inputFile) Then

                    returnResults = IIf(returnResults = "", countFile.ToString, returnResults & "," & countFile.ToString)

                End If

            Next

        Next

        ' Return success if we didn't already save 
        If returnResults = "" Then returnResults = "-1"

        Return returnResults

    End Function

    Public Function Open_andSave_DocxFile(inputFile) As Boolean

        Dim tmpDoc As Word.Document

        Try
            tmpDoc = wordApp.Documents.Open(inputFile)

            tmpDoc.Save()

            tmpDoc.Close()
        Catch
            If Err.Number <> 0 Then
                Err.Clear()
                Return False
            End If
        End Try

        Return True

    End Function

    Private Function Convert_RtfsToDocs_InAllFolders_OXML(InputFolders As String()) As String

        Me.Message_Lbl.Text = "Converting rtfs to docx..."

        Application.DoEvents()

        For Each inputFolder As String In InputFolders

            Dim rtfFilenames As String() = Directory.GetFiles(inputFolder)

            ' output to a folder named same as input folder, but in output folder, course
            Dim outputFolder As String = Path.GetFileName(inputFolder)

            outputFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder, outputFolder)

            If Not Directory.Exists(outputFolder) Then Directory.CreateDirectory(outputFolder)

            For Each rtfFilename As String In rtfFilenames

                Me.Message_Lbl.Text = Split(Me.Message_Lbl.Text, "...")(0) & "... " & Path.GetFileNameWithoutExtension(rtfFilename)

                Application.DoEvents()

                Dim docxFilename As String


                ' rename files acc to provided scheme (11 chars then a dot and language code and new extension)
                docxFilename = Mid(Path.GetFileNameWithoutExtension(rtfFilename), 1, 11) & "." & Mid(Path.GetFileNameWithoutExtension(rtfFilename), 12, 2)
                docxFilename = Path.Combine(outputFolder, docxFilename & ".docx")


                'Call CreateWordProcessingDocument(docxFilename)
                Call CreateDocx_andInsertRtf(docxFilename, rtfFilename)

                'If Not Convert_Rtf_ToDocx(rtfFilename) Then
                '    MsgBox("This rtf failed to convert to docx" & rtfFilename, vbOKOnly)
                '    Return 0
                'End If

                Me.ProgressBar1.Increment((75 / InputFolders.Count) / rtfFilenames.Count)

            Next

        Next

        Return "-1" ' success

    End Function

    Public Sub CreateDocx_andInsertRtf(ByVal WordProcessingDoc_Filename As String, ByVal RtfFilename As String)

        Dim newDocxCreationDone As Boolean

        newDocxCreationDone = CreateWordProcessingDocument(WordProcessingDoc_Filename)

        Dim insertRtfDone As Boolean

        'insertRtfDone = Insert_Rtf_ToDocx(WordProcessingDoc_Filename, RtfFilename)
        insertRtfDone = Insert_Rtf_ToDocx_RTB(WordProcessingDoc_Filename, RtfFilename)

    End Sub

    Public Function CreateWordProcessingDocument(ByVal filePath As String) As Boolean


        Using wordDocument As WordprocessingDocument = WordprocessingDocument.Create(filePath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document)

            Dim mainPart As MainDocumentPart = wordDocument.AddMainDocumentPart

            mainPart.Document = New Document

            Dim body As Body = mainPart.Document.AppendChild(New Body)

            Dim para As Paragraph = body.AppendChild(New Paragraph)

            Dim run As Run = para.AppendChild(New Run)

            run.AppendChild(New Text(""))

        End Using

        Return True

    End Function

    Public Sub OpenAndAddTextToWordDocument(ByVal filepath As String, ByVal txt As String())
        ' Open a WordprocessingDocument for editing using the filepath.
        Dim wordprocessingDocument As WordprocessingDocument = wordprocessingDocument.Open(filepath, True)

        Call InsertParagraphs_ToDocument(wordprocessingDocument, txt)

        wordprocessingDocument.MainDocumentPart.Document.Save()

        ' Close the handle explicitly.
        wordprocessingDocument.Close()
    End Sub

    Public Sub InsertParagraphs_ToDocument(Document As WordprocessingDocument, Paragraphs As String())

        ' Assign a reference to the existing document body. 
        Dim body As Body = Document.MainDocumentPart.Document.Body

        Dim para As Paragraph

        For Each inputPara In Paragraphs

            ' Add new text.
            para = body.AppendChild(New Paragraph)
            Dim run As Run = para.AppendChild(New Run)
            run.AppendChild(New Text(inputPara))

        Next inputPara

    End Sub


    Public Sub OpenAndAddTextToWordDocument_Docx(ByVal filepath As String, ByVal txt As String)

        Using docxDoc As Novacode.DocX = Novacode.DocX.Load(filepath)

            ' Open a WordprocessingDocument for editing using the filepath.

            docxDoc.InsertParagraphs(txt)

            docxDoc.Save()

        End Using

    End Sub

    Private Function Insert_Rtf_ToDocx(WordProcessingDocFilename As String, ByVal RtfToImportFilename As String) As Boolean

        'Dim wordDocument As WordprocessingDocument = WordProcessingDocument.Create(Path.Combine(Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder), Path.GetFileNameWithoutExtension(InputRtfFilename)) & ".docx", DocumentFormat.OpenXml.WordprocessingDocumentType.Document)

        Using wordDoc As WordprocessingDocument = WordprocessingDocument.Open(WordProcessingDocFilename, True)

            Dim altChunkId = "AltChunkId5"

            Dim mainDocPart As MainDocumentPart = wordDoc.MainDocumentPart

            Dim chunk As AlternativeFormatImportPart = mainDocPart.AddAlternativeFormatImportPart(AlternativeFormatImportPartType.Rtf, altChunkId)

            Dim rtfDocumentContent = File.ReadAllText(RtfToImportFilename, System.Text.Encoding.UTF8)

            Using ms As MemoryStream = New MemoryStream(System.Text.Encoding.UTF8.GetBytes(rtfDocumentContent))
                chunk.FeedData(ms)
            End Using

            Dim altChunk = New Wordprocessing.AltChunk

            altChunk.Id = altChunkId

            mainDocPart.Document.Body.InsertAfter(altChunk, mainDocPart.Document.Body.Elements(Of Wordprocessing.Paragraph)().Last)

            mainDocPart.Document.Save()

            wordDoc.Close()

        End Using

        Return True

    End Function

    Private Function Insert_Rtf_ToDocx_RTB(WordProcessingDocFilename As String, ByVal RtfToImportFilename As String) As Boolean

        Dim rtfDocumentContent = getText_wRTFTextbox(RtfToImportFilename)

        Call OpenAndAddTextToWordDocument(WordProcessingDocFilename, rtfDocumentContent)

        Return True

    End Function

    Function getText_wRTFTextbox(Inputfilename As String) As String()

        If IsNothing(Me.RTFConverter) Then Me.RTFConverter = New RichTextBox

        Me.RTFConverter.LoadFile(Inputfilename)

        getText_wRTFTextbox = Split(Me.RTFConverter.Text, vbLf)

        Me.RTFConverter.Clear()

    End Function

    Public Function countSubString(ByVal inputString As String, ByVal substringToSearch As String) As Integer
        Return System.Text.RegularExpressions.Regex.Split(inputString, substringToSearch).Length - 1
    End Function


    Function getTableRows_Positions(FileContent As String) As TableInfo()

        Const tableRowRtfMarker As String = "\trowd"

        Dim tempResults As TableInfo()

        ' do we have tables inside?
        If countSubString(FileContent, tableRowRtfMarker) > 0 Then

            ReDim tempResults(0)

            tempResults(0).TableStartPosition = InStr(1, FileContent, tableRowRtfMarker)

            If countSubString(FileContent, tableRowRtfMarker) = 2 Then
                tempResults(0).TableEndPosition = InStr(CInt(tempResults(0).TableStartPosition) + 6, FileContent, tableRowRtfMarker)
            End If
            End If

    End Function


    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing

        Try
            If Not IsNothing(Me.wordApp) Then Me.wordApp.Quit()
            Me.wordApp = Nothing
        Catch
            MsgBox("Bdk: Error " & Err.Number & " occured in " & Err.Source & " with description " & Err.Description)
            Err.Clear()
        End Try


        If My.Settings.App_AutoClean_Folders Then
            Call ClearInOutFolders()
        End If

    End Sub


    Function ClearInOutFolders() As Boolean

        Dim absInputFolder As String
        Dim absOutputFolder As String

        If Directory.Exists(Path.Combine(My.Settings.Default_App_WorkingFolder, baseInputFolder)) Then

            absInputFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseInputFolder)

            Dim inputSubDirectories As String()

            inputSubDirectories = Directory.GetDirectories(absInputFolder)

            ' Lets see if it works deleting all subfolders
            If inputSubDirectories.Length > 0 Then
                For Each subF In inputSubDirectories
                    Directory.Delete(subF, True)
                Next
            End If

            absOutputFolder = Path.Combine(My.Settings.Default_App_WorkingFolder, baseOutputFolder)

            Dim outputSubDirectories As String()

            outputSubDirectories = Directory.GetDirectories(absOutputFolder)

            ' Lets see if it works deleting all subfolders
            If outputSubDirectories.Length > 0 Then
                For Each subF In outputSubDirectories
                    Try
                        Directory.Delete(subF, True)
                    Catch
                        Err.Clear()
                    End Try
                Next
            End If

            If Directory.EnumerateDirectories(absInputFolder).Count > 0 And Directory.EnumerateDirectories(absOutputFolder).Count > 0 Then
                Return True
            Else
                Return False
            End If

        Else
            Return True
        End If

    End Function

    Private Sub ChooseWorkingFolder_Lbl_Click(sender As Object, e As EventArgs) Handles ChooseWorkingFolder_Lbl.Click

        If Me.WorkingFolderBrowserDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
            Me.CtWorkingFolder_TxtBox.Text = Me.WorkingFolderBrowserDialog.SelectedPath
            My.Settings.Default_App_WorkingFolder = Me.WorkingFolderBrowserDialog.SelectedPath
        End If

    End Sub


    Private Sub ToggleOptions_Lbl_Click(sender As Object, e As EventArgs) Handles ToggleOptions_Lbl.Click

        If Me.Width < 520 Then

            Me.Width = 830  ' Open options part of form (widen it)
        Else    ' Close it
            Me.Width = 510
        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load


        Dim ad As System.Deployment.Application.ApplicationDeployment = System.Deployment.Application.ApplicationDeployment.CurrentDeployment
        Dim sVersion As String

        Me.RTFConverter = New RichTextBox

        'Dim appVersionInfo As FileVersionInfo
        'appVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        sVersion = ad.CurrentVersion.ToString

        Me.Text = Me.Text & " (" & sVersion & ")"

        Me.Width = 510  ' Start form collapsed

        If Directory.Exists(My.Settings.Default_App_WorkingFolder) Then
            Me.CtWorkingFolder_TxtBox.Text = Path.Combine(My.Settings.Default_App_WorkingFolder, baseFolder)
            'Debug.Print("Default_App_WorkingFolder existed as " & My.Settings.Default_App_WorkingFolder)
        Else
            If MsgBox("Since this your first time running the application, we need to choose the working folder", vbOKOnly, _
                "Setting working folder") = MsgBoxResult.Ok Then

                If Me.WorkingFolderBrowserDialog.ShowDialog = System.Windows.Forms.DialogResult.OK Then

                    My.Settings.Default_App_WorkingFolder = Me.WorkingFolderBrowserDialog.SelectedPath
                    'Debug.Print("Just changed Default_App_WorkingFolder setting to " & My.Settings.Default_App_WorkingFolder)
                    Me.CtWorkingFolder_TxtBox.Text = Path.Combine(My.Settings.Default_App_WorkingFolder, baseFolder)

                Else
                    Application.Exit()
                End If
            Else
                Application.Exit()
            End If
        End If

        ' and restore user saved settings for our check-boxes
        Me.OpenOutputFolder_Cbox.Checked = My.Settings.App_Open_OutputFolder
        Me.AutoClean_Folders_Cbox.Checked = My.Settings.App_AutoClean_Folders

        '' Launch Word app upon launch, so as to make "Process" button faster
        'Me.Message_Lbl.Text = "Starting Word Application..."
        'Me.wordApp = New Word.Application
        'Me.wordApp.Visible = False
        'Me.wordApp.Options.SaveNormalPrompt = False
        'Me.Message_Lbl.Text = "Ready"

        'If IsNothing(Me.wordApp) Then
        '    MsgBox("Word app initialization failed !")
        '    End
        'End If


    End Sub

    Private Sub OpenOutputFolder_Cbox_Click(sender As Object, e As EventArgs) Handles OpenOutputFolder_Cbox.Click
        If Me.OpenOutputFolder_Cbox.Checked Then
            My.Settings.App_Open_OutputFolder = True
        Else
            My.Settings.App_Open_OutputFolder = False
        End If
    End Sub

    Private Sub AutoClean_Folders_Cbox_Click(sender As Object, e As EventArgs) Handles AutoClean_Folders_Cbox.Click
        If Me.AutoClean_Folders_Cbox.Checked Then
            My.Settings.App_AutoClean_Folders = True
        Else
            My.Settings.App_AutoClean_Folders = False
        End If
    End Sub

    Private Sub Timer1_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        timerCount = timerCount + 1
        If IsNothing(Me.wordApp) Then
            Me.Cursor = Cursors.AppStarting
            Me.Message_Lbl.Text = "Starting Word application..."
            Me.wordApp = New Word.Application
        End If

        If timerCount = 2 Then
            Me.Cursor = Cursors.Default
            Timer1.Stop()
            Me.Message_Lbl.Text = "Ready"
        End If
    End Sub

    Private Sub RichTextBox1_Click(sender As Object, e As EventArgs)

        Using cfdialog As New OpenFileDialog

            cfdialog.Title = "Choose rtf file"
            cfdialog.InitialDirectory = "C:\"
            cfdialog.Filter = "All files (*.*)|*.*|rtf files (*.rtf)|*.rtf"
            cfdialog.FilterIndex = 2
            cfdialog.RestoreDirectory = True
            cfdialog.Multiselect = False

            ' Do nothing if user does not click OK when choosing file
            If cfdialog.ShowDialog() <> DialogResult.OK Then Return

            ' And populate the listview with file names
            Me.RTFConverter.LoadFile(cfdialog.FileName)

        End Using

    End Sub

End Class