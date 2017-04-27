Imports System.IO
Imports System.IO.Compression
Imports Microsoft.VisualBasic.FileIO
Imports Word = Microsoft.Office.Interop.Word


Public Structure TableInfo
    Public TableStartPosition As Long
    Public TableEndPosition As Long
    Public TableRtfText As String
End Structure


Public Class Form1

    Private Const ConvertRtfsMessage As String = "Converting rtfs to docx..."

    Private Property Options_Opened As Boolean = False

    Public RTFConverter As RichTextBox
    Public Property wordApp As Word.Application
    Public timerCount As Integer

    Public zipIcons As New ImageList

    Const baseFolder = "CLConverter"
    Const baseInputFolder = baseFolder + "\Input"
    Const baseOutputFolder = baseFolder + "\Output"


    Private Sub ChooseZip_Btn_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ChooseZip_Btn.Click

        Me.Message_Lbl.Text = "Ready"

        Me.FilesList_Lview.SmallImageList = zipIcons
        'Me.FilesList_Lview.LargeImageList = zipIcons

        Using cfdialog As New OpenFileDialog

            Dim browsingDirectory As String
            browsingDirectory = My.Settings.Default_BrowseTo_Folder     ' Initialize quand-meme to C:\

            ' But change if user desired
            If Me.Default_BrowseToFolder_Cbox.Checked Then
                browsingDirectory = My.Settings.Default_BrowseTo_Folder
            Else    ' or just remember current folder
                If Me.ChooseZip_Btn.Tag <> "" Then
                    browsingDirectory = Me.ChooseZip_Btn.Tag
                End If
            End If

            cfdialog.Title = "Please select zip archive"
            cfdialog.InitialDirectory = browsingDirectory
            cfdialog.Filter = "All files (*.*)|*.*|Zip files (*.zip)|*.zip"
            cfdialog.FilterIndex = 2
            cfdialog.RestoreDirectory = True
            cfdialog.Multiselect = True

            ' Do nothing if user does not click OK when choosing file
            If cfdialog.ShowDialog() <> DialogResult.OK Then Return

            Dim chosenFile As String


            ' Save last used directory
            Me.ChooseZip_Btn.Tag = Path.GetDirectoryName(cfdialog.FileNames(0))

            ' And populate the listview with file names
            For Each chosenFile In cfdialog.FileNames


                If Not FileAlreadyAdded(Path.GetFileName(chosenFile)) Then

                    Dim newLviewItem As New ListViewItem

                    newLviewItem.Text = Path.GetFileName(chosenFile)

                    newLviewItem.ImageIndex = 1
                    newLviewItem.Tag = chosenFile

                    Me.FilesList_Lview.Items.Add(newLviewItem)

                    ' Save complete 

                End If

            Next chosenFile

        End Using

    End Sub

    Private Function FileAlreadyAdded(ByVal Filename As String) As Boolean

        If FilesList_Lview.Items.Count = 0 Then Return False

        Dim listItem As ListViewItem

        For Each listItem In FilesList_Lview.Items

            If listItem.Text = Filename Then
                Return True
            End If

        Next listItem

        Return False

    End Function

    Private Sub ClearList_Lbl_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles ClearList_Lbl.Click
        Me.FilesList_Lview.Clear()
        Me.Message_Lbl.Text = "Ready"
        Call ClearInOutFolders()
    End Sub

    Private Sub FilesList_Lview_DragDrop(sender As Object, e As DragEventArgs) Handles FilesList_Lview.DragDrop

        Me.Message_Lbl.Text = "Ready"

        Me.FilesList_Lview.SmallImageList = zipIcons

        Dim files() As String = e.Data.GetData(DataFormats.FileDrop)


        For Each chosenFile In files

            If Directory.Exists(chosenFile) Then

                Dim containedFiles As String()

                containedFiles = Directory.GetFiles(chosenFile)

                For Each file In containedFiles

                    If Path.GetExtension(file).ToLower = ".zip" Then

                        If Not FileAlreadyAdded(Path.GetFileName(file)) Then

                            Dim newLviewItem As New ListViewItem

                            newLviewItem.Text = Path.GetFileName(file)
                            newLviewItem.ImageIndex = 1
                            newLviewItem.Tag = file

                            Me.FilesList_Lview.Items.Add(newLviewItem)
                            ' Save complete 

                        End If

                    End If

                Next file

            Else    ' no folder selected, just one or more files

                If Not FileAlreadyAdded(Path.GetFileName(chosenFile)) Then

                    Dim newLviewItemF As New ListViewItem

                    newLviewItemF.Text = Path.GetFileName(chosenFile)
                    newLviewItemF.ImageIndex = 1
                    newLviewItemF.Tag = chosenFile

                    Me.FilesList_Lview.Items.Add(newLviewItemF)

                End If

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
            zipTempFolder = Path.Combine(hostFolder, Path.GetFileNameWithoutExtension(zipArchive.Tag.ToString))

            Using arch = ZipFile.Open(zipArchive.Tag.ToString, ZipArchiveMode.Read)

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

        Me.Message_Lbl.Text = ConvertRtfsMessage

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

                    If Len(Path.GetFileNameWithoutExtension(archFile.ToString)) > 12 Then

                        ' rename files acc to provided scheme (11 chars then a dot and language code and new extension)
                        newFileName = Mid(Path.GetFileNameWithoutExtension(archFile.ToString), 1, 11) & "." & Mid(Path.GetFileNameWithoutExtension(archFile.ToString), 12, 2)
                        newFileName = Path.Combine(archOutFolder, newFileName & ".docx")

                    Else        ' should name not seem proper, leave it such as it is
                        newFileName = Path.Combine(archOutFolder, Path.GetFileNameWithoutExtension(archFile.ToString) & ".docx")
                    End If

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

    Private Function CheckLaunchNoFiles() As Boolean

        If Me.FilesList_Lview.Items.Count = 0 Then
            Call ShowTemporaryMessage("No archives selected, Exiting...")
            Return True
        Else
            Return False
        End If

    End Function

    Private Sub Process_Btn_Click(sender As Object, e As EventArgs) Handles Process_Btn.Click

        Dim start_time As DateTime
        Dim end_time As DateTime
        Dim elapsed_time As TimeSpan

        ' Launching wihout selected archives is a no-no...
        If CheckLaunchNoFiles() Then Exit Sub

        ' DEBUG
        start_time = Now

        Me.FilesList_Lview.UseWaitCursor = True

        ' uncover the progress bar
        Me.FilesList_Lview.Height = Me.FilesList_Lview.Height - 12
        Me.ListViewBackgound_Lbl.Height = Me.ListViewBackgound_Lbl.Height - 12

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
            If MsgBox("Processing archive(s) " & archivesProcessResult & " failed!" & vbCr & vbCr & _
                   "Would you continue zipping succesfully processed files?", vbYesNo + MsgBoxStyle.Question) = MsgBoxResult.No Then

                Me.Message_Lbl.Text = "Aborted"
                Exit Sub

            End If
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

        ' Also clean temp folders if required
        Call CleanTempFolders()


        If Me.ProgressBar1.Value <> 100 Then Me.ProgressBar1.Value = 100

        Me.FilesList_Lview.UseWaitCursor = False

        Me.ProgressBar1.Value = 0

        ' cover the progress bar again... not needed
        Me.FilesList_Lview.Height = Me.FilesList_Lview.Height + 12
        Me.ListViewBackgound_Lbl.Height = Me.ListViewBackgound_Lbl.Height + 12

        Me.Message_Lbl.Text = "All done"


        'Dim ddocx As Novacode.DocX = Novacode.DocX.Load()
        'ddocx.c()
        'OpenXmlPowerTools.DocumentAssembler.()

        end_time = Now
        elapsed_time = end_time.Subtract(start_time)

        Me.Message_Lbl.Text = Me.Message_Lbl.Text & " (" & elapsed_time.TotalSeconds.ToString("0.000") & "s)"

    End Sub

    Private Sub CleanTempFolders()

        ' and also clean folders of temp files if user so desires
        If My.Settings.App_AutoClean_Folders Then
            Me.Message_Lbl.Text = "Auto-Cleaning app folders..."
            Me.Refresh()
            Call ClearInOutFolders()
        End If

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

        Me.zipIcons.Dispose()

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

            If Directory.Exists(Path.Combine(Me.WorkingFolderBrowserDialog.SelectedPath, baseFolder)) Then
                Me.CtWorkingFolder_TxtBox.Text = Path.Combine(Me.WorkingFolderBrowserDialog.SelectedPath, baseFolder)
            Else
                Me.CtWorkingFolder_TxtBox.Text = Me.WorkingFolderBrowserDialog.SelectedPath
            End If

            My.Settings.Default_App_WorkingFolder = Me.WorkingFolderBrowserDialog.SelectedPath

        End If

    End Sub


    Private Sub ToggleOptions_Lbl_Click(sender As Object, e As EventArgs) Handles ToggleOptions_Lbl.Click

        ' We're collapsed
        If Me.Options_Opened = False Then

            ' Lock controls to enable form widening
            Call Set_Anchors_forFormResize(True)

            Me.Width = Me.Width + My.Settings.App_Options_GroupBox_Width

            ' Raise flag that form is widened
            Me.Options_Opened = True

        Else    ' we're not collapsed

            Me.Width = Me.Width - My.Settings.App_Options_GroupBox_Width

            ' Raise flag that form is widened
            Me.Options_Opened = False

            ' Unlock controls right back, form resizing moves them too (options part unaccessible)
            Call Set_Anchors_forFormResize(False)

        End If


    End Sub

    Private Sub Set_Anchors_forFormResize(ByVal AnchorsSetting As Boolean)

        If AnchorsSetting Then      ' De-Anchor controls, prepare for form resize

            ' hold listview, its back label and the right side controls in place
            Me.ListViewBackgound_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Bottom
            Me.FilesList_Lview.Anchor = AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Bottom
            '
            Me.ClearList_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Top
            Me.Pin_UserformSize_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Top

            Me.ToggleOptions_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Bottom
            '
            Me.ProgressBar1.Anchor = AnchorStyles.Left + AnchorStyles.Bottom
            Me.Message_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Bottom

            ' and hold options group box in place
            Me.GroupBox1.Anchor = AnchorStyles.Left + AnchorStyles.Top

        Else    ' Re-anchor for default behaviour, meaning form resize causes listview resize and right side controls move

            ' hold listview, its back label and the right side controls in place
            Me.ListViewBackgound_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Bottom + AnchorStyles.Right
            Me.FilesList_Lview.Anchor = AnchorStyles.Left + AnchorStyles.Top + AnchorStyles.Bottom + AnchorStyles.Right
            '
            Me.ClearList_Lbl.Anchor = AnchorStyles.Right + AnchorStyles.Top
            Me.Pin_UserformSize_Lbl.Anchor = AnchorStyles.Right + AnchorStyles.Top

            Me.ToggleOptions_Lbl.Anchor = AnchorStyles.Right + AnchorStyles.Bottom
            '
            Me.ProgressBar1.Anchor = AnchorStyles.Left + AnchorStyles.Bottom + AnchorStyles.Right
            Me.Message_Lbl.Anchor = AnchorStyles.Left + AnchorStyles.Bottom + AnchorStyles.Right

            ' and hold options group box in place
            Me.GroupBox1.Anchor = AnchorStyles.Right + AnchorStyles.Top

        End If

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles Me.Load


        Dim ad As System.Deployment.Application.ApplicationDeployment = System.Deployment.Application.ApplicationDeployment.CurrentDeployment
        Dim sVersion As String

        RTFConverter = New RichTextBox

        'Dim appVersionInfo As FileVersionInfo
        'appVersionInfo = FileVersionInfo.GetVersionInfo(Application.ExecutablePath)
        sVersion = ad.CurrentVersion.ToString

        Me.Text = Me.Text & " (" & sVersion & ")"

        'Me.Width = 510  ' Start form collapsed
        Call SetForm_Dimensions(My.Settings.FormDefaultWidth, My.Settings.FormDefaultHeight)

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

        Call RestoreFormOptions()

        Call SetTooltips

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

        Me.Message_Lbl.Text = "Ready"

    End Sub

    Private Sub SetTooltips()

        With ToolTip1

            .SetToolTip(ChooseZip_Btn, "Choose one or more zip archives to process (use Ctrl/ Shift). No folders.")
            .SetToolTip(Process_Btn, "Launch processing of zip archives in the list")
            .SetToolTip(FilesList_Lview, "Drag & drop one/multiple folders and/or files here to add all zips within")
            .SetToolTip(ClearList_Lbl, "Clear list and remove temporary folders")
            .SetToolTip(Pin_UserformSize_Lbl, "Make current form size default")
            .SetToolTip(ToggleOptions_Lbl, "Show/hide options area")

            .SetToolTip(CtWorkingFolder_TxtBox, "This folder is where the app creates input/ output and the resulting archives will be (in the Output subfolder)")
            .SetToolTip(Default_BrowseTo_TxtBox, "Set this option to always start browsing from this folder when using <Choose Zip> button")
            .SetToolTip(AutoClean_Folders_Cbox, "Set this option to always clear temporary folder on program exit")
            .SetToolTip(Default_BrowseToFolder_Cbox, "Activate this option to be able to set a default browse to folder")
            .SetToolTip(OpenOutputFolder_Cbox, "Set this option to always open the folder contating the result archives on task completion")

            .SetToolTip(ChooseBrowseToFolder_Lbl, "Click to change default browse to folder")
            .SetToolTip(ChooseWorkingFolder_Lbl, "Click to change working folder ")

        End With

    End Sub

    Private Sub RestoreFormOptions()

        ' and restore user saved settings for our check-boxes
        If Directory.Exists(Path.Combine(My.Settings.Default_App_WorkingFolder, baseFolder)) Then
            Me.CtWorkingFolder_TxtBox.Text = Path.Combine(My.Settings.Default_App_WorkingFolder, baseFolder)
        Else
            Me.CtWorkingFolder_TxtBox.Text = My.Settings.Default_App_WorkingFolder
        End If

        Me.OpenOutputFolder_Cbox.Checked = My.Settings.App_Open_OutputFolder
        Me.AutoClean_Folders_Cbox.Checked = My.Settings.App_AutoClean_Folders
        Me.Default_BrowseToFolder_Cbox.Checked = IIf(My.Settings.Default_BrowseTo_Folder <> "C:\", True, False)

        If Me.Default_BrowseToFolder_Cbox.Checked Then
            Me.Default_BrowseTo_TxtBox.Text = My.Settings.Default_BrowseTo_Folder
            Me.Default_BrowseTo_TxtBox.Enabled = True
        End If


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

            'wdHandle = wordApp.ActiveWindow.Hwnd
            ' MsgBox("Word handle captured as " & wdHandle)

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


    Private Sub Update_Filelist_Icons()

        Dim iconIndex As Integer

        Select Case Me.FilesList_Lview.View
            Case System.Windows.Forms.View.List
                iconIndex = 0
            Case System.Windows.Forms.View.SmallIcon
                iconIndex = 1
            Case System.Windows.Forms.View.LargeIcon
                iconIndex = 2
        End Select

        Dim litem As ListViewItem

        For Each litem In Me.FilesList_Lview.Items

            litem.ImageIndex = iconIndex

        Next

        Me.Refresh()

    End Sub



    Private Sub ClearList_Lbl_MouseDown(sender As Object, e As MouseEventArgs) Handles ClearList_Lbl.MouseDown
        Me.ClearList_Lbl.ForeColor = System.Drawing.Color.DarkViolet
    End Sub

    Private Sub ClearList_Lbl_MouseUp(sender As Object, e As MouseEventArgs) Handles ClearList_Lbl.MouseUp
        Me.ClearList_Lbl.ForeColor = System.Drawing.Color.Red
    End Sub


    Private Sub ToggleOptions_Lbl_MouseDown(sender As Object, e As MouseEventArgs) Handles ToggleOptions_Lbl.MouseDown
        Me.ToggleOptions_Lbl.ForeColor = System.Drawing.Color.DarkViolet
    End Sub

    Private Sub ToggleOptions_Lbl_MouseUp(sender As Object, e As MouseEventArgs) Handles ToggleOptions_Lbl.MouseUp
        Me.ToggleOptions_Lbl.ForeColor = System.Drawing.Color.RoyalBlue
    End Sub


    Private Sub Default_BrowseToFolder_Cbox_Click(sender As Object, e As EventArgs) Handles Default_BrowseToFolder_Cbox.Click

        If Me.Default_BrowseToFolder_Cbox.Checked Then

            If Me.WorkingFolderBrowserDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Me.Default_BrowseTo_TxtBox.Enabled = True
                Me.Default_BrowseTo_TxtBox.Text = Me.WorkingFolderBrowserDialog.SelectedPath
                My.Settings.Default_BrowseTo_Folder = Me.WorkingFolderBrowserDialog.SelectedPath
            Else    ' revert to default
                Me.Default_BrowseTo_TxtBox.Text = ""
                My.Settings.Default_BrowseTo_Folder = "C:\"
                Me.Default_BrowseTo_TxtBox.Enabled = False
            End If
        Else
            Me.Default_BrowseTo_TxtBox.Text = ""
            My.Settings.Default_BrowseTo_Folder = "C:\"
            Me.Default_BrowseTo_TxtBox.Enabled = False
        End If

    End Sub

    Private Sub ChooseBrowseToFolder_Lbl_Click(sender As Object, e As EventArgs) Handles ChooseBrowseToFolder_Lbl.Click

        If Me.Default_BrowseTo_TxtBox.Enabled Then
            If Me.WorkingFolderBrowserDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                Me.Default_BrowseTo_TxtBox.Text = Me.WorkingFolderBrowserDialog.SelectedPath
                My.Settings.Default_BrowseTo_Folder = Me.WorkingFolderBrowserDialog.SelectedPath
            End If
        End If
    End Sub

    Private Sub SetForm_Dimensions(UserformWidth As Integer, UserformHeight As Integer)

        ' Lock controls for resize form while not showing options
        Call Set_Anchors_forFormResize(False)

        Me.Width = UserformWidth
        Me.Height = UserformHeight

        Me.Options_Opened = False

        ' and relock
        'Call Set_Anchors_forFormResize(False)

    End Sub

    Private Sub Pin_UserformSize_Lbl_Click(sender As Object, e As EventArgs) Handles Pin_UserformSize_Lbl.Click

        ' DO NOT save new size settings if userform widened
        If Not Me.Options_Opened Then

            If Me.Width <> My.Settings.FormDefaultWidth Or Me.Height <> My.Settings.FormDefaultHeight Then

                My.Settings.FormDefaultWidth = Me.Width
                My.Settings.FormDefaultHeight = Me.Height

                Call ShowTemporaryMessage("New form default size set ")

            End If

        End If

    End Sub

    Private Sub Pin_UserformSize_Lbl_MouseDown(sender As Object, e As MouseEventArgs) Handles Pin_UserformSize_Lbl.MouseDown
        Me.Pin_UserformSize_Lbl.ForeColor = System.Drawing.Color.DarkViolet
    End Sub

    Private Sub Pin_UserformSize_Lbl_MouseUp(sender As Object, e As MouseEventArgs) Handles Pin_UserformSize_Lbl.MouseUp
        Me.Pin_UserformSize_Lbl.ForeColor = System.Drawing.Color.Maroon
    End Sub

    Private Sub ShowTemporaryMessage(Message As String)

        Me.Message_Lbl.Text = Message
        Me.Refresh()
        Threading.Thread.Sleep(600)
        Me.Message_Lbl.Text = "Ready"
        Me.Refresh()

    End Sub

End Class