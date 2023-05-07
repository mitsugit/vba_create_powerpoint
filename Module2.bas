Sub ImportPNGs(ByVal folderPath As String)
    ' Import PNG files from a specified folder into PowerPoint slides
    ' and display the file name in the top-left corner of each slide.

    ' Create a new PowerPoint presentation
    Dim pptApp As PowerPoint.Application
    Set pptApp = New PowerPoint.Application
    Dim pptPres As PowerPoint.Presentation
    Set pptPres = pptApp.Presentations.Add(msoTrue)

    ' Import each PNG file as a new slide
    Dim pngFileName As String
    Dim slide As PowerPoint.slide
    Dim pic As PowerPoint.Shape
    Dim textBox As PowerPoint.Shape
    
    ' Add a backslash to the end of the folder path if it does not exist
    folderPath = NormalizeFolderPath(folderPath)
    
    pngFileName = Dir(folderPath & "*.png")
    Do While Len(pngFileName) > 0
        ' Add a new slide and import the PNG file as a picture
        Set slide = pptPres.Slides.Add(pptPres.Slides.Count + 1, ppLayoutBlank)
        Set pic = slide.Shapes.AddPicture(folderPath & pngFileName, msoFalse, msoTrue, 0, 0)

        ' Add a text box to display the file name
        Set textBox = AddFileNameTextbox(slide, pic.left + 10, pic.top + 10, slide.Master.width, 20, pngFileName)
        
        ' Get the next PNG file name
        pngFileName = Dir
    Loop

    ' Clean up
    Set pic = Nothing
    Set textBox = Nothing
    Set slide = Nothing
    Set pptPres = Nothing
    pptApp.Visible = True
    Set pptApp = Nothing
End Sub

Function NormalizeFolderPath(ByVal folderPath As String) As String
    ' Add a backslash to the end of the folder path if it does not exist
    If Right(folderPath, 1) <> "\" Then
        folderPath = folderPath & "\"
    End If
    NormalizeFolderPath = folderPath
End Function

Function AddFileNameTextbox(ByVal slide As PowerPoint.slide, ByVal left As Single, ByVal top As Single, ByVal width As Single, ByVal height As Single, ByVal fileName As String) As PowerPoint.Shape
    ' Add a text box to display the file name in the specified slide
    Dim textBox As PowerPoint.Shape
    Set textBox = slide.Shapes.AddTextbox(msoTextOrientationHorizontal, left, top, 300, 50)
    With textBox.TextFrame.TextRange
        .Text = fileName
        .Font.Name = "Arial"
        .Font.Size = 12
    End With
    With textBox.Fill
        .Transparency = 0.1
        .ForeColor.RGB = RGB(192, 192, 192)
    End With
    Set AddFileNameTextbox = textBox
End Function

Sub TestImportPNGs()
    ' Test the ImportPNGs sub with a folder path argument
    ImportPNGs "C:\new_dev\java_multithread"
End Sub
