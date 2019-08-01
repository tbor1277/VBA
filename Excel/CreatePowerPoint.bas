Sub CreatePowerPoint()

 'Add a reference to the Microsoft PowerPoint Library by:
    '1. Go to Tools in the VBA menu
    '2. Click on Reference
    '3. Scroll down to Microsoft PowerPoint X.0 Object Library, check the box, and press Okay
 
    'First we declare the variables we will be using
        Dim newPowerPoint As PowerPoint.Application
        Dim activeSlide As PowerPoint.Slide
        Dim cht As Excel.ChartObject
     
     'Look for existing instance
        On Error Resume Next
        Set newPowerPoint = GetObject(, "PowerPoint.Application")
        On Error GoTo 0
     
    'Let's create a new PowerPoint
        If newPowerPoint Is Nothing Then
            Set newPowerPoint = New PowerPoint.Application
        End If
    'Make a presentation in PowerPoint
        If newPowerPoint.Presentations.Count = 0 Then
            newPowerPoint.Presentations.Add
        End If
     
    'Show the PowerPoint
        newPowerPoint.Visible = True
    
    'Loop through each chart in the Excel worksheet and paste them into the PowerPoint
        For Each cht In ActiveSheet.ChartObjects
        
        'Add a new slide where we will paste the chart
            newPowerPoint.ActivePresentation.Slides.Add newPowerPoint.ActivePresentation.Slides.Count + 1, ppLayoutText
            newPowerPoint.ActiveWindow.View.GotoSlide newPowerPoint.ActivePresentation.Slides.Count
            Set activeSlide = newPowerPoint.ActivePresentation.Slides(newPowerPoint.ActivePresentation.Slides.Count)
                
        'Copy the chart and paste it into the PowerPoint as a Metafile Picture
            cht.Select
            ActiveChart.ChartArea.Copy
            activeSlide.Shapes.PasteSpecial(DataType:=ppPasteMetafilePicture).Select
    
        'Set the title of the slide the same as the title of the chart
            activeSlide.Shapes(1).TextFrame.TextRange.Text = cht.Chart.ChartTitle.Text
            
        'Adjust the positioning of the Chart on Powerpoint Slide
            newPowerPoint.ActiveWindow.Selection.ShapeRange.Left = 15
            newPowerPoint.ActiveWindow.Selection.ShapeRange.Top = 125
        
            activeSlide.Shapes(2).Width = 200
            activeSlide.Shapes(2).Left = 505
            
        Next
     
    AppActivate ("Microsoft PowerPoint")
    Set activeSlide = Nothing
    Set newPowerPoint = Nothing
     
End Sub
