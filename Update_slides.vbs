
Sub Update_slides()
    
'This subroutine updates all the charts and embedded objects linked to external sources   

Dim pptPresentation As Presentation
Dim PPTSlide As Slide
Dim PPTShape As Shape
Dim Data As ChartData
Dim Graph As Chart
Dim excel_file As Object
Dim xlWorkbook As Object
Dim oxl As Excel.Workbook
Dim xlapp As Excel.Application
Dim xlsheet As Excel.Worksheet

    
For Each PPTSlide In ActivePresentation.Slides

    For Each PPTShape In PPTSlide.Shapes
    
        If PPTShape.Type = msoEmbeddedOLEObject Then
             Set oxl = PPTShape.OLEFormat.Object
             Set xlapp = PPTShape.OLEFormat.Object.Application
             Set xlsheet = oxl.Worksheets(1)
            
            xlapp.ActiveWorkbook.RefreshAll
        Else
        
            If PPTShape.HasChart Or PPTShape.Type = msoChart Or _
                PPTShape.Type = msoLinkedOLEObject Then
            
                Set Graph = PPTShape.Chart
                Set Data = Graph.ChartData
                Data.Activate
                PPTShape.Chart.Refresh
                Set excel_file = Data.Workbook
                excel_file.Close
             Else
             
             On Error Resume Next
             
             End If
    
        End If
        
    Next

Next
    
'ActivePresentation.UpdateLinks

MsgBox ("Actualización terminada")

Set excel_file = Nothing
Set Data = Nothing
Set Graph = Nothing


End Sub
