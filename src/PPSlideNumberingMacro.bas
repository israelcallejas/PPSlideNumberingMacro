Attribute VB_Name = "PPSlideNumberingMacro"
'*******************************************************************************
' Program Name: PPSlideNumberingMacro
' Author: Israel Callejas
' Date: October 8, 2023
'
' Author's GitHub Profile: https://github.com/israelcallejas
'
' Description:
' This macro is designed to automatically assign page numbers to each slide in
' a PowerPoint presentation, starting from a specified initial page number. It
' utilizes an event handler to update slide numbers when a new slide is added,
' ensuring that the numbering is always accurate. The user can customize the
' initial page number by modifying the "initial_page_long" variable.
'
' This code is released under an open-source license , and you are free to use,
' modify, and distribute it. Please provide proper attribution to the author.
'
' Note: Make sure to adjust the "initial_page_long" variable to set the desired
'       initial page number at 1 (by default).
'
' GNU General Public License Version 3 (GPLv3)
'*******************************************************************************
Option Explicit
    Public AlternativeText As String
    Public is_first_start As Boolean
    
'// Doc: https://learn.microsoft.com/es-es/office/vba/powerpoint/how-to/use-events-with-the-application-object
Dim X As New WithEvents_class
Sub initialize_app()
    Set X.App = Application
End Sub
'///////
Sub start_counter_page()
    
    
    Dim initial_page_long As Long
    Dim start_number As Long
    Dim text_box_width As Single
    Dim text_box_height As Single
    Dim clg As Long
    Dim cng As Long
    Dim page_slide As Slide
    Dim text_box_shape As shape '// No necessary in this version as global
    
    initial_page_long = 3 '// Pick any page you fancy
    start_number = 1 '// The counter starts at 1 (it could also be defined as initial_page_long, etc.)

    text_box_width = 100
    text_box_height = 50
    
    remove_counter_page
    
    clg = 1
    For Each page_slide In ActivePresentation.Slides
    
        '// No duplicates when the start_counter_page or macro is activated by the user
        'If is_first_start = False Then
        '    Dim j As Long
        '    For j = page_slide.Shapes.Count To 1 Step -1
        '        If page_slide.Shapes(j).Type = msoTextBox And page_slide.Shapes(j).AlternativeText = AlternativeText Then
        '            page_slide.Shapes(j).Delete
        '        End If
        '    Next j
        'End If
        
        
        Dim slide_width As Single
        Dim slide_height As Single
        slide_width = ActivePresentation.PageSetup.SlideWidth
        slide_height = ActivePresentation.PageSetup.SlideHeight
        
        If clg >= initial_page_long Then
            Set text_box_shape = page_slide.Shapes.AddTextbox(msoTextOrientationHorizontal, Left:=slide_width - text_box_width, Top:=slide_height - text_box_height, Width:=text_box_width, Height:=text_box_height)
            text_box_shape.TextFrame.TextRange.Text = CStr(clg - initial_page_long + start_number)
            text_box_shape.AlternativeText = AlternativeText '// This could be improved, but it doesn't matter in this first version
            'AlternativeText = text_box_shape.AlternativeText
        End If
        clg = clg + 1
        
    Next page_slide
    
    On Error Resume Next '// Enable error handling;
    Dim temp As shape
    temp = text_box_shape '// Intenta acceder a la variable;
    If Err.Number <> 0 Then '// If there's an error, the variable is not declared (a value of initial_page_long greater than the number of slides is chosen);
        '// I do nothing... note that if I plan to add_new_slide, it would go into a loop when calling this Sub for enumeration
    Else
        text_box_shape.TextFrame.TextRange.Font.Name = "Arial" '// Font, no spaces problem;
    End If
    On Error GoTo 0
    
    'is_first_start = False
    
End Sub
Sub remove_counter_page()
    '// Remove the counter page
    Dim page_slide As Slide
    Dim j As Long
    For Each page_slide In ActivePresentation.Slides
        For j = page_slide.Shapes.Count To 1 Step -1 '// No problem;
            If page_slide.Shapes(j).Type = msoTextBox And page_slide.Shapes(j).AlternativeText = AlternativeText Then
                page_slide.Shapes(j).Delete
                'Exit For
            End If
        Next j
    Next page_slide
    'is_first_start = True
End Sub
Sub add_new_slide()
    '// add a new slide to the end of the presentation (no matter the location)
    ActivePresentation.Slides.Add Index:=ActivePresentation.Slides.Count + 1, Layout:=ppLayoutText
End Sub
Private Sub Workbook_Open()
    initialize_app
    is_first_start = True
    'AlternativeText = "" '// Maybe it is declared as such by default
    AlternativeText = "unique_id"
    remove_counter_page
    start_counter_page
End Sub
Sub Auto_Open()
    Call Workbook_Open
End Sub

'Sub remove_counter_page_v2()
'    Dim page_slide As Slide
'    Dim shape As shape
'
'    For Each page_slide In ActivePresentation.Slides
'        For Each shape In page_slide.Shapes
'            If shape.Type = msoTextBox And shape.AlternativeText = AlternativeText Then '// Set the AlternativeText to "unique_id" in all cases. (!);
'                shape.Delete
'                Exit For
'            End If
'        Next shape
'    Next page_slide
'End Sub
