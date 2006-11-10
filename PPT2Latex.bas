Attribute VB_Name = "ExportToLaTex"
' This macro will export the notes text from each slide in your presentation
' to the file you specify.
' Copyright (C) 2005-2006 Marco Aurélio Graciotto Silva <magsilva@gmail.com>

Dim outputFile
Dim outputPath
    
Const prefixDir = "\latex\"
Const latexSuffix = ".tex"
Const versionString = "0.2.0"

Private Function ConvertName(name As String) As String
    ConvertName = Replace(name, " ", "")
End Function

Sub Initialization()
    Dim fileSystem As Object
    Dim outputFilename As String
    
    Set fileSystem = CreateObject("Scripting.FileSystemObject")
    
    ' Sets the output directory (subdirectory "prefixDir" in the current presentation).
    outputPath = fileSystem.GetAbsolutePathName(ActivePresentation.Path) & prefixDir
    If fileSystem.FolderExists(outputPath) = False Then
        fileSystem.CreateFolder (outputPath)
    End If
    
    ' Sets the name of the main tex file.
    outputFilename = Left(ActivePresentation.name, Len(ActivePresentation.name) - 4) & latexSuffix
    Set outputFile = fileSystem.CreateTextFile(outputPath & outputFilename, True, False)
End Sub

Sub Finalization()
    outputFile.WriteLine ("\end{document} ")
    outputFile.Close
End Sub

Sub About()
    outputFile.WriteLine ("% Presentation compiled from a PowerPoint to LaTeX by ExportToLatex (v." & versionString & ")")
    outputFile.WriteLine ("% ExportToLatex is a weekend hack product by Marco Aurélio Graciotto Silva <magsilva@icmc.usp.br>")
End Sub
      
Sub AboutPresentation()
    outputFile.WriteBlankLines (1)
    With ActivePresentation
        outputFile.WriteLine ("% Source powerpoint file: " & .FullName)
        outputFile.WriteLine ("% Title: " & .BuiltInDocumentProperties("Title"))
        outputFile.WriteLine ("% Subject: " & .BuiltInDocumentProperties("Subject"))
        outputFile.WriteLine ("% Author: " & .BuiltInDocumentProperties("Author"))
        outputFile.WriteLine ("% Creation date: " & .BuiltInDocumentProperties("Creation date"))
        outputFile.WriteLine ("% Last modification date: " & .BuiltInDocumentProperties("Last save time"))
    End With
End Sub

Sub Header()
    With outputFile
        .WriteLine ("\documentclass[a4paper,oneside,12pt]{article}")
        .WriteLine ("\usepackage{tabularx}")
        .WriteLine ("\usepackage{hyperref}")
        .WriteLine ("\hypersetup{pdfborder={0 0 0}}")
        .WriteBlankLines (1)
        .WriteLine ("\title{" & ActivePresentation.BuiltInDocumentProperties("Title") & "}")
        .WriteLine ("\author{" & ActivePresentation.BuiltInDocumentProperties("Author") & "}")
        .WriteBlankLines (1)
        .WriteLine ("\begin{document}")
    End With
End Sub

Sub ExportFigure(figureShape As Shape)
    ' Yeah, PowerPoint can give the same name to different shapes...
    Static figureCounter As Integer
    figureName = "figure" & figureCounter
    figureCounter = figureCounter + 1
    
    With figureShape
        outputFile.WriteLine ("\begin{figure}[hbt]")
        outputFile.WriteLine (vbTab & "\centering")
        outputFile.WriteLine (vbTab & "\includegraphics[scale=1,height=" & Round(.Height) & ",width=" & Round(.Width) & "]{" & figureName & "}")
        If Len(.AlternativeText) <> 0 Then
            outputFile.WriteLine (vbTab & "\caption{" & .AlternativeText & "}")
        End If
        outputFile.WriteLine (vbTab & "\label{" & figureName & "}")
        outputFile.WriteLine ("\end{figure}")
        .Export outputPath & figureName & ".png", ppShapeFormatPNG
    End With
End Sub

Sub ExportTable(tableShape As Shape)
    Dim i As Integer
    Dim j As Integer
                            
    ' Yeah, PowerPoint can give the same name to different shapes...
    Static tableCounter As Integer
    tableName = "table" & tableCounter
    tableCounter = tableCounter + 1
    
    outputFile.WriteLine ("\begin{table}[hbt]")
    outputFile.WriteLine (vbTab & "\centering")
    If Len(tableShape.AlternativeText) <> 0 Then
        outputFile.WriteLine (vbTab & "\caption{" & tableShape.AlternativeText & "}")
    End If
    outputFile.Write (vbTab & "\begin{tabular}{")
    For i = 1 To tableShape.Table.Columns.Count
        outputFile.Write ("|c")
    Next i
    outputFile.WriteLine ("|}")
                            
    With tableShape.Table
        For i = 1 To .Rows.Count
            For j = 1 To .Columns.Count
                If j = 1 Then
                    outputFile.Write (vbTab & vbTab & "\hline ")
                Else
                    outputFile.Write (" & ")
                End If
                outputFile.Write (.Cell(i, j).Shape.TextFrame.TextRange.Text)
            Next j
            outputFile.WriteLine ("\\")
        Next i
    End With
    If i > 1 Then
        outputFile.Write (vbTab & vbTab & "\hline")
    End If
    
    outputFile.WriteLine (vbTab & "\end{tabular}")
    outputFile.WriteLine (vbTab & "\label{" & tableName & "}")
    outputFile.WriteLine ("\end{table}")
End Sub

Sub ExportMovie(movieShape As Shape)
    ' Yeah, PowerPoint can give the same name to different shapes...
    Static movieCounter As Integer
    movieName = "movie" & movieCounter
    movieCounter = movieCounter + 1
    
    Dim movieFilename As String
    
    
    With movieShape
        outputFile.WriteLine ("\begin{figure}[hbt]")
        outputFile.WriteLine (vbTab & "\centering")
        outputFile.WriteLine (vbTab & "\href{run:" & movieFilename & "}{Click here to show movie}")
        outputFile.WriteLine (vbTab & "\fbox{")
        outputFile.WriteLine (vbTab & vbTab & "\href{" & movieFilename & "}{\includegraphics[scale=1,height=" & Round(.Height) & ",width=" & Round(.Width) & "]{" & movieName & "}")
        outputFile.WriteLine (vbTab & "}")
        If Len(.AlternativeText) <> 0 Then
            outputFile.WriteLine (vbTab & "\caption{" & .AlternativeText & "}")
        End If
        outputFile.WriteLine (vbTab & "\label{" & movieName & "}")
        outputFile.WriteLine ("\end{figure}")
        .Export outputPath & movieName & ".png", ppShapeFormatPNG
    End With

        'outputFile.WriteLine ("\pdfannot{")
        'outputFile.WriteLine (vbTab & "/Subtype /Movie")
        'outputFile.WriteLine (vbTab & "  /T (Movie Title) ")
        'outputFile.WriteLine (vbTab & "  /Movie <> ")
        'outputFile.WriteLine (vbTab & "  /A << /ShowControls true >>")
        'outputFile.WriteLine (vbTab & "  /ANN pdfmark}")
    
End Sub

Sub ExportText(textShape As Shape)
        If textShape.HasTextFrame And textShape.TextFrame.HasText Then
        outputFile.WriteLine (textShape.TextFrame.TextRange.Text)
        End If
End Sub

Sub ExportSubtitle(textShape As Shape)
        If textShape.HasTextFrame And textShape.TextFrame.HasText Then
                outputFile.WriteLine ("\subsection{" & textShape.TextFrame.TextRange.Text & "}")
        End If
End Sub

Sub ExportTitle(textShape As Shape)
        If textShape.HasTextFrame And textShape.TextFrame.HasText Then
                outputFile.WriteLine ("\section{" & textShape.TextFrame.TextRange.Text & "}")
        End If
End Sub

Sub ExportCenterTitle(textShape As Shape)
        If textShape.HasTextFrame And textShape.TextFrame.HasText Then
                outputFile.WriteLine ("\section{" & textShape.TextFrame.TextRange.Text & "}")
        End If
End Sub


Sub ExportShape(unknownShape As Shape)
    With unknownShape
        If .Type = msoMedia Then
            Call ExportMovie(unknownShape)
        End If
    
        If .Type = msoTable Then
            If .HasTable Then
                Call ExportTable(unknownShape)
            End If
        End If

        If .Type = msoPicture Then
            Call ExportFigure(unknownShape)
        End If

                                If .Type = msoTextBox Then
                                                Call ExportText(unknownShape)
                                End If

        If .Type = msoPlaceholder Then
            Select Case .PlaceholderFormat.Type
                Case Is = ppPlaceholderChart, ppPlaceholderDate, ppPlaceholderHeader, ppPlaceholderMediaClip, ppPlaceholderOrgChart, ppPlaceholderSlideNumber, ppPlaceholderVerticalBody, ppPlaceholderVerticalTitle
                                                                                Call ExportText(unknownShape)
                    
                Case Is = ppPlaceholderBitmap, ppPlaceholderMixed, ppPlaceholderObject
                    Call ExportFigure(unknownShape)
    
                Case Is = ppPlaceholderFooter
                                                                                Call ExportText(unknownShape)
                                                     
                Case Is = ppPlaceholderTable
                                                                                Call ExportTable(unknownShape)
                           
                Case Is = ppPlaceholderBody
                                                                                Call ExportText(unknownShape)
                            
                Case Is = ppPlaceholderSubtitle
                                                                                Call ExportSubtitle(unknownShape)
            
                Case Is = ppPlaceholderCenterTitle
                                                                                Call ExportCenterTitle(unknownShape)
                                                                                
                Case Is = ppPlaceholderTitle
                                                                                Call ExportTitle(unknownShape)
            End Select
        End If
    End With
End Sub


Sub ExportSlide(currentSlide As Slide)
    Dim currentShape As Shape
    
    outputFile.WriteBlankLines (1)
    outputFile.WriteLine ("% ----------------------------------------------------------")
    outputFile.WriteLine ("% Slide " & currentSlide.SlideNumber & " (" & currentSlide.name & ")")
               
    For Each currentShape In currentSlide.Shapes
        outputFile.WriteBlankLines (1)
        Call ExportShape(currentShape)
    Next currentShape
    
    For Each currentShape In currentSlide.NotesPage.Shapes
        outputFile.WriteBlankLines (1)
        Call ExportShape(currentShape)
    Next currentShape
End Sub


Sub ExportToLatex()
    Dim currentSlide As Slide
    
    Call Initialization
    Call About
    Call AboutPresentation
    Call Header

    outputFile.WriteBlankLines (2)
    For Each currentSlide In ActivePresentation.Slides
        Call ExportSlide(currentSlide)
    Next currentSlide
      
    Call Finalization
End Sub
