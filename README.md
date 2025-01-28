# extract-pages-docx
A Visual Basic for Applications (VBA) script for extracting all the pages that start with a specific font style from a word document.

Sub SavePagesWithHeading1()
    Dim doc As Document
    Dim pageDoc As Document
    Dim rng As Range
    Dim headingRng As Range
    Dim heading1Text As String
    Dim savePath As String
    Dim currentPage As Integer
    Dim totalPages As Integer

    ' Initialize variables
    Set doc = ActiveDocument
    savePath = "C:\Users\Carlo\Desktop\Pjesmarice\Zaja pjesmarica\" ' Specific folder path

    ' Create output folder if it doesn't exist
    On Error Resume Next
    MkDir savePath
    On Error GoTo 0

    ' Total pages in the document
    totalPages = doc.ComputeStatistics(wdStatisticPages)

    ' Iterate through each page
    For currentPage = 1 To totalPages
        ' Define the range for the current page
        Set rng = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPage)
        rng.End = doc.GoTo(What:=wdGoToPage, Which:=wdGoToAbsolute, Count:=currentPage + 1).Start - 1

        ' Exclude the page break at the end of the range if present
        If Right(rng.Text, 1) = Chr(12) Then ' Chr(12) represents a page break
            rng.End = rng.End - 1
        End If

        ' Extract the first paragraph in the page range
        Set headingRng = rng.Paragraphs(1).Range

        ' Check if the first paragraph is styled as Heading 1
        If headingRng.Style = doc.Styles(wdStyleHeading1) Then
            heading1Text = headingRng.Text

            ' Remove trailing line breaks
            heading1Text = Trim(Replace(Replace(heading1Text, vbCr, ""), vbLf, ""))
        Else
            heading1Text = "Page_" & currentPage
        End If

        ' Create a new document for the page
        Set pageDoc = Documents.Add

        ' Copy the range to the new document, preserving formatting
        rng.Copy
        pageDoc.Content.PasteAndFormat (wdFormatOriginalFormatting)

        ' Remove trailing empty paragraphs
        Do While pageDoc.Content.Paragraphs.Last.Range.Text = vbCr & vbCr
            pageDoc.Content.Paragraphs.Last.Range.Delete
        Loop

        ' Save the new document with the heading as the filename
        pageDoc.SaveAs2 FileName:=savePath & heading1Text & ".docx", FileFormat:=wdFormatDocumentDefault

        ' Close the new document
        pageDoc.Close SaveChanges:=False
    Next currentPage

    ' Notify the user
    MsgBox "Pages saved successfully to " & savePath, vbInformation
End Sub
