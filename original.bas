Option Explicit

Public Const wdStyleHeading1 = -2
Public Const wdStyleTitle = -63
Public Const wdPageBreak = 7
Public Const wdStyleSubtitle = -75
Public Const wdHeaderFooterPrimary = 1
Public Const wdAdjustNone = 0
Public Const wdAdjustProportional = 1
Public Const wdOrientLandscape = 1
Public Const wdFormatXMLDocument = 12

Public Const colDataStart = 3
Public Const colDataEnd = 62

Public Const rowDataStart = 1
Public Const rowDataEnd = 7
Public Const rowHeader = 2

Public Const themesDirectory = "Macintosh HD:Applications:Microsoft Office 2011:Office:Media:Office Themes:"

Private Function createDocument() As Object
  Dim word As Object
  'create the document
  On Error Resume Next
  Set word = GetObject(, "Word.Application") 'gives error 429 if Word is not open
  If Err = 429 Then
      Err.Clear
      Set word = CreateObject("Word.Application") 'creates a Word application
      If Err = 429 Then
        Debug.Print "Failed to create Word object"
      End If
  End If
  word.Visible = True

  ' create a new document
  Set createDocument = word.documents.Add
  'apply theme
   createDocument.ApplyTheme themesDirectory & "Angles.thmx"
End Function

Private Sub saveDocument(doc As Object, filename As String)
   doc.SaveAs filename:=filename & ".docx", FileFormat:= _
        wdFormatXMLDocument, LockComments:=False, Password:="", AddToRecentFiles _
        :=False, WritePassword:="", ReadOnlyRecommended:=False, EmbedTrueTypeFonts _
        :=False, SaveNativePictureFormat:=False, SaveFormsData:=False, _
        SaveAsAOCELetter:=False, HTMLDisplayOnlyOutput:=False, MaintainCompat:= _
        False
End Sub

Public Sub produceSkillsProfile()

    Dim theTable As Object, doc As Object, sheet As Object

    'ask the user which roles to use - comma separated list of row names, e.g. "1,2,65"
    Dim rowIndex As String, colIndex As Integer
    rowIndex = InputBox("Which rows should the skills profiles be generated from?")

    'tokenize
    Dim rowsToProcess As Variant, rowToProcess As Variant
    rowsToProcess = Split(rowIndex, ",")

    'for each row
    For Each rowToProcess In rowsToProcess

        Set doc = createDocument()

        'Set header
        With doc.sections(doc.sections.Count)
         .Headers(wdHeaderFooterPrimary).Range.text = "Skills Profile"
         .Footers(wdHeaderFooterPrimary).Range.text = copyright()
        End With

        'Add skills profile header
        doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBefore ("Wynyard Group Skills Profile")
        doc.Paragraphs(doc.Paragraphs.Count).Range.Style = wdStyleTitle


        ' add skills profile table
        doc.Paragraphs.Add
        Set theTable = doc.Tables.Add(Range:=doc.Paragraphs(doc.Paragraphs.Count).Range, NumRows:=1, NumColumns:=2)

        ' process the skills profile table
        Dim newRow As Object
        Set sheet = ActiveWorkbook.Sheets("Skills Matrix") '
        Dim counter As Integer
        For counter = 1 To colDataEnd

          Dim rangeName As String, title As String, grade As String, text As String
          text = sheet.Range(counter & rowToProcess).Value

          If Not text = "N/A" Then 'only add a row if there is a value

            'add a row
            If counter <> 1 Then Set newRow = theTable.Rows.Add

            theTable.cell(counter, 1) = sheet.Cells(1, counter).Value
            theTable.cell(counter, 2) = text
            theTable.cell(counter, 1).setwidth ColumnWidth:=100, RulerStyle:=wdAdjustProportional

            ' use the information for the header
             If counter = 1 Then title = text
             If counter = 2 Then grade = text

          End If
        Next counter

        ' set the header
        Dim headerText As String
        headerText = "Skills Profile" & " - " & title & " (" & grade & ")"
        doc.sections(doc.sections.Count).Headers(wdHeaderFooterPrimary).Range.text = headerText

    Call saveDocument(doc, headerText)
    doc.Close

    Next rowToProcess

End Sub

Public Sub produceProgressionGrid()

    Dim theTable As Object, doc As Object, sheet As Object

    'ask the user which roles to use - comma separated list of column names, e.g. "A,B"
    Dim columnIndex As String, rowIndex As Integer
    columnIndex = InputBox("Which two columns should the grid be generated from?")

    'tokenize
    Dim columnsToProcess As Variant, columnToProcess As Variant
    columnsToProcess = Split(columnIndex, ",")
    Dim fromColumn As Variant, toColumn As Variant
    fromColumn = columnsToProcess(0)
    toColumn = columnsToProcess(1)

    Set doc = createDocument()
    doc.PageSetup.Orientation = wdOrientLandscape

    'Set header
    With doc.sections(doc.sections.Count)
     .Headers(wdHeaderFooterPrimary).Range.text = "Progression Matrix"
     .Footers(wdHeaderFooterPrimary).Range.text = copyright()
    End With

    'Add skills profile header
    doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBefore ("Wynyard Group Progression Matrix")
    doc.Paragraphs(doc.Paragraphs.Count).Range.Style = wdStyleTitle


    ' add skills profile table
    doc.Paragraphs.Add
    Set theTable = doc.Tables.Add(Range:=doc.Paragraphs(doc.Paragraphs.Count).Range, NumRows:=1, NumColumns:=4)

    'add table header
    theTable.cell(1, 2) = "From"
    theTable.cell(1, 3) = "To"
    theTable.cell(1, 4) = "Evidence"
    theTable.cell(1, 1).setwidth ColumnWidth:=100, RulerStyle:=wdAdjustProportional

    ' process the table
    Dim newRow As Object
    Set sheet = ActiveWorkbook.Sheets("Inverse Skills Matrix") '
    Dim counter As Integer
    For counter = 1 To rowDataEnd

      Dim rangeName As String, title As String, grade As String, leftText As String, rightText As String
      leftText = sheet.Range(fromColumn & counter).Value
      rightText = sheet.Range(toColumn & counter).Value

      'extract the SFIA levels
      Dim leftlevel As Integer, rightlevel As Integer
      leftlevel = 0
      rightlevel = 0
      If Left(leftText, 1) = "(" Then leftlevel = Int(Mid(leftText, 2, 1))
      If Left(rightText, 1) = "(" Then rightlevel = Int(Mid(rightText, 2, 1))

      Debug.Print sheet.Cells(counter, 1).Value & " " & leftlevel & " " & rightlevel

      'only add a row if there is a value in the from and/or to columns, and the level is increasing
      If ((leftText <> "N/A" And rightText <> "N/A") Or (leftText = "N/A" And rightText <> "N/A")) _
        And leftText <> rightText _
        And ((leftlevel = 0 And rightlevel = 0) Or rightlevel > leftlevel) _
        Then

        'add a row
        theTable.Rows.Add

        theTable.cell(counter + 1, 1) = sheet.Cells(counter, 1).Value
        theTable.cell(counter + 1, 2) = leftText
        theTable.cell(counter + 1, 3) = rightText
        theTable.cell(counter + 1, 1).setwidth ColumnWidth:=100, RulerStyle:=wdAdjustProportional

    End If

    ' use the information for the header
     If counter = 1 Then
       If leftText <> rightText Then
         title = leftText & " to " & rightText
       Else
         title = leftText
       End If
     End If
     If counter = 2 Then grade = leftText & " to " & rightText


    Next counter

    ' set the header
    Dim headerText As String
    headerText = "Progression Matrix " & " - " & title & " (" & grade & ")"
    doc.sections(doc.sections.Count).Headers(wdHeaderFooterPrimary).Range.text = headerText

    theTable.cell(1, 2).Range.Font.Bold = True
    theTable.cell(1, 3).Range.Font.Bold = True
    theTable.cell(1, 4).Range.Font.Bold = True

    Call saveDocument(doc, headerText)
    doc.Close


End Sub

Public Sub selfAssessments()

    Dim theTable As Object, doc As Object, sheet As Object

    'ask the user which roles to use - comma separated list of column names, e.g. "A,B,AA"
    Dim columnIndex As String, rowIndex As Integer
    columnIndex = InputBox("Which columns should the self assessments be generated from (comma separated)?")

    'tokenize
    Dim columnsToProcess As Variant, columnToProcess As Variant
    columnsToProcess = Split(columnIndex, ",")

    'for each column
    For Each columnToProcess In columnsToProcess

        Set doc = createDocument()

        'Set header
        With doc.sections(doc.sections.Count)
         .Headers(wdHeaderFooterPrimary).Range.text = "Self Assessment"
         .Footers(wdHeaderFooterPrimary).Range.text = copyright()
        End With

        'Add skills profile header
        doc.Paragraphs(doc.Paragraphs.Count).Range.InsertBefore ("Wynyard Group Self Assessment")
        doc.Paragraphs(doc.Paragraphs.Count).Range.Style = wdStyleTitle


        ' add skills profile table
        doc.Paragraphs.Add
        Set theTable = doc.Tables.Add(Range:=doc.Paragraphs(doc.Paragraphs.Count).Range, NumRows:=1, NumColumns:=3)

        ' process the skills profile table
        Dim newRow As Object
        Set sheet = ActiveWorkbook.Sheets("Inverse Skills Matrix") '
        Dim counter As Integer
        For counter = 1 To rowDataEnd

          Dim rangeName As String, title As String, grade As String, text As String
          text = sheet.Range(columnToProcess & counter).Value

          If Not text = "N/A" Then 'only add a row if there is a value

            'add a row
            If counter <> 1 Then Set newRow = theTable.Rows.Add

            Debug.Print counter & " " & sheet.Cells(counter, 1).Value

            theTable.cell(counter, 1) = sheet.Cells(counter, 1).Value
            theTable.cell(counter, 2) = text
            theTable.cell(counter, 1).setwidth ColumnWidth:=100, RulerStyle:=wdAdjustProportional

            ' use the information for the header
             If counter = 1 Then title = text
             If counter = 2 Then grade = text

          End If
        Next counter

        'add evidence header
        theTable.cell(2, 3) = "Evidence"

        ' set the header
        Dim headerText As String
        headerText = "Self Assessment" & " - " & title & " (" & grade & ")"
        doc.sections(doc.sections.Count).Headers(wdHeaderFooterPrimary).Range.text = headerText

    Call saveDocument(doc, headerText)
    doc.Close

    Next columnToProcess

End Sub

Private Function copyright() As String
  copyright = "©" & CStr(Year(Now())) & " Wynyard (NZ) Ltd"
End Function
