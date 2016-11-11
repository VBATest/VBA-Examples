Attribute VB_Name = "Readxmldoc"
Sub ReadXML()
    Call fnReadXMLByTags
End Sub

Function fnReadXMLByTags()
    Dim mainWorkBook As Workbook
    Set mainWorkBook = ActiveWorkbook
    mainWorkBook.Sheets("Sheet1").Range("A:A").Clear
    Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    Filt = "xml Files (*.xml),*.xml"
    XMLFileName = Application.GetOpenFilename(FileFilter:=Filt, Title:=Title)
    oXMLFile.Load (XMLFileName)
    Set TitleNodes = oXMLFile.SelectNodes("/catalog/book/title/text()")
    Set PriceNodes = oXMLFile.SelectNodes("/catalog/book/price/text()")
    mainWorkBook.Sheets("Sheet1").Range("A1,B1,C1").Interior.ColorIndex = 40
    mainWorkBook.Sheets("Sheet1").Range("A1,B1,C1").Borders.Value = 1
    mainWorkBook.Sheets("Sheet1").Range("A" & 1).Value = "Book ID"
    mainWorkBook.Sheets("Sheet1").Range("B" & 1).Value = "Book Titles"
    mainWorkBook.Sheets("Sheet1").Range("C" & 1).Value = "Price"
    mainWorkBook.Sheets("Sheet1").Range("D1").Value = "Total books: " & TitleNodes.Length
        For I = 0 To (TitleNodes.Length - 1)
            Title = TitleNodes(I).NodeValue
            Price = PriceNodes(I).NodeValue
        mainWorkBook.Sheets("Sheet1").Range("B" & I + 2).Borders.Value = 1
        mainWorkBook.Sheets("Sheet1").Range("C" & I + 2).Borders.Value = 1
        mainWorkBook.Sheets("Sheet1").Range("B" & I + 2).Value = Title
        mainWorkBook.Sheets("Sheet1").Range("C" & I + 2).Value = Price
    Next
    'Reading the Attributes
    Set Nodes_Attribute = oXMLFile.SelectNodes("/catalog/book")
    For I = 0 To (Nodes_Attribute.Length - 1)
        Attributes = Nodes_Attribute(I).getAttribute("id")
        mainWorkBook.Sheets("Sheet1").Range("A" & I + 2).Borders.Value = 1
        mainWorkBook.Sheets("Sheet1").Range("A" & I + 2).Value = Attributes
    Next
End Function

