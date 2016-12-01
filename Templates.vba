Private Sub CommandButton1_Click()

'    Dim prnummer As String 'prnummer = projectnummer'
'    prnummer = TextBox3.Value
'    Call UpdateBookmark("projectnummer", prnummer)
    
    Dim opdrgever As String 'opdrgever = opdrachtgever'
    opdrgever = TextBox4.Value
    Call UpdateBookmark("opdrachtgever", opdrgever)

    Dim pginas As String 'pginas = aantal pagina's'
    pginas = TextBox5.Value
    Call UpdateBookmark("paginas", pginas)

'    Dim fse As String 'fse = fase'
'    fse = ComboBox1.Value
'    Call UpdateBookmark("fase", fse)

    'Add ContentTypeProperties to Sharepoint Columns
    ActiveDocument.ContentTypeProperties("Document Type").Value = "Fasedocument"
    ActiveDocument.ContentTypeProperties("Project Number").Value = TextBox3.Value
    
    If ComboBox1.Value = "Opstartfase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Opstartfase"
    ElseIf ComboBox1.Value = "Initiatiefase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Initiatiefase"
    ElseIf ComboBox1.Value = "Definitiefase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Definitiefase"
    ElseIf ComboBox1.Value = "Ontwerp fase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Ontwerp fase"
    ElseIf ComboBox1.Value = "Aanbestedingsfase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Aanbestedingsfase"
    ElseIf ComboBox1.Value = "Uitvoeringsfase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Uitvoeringsfase"
    ElseIf ComboBox1.Value = "Nazorgfase" Then
        ActiveDocument.ContentTypeProperties("Phase").Value = "Nazorgfase"
    Else
        MsgBox ("Geen projectfase ingevoerd. Het document wordt automatisch gemarkeerd als General Document.")
        ActiveDocument.ContentTypeProperties("Phase").Value = "General Document"
        
    End If
    
    Me.Hide

End Sub

Private Sub UserForm_Activate()
    'Function to count&print pages
    Dim TtlPgs As Integer
    TtlPgs = ActiveDocument.Range.Information(wdNumberOfPagesInDocument)
    TextBox5.Value = TtlPgs

    'Add items to Dropdown menu
    ComboBox1.AddItem ("Opstartfase")
    ComboBox1.AddItem ("Initiatiefase")
    ComboBox1.AddItem ("Definitiefase")
    ComboBox1.AddItem ("Ontwerp fase")
    ComboBox1.AddItem ("Aanbestedingsfase")
    ComboBox1.AddItem ("Uitvoeringsfase")
    ComboBox1.AddItem ("Nazorgfase")
    
End Sub

Private Sub TextBox3_Change()

    OnlyNumbers

End Sub

Public Sub OnlyNumbers()

    If TypeName(Me.ActiveControl) = "TextBox" Then

        With Me.ActiveControl

            If Not IsNumeric(.Value) And .Value <> vbNullString Then

                MsgBox "In dit veld zijn alleen nummers toegestaan."

                .Value = vbNullString
            End If
        End With
    End If
End Sub

Public Sub UpdateBookmark(BookmarkToUpdate As String, TextToUse As String)
    
    Dim BMRange As Range
    Set BMRange = ActiveDocument.Bookmarks(BookmarkToUpdate).Range
    BMRange.Text = TextToUse
    ActiveDocument.Bookmarks.Add BookmarkToUpdate, BMRange

End Sub
