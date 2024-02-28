Sub Recs()
Dim styleToSearch As String
Dim tableIndex As Integer
tableIndex = 1
Dim Found As Boolean
Dim tableToInsert As Table
Dim doesTableExist As Boolean

    styleToSearch = InputBox("Enter the style to search for:")
    ' Check if the table already exists
    For Each Table In ActiveDocument.Tables
        If Table.Descr = styleToSearch + " Table" Then
            Set tableToInsert = Table
            doesTableExist = True
        End If
    Next Table
    
    ' Search the document for the specified style
    For Each para In ActiveDocument.Paragraphs
        If para.Style = styleToSearch Then
            ' Check if the paragraph is empty
            If Len(para.Range.Text) > 1 Then
                Found = False
                ' If table does't exist, create it
                If tableToInsert Is Nothing Then
                    Set tableToInsert = ActiveDocument.Tables.Add(Selection.Range, 1, 3)
                    tableToInsert.Cell(1, 1).Range.Text = "No."
                    tableToInsert.Cell(1, 2).Range.Text = "Description"
                    tableToInsert.Cell(1, 3).Range.Text = "Status"
                    tableToInsert.Descr = styleToSearch + " Table"
                    doesTableExist = True
                ' If table exists, check if the paragraph is already in the table
                Else
                    For Each Row In tableToInsert.Rows
                        If InStr(Row.Cells(2).Range.Text, para.Range.Text) > 0 Then
                            Found = True
                            Exit For
                        End If
                    Next Row
                    tableIndex = tableToInsert.Rows.Count
                End If
                ' If the paragraph is not in the table, add it
                If doesTableExist = True Then
                    If Found = False Then
                        With tableToInsert
                            .Rows.Add
                            .Cell(tableToInsert.Rows.Count, 1).Range.Text = tableIndex
                            .Cell(tableToInsert.Rows.Count, 2).Range.Text = para.Range.Text
                            .Cell(tableToInsert.Rows.Count, 3).Range.Text = "Open"
                        End With
                    End If
                End If
                tableIndex = tableIndex + 1
            End If
        End If
    Next para
End Sub





