Attribute VB_Name = "udfremove_a"
Function udf_remove_a(letter_a As String)

'Formula will appear in excel spreadsheet
'Formula remove the letter a from a string

udf_remove_a = Windows.Application.WorksheetFunction.Substitute(letter_a, "a", "")

End Function

