Attribute VB_Name = "Module2"

'===============================================================================
' Function: GetPerfumeSourceFromColor
'
' Description:
'   Populates the "Source" field with either "house site" or "BaseNotes.com", depending on the highlighting in "Perfume" field.
'   This tracks whether the notes information came directly from the perfume house or from the BaseNotes.com database.
'
' Parameters:
'   cell (Range)
'       The cell to evaluate (from the "Perfume" field)
'
'   houseColor (Long, Optional)
'       RGB color for notes data sourced from the perfume house's website.
'       Default = RGB(255, 255, 0) (Yellow)
'
'   baseNotesColor (Long, Optional)
'       RGB color for notes data sourced from BaseNotes.com.
'       Default = RGB(146, 208, 80) (Green)
'
' Returns:
'   String
'       "house website"   ? if cell color matches houseColor
'       "BaseNotes.com"   ? if cell color matches baseNotesColor
'       "" (empty)        ? if no match
'
'
' Author: Summer Mengarelli, smengare@nd.edu
' Last Updated: 2026-04-02
'===============================================================================


Function GetPerfumeSourceFromColor(cell As Range, _
                       Optional houseColor As Long = -1, _
                       Optional baseNotesColor As Long = -1) As String

    If houseColor = -1 Then houseColor = RGB(255, 255, 0)
    If baseNotesColor = -1 Then baseNotesColor = RGB(146, 208, 80)
    
    If cell Is Nothing Then
        GetPerfumeSourceFromColor = ""
        Exit Function
    End If

    Select Case cell.Interior.Color
        Case houseColor
            GetPerfumeSourceFromColor = "house website"
        Case baseNotesColor
            GetPerfumeSourceFromColor = "BaseNotes.com"
        Case Else
            GetPerfumeSourceFromColor = ""
    End Select
End Function



