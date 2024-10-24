Sub UpdatePivotTableBasedOnSelections()
    Dim pt As PivotTable
    Dim ws As Worksheet
    Dim currencySelection As String
    Dim premiumSelection As String
    Dim ccToUSDSelection As String

    ' Set the worksheet and the Pivot Table
    Set ws = ThisWorkbook.Sheets("SheetWithPivot")  ' Adjust to your sheet name
    Set pt = ws.PivotTables("PivotTable1")  ' Adjust to your Pivot Table name

    ' Get the user's selections
    currencySelection = ws.Range("A1").Value  ' Currency dropdown in A1 (Original/Settlement)
    premiumSelection = ws.Range("A2").Value  ' Premiums dropdown in A2 (GWP, NWP, GPP, NPP)
    ccToUSDSelection = ws.Range("A3").Value  ' Convert to USD dropdown in A3 (Yes/No)

    ' Step 1: Show either the Original or USD-Converted Columns
    ' Assuming that for each premium type, you have two columns: Original and USD Converted
    Select Case currencySelection
        Case "Original"
            ' If Original is selected, show the Original Currency columns
            pt.PivotFields("GWP (Original)").Orientation = xlDataField
            pt.PivotFields("NWP (Original)").Orientation = xlDataField
            pt.PivotFields("GPP (Original)").Orientation = xlDataField
            pt.PivotFields("NPP (Original)").Orientation = xlDataField

            ' Hide the USD-Converted fields
            pt.PivotFields("GWP (USD)").Orientation = xlHidden
            pt.PivotFields("NWP (USD)").Orientation = xlHidden
            pt.PivotFields("GPP (USD)").Orientation = xlHidden
            pt.PivotFields("NPP (USD)").Orientation = xlHidden

        Case "Settlement"
            ' If Settlement is selected, show the USD Converted Columns
            pt.PivotFields("GWP (USD)").Orientation = xlDataField
            pt.PivotFields("NWP (USD)").Orientation = xlDataField
            pt.PivotFields("GPP (USD)").Orientation = xlDataField
            pt.PivotFields("NPP (USD)").Orientation = xlDataField

            ' Hide the Original Currency fields
            pt.PivotFields("GWP (Original)").Orientation = xlHidden
            pt.PivotFields("NWP (Original)").Orientation = xlHidden
            pt.PivotFields("GPP (Original)").Orientation = xlHidden
            pt.PivotFields("NPP (Original)").Orientation = xlHidden
    End Select

    ' Step 2: Adjust for Premium Selection (GWP, NWP, GPP, NPP)
    ' Ensure that only the selected premium type is shown in the Pivot Table
    With pt
        .ClearTable
        Select Case premiumSelection
            Case "GWP"
                pt.AddDataField pt.PivotFields("GWP (Original)")
            Case "NWP"
                pt.AddDataField pt.PivotFields("NWP (Original)")
            Case "GPP"
                pt.AddDataField pt.PivotFields("GPP (Original)")
            Case "NPP"
                pt.AddDataField pt.PivotFields("NPP (Original)")
        End Select
    End With

    ' Step 3: Refresh the Pivot Table
    pt.RefreshTable
End Sub



--python
df = pd.DataFrame(data)

data = {
    'Claim_id': ['ID4562019', 'ID4562019', 'ID4572019', 'ID4572019'],
    'Month': [7, 7, 8, 8],
    'Request_Number': [1, 2, 1, 2],
    'Claim_Amount': [250, 300, 300, 350]

# Group by 'Claim_id' and 'Month', and then find the max Request_Number and corresponding Claim_Amount
result = df.loc[df.groupby(['Claim_id', 'Month'])['Request_Number'].idxmax()]

# Reset the index if needed and display the result
result = result.reset_index(drop=True)
print(result)
