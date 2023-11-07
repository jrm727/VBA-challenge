# VBA-challenge
References: 
     https://www.statology.org/vba-percentage-format/
     'format as a percent 
     ws.Cells(ticker_name_summary, 11).NumberFormat = "0.00%"


  https://www.educba.com/vba-max/
  &
  Further Assistance From: Kelci Griffin (Classmate) via Slack Channel
  
  'Find Greatest % Increase/Decrease and Volume 
   For i = 2 To Last_Row_Ticker_Summary
   
If ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.MAX(ws.Range("K2:K" &     Last_Row_Ticker_Summary)) Then
  ws.Cells(2, 15).Value = ws.Cells(i, 9).Value
  ws.Cells(2, 16).Value = ws.Cells(i, 11).Value
  ws.Cells(2, 16).NumberFormat = "0.00%"
  
  ElseIf ws.Cells(i, 11).Value = ws.Application.WorksheetFunction.Min(ws.Range("K2:K" & Last_Row_Ticker_Summary)) Then
  ws.Cells(3, 15).Value = ws.Cells(i, 9).Value
  ws.Cells(3, 16).Value = ws.Cells(i, 11).Value
  ws.Cells(3, 16).NumberFormat = "0.00%"
                
  ElseIf ws.Cells(i, 12).Value = ws.Application.WorksheetFunction.MAX(ws.Range("L2:L" & Last_Row_Ticker_Summary)) Then
  ws.Cells(4, 15).Value = Cells(i, 9).Value
  ws.Cells(4, 16).Value = Cells(i, 12).Value
   
