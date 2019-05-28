' EASY

Sub Easy()
	Dim CurrentTicker, NextTicker
	Dim NumOfTrades
	Dim CurrentRow, RunningVol
	
	NumOfTrades = Cells(Rows.Count, 1).End(xlUp).Row
	RunningVol = 0
	CurrentRow = 2
	 
	for i = 2 to NumOfTrades
		CurrentTicker = Cells(i, 1).Value
		NextTicker = Cells(i + 1, 1).Value
		RunningVol = RunningVol + Cells(i, 7).Value

		If CurrentTicker <> NextTicker Then
			Cells(CurrentRow, 9).Value = CurrentTicker
			Cells(CurrentRow, 12).Value = RunningVol
			RunningVol = 0
			CurrentRow = CurrentRow + 1
		End If
	Next
End Sub


' MODERATE

Sub Moderate()
	Dim CurrentTicker, NextTicker
	Dim NumOfTrades, RunningVol
	Dim CurrentRow

	Dim TickerFirstTime As Boolean
	Dim OpenPriceYear, YearlyChange, PercentChange, ClosePriceYear
	
	NumOfTrades = Cells(Rows.Count, 1).End(xlUp).Row
	RunningVol = 0
	CurrentRow = 2
	TickerFirstTime = True
	 
	for i = 2 to NumOfTrades
		If TickerFirstTime = True Then
			OpenPriceYear = Cells(i, 3).Value
			TickerFirstTime = False
		End If
		
		CurrentTicker = Cells(i, 1).Value
		NextTicker = Cells(i + 1, 1).Value
		RunningVol = RunningVol + Cells(i, 7).Value

		If CurrentTicker <> NextTicker Then
			ClosePriceYear = Cells(i, 6).Value
			YearlyChange = ClosePriceYear - OpenPriceYear
			PercentChange = YearlyChange / OpenPriceYear
			Cells(CurrentRow, 9).Value = CurrentTicker
			Cells(CurrentRow, 10).Value = YearlyChange
				If YearlyChange >= 0 Then
					Cells(CurrentRow, 10).Interior.Color = vbGreen
				Else
					Cells(CurrentRow, 10).Interior.Color = vbRed
				End If
			Cells(CurrentRow, 11).Value = PercentChange
			Cells(CurrentRow, 11).NumberFormat = "0.00%"
			Cells(CurrentRow, 12).Value = RunningVol
			RunningVol = 0
			CurrentRow = CurrentRow + 1
			TickerFirstTime = True
		End If
	Next
End Sub


' HARD

Sub Hard()
	Dim CurrentTicker, NextTicker
	Dim TickerFirstTime As Boolean
	Dim NumOfTrades
	Dim CurrentRow, RunningVol
	Dim OpenPriceYear, ClosePriceYear, YearlyChange, PercentChange
	
	NumOfTrades = Cells(Rows.Count, 1).End(xlUp).Row
	RunningVol = 0
	TickerFirstTime = True
	CurrentRow = 2
	 
	for i = 2 to NumOfTrades
		If TickerFirstTime = True Then
			OpenPriceYear = Cells(i, 3).Value
			TickerFirstTime = False
		End If
		
		CurrentTicker = Cells(i, 1).Value
		NextTicker = Cells(i + 1, 1).Value
		RunningVol = RunningVol + Cells(i, 7).Value

		If CurrentTicker <> NextTicker Then
			ClosePriceYear = Cells(i, 6).Value
			YearlyChange = ClosePriceYear - OpenPriceYear
			PercentChange = YearlyChange / OpenPriceYear
			Cells(CurrentRow, 9).Value = CurrentTicker
			Cells(CurrentRow, 10).Value = YearlyChange
				If YearlyChange >= 0 Then
					Cells(CurrentRow, 10).Interior.Color = vbGreen
				Else
					Cells(CurrentRow, 10).Interior.Color = vbRed
				End If
			Cells(CurrentRow, 11).Value = PercentChange
			Cells(CurrentRow, 11).NumberFormat = "0.00%"
			Cells(CurrentRow, 12).Value = RunningVol
			RunningVol = 0
			CurrentRow = CurrentRow + 1
			TickerFirstTime = True
		End If
	Next

	Dim NumOfTickers, CurrentRowTicker
	Dim IncreaseTicker, DecreaseTicker, VolTicker
	Dim CurrentPercentIncrease, CurrentPercentChange
	Dim GreatestPercentIncrease, GreatestPercentDecrease, GreatestVol

	NumOfTickers = Cells(Rows.Count, 9).End(xlUp).Row
	
	CurrentRow = 2

	GreatestPercentIncrease = 0
	GreatestPercentDecrease = Cells(CurrentRow, 11).Value
	GreatestVol = 0

	For Count = 2 to NumOfTickers
		CurrentRowTicker = Cells(CurrentRow, 9).Value
		CurrentPercentChange = Cells(CurrentRow, 11).Value
		CurrentVol = Cells(CurrentRow, 12).Value

		If CurrentPercentChange > GreatestPercentIncrease Then
			GreatestPercentIncrease = CurrentPercentChange
			IncreaseTicker = CurrentRowTicker
		End If

		If CurrentPercentChange < GreatestPercentDecrease Then
			GreatestPercentDecrease = CurrentPercentChange
			DecreaseTicker = CurrentRowTicker
		End If

		If CurrentVol > GreatestVol Then
			GreatestVol = CurrentVol
			VolTicker = CurrentRowTicker
		End If
		CurrentRow = CurrentRow + 1
	Next Count
		
	Range("R2").Value = IncreaseTicker
	Range("S2").Value = GreatestPercentIncrease
	Range("S2").NumberFormat = "0.00%"
	Range("R3").Value = DecreaseTicker
	Range("S3").Value = GreatestPercentDecrease
	Range("S3").NumberFormat = "0.00%"
	Range("R4").Value = VolTicker
	Range("S4").Value = GreatestVol
End Sub


' CHALLENGE

Sub Challenge()
    Dim xSh As Worksheet
    Application.ScreenUpdating = False
    For Each xSh In Worksheets
        xSh.Select
        Call HardChallenge
    Next
    Application.ScreenUpdating = True
End Sub

Sub HardChallenge()
	Dim CurrentTicker, NextTicker
	Dim TickerFirstTime As Boolean
	Dim NumOfTrades
	Dim CurrentRow, RunningVol
	Dim OpenPriceYear, ClosePriceYear, YearlyChange, PercentChange
	
	NumOfTrades = Cells(Rows.Count, 1).End(xlUp).Row
	RunningVol = 0
	TickerFirstTime = True
	CurrentRow = 2
	 
	for i = 2 to NumOfTrades
		If TickerFirstTime = True Then
			OpenPriceYear = Cells(i, 3).Value
			TickerFirstTime = False
		End If
		
		CurrentTicker = Cells(i, 1).Value
		NextTicker = Cells(i + 1, 1).Value
		RunningVol = RunningVol + Cells(i, 7).Value

		If CurrentTicker <> NextTicker Then
			ClosePriceYear = Cells(i, 6).Value
			YearlyChange = ClosePriceYear - OpenPriceYear
			PercentChange = YearlyChange / OpenPriceYear
			Cells(CurrentRow, 9).Value = CurrentTicker
			Cells(CurrentRow, 10).Value = YearlyChange
				If YearlyChange >= 0 Then
					Cells(CurrentRow, 10).Interior.Color = vbGreen
				Else
					Cells(CurrentRow, 10).Interior.Color = vbRed
				End If
			Cells(CurrentRow, 11).Value = PercentChange
			Cells(CurrentRow, 11).NumberFormat = "0.00%"
			Cells(CurrentRow, 12).Value = RunningVol
			RunningVol = 0
			CurrentRow = CurrentRow + 1
			TickerFirstTime = True
		End If
	Next

	Dim NumOfTickers, CurrentRowTicker
	Dim IncreaseTicker, DecreaseTicker, VolTicker
	Dim CurrentPercentIncrease, CurrentPercentChange
	Dim GreatestPercentIncrease, GreatestPercentDecrease, GreatestVol

	NumOfTickers = Cells(Rows.Count, 9).End(xlUp).Row
	
	CurrentRow = 2

	GreatestPercentIncrease = 0
	GreatestPercentDecrease = Cells(CurrentRow, 11).Value
	GreatestVol = 0

	For Count = 2 to NumOfTickers
		CurrentRowTicker = Cells(CurrentRow, 9).Value
		CurrentPercentChange = Cells(CurrentRow, 11).Value
		CurrentVol = Cells(CurrentRow, 12).Value

		If CurrentPercentChange > GreatestPercentIncrease Then
			GreatestPercentIncrease = CurrentPercentChange
			IncreaseTicker = CurrentRowTicker
		End If

		If CurrentPercentChange < GreatestPercentDecrease Then
			GreatestPercentDecrease = CurrentPercentChange
			DecreaseTicker = CurrentRowTicker
		End If

		If CurrentVol > GreatestVol Then
			GreatestVol = CurrentVol
			VolTicker = CurrentRowTicker
		End If
		CurrentRow = CurrentRow + 1
	Next Count
		
	Range("R2").Value = IncreaseTicker
	Range("S2").Value = GreatestPercentIncrease
	Range("S2").NumberFormat = "0.00%"
	Range("R3").Value = DecreaseTicker
	Range("S3").Value = GreatestPercentDecrease
	Range("S3").NumberFormat = "0.00%"
	Range("R4").Value = VolTicker
	Range("S4").Value = GreatestVol
End Sub


' EXTRA!

Sub Reset()
	Dim NumOfTickers

	NumOfTickers = Cells(Rows.Count, 9).End(xlUp).Row
	
	for i = 2 to NumOfTickers
		for j = 9 to 12
			Cells(i, j).Clear
		next j
	Next
	
	Range("R2:S4").Clear
End Sub
