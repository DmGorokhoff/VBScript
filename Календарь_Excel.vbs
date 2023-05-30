    cl_dayCur = Day(Date)
	cl_WeekdayCur = Weekday(Date, vbMonday)	
	cl_year = Year(Date)
    cl_month = Month(Date)

Set objShellApp = CreateObject("Shell.Application")
	objShellApp.MinimizeAll
set app = CreateObject("Excel.Application")
	app.visible = true
	app.workbooks.add
	app.worksheets(1).Name = cl_year & " г."	
		
fillRange()

Sub fillRange()

    'Dim arr28(), arr29(), arr30(), arr31(), arrMonth(), arrDayWeek(), arrRngMonth()
    Dim ind    
    'app.Cells.Clear
    
Set rngJan = app.Range("A3:G8")
Set rngDWeekJan = app.Range("A2:G2")
Set rngFeb = app.Range("I3:O8")
Set rngDWeekFeb = app.Range("I2:O2")
Set rngMar = app.Range("Q3:W8")
Set rngDWeekMar = app.Range("Q2:W2")
Set rngApr = app.Range("Y3:AE8")
Set rngDWeekApr = app.Range("Y2:AE2")
Set rngMay = app.Range("AG3:AM8")
Set rngDWeekMay = app.Range("AG2:AM2")
Set rngJun = app.Range("AO3:AU8")
Set rngDWeekJun = app.Range("AO2:AU2")
Set rngJul = app.Range("A12:G17")
Set rngDWeekJul = app.Range("A11:G11")
Set rngAug = app.Range("I12:O17")
Set rngDWeekAug = app.Range("I11:O11")
Set rngSep = app.Range("Q12:W17")
Set rngDWeekSep = app.Range("Q11:W11")
Set rngOkt = app.Range("Y12:AE17")
Set rngDWeekOkt = app.Range("Y11:AE11")
Set rngNov = app.Range("AG12:AM17")
Set rngDWeekNov = app.Range("AG11:AM11")
Set rngDec = app.Range("AO12:AU17")
Set rngDWeekDec = app.Range("AO11:AU11")
	
	arrDate = array(("01.01." & cl_year), ("01.02." & cl_year), ("01.03." & cl_year), ("01.04." & cl_year), ("01.05." & cl_year), ("01.06." & cl_year), ("01.07." & cl_year), ("01.08." & cl_year), ("01.09." & cl_year), ("01.10." & cl_year), ("01.11." & cl_year), ("01.12." & cl_year), ("01.01." & (cl_year + 1)))
    arrMonth = Array("Янв", "Фев", "Мар", "Апр", "Май", "Июн", "Июл", "Авг", "Сен", "Окт", "Ноя", "Дек")
    arrRngDWeek = Array(rngDWeekJan, rngDWeekFeb, rngDWeekMar, rngDWeekApr, rngDWeekMay, rngDWeekJun, rngDWeekJul, rngDWeekAug, rngDWeekSep, rngDWeekOkt, rngDWeekNov, rngDWeekDec)
    arrRngMonth = Array(rngJan, rngFeb, rngMar, rngApr, rngMay, rngJun, rngJul, rngAug, rngSep, rngOkt, rngNov, rngDec)
    arrCellMonth = Array("A1", "I1", "Q1", "Y1", "AG1", "AO1", "A10", "I10", "Q10", "Y10", "AG10", "AO10")    
    arrDayWeek = Array("пн", "вт", "ср", "чт", "пт", "сб", "вс")
    arr31 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30, 31)
    arr30 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29, 30)
    arr29 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28, 29)
    arr28 = Array(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18, 19, 20, 21, 22, 23, 24, 25, 26, 27, 28)
    
	'Заполнение полей календаря названиями месяцев с указанием года
    For i = 0 To 11
        app.Range(arrCellMonth(i)) = arrMonth(i) & " " & cl_year & " г."
		app.Range(arrCellMonth(i)).font.bold = true
    Next
        
    'Заполнение полей календаря названиями дней недели
	For j = 0 To 11
        For index = 0 To 6
            arrRngDWeek(j)(index + 1) = arrDayWeek(index)
			arrRngDWeek(j)(index + 1).font.bold = true
            arrRngDWeek(j)(6).Font.Color = vbRed
            arrRngDWeek(j)(7).Font.Color = vbRed
        Next
    Next
        
	'Заполнение полей календаря числами
    For ii = 0 To 11
        quantityDayMonth = Day(CDate(arrDate(ii + 1)) - 1)
        ind = Weekday(arrDate(ii), vbMonday)
        Select Case quantityDayMonth
            Case 31
                For Each i In arr31
                    arrRngMonth(ii)(ind) = i
                    ind = ind + 1
                Next
                    arrRngMonth(ii).Columns(6).Font.Color = vbRed
                    arrRngMonth(ii).Columns(7).Font.Color = vbRed            
            Case 30
                For Each i In arr30
                    arrRngMonth(ii)(ind) = i
                    ind = ind + 1
                Next
                    arrRngMonth(ii).Columns(6).Font.Color = vbRed
                    arrRngMonth(ii).Columns(7).Font.Color = vbRed					
			Case 29
                For Each i In arr29
                    arrRngMonth(ii)(ind) = i
                    ind = ind + 1
                Next
                    arrRngMonth(ii).Columns(6).Font.Color = vbRed
                    arrRngMonth(ii).Columns(7).Font.Color = vbRed
			Case 28
                For Each i In arr28
                    arrRngMonth(ii)(ind) = i
                    ind = ind + 1
                Next
                    arrRngMonth(ii).Columns(6).Font.Color = vbRed
                    arrRngMonth(ii).Columns(7).Font.Color = vbRed					
        End Select
    Next    
	'numberSearch = Weekday(DateAdd("d", -(day(date)), date) , vbMonday)
		
	for each count in arrRngMonth(cl_month-1)
		seacheInd = seacheInd + 1
		if count = cl_dayCur then
				app.Range("A3:AU7").Columns.AutoFit
				app.Range("A12:AU17").Columns.AutoFit
				'arrRngMonth(cl_month - 1)(cl_dayCur + 1).Interior.Color = vbGreen
				arrRngMonth(cl_month - 1)(seacheInd).Interior.Color = RGB (255,204,255)
				'arrRngMonth(cl_month - 1)(cl_WeekdayCur).Borders.Weight = xlThick	
				arrRngMonth(cl_month - 1)(seacheInd).Borders.Color = vbRed
				'arrRngMonth(cl_month - 1)(seacheInd).Borders.Weight = true
				arrRngMonth(cl_month - 1)(seacheInd).font.bold = true
				arrRngMonth(cl_month - 1)(seacheInd).font.color = vbRed
				arrRngMonth(cl_month - 1)(seacheInd).font.size = 11
				'msgbox "count.value = " & count.value & "; seacheInd = " & seacheInd
			exit for
		end if		
	next
	app.ScreenUpdating = True
    msgbox "The calendar is ready!", vbInformation, "Message"
	
End Sub