wb = load_workbook(path_RESULT.filename)
ws=wb.active

i=1
while i<=6:
	ws.insert_cols(49)
	#ws.cell(row=1, column=48+i).value = 'NEWS'
	i=i+1
wb.save(path_RESULT.filename)
