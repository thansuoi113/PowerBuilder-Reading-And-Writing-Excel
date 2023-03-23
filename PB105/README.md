# Direct reading and writing Excel files with Powerbuilder

![Image](https://github.com/freemantau/tplib_powerbuilder/blob/master/demo.png?raw=true)
## Writing data
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = Create n_tp_excel

book.createxls( cxls.TYPE_XLS)
//
int logoid
logoid = book.addpicture( "logo.png")
//
////fonts
n_tp_xlsfont textfont
textfont = book.addfont( )
textfont.setsize( 8)
textfont.setname( "Arial")

n_tp_xlsfont titlefont
titlefont = book.addfont( textfont )
titlefont.setsize( 38)
titlefont.setcolor( cxls.COLOR_GRAY25)

n_tp_xlsfont font12,font10
font12 = book.addfont( textfont )
font10 = book.addfont( textfont)
font12.setsize( 12)
font10.setsize( 10)

////format
n_tp_xlsformat textformat
textformat = book.addformat( )
textformat.setfont( textfont)
textformat.setalignh( cxls.ALIGNH_LEFT)

n_tp_xlsformat titleformat
titleformat = book.addformat( )
titleformat.setfont( titlefont)
titleformat.setalignh( cxls.ALIGNH_RIGHT)

n_tp_xlsformat companyformat
companyformat = book.addformat( )
companyformat.setfont( font12)

n_tp_xlsformat dateformat
dateformat = book.addformat(textformat)
dateformat.setnumformat( book.addcustomnumformat( "[$-409]mmmm\ d\,\ yyyy;@"))

n_tp_xlsformat phoneformat
phoneformat = book.addformat(textformat)
phoneformat.setnumformat( book.addcustomnumformat( "[<=9999999]###\-####;\(###\)\ ###\-####"))

n_tp_xlsformat borderformat
borderformat = book.addformat(textformat )
borderformat.setborder( cxls.BORDERSTYLE_THIN)
borderformat.setbordercolor( cxls.COLOR_GRAY25)
borderformat.setalignv( cxls.ALIGNV_CENTER)

n_tp_xlsformat percentformat
percentformat = book.addformat(borderformat )
percentformat.setnumformat( book.addcustomnumformat( "#%_)"))
percentformat.setalignh( cxls.ALIGNH_RIGHT)

n_tp_xlsformat textrightformat
textrightformat = book.addformat( textformat )
textRightFormat.setAlignH(cxls.ALIGNH_RIGHT);
textRightFormat.setAlignV(cxls.ALIGNV_CENTER);

n_tp_xlsformat thankformat
thankformat = book.addformat( )
thankFormat.setFont(font10);
thankFormat.setAlignH(cxls.ALIGNH_CENTER);

n_tp_xlsformat dollarformat
dollarformat = book.addformat( borderformat )
dollarformat.setnumformat( book.addcustomnumformat( "_($* # ##0.00_);_($* (# ##0.00);_($* -??_);_(@_)"))
//actions
n_tp_xlssheet sheet
sheet = book.addsheet( "Sales Receipt表")


sheet.setdisplaygridlines(false)

sheet.setCol(1, 1, 36)
sheet.setCol(0, 0, 10)
sheet.setCol(2, 4, 11)

sheet.setRow(2, 47.25)
sheet.writeStr(2, 1, "Sales Receipt", titleFormat)
sheet.setMerge(2, 2, 1, 4)
sheet.setPicture(2, 1, logoId,1.0,0,0)


sheet.writeStr(4, 0, "Apricot Ltd.", companyFormat)
sheet.writeStr(4, 3, "Date:", textFormat)


sheet.writeFormula(4, 4, "TODAY()", dateFormat)

sheet.writeStr(5, 3, "Receipt #:", textFormat)


sheet.writeNum(5, 4, 652, textFormat)

sheet.writeStr(8, 0, "Sold to:", textFormat)
sheet.writeStr(8, 1, "John Smith", textFormat)
sheet.writeStr(9, 1, "Pineapple Ltd.", textFormat)
sheet.writeStr(10, 1, "123 Dreamland Street", textFormat)
sheet.writeStr(11, 1, "Moema, 52674", textFormat)


sheet.writeNum(12, 1, 2659872055, phoneFormat)

sheet.writeStr(14, 0, "Item #", textFormat)
sheet.writeStr(14, 1, "Description", textFormat)
sheet.writeStr(14, 2, "Qty", textFormat)
sheet.writeStr(14, 3, "Unit Price", textFormat)
sheet.writeStr(14, 4, "Line Total", textFormat)

int row,col
string s
for row = 15 to 37
	sheet.setRow(row, 15)
	for col = 0 to 2
		//TODO
		sheet.writeBlank(row, col, borderFormat)
	next
	//TODO
	sheet.writeBlank(row, 3, dollarFormat)

	//TODO
	//s = sprintf('IF(C{1}>0;ABS(C{2}*D{3});"")',row + 1,row + 1,row + 1)
	s = 'IF(C' +string(row + 1)+ '>0;ABS(C' +string(row + 1)+ '*D' +string(row + 1)+ ');"")'
	sheet.writeFormula(row, 4, s, dollarFormat)
next 



sheet.writeStr(38, 3, "Subtotal ", textRightFormat)
sheet.writeStr(39, 3, "Sales Tax ", textRightFormat)
sheet.writeStr(40, 3, "Total ", textRightFormat)
sheet.writeFormula(38, 4, "SUM(E16:E38)", dollarFormat)
sheet.writeNum(39, 4, 0.2, percentFormat)
sheet.writeFormula(40, 4, "E39+E39*E40", dollarFormat)
sheet.setRow(38, 15)
sheet.setRow(39, 15)
sheet.setRow(40, 15)

sheet.writeStr(42, 0, "Thank you for your business!", thankFormat)
sheet.setMerge(42, 42, 0, 4)

// items

sheet.writeNum(15, 0, 45, borderFormat)
sheet.writeStr(15, 1, "Grapes", borderFormat)
sheet.writeNum(15, 2, 250, borderFormat)
sheet.writeNum(15, 3, 4.5, dollarFormat)

sheet.writeNum(16, 0, 12, borderFormat)
sheet.writeStr(16, 1, "Bananas", borderFormat)
sheet.writeNum(16, 2, 480, borderFormat)
sheet.writeNum(16, 3, 1.4, dollarFormat)

sheet.writeNum(17, 0, 19, borderFormat)
sheet.writeStr(17, 1, "Apples", borderFormat)
sheet.writeNum(17, 2, 180, borderFormat)
sheet.writeNum(17, 3, 2.8, dollarFormat)

book.save("receipt.xls",false)

destroy book




Messagebox('','complete!')
```
</details>

## reading data
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = Create n_tp_excel
string sout
sout = ''
long row,col,rowlast,collast
int celltype
if book.load( "data.xlsx") Then 
	n_tp_xlssheet sheet
	sheet = book.getsheet( 2)
	rowlast = sheet.lastrow( )
	collast = sheet.lastcol( )
	for row = 0 to rowlast - 1
		for col = 0 to collast - 1
			celltype = sheet.celltype(row,col)
			sout += "(" + string(row) + "," + string(col) + ") = " 
			if sheet.isformula( row,col /*long col */) Then
				sout += sheet.readformula( row,col /*long col */) + "[formula]~n"
			else
				choose case celltype
					case cxls.celltype_empty
						sout += "[empty]~n"
					case cxls.celltype_number
						sout += string(sheet.readnum( row,col /*long col */)) + "[number]~n"
					case cxls.celltype_datetime	/*use dateunpack*/
						sout += string(sheet.readnum( row,col /*long col */)) + "[date]~n"
					case cxls.celltype_string
						sout += sheet.readstr( row,col /*long col */) + "[string]~n"
					case cxls.celltype_boolean
						sout += string(sheet.readbool( row,col /*long col */)) + "[boolean]~n"	
					case cxls.celltype_blank
						sout += "[blank]~n"
					case cxls.celltype_error
						sout += "[error]~n"
				end choose
			end if
		next
	next	
end if

messagebox('reading data',sout)
destroy book
```
</details>

## Placing pictures
<details>
<summary>show code</summary>
	
```cpp
boolean lb
n_tp_excel book
book = Create n_tp_excel
lb = book.createxls( cxls.TYPE_XLSX)
long id
if lb Then
	id = book.addpicture( "1.jpg")
	if id = -1 Then
		messagebox('','picture not found')
		return
	end if
	n_tp_xlssheet sheet
	sheet = book.addsheet( "sheet1")
	sheet.setpicture( 10/*long row*/,1 /*long col*/, id/*long pictureid*/,1 /*double scale*/,0 /*long offset_x*/,0 /*long offset_y */)

	if book.save( "Placing pictures.xlsx",false /*boolean usetempfile */) Then
		messagebox('','complete')
	end if	
end if
destroy book
```
</details>

## Writing formulas
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)

n_tp_xlsformat alFormat
alFormat = book.addFormat()
alFormat.setAlignH(cxls.ALIGNH_LEFT)

n_tp_xlsformat arformat
arFormat = book.addFormat()
arFormat.setAlignH(cxls.ALIGNH_RIGHT)

n_tp_xlsformat alignDateFormat
alignDateFormat = book.addFormat(alFormat)
alignDateFormat.setNumFormat(cxls.NUMFORMAT_DATE)

n_tp_xlsfont linkFont
linkFont = book.addFont()
linkFont.setColor(cxls.COLOR_BLUE)
linkFont.setUnderline(cxls.UNDERLINE_SINGLE)

n_tp_xlsformat linkFormat
linkFormat = book.addFormat(alFormat)
linkFormat.setFont(linkFont)

n_tp_xlssheet sheet
sheet = book.addSheet("Sheet1")


sheet.setCol(0, 0, 27)
sheet.setCol(1, 1, 10)

sheet.writeNum(2, 1, 40, alFormat)
sheet.writeNum(3, 1, 30, alFormat)
sheet.writeNum(4, 1, 50, alFormat)

sheet.writeStr(6, 0, "SUM(B3:B5) = ", arFormat)        
sheet.writeFormula(6, 1, "SUM(B3:B5)", alFormat)        
sheet.writeStr(7, 0, "AVERAGE(B3:B5) = ", arFormat)        
sheet.writeFormula(7, 1, "AVERAGE(B3:B5)", alFormat)        
sheet.writeStr(8, 0, "MAX(B3:B5) = ", arFormat)        
sheet.writeFormula(8, 1, "MAX(B3:B5)", alFormat)        
sheet.writeStr(9, 0, "MIX(B3:B5) = ", arFormat)        
sheet.writeFormula(9, 1, "MIN(B3:B5)", alFormat)
sheet.writeStr(10, 0, "COUNT(B3:B5) = ", arFormat)      
sheet.writeFormula(10, 1, "COUNT(B3:B5)", alFormat)

sheet.writeStr(12, 0, 'IF(B7 > 100;"large";"small") = ', arFormat)      
sheet.writeFormula(12, 1, 'IF(B7 > 100;"large";"small")', alFormat)

sheet.writeStr(14, 0, "SQRT(25) = ", arFormat)      
sheet.writeFormula(14, 1, "SQRT(25)", alFormat)
sheet.writeStr(15, 0, "RAND() = ", arFormat)      
sheet.writeFormula(15, 1, "RAND()", alFormat)
sheet.writeStr(16, 0, "2*PI() = ", arFormat)      
sheet.writeFormula(16, 1, "2*PI()", alFormat)

sheet.writeStr(18, 0, 'UPPER("libxl") = ', arFormat)      
sheet.writeFormula(18, 1, 'UPPER("libxl")', alFormat)
sheet.writeStr(19, 0, 'LEFT("window";3) = ', arFormat)      
sheet.writeFormula(19, 1, 'LEFT("window";3)', alFormat)
sheet.writeStr(20, 0, 'LEN("string") = ', arFormat)      
sheet.writeFormula(20, 1, 'LEN("string")', alFormat)

sheet.writeStr(22, 0, "DATE(2010;3;11) = ", arFormat)      
sheet.writeFormula(22, 1, "DATE(2010;3;11)", alignDateFormat)
sheet.writeStr(23, 0, "DAY(B23) = ", arFormat)      
sheet.writeFormula(23, 1, "DAY(B23)", alFormat)
sheet.writeStr(24, 0, "MONTH(B23) = ", arFormat)      
sheet.writeFormula(24, 1, "MONTH(B23)", alFormat)
sheet.writeStr(25, 0, "YEAR(B23) = ", arFormat)      
sheet.writeFormula(25, 1, "YEAR(B23)", alFormat)
sheet.writeStr(26, 0, "DAYS360(B23;TODAY()) = ", arFormat)      
sheet.writeFormula(26, 1, "DAYS360(B23;TODAY())", alFormat)

sheet.writeStr(28, 0, "B3+100*(2-COS(0)) = ", arFormat)      
sheet.writeFormula(28, 1, "B3+100*(2-COS(0))", alFormat)
sheet.writeStr(29, 0, "ISNUMBER(B29) = ", arFormat)      
sheet.writeFormula(29, 1, "ISNUMBER(B29)", alFormat)
sheet.writeStr(30, 0, "AND(1;0) = ", arFormat)      
sheet.writeFormula(30, 1, "AND(1;0)", alFormat)

sheet.writeStr(32, 0, "HYPERLINK() = ", arFormat)
sheet.writeFormula(32, 1, 'HYPERLINK("http://www.libxl.com")', linkFormat)

if book.save("formula.xlsx",false) then
		messagebox('','complete')
end if	
destroy book
```
</details>

## Reading and writing date/time values

<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)

n_tp_xlsformat format1,format2,format3,format4
format1 = book.addFormat()
format1.setNumFormat(cxls.NUMFORMAT_DATE)

format2 = book.addFormat()
format2.setNumFormat(cxls.NUMFORMAT_CUSTOM_MDYYYY_HMM)

format3 = book.addFormat()
format3.setNumFormat(book.addCustomNumFormat("d mmmm yyyy"))

format4 = book.addFormat()
format4.setNumFormat(cxls.NUMFORMAT_CUSTOM_HMM_AM)

n_tp_xlssheet sheet
sheet = book.addSheet("Sheet1")


sheet.setCol(1, 1, 15)

// writing

sheet.writeNum(2, 1, book.datePack(2010, 3, 11,0,0,0,0), format1)       
sheet.writeNum(3, 1, book.datePack(2010, 3, 11, 10, 25, 55,0), format2)
sheet.writeNum(4, 1, book.datePack(2010, 3, 11,0,0,0,0), format3)       
sheet.writeNum(5, 1, book.datePack(2010, 3, 11, 10, 25, 55,0), format4)

// reading

long year, month, day , hour, min, sec,msec

book.dateUnpack(sheet.readNum(3, 1), ref year, ref month, ref day, ref hour, ref min, ref sec , ref msec)

messagebox('datetime',			"year:" 		+ string(year) + "~n" + &
										"month:" 	+ string(month) +  "~n" + &
										"day:" 		+ string(day) +  "~n" + &
										"hour:"		+ string(hour) +  "~n" + &
										"min:"	 	+ string(min) +  "~n" + &
										"second:" 	+ string(sec) +  "~n" + &
										"msec:"	 	+ string(msec) )	


if book.save("datetime.xlsx",false) then
	Messagebox('','complete')
end if
destroy book
```
</details>

## Access to sheet by name
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = Create n_tp_excel

long sheetcount,sheetindex
string sheetname
if book.load( "data.xlsx") Then
	sheetcount = book.getsheetcount( )
	for sheetindex = 0 to sheetcount  - 1	
		sheetname = book.getsheetname( sheetindex)
		if sheetname = "mysheetname" Then
			/*get your sheetbyname*/
			/*
				sheet  = book.getsheet( sheetindex)
			*/
		end if
	next
end if
```
</details>

## Merging cells
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)

n_tp_xlsformat format
format = book.addFormat();
format.setAlignH(cxls.ALIGNH_CENTER);
format.setAlignV(cxls.ALIGNV_CENTER);

n_tp_xlssheet sheet
sheet = book.addSheet("Sheet1")

sheet.writeStr(3, 1, "Hello World !", format)

sheet.setMerge(3, 5, 1, 5)

sheet.setMerge(7, 20, 1, 2)
sheet.setMerge(7, 20, 4, 5)

sheet.writeNum(7, 1, 1, format)
sheet.writeNum(7, 4, 2, format)


if book.save( "merge.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
```
</details>

## Inserting rows and columns
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)


n_tp_xlssheet sheet
sheet = book.addSheet("Sheet1")

n_tp_xlsformat format
format = book.addformat()

long row,col
for row = 1 to 30
  for col = 0 to 10
		sheet.writeNum(row, col,rand(10),format)
	next
next

sheet.insertRow(5, 10,false)
sheet.insertRow(20, 22,false)

sheet.insertCol(4, 5,false)
sheet.insertCol(8, 8,false)


if book.save( "insert.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
```
</details>

## Using number formats
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)

n_tp_xlssheet sheet
sheet = book.addSheet("my")

sheet.setCol(0, 0, 38)
sheet.setCol(1, 1, 10)    


// built-in number formats
n_tp_xlsformat format1,format2,format3,format4,format5,format6,format7,format8,format9,format10,format11,format12

format1 = book.addFormat()
format1.setNumFormat(cxls.NUMFORMAT_NUMBER_D2)

sheet.writeStr(3, 0, "NUMFORMAT_NUMBER_D2")
sheet.writeNum(3, 1, 2.5681, format1)

format2 = book.addFormat()
format2.setNumFormat(cxls.NUMFORMAT_NUMBER_SEP)

sheet.writeStr(4, 0, "NUMFORMAT_NUMBER_SEP")
sheet.writeNum(4, 1, 2500000, format2)

format3 = book.addFormat()
format3.setNumFormat(cxls.NUMFORMAT_CURRENCY_NEGBRA)

sheet.writeStr(5, 0, "NUMFORMAT_CURRENCY_NEGBRA")
sheet.writeNum(5, 1, -500, format3)

format4 = book.addFormat()
format4.setNumFormat(cxls.NUMFORMAT_PERCENT)

sheet.writeStr(6, 0, "NUMFORMAT_PERCENT")
sheet.writeNum(6, 1, -0.25, format4)

format5 = book.addFormat()
format5.setNumFormat(cxls.NUMFORMAT_SCIENTIFIC_D2)

sheet.writeStr(7, 0, "NUMFORMAT_SCIENTIFIC_D2")
sheet.writeNum(7, 1, 890, format5)

format6 = book.addFormat()
format6.setNumFormat(cxls.NUMFORMAT_FRACTION_ONEDIG)

sheet.writeStr(8, 0, "NUMFORMAT_FRACTION_ONEDIG")
sheet.writeNum(8, 1, 0.75, format6)

format7 = book.addFormat()
format7.setNumFormat(cxls.NUMFORMAT_DATE)

sheet.writeStr(9, 0, "NUMFORMAT_DATE")
sheet.writeNum(9, 1, book.datePack(2020, 5, 16,0,0,0,0), format7)

format8 = book.addFormat()
format8.setNumFormat(cxls.NUMFORMAT_CUSTOM_MON_YY)

sheet.writeStr(10, 0, "NUMFORMAT_CUSTOM_MON_YY")
sheet.writeNum(10, 1, book.datePack(2020, 5, 16,0,0,0,0), format8)

// custom number formats

format9 = book.addFormat()
format9.setNumFormat(book.addCustomNumFormat("#.###"))    

sheet.writeStr(12, 0, "#.###")
sheet.writeNum(12, 1, 20.5627, format9)

format10 = book.addFormat()
format10.setNumFormat(book.addCustomNumFormat("#.00"))    

sheet.writeStr(13, 0, "#.00")
sheet.writeNum(13, 1, 4.8, format10)

format11 = book.addFormat()
format11.setNumFormat(book.addCustomNumFormat('0.00 "dollars"'))    

sheet.writeStr(14, 0, '0.00 "dollars"')
sheet.writeNum(14, 1, 1.23, format11)

format12 = book.addFormat()
format12.setNumFormat(book.addCustomNumFormat("[Red][<=100];[Green][>100]"))    

sheet.writeStr(15, 0, "[Red][<=100];[Green][>100]")
sheet.writeNum(15, 1, 60, format12)



if book.save( "numformats.xlsx"/*string filename*/,false /*boolean usetempfile */) then
	messagebox('','complete')
end if
destroy book
```
</details>

## Aligning, colors and borders
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel

book.createxls( cxls.type_xlsx)

n_tp_xlssheet sheet
sheet = book.addSheet("my")

sheet.setDisplayGridlines(false)

sheet.setCol(1, 1, 30)
sheet.setCol(3, 3, 11.4)
sheet.setCol(4, 4, 2)
sheet.setCol(5, 5, 15)
sheet.setCol(6, 6, 2)
sheet.setCol(7, 7, 15.4)

string nameAlignH[] = {"ALIGNH_LEFT", "ALIGNH_CENTER", "ALIGNH_RIGHT"}
int alignH[] = {cxls.ALIGNH_LEFT, cxls.ALIGNH_CENTER, cxls.ALIGNH_RIGHT}

int i

for i = 1 to upperbound(alignH)
	n_tp_xlsformat format
	format = book.addFormat()
	format.setAlignH(alignH[i])
	format.setBorder(cxls.borderstyle_thin)
	sheet.writeStr(i * 2 + 2, 1, nameAlignH[i], format)
next  

string nameAlignV[] = {"ALIGNV_TOP", "ALIGNV_CENTER", "ALIGNV_BOTTOM"}
long alignV[] = {cxls.ALIGNV_TOP, cxls.ALIGNV_CENTER, cxls.ALIGNV_BOTTOM}

for i = 1 to upperbound(alignV)
	format = book.addFormat()
	format.setAlignV(alignV[i])
	format.setBorder(cxls.borderstyle_thin)
	sheet.writeStr(4, i * 2 + 1, nameAlignV[i], format)
	sheet.setMerge(4, 8, i * 2 + 1, i * 2 + 1)
next

string nameBorderStyle[] = {"BORDERSTYLE_MEDIUM", "BORDERSTYLE_DASHED", &
										 "BORDERSTYLE_DOTTED", "BORDERSTYLE_THICK",& 
										 "BORDERSTYLE_DOUBLE", "BORDERSTYLE_DASHDOT"}
long borderStyle[] = {cxls.BORDERSTYLE_MEDIUM, cxls.BORDERSTYLE_DASHED,cxls.BORDERSTYLE_DOTTED, &
								cxls.BORDERSTYLE_THICK, cxls.BORDERSTYLE_DOUBLE, cxls.BORDERSTYLE_DASHDOT}

for i = 1 to upperbound(nameBorderStyle)       
	format = book.addFormat()
	format.setBorder(borderStyle[i])
	sheet.writeStr(i * 2 + 12, 1, nameBorderStyle[i], format)
next 

string nameColors[] = {"COLOR_RED", "COLOR_BLUE", "COLOR_YELLOW", &
								  "COLOR_PINK", "COLOR_GREEN", "COLOR_GRAY25"}
long colors[] = {cxls.COLOR_RED, cxls.COLOR_BLUE, cxls.COLOR_YELLOW, cxls.COLOR_PINK, cxls.COLOR_GREEN, &
				 cxls.COLOR_GRAY25}
long fillPatterns[] = {cxls.FILLPATTERN_GRAY50, cxls.FILLPATTERN_HORSTRIPE, &
								 cxls.FILLPATTERN_VERSTRIPE, cxls.FILLPATTERN_REVDIAGSTRIPE,&
								 cxls.FILLPATTERN_THINVERSTRIPE, cxls.FILLPATTERN_THINHORCROSSHATCH}


for i = 1 to upperbound(nameColors)
	n_tp_xlsformat format1
	format1 = book.addFormat()
	format1.setFillPattern(cxls.FILLPATTERN_SOLID)
	format1.setPatternForegroundColor(colors[i])
	sheet.writeBlank(i * 2 + 12, 3, format1)
	
	n_tp_xlsformat format2
	format2 = book.addFormat()
	format2.setFillPattern(fillPatterns[i])
	format2.setPatternForegroundColor(colors[i])
	sheet.writeBlank(i * 2 + 12, 5, format2)
	
	n_tp_xlsfont font
	font = book.addFont()
	font.setColor(colors[i])
	
	n_tp_xlsformat format3
	format3 = book.addFormat()
	format3.setBorder(cxls.borderstyle_thin)
	format3.setBorderColor(colors[i])        
	format3.setFont(font)
	sheet.writeStr(i * 2 + 12, 7, nameColors[i], format3)
next


if book.save("acb.xlsx",false) then
	messagebox('','complete')
end if

destroy book
```
</details>

## Customizing fonts
<details>
<summary>show code</summary>
	
```cpp
n_tp_excel book
book = create n_tp_excel
book.createxls( cxls.type_xlsx)

n_tp_xlssheet sheet
sheet = book.addSheet("Sheet1")

string fonts[] = {"Aria", "Arial Black", "Comic Sans MS", "Courier New",&
							"Impact", "Times New Roman", "Verdana"}

int i
for i = 1 to upperbound(fonts)
	n_tp_xlsfont font
	font = book.addFont()
	font.setSize(16)
	font.setName(fonts[i])
	n_tp_xlsformat format
	format = book.addFormat()
	format.setFont(font)
	sheet.writeStr(i + 1, 3, fonts[i], format)
next

int fontSize[] = {8, 10, 12, 14, 16, 20, 25}

for i = 1 to upperbound(fontSize) 
	font = book.addFont()
	font.setSize(fontSize[i])        
	format = book.addFormat()
	format.setFont(font)
	sheet.writeStr(i + 1, 7, "Text", format)
next  

font = book.addFont()
font.setSize(16)
format = book.addFormat()        
format.setRotation(255)
format.setFont(font)
sheet.writeStr(2, 9, "Vertica", format)
sheet.setMerge(2, 8, 9, 9)

n_tp_xlsfont boldFont
boldFont = book.addFont()
boldFont.setBold(true)
n_tp_xlsformat boldFormat
boldFormat = book.addFormat()
boldFormat.setFont(boldFont)

n_tp_xlsfont italicFont
italicFont = book.addFont()
italicFont.setItalic(true) 
n_tp_xlsformat italicFormat
italicFormat = book.addFormat()
italicFormat.setFont(italicFont)

n_tp_xlsfont underlineFont
underlineFont = book.addFont()
underlineFont.setUnderline(cxls.UNDERLINE_SINGLE)
n_tp_xlsformat underlineFormat
underlineFormat = book.addFormat()
underlineFormat.setFont(underlineFont)

n_tp_xlsfont strikeoutFont
strikeoutFont = book.addFont()
strikeoutFont.setStrikeOut(true)
n_tp_xlsformat strikeoutFormat
strikeoutFormat = book.addFormat()
strikeoutFormat.setFont(strikeoutFont)

sheet.writeStr(2, 1, "Norma")
sheet.writeStr(3, 1, "Bold", boldFormat)      
sheet.writeStr(4, 1, "Italic", italicFormat)      
sheet.writeStr(5, 1, "Underline", underlineFormat)      
sheet.writeStr(6, 1, "Strikeout", strikeoutFormat)    

if book.save("fonts.xlsx",false) Then
	messagebox('','complete')
end if

destroy book
```
</details>

