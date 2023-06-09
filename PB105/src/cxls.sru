$PBExportHeader$cxls.sru
forward
global type cxls from nonvisualobject
end type
end forward

global type cxls from nonvisualobject
end type
global cxls cxls

type variables
//File Type
Constant Int TYPE_XLS = 0
Constant Int TYPE_XLSX	 = 	1

//cell type
Constant UInt CELLTYPE_EMPTY 	 = 0
Constant UInt CELLTYPE_NUMBER = 1
Constant UInt CELLTYPE_STRING = 2
Constant UInt CELLTYPE_BOOLEAN 	 = 3
Constant UInt CELLTYPE_BLANK 	 = 4
Constant UInt CELLTYPE_ERROR 	 = 5
Constant UInt CELLTYPE_DATETIME 	 = 6

//alignh
Constant UInt ALIGNH_GENERAL 	 = 	0
Constant UInt ALIGNH_LEFT		 = 	1
Constant UInt ALIGNH_CENTER	 = 	2
Constant UInt ALIGNH_RIGHT		 = 	3
Constant UInt ALIGNH_FILL		 = 	4
Constant UInt ALIGNH_JUSTIFY	 = 	5
Constant UInt ALIGNH_MERGE		 = 	6
Constant UInt ALIGNH_DISTRIBUTED	 = 	7

//alignv
Constant UInt ALIGNV_TOP		 = 	0
Constant UInt ALIGNV_CENTER	 = 	1
Constant UInt ALIGNV_BOTTOM	 = 	2
Constant UInt ALIGNV_JUSTIFY	 = 	3
Constant UInt ALIGNV_DISTRIBUTED	 = 	4

//borderstyle
Constant Long  BORDERSTYLE_NONE				 = 	0
Constant Long  BORDERSTYLE_THIN				 = 	1
Constant Long  BORDERSTYLE_MEDIUM			 = 	2
Constant Long  BORDERSTYLE_DASHED			 = 	3
Constant Long  BORDERSTYLE_DOTTED			 = 	4
Constant Long  BORDERSTYLE_THICK				 = 	5
Constant Long  BORDERSTYLE_DOUBLE			 = 	6
Constant Long  BORDERSTYLE_HAIR				 = 	7
Constant Long  BORDERSTYLE_MEDIUMDASHED	 = 	8
Constant Long  BORDERSTYLE_DASHDOT			 = 	9
Constant Long  BORDERSTYLE_MEDIUMDASHDOT	 = 	10
Constant Long  BORDERSTYLE_DASHDOTDOT		 = 	11
Constant Long  BORDERSTYLE_MEDIUMDASHDOTDOT		 = 	12
Constant Long  BORDERSTYLE_SLANTDASHDOT			 = 	13

//numformat
Constant Long NUMFORMAT_GENERAL	 = 	0 //general format	
Constant Long NUMFORMAT_NUMBER	 = 	1 //general number	1000
Constant Long NUMFORMAT_NUMBER_D2	 = 	2 //number with decimal point	1000.00
Constant Long NUMFORMAT_NUMBER_SEP	 = 	3 //number with thousands separator	100,000
Constant Long NUMFORMAT_NUMBER_SEP_D2	 = 	4 //number with decimal point and thousands separator	100,000.00
Constant Long NUMFORMAT_CURRENCY_NEGBRA	 = 	5 //monetary value, negative in brackets	(1000$)
Constant Long NUMFORMAT_CURRENCY_NEGBRARED	 = 	6 //monetary value, negative is red in brackets	(1000$)
Constant Long NUMFORMAT_CURRENCY_D2_NEGBRA	 = 	7 //monetary value with decimal point, negative in brackets	($1000.00)
Constant Long NUMFORMAT_CURRENCY_D2_NEGBRARED	 = 	8 //monetary value with decimal point, negative is red in brackets	($1000.00)
Constant Long NUMFORMAT_PERCENT	 = 	9 //percent value, multiply the cell value by 100	75%
Constant Long NUMFORMAT_PERCENT_D2	 = 	10 //percent value with decimal point, multiply the cell value by 100	75.00%
Constant Long NUMFORMAT_SCIENTIFIC_D2	 = 	11 //scientific value with E character and decimal point	10.00E+1
Constant Long NUMFORMAT_FRACTION_ONEDIG	 = 	12 //fraction value, one digit	10 1/2
Constant Long NUMFORMAT_FRACTION_TWODIG	 = 	13 //fraction value, two digits	10 23/95
Constant Long NUMFORMAT_DATE						 = 	14 //date value, depends on OS settings	3/11/2009
Constant Long NUMFORMAT_CUSTOM_D_MON_YY		 = 	15 //custom date value	11-Mar-09
Constant Long NUMFORMAT_CUSTOM_D_MON			 = 	16 //custom date value	11-Mar
Constant Long NUMFORMAT_CUSTOM_MON_YY			 = 	17 //custom date value	Mar-09
Constant Long NUMFORMAT_CUSTOM_HMM_AM			 = 	18 //custom date value	8:30 AM
Constant Long NUMFORMAT_CUSTOM_HMMSS_AM		 = 	19 //custom date value	8:30:00 AM
Constant Long NUMFORMAT_CUSTOM_HMM				 = 	20 //custom date value	8:30
Constant Long NUMFORMAT_CUSTOM_HMMSS			 = 	21 //custom date value	8:30:00
Constant Long NUMFORMAT_CUSTOM_MDYYYY_HMM			 = 	22 //custom datetime value	3/11/2009 8:30
Constant Long NUMFORMAT_NUMBER_SEP_NEGBRA			 = 	37 //number with thousands separator, negative in brackets	(4,000)
Constant Long NUMFORMAT_NUMBER_SEP_NEGBRARED		 = 	38 //number with thousands separator, negative is red in brackets	(4,000)
Constant Long NUMFORMAT_NUMBER_D2_SEP_NEGBRA		 = 	39 //number with thousands separator and decimal point, negative in brackets	(4,000.00)
Constant Long NUMFORMAT_NUMBER_D2_SEP_NEGBRARED	 = 	40 //number with thousands separator and decimal point, negative is red in brackets	(4,000.00)
Constant Long NUMFORMAT_ACCOUNT		 = 	41 //account value	5,000
Constant Long NUMFORMAT_ACCOUNTCUR		 = 	42 //account value with currency symbol	$	5,000
Constant Long NUMFORMAT_ACCOUNT_D2		 = 	43 //account value with decimal point	5,000.00
Constant Long NUMFORMAT_ACCOUNT_D2_CUR		 = 	44 //account value with currency symbol and decimal point	$	5,000.00
Constant Long NUMFORMAT_CUSTOM_MMSS		 = 	45 //custom time value	30:55
Constant Long NUMFORMAT_CUSTOM_H0MMSS		 = 	46 //custom time value	20:30:55
Constant Long NUMFORMAT_CUSTOM_MMSS0		 = 	47 //custom time value	30:55.0
Constant Long NUMFORMAT_CUSTOM_000P0E_PLUS0		 = 	48 //custom value	15.2E+3
Constant Long NUMFORMAT_TEXT		 = 	49 //text value	any text

//color
Constant Long COLOR_BLACK = 8
Constant Long COLOR_WHITE =  9
Constant Long COLOR_RED =  10
Constant Long COLOR_BRIGHTGREEN = 	11
Constant Long COLOR_BLUE =  	12
Constant Long COLOR_YELLOW =  13
Constant Long COLOR_PINK = 14
Constant Long COLOR_TURQUOISE = 	15
Constant Long COLOR_DARKRED =     16
Constant Long COLOR_GREEN =  	17
Constant Long COLOR_DARKBLUE =  18
Constant Long COLOR_DARKYELLOW =  19
Constant Long COLOR_VIOLET =  	20
Constant Long COLOR_TEAL =  	21
Constant Long COLOR_GRAY25 =  	22
Constant Long COLOR_GRAY50 =  	23
Constant Long COLOR_PERIWINKLE_CF = 	24
Constant Long COLOR_PLUM_CF =  		25
Constant Long COLOR_IVORY_CF =  	26
Constant Long COLOR_LIGHTTURQUOISE_CF =  27
Constant Long COLOR_DARKPURPLE_CF =  	28
Constant Long COLOR_CORAL_CF =  	29
Constant Long COLOR_OCEANBLUE_CF =  	30
Constant Long COLOR_ICEBLUE_CF =          	31
Constant Long COLOR_DARKBLUE_CL =  	32
Constant Long COLOR_PINK_CL =  		33
Constant Long COLOR_YELLOW_CL =  	34
Constant Long COLOR_TURQUOISE_CL =  	35
Constant Long COLOR_VIOLET_CL =  	36
Constant Long COLOR_DARKRED_CL =  	37
Constant Long COLOR_TEAL_CL =             38
Constant Long COLOR_BLUE_CL =  		39
Constant Long COLOR_SKYBLUE =  		40
Constant Long COLOR_LIGHTTURQUOISE =  	41
Constant Long COLOR_LIGHTGREEN =  	42
Constant Long COLOR_LIGHTYELLOW =  	43
Constant Long COLOR_PALEBLUE =  	44
Constant Long COLOR_ROSE =  		45
Constant Long COLOR_LAVENDER =           46
Constant Long COLOR_TAN =  		47
Constant Long COLOR_LIGHTBLUE =  	48
Constant Long COLOR_AQUA =  		49
Constant Long COLOR_LIME =  		50
Constant Long COLOR_GOLD =  		51
Constant Long COLOR_LIGHTORANGE =  52
Constant Long COLOR_ORANGE =  53
Constant Long COLOR_BLUEGRAY =  54
Constant Long COLOR_GRAY40 =       55
Constant Long COLOR_DARKTEAL =  56
Constant Long COLOR_SEAGREEN =  57
Constant Long COLOR_DARKGREEN =  58
Constant Long COLOR_OLIVEGREEN =  59
Constant Long COLOR_BROWN =  60
Constant Long COLOR_PLUM =  61
Constant Long COLOR_INDIGO =  62
Constant Long COLOR_GRAY80 =     63
Constant Long COLOR_DEFAULT_FOREGROUND = 64 // 0x0040 =  64
Constant Long COLOR_DEFAULT_BACKGROUND = 65 //0x0041 =  65
Constant Long COLOR_TOOLTIP = 81 //0x0051 
Constant Long COLOR_NONE = 	127 //0x7F =  
Constant Long COLOR_AUTO = 	32767 //0x7FFF


Constant Long UNDERLINE_NONE	 = 	0
Constant Long UNDERLINE_SINGLE	 = 	1
Constant Long UNDERLINE_DOUBLE	 = 	2
Constant Long UNDERLINE_SINGLEACC = 33 //0x21, 
Constant Long UNDERLINE_DOUBLEACC = 34 //0x22};


Constant Long PAPER_DEFAULT = 0
Constant Long PAPER_LETTER = 1
Constant Long PAPER_LETTERSMALL = 2
Constant Long PAPER_TABLOID = 3
Constant Long PAPER_LEDGER = 4
Constant Long PAPER_LEGAL = 5
Constant Long PAPER_STATEMENT = 6
Constant Long PAPER_EXECUTIVE = 7
Constant Long PAPER_A3 = 8
Constant Long PAPER_A4 = 9
Constant Long PAPER_A4SMALL = 10
Constant Long PAPER_A5 = 11
Constant Long PAPER_B4 = 12
Constant Long PAPER_B5 = 13
Constant Long PAPER_FOLIO = 14
Constant Long PAPER_QUATRO = 15
Constant Long PAPER_10x14 = 16
Constant Long PAPER_10x17 = 17
Constant Long PAPER_NOTE = 18
Constant Long PAPER_ENVELOPE_9 = 19
Constant Long PAPER_ENVELOPE_10 = 20
Constant Long PAPER_ENVELOPE_11 = 21
Constant Long PAPER_ENVELOPE_12 = 22
Constant Long PAPER_ENVELOPE_14 = 23
Constant Long PAPER_C_SIZE = 24
Constant Long PAPER_D_SIZE = 25
Constant Long PAPER_E_SIZE = 26
Constant Long PAPER_ENVELOPE_DL = 27
Constant Long PAPER_ENVELOPE_C5 = 28
Constant Long PAPER_ENVELOPE_C3 = 29
Constant Long PAPER_ENVELOPE_C4 = 30
Constant Long PAPER_ENVELOPE_C6 = 31
Constant Long PAPER_ENVELOPE_C65 = 32
Constant Long PAPER_ENVELOPE_B4 = 33
Constant Long PAPER_ENVELOPE_B5 = 34
Constant Long PAPER_ENVELOPE_B6 = 35
Constant Long PAPER_ENVELOPE = 36
Constant Long PAPER_ENVELOPE_MONARCH = 37
Constant Long PAPER_US_ENVELOPE = 38
Constant Long PAPER_FANFOLD = 39
Constant Long PAPER_GERMAN_STD_FANFOLD = 40
Constant Long PAPER_GERMAN_LEGAL_FANFOLD = 41
Constant Long PAPER_B4_ISO = 42
Constant Long PAPER_JAPANESE_POSTCARD = 43
Constant Long PAPER_9x11 = 44
Constant Long PAPER_10x11 = 45
Constant Long PAPER_15x11 = 46
Constant Long PAPER_ENVELOPE_INVITE = 47
Constant Long PAPER_US_LETTER_EXTRA  = 50
Constant Long PAPER_US_LEGAL_EXTRA = 51
Constant Long PAPER_US_TABLOID_EXTRA = 52
Constant Long PAPER_A4_EXTRA = 53
Constant Long PAPER_LETTER_TRANSVERSE = 54
Constant Long PAPER_A4_TRANSVERSE = 55
Constant Long PAPER_LETTER_EXTRA_TRANSVERSE = 56
Constant Long PAPER_SUPERA = 57
Constant Long PAPER_SUPERB = 58
Constant Long PAPER_US_LETTER_PLUS = 59
Constant Long PAPER_A4_PLUS = 60
Constant Long PAPER_A5_TRANSVERSE = 61
Constant Long PAPER_B5_TRANSVERSE = 62
Constant Long PAPER_A3_EXTRA = 63
Constant Long PAPER_A5_EXTRA = 64
Constant Long PAPER_B5_EXTRA = 65
Constant Long PAPER_A2 = 66
Constant Long PAPER_A3_TRANSVERSE = 67
Constant Long PAPER_A3_EXTRA_TRANSVERSE = 68
Constant Long PAPER_JAPANESE_DOUBLE_POSTCARD = 69
Constant Long PAPER_A6 = 70
Constant Long PAPER_JAPANESE_ENVELOPE_KAKU2 = 71
Constant Long PAPER_JAPANESE_ENVELOPE_KAKU3 = 72
Constant Long PAPER_JAPANESE_ENVELOPE_CHOU3 = 73
Constant Long PAPER_JAPANESE_ENVELOPE_CHOU4 = 74
Constant Long PAPER_LETTER_ROTATED = 75
Constant Long PAPER_A3_ROTATED = 76
Constant Long PAPER_A4_ROTATED = 77
Constant Long PAPER_A5_ROTATED = 78
Constant Long PAPER_B4_ROTATED = 79
Constant Long PAPER_B5_ROTATED = 80
Constant Long PAPER_JAPANESE_POSTCARD_ROTATED = 81
Constant Long PAPER_DOUBLE_JAPANESE_POSTCARD_ROTATED = 82
Constant Long PAPER_A6_ROTATED = 83
Constant Long PAPER_JAPANESE_ENVELOPE_KAKU2_ROTATED = 84
Constant Long PAPER_JAPANESE_ENVELOPE_KAKU3_ROTATED = 85
Constant Long PAPER_JAPANESE_ENVELOPE_CHOU3_ROTATED = 86
Constant Long PAPER_JAPANESE_ENVELOPE_CHOU4_ROTATED = 87
Constant Long PAPER_B6 = 88
Constant Long PAPER_B6_ROTATED = 89
Constant Long PAPER_12x11 = 90
Constant Long PAPER_JAPANESE_ENVELOPE_YOU4 = 91
Constant Long PAPER_JAPANESE_ENVELOPE_YOU4_ROTATED = 92
Constant Long PAPER_PRC16K = 93
Constant Long PAPER_PRC32K = 94
Constant Long PAPER_PRC32K_BIG = 95
Constant Long PAPER_PRC_ENVELOPE1 = 96
Constant Long PAPER_PRC_ENVELOPE2 = 97
Constant Long PAPER_PRC_ENVELOPE3 = 98
Constant Long PAPER_PRC_ENVELOPE4 = 99
Constant Long PAPER_PRC_ENVELOPE5 = 100
Constant Long PAPER_PRC_ENVELOPE6 = 101
Constant Long PAPER_PRC_ENVELOPE7 = 102
Constant Long PAPER_PRC_ENVELOPE8 = 103
Constant Long PAPER_PRC_ENVELOPE9 = 104
Constant Long PAPER_PRC_ENVELOPE10 = 105
Constant Long PAPER_PRC16K_ROTATED = 106
Constant Long PAPER_PRC32K_ROTATED = 107
Constant Long PAPER_PRC32KBIG_ROTATED = 108
Constant Long PAPER_PRC_ENVELOPE1_ROTATED = 109
Constant Long PAPER_PRC_ENVELOPE2_ROTATED = 110
Constant Long PAPER_PRC_ENVELOPE3_ROTATED = 111
Constant Long PAPER_PRC_ENVELOPE4_ROTATED = 112
Constant Long PAPER_PRC_ENVELOPE5_ROTATED = 113
Constant Long PAPER_PRC_ENVELOPE6_ROTATED = 114
Constant Long PAPER_PRC_ENVELOPE7_ROTATED = 115
Constant Long PAPER_PRC_ENVELOPE8_ROTATED = 116
Constant Long PAPER_PRC_ENVELOPE9_ROTATED = 117
Constant Long PAPER_PRC_ENVELOPE10_ROTATED = 118


Constant Long FILLPATTERN_NONE				 = 	0
Constant Long FILLPATTERN_SOLID				 = 	1
Constant Long FILLPATTERN_GRAY50				 = 	2
Constant Long FILLPATTERN_GRAY75				 = 	3
Constant Long FILLPATTERN_GRAY25				 = 	4
Constant Long FILLPATTERN_HORSTRIPE			 = 	5
Constant Long FILLPATTERN_VERSTRIPE			 = 	6
Constant Long FILLPATTERN_REVDIAGSTRIPE	 = 	7
Constant Long FILLPATTERN_DIAGSTRIPE		 = 	8
Constant Long FILLPATTERN_DIAGCROSSHATCH	 = 	9
Constant Long FILLPATTERN_THICKDIAGCROSSHATCH	 = 	10
Constant Long FILLPATTERN_THINHORSTRIPE			 = 	11
Constant Long FILLPATTERN_THINVERSTRIPE			 = 	12
Constant Long FILLPATTERN_THINREVDIAGSTRIPE		 = 	13
Constant Long FILLPATTERN_THINDIAGSTRIPE			 = 	14
Constant Long FILLPATTERN_THINHORCROSSHATCH		 = 	15
Constant Long FILLPATTERN_THINDIAGCROSSHATCH		 = 	16
Constant Long FILLPATTERN_GRAY12P5					 = 	17
Constant Long FILLPATTERN_GRAY6P25					 = 	18


end variables
on cxls.create
call super::create
TriggerEvent( this, "constructor" )
end on

on cxls.destroy
TriggerEvent( this, "destructor" )
call super::destroy
end on

