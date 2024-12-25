let
// 	ПОДКЛЮЧЕНИЕ БИБЛИОТЕК
	F      = let path = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\2. PQ запросы\F.pq"      in Expression.Evaluate(Text.FromBinary(File.Contents(path)), #shared),
	F_Form = let path = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\2. PQ запросы\F_Form.pq" in Expression.Evaluate(Text.FromBinary(File.Contents(path)), #shared),
D = [ 
//	ВАЖНО
	ОПИСАНИЕ = [
		r1  = " определеится и именованием и поназывать едионообразно колонки, + на русском и псевдоним на En"
	],
	README = [
		to  =  // 	Используемые функции функция какая?
			{" 	
				F[Column.RemoveNull]
				F[GetQueryValues]
				F[LoadCSV]
				F[Reorder.AllColumns]
				F[Row.RemoveNullColumns]
				F[Table.AddedRows]
				F[Table.ClearAll]
				F[Table.ReplaceErr]
				F[Table.SkipUp]
				F[Table.TransformContent]
				F[Table.TypedField]
				F[Text.Clear]
				F[Text.CompareText]
				F[Text.DeletBracket]
				F[Text.MultiReplace]
				F[Text.RemoveIntoBracket]
				F[Text.RemoveMulti]
				F[Text.ReplaceBetweenDelimiters]
				F[Text.ReplaceMulti]
				F[Type.AllColumnText]
				F[fxAddMultiColumns]
				F[fxJoinSpr]
				F[fxSizeFile]
			"},
		r1  = " определеится и именованием и поназывать едионообразно колонки, + на русском и псевдоним на En",
		r2  = " добавить ID где надо и сделатиь его числовым наверное, чтобы по названию было видно тип столюца",
		r3  = " функция убрать 0 на null",
		r4  = " если схемы отчета езе нее в tableSherma то поумолчанию тогда",
		r5  = " СДЕЛАТЬ [ Mart = , HOM] и т.д. фунцкию",
		r6  = " tableAddedRows2 = функцию сделать",
		r7  = " именование файлов и полей в датасетах CSV и XLS непонятное, одинаковое дб, влияет на что то  разобраться",
		r8  = " наличие вычисляемых таблиц внутри записи замедляет ее загрузку в проект",
		r9  = " документировать основные функции LoadConRep",
		r10 = " сделать Get фугкция параметров по полю, а не значени.",
		r11 = " из F_Form в самом саом конце можно убрать выгрузку функций, которые тормозят весь запрос",
		r12 = " в шаге отчетов Load колонок может быть не равно. ОБработать. Когда выбираем из документа колонки",
		r13 = " //! Удалить при вставке Ошибка в проекте из-за fxDataSet = в шаге prm = prm_DataSet если prm_DataSet с _ в конце",
		R   = [ // порядок работы с новым документом - если код находится во внешних файлах
			/* 	
				------------------ Перенос кода ---------------------------
				ВАРИАНТ 1 - запросы в файлах на диске/интернете
				1. Грузим загрузчик библиотек 	 	-	Копируем fxParam в новый файл. Называем запрос Load
				2. Грузим основной пакет функций 	-	= Load("PackGlobal"). Называем запрос F
				3. Грузим все остальные пакеты 		-  	= Load("PackForm") - функции отчетности. Называем запрос fxPackForm
				
				ВАРИАНТ 2 - запросы в самом файле Power BI или Excel
				1. Тоже самое только копируем все в файл
			
				------------------ Дата отчета ----------------------------
				ВАЖНО!!!  
					1. ДАТУ ОТЧЕТА БЕРЕМ ИЗ НАЗВАНИЯ ФАЙЛА - если отчетов мало, прислали например
					2. ДАТУ ОТЧЕТА БЕРЕМ ИЗ ФАЙЛОВ - если отчетов много много и они одинаковой структуры!

				Маска начала названия файла: если вручную называть файлы, если их не много
					КОД_МЕСЯЦ_ГОД => 
						КОДОТЧЕТА 	- 2 знака (лучше 3). Но это не важно, резделять порзиции по разделителю_
						МЕСЯЦ	 	- 2 знака
						ГОД			- 2 знака
					
					В ОПС так = 011216 => 01 код 12 месяц 16 год
					Сделать так, если мне присылают отчеты: 06R_12_24
				*/
			]
		],

//	ССЫЛКИ
	link = [
		link  = // хорошие скрины по финансовым отчетам
			"https://www.agile-insights.com.au/power-bi-solutions/financial-analytics-dashboards/",
		link2 = // отличный учебник с baguzin, тут - система типов
			"https://baguzin.ru/wp/yazyk-m-power-query-sistema-tipov-osnovy/", 
		link3 =3, // хорошие скрины по финансовым отчетам
		link4 =4 // хорошие скрины по финансовым отчетам
		],
// 	ПАРАМЕТРЫ
	prm_DataSet  = "2 Homel" /* параметр PowerBI: выбор датасета (для PBI, надо поставить "_" в имени параметра) */
		meta // ? окончание можно не писать, неомер берется сплитом
			[	IsParameterQuery 	= true,
				List 				= {"1 Mart", "2 Homel", "3 BKS", "4 Дрогичин"},
				DefaultValue 		= "2 Homel",
				Type 				= "Text",
				IsParameterQueryRequired = true	]
		,

	prm_LoadVar  = "csv" 	 /* параметр PowerBI: выбор вида импорта (для PBI, надо поставить "_" в имени параметра) */
		 meta
			[	IsParameterQuery	= true,
				List				= {"csv", "xls"},
				DefaultValue		= "csv",
				Type				= "Text",
				IsParameterQueryRequired = true	]
		,
	
	prm_LoadMode = 1 		 /* параметр PowerBI: вариант загрузки xls (0(full)- полные данные, 1(short) - один отчет, за 1 год, по 1 предприятию (если их много) */
		meta [IsParameterQuery=true, List={0, 1}, DefaultValue=0, Type="Number", IsParameterQueryRequired=true]
		/*	
			Варианты по xls:
				0	-	грузим в модель все данные
				1	-	грузим в модель только одну строку: 1 отчет, 1 предприятие, 1 дата.
						Table = данные
		*/
		,
	
	Parameters   = // параметры проекта. В модели надо держать запрос в загрузке, для вывода параметров на дашборд!
		/*
			Описание:
			Связь:
				Запрос ParametersMeta = Value.Metadata( F_Form[Parameters] ) - для чего, связан ли с другими отчетами?!
			Непонятки 
				как параметр loadVar = Record.FieldOrDefault( [1 = "csv", 2 = "xlsx"], Text.From(1), null),
							 loadVar = F[GetQueryValues]("prm_LoadVar"),
							 loadVar = prm_LoadVar,	// loadVar = F[GetQueryValues]("prm_LoadVar"),
			Улучшение:
				1	Путь для DataShema - отдельный параметр, не выгружаемый 
				2  	StartDate и EndDate сделать через функцию ибрать значение из колонки, например:
						StartDate = Date.StartOfYear(Date.From(List.Min(from[Date])))
						EndDate   = Date.EndOfYear(Date.From(List.Max(from[Date])))
		*/

		let	from = [
				Mart  = [ 
					Company      = "Фабрика 8 Марта",
					DataSetName  = "Mart",
					DataSetNomer = 1,
					loadVar      = prm_LoadVar,	
					PERIOD_RANGE = "Q",
			
					path         = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности",
					pathXLSX     = path & "\1. Отчетность\Фабрика",
					pathCSV      = path & "\4. CSV выгрузка\Фабрика",
					
					FilesAll 	 = F_Form[LoadFolder](DataSetName,3),
					FilesXLSX 	 = F_Form[Connect](DataSetName)(),
					FilesCSV	 = F_Form[LoadFolderCSV](pathCSV),
					
					lbl_Period   = "c 2021 по 2022 год",
					StartDate    = "01.01.2021",	// Date.FromText("01.01.2021"),
					EndDate      = "31.12.2022"		// Date.FromText("01.01.2022"),
					],

				Homel = [
					Company      = "Гомельское ОПО",
					DataSetName  = "Homel",
					DataSetNomer = 2,
					loadVar      = prm_LoadVar,
					
					path         = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности",
					pathXLSX     = path & "\1. Отчетность\Гомель",
					// pathXLSX     = path & "\0. Исходники\Исходники ГомельскоеОПО (2015-2023)",
					pathCSV      = path & "\4. CSV выгрузка\Гомель",
					
					FilesAll 	 = F_Form[LoadFolder](DataSetName,3),
					FilesXLSX 	 = F_Form[Connect](DataSetName)(),
					FilesCSV	 = F_Form[LoadFolderCSV](pathCSV),
					
					PERIOD_RANGE = "Y",
					lbl_Period   = "c 2016 по 2021 год",
					StartDate    = "01.01.2016",
					EndDate      = "31.12.2023"
					],

				BKS   = [
					/* 
						Особенности датасета:
							1 Один файл xls 
					*/
					Company      = "Белкоопсоюз",
					DataSetName  = "BKS",
					DataSetNomer = 3,
					loadVar      = prm_LoadVar,
					
					path		 = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности",
					pathXLSX     = path & "\1. Отчетность\БКС",
					pathCSV      = path & "\4. CSV выгрузка\Белкоопсоюз",
					
					FilesAll 	 = F_Form[LoadFolder](DataSetName,3),
					FilesXLSX 	 = F_Form[Connect](DataSetName)(),
					FilesCSV	 = F_Form[LoadFolderCSV](pathCSV),
					
					PERIOD_RANGE = "Y",
					lbl_Period   = "c 2016 по 2022 год",
					StartDate    = "01.01.2016",
					EndDate      = "31.12.2022"
					],

				Drog  = [
					Company      = "Догичинское райпо",
					DataSetName  = "Drog",
					DataSetNomer = 4,
					loadVar      = prm_LoadVar,
					
					path         = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности",
					pathXLSX     = path & "\1. Отчетность\Дрогичин",
					pathCSV      = path & "\4. CSV выгрузка\Дрогичин",
					
					FilesAll 	 = F_Form[LoadFolder](DataSetName,3),
					FilesXLSX 	 = F_Form[Connect](DataSetName)(),
					FilesCSV	 = F_Form[LoadFolderCSV](pathCSV),
					
					PERIOD_RANGE = "Y",
					
					lbl_Period   = "c 2022 по 2023 год",
					StartDate    = "01.01.2022", 
					EndDate      = "31.12.2023"
					],
				
				// 	варианты параметров для пробы
					test  = [ 
						Year         		= {2022..2023},
						IsReportViewGroup 	= true meta [Type="Logical"],
						sec 				= #sections[Section1]
					]	
			],

			// 	функция выбора датасета по номеру
				fxParam = (DataSet as number) =>
					Record.FieldOrDefault([   
							1 = from[Mart],
							2 = from[Homel],
							3 = from[BKS],
							4 = from[Drog],
							0 = 0
						], Text.From(DataSet), null 
					),

			// 	функция определения номера датасета
				fxDataSet = let 
					// prm = F[GetQueryValues]("prm_DataSet"),		// значение параметра
					prm = prm_DataSet,	// из за этого параметра возможна ошибка //!
					res = Number.From(Text.Split(prm," "){0}) 	// берем часть текста до первого пробела
				in 	res,

			//  сохраняем метки на шаг from (список record), чтобы потом использовать шаг, как окончание запроса
				/* Пример: = Value.Metadata(Parameters)[Homel][pathCSV] - возвращает путь к папке */
				to = fxParam(fxDataSet) meta [
						Homel = from[Homel],
						Mart  = from[Mart],
						BKS   = from[BKS],
						Drog  = from[Drog]
					]
		in	to,

// 	ПАРАМЕТРЫ 2 (КОПИРУЕТСЯ В ПРОЕКТ)
	Calendar     = "" // календарь
		// "# КАЛЕНДАРЬ" "_CALENDAR"
		// = F_Calendar[fxCalendarKill](Date.From(Parameters[StartDate]), Date.From(Parameters[EndDate]), null, true)  // #date(2016, 12, 1) - #date(2021, 12, 31)
		,
	DS 			 = // #1 датасеты всех проектов 
		/* 
			Описание:
				! Для чего нужен и кто использует? Добавить и по другим функциям ответ на эти вопросы
				! Не понятное использование

				Таблица из 3-x колонок, в зависимости от параметра xls или csv
				------------------------------
				Index	Name	Vid		Value	
				------------------------------
				1		Mart	xls		Table	
				2		Homel	xls		Table	
				3		BKS		xls		Table	
				4		Drog	xls		Table	
				------------------------------
			
			Пример: 
				F_Form[DS]( )	- все датасеты, и xls и csv
				F_Form[DS](1)	- датасеты xls и csv
				F_Form[DS](2)	- датасеты csv
			
			А грузится то что, когда и то и то?
		*/
		(optional n as nullable number) =>
			let	
				n 	 = if prm_LoadVar = "xls" then 1 else 2,
				data = 
					#table(
						type table[Index = number, Name = text, xls = table, csv = table],
						{
							{1, "Mart",  DataSet_XLS_Mart,  DataSet_CSV_Mart},
							{2, "Homel", DataSet_XLS_Homel, DataSet_CSV_Homel},
							{3, "BKS",   DataSet_XLS_BKS,   DataSet_CSV_BKS},
							{4, "Drog",  DataSet_XLS_Drog,  DataSet_CSV_Drog}
						}
					),
				unp = 	Table.UnpivotOtherColumns(data, {"Index", "Name"}, "Vid", "Value"),
				sel = 	Table.SelectRows(unp, each 
							if 		n = 1	then [Vid] = "xls"
							else if n = 2	then [Vid] = "csv"
							else true
						)
			in	
				sel,
	DS_CSV 		 = // Record с проектами. Пока не использую 
		[	
			DataSet_CSV_Mart  = DataSet_CSV_Mart, 
			DataSet_CSV_Homel = DataSet_CSV_Homel, 
			DataSet_CSV_BKS   = DataSet_CSV_BKS,
			DataSet_CSV_Drog  = DataSet_CSV_Drog
		],
	DS_XLS 		 = // Record с проектами. Пока не использую 
		[
			DataSet_XLS_Mart  = DataSet_XLS_Mart, 
			DataSet_XLS_Homel = DataSet_XLS_Homel, 
			DataSet_XLS_BKS   = DataSet_XLS_BKS,
			DataSet_XLS_Drog  = DataSet_XLS_Drog
		],

	DataShema    = // #2 схема таблиц. На входе CSV, на выходе Record: Name (вид отчета) + Value (таблица со схемой)
		/*  
			Описание:
				схема таблиц всех отчетов с типизацией столбцов полученная из файла CSV
				нужна для типизации столбцов при импорте готовых отчетов из CSV
		*/
		let
			from = F[LoadCSV]("G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\4. CSV выгрузка\tableSchema.csv"),
			typ  = Table.TransformColumnTypes(from,{{"Position", Int64.Type}}),
			srt  = Table.Sort(typ,{"Position", Order.Ascending}),
			grp  = Table.Group(srt, {"Name"}, {{"Value", each _}}),
			add  = Table.AddColumn(grp, "СountRows", each Table.RowCount([Value])),
			rec  = Record.FromTable(add)
		in
			rec,
		
		/*  Вариант с загрузкий схем таблиц из имеющихся в проекте таблиц
			let
					// from = Table.SelectRows(Record.ToTable(#shared), each Text.StartsWith([Name], "DataSet_XLS")),
					tbl  = DS, // DataSet

					
					exp = [ // Разворачивание и преобразование столбца "Value", чистка ошибок
						nms = {"Name", "Value"},
						exp = Table.ExpandTableColumn(tbl, "Value", nms, {"Атрибут", "Значение"}),
						trn = Table.TransformColumns(exp, {"Значение", each Table.FirstN(_, 1)}),
						cln = Table.RemoveRowsWithErrors(trn, {"Значение"})
					][cln],

				
					grp = [ //  Группировка по "Атрибут" и применение функции
						f=(rows)=> F_Form[form.Shema](rows[Значение]{0}),
						to = Table.Group(exp, {"Атрибут"}, {{"Количество", f}})
					][to],

					ren = [	//  Переименование столбцов перед преобразованием в запись
						a= Table.RenameColumns(grp, {{"Атрибут", "Name"}, {"Количество", "Value"}}),
						 = Record.FromTable(ren)
					][rec]
			in	ren
		*/

	DataShemaExport = ""// Проверить таблица со схемами всех отчетов для формированоя SQL запроса для формирования базы данных
		/*
			Описать когда применять!
		*/

		// 	let
		// 		tbl = Record.ToTable([
		// 			Mart  = DataSet_XLS_Mart, 
		// 			Homel = DataSet_XLS_Homel, 
		// 			BKS   = DataSet_XLS_BKS,
		// 			Drog  = DataSet_XLS_Drog
		// 			]),

		// 	//  Разворачивание и преобразование столбца "Value", чистка ошибок
		// 		exp = [
		// 			nms = {"Name", "Value"},
		// 			exp = Table.ExpandTableColumn(tbl, "Value", nms, {"Атрибут", "Значение"}),
		// 			trn = Table.TransformColumns(exp, {"Значение", each Table.FirstN(_, 1)}),
		// 			cln = Table.RemoveRowsWithErrors(trn, {"Значение"})
		// 		][cln],

		// 	//  Группировка по "Атрибут" и применение функции
		// 		grp = [
		// 			f=(rows)=> F_Form[form.Shema](rows[Значение]{0}),
		// 			to = Table.Group(exp, {"Атрибут"}, {{"Количество", f}})
		// 		][to],

		// 	//  Переименование столбцов перед преобразованием в запись
		// 		ren = Table.RenameColumns(grp, {{"Атрибут", "Name"}, {"Количество", "Value"}}),
		// 		nms = List.Distinct(List.Combine(List.Transform(Table.Column(ren, "Value"), each Table.ColumnNames(_)))),
		// 		exp2 = Table.ExpandTableColumn(ren, "Value", nms, nms)
		// in
		// 	exp2

		// let
		// 	from = Record.ToTable(DataShema),
		// 	nms  = List.Distinct(List.Combine(List.Transform(Table.Column(from, "Value"), each Table.ColumnNames(_)))),
		// 	exp  = Table.ExpandTableColumn(from, "Value", nms, nms)
		// in
		// 	exp
		,
	
// 	ФУНКЦИИ
	//? НОВЫЕ ВЕРСИИ ФУНКЦИЙ!
	LoadFolder_ = // #1 подключаемся к папке, получаем список файлов XLSX с подпапками, чистим  (ФОРМИРУЕМ ИСТОЧНИК  )
		/* 
			Описание:
				1. подключаемся к папке с файлами по имени датасета или по полному пути
				2. выбираем номер шага загрузки. По умолчанию выгружаем последний шаг
				3. чистим папку от временных и скрытых файлов, оставляем файлы .xlsx, преобразуем бинарники
				4. в книгах excel оставляем только листы (Sheet)
			
			Аргументы:
				optional options as record
					Например [ДатаCетИмя = "Homel", НомерШага = null]
			
			Пример:	
				= F[LoadFolder_]
					([ДатаCетИмя = "Homel"])					// возвращается последний шаг
					([ДатаCетИмя = "Homel", НомерШага = null])	// возвращается последний шаг
					([ДатаCетИмя = "Homel", НомерШага = 3])		// возвращается третий шаг
					([ДатаCетИмя = "Homel", НомерШага = 123])	// возвращается надпись об ошибке
			
			Улучшения:
				1 //! Folder.Contents – собираем данные из папки, но выборочно
				1. Если датасет не выбран, то меняем например на Homel 
				2. Сделать макеты таблиц всех отчетов для замены, если ничего нет или не выбрано.
				3. Макеты таблиц с данными по Условной организации. Ее надо создать
				4. Это могут быть данные из одного файла XlSX с данными всех отчетов
				5. Данные разных организаций можно совместить, чтобы заполнить базу и визуализации
			*/
		[
			func   = (optional options as record) => [ 	
				//	0. формируем список аргументов 
					DataSetName2  = 
						let value = Record.FieldOrDefault(options, "ДатаCетИмя", null)
						in  if List.Contains({null, 0, ""}, value) then null else Text.From(value),
					
					DataSetName  = let value = Record.FieldOrDefault(options, "ДатаCетИмя", null) in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
					STEP         = let value = Record.FieldOrDefault(options, "НомерШага", null)  in if List.Contains({null, 0, ""}, value) then null else Number.From(value),

				//	1. формируем path к папке с файлами XLSX //? датасет имя 
					path =  Record.FieldOrDefault(
								[ 
									Homel = Value.Metadata(Parameters) [Homel] [pathXLSX],
									Mart  = Value.Metadata(Parameters) [Mart]  [pathXLSX],
									BKS   = Value.Metadata(Parameters) [BKS]   [pathXLSX],
									Drog  = Value.Metadata(Parameters) [Drog]  [pathXLSX]
								], 
								DataSetName, 
								if  Text.Contains(DataSetName,"\")	      // если имя содержит "\" (как часть пути к папке)
									then DataSetName					  // то это путь к папке (для датасетов не из параметров)
									else "Error. Ошибка в имени датасета" // то ошибка 	
							),

				// 	2. подключаемся к папке + фильтруем колонки + немного меняем формат колонок
					from  = Folder.Files(path),
					
					nms	  = {"Name","Extension", "Attributes", "Date created", "Content"} & {"Folder Path"},
					Sel	  = Table.SelectColumns(from, nms),
					Trans = Table.TransformColumns(Sel, {
						{"Date created", (x)=> DateTime.Date(x), type date},
						{"Folder Path",  (x)=> Text.AfterDelimiter(x, "\", {2, RelativePosition.FromEnd}), type text}
					}),

				// 	3. фильтруем строки
					FilteredRows 	= Table.SelectRows(Trans, each 
						(Text.Start([Extension], 5) = ".xlsx" and 	// оставляем только .xlsx
						[Attributes][Hidden] = false ) and			// убираем скрытые
						not Text.StartsWith([Name], "~$")			// убираем временные
					),

				//  4. преобразуем бинарники в таблицы и фильтруем строки в подтаблицах
					TransformBinary = Table.TransformColumns(FilteredRows,         
						{"Content", each let 
							TransBinary  = Excel.Workbook(_),
							NoPrintAreas = Table.SelectRows(TransBinary, each not Text.Contains([Name], "$")),
							OnlySheets   = Table.SelectRows(NoPrintAreas, each [Kind]="Sheet")
							in OnlySheets
						} 
					),
					
					// добавляем колонки и фильтруем таблицу
					AddTwoColumns = let 
						AddSize   = Table.AddColumn(TransformBinary, "Size", (x)=> F[fxSizeFile](x[Attributes][Size]), type text),
						AddShName = Table.AddColumn(AddSize, "SheetsName",   (x)=> Text.Combine(x[Content][Name],","), type text),
						nms	 	  = {"Name", "Content"} & {"Size"} & {"Folder Path","SheetsName"},
						SelCol	  = Table.SelectColumns(AddShName, nms)
					in  SelCol,

				//  5. фильтруем колонки
					SelColumns4 = Table.SelectColumns(AddTwoColumns, {"Name", "Content"} & {"Folder Path"} & {"Size"}),

				// 	возвращаем разные этапы подключения к папке
					Return = Record.FieldOrDefault(
						[	1 = Trans, 			// ...
							2 = FilteredRows, 	// ...
							3 = AddTwoColumns,	// ...
							4 = SelColumns4		// ...
						], 	
							if   STEP = null 
							then Text.From(4) 		// последний шаг
							else Text.From(STEP)	// заданный шаг
						, 	
							"Error. Нет такого номера шага в функции LoadFolder")
				][Return],
				
			fnType = [ 
					optionsType	= 
						type [	
							ДатаCетИмя	= (type nullable text meta [Documentation.FieldCaption = "ДатаCетИмя",Documentation.SampleValues = {"Homel"}]),
							НомерШага 	= (type nullable logical meta [Documentation.FieldCaption = "НомерШага (от 1 до 5)",Documentation.SampleValues = {"1"}])
						],

					optionsMeta	= 
						[ 	Documentation.Name = "LoadFolder",
							Documentation.LongDescription = "подключаемся к папке, получаем список файлов XLSX с подпапками, чистим список. Шагов - 4. Аргументов - 2",
							Documentation.Examples = {
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"")",		Result="список файлов xls"], 
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"", 3)",	Result="список файлов xls на шаге 3"], 
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"", 123)",Result="сообщение об ошибке"]
							}
						],

					return 	    = type function (optional options as optionsType) as table meta optionsMeta
				][return],

			result = Value.ReplaceType(func,fnType)
		]  [result],

	Connect_    = // #2 подключаемся к файлам, чистим внутрянку, добавляем 4 техколонки
		// = Connect( [ДатаCетИмя = "Homel", ОтчетИмя = null, МаскаИмениПапки = "стат", boolfilterSheet = null, НомерШага = null)]
		// = Connect( [ДатаCетИмя = "Homel", ОтчетИмя = "opu", НомерШага = null])
		[	
			func   = (optional options as record) => [ 	
					//	0  формируем cписок аргументов //? сделать обязательные аргументы + по умолчанию, если они не введены
						DATASET_NAME = let value = Record.FieldOrDefault(options, "ДатаCетИмя", null) 		in if value = null then null else Text.From(value),
						REPORT_NAME  = let value = Record.FieldOrDefault(options, "ОтчетИмя", null) 		in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
						FOLDER_MASK  = let value = Record.FieldOrDefault(options, "МаскаИмениПапки", null) 	in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
						SHEET_FILTER = let value = Record.FieldOrDefault(options, "boolfilterSheet", null) 	in if value = null then null else Logical.From(value),
						STEP         = let value = Record.FieldOrDefault(options, "НомерШага", null) 		in if List.Contains({null, 0, ""}, value) then null else Number.From(value),

					//	#1 подключаемся к папке с файлами //* через LoadFolder_
						loadFile = F_Form[LoadFolder_]([ДатаCетИмя = DATASET_NAME]),
					
					//	#2 обрабатываем файлы и готовим их к анализу //* DATASET_NAME 
						// transform = Record.FieldOrDefault([ Homel = Homel ], DATASET_NAME, null),
						transform   = 	
							Record.FieldOrDefault(
								[
									Homel = Homel, 
									BKS   = BKS
								],  
								DATASET_NAME, null),

					//	#3 фильтруем файлы по form_ID //* REPORT_NAME
						filter = [
							nms = {"Company","DataSet","form_ID","Date"} & {"Content"},
							col = Table.SelectColumns(transform, nms),
							row = Table.SelectRows(col, each [form_ID] = REPORT_NAME),
							to  = try row otherwise "отчета """ & REPORT_NAME & """ в папке c маской """ & FOLDER_MASK & """ нет"
						][to],
						
					//	#r выгружаем с учетом параметров функции. Идет последним шагом! 
						mode = 
							if REPORT_NAME = null 		// СЕТ +, ОТЧЕТ -, ШАГ -
							then transform				// то #2 -> выводим ВСЕ файлы после трансформации + можно шаг Connect
							else if  STEP = null 		// СЕТ +, ОТЧЕТ +, ШАГ -
								then filter				// то #3 -> выводим отфильтрованные файлы
								else transform,			// то #2 -> выводим ВСЕ файлы после трансформации + можно шаг Connect
					
					//	1  список названий выходных столбцов этой функции 
						nms  = 	{"Company", "DataSet", "form_ID"} & {"Date"} & {"Content"} & {"Folder Path"},

					//  2  список датасетов (с каждым работать индивидуально) //? тут шаги - 5, для каждого датасета могут быть разные
						Homel = [ 
							// 	1 получаем список файлов в папке с учетом всех вложенностей
								// loadFile = F_Form[LoadFolder]("Homel", null),
							
							// 	2 фильтруем файлы в подтаблицах
								selectFiles = [
									lst = if   Text.Lower(FOLDER_MASK) = null
										then {""}
										else {FOLDER_MASK},
									f = each List.ContainsAny( { Text.Lower([Folder Path]) }, lst, Text.Contains),
									sel = Table.SelectRows(loadFile, f)
								] [sel],

							// 	3 фильтруем листы в подтаблицах
								// а если несколько листов попадает в такую выборку?!
								selectSheet =
									Table.TransformColumns(selectFiles,{"Content", (x)=>
										let sheet = Table.SelectRows(x, each Text.Length([Name]) = 15) 	// оставляем листы с длиной названия = 15 знаков - код организации
										in	sheet{0}[Data] 												// данные первого листа, //! Но... может быть несколько листов!
									}),

							// 	4 добавляем колонки с допинформацией о файле/отчете. Источник - данные из файла
								// 	в .xls дата = 01.16, в .xlsx = 01.12.2016 => возникает ошибка => обрабатываем ее через try/otherwise
								addColumns =
									let
										fxDate =(x)=> // ячейка [A1] - дата отчета
											try 	  Date.From("01.01.20" & Text.End(x[Content][Column1]{0},2))  // если xls  => 2 end знака зачения в ячейке с датой (в .xls дата = 01.16)
											otherwise Date.From(x[Content][Column1]{0} + #duration(30, 0, 0, 0)), // если xlsx => + 30 дней (в .xlsx = 01.12.2016 => возникает ошибка => обрабатываем ее через try/otherwise)
										fxCompany =(x)=> "Homel", 	//Text.From (x[Content][Column1]{2}), // ячейка [A1] - организация - ГОМЕЛЬСКОЕ РАЙПО / ГОМЕЛЬСКИЙ ФИЛИАЛ
										fxDataSet =(x)=> 2, 		// ячейка [A1] - код предприятия - тут не надо - Text.From (x[Content][Column1]{1})
										fxReport  =(x)=> Text.Start(x[Content][Column2]{0}, 4), // 4 первые знака = код отчета, например "0112"
										
										var = false, // true - берем данные с листа, false - руками заполняем
										
										// переключение режима откуда брать сделать в самой функции!!! а не через var дальше
										to  =
											F[fxAddMultiColumns](selectSheet,
												{
													{ "Date",    	if var then fxDate 	  else fxDate, 			type date},
													{ "DataSet", 	if var then fxDataSet else each 2, 			Int64.Type},
													{ "Company", 	fxCompany, 	type text},
													{ "form_ID",  	if var then fxReport  else fxReport, 		type text}
												}
											)
									in to,

							// 	5 меняем код отчета по нормам БКС на код отчета из моего справочника форм отчетности
								//	выбираем из справочника нужное название ЕСЛИ ФАЙЛОВ МНОГО (или берем с листа, если отчет один или руками ставим)
								transform = [
									f=(txt)=> // значения колонки 1 словаря, заменяю на колонку 0
										let dict = Table.ToRows(F_Form[spr.REPORT][Спр_ФормыОтчетности]), // список списков
											to   = List.First(List.Select(dict, (col)=>col{1}=txt)){0}
										in 	to,
									to = Table.TransformColumns(addColumns, {"form_ID", f, type text})
								][to],

							// 	6 перегруппируем листы и фильтруем колонки
								reorder	= Table.ReorderColumns(transform, nms),
								select  = Table.SelectColumns(reorder, nms),
								remove  = Table.RemoveColumns(select, {"Folder Path"}), 
								sort	= Table.Sort(remove, {{"form_ID", Order.Ascending}, {"Date", Order.Ascending}} ),

							//	  возвращаем данные конкретного шага max - 5
								Return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectFiles, 
										3 = selectSheet,
										4 = addColumns,
										5 = sort
									], 
										if   STEP = null
										then Text.From(5) 
										else Text.From(STEP)
									,
										"нет такого шага")
							] [Return],
						BKS   = [ 
							// 	1 подключаемся к папке
								// loadFile = F_Form[LoadFolder]("Drog"),

							// 	2 фильтруем файлы по маске имени папки: "" - оставляем все файлы по всех подпапках
								selectFiles = [
									lst = 	if   Text.Lower(FOLDER_MASK) = null
											then {""}
											else {Text.Lower(FOLDER_MASK)},
									f 	= each List.ContainsAny( { Text.Lower([Folder Path]) }, lst, Text.Contains),
									sel = Table.SelectRows(loadFile, f)
								] [sel],

							// 	3 фильтруем листы в подтаблицах (перенести шаг позднее)
								selectSheet = [
									f1 = each try Value.Is(Number.From([Name]), type number) otherwise null,
									f2 = each [Name] = "СВОД В РАЗРЕЗЕ",
									
									to = Table.TransformColumns(selectFiles, {"Content", (x)=> 
											// Table.SelectRows(x, f1)
											Table.SelectRows(x, 
												if SHEET_FILTER = null	// если опциональный параметр не указан
												then f1				// оставляем листы с номерами в названии, это коды организаций
												else f2
											)	
										})
								] [to],

							// 	4 добавляем колонки с допинформацией
								addColumns =
									let
									// 	вытягиваем дату из ячейки А1 первого листа с номером в названии первого файла отчета
										// 	fxDate    	=(x)=> Date.From(x[Content][Data]{0}[Column1]{0}),  // дата из отчета
											fxDate    =(x)=> 
												// конец года, чтобы везде одинаково было
												try Date.EndOfYear(Date.From(x[Content][Data]{0}[Column1]{0}))
												otherwise null, // ошибка если даты нет в отчете
																			
											fxCompany =(x)=> "BKS",
											fxDataSet =(x)=> 3,
									
										//	fxReport	формируем reportName для новой колонки по первым символам названия файла //! важный шаг
											fxReport  =(x)=> Record.FieldOrDefault([
												03 	= "v_3",
												07 	= "v_7",
												po 	= "v_12po",		// 12p ОПО
												pr 	= "v_12pr",		// 12p райпо
												PG 	= "s_1p1",		// 1п ОПС
												PQ 	= "s_1p2",		// 1п райпо
												11  = "v_11z",		// 11z заготовки
												7R 	= "s_1torg"		// 1 торг розница
											],  Text.Start(x[Name],2), null),

										// переключение режима откуда брать сделать в самой функции!!! а не через var дальше
										to  =
											F[fxAddMultiColumns](selectSheet, {
												{ "Date",    	fxDate, 	type date},
												{ "DataSet", 	fxDataSet , Int64.Type},
												{ "Company", 	fxCompany, 	type text},
												{ "form_ID",  	fxReport, 	type text}
											})
									in to,

							// 	5 перегруппируем листы и фильтруем колонки
								reorder 	= Table.ReorderColumns(addColumns, nms),
								selectRows	= Table.SelectRows(reorder, each [form_ID] <> null),
								selectCol 	= Table.SelectColumns(selectRows, nms),
								remove  	= Table.RemoveColumns(selectCol, {"Folder Path"}), 
								sort	    = Table.Sort(remove, {
												{"Date", 	Order.Ascending}, 
												{"form_ID",	Order.Ascending}
											}),
							
							//	  возвращаем данные конкретного шага - max 5	
								return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectFiles, 
										3 = selectSheet,
										4 = addColumns,
										5 = sort
									], 
										if   STEP = null
										then Text.From(5) 
										else Text.From(STEP)
									,
										"нет такого шага")
							] [return]			
				] [mode],

			fnType = [ 
					optionsType	= 
						type [	
							ДатаCетИмя 		= (type nullable text meta [
								Documentation.FieldCaption 	= "ДатаCетИмя (Имя датасета)", 
								Documentation.SampleValues 	= {"""Homel+"""},
								Documentation.Description 	= "Имя набора данных для обработки"
								]),
							ОтчетИмя 		= (type nullable text meta [
								Documentation.FieldCaption = "ОтчетИмя (ID отчета)", 	 
								Documentation.SampleValues = {"""bal"""},
								Documentation.Description = "Идентификатор отчета"
								]),
							МаскаИмениПапки = (type nullable text meta [
								Documentation.FieldCaption = "МаскаИмениПапки (Маска имени папки)", 
								Documentation.SampleValues = {"""стат"""},
								Documentation.Description = "Маска для имени папки"
								]),
							boolfilterSheet = (type nullable logical meta [
								Documentation.FieldCaption = "boolfilterSheet (Фильтр листов)", 
								Documentation.SampleValues = {true},
								Documentation.Description = "Применить ли фильтр к листам"
								]),
							НомерШага 		= (type nullable number meta [
								Documentation.FieldCaption = "НомерШага (Номер шага запроса: Homel 1-5)", 
								Documentation.SampleValues = {"1"},
								Documentation.Description = "Номер шага выполнения запроса"
								])
						],

					optionsMeta	= 
						[ 	
							Documentation.Name = "Connect",
							Documentation.LongDescription = 
								"	<b>Подключаемся к файлам, чистим внутрянку, добавляем 4 техколонки </b> <br>
									= F[Connect]( [ДатаCетИмя = ""Homel""] ) <br>
									= F[Connect]( [ДатаCетИмя = ""Homel"", ОтчетИмя = ""opu""] ) <br>
									= F[Connect]( [ДатаCетИмя = ""Homel"", ОтчетИмя = ""opu"", НомерШага = 1]) <br>
									= F[Connect]( [ДатаCетИмя = ""Homel"", ОтчетИмя = null,    НомерШага = null, МаскаИмениПапки = ""бух"", boolfilterSheet = null)]
								",
							Documentation.Examples = 
								{
									[
										Description = "", 
										Code 		= "=F_Form[Connect](""Homel"")()",
										Result		= "все обработанные файлы в папке"
									], 
									[	Description = "", 
										Code 		= "=F_Form[Connect](""Homel"",""bal"")()",
										Result		= "все отчеты с ID = bal в папке"
									], 
									[	Description = "", 
										Code 		= "=F_Form[Connect](""Homel"",""null"",""стат"")",
										Result		= "все обработанные файлы в папке + маска имени подпапки"
									], 
									[	Description = "", 
										Code 		= "=F_Form[Connect](""Homel"")(1)",
										Result		= "все отчеты в папке после общей чистки, шаг 1 (макс 5)"
									]
								}       
						],

					return 	    = type function (optional options as optionsType) as table meta optionsMeta // можно без "meta optionsMeta"
				]  [return],
			
			result = Value.ReplaceType(func, fnType)
		]  [result],

	Report_     = // #3 обрабатываем каждый файл исходя из вида отчета (form_ID) и датасета (Company)
		(options as record)  =>
		[
			// 	аргументы сделть имена полей как имена колонок в таблице: Company и form_ID
				DATASET_NAME = let value = Record.FieldOrDefault(options, "ДатаCет", null) 	in if value = null then null else Text.From(value),
				REPORT_NAME  = let value = Record.FieldOrDefault(options, "Отчет", null)   	in if List.Contains({null, 0, ""}, value) then null else Text.From(value),

			// 	functionMap	- карта функций
				functionMap = [
					Homel   = F_Form[rep.BAL] ("Homel"),
					BKS 	= F_Form[rep.BAL] ("BKS")
				],
			
			//	from		- отфильтрованная таблица с нужными отчетами по заданному датасету
				from = 	F_Form[Connect_] ([ДатаCетИмя = DATASET_NAME, ОтчетИмя = REPORT_NAME]),
			
			// 	проверку сделать по виду отчета! т.к. обрабатывает любое содержимое
				to   = 	
					Table.AddColumn(from, "DataTransform", each 
							 if [form_ID] = "bal" then F_Form[rep.BAL4] ([ДатаCетИмя = DATASET_NAME, Таблица = _[Content] ]) // это по новому
						else if [form_ID] = "opu" then 2
						else "нет данных"
						// let func = Record.FieldOrDefault(functionMap, [Company], (_)=>_)
						// in  func([Content])
					),
			//	удаляем столбец с исходниками + фильтруем по имени отчета // ? может пораньше делать это, до обработки
				del = Table.RemoveColumns(to, {"Content"}),
				// sel = Table.SelectRows(to, each ([form_ID] = REPORT_NAME)),
			
			// 	раскрывать таблицы надо сгруппировав их по виду отчетов
				exp = Table.ExpandTableColumn(del, "Step2", Table.ColumnNames(del[Step2]{0}))
		]  [to],

	SelectRow_  = // #3.1 фильтруем строки в подтаблицах по "Статья_ID" в rep.BAL3 и др на шаге Load (для отчетов которые имеют номер строки)
			/* 
				Описание:
					Функция добавляет колонку с отфильтрованными подтаблицами
				Пример:
					= 	F_Form[form.SelectRowsID](
							STEP1_Connect,
							{"Column1","Column2","Column3","Column4"},
							{"Статья", "Строка_ID", "На начало", "На конец"},
							"Строка_ID",
							10,
							1000
						)
				Улучшение:
					сделать проверку соответствия длин списков nmsOld и nmsNew
			*/
			[
				func   = (options as record) as table => [ 	
					// 1 формируем список аргументов 
						tbl  	 = let value = Record.FieldOrDefault(options, "tbl", null)      in if value is table then value else null,
						nmsOld	 = let value = Record.FieldOrDefault(options, "nmsOld", null)   in if value is list then value else {} ,
						nmsNew	 = let value = Record.FieldOrDefault(options, "nmsNew", null)   in if value is list then value else {} ,
						colName  = let value = Record.FieldOrDefault(options, "colName", null)  in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
						kodStart = let value = Record.FieldOrDefault(options, "kodStart", null) in if List.Contains({null, 0, ""}, value) then null else Number.From(value),
						kodEnd	 = let value = Record.FieldOrDefault(options, "kodEnd", null)   in if List.Contains({null, 0, ""}, value) then null else Number.From(value),
						subTable = let value = Record.FieldOrDefault(options, "subTable", null) in if Value.Type(value) = type logical then value else null,

					// 2 обрабатываем содержимое файлы, каждого или всех вместе
						fxOne = (table)=> // по новому, если надо отибработать отдельные файлы в подтаблицах
							[ 
							//	переименовываем колонки в таблице
								rename  = Table.RenameColumns(table, List.Zip({nmsOld, nmsNew} )),
								clear   = F[Table.ClearAll](rename," #(lf)#(00A0)", false, false),

							//	делаем колонку числовой, если там не число и возникает ошибка то меняем на null
								tr      = Table.TransformColumns(clear, {colName, try Number.From otherwise null, type number}),

							// 	удаляем errors, который иногда появляется, когда там текст или еще чего
								err     = Table.RemoveRowsWithErrors(tr, {colName}),
								lst     = {null, kodStart, kodEnd},
								f=(x)=> x>=lst{1} and x<=lst{2},

							//	оставляем строки в диапазоне номеров в колонке "Статья_ID"
								selRow  = Table.SelectRows(err, (row)=> f(Record.Field(row, colName) ) )
							]  [selRow],

						fxAll =	(table)=> // по старому, если одним запросом вывести все по виду отчета
							[ 
							//	добавляем колонку Step2 с подтаблицами, с отфильтроваными строками отчета
								add = Table.AddColumn(table, "Step2", (x)=> [
										//	проверяю наличие колонки. В одних отчетах подтаблицы в "Content", в других, как в БКС "Data" после разворота листов с предприятиями
											//!  сделать вариант, если грузить надо с листа один отчет, там нет колонки дата и контент
											selCol  = 	if   Table.HasColumns(tbl, "Content")
														then Table.SelectColumns(x[Content], nmsOld)
														else Table.SelectColumns(x[Data], nmsOld),
										
										//	обрабатываем отдельно каждый отчет 
											to = fxOne(selCol)
										]  [to]
									),

							// 	удаляем таблицу с исходниками (либо "Content" - Homel, либо "Data" - BKS, по наличию их на входе)
								rem	= Table.RemoveColumns(add, {"Content","Data"}, MissingField.Ignore)
							]  [rem],

					// 3 выводим резальтат 
						to = if subTable = true 
							 then fxAll(tbl) 
							 else fxOne(tbl)

					]  [to],
				fnType = [ 
						optionsType	= 
							type [	
								tbl	     = (type table meta [Documentation.FieldCaption = "tbl (таблица), table",Documentation.SampleValues = {"Source"}]),
								subTable = (type logical meta [Documentation.FieldCaption = "Источник подтаблица, true / false",Documentation.SampleValues = {"true / false"}]),
								nmsOld 	 = (type list meta [ // если type list, то можно дет выбрать из списка и поставить галочки
												Documentation.FieldCaption 		= "nmsOld (список старых имен колонок)",
												Documentation.FieldDescription 	= "Выберите допустимые поля ", // как можно увидеть этот текст?
												Documentation.SampleValues 		= {"{""Column1"",""Column2""}"},
												Documentation.AllowedValues 	= {10,110} // как сделать при этом произволльный ввод?
											]),
								nmsNew 	 = (type any meta [Documentation.FieldCaption = "nmsNew (список новых имен колонок)",Documentation.SampleValues = {"{""Column1"",""Column2""}"}]),
								colName	 = (type nullable text meta [Documentation.FieldCaption = "colName (имя колонки)",Documentation.SampleValues = {"""Column1"""}]),
								kodStart = (type nullable number meta [Documentation.FieldCaption = "kodStart (начальная строка)",Documentation.SampleValues = {10}]),
								kodEnd 	 = (type nullable number meta [Documentation.FieldCaption = "kodEnd (конечная строка)",Documentation.SampleValues = {1000}])
							],

						optionsMeta	= 
							[ 	Documentation.Name = "SelectRow",
								Documentation.LongDescription = "...",
								Documentation.Examples = {
									[Description = "", Code = "=...",Result="..."], 
									[Description = "", Code = "=...",Result="..."]
								}
							],

						return 	    = type function (options as optionsType) as table meta optionsMeta
					][return],

				result = Value.ReplaceType(func,fnType)
			]  
				[result],

	// старые варианты, убрать потом -------------------------------------------
	LoadReport_  = // #3 - старое - выгружаем готовый обработанный отчет для модели (с функцией tableBlank) (ИСПОЛЬЗУЕМ DATASET  )
		/*
			Описание: 
				1.	если можно сразу грузить отчет через F_Form[rep.BAL]("Homel")
				2. 	если отчет в базе есть, то выводим его, если нет, то пустую таблицу с заполненной датой и возможно другими тестовыми данными
			Пример: 
				= fxLoadReport("bal")	 
				= fxLoadReport("bal", null)	 	
				= fxLoadReport("bal", "Homel")
			Улучшение:
				1. 	опциональным агрументом или номер дата сета И название, с ним удобнее
			Исправить:
				1.	! сбиваются типы некоторых отчетов, например ofr. Именно после этой функции	
					неверное это тлько загрузки CSV касается, по которым нет в tableShema типов полей
		*/
		[
			func   = 
				(reportName as text, optional datasetName_ as nullable text) =>
				// (optional options as record) =>
				// let 
				[
				//	0  формируем cписок аргументов 
					// datasetName_ = let value = Record.FieldOrDefault(options, "ДатаCетИмя", null) in if value = null then null else Text.From(value),
					// reportName   = let value = Record.FieldOrDefault(options, "ОтчетИмя", null)   in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
				
				//	если datasetName не задан, то берем его из параметров F_Form[Parameters][DataSetName]
					datasetName =
						if   datasetName_ = null
						then F_Form[Parameters][DataSetName]
						else datasetName_,
				
				//	формируем имя запроса по типу и датасету DataSet_XLS_Homel или например DataSet_CSV_BKS
					QueryName  = 
						if 	 prm_LoadVar = "xls"
						then "DataSet_XLS_" & datasetName
						else "DataSet_CSV_" & datasetName,
				
				// 	проверяем, если ли такой запрос
					CheckQuery = 
						(namePattern as text) as logical =>
							let
								SharedItems     = Record.FieldNames(#shared),
								MatchingQueries = List.Select(SharedItems, each Text.Contains(_, namePattern)),
								QueriesExist    = List.Count(MatchingQueries) > 0
							in
								QueriesExist,

					CheckName  = // функция проверки наличия значения в колонке Name запроса DS = F_Form[DS]()
						(datasetName as text, eportName as text) as logical =>
							let
							//  Source   = Expression.Evaluate(datasetName, #shared), 	//!ошибка если отсюда грузить! таблица DataSet_XLS_Mart
								Source   = Record.Field(F_Form, datasetName),			// таблица DataSet_XLS_Mart
								NameList = Table.Column(Source, "Name"),				// список отчетов
								Exists   = List.Contains(NameList, reportName) 			// true/false
							in
								Exists,

					GetTable   = // функция получения таблицы из колонки Value по имени в колонке Name
						(datasetName as text, reportName as text) as table =>
							let
								// Source       = Expression.Evaluate(datasetName, #shared), // ошибка загрузки, в чате вопрос был
								Source1  = 
									Record.FieldOrDefault([
										DataSet_XLS_Mart  = DataSet_XLS_Mart, 
										DataSet_XLS_Homel = DataSet_XLS_Homel, 
										DataSet_XLS_BKS   = DataSet_XLS_BKS,
										DataSet_XLS_Drog  = DataSet_XLS_Drog
									],  datasetName, null),

								Source2  = 
									Record.FieldOrDefault([
										DataSet_CSV_Mart  = DataSet_CSV_Mart, 
										DataSet_CSV_Homel = DataSet_CSV_Homel, 
										DataSet_CSV_BKS   = DataSet_CSV_BKS,
										DataSet_CSV_Drog  = DataSet_CSV_Drog
									],  datasetName, null),	

								Source =	if prm_LoadVar = "xls"
											then Source1
											else Source2,

								FilteredRows = Table.SelectRows(Source, each [Name] = reportName),
								ResultTable  = 
									if Table.RowCount(FilteredRows) > 0 
									then Record.Field(FilteredRows{0}, "Value")
									else error "Значение не найдено"
							in
								ResultTable,

							
					IsQuery    = CheckQuery(datasetName),
					IsReport   = if IsQuery then CheckName(QueryName, reportName) else null,
					IsData     = if IsReport then GetTable(QueryName, reportName) else null,

					tableBlank = F_Form[form.BlankTable](DataShema, reportName),

					//	Выгружаем таблицы
						res =   if   IsQuery = true and IsReport = true // если найден отчет с такими параметрами в DataSets
								then IsData
								else tableBlank

				], //  [res]

			fnType = 
					type function (
						reportName as (type text meta [Documentation.FieldCaption = "Имя датасета",Documentation.SampleValues = {"""Homel"""}]),
						optional datasetName_  as (type nullable text meta [Documentation.FieldCaption = "form_ID отчета ",Documentation.SampleValues = {"""bal"""}])
					)
					as any 
						meta [ 	Documentation.Name = "LoadReport",
								Documentation.LongDescription = 
								" <b> Обрабатываем и выводим готовый нормализованный отчет или пустографку </b> <br>
									=F_Form[LoadReport](""bal"")			- выводим отчет активного датасета <br> 
									=F_Form[LoadReport](""bal"", null)		- то же <br> 
									=F_Form[LoadReport](""bal"", ""Mart"")	- выводим отчет датасета ""Mart"" <br> 
								",
								Documentation.Examples = 
								{
									[Description = "", Code = "=F_Form[LoadReport](""bal"", ""Mart"")",Result="отчет ""bal"" датасета ""Mart"" "]
								}       
							],
			result = Value.ReplaceType(func,fnType)
		]  [result],			
	
	SelectRow_2  = // #3 - старое - не через добавление колонки 
		[
			func   = (options as record) as table => [ 	
				// 1 формируем список аргументов 
					tbl  	 = let value = Record.FieldOrDefault(options, "tbl", null)     in if value is table then value else null,
					nmsOld	 = let value = Record.FieldOrDefault(options, "nmsOld", null)  in if value is list then value else {} ,
					nmsNew	 = let value = Record.FieldOrDefault(options, "nmsNew", null)  in if value is list then value else {} ,
					colName  = let value = Record.FieldOrDefault(options, "colName", null) in if List.Contains({null, 0, ""}, value) then null else Text.From(value),
					kodStart = let value = Record.FieldOrDefault(options, "kodStart", null)  in if List.Contains({null, 0, ""}, value) then null else Number.From(value),
					kodEnd	 = let value = Record.FieldOrDefault(options, "kodEnd", null)    in if List.Contains({null, 0, ""}, value) then null else Number.From(value),
			
				// 
					f =	[
						//	переименовываю
							rename  = Table.RenameColumns(tbl, List.Zip({nmsOld, nmsNew} )),
							clear   = F[Table.ClearAll](rename," #(lf)#(00A0)", false, false),

						//	делаем колонку числовой, если там не число и возникает ошибка то меняем на null
							tr      = Table.TransformColumns(clear, {colName, try Number.From otherwise null, type number}),

						// 	удаляем errors, который иногда появляется, когда там текст или еще чего
							err     = Table.RemoveRowsWithErrors(tr, {colName}),
							lst     = {null, kodStart, kodEnd},
							f=(x)=> x>=lst{1} and x<=lst{2},

						//	оставляем строки в диапазоне номеров в колонке "Статья_ID"
							selRow  = Table.SelectRows(err, (row)=> f(Record.Field(row, colName) ) )
						]  [selRow]
				][f]
			// fnType = [ 
			// 		optionsType	= 
			// 			type [	
			// 				tbl	     = (type table meta [Documentation.FieldCaption = "tbl (таблица), table",Documentation.SampleValues = {"Source"}]),
			// 				nmsOld 	 = (type list meta [ // если type list, то можно дет выбрать из списка и поставить галочки
			// 								Documentation.FieldCaption 		= "nmsOld (список старых имен колонок)",
			// 								Documentation.FieldDescription 	= "Выберите допустимые поля ", // как можно увидеть этот текст?
			// 								Documentation.SampleValues 		= {"{""Column1"",""Column2""}"},
			// 								Documentation.AllowedValues 	= {10,110} // как сделать при этом произволльный ввод?
			// 							]),
			// 				nmsNew 	 = (type any meta [Documentation.FieldCaption = "nmsNew (список новых имен колонок)",Documentation.SampleValues = {"{""Column1"",""Column2""}"}]),
			// 				colName	 = (type nullable text meta [Documentation.FieldCaption = "colName (имя колонки)",Documentation.SampleValues = {"""Column1"""}]),
			// 				kodStart = (type nullable number meta [Documentation.FieldCaption = "kodStart (начальная строка)",Documentation.SampleValues = {10}]),
			// 				kodEnd 	 = (type nullable number meta [Documentation.FieldCaption = "kodEnd (конечная строка)",Documentation.SampleValues = {1000}])
			// 			],

			// 		optionsMeta	= 
			// 			[ 	Documentation.Name = "SelectRow",
			// 				Documentation.LongDescription = "...",
			// 				Documentation.Examples = {
			// 					[Description = "", Code = "=...",Result="..."], 
			// 					[Description = "", Code = "=...",Result="..."]
			// 				}
			// 			],

			// 		return 	    = type function (options as optionsType) as table meta optionsMeta
			// 	][return],

			// result = Value.ReplaceType(func,fnType)
		]  
			[func],
	//	-------------------------------------
	// 	ТЕСТИРОВАНИЕ ЗАПРОСОВ
	//	-------------------------------------
	TEST_Function = 
		[
			Name        = "Homel",
			readme1     = "// Описание функций",
			LoadFolder  = F_Form[LoadFolder],
			Connect     = F_Form[Connect],
			LoadReport  = F_Form[LoadReport],
			rep.BAL     = F_Form[rep.BAL],
			
			readme2     = "// Результат выполнения",
			LoadFolder_ = F_Form[LoadFolder](Name),
			Connect_    = F_Form[Connect](Name)(),
			LoadReport_	= F_Form[LoadReport]("opu",Name),
			rep.BAL_    = F_Form[rep.BAL](Name)
		]
		,
	//	-------------------------------------
	// 	#INFO
	//	-------------------------------------
	getListFuncUDF1  = // список используемых в проекте функций из библиотеки //! F 
		/* 
			= F_Form[getListFuncUDF1] 
		*/
		let
			txt  = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\2. PQ запросы\F_Form.pq",
			from = Table.FromColumns({Lines.FromBinary(File.Contents(txt), null, null, 65001)}),
			indx = Table.AddIndexColumn(from, "Индекс", 1, 1, Int64.Type),
			add  = Table.AddColumn(indx, "TextEachFromTo", each try F[Text.EachFromTo]([Column1], "F[", "]"){0} otherwise null),
			grp  = Table.Group(add, {"TextEachFromTo"}, {
						{"All", 		each List.Distinct([Column1]) },
						{"Количество", 	each List.Count([Индекс]) },
						{"Строки", 		each Text.Combine(List.Transform([Индекс], Text.From), ",") }
					}),
			sel  = Table.SelectRows(grp, each ([TextEachFromTo] <> null and [TextEachFromTo] <> "F["""")),]"))
		in
			sel,
	getListFuncUDF2  = // список используемых в проекте функций из библиотеки //! F_Form 
		/* 
			= F_Form[getListFuncUDF2] 
		*/
		let
			txt  = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\2. PQ запросы\F_Form.pq",
			from = Table.FromColumns({Lines.FromBinary(File.Contents(txt), null, null, 65001)}),
			indx = Table.AddIndexColumn(from, "Индекс", 1, 1, Int64.Type),
			add  = Table.AddColumn(indx, "TextEachFromTo", each try F[Text.EachFromTo]([Column1], "F_Form[", "]"){0} otherwise null),
			grp  = Table.Group(add, {"TextEachFromTo"}, {
						// {"function", each Text.BetweenDelimiters([Column1], "F_Form[", "]"){0}},
						{"All", each List.Distinct([Column1]) },
						{"Количество", each List.Count([Индекс]) },
						{"Строки", each Text.Combine(List.Transform([Индекс], Text.From), ",") }
					}),
			sel  = Table.SelectRows(grp, each ([TextEachFromTo] <> null and [TextEachFromTo] <> "F_Form["""")),]")),
			nam  = Table.AddColumn(sel, "function", each Text.BetweenDelimiters([TextEachFromTo], "F_Form[", "]")),

			reo  = Table.ReorderColumns(nam,{"TextEachFromTo", "function", "All", "Количество", "Строки"}),
    		sel2 = Table.SelectRows(reo, each not Text.StartsWith([function], """")),
			sort = Table.Sort(sel2,{{"function", Order.Ascending}})
		in
			sort,		
	
	AllReportFromCSV = // #INFO все отчеты по всем датасетам из CSV файлов одной таблицей
		/* 	
			Пример:	
				= F_Form[AllReportFromCSV]
			Результат:
				На выходе 3 колонки: датасет, имя файла, Table
			Использование:
				Можно анализировать в какой базе, какие отчеты на сколько строки и т.п.
				Матрица: форма + статья в строки, Company в столбцы, количество статья в значения
			Варианты:
				первый шаг можно заменить на - rec = F_Form[ConnectCSV]
	    */
		let
			rec = [ 	
					Mart  = F_Form[LoadFolderCSV]( Value.Metadata( F_Form[Parameters] ) [Mart]  [pathCSV] ),
					Homel = F_Form[LoadFolderCSV]( Value.Metadata( F_Form[Parameters] ) [Homel] [pathCSV] ),
					BKS   = F_Form[LoadFolderCSV]( Value.Metadata( F_Form[Parameters] ) [BKS]   [pathCSV] ),
					Drog  = F_Form[LoadFolderCSV]( Value.Metadata( F_Form[Parameters] ) [Drog]  [pathCSV] )
				],
			RecordToTable 	    = Record.ToTable(rec),
			ExpandTableColumn   = Table.ExpandTableColumn(RecordToTable, "Value", {"Name", "Content", "Folder Path"}, {"Name.1", "Content", "Folder Path"}),
			TableAddColumn 	    = Table.AddColumn(ExpandTableColumn, "tmp", (x)=> F[LoadBinCSV](x[Content]) ),
			TableRemoveColumns  = Table.RemoveColumns(TableAddColumn,{"Content", "Folder Path"}),
			
			// 	раскрываем все таблицы, если нужно
			AllFieldNamesTables = F[AllFieldNamesTables](TableRemoveColumns, "tmp"),
			ExpandTableColumns  = Table.ExpandTableColumn(TableRemoveColumns, "tmp", AllFieldNamesTables)
		in
			TableRemoveColumns,
	//	-------------------------------------
	// 	РАБОТА С CSV
	//	-------------------------------------
	LoadFolderCSV  = // #1 подключаемся к папке, получаем очищенный список файлов CSV в папке с учетом подпапок
		(path as text) =>
		let
			from = Folder.Files(path),
			
			// 	фильтруем колонки и строки в списке файлов
				FilteredRows 		= Table.SelectRows(from, each Text.Start([Extension], 5) = ".csv" and [Attributes][Hidden] = false),
				NoTemps   			= Table.SelectRows(FilteredRows, each not Text.StartsWith([Name], "~$")),
				RemovedOtherColumns	= Table.SelectColumns(NoTemps, {"Name", "Content"} & {"Folder Path"}),
						
			// 	выводим последние 2 уровня названия папки
				NameFolder  		= Table.TransformColumns(RemovedOtherColumns, {
										{"Folder Path", each Text.AfterDelimiter(_, "\", {2, RelativePosition.FromEnd}), type text}}),
			
			SelectColumns		= Table.SelectColumns(NameFolder,{"Name", "Content"} & {"Folder Path"})
		in
			SelectColumns,

	ConnectCSV     = // #2 подключаемся к готовым файлам выгрузки в формате CSV (пока не используется)
		/*	
			Пример: подключение к папке = F_Form[ConnectCSV][Mart]	
		*/
		[ 	
			Mart  = F_Form[LoadReportCSV] ( Value.Metadata( F_Form[Parameters] ) [Mart]  [pathCSV] ),
			Homel = F_Form[LoadReportCSV] ( Value.Metadata( F_Form[Parameters] ) [Homel] [pathCSV] ),
			BKS   = F_Form[LoadReportCSV] ( Value.Metadata( F_Form[Parameters] ) [BKS]   [pathCSV] ),
			Drog  = F_Form[LoadReportCSV] ( Value.Metadata( F_Form[Parameters] ) [Drog]  [pathCSV] )
		],

	LoadReportCSV  = // #3 Важная функция: загрузка проекта из csv файла 
		/*
			Описание:
			Пример :
				= F_Form[LoadReportCSV]("Mart","bal","f_bal.csv")

			Улучшения:
				1	Если нет в схемы отчета в файле схем, то типа обработать. Или все тектовые или езще как
				2	F[LoadCSV] сделать функцию тут, стараться особых без ссылок на библиотеку
			
			Проверка:
				[
					datasetName = "Homel",
					reportName  = "pril",
					fileName    = "_f_pril.csv",
					nSteps      = 3,
				
					tblShema = DataShema,                                               //  Record  - запись с таблицами колонок отчетов и их типами
					param    = Record.Field(Value.Metadata(Parameters), datasetName),   //  Record  - параметры датасетов
					nms      = Record.Field(tblShema, reportName) [Header],             //  lst     - поля отчета согласно схем отчетных таблиц
					path     = param[pathCSV] & "\" & fileName,                         //  txt     - полный путь к файлу
					step     = if nSteps = null then "" else Text.From(nSteps),         //  txt     - тестовый номер шага

					load      = F[LoadCSV](path),                     					// грузим Table.Buffer?!
					reorder   = F[Reorder.AllColumns](load, nms),                   	// переупорядочиваем колонки
					shema     = F_Form[form.SetType2](reorder, tblShema, reportName),   // присваиваем колонкам типы конкретного отчета
					
					result    = Record.FieldOrDefault( [1 = load, 2 = reorder, 3 = shema], step, shema)
			]
		*/
		(datasetName as text, reportName as text, fileName as text, optional nSteps as number)=>
			[
				// datasetName = "Mart",
				// reportName  = "bal",
				// fileName    = "_f_bal.csv",
				// nSteps      = 3,
				tblShema = DataShema,                                                //  Record  - запись с таблицами колонок отчетов и их типами
				
				param    = Record.Field(Value.Metadata(Parameters), datasetName),    //  Record  - параметры датасетов
				nms      = Record.Field(tblShema, reportName) [Header],              //  lst     - поля отчета из файла схем отчетов
				path     = param[pathCSV] & "\" & fileName,                          //  txt     - полный путь к файлу
				step     = if nSteps = null then "" else Text.From(nSteps),          //  txt     - тестовый номер шага

				load      = F[LoadCSV](path),                     					 // грузим Table.Buffer?!
				nms2	  = Table.ColumnNames(load),								 // список колонок по факту, если отчета нет в базе
				reorder   = F[Reorder.AllColumns](load, try nms otherwise nms2),	 // переупорядочиваем колонки, по схеме или по факту
				shema     = try F_Form[form.SetType2](reorder, tblShema, reportName) // присваиваем колонкам типы конкретного отчета
							otherwise F[Table.TypedField](reorder),   				 // если отчета нет в базе, то типы по шаблону
				
				result    = Record.FieldOrDefault( [1 = load, 2 = reorder, 3 = shema], step, shema)
			
			]  [result],
		
	//	-------------------------------------
	// 	РАБОТА С XLSX
	//	-------------------------------------
	LoadFolder 	   = // #1 подключаемся к папке, получаем список файлов XLSX с подпапоками, чистим (ФОРМИРУЕМ ИСТОЧНИК  )
		/* 
			Описание:
				1	подключаемся к папке с файлами по имени датасета или по полному пути
				2	выбираем номер шага загрузки. По умолчанию выгружаем последний шаг
				3	чистим папку от временных и скрытых файлов, оставляем файлы .xlsx, преобразуем бинарники
				4	в книгах excel оставляем только листы (Sheet)
			
			Аргументы:
				DataSetName as text		-	путь к папке (формируется из параметров в шаге 1)
				optional n as number	-	номер шага в алгоритме выполнения функции
			
			Пример:	
				= F_Form[LoadFolder]("Homel")			// возвращается последний шаг
				= F_Form[LoadFolder]("Homel", null)		// возвращается последний шаг
				= F_Form[LoadFolder]("Homel", 3)		// возвращается третий шаг
				= F_Form[LoadFolder]("Homel", 123)		// возвращается надпись об ошибке
			
			Улучшения:
				1  	//! Folder.Contents – собираем данные из папки, но выборочно
				2 	Если датасет не выбран, то меняем например на Homel 
				3 	Сделать макеты таблиц всех отчетов для замены, если ничего нет или не выбрано.
				4 	Макеты таблиц с данными по Условной организации. Ее надо создать
				5 	Это могут быть данные из одного файла XlSX с данными всех отчетов
				6	Данные разных организаций можно совместить, чтобы заполнить базу и визуализации
			*/
		[	
			func   = 
				(DataSetName as text, optional n as number) =>
				let
				//	1. формируем path к папке с файлами XLSX
					path = 
						Record.FieldOrDefault(
							[   
								Mart  = Value.Metadata(Parameters) [Mart]  [pathXLSX],
								Homel = Value.Metadata(Parameters) [Homel] [pathXLSX],
								BKS   = Value.Metadata(Parameters) [BKS]   [pathXLSX],
								Drog  = Value.Metadata(Parameters) [Drog]  [pathXLSX]
							], 
							DataSetName, 
							if  Text.Contains(DataSetName,"\")	 	  // если имя содержит "\" (как часть пути к папке)
								then DataSetName					  // то это путь к папке (для датасетов не из параметров)
								else "Error. Ошибка в имени датасета" // то ошибка 	
					),

				// 	2. подключаемся к папке + фильтруем колонки + немного меняем формат колонок
					from  = Folder.Files(path),
					
					nms	  = {"Name","Extension", "Attributes", "Date created", "Content"} & {"Folder Path"},
					Sel	  = Table.SelectColumns(from, nms),
					Trans = Table.TransformColumns(Sel, {
						{"Date created", (x)=> DateTime.Date(x), type date},
						{"Folder Path",  (x)=> Text.AfterDelimiter(x, "\", {2, RelativePosition.FromEnd}), type text}
					}),

				// 	3. фильтруем строки
					FilteredRows 	= Table.SelectRows(Trans, each 
						(Text.Start([Extension], 5) = ".xlsx" 	// оставляем только .xlsx
						and [Attributes][Hidden] = false ) 		// убираем скрытые
						and not Text.StartsWith([Name], "~$")	// убираем временные
					),

				//  4. преобразуем бинарники в таблицы и фильтруем строки в подтаблицах
					TransformBinary = Table.TransformColumns(FilteredRows,         
						{"Content", each 
							let TransBinary  = Excel.Workbook(_),
								NoPrintAreas = Table.SelectRows(TransBinary, each not Text.Contains([Name], "$")),
								OnlySheets   = Table.SelectRows(NoPrintAreas, each [Kind]="Sheet")
							in 	OnlySheets
						} 
					),
					
					// добавляем колонки и фильтруем таблицу
					AddTwoColumns = let 
						AddSize   = Table.AddColumn(TransformBinary, "Size", (x)=> F[fxSizeFile](x[Attributes][Size]), type text),
						AddShName = Table.AddColumn(AddSize, "SheetsName",   (x)=> Text.Combine(x[Content][Name],","), type text),
						nms	 	  = {"Name", "Content"} & {"Size"} & {"Folder Path","SheetsName"},
						SelCol	  = Table.SelectColumns(AddShName, nms)
					in  SelCol,

				//  5. фильтруем колонки
					SelColumns4 = Table.SelectColumns(AddTwoColumns, {"Name", "Content"} & {"Folder Path"} & {"Size"}),

				// 	возвращаем разные этапы подключения к папке
					Return = Record.FieldOrDefault(
						[	1 = Trans, 			// ...
							2 = FilteredRows, 	// ...
							3 = AddTwoColumns,	// ...
							4 = SelColumns4		// ...
						], 	
							if   n = null 
							then Text.From(4) 	// последний шаг
							else Text.From(n)	// заданный шаг
						, 	
							"Error. Нет такого номера шага в функции LoadFolder")
				in	Return,

			fnType = 
				type function (
					DataSetName as (type text  meta [Documentation.FieldCaption = "Имя датасета",Documentation.SampleValues = {"Homel"}]),
					optional n as (type number meta [Documentation.FieldCaption = "Номер шага (от 1 до 5)",Documentation.SampleValues = {"1"}])
				)
				as table 
					meta [ 	Documentation.Name = "LoadFolder",
							Documentation.LongDescription = "подключаемся к папке, получаем список файлов XLSX с подпапками, чистим список",
							Documentation.Examples = 
							{
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"")",Result="список файлов xls"], 
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"", 3)",Result="список файлов xls на шаге 3"], 
								[Description = "", Code = "=F_Form[LoadFolder](""Homel"", 123)",Result="сообщение об ошибке"]
							}       
						],
			result = Value.ReplaceType(func,fnType)
		]  [result],

	Connect    	   = // #2 подключаемся к файлам, чистим внутрянку, добавляем 4 техколонки         (ФОРМИРУЕМ DATASET ?!)
		/*	
			Описание:
				Импорт из файлов в папке всех листов -> списк таблиц ПОСЛЕ ПЕРВОЙ ЧИСТКИ ПАПКИ И ФАЙЛОВ
				На выходе 6 столбцов: 4 тех + 1 "Contents" + 1 "Folder Path"
		
			Пример:
				1	фильтруем файлы (наличие символов) по маске имени папки (если там лежат отчеты одного вида)
					Connect = 	F_Form[Connect](
									F_Form[GetName](datasetName),
									null,
									"Фоpма 12-п (месячная, срочная) ОПО"
								)(),
				
				2	фильтруем файлы (2 знака) по маске имени файла (если там лежат самые разные отчеты)
					Connect = 	F_Form[Connect](
									F_Form[GetName](datasetName),
									"v_12po",
									null
								)(),
				
				= F_Form[Connect]("Homel")()			// все обработанные файлы в папке  //! Основной вариант
				= F_Form[Connect]("Homel","bal")()		// все отчеты данного вида в папке //! Основной вариант
				
				= F_Form[Connect]("Homel")(1)			// все отчеты в папке после общей чистки, шаг 1 (макс 5)
				= F_Form[Connect]("Homel", null)(1)		// результат то же, аналог F_Form[LoadFolder]("Homel")
				= F_Form[Connect]("Homel","bal")(1)		// результат то же

				= F_Form[Connect]("Homel", null, "стат")()	// все обработанные файлы в папке + маска имени подпапки 
				= F_Form[Connect]("Drog",  null, "бух" )()	// все отчеты в папке с бухотчетностью

			Дополнительно:
				= F_Form[LoadFolder]("Homel",1)			//? проверяем исходное состояние папки
				= F_Form[LoadFolder]("Homel")			//? проверяем конечное состояние папки

			Улучшение:
				1. Сделать описание функции
				2. Понятное описание переменных в окне ввода аргументов через meta, описание функции
				3. Придумать что-то с if n = null then Text.From(5) else Text.From(n) в шаге Return
				4. Сделать функцию по фильтрации папки шага selectFiles.   В каждом датасете используется
				5. Сделать функцию по фильтрации листов шага selectSheets. В каждом датасете используется
				6. номер шага потом можно убрать, т.к. вторые скобки ставить забываешь иногда
				7. аргументы последовательно может сделать?! Фильтруем файлы, фильтруем отчет и только потом листы
			
			Важно: получить инфомацию об Организации, НомереДатаСета, Названия отчета для заполнения технических колонок:
				1. Из названия файла Код отчета остальное руками
				2. Из содержимого файла
				3. В ручную, если проект разовый
				Дальше это и делаем
			*/
		[	
			func   = 
				(	DATASET_NAME 		 	as text, 
					optional REPORT_NAME 	as nullable text,
					optional maskNameFolder as nullable text,
					optional filterSheet 	as nullable number
				) => 
				( optional n as nullable number) =>
				[ 	
					//	#1 подключаемся к папке с файлами 
						loadFile =	
							F_Form[LoadFolder](DATASET_NAME),
					
					//	#2 обрабатываем файлы конкретного датасета "Имя поля = Имя шага" и готовим их к анализу 
						transform   = 	Record.FieldOrDefault([
										Mart  = Mart, 
										Homel = Homel, 
										BKS   = BKS,
										Drog  = Drog
									],  DATASET_NAME, null),

					//	#3 фильтруем файлы
						filter = [
							FilterRowAndCol = (tbl as table, reportName as text) =>
								Table.SelectColumns(
									Table.SelectRows(tbl, each [form_ID] = reportName), 
									{"Company","DataSet","form_ID","Date", "Content"}
								),
							to = try FilterRowAndCol(transform, REPORT_NAME)
								 otherwise "отчета """ & REPORT_NAME & """ в папке c маской """ & maskNameFolder & """ нет"
							][to],
						
					//	#r выгружаем с учетом параметров функции. Идет последним шагом! 
						mode = 
							if REPORT_NAME = null 		// СЕТ +, ОТЧЕТ -, ШАГ -
							then transform				// то #2 -> выводим ВСЕ файлы после трансформации + можно шаг Connect
							else if  n = null 			// СЕТ +, ОТЧЕТ +, ШАГ -
								then filter				// то #3 -> выводим отфильтрованные файлы
								else transform,			// то #2 -> выводим ВСЕ файлы после трансформации + можно шаг Connect
					
					//	1 список названий выходных столбцов этой функции 
						nms  = 	{"Company", "DataSet", "form_ID"} &
								{"Date"} & {"Content"} & {"Folder Path"},

					//  2 список датасетов (с каждым работать индивидуально)
						Mart  = [ 
							// 	1 получаем список файлов в папке с учетом всех вложенностей
								// loadFile = F_Form[LoadFolder]("Mart"),

							//  2 меняем значения в колонке с именем файла на дату, полученную из имени файла
								dateColumn = [
									f=(x)=> Date.From(Text.Middle(x, 2, 10)),	// 10 символов со смещенеим 2
									trs = Table.TransformColumns(loadFile, {"Name", f, Date.Type}),
									to  = Table.RenameColumns(trs, {"Name", "DateTmp"}) // так как при развороте создается колонка с таким же именем
								][to],

							// 	3 разворачиваем подтаблицы и фильтруем листы (лист "Баланс" + "Прил.2"+ "Прил.3"+ ... + "Прил.5")
								selectSheet = [
									lst = {"Name", "Data", "Kind"},
									exp = Table.ExpandTableColumn(dateColumn, "Content", lst, lst),
									ren = Table.RenameColumns(exp, {"Data", "Content"}),
									del = Table.RemoveColumns(Table.SelectRows(ren, each ([Kind] = "Sheet")), {"Kind"}),
									sel = Table.SelectRows(del, each List.ContainsAny({[Name]}, {"Баланс","Прил"}, Text.Contains))
								]  [sel], // del если посмотреть все листы в книге

							// 	4 добавляем колонки с допинформацией о файле/отчете.
								addColumns =
									let
										fxDate    =(x)=> x[DateTmp], // из названия - 10 символов начиная со второго
										fxCompany =(x)=> "Mart", 	 // ручками
										fxDataSet =(x)=> 1, 		 // сделать из параметров
										fxReport  =(x)=> x[Name], 	 // имя отчета = название листа
										to  =
											F[fxAddMultiColumns](selectSheet,
												{
													{ "Date",    	fxDate, 	type date},
													{ "DataSet", 	fxDataSet , Int64.Type},
													{ "Company", 	fxCompany, 	type text},
													{ "form_ID",  fxReport, 	type text}
												}
											)
									in to,

							//  5 меняем название отчетов в файле (= имя листа) на нужное справочное
								replace = 
									let	replaceList  = {
											{"Баланс", 	"bal"},
											{"Прил.2", 	"opu"},
											{"Прил.3", 	"osk"},
											{"Прил.4", 	"odds"},
											{"Прил.5",  "ocf"}
										},
										dict = List.Buffer(replaceList),
										func = (txt)=> List.Accumulate(dict, txt, (s,x)=>Text.Replace(s,x{0},x{1}))
									in 	Table.TransformColumns(addColumns, {"form_ID", func}),

							// 	6 фильтруем, перегруппируем и сортируем колонки
								reorder 	= 	Table.ReorderColumns(replace, nms),
								select  	= 	Table.SelectColumns(reorder, nms),
								remove  	= 	Table.RemoveColumns(select, {"Folder Path"}), 
								sortToDate  = 	Table.Sort(remove, {
													{"Date", 	Order.Ascending}, 
													{"form_ID",	Order.Ascending}
												}),
							
							//	  возвращаем данные конкретного шага	
								Return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectSheet, 
										3 = addColumns,
										4 = replace,
										5 = sortToDate
									], 
										if   n = null
										then Text.From(5) 
										else Text.From(n)
									,
										"нет такого шага")
							] [Return],

						Homel = [ 
							// 	1 получаем список файлов в папке с учетом всех вложенностей
								// loadFile = F_Form[LoadFolder]("Homel", null),
							
							// 	2 фильтруем файлы в подтаблицах
								selectFiles = [
									lst = if   Text.Lower(maskNameFolder) = null
										then {""}
										else {maskNameFolder},
									f = each List.ContainsAny( { Text.Lower([Folder Path]) }, lst, Text.Contains),
									sel = Table.SelectRows(loadFile, f)
								] [sel],

							// 	3 фильтруем листы в подтаблицах
								// а если несколько листов попадает в такую выборку?!
								selectSheet =
									Table.TransformColumns(selectFiles,{"Content", (x)=>
										let sheet = Table.SelectRows(x, each Text.Length([Name]) = 15) 	// оставляем листы с длиной названия = 15 знаков - код организации
										in	sheet{0}[Data] 												// данные первого листа, //! Но... может быть несколько листов!
									}),

							// 	4 добавляем колонки с допинформацией о файле/отчете. Источник - данные из файла
								// 	в .xls дата = 01.16, в .xlsx = 01.12.2016 => возникает ошибка => обрабатываем ее через try/otherwise
								addColumns =
									let
										fxDate =(x)=> // ячейка [A1] - дата отчета
											try 	  Date.From("01.01.20" & Text.End(x[Content][Column1]{0},2))  // если xls  => 2 end знака зачения в ячейке с датой (в .xls дата = 01.16)
											otherwise Date.From(x[Content][Column1]{0} + #duration(30, 0, 0, 0)), // если xlsx => + 30 дней (в .xlsx = 01.12.2016 => возникает ошибка => обрабатываем ее через try/otherwise)
										fxCompany =(x)=> "Homel", 	//Text.From (x[Content][Column1]{2}), // ячейка [A1] - организация - ГОМЕЛЬСКОЕ РАЙПО / ГОМЕЛЬСКИЙ ФИЛИАЛ
										fxDataSet =(x)=> 2, 		// ячейка [A1] - код предприятия - тут не надо - Text.From (x[Content][Column1]{1})
										fxReport  =(x)=> Text.Start(x[Content][Column2]{0}, 4), // 4 первые знака = код отчета, например "0112"
										
										var = false, // true - берем данные с листа, false - руками заполняем
										
										// переключение режима откуда брать сделать в самой функции!!! а не через var дальше
										to  =
											F[fxAddMultiColumns](selectSheet,
												{
													{ "Date",    	if var then fxDate 	  else fxDate, 			type date},
													{ "DataSet", 	if var then fxDataSet else each 2, 			Int64.Type},
													{ "Company", 	fxCompany, 	type text},
													{ "form_ID",  if var then fxReport  else fxReport, 		type text}
												}
											)
									in to,

							// 	5 меняем код отчета по нормам БКС на код отчета из моего справочника форм отчетности
								//	выбираем из справочника нужное название ЕСЛИ ФАЙЛОВ МНОГО (или берем с листа, если отчет один или руками ставим)
								transform = [
									f=(txt)=> // значения колонки 1 словаря, заменяю на колонку 0
										let dict = Table.ToRows(F_Form[spr.REPORT][Спр_ФормыОтчетности]), // список списков
											to   = List.First(List.Select(dict, (col)=>col{1}=txt)){0}
										in 	to,
									to = Table.TransformColumns(addColumns, {"form_ID", f, type text})
								][to],

							// 	6 перегруппируем листы и фильтруем колонки
								reorder	= Table.ReorderColumns(transform, nms),
								select  = Table.SelectColumns(reorder, nms),
								remove  = Table.RemoveColumns(select, {"Folder Path"}), 
								sort	= Table.Sort(remove, {{"form_ID", Order.Ascending}, {"Date", Order.Ascending}} ),

							//	  возвращаем данные конкретного шага	
								Return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectFiles, 
										3 = selectSheet,
										4 = addColumns,
										5 = sort
									], 
										if   n = null
										then Text.From(5) 
										else Text.From(n)
									,
										"нет такого шага")
							] [Return],

						BKS   = [ 
							// 	1 подключаемся к папке
								// loadFile = F_Form[LoadFolder]("Drog"),

							// 	2 фильтруем файлы по маске имени папки: "" - оставляем все файлы по всех подпапках
								selectFiles = [
									lst = 	if   Text.Lower(maskNameFolder) = null
											then {""}
											else {Text.Lower(maskNameFolder)},
									f 	= each List.ContainsAny( { Text.Lower([Folder Path]) }, lst, Text.Contains),
									sel = Table.SelectRows(loadFile, f)
								] [sel],

							// 	3 фильтруем листы в подтаблицах (перенести шаг позднее)
								selectSheet = [
									f1 = each try Value.Is(Number.From([Name]), type number) otherwise null,
									f2 = each [Name] = "СВОД В РАЗРЕЗЕ",
									
									to = Table.TransformColumns(selectFiles, {"Content", (x)=> 
											// Table.SelectRows(x, f1)
											Table.SelectRows(x, 
												if filterSheet = null	// если опциональный параметр не указан
												then f1				// оставляем листы с номерами в названии, это коды организаций
												else f2
											)	
										})
								] [to],

							// 	4 добавляем колонки с допинформацией
								addColumns =
									let
									// 	вытягиваем дату из ячейки А1 первого листа с номером в названии первого файла отчета
										// 	fxDate    	=(x)=> Date.From(x[Content][Data]{0}[Column1]{0}),  // дата из отчета
											fxDate    =(x)=> 
												// конец года, чтобы везде одинаково было
												try Date.EndOfYear(Date.From(x[Content][Data]{0}[Column1]{0}))
												otherwise null, // ошибка если даты нет в отчете
																			
											fxCompany =(x)=> "BKS",
											fxDataSet =(x)=> 3,
									
										//	fxReport	формируем reportName для новой колонки по первым символам названия файла //! важный шаг
											fxReport  =(x)=> Record.FieldOrDefault([
												03 	= "v_3",
												07 	= "v_7",
												po 	= "v_12po",		// 12p ОПО
												pr 	= "v_12pr",		// 12p райпо
												PG 	= "s_1p1",		// 1п ОПС
												PQ 	= "s_1p2",		// 1п райпо
												11  = "v_11z",		// 11z заготовки
												7R 	= "s_1torg"		// 1 торг розница
											],  Text.Start(x[Name],2), null),

										// переключение режима откуда брать сделать в самой функции!!! а не через var дальше
										to  =
											F[fxAddMultiColumns](selectSheet, {
												{ "Date",    	fxDate, 	type date},
												{ "DataSet", 	fxDataSet , Int64.Type},
												{ "Company", 	fxCompany, 	type text},
												{ "form_ID",  	fxReport, 	type text}
											})
									in to,

							// 	5 перегруппируем листы и фильтруем колонки
								reorder 	= Table.ReorderColumns(addColumns, nms),
								selectRows	= Table.SelectRows(reorder, each [form_ID] <> null),
								selectCol 	= Table.SelectColumns(selectRows, nms),
								remove  	= Table.RemoveColumns(selectCol, {"Folder Path"}), 
								sort	    = Table.Sort(remove, {
												{"Date", 	Order.Ascending}, 
												{"form_ID",	Order.Ascending}
											}),
							
							//	  возвращаем данные конкретного шага	
								return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectFiles, 
										3 = selectSheet,
										4 = addColumns,
										5 = sort
									], 
										if   n = null
										then Text.From(5) 
										else Text.From(n)
									,
										"нет такого шага")
							] [return],
						
						Drog  = [ 
							// 	1 подключаемся к папке
								// loadFile = F_Form[LoadFolder]("Drog"),

							// 	2 фильтруем файлы в папке по маске имени папки (можно {"бух","стат"})
								selectFiles = [
									lst = 	if   Text.Lower(maskNameFolder) = null
											then {"бух","стат"} 	// маска имени папки, 2 папки, все файлы
											else {Text.Lower(maskNameFolder)},
									f 	= each List.ContainsAny( { Text.Lower([Folder Path]) }, lst, Text.Contains),
									sel = Table.SelectRows(loadFile, f)
								] [sel],

							// 	3 фильтруем листы в файлах с отчетами по маске имени папки (т.к. у них разное наполнение)
								selectSheet = 
									F[Table.TransformContent](
										selectFiles, 
										"Content",		//	имя колонки для замены значений в подтаблциах
										"Folder Path",	// 	имя колонки для поиска текста по маске
										{ 				// 	col{1} - номер колонки c таблицей, Content
											{"ухот",   (col)=> Table.SelectRows(col{1}, each [Name] = "СВОД")},
											{"4-торг", (col)=> Table.SelectRows(col{1}, each Text.Length([Name]) = 15)},
											{"12ф",	   (col)=> Table.SelectRows(col{1}, each Text.Contains([Name],"Раздел"))}
										}
									),
								
								// обрабатываем каждую подтаблицу (файл) в зависимости от числа строк (количество листов после фильтрации) в них 
								dataselectSheet = 
									Table.TransformColumns(selectSheet, 
										{"Content", each let
											rowCount = Table.RowCount(_),
											result   = 
												
												// если лист a файле олдин
													 if rowCount = 1 then [Data]{0} 
												
												// если листов в файле более одного
												else if rowCount > 1 then 
													let a = Table.TransformColumns(_, {"Data", (i)=>F[Table.ClearAll](i," #(lf)#(00A0)", 1, 1)}),
														b = Table.Combine(a[Data])
													in  b
												
												else _ 
											in	result
										}
									),

							// 	4 добавляем колонки с допинформацией
								addColumns =
									let
										// функции
											fxCompany =(x)=> "Drog",
											fxDataSet =(x)=> 4,
											fxDate    =(x)=> 
												[	txt = Text.SplitAny(x[Name],"_ "){0}, // первые 2 знака в имени файла (до "_ ")
													to  = Date.From("01" & "." & Text.Start(txt, 2) & "." & Text.End(txt, 2))
												]  [to],
											fxReport  =(x)=> 
												Record.FieldOrDefault(
													[01 = "bal", 02 = "opu"], 
												Text.Start(x[Name],2), 
												null),

										// 	переключение режима откуда брать сделать в самой функции!!! а не через var дальше
											to  =
												F[fxAddMultiColumns](dataselectSheet,
													{
														{ "Date",    	fxDate, 	type date},
														{ "DataSet", 	fxDataSet , Int64.Type},
														{ "Company", 	fxCompany, 	type text},
														{ "form_ID",  	fxReport, 	type text}
													}
												)
									in to,

							// 	5 перегруппируем листы и фильтруем колонки
								reorder = Table.ReorderColumns(addColumns, nms),
								select  = Table.SelectColumns(reorder, nms),
								result  = Table.RemoveColumns(select, {"Folder Path"}), 
							
							//	  возвращаем данные конкретного шага	
								Return = Record.FieldOrDefault([
										1 = loadFile, 
										2 = selectFiles, 
										3 = dataselectSheet,
										4 = addColumns,
										5 = result
									], 
										if   n = null
										then Text.From(5) 
										else Text.From(n)
									,
										"нет такого шага")
							] [Return]
				] [mode],
			fnType = 
				type function (
					DATASET_NAME as (type text meta [Documentation.FieldCaption = "Имя датасета",Documentation.SampleValues = {"""Homel"""}]),
					optional REPORT_NAME    as (type nullable text meta [Documentation.FieldCaption = "ID отчета ",Documentation.SampleValues = {"""bal"""}]),
					optional maskNameFolder as (type nullable text meta [Documentation.FieldCaption = "Маска имени папки",Documentation.SampleValues = {"""стат"""}]),
					optional filterSheet    as (type nullable number)
				)
				as table 
					meta [ 	Documentation.Name = "Connect",
							Documentation.LongDescription = 
							" <b>Подключаемся к файлам, чистим внутрянку, добавляем 4 техколонки </b> <br>
								= F_Form[Connect](""Homel"")()			- все файлы в папке	после общей чистки <br> 
								= F_Form[Connect](""Homel"")(1)			- все файлы в папке после общей чистки, шаг 1 из 5 <br>
								= F_Form[Connect](""Homel"",""bal"")()	- все отчеты данного вида в папке	<br>
								= F_Form[Connect](""Homel"", null)(1)	- то же, аналог F_Form[LoadFolder](""Homel"") <br>
								= F_Form[Connect](""Homel"",""bal"")(1)	- то же <br>
								= F_Form[Connect](""Homel"", null, ""стат"")() - все обработанные файлы в папке + маска имени подпапки <br>
							",
							Documentation.Examples = 
							{
								[Description = "", Code = "=F_Form[Connect](""Homel"")()",Result="все обработанные файлы в папке"], 
								[Description = "", Code = "=F_Form[Connect](""Homel"",""bal"")()",Result="все отчеты с ID = bal в папке"], 
								[Description = "", Code = "=F_Form[Connect](""Homel"",""null"",""стат"")",Result="все обработанные файлы в папке + маска имени подпапки"], 
								[Description = "", Code = "=F_Form[Connect](""Homel"")(1)",Result="все отчеты в папке после общей чистки, шаг 1 (макс 5)"]
							}       
						],
			result = Value.ReplaceType(func,fnType)
		]  [result],

	LoadReport     = // #3 выгружаем готовый обработанный отчет для модели (с функцией tableBlank) (ИСПОЛЬЗУЕМ DATASET  )
		/*
			Описание: 
				1.	если можно сразу грузить отчет через F_Form[rep.BAL]("Homel")
				2. 	если отчет в базе есть, то выводим его, если нет, то пустую таблицу с заполненной датой и возможно другими тестовыми данными
			Пример: 
				= fxLoadReport("bal")	 
				= fxLoadReport("bal", null)	 	
				= fxLoadReport("bal", "Homel")
			Улучшение:
				1. 	опциональным агрументом или номер дата сета И название, с ним удобнее
			Исправить:
				1.	! сбиваются типы некоторых отчетов, например ofr. Именно после этой функции	
					неверное это тлько загрузки CSV касается, по которым нет в tableShema типов полей
		*/
		[
			func   = 
				(reportName as text, optional datasetName_ as nullable text) =>
				// let 
				[
						datasetName =
						if   datasetName_ = null
						then F_Form[Parameters][DataSetName]
						else datasetName_,

					QueryName  = 
						if 	 prm_LoadVar = "xls"
						then "DataSet_XLS_" & datasetName
						else "DataSet_CSV_" & datasetName,

					CheckQuery = 
						(namePattern as text) as logical =>
							let
								SharedItems     = Record.FieldNames(#shared),
								MatchingQueries = List.Select(SharedItems, each Text.Contains(_, namePattern)),
								QueriesExist    = List.Count(MatchingQueries) > 0
							in
								QueriesExist,

					CheckName  = // функция проверки наличия значения в колонке Name запроса DS = F_Form[DS]()
						(datasetName as text, eportName as text) as logical =>
							let
							//  Source   = Expression.Evaluate(datasetName, #shared), 	//!ошибка если отсюда грузить! таблица DataSet_XLS_Mart
								Source   = Record.Field(F_Form, datasetName),			// таблица DataSet_XLS_Mart
								NameList = Table.Column(Source, "Name"),				// список отчетов
								Exists   = List.Contains(NameList, reportName) 			// true/false
							in
								Exists,

					GetTable   = // функция получения таблицы из колонки Value по имени в колонке Name
						(datasetName as text, reportName as text) as table =>
							let
								// Source       = Expression.Evaluate(datasetName, #shared), // ошибка загрузки, в чате вопрос был
								Source1  = 
									Record.FieldOrDefault([
										DataSet_XLS_Mart  = DataSet_XLS_Mart, 
										DataSet_XLS_Homel = DataSet_XLS_Homel, 
										DataSet_XLS_BKS   = DataSet_XLS_BKS,
										DataSet_XLS_Drog  = DataSet_XLS_Drog
									],  datasetName, null),

								Source2  = 
									Record.FieldOrDefault([
										DataSet_CSV_Mart  = DataSet_CSV_Mart, 
										DataSet_CSV_Homel = DataSet_CSV_Homel, 
										DataSet_CSV_BKS   = DataSet_CSV_BKS,
										DataSet_CSV_Drog  = DataSet_CSV_Drog
									],  datasetName, null),	

								Source =	if prm_LoadVar = "xls"
											then Source1
											else Source2,

								FilteredRows = Table.SelectRows(Source, each [Name] = reportName),
								ResultTable  = 
									if Table.RowCount(FilteredRows) > 0 
									then Record.Field(FilteredRows{0}, "Value")
									else error "Значение не найдено"
							in
								ResultTable,

							
					IsQuery    = CheckQuery(datasetName),
					IsReport   = if IsQuery then CheckName(QueryName, reportName) else null,
					IsData     = if IsReport then GetTable(QueryName, reportName) else null,

					tableBlank = F_Form[form.BlankTable](DataShema, reportName),

					//	Выгружаем таблицы
						res =   if   IsQuery = true and IsReport = true // если найден отчет с такими параметрами в DataSets
								then IsData
								else tableBlank
				],// in	res,

			fnType = 
					type function (
						reportName as (type text meta [Documentation.FieldCaption = "Имя датасета",Documentation.SampleValues = {"""Homel"""}]),
						optional datasetName_  as (type nullable text meta [Documentation.FieldCaption = "form_ID отчета ",Documentation.SampleValues = {"""bal"""}])
					)
					as any 
						meta [ 	Documentation.Name = "LoadReport",
								Documentation.LongDescription = 
								" <b> Обрабатываем и выводим готовый нормализованный отчет или пустографку </b> <br>
									=F_Form[LoadReport](""bal"")			- выводим отчет активного датасета <br> 
									=F_Form[LoadReport](""bal"", null)		- то же <br> 
									=F_Form[LoadReport](""bal"", ""Mart"")	- выводим отчет датасета ""Mart"" <br> 
								",
								Documentation.Examples = 
								{
									[Description = "", Code = "=F_Form[LoadReport](""bal"", ""Mart"")",Result="отчет ""bal"" датасета ""Mart"" "]
								}       
							],
			result = Value.ReplaceType(func,fnType)
		]  [result],			
	
	//	-------------------------------------
	// 	ЗАГРУЗКА И ПРЕДВАРИТЕЛЬНАЯ ЧИСТКА
	//	-------------------------------------
	form.SelectRowsID = // #! фильтруем строки в подтаблицах по "Статья_ID" на шаге Load (для отчетов которые имеют номер строки) . Надо документировать функцию 
		/* 
			Описание:
				Функция добавляет колонку с отфильтрованными подтаблицами
			Пример:
				= 	F_Form[form.SelectRowsID](
						STEP1_Connect,
						{"Column1","Column2","Column3","Column4"},
						{"Статья", "Строка_ID", "На начало", "На конец"},
						"Строка_ID",
						10,
						1000
					)
			Улучшение:
				сделать проверку соответствия длин списков nmsOld и nmsNew
		*/
		(	tbl 		as table,	//
			nmsOld 		as list,	//
			nmsNew 		as list,	//
			columnName 	as text, 	//
			kodStart 	as number, 	//
			kodEnd 		as number	//
		) as table=>
			let
			// 1 добавляем новую колонку "Step2", в которой будут отфильтрованые строки
				from =
					Table.AddColumn(tbl, "Step2", (x)=>
						let
						//	проверяю наличие колонки. В одних отчетах подтаблицы в "Content", в других, как в БКС "Data" после разворота листов с предприятиями
						//!  сделать вариант, если грузить надо с листа один отчет, там нет колонки дата и контент
							selCol  = 	if   Table.HasColumns(tbl, "Content")
										then Table.SelectColumns(x[Content], nmsOld)
										else Table.SelectColumns(x[Data], nmsOld),

							rename  = Table.RenameColumns(selCol, List.Zip({nmsOld, nmsNew} )),
							clear   = F[Table.ClearAll](rename," #(lf)#(00A0)", false, false),

						//	делаем колонку числовой, если там не число и возникает ошибка то меняем на null
							tr      = Table.TransformColumns(clear, {columnName, try Number.From otherwise null, type number}),

						// 	удаляем errors, который иногда появляется, когда там текст или еще чего
							err     = Table.RemoveRowsWithErrors(tr, {columnName}),

							lst     = {null, kodStart, kodEnd},
							f=(x)=> x >= lst{1} and x <= lst{2},

						//	оставляем строки в диапазоне номеров в колонке "Статья_ID"
							selRow  = Table.SelectRows(err, (row)=> f(Record.Field(row, columnName) ) )
						in	selRow
					),
				
			// 2 удаляем колонку с исходнами данными при ее наличии
				to	= Table.RemoveColumns(from, {"Content","Data"}, MissingField.Ignore)
			in	to,
	//	-------------------------------------
	//* привести другие варианты обработки например баланса, не по номерам строк
	//	-------------------------------------
	//	СХЕМЫ ТАБЛИЦ (ДЛЯ ЭКСПОРТА)
	//	-------------------------------------
	form.Shema      = // 1 функция создает таблицу со схемой столбцов входящей таблицы + столбец примера из ее первой строки
		/*
			Описание:	Выводит схему столбцов одного отчета
			Пример:		= F_Form[form.Shema](fxLoadReport("opu"))	
		*/
		(x as table, optional n as number)=>
			let
				nCols = if n = null then 5 else n,
				from  = Table.FirstN(x,1),	// берем первую строку
				tbl   = Table.Transpose(Table.DemoteHeaders(from)),
				nms	  = List.Buffer(Table.ColumnNames(Table.Schema(from))),
				shema = Table.SelectColumns(Table.Schema(from), List.FirstN(nms, nCols)),	// оставляю первые 5 колонок из параметров столбцов
				join  = Table.Join(shema, {"Name"}, tbl, {"Column1"}, JoinKind.LeftOuter),
				del   = Table.RemoveColumns(join,{"Column1"}),
				ren   = Table.RenameColumns(del,{{"Column2", "Example"},{"Name", "Header"}})
			in
				ren,

	form.Shema2     = // * формирует таблицу из одной колонки: "Заголовок / ТипКолонки", например "Date / Type.Date"
		(tbl as table, txt as text)=>
			let
				// Примерные списки
				List1 = Table.ColumnNames(tbl), // имена колонок
				List2 = List.Transform(Table.Schema(tbl)[TypeName], Text.From ), // тип колонки
				
				// Проверяем, что оба списка одной длины и присваиваем List2 новое значение
				CombinedList = List.Transform(List.Zip({List1, List2}), each Text.Combine(_, " / ")),
				
				// 
				CombinedList2 = 
					if   List.Count(List1) = List.Count(List2) 
					then List.Transform(List.Zip({List1, List2}), each Text.Combine(_, " / ")) 
					else error "Списки разной длины",        
				
				// Преобразуем результат в таблицу
				ResultTable = Table.FromList(CombinedList, Splitter.SplitByNothing(), {txt}) // type table [txt = text]
			in
				ResultTable,

	form.SetType    = // 3 типизация столбцов в одной таблице по схеме в другой таблице (столбцы совпадают!) //! совместить
		(tblSource as table, tblShema as record, reportName as text)=>
			let 
				// 	функция определения типа значения по текстовому параметры. "Text.Type" => type text
					fxTypeFromText = 
						(txt as text) =>
								if txt = "Text.Type" 	 then type text
							else if txt = "Int64.Type" 	 then type number
							else if txt = "Number.Type"  then type number
							else if txt = "Date.Type" 	 then type datetime
							else if txt = "Logical.Type" then type logical
							else type any,
				
				// 	функция возвращает список пар значений для трансформации типов
					lst =
						List.Transform(
							Table.ToRecords( // список Record-ов
								Record.Field(tblShema, reportName)
									[ [Header], [TypeName] ] // таблица из содержимого Record с двумя полями
							),
							each { [Header], fxTypeFromText([TypeName]) } // пары значений Заголовок/Тип
						),

				// 	трансформация типов столбцов таблицы
					to = Table.TransformColumnTypes(tblSource, lst)
			in  to,

	form.SetType2   = // 3 типизация столбцов в одной таблице по схеме в другой таблице (столбцы НЕ совпадают!)
		/* ситуация может возникать, когда добавил столбцы, а в файл CSV еще не выгрузил */
    	(	TBL_SOURCE as table, 
			TBL_SHEMA as record, 
			REPORT_NAME as text) =>
        let 
            // функция определения типа значения по текстовому параметру. "Text.Type" => type text
				fxTypeFromText = 
					(txt as text) =>
						if      txt = "Text.Type" 	 then type text
						else if txt = "Int64.Type" 	 then type number
						else if txt = "Number.Type"  then type number
						else if txt = "Date.Type" 	 then type datetime
						else if txt = "Logical.Type" then type logical
						else type any,

            // 	получение схемы таблицы
           		schemaTable = Record.Field(TBL_SHEMA, REPORT_NAME),

            // 	создание списка столбцов и их типов из схемы
            	schemaColumns = Table.ToRecords(schemaTable [ [Header], [TypeName] ]),

            // 	создание списка типов для существующих столбцов в tblSource
				existingColumns = Table.ColumnNames(TBL_SOURCE),
				columnTypes = List.Transform(
					schemaColumns,
					each if   List.Contains(existingColumns, [Header])
						 then { [Header], fxTypeFromText([TypeName]) }
						 else null
				),
				filteredTypes = List.RemoveNulls(columnTypes),

            // 	добавление столбцов, которые отсутствуют в TBL_SHEMA, но есть в tblSource, с типом any
				missingColumns    = List.Difference(existingColumns, List.Transform(filteredTypes, each _{0})),
				additionalColumns = List.Transform(missingColumns, each {_, type any}),
				finalTypes 		  = List.Combine({filteredTypes, additionalColumns}),

            // 	трансформация столбцов
            	to = Table.TransformColumnTypes(TBL_SOURCE, finalTypes)
        in to,
	
	form.BlankTable = // # вставляем новую пустую таблицу с полями заданного отчета
		/*
			Описание: 
				если отчет в базе есть, то выводим его, если нет, то пустую таблицу с заполненной датой и возможно другими тестовыми данными
				если данных по отчету нет, то 
					1. 	грузим пустую таблицу с нужными полями отчета - если в схемах есть макет такого отчета
					2.	грузим пустую таблицу с универсальными полями отчета согласно имени запроса

			Пример:
				= Value.Metadata(form.BlankTable(DataShema, "opu"))

			Результат:
				на выходе запись с 4-мя вариантами таблиц
			
			Улучшение: 
				1 	задокументировать функцию, чтобы из проекта были видны параметры и примеры
				2 	типизацию столбцов сделать через имена полей. Фиксируем поля с датой и числами, остальное тект
					в танке можно найти по тегу "обновляем типы вручную, но пачками"

		*/
		
		(	TBL_SHEMA   as record,			// таблица с типами колонок всех отчетов
			REPORT_NAME as text 
		)=>		
		
			let 
				//	схемы таблиц с типами колонок всех отчетов (если аргумент, то TBL_SHEMA as record)
					//TBL_SHEMA = F[GetQueryValues]("DataShema"),

				// 	имена колонок конкретного отчета или эталонного отчета
					nms  = Record.Field(TBL_SHEMA, REPORT_NAME)[Header],
			
				// 	ЗНАЧЕНИЯ строк. Если полей с такими названиями нет, ошибки не будет, данные просто не будут добавлены
					
					// 	1 добавляем значение даты, так как с этой колонкой существует связь. Добавить другие связи если будут
					dataList = {
							Record.FromList( 
								List.Transform(nms, (x)=> if x = "Date" then #date(2023, 12, 1) else null), 
								nms) 
						},
					
					dataList2 = {
							Record.FromList(List.Transform(nms, (x)=> if x="Date" then #date(2023, 12, 1) else null),nms),
							Record.FromList(List.Transform(nms, (x)=> if x="Статья" then "Запасы" else null), nms)
						},	

				//	Таблица 1 - добавляем в пустую таблицу с полями отчета строку с датой. Надо УПРОСТИТЬ
					// 	добавление новых строк к существующей таблице можно через Table.Combine
					tableAddedRows = F[Table.AddedRows](dataList),

				// 	если ошибка в названии отчета, но функция применена в конкретном запросе
				// 	меняем тип всех колонок таблицы по шаблону - таблица "tableShema"
				// ? можно унивесальную функцию заделать, как в F[Table.AddedRows] в конце
					tableTyped = 
						F_Form[form.SetType2](
							tableAddedRows, 
							TBL_SHEMA, 
							REPORT_NAME
						),

				//  Таблица 2 - пустая таблица с универсальными колонками //! F_Form[nmsTech]?!
					// колонка Date обязательно, так как по ней связи ексть => добавить надо поля с корорыми есть связи
					tableAddedRows2 = F[Table.AddedRows]({
						[Company = "Нет отчета с таким кодом", Date = #date(2022, 12, 1)]
					}),
				
				//	варианты вывода:
					tableZero1 = Table.PromoteHeaders(Table.FromList({""})), // нет таблицы
					tableZero2 = Table.FromList({""}),	// пустая таблица (1 строка и 1 колонка)
					tableZero3 = tableAddedRows2,		// таблица с универсальными полями
					return 	   = try 	   tableTyped
								 otherwise tableZero3 
			in
				return meta [
					Таблица с полями отчета = tableTyped, 
					Нет таблицы 			= tableZero1, 
					Пустая таблица 			= tableZero2, 
					Универсальная таблица 	= tableZero3
				],

	// 	-------------------------------------
	// 	ПОЛЕЗНОСТИ
	// 	-------------------------------------
	fxAddTechColumn = // добавляет 4 техколонки для одиночных отчетов на 1 дату //! возможно лучше сделать запись и потом добавить запись и раскрыть ее
		/*  
			= F_Form[Table.AddTechColumn](tbl,"01/01/2024",1,"Drog","11z")
		*/
		(tbl, _date, _set, _com, _form)=>
			F[fxAddMultiColumns](tbl,
			{
				{ "Date",    	each Date.FromText(_date), type date},
				{ "DataSet", 	each _set, 	Int64.Type},
				{ "Company", 	each _com, 	type text},
				{ "form_ID",  	each _form,	type text}
				}
			)
		,
	fxRowCount 		= // добавляем колонку с числом строк в подтаблицах 
		(tbl)=> Table.AddColumn(tbl, "СountRows", each Table.RowCount([Value])),
	form.TextOffset = // делает отступ из двух пробелов в одной колонке, по условию (номер строки) в другой колонке. Но 1 Колонка
		/* 	
			Описание
				Делаем отступ из двух пробелом перед первым символом если номер строки не кратный 10 (последний символ <> 0)
			Особенности
				1	Могут быть исключения. Например 411 - статья "Паевой фонд". Это доп колонка в отчетах ПК
				2	Сделать условие, еслои колонок две. Но в эжтом примере это врядли будет
			Пример:	
				1	= F_Form[form.OffsetColumn](from, "Статья", "ID_Строка", 10)
		*/
		[ 	func = 
				(	tbl as table, 				// таблица 
					columnTarget as text,		// столбец "Статья"
					columnCriteria as text,		// столбец "Строка_ID"
					delitel as number			// челое число на которое надо делить номер строки, чтобы появился искомый остаток (10, 100)
				) => 
					Table.ReplaceValue(tbl, 
						each Record.Field(_, columnTarget),
						each if Number.Mod(Record.Field(_, columnCriteria), delitel) <> 0  
							then "    " & Record.Field(_, columnTarget)
							else          Record.Field(_, columnTarget),
						Replacer.ReplaceValue, 
						{columnTarget}
					),
			
			typ = type function (tbl as table, columnTarget as text, columnCriteria as text,delitel as number) as table 
					meta [ 	Documentation.Name            = "form.OffsetTextInColumn",
							Documentation.LongDescription = "Делаем отступ из двух пробелом перед первым символом если номер строки не кратный 10 (последний символ <> 0)" ],
				
			result = Value.ReplaceType(func,typ) 
		]  [result],
	
	getColumnNames 	= // # функция создает список имен столбцов, "Column" + номер //? только не get a setColumnNames
		/*
			Пример:	
				nms = getColumnNames({1..10})		{"Column1", "Column2"}
				nms = getColumnNames({1,2,10})		{"Column1", "Column2", "Column10"}
			
			Варианты: 
				nms = List.Generate(()=>1, (x)=>x<=10, (x)=>x+1, (x)=>"Column" & Text.From(x)),
			
			Улучшения:
				1 добавить первый элемент
					= {"Статья"} & List.Transform({1..10}, (x)=>"Column" & Number.ToText(x))
		*/
		(numbers as list) as list =>
			List.Transform(numbers, each "Column" & Number.ToText(_)),

	formTrans 		= // в файле Excel переводит список отчетов в код для вставки сюда. //!непонятно + ошибка
		let
			from = Excel.CurrentWorkbook(){[Name="Спр_ФормОтчетности"]}[Content],
			typ  = F[Type.AllColumnText](from),
			nms  = Table.ColumnNames(typ),
			nmsF = List.First(nms),
			nmsL = List.Last(nms),
			ReplaceNulls = Table.ReplaceValue(typ, null, "", Replacer.ReplaceValue, nms),
			// добавляю кавычки и двойную табуляцию
			Add1 = Table.TransformColumns(ReplaceNulls, List.Transform(nms, each {_, each """" & _ & ""","})), // #(tab)
			Add2 = Table.TransformColumns(Add1, {nmsF, each "{" & _ & ""}),		// добавляю {
			Add3 = Table.TransformColumns(Add2, {nmsL, each _ & "}" }),			// добавляю }
			del  = Table.ReplaceValue(Add3,",","",Replacer.ReplaceText,{nmsL}),	// убираю запятую в последней колонк
			Add4 = Table.TransformColumns(del, {nmsL, each _ & "," }),			// добавляю }
	 		
			// Удаление последнего символа в последней строке последней колонки
				delLastChar = [
					// 	Добавление индекса
						IndexedTable = Table.AddIndexColumn(Add4, "Index", 1, 1, Int64.Type),
					
					// 	Определение индекса последней строки
						LastRowIndex = Table.RowCount(IndexedTable),

					// 	Удаление последнего символа в последней строке последней колонки
						ModifiedRows = Table.TransformRows(
							IndexedTable,
							each if [Index] = LastRowIndex 
								then Record.TransformFields(_, {nmsL, each Text.Start(_, Text.Length(_) - 1)})
								else _
						),

					// 	Список записей в таблицу
						FromRecords = Table.FromRecords(ModifiedRows),

					// 	Удаление временного индекса
    					FinalTable = Table.RemoveColumns(FromRecords, "Index")	
					
					]  [FinalTable],
    		
			// 	установка запятой на нужной позиции, т.к. tab не помогает при вставке для выравнивания
				AddComma = let 
					lst = {nmsF,nms{1},nms{2},nms{3},nms{4}},
					del = Table.ReplaceValue(delLastChar,",","",Replacer.ReplaceText,lst),
					f=(x)=>each Text.Insert(Text.PadEnd(_, x, " "), x, ","), // Дополнение пробелами
					to  = 	Table.TransformColumns(del, { 
							{nmsF,   f(10)},
							{nms{1}, f(6)},
							{nms{2}, f(11)},
							{nms{3}, f(52)},
							{nms{4}, f(8)}
						} )
				in to
		in
			AddComma,

	GetName  		= // text  возвращает имя DataSeta
		/*
			Описание: 	получаем имя датасета (Key) из метаданных Parameters
			Результат:	имя датасета, активное или любое из базы
			Пример:		= F_Form[GetName]("BKS")
		*/
		(optional Key as text) =>
		let
			from 	  = Parameters,
			Param	  = "DataSetName",
			Metadata  = Value.Metadata(from),
			RecordVal = Record.FieldOrDefault(Metadata, Key, null),
			
			// проверка наличия датасета с таким именем
			DataBase  = 
				if 	 RecordVal <> null and Record.HasFields(RecordVal, Param)  // проверка наличия такого параметра
				then Record.Field(RecordVal, Param)
				else "нет датасета с таким именем",	

			// возвращаем текстовое имя датасета, если он есть в параметрах
			ActualSet =	if Key = null 					// 	если аргумент не задан
						then Parameters[DataSetName] 	//	то активный сет
						else DataBase					//	иначе нужный сет	
		in
			ActualSet,

	GetPath  		= // text  возвращает путь к таблице с файлами XLSX в папке по метаданным Parameters
		/*
			Описание: 	Получаем метаданных Parameters значение отдельного параметра.
			Результат:	путь к источнику данных
			Пример:		= F_Form[GetPath]("BKS")
		*/
		(Key as text) =>
		let
			from 	  = Parameters,
			Param	  = "pathXLSX",
			Metadata  = Value.Metadata(from),
			RecordVal = Record.FieldOrDefault(Metadata, Key, null),
			DataBase  = 
				if 	 RecordVal <> null and Record.HasFields(RecordVal, Param)  // проверка наличия такого параметра
				then Record.Field(RecordVal, Param)
				else null
		in
			DataBase,

	GetFiles 		= // table возвращает таблицу со всеми файлами XLSX в папке по метаданным Parameters
		/*
			Описание:
				Получаем метаданных Parameters значение отдельного параметра.
				Это связано с тем, что в Parameters возвращают только параметры активного датасета
			
			Результат:
				путь к источнику данных
			
			= Value.Metadata(Parameters)[BKS][FileaXLS]
			= F_Form[GetFiles]("BKS")
		*/
		(Key as text) as table =>
		let
			from 	  = Parameters,
			Param	  = "FilesXLSX",
			Metadata  = Value.Metadata(from),
			RecordVal = Record.FieldOrDefault(Metadata, Key, null),
			DataBase  = 
				if 	 RecordVal <> null and Record.HasFields(RecordVal, Param) 
				then Record.Field(RecordVal, Param)
				else null
		in
			DataBase,
	
	Q 				= // добавляет заданное число неразрывных пробелов (для отступов в подстатьях)
		(numSpaces as number) as text => 
			Text.Repeat( Character.FromNumber(160), numSpaces ),
		
	// 	-------------------------------------
	// 	ТЕСТОВЫЕ ФУНКЦИИ
	// 	-------------------------------------
	LoadAllCSV 	       = // массовая загрузка из папки. Сделать функцию + путь
		let
			Источник = Folder.Files("G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\4. CSV выгрузка\Белкоопсоюз"),
			csv =(x)=> let
				from = Csv.Document(x,[Delimiter=";", Encoding=65001, QuoteStyle=QuoteStyle.None]),
				to   = Table.PromoteHeaders(from, [PromoteAllScalars=true])
				in	to,
			add = Table.AddColumn(Источник, "Преобразовать файл", each csv([Content])),
			ren = Table.RenameColumns(add, {"Name", "Source.Name"}),
			sel = Table.SelectColumns(ren, {"Source.Name", "Преобразовать файл"}),
			row = Table.SelectRows(sel, each [Source.Name] <> "# S_4TORG.csv"),
			nms = List.Distinct(List.Combine(List.Transform(Table.Column(row, "Преобразовать файл"), each Table.ColumnNames(_)))),
			exp = Table.ExpandTableColumn(row, "Преобразовать файл", nms),
			typ = (tbl as table) =>
				let	nms = Table.ColumnNames(tbl),
					lstDate   = {"Date"},
					lstNumber = {"Строка_ID", "Значение","Сумма","Value", "DataSet",
								"Index",
								"На начало","На конец",
								"Текущий период","Прошлый период"},

					dlst = List.Select(nms,(i)=>List.Contains(lstDate,   i, (x,y)=>Text.Contains(y,x))),
					nlst = List.Select(nms,(i)=>List.Contains(lstNumber, i, (x,y)=>Text.Contains(y,x))),
					tlst = List.Difference(nms, dlst & nlst),

					f=(x,y)=>List.Transform(x,(i)=>{i,y}),
					transflst = f(dlst, DateTime.Type) &
								f(nlst, Number.Type) & 
								f(tlst, Text.Type),
					
					to = Table.TransformColumnTypes(tbl, transflst)	
				in  to,

			to  = typ(exp)
		in
			to,
	
	form.FormatText    = // добавляет 0 в 8-ю и 9-ю позицию, если их там нет
			(input as text) as text =>
			// 
			// Result: 
				let
					// Подсчитываем количество символов без точек
					numChars = Text.Length(Text.Remove(input, ".")),
					// Проверяем, нужно ли добавлять нули
					result =
						try List.Buffer(
							if numChars = 7 
							then input & "00"
							else if numChars = 8    // Если символов 8, добавляем 0
								then input & "0"
								else input           // Если символов не 7 и не 8, оставляем как есть
						) otherwise input
				in
					result,
	
	form.TransformText = // ставит разделитель на заданные позиции (для кодов классификаторов)	
			/*
				В статотчете 4torg пробовал. До уцонка не довел. Для определения level хотел приспособить
				Пример:	= TransformText("47.21.11.100", {2,3,4,5,6,7,8,9})
						= F[TransformText]("47.21.11.100", {2,3,4,5,6,7,8,9})
				Result: 4|7|2|1|1|1|1|0|0
			*/
			(input as text, positions as list, optional char as text) as text =>
			let
				textRemove  = Text.Remove(input, "."),
				delimeter   = "|",
				charList    = List.Buffer(Text.ToList(textRemove)),
				LitPosition = List.Buffer(positions),

				transformedList = List.Combine( // Добавление "/" на указанные позиции
					List.Transform(
						{0..List.Count(charList) - 1},
						each if List.Contains(positions, _ + 1)
							then {delimeter, charList{_}}
							else {charList{_}}
					)
				),
				transformedText = Text.Combine(transformedList, "") // Преобразование списка символов обратно в строку
			in  transformedText,
	func = "",

//	ПЕРЕМЕННЫЕ
	nmsTech  = {"Company", "DataSet", "form_ID"},				// названия технических колонок
	nmsTech2 = {"Company", "DataSet", "form_ID"} & {"Date"},	// названия технических колонок + дата

// 	СПРАВОЧНИКИ
	// общие
	spr.FAQ    = [ // код для копирования в проект только справочников
		from = Table.SelectRows(Record.ToTable(F_Form), each Text.StartsWith([Name], "spr"))
		],
	spr.REPORT = [ // Формы отчетности и др.
		Спр_form = [], 			// сделать справочник форм (словарь по части имени из имени файла) для отчетов других организаций
		Спр_ФормыОтчетностиБКС = // Справочники форм отчетности для потребкооперации
			/* 
				Описание
					берем символы до 12. Это код отчета в названии файла. Можно потом добавлять 12 для указаня что месячная отчетность
					! в функции Connect для form_ID_BKS берутся 4!!! первые знака. Плоэтому надо не 07 а 0712 вводить код
				Улучшения 
					
			*/
			[ Источник = 
				#table(type table
					[	
						ID        		= text, 	// символы до "ГОМЕЛЬСКИЙ" + 4 символа левее (1223). До этого это код отчета!
						form_ID_BKS	 	= text, 	// 
						form_ExportName = text,		// 
						form_NameFull  	= text		// 
					],
					{ // ID    		form_ID 	form_Export 	form_NameFull
						{"01"   ,	"bal",		"f_bal"    	,	"Бухгалтерский баланс"                              },
						{"02"   ,	"opu",		"f_opu"    	,	"Отчет и прибылях и убытках"                        },
						{"03"   ,	"v_3",		"v_3bks"   	,	"Отчет о доходах расходах и финансовых результатах" },
						{"04"   ,	"odds",		"f_odds"   	,	"Отчет о движении денежных средств"                 },
						{"05"   ,	"pril",		"f_pril"   	,	"Приложение к бухгалтерскому Балансу"               },
						{"06"   ,	"osr",		"s_ras1"   	,	"Отчет о состоянии расчетов"                        },
						{"07"   ,	"v_7",		"v_ospp"   	,	"Отчет о себестоимости произведенной продукции"     },
						{"11"   ,	"v_11",		"v_osks"   	,	"Отчет о состоянии кооперативной собственности"     },
						{"12"   ,	"ofr",		"s_fr1"    	,	"Отчет о финансовых результатах"                    },
						{"16"   ,	"f_16",		"s_ocf"    	,	"Отчет о целевом финансировании"                    },
						{"18"   ,	"f_18",		"f_18"     	,	"Приложение к отчету ОПУ, ФР по отраслям"           },
						{"22"   ,	"f_22",		"f_22"     	,	"Отчет по инвестиционному фонду"		            },
						{"30"   ,	"f_30",		"f_osk"    	,	"Отчет об изменении капитала"                       },
						
						{"w1"   ,	"w_1",		"w_070-3"  	,	"Прочие доходы по текущей деятельности"				},
						{"w2"   ,	"w_2",		"w_080-3"  	,	"Прочие расходы по текущей деятельности"			},
						{"w3"   ,	"w_3",		"w_100-3"  	,	"Доходы по инвестиционной деятельности"				},
						{"w4"   ,	"w_4",		"w_110-3"  	,	"Расходы по инвестиционной деятельности"			},
						{"w5"   ,	"w_5",		"w_120-3"  	,	"Доходы по финансовой деятельности"					},
						{"w6"   ,	"w_6",		"w_130-3"	,	"Расходы по финансовой деятельности"				},

						{"w16"   ,	"w_16",		"w_lik"    	,	"Меры ликвидации внесистемной просроченной ДЗ"		},
						{"w17"   ,	"w_17",		"w_rbp"    	,	"Расходы будущих периодов"							},
						{"w18"   ,	"w_18",		"w_dbp"    	,	"Доходы будущих периодов"							},
						{"w24"   ,	"w_24",		"w_kv"    	,	"Расшифровка капитальных вложений по отраслям"		},
						
						{"w35"   ,	"w_35",		"w_35"    	,	"Информация о исправлении ошибок прошлых лет"		},
						{"w36"   ,	"w_36",		"w_36"    	,	"Отчет о раходах на потребление"					},

						{"w240"   ,	"w_240",	"w_240"    	,	"Отчет о ДЗ за товары, работы, услуги"				},
						{"W340"   ,	"W_340",	"W_340"    	,	"Отчет о движении основных средств"					},
						
						{"W30"   ,	"W_30",		"W_30"    	,	"Отчет о составе и движении работников бухгалтерии" },
						{"06R"	,	"osr",		"s_ras2"   	,	"Отчет о состоянии расчетов new"                    }
					}
				)
			] [Источник],
		
		Спр_ФормыОтчетности = [ // Справочники форм отчетности для потребкооперации
			/* 
				! в функции Connect для form_ID_BKS берутся 4!!! первые знака. Плоэтому надо не 07 а 0712 вводить код
				Улучшения 
					form_ID_BKS - сделать можно по другим дасетам
					добработать надо. Когда и где используются поля. Справочно только или связать есть
			*/
			Источник = #table(type table
			[	
				form_ID        	= text, 	// 1 основной код. Много где используется. Значения не уникальные!
				form_ID_BKS	 	= text, 	//!2 первые символы в названи файлов. Могут быть разные. По ним идет поиск-замена.
											// используются ли они из справочника или только справочно приведены?!
				form_ExportName = text,		// 3
				form_NameFull  	= text,		// 4
				form_NameShort 	= text, 	// 5
				form_Descript  	= text,		// 6
				form_Index 		= number	// 7
			],
			{ // form_ID    form_ID_BKS form_Export 	form_NameFull											form_NameShort
				{"bal"    ,	"0112",		"f_bal"    ,	"Баланс"                                            ,	"ББ"    ,	"",	"1"},
				{"opu"    ,	"0212",		"f_opu"    ,	"Отчет и прибылях и убытках"                        ,	"OPU_"  ,	"",	"2"},
				{"odds"   ,	"0412",		"f_odds"   ,	"Отчет о движении денежных средств"                 ,	"ODDS_" ,	"",	"3"},
				{"osk"    ,	"3012",		"f_osk"    ,	"Отчет об изменении капитала"                       ,	"ОИК"   ,	"",	"4"},
				{"pril"   ,	"0512",		"f_pril"   ,	"Приложение к бухгалтерскому Балансу"               ,	"ПББ"   ,	"",	"5"},
				{"ocf"    ,	"",			"s_ocf"    ,	"Отчет о целевом финансировании"                    ,	"ОСР"   ,	"",	"6"},
				{"osr"    ,	"0612",		"s_ras1"   ,	"Отчет о состоянии расчетов"                        ,	"ОСР"   ,	"",	"6"},
				{"osr"    ,	"06R1",		"s_ras2"   ,	"Отчет о состоянии расчетов new"                    ,	"ОСР"   ,	"",	"7"},
				{"ofr"    ,	"1212",		"s_fr1"    ,	"Отчет о финансовых результатах"                    ,	"ОФР"   ,	"",	"8"},
				{"ofr"    ,	"12f1",		"s_fr2"    ,	"Отчет о финансовых результатах new?!"              ,	"ОФР"   ,	"",	"9"},
				{"v_3"    ,	"0312",		"v_3bks"   ,	"Отчет о доходах расходах и финансовых результатах" ,	"ОДРФР" ,	"",	"10"},
				{"v_7"    ,	"0712",		"v_ospp"   ,	"Отчет о себестоимости произведенной продукции"     ,	"ОСПП"  ,	"",	"11"},
				{"v_12po" , "po"  ,		"v_12po"   ,	"Отчет о производстве продукции (ОПО) 12-п" 		,	"ОПП1"  ,	"",	"12"},
				{"v_12pr" , "pr"  ,		"v_12pr"   ,	"Отчет о производстве продукции (РПО) 12-п" 		,	"ОПП2"  ,	"",	"12"},
				{"s_1p1"  ,	"pg"  ,		"s_1pg"    ,	"Отчет 1-п (натура годовая) ОПО"                    ,	"ОСПП"  ,	"",	"13"},
				{"s_1p2"  ,	"pq"  ,		"s_1pq"    ,	"Отчет 1-п (натура годовая) райпо"                  ,	"ОСПП"  ,	"",	"14"},
				{"s_1torg",	"7R"  ,		"s_1torg"  ,	"форма 1-торг (розница годовая) ОПО"                ,	"1торг" ,	"",	"15"},
				{"s_4torg",	"??"  ,		"s_1torg"  ,	"форма 4-торг (продажи) ОПО"  			            ,	"4торг" ,	"",	"16"},
				{"s_11z",	"11z" ,		"s_11z"    ,	"форма 11-заготовки (квартальная) ОПО"              ,	"1торг" ,	"нараст",	"17"}

			}
			)
			] [Источник],

		Спр_Год =  // Справочники периодов
			Table.FromList({2017..2024}, Splitter.SplitByNothing(), type table [Год = number]),

		Спр_Область = // Справочники областных городов
			let Источник = Table.Buffer(#table(type table [
						Область_ID 	= number,
						Область 	= text	],
					{
						{1,"Брест"},
						{2,"Витебск"},
						{3,"Гомель"},
						{4,"Гродно"},
						{5,"Минск"},
						{6,"Могилев"}
					}))
			in Источник
		],

	spr.BKS    = [ // Справочник предприятий + другие по БКС
		Спр_ГруппаПредприятий =
			/* 
				Group_ID	НазваниеГруппы
				1			Филиалы
				2			УП
			*/
			let
				Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw5MKOC7uBeMOF3Re7lWJ1opWMQMKLL8xXio0FAA==", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Group_ID = _t, НазваниеГруппы = _t]),
				#"Измененный тип" = Table.TransformColumnTypes(Источник,{{"Group_ID", Int64.Type}, {"НазваниеГруппы", type text}})
			in
				#"Измененный тип",

		Спр_Предприятия  =
			/*
				Предприятие_ID	Предприятие_НазваниеИсх	Предприятие_Группа1
				1 	Брестский ОПС	1
				2	Витебский ОПО	1
				3	Гомельский ОПО	1
				4	Гродненское ОПО	1
				5	Минский ОПО	1
				6	Могилевский ОПС	1
				20	УП "Белкоопмех"	2
				24	ЧУП "Калинковичское зв-во БКС"	2
				26	Молодечненское зверохозяйство	2
				27	ЧУП "Пинское зв-во БКС"	2
				29	ЧПУП "Борисовский КТМ "БКС"	2
				33	ЧУП "З-д Белкооппрогресс "БКС"	2
				34	УП "Пищевой комбинат Белкоопсою...	2
				39	УП "ИнтерПродТрейд"	2
				40	ЧУП "Белкоопвнешторг"	2
				41	ЧТПУП "Поречье БКС"	2
				44	ЧУП "ВЦ Белкоопсоюза"	2
				49	УО "Гомельский БТЭУ"	2
				50	УО "Барановичский ТК"	2
			*/
			
			let	Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("jVPbasJAFPyVkOcKXmKL36L+iwbEFqRF+9ASxGCE9qmg0eCtWX9h9o86Z6NmEyrta3Zmzsyck3bbrbl3Lsa6h0T3ta/7OGCHvYMZQsz5VnO7d223LqgJdtpHglURNbuiGoJ6hcI3UUc9uoXzDI4zFTZIiU0NUiEpI5uCnFIjvaV1nyEU1nw9Uiu+maEq0AVCp8PE4tDMVDiJXz3ouHyvZ1Dj8PMCDrAkmCYMIWYNw9zwFnGF35RDzQBzW+Xq7WiSJnpYTksyP7MIPeCHrX7BXtYgernMQ8FMeG3jr/GtjBfmmRVH7chVdksBIkzNc5HfaBTmvlWwcUq9ncwO1+fr6V9EcgnPqpzG9ZMsiBSOPZg7WZlal9ovSYtJ/cx0S9tRy5J7Zwu+KS88X1IkPrDHxqJ41UKIov1Y1qEfKSPFrG1aLaNFeXuhAXGHeiS9l9vyiiczwcd/EnnnRDPh/PrnjOnhCwuL06xanDG76/E808JhGmaE4MLq/gA=", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Предприятие_ID = _t, Предприятие_НазваниеИсх = _t, Предприятие_Группа1 = _t]),
				typed    = Table.TransformColumnTypes(Источник,{
					{"Предприятие_ID", Int64.Type}, 
					{"Предприятие_Группа1", Int64.Type}
				})
			in 	typed,

		Спр_Предприятия2 = // полный список, копируется их файла Excel
			let
				Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("lVlrUhtHEL7Kln4bl7R63yB3AN+BI4BIsKtwmSIhISWIhExcSf4JPWCRxHKF2RulH7MzO7urmaZcNi40X3dPT/fX34wODxutxoeGuspO1Do7zUbZqdqoRL1EaqKmagaf4efNxqcPh40YV/6qkmyk1uqxtPID/LhSt/DfiQNqI+g3laodgLbZVyGsQzCIKlUr9QrQVwKmah3tQXQRcQemX4UuegxI1RIWb8HFwrv5IS7/B37/oKbREaRMpRBeApi0iByr7+qOPh6r2VEDQLHNBWVQ42tzcgPoSRnVZs8ad3OgVhGYByQlJFVv6o0SteRDjGp9d5oF3y56gQnOvsCx4o6W9fhW1Gxyqag5rJoDBvYN+/9sDoZ+/QIGU1jXpuzlyJiRa4xQPZuMBVDtamkGEFw2GF12jicqjY+K55qziBhTRgFcH3F/gh+dDwlmYDDvi5EKcAwfY9XgiZ0HMa0mY1Jo1xPhnlp0yrfZGfhIpH5ixlxmv7yjKlpt7sA51CL3ksATnfBUuhc612l2Art5oqoQYIgXZtgNEJfQz8Bp0SqbQgQXVCEbaisHHDc94Hl2GUGK7tVVDXDoAEslpZKITh46nKxhbb9UjTAneULPaYJMUAqrRgbNCjmWW12T41hNa+AUw98HaovYsUNqu6NGVNNhS5V6zmPQMQb3dJqLZ7f1sQ2Ksd3AFscR/IO1gXn6qtYYYLVd/AEO3TOfumNrU81AseIDpltuMmlXUU0TBMyU6oImBvb2Wj3pSGEUL2jP82rahu3aKKpt5YkitiMHxhOR3sahF+qsn2l6J8giGh0bdJwPkDms38ISD0FbVJt1DlQw0BlqjxXvU+azyyNoC9l+xIaXeOxpkaT1jhBFw+eaVn/R5BmMjijjFvb0VpIePhSPkVv4MIWOdtXY/vh4kKAig31J1mtqT5F+zFkFY6N83xMKpZgoMqZ4qii7m6Anyvd3qt80JBEsivjjAU5oLqzadsvpPFdw4zzQXLOgCj39CGU2h2OxTWgtuT081aGfv9tO37FTLm6aUVM46XEJ2jYNTPlZhpqobVC6cc+gGeYHOBGo/61Ir825hWsZWBXYXqfdvJ8ekStBEAs8USX9bnwEVlMe/6ADFTGZRQ4YiRT86K88izGCkfi25NGPNcIRluyyS5E/LRx1xj/7GcmiWDqibdHZcr9PoPRGEBeelOx0ueenUOsjksIbR5l4kX1mmLXZUyDCgWYkKHjoeedi6fUzZFbCfhzRSS8Nc4awrCF/oMSjnvbMD4Npuxqk3C7U1Q/Ad1Z5WmiLqVBD99VYjXi1JlxuwrYDIUGSJylITLO+44rdXNmiyNzBHyR+HBpPZAVH/QoDqNphaTiBPDMn0shNgPcWH1l47bjBIPgRKbtSPe9RLGy/U1Asx6yjdUI8rdAxUP3Akuq7+SK7oKEmQGrJkh5TGgOVbWHMk8elR5YgTN+WL3WZCp1RA95wOklHCr3ld2zU2lJfA01fxzQ1AixrYcNc56xE65kl77RikSaCePInqgsopmMccjIclccEfR2zEhdVFd+yJzBALZEHMZ2cihISzqkwgUzOM6oKpKGdDKXl2A7Pa2k0XBBGhfEfTGqrxUKYnvuWNkMZDFlMNIE9q4RTOydlWCNs2FK3eDPhe653JncNSj+ElZ8OAyjd3vxM+WIQPkjHckkKrFgMzwfTGuiZ9iWH9bjftoXXIt/yvm5PIPALRzv5QNzTdI19EwGGubLAmhJlTV90sm+khUVbt7cc2fJYUwady4pEhXuh8oH5wQzaeFUY1CEQVcJfOI3oGvskAumHs7OaBz0frG86mcW6GDhgIL8NLJ0bow/Gkgknf351pp+iPbJmyl8kHqFwi93oA9KRP0CTrKQnF2u1g1218N+fDabrKqQZ3rrM3S2/rOmvKgw/WbTLdPmj2oZ4GROsX/trkO5tz5a2+b6Asfql8AyPKruo2uk199jR37DcO8RqYeXAt4VBB3udqpsaUMcBlROtPc5qod1SmG5rpuizJsz+cM/p7HnuNcBB7NvfGJ8Ya0DV0yxSZ/2rYcWIu1OX5oQmSrve03kCY73CCIV9YLq8V9OeQelvhS6YmiQI/RyAJ3QqWa+f8FKq8L0XPruept81XSsDl3OL0SMwoeSIdjHMZ2xSuFT6MfYajy/63yRPABbbySeV+/1oAMVfw5IgTsKPrhZHOQRxC77e3hOlnjjkhwcHVqMESSPnBzFhrlQDCF38c2Ed5V/srEsPDgEUdcW/sG4rRDhfK1/l71da8RAdaa7nOyX+4pyocZSdmK401kpjp1oANH5mJV1s4S5TlcOpkdQMjeN+1Go2oc7NX7uriMQvSHKKHuTFnApsXYhfP1Z2W1EP0YNmFJOVViE3QZLOjbTrQqlqff9LLljpVq1UtXXQyrBqxVW1aMFnolO3nTtpAJ26AOpk0347n/4H", BinaryEncoding.Base64), Compression.Deflate)), let _t = ((type nullable text) meta [Serialized.Text = true]) in type table [Код = _t, Название_исходное = _t, Группа = _t, Область = _t]),
				typ = Table.TransformColumnTypes(Источник,{
						{"Код", Int64.Type}, 
						{"Название_исходное", type text}, 
						{"Группа", Int64.Type}, 
						{"Область", Int64.Type}
					})
			in	typ,
		
		Спр_Райпо = 
			#table(type table
				[
					Код = nullable Int64.Type, 
					Название_исходное = nullable text, 
					Группа = nullable Int64.Type, 
					Область = nullable Int64.Type
				],
				{
					{1             ,"Брестский ОПС"                             ,1,0},
					{2              ,"Витебский ОПС,ОБЛПО"                       ,1,0},
					{3              ,"Гомельский ОПС,ОБЛПО"                      ,1,0},
					{4              ,"Гродненское ОБЛПО"                         ,1,0},
					{5              ,"Минский ОПС,ОБЛПО"                         ,1,0},
					{6              ,"Могилевский ОПС"                           ,1,0},
					{29             ,"ЧПУП ""Борисовский КТМ""БКС"""             ,2,0},
					{32             ,"УП ""Гомельский ЗТО"""                     ,2,0},
					{33             ,"ЧУП ""З-д Белкооппрогрес БКС"""            ,2,0},
					{40             ,"УП ""Белкоопвнешторг БКС"""                ,2,0},
					{1001           ,"Барановичское райпо"                       ,3,1},
					{1002           ,"Березовское райпо"                         ,3,1},
					{1003           ,"Брестское райпо"                           ,3,1},
					{1004           ,"Ганцевичское райпо"                        ,3,1},
					{1005           ,"Дрогичинское райпо"                        ,3,1},
					{1007           ,"Ивановское райпо"                          ,3,1},
					{1008           ,"Ивацевичское райпо"                        ,3,1},
					{1009           ,"Каменецкое райпо"                          ,3,1},
					{1010           ,"Кобринское райпо"                          ,3,1},
					{1011           ,"Лунинецкое райпо"                          ,3,1},
					{1012           ,"Ляховичское райпо"                         ,3,1},
					{1013           ,"Малорицкое райпо"                          ,3,1},
					{1014           ,"Пинское райпо"                             ,3,1},
					{1015           ,"Пружанское райпо"                          ,3,1},
					{1016           ,"Столинское райпо"                          ,3,1},
					{1018           ,"ЧУП ""Брестский рынок"""                   ,3,1},
					{1020           ,"ЧУП ""Брестская МРБ"""                     ,3,1},
					{1029           ,"ЧУП ""Ивацевичи Коопстрой"""               ,3,1},
					{1032           ,"ЧУП ""Брестский кооптранс"""               ,3,1},
					{1080           ,"ЧПУП ""Березовский ККП"""                  ,3,1},
					{1082           ,"Ф-л ""Кооппром"" Ивановского райпо"        ,3,1},
					{1084           ,"Ф-л Ивацевичского райпо ""ККП"""           ,3,1},
					{1088           ,"Ф-л ""ЗПК Полесье"" Лунинецкого райпо"     ,3,1},
					{1090           ,"ЧУП ""Пинский кооппром"" Пинского райпо"   ,3,1},
					{1091           ,"Ф-л ""ККП"" Пружанского райпо"             ,3,1},
					{1092           ,"ЧУП ""Беловежские вина"""                  ,3,1},
					{1093           ,"Ф-л ""ККП"" Столинского райпо"             ,3,1},
					{2001           ,"Бешенковичский филиал"                     ,3,2},
					{2002           ,"Браславское райпо"                         ,3,2},
					{2003           ,"Верхнедвинский филиал"                     ,3,2},
					{2005           ,"Глубокское райпо"                          ,3,2},
					{2006           ,"Городокское райпо"                         ,3,2},
					{2007           ,"Докшицкий филиал"                          ,3,2},
					{2009           ,"Лепельский филиал"                         ,3,2},
					{2010           ,"Лиозненское райпо"                         ,3,2},
					{2011           ,"Миорское райпо"                            ,3,2},
					{2014           ,"Поставский филиал"                         ,3,2},
					{2015           ,"Россонское райпо"                          ,3,2},
					{2016           ,"Сенненский филиал"                         ,3,2},
					{2017           ,"Толочинское райпо"                         ,3,2},
					{2018           ,"Ушачский филиал"                           ,3,2},
					{2031           ,"ЧУП ""Витебская универс.база"""            ,3,2},
					{2032           ,"ЧУП ""Полоцкая универс.база"""             ,3,2},
					{2037           ,"ЧУП ""Глубокская ПМК"""                    ,3,2},
					{3001           ,"Брагинский филиал"                         ,3,3},
					{3002           ,"Буда-Кошелевское райпо"                    ,3,3},
					{3004           ,"Гомельский филиал"                         ,3,3},
					{3005           ,"Добрушское райпо"                          ,3,3},
					{3006           ,"Ельское райпо"                             ,3,3},
					{3007           ,"Житковичский филиал"                       ,3,3},
					{3008           ,"Жлобинское райпо"                          ,3,3},
					{3009           ,"Калинковичское райпо"                      ,3,3},
					{3010           ,"Кормянское райпо"                          ,3,3},
					{3011           ,"Лельчицкий филиал"                         ,3,3},
					{3012           ,"Лоевское райпо"                            ,3,3},
					{3015           ,"Октябрьский филиал"                        ,3,3},
					{3016           ,"Петриковский филиал"                       ,3,3},
					{3017           ,"Речицкое райпо"                            ,3,3},
					{3018           ,"Рогачевский филиал"                        ,3,3},
					{3019           ,"Светлогорский филиал"                      ,3,3},
					{3020           ,"Хойникское райпо"                          ,3,3},
					{3030           ,"ЧУП ""Гомельская УТБ"""                    ,3,3},
					{3031           ,"ТУП ""Калинковичская МРБ"""                ,3,3},
					{3032           ,"ЧУП ""Донаприс"""                          ,3,3},
					{3049           ,"ЧУП ""Коопреммонтажналадка"""              ,3,3},
					{3084           ,"ОСП ""Произв. комбинат"" Лельчицкого райпо",3,3},
					{4001           ,"Беpестовицкий филиал"                      ,3,4},
					{4002           ,"Волковысский филиал"                       ,3,4},
					{4003           ,"Воpоновский филиал"                        ,3,4},
					{4004           ,"Гpодненский филиал"                        ,3,4},
					{4005           ,"Дятловский филиал"                         ,3,4},
					{4006           ,"Зельвенский филиал"                        ,3,4},
					{4007           ,"Ивьевский филиал"                          ,3,4},
					{4008           ,"Коpеличский филиал"                        ,3,4},
					{4009           ,"Лидский филиал"                            ,3,4},
					{4010           ,"Мостовский филиал"                         ,3,4},
					{4011           ,"Hовогpудский филиал"                       ,3,4},
					{4012           ,"Остpовецкий филиал"                        ,3,4},
					{4013           ,"Ошмянский филиал"                          ,3,4},
					{4014           ,"Свислочский филиал"                        ,3,4},
					{4015           ,"Слонимский филиал"                         ,3,4},
					{4016           ,"Смоpгонский филиал"                        ,3,4},
					{4017           ,"Щучинский филиал"                          ,3,4},
					{4063           ,"ЧУП ""Специализиpованная ПМК"""            ,3,4},
					{5001           ,"Березинское райпо"                         ,3,5},
					{5002           ,"Борисовское райпо"                         ,3,5},
					{5003           ,"Вилейское райпо"                           ,3,5},
					{5004           ,"Воложинское райпо"                         ,3,5},
					{5005           ,"Дзержинское райпо"                         ,3,5},
					{5006           ,"Клецкое райпо"                             ,3,5},
					{5007           ,"Копыльское райпо"                          ,3,5},
					{5008           ,"Крупское райпо"                            ,3,5},
					{5009           ,"Логойское райпо"                           ,3,5},
					{5010           ,"Любанское райпо"                           ,3,5},
					{5011           ,"Минское райпо"                             ,3,5},
					{5012           ,"Молодечненское райпо"                      ,3,5},
					{5013           ,"Мядельское райпо"                          ,3,5},
					{5014           ,"Несвижское райпо"                          ,3,5},
					{5015           ,"Пуховичское райпо"                         ,3,5},
					{5017           ,"Смолевичское райпо"                        ,3,5},
					{5018           ,"Солигорское райпо"                         ,3,5},
					{5019           ,"Стародорожское райпо"                      ,3,5},
					{5020           ,"Столбцовское райпо"                        ,3,5},
					{5021           ,"Узденское райпо"                           ,3,5},
					{5022           ,"Червенское райпо"                          ,3,5},
					{5052           ,"ЧУП ""Слуцкая база ОПС"""                  ,3,5},
					{5053           ,"ЧУП ""Коопкниготорг"""                     ,3,5},
					{5057           ,"ЧУП ""Минсккоопвторресурсы"""              ,3,5},
					{5060           ,"ЧУП ""Минский ОРМК"""                      ,3,5},
					{5063           ,"ЧУП ""Клецкий ППЗ"""                       ,3,5},
					{5064           ,"ЧУП ""Червенский ОСЗ"""                    ,3,5},
					{5065           ,"ЧУП ""Молодечно ПК"""                      ,3,5},
					{5079           ,"ЧУП ""Слуцкий рынок"""                     ,3,5},
					{5082           ,"ЧУП ""Клецкий КЗП"""                       ,3,5},
					{5083           ,"ЧУП ""Копыльский кооппром"""               ,3,5},
					{5085           ,"ЧУП ""Любанский кооппром"""                ,3,5},
					{5089           ,"ЧУП ""Стародорожский кооппром"""           ,3,5},
					{6001           ,"Белыничское райпо"                         ,3,6},
					{6003           ,"Быховское райпо"                           ,3,6},
					{6004           ,"Глусское райпо"                            ,3,6},
					{6005           ,"Горецкое райпо"                            ,3,6},
					{6006           ,"Дрибинское райпо"                          ,3,6},
					{6007           ,"Кировское райпо"                           ,3,6},
					{6009           ,"Кличевское райпо"                          ,3,6},
					{6010           ,"Костюковичское райпо"                      ,3,6},
					{6014           ,"Могилевское райпо"                         ,3,6},
					{6015           ,"Мстиславское райпо"                        ,3,6},
					{6016           ,"Осиповичское райпо"                        ,3,6},
					{6017           ,"Славгородское райпо"                       ,3,6},
					{6018           ,"Хотимское райпо"                           ,3,6},
					{6019           ,"Чаусское райпо"                            ,3,6},
					{6020           ,"Чериковское райпо"                         ,3,6},
					{6021           ,"Шкловское райпо"                           ,3,6},
					{6032           ,"УП ""Бобруйский торговый центр"""          ,3,6},
					{6052           ,"ЧУП ""Могилевская СПМК"""                  ,3,6},
					{6053           ,"ЧУП ""Бобруйская ПМК"""                    ,3,6},
					{227100000000000,"УП "" Винный мастер"""                     ,3,2},
					{251600080200001,"УП ""Слуцкий рынок"""                      ,3,2},
					{253100000000000,"Березинский филиал"                        ,3,2},
					{253500000000000,"Дзержинский филиал"                        ,3,2},
					{253900000000000,"Логойский филиал"                          ,3,2},
					{254100000000000,"Минский филиал"                            ,3,2},
					{254900000000000,"Столбцовский филиал"                       ,3,2}
				}
			)

		],

	// отдельные
	spr.BAL    = [ // Бухгалтерский БАЛАНС
		Спр_ББ_Часть   = // Справочники частей
			let Источник = #table(type table
					[
						Часть_ID 		  = number,
						ЧастьИмяПолное	  = text,
						ЧастьИмяКраткое   = text,
						ЧастьАббревиутура = text
					],
					{
						{1,	"Активы",								"Актив",	"А"},
						{2,	"Собственный капитал и обязательства", 	"Пассив",	"П"}
					}
				)
			in 	Источник,

		Спр_ББ_Раздел  = // Справочники разделов
			let	Источник = #table(type table
					[
						Раздел_ID = number,
						РазделПрефикс= text,
						РазделИмя = text,
						РазделИмяПолное = text,
						РазделИмяСокращенное = text,
						РазделИмяАббревиатура = text
					],
					{
						{1, "Раздел 1", "Долгосрочные активы",         "Раздел 1 Долгосрочные активы",         "ДолгАктивы",  "ДА" },
						{2, "Раздел 2", "Краткосрочные активы",        "Раздел 2 Краткосрочные активы",        "КраткАктивы", "КА" },
						{3, "Раздел 3", "Собственный капитал",         "Раздел 3 Собственный капитал",         "СобКапитал",  "СК" },
						{4, "Раздел 4", "Долгосрочные обязательства",  "Раздел 4 Долгосрочные обязательства",  "ДолгОбяз",    "ДО" },
						{5, "Раздел 5", "Краткосрочные обязательства", "Раздел 5 Краткосрочные обязательства", "КраткОбяз",   "КО" },
						{0, "Раздел 6", "ЗаBAL_овые счета",        	   "Раздел 0 ЗаBAL_овые счета",            "ЗабСчета",    "ЗС" }
					} ),
				selectRow = Table.SelectRows(Источник, each ([Раздел_ID] <> 0))
			in 	selectRow,

		Спр_ББ_Группа  = // Справочники групп статей
			let	Источник = #table(type table
					[
						Группа_ID = number,
						ГруппаИмя = text
					],
					{
						{1, "Группа"},
						{2, "Подгруппа"}
					}
				)
			in 	Источник,

		Спр_ББ_Уровень = // Справочники уровней
			let Источник = Table.FromRows(Json.Document(Binary.Decompress(Binary.FromText("i45WMlTSUbqw+GLDhX0XNl3YemHvxR4FsNDyCxsuNl5sutijFKsTrWSEqQostODChgvbL2wBCu0GqzPGVAcWWnixCWge0LSL/WB1JpjqwELzgQJbQPYiVMcCAA==", BinaryEncoding.Base64), Compression.Deflate)),
				let _t = ((type nullable text) meta [Serialized.Text = true])
				in type table [Уровень_ID = _t, УровеньПрефикс = _t, УровеньИмя = _t]),
				tr = Table.TransformColumnTypes(Источник,{{"Уровень_ID", Int64.Type}, {"УровеньИмя", type text}})
			in  tr
		],

	spr.OPU    = [ // Отчет о прибылях и убытках
		Спр_OPU_ВидДеятельности =
			let Источник = #table(type table
					[
						ВидДеятельности_ID  = Int64.Type,
						ДиапазонСтрок_ID    = text,
						ВидДеятельностиИмя  = text,
						ВидДеятельностиIDИмя  = text
					],
					{
						{1, "10..90",    "Текущая деятельность",		"1 Текущая деятельность"},
						{2, "100..119",  "Инвестиционная деятельность",	"2 Инвестиционная деятельность"},
						{3, "120..139",  "Финансовая деятельность",		"3 Финансовая деятельность"},
						{4, "140..1000", "Другие показатели",			"4 Другие показатели"}
					}
				)
			in Источник,

		Спр_OPU_ВидПрибыли 		= // Справочники видов показателей прибыли
			let Источник = #table(type table [
				Прибыль_ID = Int64.Type,
				ПрибыльВид = text ],
				{
					{1, "Показатели прибыли"},
					{2, "Другие показатели"}
				} )
			in Источник,

		Спр_OPU_Группа 			= // Справочники групп статей
			let Источник = #table(type table [
				Группа_ID = number,
				ГруппаИмя = text    ],
				{
					{1,"Группа"},
					{2,"Подгруппа"}
				} )
			in Источник
		],

	spr.ODDS   = [ // Отчет о движении денежных средств
		Спр_ODDS_ВидДеятельности = //
			let Источник = #table(type table
					[
						ВидДеятельности_ID   = Int64.Type,
						ДиапазонСтрокДеят_ID = text,
						ВидДеятельностиИмя 	 = text,
						ВидДеятельностиIDИмя = text
					],
					{
						{1, "20..40",	"Текущая деятельность",			"1 Текущая деятельность"},
						{2, "50..70",	"Инвестиционная деятельность",	"2 Инвестиционная деятельность"},
						{3, "80..100",	"Финансовая деятельность",		"3 Финансовая деятельность"},
						{4, "110..140",	"Другое",						"4 Другое"}
					}
					)
			in  Table.SelectColumns(Источник, {"ВидДеятельности_ID","ВидДеятельностиИмя","ВидДеятельностиIDИмя"}),

		Спр_ODDS_ВидДвижения 	 = // Справочники видов движения
			let Источник = #table(type table
					[
						ВидДвижения_ID  	 = Int64.Type,
						ДиапазонСтрокДвиж_ID = text,
						ВидДвиженияИмя 		 = text,
						ВидДвиженияIDИмя	 = text
					],
					{
						{1, "20..24,50..55,80..84",	"Поступило",	  "1 Поступило"},
						{2, "30..34,60..64,90..95",	"Выбыло",		  "2 Выбыло"},
						{3, "40",					"РезультатТек",	  "3 РезультатТек"},
						{4, "70",					"РезультатИнв",	  "3 РезультатИнв"},
						{5, "100",					"РезультатФин",	  "3 РезультатФин"},
						{6, "110",					"РезультатОбщий", "1 РезультатОбщий"},
						{7, "120..140",				"Другие",		  "2 Другие"}
					}
				)
			in  Table.SelectColumns(Источник, {"ВидДвижения_ID","ВидДвиженияИмя","ВидДвиженияIDИмя"})

		],

	spr.OSK    = [ // Отчет об изменении собственного капитала
		Спр_OSK_Раздел =
			let	from = #table(type table
				[
					Раздел_ID 	  = number,
					РазделПрефикс = text,
					РазделИмя 	  = text,
					СтрокаНач     = number,
					СтрокаКон     = number
				],
				{
					{1, "Раздел 1", "Позапрошлый год", 10, 40},
					{2, "Раздел 2", "Прошлый год", 50, 130},
					{3, "Раздел 3", "Текущий год", 140, 200}
				}),
				to = Table.AddColumn(from, "РазделПолноеИмя", each [РазделПрефикс] & " " & [РазделИмя])
			in	Table.SelectColumns(to, {"Раздел_ID","РазделПрефикс","РазделИмя"})
		],

	spr.PRIL   = [ // Приложение к Балансу
		Спр_Разделы 	= 
				let Источник = #table(type table
					[   Раздел_ID     = number,
						РазделПрефикс = text,
						РазделИмя     = text,
						Str1          = number,
						Str2          = number],
					{   {1, "Раздел 1", "Основные средства и нематериальные активы", 		1, 36},
						// {2, "Раздел 1", "Справочно. ", 1, 5},
						// {3, "Раздел 1", "Источники финансирования сложений ", 30, 36},
						{2, "Раздел 2", "Незавершенные вложения во внеоборотные активы", 	50, 65},
						{3, "Раздел 3", "Резервы по сомнительным долгам", 					70, 70},
						{4, "Раздел 4", "Финансовые вложения", 								80, 95},
						{5, "Раздел 5", "Полученные кредиты и займы", 						140, 171},
						{6, "Раздел 6", "Справка о платежах в бюджет и ФСЗН", 				1010, 1320},
						{7, "Раздел 7", "ЗаBAL_овые счета", 								1401, 1411}
					}),

					AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
					Reorder = Table.ReorderColumns(AddCol,
							let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col = 3  // 0 - если в начало
							in  List.InsertRange(List.Difference(nms, lst ), col, lst) ),
					SelCol = Table.SelectColumns(Reorder, List.FirstN(Table.ColumnNames(Reorder), 4) )
				in  SelCol,

		Спр_Заголовки 	= // заголовки разделов
				let Источник = #table(type table
					[
						IndexIn = text,      // порядковый номер в заголовка пределах раздела отчета
						ЗаголовокИмя = text, // имя заголовка
						НомерРаздела = text  // номера разделов, в которых есть этот заголовок
					],
					{
						{"1", "На начало года","1"},
						{"2", "Поступило","1"},
						{"3", "Выбыло","1"},
						{"4", "На конец отчетного периода всего","1"},
						{"5", "На конец отчетного периода всего, и.т.ч. полностью самортизированных","1"},

						{"1", "На начало года","2"},
						{"2", "Поступило","2"},
						{"3", "Выбыло","2"},
						{"4", "На конец периода","2"},

						{"1", "На начало года","3"},
						{"2", "Создано","3"},
						{"3", "Списано","3"},
						{"4", "На конец периода","3"},

						{"1", "На начало года","4,7"},
						{"2", "На конец периода","4,7"},

						{"1", "На начало года","5"},
						{"2", "Получено","5"},
						{"3", "Погашено","5"},
						{"4", "На конец периода","5"},

						{"1", "Пpичитается по расчету","6"},
						{"2", "Изpасходовано","6"},
						{"3", "Фактически уплачено","6"}
					}),

					// разворачиваем по строкам через запятую
						ExpandList = Table.ExpandListColumn(
							Table.TransformColumns(Источник,
									{{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
									let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}),
									"НомерРаздела"),

					// 	объединяем два столбца в новый используя функцию
						CombineCol =
							Table.CombineColumns(ExpandList,
									{"НомерРаздела", "IndexIn"},
									Combiner.CombineTextByDelimiter("", QuoteStyle.None),
									"Заголовок_ID"
							),
					// 	Написать надо что получилось в итоге!!!
						TransCol   = Table.TransformColumnTypes(CombineCol,{{"Заголовок_ID", Int64.Type}})
				in  TransCol
		],

	spr.BKS3   = [ // Отчет о доходах, расходах и финрезультатах "Форма 3 БКС"
		Спр_Разделы   = // Названия разделов отчета 3 БКС
			let Источник = #table(type table
				[
					Раздел_ID     = number,
					РазделПрефикс = text,
					РазделСокр    = text,
					РазделИмя     = text,
					Str1          = number,
					Str2          = number
				],
				{
					{1, "Раздел 1", "Раздел 1. ИЗДЕРЖКИ ОБРАЩЕНИЯ", "ИЗДЕРЖКИ ОБРАЩЕНИЯ (РАСХОДЫ) НА РЕАЛИЗОВАННЫЕ ТОВАРЫ И ТОВАРЫ, ПЕРЕДАННЫЕ ФИЛИАЛАМИ", 10100, 10700},
					{2, "Раздел 2", "Раздел 2. ГРУППИРОВКА ИЗДЕРЖЕК ОБРАЩЕНИЯ", "ГРУППИРОВКА ИЗДЕРЖЕК ОБРАЩЕНИЯ ПО ЭЛЕМЕНТАМ ЗАТРАТ (строка 11500)", 21000, 25000},
					{3, "Раздел 3", "Раздел 3. ДОХОДЫ, РАСХОДЫ И ФИНАНСОВЫЕ РЕЗУЛЬТАТЫ", "ДОХОДЫ, РАСХОДЫ И ФИНАНСОВЫЕ РЕЗУЛЬТАТЫ ОТ РЕАЛИЗАЦИИ ТОВАРОВ И ПЕРЕДАЧИ ТОВАРОВ ФИЛИАЛАМИ", 30000, 38005},
					{4, "Раздел 4", "Раздел 4. СТАТИСТИЧЕСКИЕ ПОКАЗАТЕЛИ", "ОСНОВНЫЕ СТАТИСТИЧЕСКИЕ ПОКАЗАТЕЛИ ТОРГОВО-ЗАКУПОЧНОЙ ДЕЯТЕЛЬНОСТИ", 40000, 47000},
					{5, "Раздел 5", "Раздел 5. ПРОДАЖА ТОВАРОВ НАСЕЛЕНИЮ В КРЕДИТ", "ДАННЫЕ О ПРОДАЖЕ ТОВАРОВ НАСЕЛЕНИЮ В КРЕДИТ В РОЗНИЧНОЙ ТОРГОВЛЕ", 50000, 53100},
					{6, "Раздел 6", "Раздел 6. ЗАПАСЫ ТОВАРОВ НА КОНЕЦ ПЕРИОДА", "ЗАПАСЫ ТОВАРОВ И ТАРЫ НА КОНЕЦ ОТЧЕТНОГО ПЕРИОДА", 60000, 61000}
				}
				),

				AddCol  = 
					Table.AddColumn(Источник, "РазделИмяПрефикс", 
						each [РазделПрефикс] & " " & [РазделИмя]
					),

				Reorder =
					Table.ReorderColumns(AddCol,
						let
							nms = Table.ColumnNames(AddCol),
							lst = {"РазделИмяПрефикс"},
							col = 3	// 0 - если в начало
						in
							List.InsertRange(List.Difference(nms, lst), col, lst) ),

				SelCol =
					Table.SelectColumns(
						Reorder,
						List.FirstN(Table.ColumnNames(Reorder), 4)
					)
			in  SelCol,


		Спр_Заголовки = // заголовки разделов
			let Источник = #table(type table
				[
					IndexIn = text,       // порядковый номер в заголовка пределах раздела отчета
					ЗаголовокИмя = text,  // имя заголовка
					ЗаголовокСокр = text, // имя заголовка
					НомерРаздела = text   // номера разделов, в которых есть этот заголовок
				],
				{
					{"1", "Розничная торговля",		"Розница",	"1,2,3,6"},
					{"2", "Оптовая торговля",		"Опт",		"1,2,3,6"},
					{"3", "Всего по торговле",		"Торговля",	"1,2,3,6"},
					{"4", "Общественное питание",	"Общепит",	"1,2,3,6"},
					{"5", "Заготовки",				"Заготовки","1,2,3,6"},
					{"6", "Итого",					"Итого",	"1,2,3,6"},
					{"1", "Отчетный период",		"Отчетный",	"4"},
					{"2", "Прошлый период",			"Прошлый",	"4"},
					{"1", "Сумма",					"сумма",	"5"}
				}),

				// разворачиваем по строкам через запятую
					ExpandList =
						Table.ExpandListColumn(
							Table.TransformColumns(Источник, {
								{"НомерРаздела",
									Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
									let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}
								}
							), "НомерРаздела"
						),

				// 	дубликат столбца, потому как "IndexIn" объединится
					dubl   = Table.DuplicateColumn(ExpandList, "IndexIn", "IndexIn+"),

				// 	объединяем два столбца в новый используя функцию
					CombineCol =
						Table.CombineColumns(
							dubl,
							{"НомерРаздела", "IndexIn+"},
							Combiner.CombineTextByDelimiter("", QuoteStyle.None),
							"Заголовок_ID"
						),

				TransCol   = Table.TransformColumnTypes(CombineCol, {"Заголовок_ID", Int64.Type} )
			in  TransCol
		],

	spr.BKS7   = [ // Отчет "Форма 7 БКС"
		Спр_Разделы   = // Названия разделов
			let Источник = #table(type table
				[   Раздел_ID     = number,
					РазделПрефикс = text,
					РазделИмя     = text,
					Str1          = number,
					Str2          = number],
				{   {1, "Раздел 1", "ЗАТРАТЫ НА ПРОИЗВОДСТВО В ОТРАСЛЯХ ПРОМЫШЛЕHHОСТИ", 10, 170},
					{2, "Раздел 2", "ГРУППИРОВКА ЗАТРАТ НА ПРОИЗВОДСТВО ПО ЭЛЕМЕНТАМ   (строка 160)", 210, 260},
					{3, "Раздел 3", "ДОХОДЫ, РАСХОДЫ И ФИНАНСОВЫЕ РЕЗУЛЬТАТЫ ОТ РЕАЛИЗАЦИИ ПРОДУКЦИИ (РАБОТ, УСЛУГ)", 300, 406},
					{4, "Раздел 4", "ДОХОДЫ, РАСХОДЫ И ФИНАНСОВЫЕ РЕЗУЛЬТАТЫ ПО ДРУГИМ ВИДАМ ДЕЯТЕЛЬНОСТИ", 410, 496},
					{5, "Раздел 5", "ГРУППИРОВКА ЗАТРАТ НА ПРОИЗВОДСТВО ПО ДРУГИМ ВИДАМ ДЕЯТЕЛЬНОСТИ ПО ЭЛЕМЕНТАМ", 500, 560},
					{6, "Раздел 6", "Справочно к разделу 4", 420, 420} }),
				AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя]),
				Reorder = Table.ReorderColumns(AddCol,
						let nms = Table.ColumnNames(AddCol), lst = {"РазделИмяПрефикс"}, col =3  // 0 - если в начало
						in  List.InsertRange(List.Difference(nms, lst ), col, lst) ),
				SelCol = Table.SelectColumns(Reorder, List.FirstN(Table.ColumnNames(Reorder), 4) )
		 in  SelCol,

		Спр_Заголовки = // заголовки разделов
			let Источник = #table(type table
				[
					IndexOut = number,     	// порядковый номер видов деятельности для сортировки в отчете
					IndexIn  = text,      	// порядковый номер заголовка в пределах раздела отчета!
					ЗаголовокИмя = text, 	// имя заголовка
					НомерРаздела = text,  	// номера разделов, в которых есть этот заголовок
					ВидДеят = text 			// виды деятельности (сокращенный вариант без итогов)
				],
				{
					{0,	 "1",  "Всего по промышленности",				"1,2,3", 	""},
					{1,	 "2",  "хлеба и хлебобулочных изделий",			"1,2,3", 	"хлебные изделия"},
					{2,	 "3",  "кондитерских изделий",					"1,2,3", 	"кондитерка"},
					{3,	 "4",  "колбасных изделий",						"1,2,3", 	"колбасные изделия"},
					{4,	 "5",  "полуфабрикатов мясных",					"1,2,3", 	"п/ф мясные"},
					{5,	 "6",  "плодоовощных консервов",				"1,2,3", 	"консервы"},
					{6,	 "7",  "алкогольных напитков",					"1,2,3", 	"алкогольные напитки"},
					{7,	 "8",  "безалкогольных напитков",				"1,2,3", 	"безалгогольные напитки"},
					{8,	 "9",  "пpочих пищевых продуктов",				"1,2,3", 	"прочие пищевые"},
					{9,	 "10", "пpочих непищевых продуктов",			"1,2,3", 	"прочие непищевые"},
					{0,	 "1",  "Всего по другим видам деятельности",	"4,5", 		""},
					{10, "2",  "сельское хозяйство",					"4,5", 		"сельское хозяйство"},
					{11, "3",  "транспорт строительство",				"4,5",		"транспорт, строительство"},
					{12, "4",  "информационно техническое обслуживание","4,5,",		"инфо и тех обслуживание"},
					{13, "5",  "рынки",									"4,5",		"рынки"},
					{14, "6",  "другие виды деятельности",				"4,5",		"другие виды"},
					{15, "7",  "сдача внаем имущества",					"4,5",		"сдача внаем имущества"},
					{0,	 "1",  "За 12 месяцев текущего года",			"6",		""},
					{0,	 "2",  "За 12 месяцев прошлого года",			"6",		""}
				}),

				// разворачиваем по строкам через запятую
				ExpandList =
					Table.ExpandListColumn(
						Table.TransformColumns(Источник,
								{
									{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
									let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}
								}),
						"НомерРаздела"
					),

				// дубликат столбца, потому как "IndexIn" объединится
					dubl   = Table.DuplicateColumn(ExpandList, "IndexIn", "IndexIn+"),

				// объединяем два столбца в новый используя функцию
				CombineCol =
					Table.CombineColumns(
						dubl,
						{"НомерРаздела", "IndexIn+"},
						Combiner.CombineTextByDelimiter("", QuoteStyle.None),
						"Заголовок_ID"
					),

				TransCol   = Table.TransformColumnTypes(CombineCol, {"Заголовок_ID", Int64.Type} )
			in  TransCol
		],

	spr.1P     = [ // Отчет "Форма 1п_райпо" и "1п_ОПС"
		Спр_Разделы    =
					let Источник = #table(type table
						[   ID_Раздел     = number,
							РазделПрефикс = text,
							РазделИмя     = text
						],
						{   {1, "Раздел 1", "Раздел I. Производство промышленной продукции (услуг промышленного характера)"},
							{2, "Раздел 2", "Раздел II. Основные показатели работы промышленности, тыс. руб."}
						}),
						AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя])
					in  AddCol,

		Спр_ЕдИзм      = // справочник единиц измерения (для замены)
					// lst1 = List.Buffer({"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"}),
					let lst1 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"},
						lst2 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"},
						from = #table( {"Что","На что"}, List.Zip({lst1, lst2}))
					in  from,

		Спр_ВидТоваров =
					let Источник = #table(type table
						[
							ID_КатегорияТоваров = number,
							КатегорияТоваров = text
						],
						{
							{1,"Продовольственные товары"},
							{2,"Непродовольственные товары"}
						})
					in  Источник,

		Спр_Заголовки  = // заголовки разделов
					let Источник = #table(type table
						[
							IndexIn = text,      // порядковый номер заголовка в пределах раздела отчета!
							ЗаголовокИмя = text, // имя заголовка
							НомерРаздела = text  // номера разделов, в которых есть этот заголовок
						],
						{
							{"1", "Наименование вида продукции (услуги)","1"},
							{"2", "Единица измерения","1"},
							{"3", "Фактическое производство","1"},
							{"4", "Фактическое производство за предыдущий год","1"},
							{"5", "Темп роста, %","1"},
							{"6", "Стоимость продукции в фактических ценах","1"},
							{"7", "Стоимость продукции в фактических ценах за предыдущий год","1"},

							{"1", "Hаименование показателя","2"},
							{"2", "Отчетный год","2"},
							{"3", "Предыдущий год","2"},
							{"4", "Темп изменения, %","2"}
						}),

						// разворачиваем по строкам через запятую
						ExpandList = Table.ExpandListColumn(Table.TransformColumns(Источник,
										{{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
										let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}),
										"НомерРаздела"),

						// объединяем два столбца в новый используя функцию НомерРаздела + Индекс
						CombineCol = Table.CombineColumns(ExpandList,
										{"НомерРаздела", "IndexIn"},
										Combiner.CombineTextByDelimiter("", QuoteStyle.None),
										"Заголовок_ID"),
						TransCol   = Table.TransformColumnTypes(CombineCol,{{"Заголовок_ID", Int64.Type}})
					in  TransCol
		],

	spr.1TORG  = [ // 1 торг розница 3 и 4 разделы
		Спр_Разделы   =
			let Источник = #table(type table
				[   Раздел_ID     = number,
					РазделПрефикс = text,
					РазделИмя     = text
				],
				{   {3, "Раздел 3", "Раздел III. ТОРГОВЫЕ ОБЪЕКТЫ"},
					{4, "Раздел 4", "Раздел IV. РОЗНИЧНАЯ ТОРГОВЛЯ ФИРМЕННЫХ МАГАЗИНОВ"}
				}),
				AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя])
			in  AddCol,

		Спр_Заголовки =
			let
				Источник = #table(type table
					[
						IndexIn = text,      // порядковый номер заголовка в пределах раздела отчета!!!
						ЗаголовокИмяКраткое = text, // имя заголовка
						ЗаголовокИмя = text, // имя заголовка
						НомерРаздела = text  // номера разделов, в которых есть этот заголовок
					],
					{
						{"1", "Торговых объектов, ед", 			"Количество торговых объектов, ед", "3"},
						{"2", "Торговая площадь, кв.м", 		"Торговая площадь, кв.м", "3"},
						
						{"3", "Торговых объектов, ед", 	"Количество торговых объектов в сельской местности, ед", "3"},
						{"4", "Торговая площадь, кв.м", 	"Торговая площадь торговых объектов в сельской местности, кв.м", "3"},
						
						// добавленные потом колонки
						{"5", "Торговых объектов, ед", 	"Количество торговых объектов в городе, ед", "3"},
						{"6", "Торговая площадь, кв.м", 	"Торговая площадь торговых объектов в городе, кв.м", "3"},
						
						// {"1", "Торговых объектов, ед", 			"Количество торговых объектов, ед", "3"},
						// {"2", "Торговая площадь, кв.м", 		"Торговая площадь, кв.м", "3"},
						// {"3", "Торговых объектов село, ед", 	"Количество торговых объектов в сельской местности, ед", "3"},
						// {"4", "Торговая площадь село, кв.м", 	"Торговая площадь торговых объектов в сельской местности, кв.м", "3"},
						// {"5", "Торговых объектов город, ед", 	"Количество торговых объектов в городе, ед", "3"},
						// {"6", "Торговая площадь город, кв.м", 	"Торговая площадь торговых объектов в городе, кв.м", "3"},
						{"1", "ЕдИзм", "ЕдИзм",	"4"},
						{"2", "Отчетный год","За отчетный год","4"}
					}),

				// разворачиваем по строкам через запятую
					ExpandList =
						Table.ExpandListColumn(
							Table.TransformColumns(Источник,
								{{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
								let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}} ),
							"НомерРаздела"
						),

				// 	объединяем два столбца в новый используя функцию НомерРаздела + Индекс
					CombineCol =
						Table.CombineColumns(
							ExpandList,
							{"НомерРаздела", "IndexIn"},
							Combiner.CombineTextByDelimiter("", QuoteStyle.None),
							"Заголовок_ID"
						),

					TransCol   = Table.TransformColumnTypes(CombineCol,{{"Заголовок_ID", Int64.Type}})
			in
				TransCol
		],
	
	spr.4TORG  = [ // 1 торг  
		Спр_Разделы   = 
			let Источник = #table(type table
				[   Раздел_ID     = number,
					РазделПрефикс = text,
					РазделИмя     = text
				],
				{   {1, "Раздел 1", "Раздел I. СТРУКТУРА ПРОДАЖИ И ЗАПАСЫ ТОВАРОВ ПО ВИДАМ"},
					{2, "Раздел 2", "Раздел II. РОЗНИЧНЫЙ ТОВАРООБОРОТ"}
				}),
				AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя])
			in  AddCol,

		Спр_ЕдИзм     = // для замены
			let lst3 = {"т","т","дал","тыс.шт.","ц","шт.","тыс.pуб.","куб.м","кв.м"},
				lst4 = {"т","т","дал","тыс.шт.","ц","шт.","тыс.pуб.","куб.м","кв.м"},

				lst1 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"},
				lst2 = {"Гкал", "м.п.", "млн. руб.", "млн. шт.", "млн.руб.", "т.дал", "тонн", "туб", "тыс. кв. м", "тыс. руб.", "тыс. штук","тыс.дал", "тыс.кВт.ч", "тыс.пар", "тыс.руб.", "тыс.штук"},

				from = #table( {"Что","На что"}, List.Zip({lst3, lst4}))
			in  from,

		Спр_Заголовки = 
			let Источник = #table(type table
				[
					IndexIn = text,      // порядковый номер заголовка в пределах раздела отчета!
					ЗаголовокИмя = text, // имя заголовка
					НомерРаздела = text  // номера разделов, в которых есть этот заголовок
				],
				{
					{"1", "Код по СК 33.004-2020", 	"1"},
					{"2", "Наименование показателя","1"},
					{"3", "Код  стр.",				"1"},
					{"4", "ЕдИзм",					"1"},
					{"5", "Продано Всего",			"1"},
					{"6", "Продано Отеч",			"1"},
					{"7", "Запасы на конец Всего",	"1"},
					{"8", "Запасы на конец Отеч",	"1"},

					{"1", "Наименование показателя","2"},
					{"2", "Код  стр.",				"2"},
					{"3", "ЕдИзм",					"2"},
					{"4", "Розничный товарооборот Всего",	"2"},
					{"5", "Розничный товарооборот Отеч",	"2"}

				}),

				// 	разворачиваем по строкам через запятую
					ExpandList =
						Table.ExpandListColumn(
							Table.TransformColumns(Источник,
								{{"НомерРаздела", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
								let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}} ),
							"НомерРаздела"
						),

				// 	объединяем два столбца в новый используя функцию НомерРаздела + Индекс
					CombineCol =
						Table.CombineColumns(
							ExpandList,
							{"НомерРаздела", "IndexIn"},
							Combiner.CombineTextByDelimiter("", QuoteStyle.None),
							"Заголовок_ID"
						),
				// 	меняем тип столбца
					TransCol   =
						Table.TransformColumnTypes(
							CombineCol,
							{{"Заголовок_ID", Int64.Type}}
						)
			in  
				TransCol
		],
	spr.11z    = [ // Отчет по заготовкам
		Спр_Заголовки = 		
				#table(
					type table
					[	
						Header_ID 	= nullable Int64.Type, 
						Header_Old	= nullable text, 
						Header	 	= nullable text
					],
					{
						 {21,"Наименование показателей"                                                              ,"Показатель"                                       }
						,{22,"Един.    измер."                                                                       ,"Ед. изм"                                          }
						,{23,"№ стро-ки"                                                                             ,"Код строки"                                       }
						,{1 ,"Нали-     чие  на начало отчетно-го года"                                              ,"Остатки на начало"                                }
						,{2 ,"Закуплено (всего), количество"                                                         ,"Закуплено, количество"                            }
						,{3 ,"Закуплено (всего), стоимость (тыс.руб.)"                                               ,"Закуплено, стоимость"                             }
						,{4 ,"В том числе у населения, количество"                                                   ,"Закуплено у населения, количество"                }
						,{5 ,"В том числе у населения, стоимость (тыс.руб.)"                                         ,"Закуплено у населения, стоимость"                 }
						,{6 ,"Реализовано всего"                                                                     ,"Реализовано, количество"                          }
						,{7 ,"Реализовано торговли ПК"                                                               ,"Реализовано (торговле ПК)"                        }
						,{8 ,"Реализовано общественного питания ПК"                                                  ,"Реализовано (общепиту ПК)"                        }
						,{9 ,"Реализовано промышлен ной пере-работки ПК"                                             ,"Реализовано (промышленности ПК)"                  }
						,{10,"потери и недостачи (пересор-тица)"                                                     ,"Потери и недостачи"                               }
						,{11,"Реализовано по прочим каналам сбыта всего"                                             ,"Реализовано (прочим каналам сбыта)"               }
						,{12,"Реализовано по прочим каналам сбыта в т.ч. промышленным перерабатыва-ющим организациям","Реализовано (прочим каналам сбыта промышленность)"}
						,{13,"Наличие на  конец отчетного периода"                                                   ,"Остатки на конец"                                 }
						,{14,"Средняя цена"                                                                          ,"Средняя цена"                                     }
					}
				),
		
		спр_Групп =
			#table(type table
				[
					ID_Группы = nullable Int64.Type, 
					Название = nullable text
				],
				{
					{1,	"Продукция растениеводства"},
					{2,	"Дикорастущая продукция"   },
					{3,	"Продукция животноводства" },
					{4,	"Кожевенное сырье:"        },
					{5,	"Вторичное сырье"          }
				}
			)
		],
	
	spr.12po   = [ // отчет о производстве 12-п (есть по райпо, есть по ОПО) справочники одинаковые наверное
		Спр_Разделы   = 
			let Источник = #table(type table
				[   Раздел_ID     = number,
					РазделПрефикс = text,
					РазделИмя     = text
				],
				{   {1, "Раздел 1", "Основные показатели промышленной деятельности"},
					{2, "Раздел 2", "Сведения о видах экономической деятельности по организации"},
					{3, "Раздел 3", "Количество пpоизведенной промышленной пpодукции"},
					{4, "Раздел 4", "Стоимость промышленной пpодукции, приведенной в разделе III"}
				}),
				AddCol  = Table.AddColumn(Источник, "РазделИмяПрефикс", each [РазделПрефикс] & " " & [РазделИмя])
			in  AddCol,

		Спр_Заголовки 	= // заголовки разделов
				let Источник = #table(type table
					[
						Index	 	= text, // порядковый номер заголовка пределах раздела отчета
						Header 		= text, // имя заголовка
						Razdel_ID 	= text  // номера разделов, в которых есть этот заголовок
					],
					{
						{"А", "Показатель","1,2"},
						{"Б", "Код показателя","1,2"},
						
						{"2", "Фактически за отчетный год, тыс. руб.","1"},
						{"4", "Фактически за предыдущий год, тыс. руб.","1"},
						{"6", "Темп роста по количеству за период, %","1"},

						{"2",  "Объем производства за отчетный год","2"},
						{"4",  "Объем производства за прошлый год","2"},
						{"5",  "Стоимость переработанного давальческого сырья за отчетный год","2"},
						{"6",  "Стоимость переработанного давальческого сырья за прошлый год","2"},
						{"10", "Темп роста по объёму производства за год","2"},

						{"А",  "Вид продукции","3"},
						{"Б",  "ЕдИзм","3"},
						{"3",  "Фактически произведено за отчетный год","3"},
						{"8",  "Фактически произведено за предыдущий год","3"},
						{"11", "Запасы готовой продукции на конец периода","3"},
						{"13", "Темпы роста по количеству за год","3"},

						{"А",  "Вид продукции","4"},
						{"1" , "Фактически произведено за отчетный год","4"},
						{"2" , "Фактически произведено за прошлый год","4"},
						{"3" , "Стоимость переработанного давальческого сырья за отчетный год","4"},
						{"5" , "Темп роста по объёму производства за год","4"}
					}),

					// 	разворачиваем по строкам через запятую
						ExpandList = Table.ExpandListColumn(
							Table.TransformColumns(Источник,
									{{"Razdel_ID", Splitter.SplitTextByDelimiter(",", QuoteStyle.Csv),
									let itemType = (type nullable text) meta [Serialized.Text = true] in type {itemType}}}),
									"Razdel_ID"),

					// 	объединяем два столбца в новый используя функцию
						CombineCol =
							Table.CombineColumns(ExpandList,
									{"Razdel_ID", "Index"},
									Combiner.CombineTextByDelimiter("", QuoteStyle.None),
									"Header_ID"
							),
					// 	Написать надо что получилось в итоге!!!
						TransCol   = Table.TransformColumnTypes(CombineCol,{{"Header_ID", type text}}) // Int64.Type
				in  TransCol
		 	],

	spr.None = "",

//	ОТЧЕТЫ
	rep.FAQ   = [],// пояснения
		/*	загрузка любого отчета после импорта библиотеки "F" и "F_Form"
				= 	F_Form[rep.BAL]( F_Form[Connect][Homel] )
		*/
	rep.BAL4  = // БАЛАНС
		(optional options as record) => [
		//	АРГУМЕНТЫ   Шаг 0. Список аргументов 
           	datasetName = let value = Record.FieldOrDefault(options, "ДатаCетИмя", null)  in if value = 0 then null else value,
			tbl         = let value = Record.FieldOrDefault(options, "Таблица", null)     in if List.Contains({null, 0, ""}, value) then null else value,
			
		// 	ПОДКЛЮЧЕНИЕ Шаг 1. На входе спиcок файлов с техколонками в т.ч. колонка Company
			flagSubTable = if tbl = null then true else false, // true = работаем с таблицей, false = работаем с подтаблицей
			Connect = 
				if 		flagSubTable = false
				then 	tbl
				else 	F_Form[Connect_]( [								
							ДатаCетИмя 	= F_Form[GetName](datasetName),	// активное имя датасета сделать проверку на null = ?? 0
							ОтчетИмя 	= "bal", 						// если фильтровать отчеты по имени файла form_ID в папке с файлами отчетов
							МаскаИмениПапки = null, 					// если фильтроватиь отчеты по папкам с отчетами
							НомерШага 	= null
						]),
			
		//	ЗАГРУЗКА	Шаг 2. Фильтруем строки в каждом отчете (на входе ...)
			Load =
				F_Form[SelectRow_]([
					subTable = flagSubTable, // работаем с таблицей или подтаблицами
					tbl 	 = Connect, 	 // на входе колонка с таблицей или подтаблицами
					nmsOld 	 = F_Form[getColumnNames]({1,2,3,4}),
					nmsNew 	 = {"Статья", "Строка_ID", "На начало", "На конец"},
					colName  = "Строка_ID",
					kodStart = 10,
					kodEnd   = 1000
				]),
					
		//	РАЗВОРОТ	Шаг 3. Трансформируем исходные столбцы (чистим, меняем, типизируем)
			Expand  = 
				let from  = Load,
					fxAll = 
						let
							// 	список начальных колонок итоговой таблицы: общие + обязательные (форму еще не добавили)
								lst1 = {"Company", "DataSet", "form_ID", "Date"},
								lst2 = {"Статья", "Строка_ID", "На начало", "На конец"},

							// 	надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
								EXPAND = Table.ExpandTableColumn(from, "Step2", lst2, lst2),

							// 	типизируем и причесываем исходные столбцы + 4 техколонки
								Transform =
									Table.TransformColumns(
										EXPAND,
										{
											{ lst1{0}, (x)=> x, type text},               	// организация
											{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // номер датасета
											{ lst1{2}, (x)=> "bal", type text},             // отчет - Сделать join со Справочникиом отчетов
											{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата отчета

											{ lst2{0}, (x)=> x, type text}, 													// статья
											{ lst2{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}, 					// строка
											{ lst2{2}, (x)=> if Number.From(x) = 0 then null else Number.From(x), type number}, 	// на начало Int64.Type
											{ lst2{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), type number}  	// на конец
										}
									)
						in	Transform,
					
					fxOne = 
						let 
							trans   = // типизируем и причесываем исходные столбцы //! добавить техколонки внутрь подтаблиц!!
								Table.TransformColumns(from,{
									{ "Статья", 	(x)=> x, type text},	
									{ "Строка_ID", 	(x)=> try Number.From(x) otherwise null, Int64.Type},
									{ "На начало", 	(x)=> if Number.From(x) = 0 then null else Number.From(x), type number},
									{ "На конец", 	(x)=> if Number.From(x) = 0 then null else Number.From(x), type number} 
								})
						in 	trans,
					return = if flagSubTable then fxAll else fxOne
				in  return,
			
		//	ЧИСТКА	    Шаг 4. Трансформируем колонки (удаляем, меняем, добавляем)
			Clean 	 =
			// добавить в аргументы имя колонки или список колонок типа optional? ColNamesOrList + через если
				let	tbl = Expand,

				// 	удаляем лишний текст в колонке
					removeText = F[Text.RemoveMulti](tbl, "Статья", {"в том числе: ", " (лизинга)"}),

				// 	добавляем отступ в колонке "Статья"
					offsetText = F_Form[form.TextOffset](removeText, "Статья", "Строка_ID", 10)
				
				in 	offsetText,

		//	ОБОГАЩЕНИЕ  Шаг 5. Добавляем 4 колонки иерархии: Часть, Раздел, Статья, Подстатья (на основе колонки "Номер строки")
			//	Обогащение 1:  добавляем 3 колонки Часть_ID + Раздел_ID + Группа_ID + Уровень_ID
				Add1	 =
				let	tbl  = Clean,
					
					// 	Функция для получения значения Строка_ID
    				GetID = each [Строка_ID],

					// 	добавляем колонки
					AddColumns = F[fxAddMultiColumns](tbl, {
						
						// 	статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
							{ 	"Часть_ID",  each
									if List.Contains ({110..399}, GetID(_)) = true
									then 1
									else 2
								, Int64.Type},

						// 	номер раздела = первому символу номера строки. 0 - не вошедшие в разделы
							{ 	"Раздел_ID", each
									let f=(x as text) => Record.FieldOrDefault([1=1, 2=2, 4=3, 5=4, 6=5, 0=0], x, null),
										Номер = Text.Start(Text.From(GetID(_)),1)
									in f(Номер)
								, Int64.Type},

						// 	если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
							{ 	"Группа_ID", each
									if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1))
									then 1
									else 2
								, Int64.Type},

						// 	уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (= 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
							{	"Уровень_ID", each
										 if List.Contains ({300, 700}, GetID(_)) then 1
									else if List.Contains ({190, 290, 390, 490, 590, 690}, GetID(_)) then 2
									else if Number.Mod(GetID(_), 10) <> 0  and
											not List.Contains ({411}, GetID(_)) then 4 // исключение для строки 411 "Паевые взносы"
									else 3
								, Int64.Type}
						})

				in	AddColumns,
			
			//	Обогащение 2:  добавляем колонки с цифровым ID атрибутов
				Add2	=
				let	tbl  = Add1,
				
				// 	Функция для получения значения Строка_ID
    				GetID = each [Строка_ID],

				// 	Списки строк для проверки
					group_not_calculation = {190, 290, 390, 490, 590, 690, 300, 700, 630, 211, 212, 213, 214, 215, 216},
					group_A1 = {260, 270},
					group_A2 = {210, 250, 280},
					group_A3 = {220, 230, 240, 150, 170},
					group_A4 = {110, 120, 130, 140, 160, 180},
					group_P1 = {632, 633, 634, 635, 636, 637, 638, 639},
					group_P2 = {610, 620, 640, 650, 660, 670, 631},
					group_P3 = {510, 520, 530, 540, 550, 560},
					group_P4 = {410, 420, 430, 440, 450, 460, 470, 480},

				// 	Добавляем колонки с цифровым ID атрибутов
					AddColumn = F[fxAddMultiColumns](tbl, {
					
					// 	1 "Ликвидность_Статус" - выбираем строки, которые НЕ включаются в расчет = false
						{"Ликвидность_Статус", each
								if   List.Contains (group_not_calculation, GetID(_) ) = true
								then false
								else true,
							type logical},

					// 	2 "Ликвидность_Группа" - те что включаются - делим по категориям 4 + 4
						{"Ликвидность_Группа", each
							let id = GetID(_)
							in 	     if List.Contains (group_A1, id) then "А1"
								else if List.Contains (group_A2, id) then "А2"
								else if List.Contains (group_A3, id) then "А3"
								else if List.Contains (group_A4, id) then "А4"
								else if List.Contains (group_P1, id) then "П1"
								else if List.Contains (group_P2, id) then "П2"
								else if List.Contains (group_P3, id) then "П3"
								else if List.Contains (group_P4, id) then "П4"
								else null,
							type text}, // мб сделать null ? было "0"

					// 	3 "Ликвидность_Категория" - категория Ликвидности АП1 ... АП4
						{"Ликвидность_Категория", each
							let id = GetID(_)
							in 	 	 if List.Contains (group_A1 & group_P1, id) then "АП1"
								else if List.Contains (group_A2 & group_P2, id) then "АП2"
								else if List.Contains (group_A3 & group_P3, id) then "АП3"
								else if List.Contains (group_A4 & group_P4, id) then "АП4"
								else "0",
							type text}
					})

				in	AddColumn,
			
			//	Обогащение 3:  добавляем разные колонки из Справочникиов через Join
				Add3	=
				let	from  = Add2,
					// 	создаем Справочники статей из 2 + 1 колонок с рабочими названиями статей
						Спр_ББ_Статья =
							let
								nms      = {"Строка_ID"} & {"Статья", "СтатьяИмя"} & {"СтатьяСтрока"},
								distinct = Table.Distinct(Table.SelectColumns(from, {nms{0}, nms{1}}), {nms{0}} ),
								rename   = Table.RenameColumns(distinct, {{nms{1}, nms{2}}} ),

								// МЕНЯЕМ текст в столбце "СтатьяИмя"
								lst  = {
										{"Налог на добавленную стоимость",                      "НДС"},
										{"Hалог на добавленную стоимость",                      "НДС"}, // в чем отличие?! Дрогичин
										{"НДС по приобретенным товарам, работам, услугам",      "НДС по приобретенным Т,Р,У"},
										{"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
										{"Долгосрочные обязательства по лизинговым платежам",   "Долгосрочные обязательства по лизингу"},
										{"Денежные средства и эквиваленты денежных средств",    "Денежные средства и их эквиваленты"},
										{"кредиторская задолженность",                          "КЗ"},
										{"Вложения в долгосрочные активы",                      "Вложения в ДА"},
										{"дебиторская задолженность",                           "ДЗ"}
									},
								replace   = F[Text.ReplaceMulti](rename, nms{2}, lst),
								
								addColumn = Table.AddColumn(replace, nms{3}, each Text.From([Строка_ID]) & " " & [СтатьяИмя])
							in  addColumn,

					// 	добавляем значения колонок с ID (если переключатель IsReportViewGroup = True ?!)
						Join = [
							j  = from,
							j1 = Table.Join(j,	"Часть_ID",     F_Form[spr.BAL][Спр_ББ_Часть],   "Часть_ID"),
							j2 = Table.Join(j1, "Раздел_ID",    F_Form[spr.BAL][Спр_ББ_Раздел],  "Раздел_ID"),
							j3 = Table.Join(j2, "Группа_ID",    F_Form[spr.BAL][Спр_ББ_Группа],  "Группа_ID"),
							j4 = Table.Join(j3, "Уровень_ID",	F_Form[spr.BAL][Спр_ББ_Уровень], "Уровень_ID"),
							j5 = Table.Join(j4, "Строка_ID",    Спр_ББ_Статья, 		 	 		 "Строка_ID") // ? зачем это
						]  [j5]
				in  Join,

		//	ВЫВОД       Шаг 6. Наводим красоту перед выгрузкой в модель
			End = 
				let tbl = Add3,
					sel = F[Row.RemoveNullColumns](tbl, {"На начало","На конец"})
				in	sel
		] [End],

	rep.BAL   = // БАЛАНС //? использовать опциональные параметры, с названиями, как в расписании
		(optional datasetName as nullable text) => 
		//	ПОЯСНЕНИЯ 
			/*	
				Описание:
					на входе таблица со всеми файлами в папке, например  = F_Form[Connect][Homel]
					на входе должны быть: 4 техколонки + колонка с подтаблицами (4 колонки)
			*/
		let 
		// 	ПОДКЛЮЧЕНИЕ Шаг 1. На входе спиок файлов с техколонками в т.ч. колонка Company
			Connect = 	F_Form[Connect](
							F_Form[GetName](datasetName),	// активное имя датасета сделать проверку на null = ?? 0
							"bal",							// form_ID, если фильтровать отчеты по имени файла в общей папке с файлами отчетов
							null							// маска имени папки, если фильтроватиь отчеты по папкам с отчетами
						) (),
				
		// 	ЗАГРУЗКА    Шаг 2. 
			Load	= 
				let	tbl  = Connect,
					
					// 	определяем имя датасета - смотрим первую строку в отсортированной по датасету таблице
						company = tbl[Company]{0},	
					
					// 	фильтруем колонки в подтаблицах с учетом датасета (нужные значения могут быть в разных столбцах)
						nmsOld = if company = "Homel" then F_Form[getColumnNames]({1,2,3,4})
							else if company = "Mart"  then F_Form[getColumnNames]({1,5,6,7}) 
							else if company = "Drog"  then {"Column1","Column2","Column3","Column4"}
							else null,

					//	фильтруем строки в подтаблицах с отчетами по диапазону номеров строк отчета
						to	= 	F_Form[form.SelectRowsID](
									tbl,
									nmsOld,
									{"Статья", "Строка_ID", "На начало", "На конец"},
									"Строка_ID",
									10,
									1000
								)
				in  to,

		//	РАЗВОРОТ    Шаг 3. Разворачиваем и типизируем 4 + 4 колонки
			Expand 	 =
				let tbl = Load,
				/* 	На входе:
					- 4 общие колонки: 			Название организации + Номер датасета + Вид ответа + Дата отчета
					- 4 обязательные колонки:   Статья + Строка + На начало + На конец
				*/
				// 	список начальных колонок итоговой таблицы: общие + обязательные (форму еще не добавили)
					lst1 = {"Company", "DataSet", "form_ID", "Date"},
					lst2 = {"Статья", "Строка_ID", "На начало", "На конец"},

				// 	Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
				// 	надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
					EXPAND = Table.ExpandTableColumn(tbl, "Step2", lst2, lst2),

				// 	меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
					Transform =
						Table.TransformColumns(
							EXPAND,
							{
								{ lst1{0}, (x)=> x, type text},               	// организация
								{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // номер датасета
								{ lst1{2}, (x)=> "bal", type text},             // отчет - Сделать join со Справочникиом отчетов
								{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата отчета

								{ lst2{0}, (x)=> x, type text}, 													// статья
								{ lst2{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}, 					// строка
								{ lst2{2}, (x)=> if Number.From(x) = 0 then null else Number.From(x), type number}, 	// на начало Int64.Type
								{ lst2{3}, (x)=> if Number.From(x) = 0 then null else Number.From(x), type number}  	// на конец
							}
						)
				in
					Transform,

		//	ЧИСТКА	    Шаг 4. Трансформируем колонки (чистим, меняем, обогащаем)
			Clean 	 =
			// добавить в аргументы имя колонки или список колонок типа optional? ColNamesOrList + через если
				let	tbl = Expand,

				// 	удаляем лишний текст в колонке
					removeText = F[Text.RemoveMulti](tbl, "Статья", {"в том числе: ", " (лизинга)"}),

				// 	добавляем отступ в колонке "Статья"
					offsetText = F_Form[form.TextOffset](removeText, "Статья", "Строка_ID", 10)
				
				in 	offsetText,

		//	ОБОГАЩЕНИЕ  Шаг 5. Добавляем 4 колонки иерархии: Часть, Раздел, Статья, Подстатья (на основе колонки "Номер строки")
			//	Обогащение 1:  добавляем 3 колонки Часть_ID + Раздел_ID + Группа_ID + Уровень_ID
				Add1	 =
				let	tbl  = Clean,
					
					// 	Функция для получения значения Строка_ID
    				GetID = each [Строка_ID],

					// 	добавляем колонки
					AddColumns = F[fxAddMultiColumns](tbl, {
						
						// 	статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
							{ 	"Часть_ID",  each
									if List.Contains ({110..399}, GetID(_)) = true
									then 1
									else 2
								, Int64.Type},

						// 	номер раздела = первому символу номера строки. 0 - не вошедшие в разделы
							{ 	"Раздел_ID", each
									let f=(x as text) => Record.FieldOrDefault([1=1, 2=2, 4=3, 5=4, 6=5, 0=0], x, null),
										Номер = Text.Start(Text.From(GetID(_)),1)
									in f(Номер)
								, Int64.Type},

						// 	если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
							{ 	"Группа_ID", each
									if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1))
									then 1
									else 2
								, Int64.Type},

						// 	уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (= 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
							{	"Уровень_ID", each
										 if List.Contains ({300, 700}, GetID(_)) then 1
									else if List.Contains ({190, 290, 390, 490, 590, 690}, GetID(_)) then 2
									else if Number.Mod(GetID(_), 10) <> 0  and
											not List.Contains ({411}, GetID(_)) then 4 // исключение для строки 411 "Паевые взносы"
									else 3
								, Int64.Type}
						})

				in	AddColumns,
			
			//	Обогащение 2:  добавляем колонки с цифровым ID атрибутов
				Add2	=
				let	tbl  = Add1,
				
				// 	Функция для получения значения Строка_ID
    				GetID = each [Строка_ID],

				// 	Списки строк для проверки
					group_not_calculation = {190, 290, 390, 490, 590, 690, 300, 700, 630, 211, 212, 213, 214, 215, 216},
					group_A1 = {260, 270},
					group_A2 = {210, 250, 280},
					group_A3 = {220, 230, 240, 150, 170},
					group_A4 = {110, 120, 130, 140, 160, 180},
					group_P1 = {632, 633, 634, 635, 636, 637, 638, 639},
					group_P2 = {610, 620, 640, 650, 660, 670, 631},
					group_P3 = {510, 520, 530, 540, 550, 560},
					group_P4 = {410, 420, 430, 440, 450, 460, 470, 480},

				// 	Добавляем колонки с цифровым ID атрибутов
					AddColumn = F[fxAddMultiColumns](tbl, {
					
					// 	1 "Ликвидность_Статус" - выбираем строки, которые НЕ включаются в расчет = false
						{"Ликвидность_Статус", each
								if   List.Contains (group_not_calculation, GetID(_) ) = true
								then false
								else true,
							type logical},

					// 	2 "Ликвидность_Группа" - те что включаются - делим по категориям 4 + 4
						{"Ликвидность_Группа", each
							let id = GetID(_)
							in 	     if List.Contains (group_A1, id) then "А1"
								else if List.Contains (group_A2, id) then "А2"
								else if List.Contains (group_A3, id) then "А3"
								else if List.Contains (group_A4, id) then "А4"
								else if List.Contains (group_P1, id) then "П1"
								else if List.Contains (group_P2, id) then "П2"
								else if List.Contains (group_P3, id) then "П3"
								else if List.Contains (group_P4, id) then "П4"
								else null,
							type text}, // мб сделать null ? было "0"

					// 	3 "Ликвидность_Категория" - категория Ликвидности АП1 ... АП4
						{"Ликвидность_Категория", each
							let id = GetID(_)
							in 	 	 if List.Contains (group_A1 & group_P1, id) then "АП1"
								else if List.Contains (group_A2 & group_P2, id) then "АП2"
								else if List.Contains (group_A3 & group_P3, id) then "АП3"
								else if List.Contains (group_A4 & group_P4, id) then "АП4"
								else "0",
							type text}
					})

				in	AddColumn,
			
			//	Обогащение 3:  добавляем разные колонки из Справочникиов через Join
				Add3	=
				let	from  = Add2,
					// 	создаем Справочники статей из 2 + 1 колонок с рабочими названиями статей
						Спр_ББ_Статья =
							let
								nms      = {"Строка_ID"} & {"Статья", "СтатьяИмя"} & {"СтатьяСтрока"},
								distinct = Table.Distinct(Table.SelectColumns(from, {nms{0}, nms{1}}), {nms{0}} ),
								rename   = Table.RenameColumns(distinct, {{nms{1}, nms{2}}} ),

								// МЕНЯЕМ текст в столбце "СтатьяИмя"
								lst  = {
										{"Налог на добавленную стоимость",                      "НДС"},
										{"Hалог на добавленную стоимость",                      "НДС"}, // в чем отличие?! Дрогичин
										{"НДС по приобретенным товарам, работам, услугам",      "НДС по приобретенным Т,Р,У"},
										{"Долгосрочные активы, предназначенные для реализации", "Долгосрочные активы для реализации"},
										{"Долгосрочные обязательства по лизинговым платежам",   "Долгосрочные обязательства по лизингу"},
										{"Денежные средства и эквиваленты денежных средств",    "Денежные средства и их эквиваленты"},
										{"кредиторская задолженность",                          "КЗ"},
										{"Вложения в долгосрочные активы",                      "Вложения в ДА"},
										{"дебиторская задолженность",                           "ДЗ"}
									},
								replace   = F[Text.ReplaceMulti](rename, nms{2}, lst),
								
								addColumn = Table.AddColumn(replace, nms{3}, each Text.From([Строка_ID]) & " " & [СтатьяИмя])
							in  addColumn,

					// 	добавляем значения колонок с ID (если переключатель IsReportViewGroup = True ?!)
						Join = [
							j  = from,
							j1 = Table.Join(j,	"Часть_ID",     F_Form[spr.BAL][Спр_ББ_Часть],   "Часть_ID"),
							j2 = Table.Join(j1, "Раздел_ID",    F_Form[spr.BAL][Спр_ББ_Раздел],  "Раздел_ID"),
							j3 = Table.Join(j2, "Группа_ID",    F_Form[spr.BAL][Спр_ББ_Группа],  "Группа_ID"),
							j4 = Table.Join(j3, "Уровень_ID",	F_Form[spr.BAL][Спр_ББ_Уровень], "Уровень_ID"),
							j5 = Table.Join(j4, "Строка_ID",    Спр_ББ_Статья, 		 	 		 "Строка_ID") // ? зачем это
						]  [j5]
				in  Join,

		//	ВЫВОД       Шаг 6. Наводим красоту перед выгрузкой
			end = 
				let tbl = Add3,
					sel = F[Row.RemoveNullColumns](tbl, {"На начало","На конец"})
				in	sel
		in 	end,

	rep.OPU   = // ОТЧЕТ О ПРИБЫЛЯХ И УБЫТКАХ --- 
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"opu") () ,

		// 	ЗАГРУЗКА
			Load	=
				let	
					//	предыдущий шаг
						from = Connect,

					// 	выбираем колонки с данными из разных датасетов. Должно остаться 4 колонки.
						company = from[Company]{0},
						nmsOld = if company = "Homel" then {"Column1","Column2","Column3","Column4"}
							else if company = "Mart"  then {"Column1","Column6","Column7","Column11"}
							else if company = "Drog"  then {"Column1","Column2","Column3","Column4"}
							else null,

					//	фильтруем строки в подтаблицах с отчетами по диапазону номеров строк отчета
						to = F_Form[form.SelectRowsID](
							from,
							nmsOld,
							{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
							"Строка_ID", 9, 1000
							)
				in  to,
		//	РАЗВОРОТ
			Expand 	 =
				let tbl = Load,
					
					// 	разворачиваем подтаблицы
						lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
						Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew),

					// 	трансформируем столбцы с указанием типа
						lst1 = {"Company", "DataSet", "form_ID", "Date"},

					//	причесываем колонки //! можнои другие преобразования тут сделать
						tr = Table.TransformColumns(Expand,
							{
								{ lst1{0}, (x)=> x, type text},               	// организация
								{ lst1{1}, (x)=> Number.From(x), Int64.Type},   // датасет
								{ lst1{2}, (x)=> "opu", type text},         	// отчет Сделать join со справончиком отчетов
								{ lst1{3}, (x)=> Date.From(x), Date.Type},      // дата

								{"Статья",         (x)=> x, Text.Type},
								{"Строка_ID",      each try Number.From(_) otherwise null, type number},
								{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
								{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
							}
						)
				in 	tr,
		//	ЧИСТКА
			Clean 	 =
			[	tbl = Expand,
				// 	меняем 0 на null, чтобы 0 в визуализации не появлялся
					replace0ToNull = Table.ReplaceValue(tbl, 0, null, Replacer.ReplaceValue, {"Текущий период", "Прошлый период"}),

				// 	убираем текст внутри скобок + сами скобки + текст после "(" по условию
					removeText1 =
						let
							f=(txt)=>
								if   F[Text.CompareText]({"010","100","±"}, txt) = true 			// проверяем строку на наличие символов
								then Text.BeforeDelimiter(txt, "(", {0, RelativePosition.FromEnd}) 	// оставляем текст ДО, а "(" и ПОСЛЕ убираем
								else txt
						in
							Table.TransformColumns(replace0ToNull, {"Статья", f, type text} ),

				// 	удаляем текст в колонке "Статья" //! сделать универсально
					removeText2 = F[Text.RemoveMulti](removeText1, "Статья", {"В том числе: ", ""}),

				// 	делаем отступ, если номер строки не четный //! сделать неразрывными пробелами и в переменные Q1 Q2
					offsetText = F_Form[form.TextOffset](removeText2, "Статья", "Строка_ID", 10),

					to = offsetText
			] [to],

		//	ОБОГАЩЕНИЕ
			Add	 = // добавляем 3 колонки Группа_ID + Прибыль_ID + ВитдДеят_ID
				let	tbl  = Clean,
				// 	1 добавляем колонки руками
					add1 = F[fxAddMultiColumns](tbl, {
						// 	если первая буква большая, то это 1-Группа, иначе 2-Подгруппа
						{ "Группа_ID", each
							if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true
							then 1
							else 2, Int64.Type},

						// 	1 - статьи со словом "прибыль" в названии
						{ "Прибыль_ID", each
							if List.Contains ({30,60,90,140,150,201,210,240},[Строка_ID]) = true
							then 1
							else 2, Int64.Type},

						// 	1-Текущая, 2-Инвестиционная, 3-финансовая деятельность
						{ "ВидДеят_ID", each
								 if List.Contains ({10..90},  [Строка_ID]) = true then 1
							else if List.Contains ({100..119},[Строка_ID]) = true then 2
							else if List.Contains ({120..139},[Строка_ID]) = true then 3
							else 4, Int64.Type},

						// 	уровень иерархии
						{ "Уровень_ID", each
								if Number.Mod([Строка_ID], 10) <> 0 = true
								then 1
								else 2
							, Int64.Type},

						// 	Доходы, Расходы
						{ "ДохРасх_ID", each
									 if List.Contains ({10,70,101,102,103,104,121,122},[Строка_ID])      = true then "Доходы"
								else if List.Contains ({20,30,40,50,80,111,112,131,132,133},[Строка_ID]) = true then "Расходы"
								else null
							, type text},

						//	Номер года
						{ "Год", each Date.Year([Date]), Int64.Type}
					} ),
			
				// 	2.1 добавляем колонки из Справочникиов
					// нужно ли это тут?! Может в справочники перенести.
					Спр_OPU_Статья =
						let
							Источник        = Table.Distinct(tbl[[Строка_ID],[Статья]], {"Строка_ID"}),
							AddCol_СтрокаID = Table.AddColumn(Источник, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),
							RenameColumns   = Table.RenameColumns(AddCol_СтрокаID,{{"Статья", "СтатьяИмя"}}),

							// Делаем множественную замену в подстроке столбца //! думаю подойдет обычная замена в шаге Clean
							МультиЗамена = [
								dict = List.Buffer(Table.FromValue({ // список списков
										{"товаров, работ, услуг",    			""},
										{"основных средств",                    "ОС"},
										{"нематериальных активов",              "НА"},
										{"долгосрочных активов",                "ДА"},
										{"инвестиционной деятельности",         "ИнвДеят"},
										{"финансовой деятельности",             "ФинДеят"},
										{"отложенных налоговых активов",        "ОНА"},
										{"отложенных налоговых обязательств",   "ОНО"}
									}) [Value]),
								f = (x)=> List.Accumulate(dict, x, (state, current) => Text.Replace(state, current{0}, current{1} )),
								b1   = Table.TransformColumns(RenameColumns, {"СтатьяИмя", f}), //! в двух колонках сразу
								b2   = Table.TransformColumns(b1, {"СтатьяСтрока", f})
								][b2],

							// зачем?
							AddIndexColumn = Table.AddIndexColumn(МультиЗамена, "СтрокаИндекс", 1, 1, Int64.Type)
						in  AddIndexColumn,

				// 	2.2 объединяем таблицу фактов со Справочникиами с формированием дополнительных колонок-срезов
					/*	если БКС то джойнить надо будет и организацию. 
						возмодно сделать надо универсальный справочник: филиалы, предприятия, структурные подразделения
						Join вынести в отдельный шаг со своими функциями */
					Join =
						let tbl = add1,
							j1 = Table.Join(tbl,"ВидДеят_ID", F_Form[spr.OPU][Спр_OPU_ВидДеятельности], "ВидДеятельности_ID"),
							j2 = Table.Join(j1, "Прибыль_ID", F_Form[spr.OPU][Спр_OPU_ВидПрибыли],      "Прибыль_ID"),
							j3 = Table.Join(j2,	"Группа_ID",  F_Form[spr.OPU][Спр_OPU_Группа],          "Группа_ID"),
							j4 = Table.Join(j3, "Строка_ID",  Спр_OPU_Статья,							 "Строка_ID")
						in  j4
				
				in 	Join,

		//	ВЫВОД
			end	 =
				let tbl = Add,
					tbl2 = F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"})
				in	tbl2
		in 	end,
		
	rep.ODDS  = // ОТЧЕТ О ДВИЖЕНИИ ДЕНЕЖНЫХ СРЕДСТВ
		(optional datasetName as nullable text) => 
		let
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"odds") () ,
			
			Load	=
				let	tbl = Connect,
					
					// 	выбираем колонки с данными из разных датасетов
						company = tbl[Company]{0},
						nmsOld = if company = "Homel" then {"Column1","Column2","Column3","Column4"}
							else if company = "Mart"  then {"Column1","Column6","Column7","Column11"}
							else if company = "Drog"  then {"Column1","Column2","Column3","Column4"}
							else null,

					//	фильтруем строки в подтаблицах с отчетами по диапазону номеров строк отчета
						to = F_Form[form.SelectRowsID](
							tbl,
							nmsOld,
							{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
							"Строка_ID", 9, 1000
							)
				in  to,

			Expand  =
				let	tbl = Load,
					// разворачиваем подтаблицы
					lstColumnNew = {"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
					Expand = Table.ExpandTableColumn(tbl, "Step2", lstColumnNew, lstColumnNew ),

					// трансформируем столбцы с указанием типа
					TransformAllColumns = Table.TransformColumns(Expand,
						{
							{ "Company", 	(x)=> x, type text},               	// организация
							{ "DataSet", 	(x)=> Number.From(x), Int64.Type},  // датасет
							{ "form_ID", 	(x)=> "odds", type text},         	// отчет Сделать join со справончиком отчетов
							{ "Date", 		(x)=> Date.From(x), Date.Type},     // дата

							{"Статья",         (x)=> x, Text.Type},
							{"Строка_ID",      each try Number.From(_) otherwise null, type number},
							{"Текущий период", each try Number.From(_) otherwise _, Number.Type},
							{"Прошлый период", each try Number.From(_) otherwise _, Number.Type}
						}
					),

					to = TransformAllColumns
				in to,

			Clean   =
				let	tbl = Expand,
					// меняем 0 на null (чтобы в визуализации не появлялся)
					Replace0ToNull = Table.ReplaceValue(tbl,
						0, null, Replacer.ReplaceValue,
						{"Текущий период", "Прошлый период"}
					),

					// удаляем лишний текст
					removeList  = {"В том числе: ", "В том числе:", " - всего"},
					multiRemove = F[Text.RemoveMulti](Replace0ToNull,"Статья", removeList),

					// трансформируем текст
					transform   = Table.TransformColumns(multiRemove,
						{"Статья", (x)=> [
								txt = F[Text.ReplaceBetweenDelimiters](x, "(",")"), 		// убираем текст в скобках
								to  =  	if   Text.Contains(txt,"Остаток денежных средств") 	// удаляем даты "... на 01.01.2015" и т.д.
										then Text.BeforeDelimiter(txt," на")
										else txt
							] [to]
						}),

					// делаем отступ для подстатей
					addOffset = F_Form[form.TextOffset](transform, "Статья", "Строка_ID", 10),

					endClean = addOffset
				in 	endClean,

			Add1    =
				let	tbl = Clean,
					// добавляем колонки
					AddColumns = F[fxAddMultiColumns](tbl, {

						{ "Группа_ID",  each
								if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true
								then 1
								else 2 , Int64.Type}, // 1- Группа, 2 - Подгруппа

						{ "ВидДеятельности_ID",
								each if List.Contains ({20..40},[Строка_ID]) = true then 1 // "Текущая"
								else if List.Contains ({50..70},[Строка_ID]) = true then 2 // "Инвестиционная"
								else if List.Contains ({80..100},[Строка_ID]) = true then 3 // "Финансовая"
								else 4, Int64.Type}, // "Другое"

						{ "ВидДвижения_ID",
								each if List.Contains ({20..24, 50..55, 80..84},[Строка_ID]) = true then 1 // "Поступление"
								else if List.Contains ({30..34, 60..64, 90..95},[Строка_ID]) = true then 2 // "Выбытие"
								else if List.Contains ({40},[Строка_ID]) = true then 3 // "РезультатТек"
								else if List.Contains ({70},[Строка_ID]) = true then 4 // "РезультатИнв"
								else if List.Contains ({100},[Строка_ID]) = true then 5 // "РезультатФин"
								else if List.Contains ({110},[Строка_ID]) = true then 6 // "РезультатОбщий"
								else 7 /* "Другое" */, Int64.Type},

						{"Год", each Date.Year([Date]), Int64.Type}
					}),

					endAdd1 = AddColumns
				in	endAdd1,

			Add2    =
				let	tbl = Add1,
					Спр_ODDS_Статья =
						let
							// берем две колонки "Строка_ID","Статья"
							from = 	Table.Distinct(Table.SelectColumns(tbl, {"Строка_ID","Статья"}), {"Строка_ID"} ),

							// УДАЛЯЕМ лишний текст
							lst    = {" по отношению к белорусскому рублю", ""},
							REMOVE = F[Text.RemoveMulti](from, "Статья", lst),

							// ДОБАВЛЯЕМ колонку "СтатьяСтрока" соединяя имеющиеся
							AddColumn = Table.AddColumn(REMOVE, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),

							// меняем имя колонки "Статья", чтобы сделать в ней рабочие сокращенные названия статей
							RenameCol = Table.RenameColumns(AddColumn, {{"Статья", "СтатьяИмя"}} ),

							// МЕНЯЕМ текст в подстроке столбца "СтатьяИмя" для удобства работы
							replaceList  = {
									{"текущей, инвестиционной и финансовой деятельности", 	"ТИФ деят"},
									{"от покупателей продукции, товаpов, заказчиков pабот, услуг", 	"покупателей товаров"},
									{"от покупателей материалов и других запасов", 					"покупателей других запасов"},
									{"доходы от участия в уставном капитале других организаций", 	"доходы от дивидендов"},
									{"основных средств, нематериальных активов и других долгосрочных активов", "ОС, НА и других ДА"},
									{"вклады собственника имущества (учредителей, участников)", 	"вклады собственника"},
									{"на выплаты дивидендов и других доходов от участия в уставном капитале организации", "на выплаты дивидендов"},
									{"текущей деятельности",		"ТекДеят"},
									{"инвестиционной деятельности", "ИнвДеят"},
									{"финансовой деятельности", 	"ФинДеят"}/*,
									{"продукции, товаpов, заказчиков pабот, услуг","товаров"}*/
								},

							// заменяем содержимое колонок
							REPLACE1 = F[Text.ReplaceMulti](RenameCol, "СтатьяИмя", replaceList),
							REPLACE2 = F[Text.ReplaceMulti](REPLACE1, "СтатьяСтрока", replaceList)

						in
							REPLACE2,

					Join =
						let j0 = tbl,
						   	j1 = Table.Join(j0,	"ВидДеятельности_ID", F_Form[spr.ODDS][Спр_ODDS_ВидДеятельности],  	"ВидДеятельности_ID"),
							j2 = Table.Join(j1, "Строка_ID",  		  Спр_ODDS_Статья,						 		"Строка_ID"),
							j3 = Table.Join(j2, "ВидДвижения_ID", 	  F_Form[spr.ODDS][Спр_ODDS_ВидДвижения],      	"ВидДвижения_ID")
						in  j3,

					endAdd2 = Join
				in 	endAdd2,

			End     =
				let	tbl  = 	Add2,
					tbl1 = 	F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"}),
					tbl2 = 	if tbl[Company]{0} <> "Mart"
							then 1
							else 2
				in	tbl
		in 	End,

	rep.OSK   = // ОТЧЕТ ОБ ИЗМЕНЕНИИ СОБСТВЕННОГО КАПИТАЛА
		(optional datasetName as nullable text) => 
		/*
			Описание:
				По нормальному этим отчетом не занимался еще. А надо ли?!
				непонятно в каком видже надо представить в конце.
				нужно ди делать анпивот или две колонки очстаток на начало + остаток на конец
				таблица сложная
		*/
		let
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"osk") () ,

			Load	=
				let	tbl = Connect,
					// 	выбираем колонки в подтаблицах с учетом файлов датасета Column1...Column10
						company = tbl[Company]{0},
						nmsOld 	= if company = "Homel" then F_Form[getColumnNames]({1..10})
							 else if company = "Mart"  then F_Form[getColumnNames]({1,5..7,9..14})
							 else if company = "Drog"  then F_Form[getColumnNames]({1..4})
							 else null,
						nmsNew 	= {"Показатель", "Строка_ID", "1", "2","3","4","5","6","7","8"},

					//	фильтруем строки в подтаблицах с отчетами по диапазону номеров строк отчета
						to	= 	F_Form[form.SelectRowsID](tbl, nmsOld, nmsNew, "Строка_ID", 10, 200)
				in 	to,

			Expand  =
				let	tbl = Load,
					// 	список заголовков в подтаблицах отчета - 10 шт.
						lstHeader =
							{
								"Статья", "Строка_ID",
								"Уставный капитал", "Неоплаченная часть уставного капитала",
								"Собственные акции (доли в уставном капитале)", "Резервный капитал", "Добавочный капитал",
								"Нераспределенная прибыль (непокрытый убыток)", "Чистая прибыль (убыток)",
								"Итого"
							},

					// 	разворачиваем таблицу
						Expand =
							Table.ExpandTableColumn(tbl,
								"Step2",
								Table.ColumnNames(Record.Field(tbl{0}, "Step2")), // список имен колонок таблицы из 1-й строки - 10 штук
								lstHeader//  & {"Index"} // зачем эта колонка?
							) // + 1 колонка
				in Expand,

			Clean   =	// на входе 4 технические колонки + 10 колонок отчета + ?1 колонка с индексом
				let	tbl = Expand,
					// 	делаем последние 8 колонок числовыми
						Types =
							Table.TransformColumnTypes(tbl,
								List.Transform(
									List.LastN(Table.ColumnNames(tbl), 8),
									each {_, type number}
								)
							),

					// 	сортируем колонку 5 - это "Строка_ID" (пока не переименовывали колонки) //? Зачем?
						Sort = Table.Sort(Types, {Table.ColumnNames(tbl){5}, Order.Ascending} ),

					// 	удаляем лишний текст в подстроке столбца
						MultiRemove = F[Text.RemoveMulti](Sort, "Статья", 
							{"В том числе: ", "в том числе: "} 
						),

					//  меняем на правильные названия некоторых статей 
						MultiReplace =
							F[Text.MultiReplace](MultiRemove, "Строка_ID", "Статья", 
								[	10 = "Остаток на начало",
									30 = "Корректировки ошибок",
									40 = "Скорректированный остаток",
									50 = "Увеличение собственного капитала - всего",
									100 = "Остаток на начало 1",
									110 = "Остаток на начало 2",
									140 = "Остаток на начало скорректированный",
									150 = "Увеличение собственного капиатала - всего",
									151 = "чистая прибыль",
									161 = "чистая убыток",
									200 = "Остаток на конец"
								]
							)

				in MultiReplace,

			Add  =
				let	tbl = Clean,
				
					// делаем отступ, если номер строки нечетный
						offsetText =
							F_Form[form.TextOffset](tbl, "Статья", "Строка_ID", 10),

					// убираем строки с позапрошлым годом + 2 колонки
						// Remove  = Table.RemoveColumns(offsetText,{"Итого", "Index"}),
						Remove  = Table.RemoveColumns(offsetText,{"Итого"}),
						AddCol  = Table.AddColumn(Remove, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья]),

					// переворачиваем таблицу. Надо ли?!
						Unpivot =
							Table.UnpivotOtherColumns(AddCol,
								{"Company", "DataSet", "form_ID", "Date"} & {"Статья", "Строка_ID", "СтатьяСтрока"},
								"Показатель", "Значение"
							),

					// 	меняем 0 на null (чтобы в визуализации не появлялся) - можно убрать и оставить фильтр идущий следом + <> 0
						Replace0ToNull =
							Table.ReplaceValue(Unpivot,
								0, null, Replacer.ReplaceValue,
								{"Значение"}
							),

					//  удаляем пустые строки
						Select  = 	Table.SelectRows(Replace0ToNull, each ([Значение] <> null)),

					//  Зачем тут менять тип колонки? //!
						Typed   = 	Table.TransformColumnTypes(Select, {"Статья", type text} ),

					// добавляем новые колонки
						AddColumns =
							F[fxAddMultiColumns](Typed, {
								{ 	"Раздел_ID",
										each if List.Contains ({10},[Строка_ID])       = true then 2 // Позапрошлый год" ?!
										else if List.Contains ({20..110},[Строка_ID])  = true then 2 // Прошлый год
										else if List.Contains ({120..200},[Строка_ID]) = true then 3 // Текущий год
										else 4 // "Другое"
									, Int64.Type
								},

								// 	уровень иерархии: 1 - группа, 2 - подгруппа
								{ 	"Уровень_ID", each
										if Number.Mod([Строка_ID], 10) = 0 = true
										then 1
										else 2
									, Int64.Type
								}
							})
				in AddColumns,

			Add2 =
				let	tbl = Add,
					Join =
						let j  = Table.Join(tbl, "Раздел_ID", F_Form[spr.OSK][Спр_OSK_Раздел], "Раздел_ID")
						in  j
				in Join,
			
			End  =
				let	tbl = Add2,
					tbl2 = "тут нет пустых строк пока"
				in	tbl	
		
		in 	End,	
		
	rep.PRIL  = // ПРИЛОЖЕНИЕ К БАЛАНСУ 
		(optional datasetName as nullable text) => 
		let 
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"pril") () ,

			Load    = 
				F_Form[form.SelectRowsID](Connect,
					{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7"},
					{"Статья", "Строка_ID", "1", "2", "3", "4", "5"},
					"Строка_ID", 0, 2000),

			Expand  = 
				let	tbl = Load,
					
					// 	список колонок итоговой таблицы, сквозной, от 2 до 5 в каждом разделе
						// lstNewColNаmes = {"Статья", "Строка_ID", "1", "2", "3", "4", "5"},

					//  список заголовков в подтаблицах отчета -  шт.
						lstHeader = Table.ColumnNames(tbl[Step2]{0}), // 7 колонок

					// 	разворачиваем подтаблицы (список столбцов без первой колонки "Date", которой внутри нет)
						EXPAND = Table.ExpandTableColumn(tbl, "Step2", lstHeader,  lstHeader),

					//  список всех колонок после разворота
						nms = List.Skip(Table.ColumnNames(EXPAND), 4),

					// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
					Transform = Table.TransformColumns(EXPAND,
						{
							{ nms{0}, (x)=> x, type text},  // Статья
							{ nms{1}, (x)=> x, Int64.Type}, // Строка
							{ nms{2}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{4}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{5}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
							{ nms{6}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
						}
					),
					
					to = Transform
				in	to,

			Clean   = 
				let	tbl = Expand,

					// 	удаляем лишнее
						replace = F[Text.ReplaceMulti](tbl, "Статья", {
								{"в т.ч.: ", ""},
								{" (сумма стр.1010-1130)", ""},
								{" (сумма стр.1010-1150)", ""},
								{" (сумма стр. 1210, 1220)", ""},
								{" - всего", ":"},
								{" - ВСЕГО", ":"},
								{" (авансы выданные, предоплата)", ""}
							}),

					// 	делаем отступ для подстатей
						addOffset = F_Form[form.TextOffset](replace, "Статья", "Строка_ID", 10),

					//  меняем названия некоторых статей
					//  	надо что то другое придумать, чтобы разные названия полей в разные года свести
					//  	к одному значени. Например было "Здания" и "Соружения", стало "Здания и сооружения"
					//  	в начале может поменять/объединить/удалить
						Replace =
							Table.ReplaceValue(addOffset, (x)=>x[Статья],
								each
									if Date.Year([Date]) = 2020 or Date.Year([Date]) = 2021
									then Record.Field([11 = "    здания "], Text.From([Строка_ID])) // и сооружения
									else
										if Date.Year([Date]) = 2016
										then Record.Field([11 = "    здания и сооружения"], Text.From([Строка_ID]))
											else Record.Field([11 = _[Статья] ], Text.From([Строка_ID]) )
									,
								Replacer.ReplaceValue, {"Статья"}
							),

					// 	разворачиваем
						Unpivot =
							Table.UnpivotOtherColumns(Replace,
								{"DataSet","Company","form_ID", "Date", "Статья","Строка_ID"},
								"Заголовок", "Значение"
							),

					to = Unpivot
				in to,

			Add  	= 
				let	tbl = Clean,
					// 	добавляем колонки
						AddedColumns = F[fxAddMultiColumns](tbl, {
							{ "Раздел_ID",
								each if List.Contains ({1..36},      [Строка_ID]) = true then 1	// Раздел 1
								else if List.Contains ({50..65},     [Строка_ID]) = true then 2	// Раздел 2
								else if List.Contains ({70..70},     [Строка_ID]) = true then 3	// Раздел 3
								else if List.Contains ({80..95},     [Строка_ID]) = true then 4	// Раздел 4
								else if List.Contains ({140..171},   [Строка_ID]) = true then 5	// Раздел 5
								else if List.Contains ({1010..1320}, [Строка_ID]) = true then 6	// Раздел 6
								else if List.Contains ({1401..1411}, [Строка_ID]) = true then 7	// Раздел 7
								else 9
							, Int64.Type},

							{ "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья] , type text},

							//
							{ "Заголовок_ID",
								each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок]))
							, Int64.Type}
						}),

					// 	добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
						Join =
							let tbl = AddedColumns,
								j1  = Table.Join(tbl, "Заголовок_ID", F_Form[spr.PRIL][Спр_Заголовки],   "Заголовок_ID"),
								j2  = Table.Join(j1,  "Раздел_ID",    F_Form[spr.PRIL][Спр_Разделы],     "Раздел_ID")
							in  j1,

					//	возвращаем
						to = Join
				in to,

			End 	= 
				let to = Add
				in	to
		
		in	End	,

	rep.OFR   = // ОТЧЕТ О ФИНАНСОВЫХ РЕЗУЛЬТАТАХ (Форма №12) - был 12ф, стал 12.
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"ofr") () ,

			Load    =
					F_Form[form.SelectRowsID](Connect,
						{"Column1","Column2","Column3","Column4"},
						{"Статья", "Строка_ID", "Текущий период", "Прошлый период"},
						"Строка_ID", 0, 1000),

			Expand  =
				let	tbl = Load,
				/*  Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
					надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0}) */
					//  список заголовков в подтаблицах отчета -  шт.
						lstHeader = Table.ColumnNames(tbl[Step2]{0}),

					// разворачиваем таблицу
						expand = Table.ExpandTableColumn(tbl, "Step2", lstHeader,lstHeader)

				in	expand,

			Clean   =
				let	tbl = Expand,
					//	список столбцов (прOPU_скаем первые 4 технические столбца)
						nms = List.Skip(Table.ColumnNames(tbl), 4),

					// 	типизируем таблицу
						transform = Table.TransformColumns(tbl,
							{
								{ nms{0}, (x)=> F[Text.ReplaceBetweenDelimiters](x," (",")"), type text},
								{ nms{1}, (x)=> try Number.From(x) otherwise null, Int64.Type },
								{ nms{2}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type},
								{ nms{3}, (x)=> if  Number.From(x) = 0 then null else Number.From(x), Int64.Type}
							}
						),

					//  меняем названия некоторых статей
						Replace =
							Table.ReplaceValue(transform, (x)=>x[Статья],
								each Record.Field([
										29 = "Средства из бюджета на покрытие убытков, на возмещение затрат",
										32 = "   на погашение основного долга",
										33 = "   на погашение процентов по кредиту, займу"
									], Text.From([Строка_ID])),
								Replacer.ReplaceValue, {"Статья"}
							),

						Replace2 = F[Text.ReplaceMulti](Replace, "Статья", {
							{"налог на добавленную стоимость", "НДС"},
							{"реализации продукции, товаров, работ, услуг", "реализации"}
						}),

					to = Replace2
				in 	to,

			Add  =
				let	tbl = Expand,
					AddCol  = Table.AddColumn(tbl, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья], type text),
					to = AddCol
				in to,

			End     =
				let	tbl  = F[Table.TypedField](Add),
					tbl2 = F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"})
				in	tbl2
		in End,

	rep.OSR   = // ОТЧЕТ О СОСТОЯНИИ РАСЧЕТОВ
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"osr") () ,

			Load    =
				let	tbl = Connect,
					load = // Чистим подтаблицы = лист с отчетом. Для Гомеля "Step 1" нет, так как каждый отчет в отдельном файле. from = заменяет его
						Table.AddColumn(tbl,
							"Step2", // создаем колонку с подтаблицами
								(x)=>
									let
									// фильтруем колонки
										select = Table.SelectColumns(x[Content], {"Column1", "Column2", "Column3", "Column4"} ),

									// чистим текст во всей таблице (" " двойные пробелы, табуляцию,)
										clear   = F[Table.ClearAll](select," #(lf)#(00A0)",false,false),

									// список имен колонок подтаблицы
										ColNames = Table.ColumnNames(clear),

									// меняем тип данных второй колонки - "НомерСтроки" для дальнейшей фильтрации
										tr = Table.TransformColumns(clear,{
												{ ColNames{1}, (x)=> try Number.From(x) otherwise null, Int64.Type}
											} ),

									// фильтруем строки в колонке ПО ИМЕНИ колонки, не по [полю]
									// Table.SelectRows(RenameColumns, each {ColNames{1}} <> null), не работает!
										f=(x)=> (x <> null) and (x > 0) and (x < 1000),
										SelectRow = Table.SelectRows (tr,
												(row) => f (Record.Field (row, ColNames{1} ) )
											),

									// список колонок итоговой таблицы
										lstNewColNаmes = {"Статья", "Строка_ID", "Всего", "Просроченная"},

									// меняем названия колонок кроме первой (первая Date)
										rename = Table.RenameColumns(SelectRow, List.Zip({ColNames, lstNewColNаmes}) )

								in rename
							),
					remove   = Table.RemoveColumns(load, {"Content"})
				in 	remove,

			Expand	=
				let	tbl = Load,
							lst =Table.ColumnNames(tbl[Step2]{0}), // {"Статья", "Строка_ID", "Всего", "Просроченная"},

						// Разворачиваем подтаблицы. Список столбцов без первой колонки "Date", которой внутри нет
						// надежней будет колонки определять сразу до разворота  ColNames = Table.ColumnNames(from2[Step2]{0})
							EXPAND = Table.ExpandTableColumn(tbl,  "Step2", lst, lst),

						// меняем тип колонок с трансформацией. Можно делать это и в подтаблицах, до разворота
							Transform = Table.TransformColumns(EXPAND,
								{
									{ lst{0}, (x)=> x, type text},
									{ lst{1}, (x)=> try Number.From(x) otherwise null, Int64.Type},
									{ lst{2}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type},
									{ lst{3}, (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
								}
							)
					in 	Transform,

			Clean	=
				let	tbl = Expand,
						// Убираем текст внутри скобок и сами скобки.
						// Проверяем колонку x на наличие символ из списка. Если есть - оставляем текст ДО, а "(" и ПОСЛЕ нее все убираем
							УбираемТекстПослеСкобки =
								Table.TransformColumns ( tbl, // можно просто убрать все в скобках!
									{ "Статья", (x) =>
										if   F[Text.CompareText]({"50","66","70","09","76","60","69","68","202"}, x) = true
										then Text.BeforeDelimiter(x, "(", {0, RelativePosition.FromEnd})
										else x, type text }
								),

						sel = Table.SelectRows(УбираемТекстПослеСкобки, each [Строка_ID] < 200),

						Remove = F[Text.RemoveMulti](sel, "Статья",
							{" - всего", "Из стр. 101 - ", "Из стр. 110 - ", "Из стр. 101 графы 2 - ",
							" (в отпускных ценах)"," (договоров)","из нее ","из них "})

				in 	Remove,

			Add	=
				let	tbl = Clean,
					addOffset = Table.ReplaceValue(tbl, // сделать функцию. Ждобавить тип функций - проверка типа is ...
							each [Статья],
							each if not List.ContainsAny ({Text.Start([Статья],1) },{"А".."Я"},Text.Contains) = true
								then "    " & [Статья]
								else [Статья],
							Replacer.ReplaceValue,{"Статья"}
						),
					addColumn = Table.AddColumn(addOffset, "СтатьяСтрока", each Text.From([Строка_ID]) & " " & [Статья])
				in  addColumn,
			
			End     =
				let	tbl  = Add,
					tbl2 = F[Row.RemoveNullColumns](tbl, {"Всего"})
				in	tbl2
		
		in End,

	rep.BKS3  = // ФОРМА 3 БКС "Отчет о доходах, расходах и финансовых результатах в торговле" (готово для BKS и HOMEL) 
		/* = F_Form[rep.BKS3] ([ДатаCетИмя="Homel"]) */
		(optional options as record) =>	// (optional datasetName as nullable text) => 
		let 
		//	АРГУМЕНТЫ   Шаг 1. Список аргументов 
           	datasetName = 
				let value = Record.FieldOrDefault(options, "ДатаCетИмя", null) 
				in if value = 0 then null else value,
		// 	ПОДКЛЮЧЕНИЕ Шаг 2. Connect 	 - на входе спиок файлов с техколонками в т.ч. колонка Company
			Connect = 	
				F_Form[Connect_]([
					ДатаCетИмя 		= F_Form[GetName](datasetName), //"Homel" - активное имя датасета сделать проверку на null = ?? 0
					ОтчетИмя 		= "v_3",	//"v_3" = form_ID, если фильтровать отчеты по имени файла в общей папке с файлами отчетов
					МаскаИмениПапки = null,		// маска имени папки, если фильтроватиь отчеты по папкам с отчетами
					boolfilterSheet = false,
					НомерШага 		= null	
				]),
		// 	ЗАГРУЗКА    Шаг 3. SelectRow - обрабатываем подтаблицы с сортировкой строк 
			Load    = // по БКС все нормально, а по Homel надо делать обработку подтаблиц наверное раньше в Connect еще 
				let	tbl = Connect,
					// 	функции обработки отдельных датасетов
						fxBKS=(x)=> [
							//	выбираем один период или одно предприятие для проверки //! пока этот кусок не обработан
								filter =
									let 
										/* from = 
											if 	prm_LoadMode = 0 
												then tbl
												else if 1=1
													//	оставляем данные за 1 год по ВСЕМ предприятим
													then Table.SelectRows(tbl, (x)=> x[Date] = #date(2021, 12, 1) ) // данные за 1 год
													// 	оставляем данные по 1 предприятию за ВСЕ периоды
													else Table.TransformColumns(tbl, {"Content", (x)=>Table.SelectRows(x, each [Name]="1")})
										*/
										from  = tbl,
										from1 = Table.TransformColumns(tbl, {"Content", (x)=>Table.SelectRows(x, each [Name]="1")}),
										from2 = Table.SelectRows(tbl, (x)=> x[Date] = #date(2021, 12, 1) )
									in  from,
							// 	раскрываем предприятия
								expand = Table.ExpandTableColumn(x, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
							// 	фильтруем колонки (оставляем 9 колонок (макс кол-во с данными для этого отчета), т.к. в некоторых отчетах есть больше колонок и вылазит ошибка
							// 	фильтруем строки (название листа = число)
								SelCol = Table.SelectColumns(expand, List.Generate(()=> 1, (x)=> x<=9, (x)=>x+1, (x)=>"Column" & Text.From(x) )),        // оставляем 9 колонок
								RemCol = Table.RemoveColumns(SelCol, {"Column1"}),  // убираем первую колонку с нумерацией статей
								SelRow = Table.SelectRows(RemCol, each try Value.Is( Number.From([Column3]), type number) = true otherwise null),   // оставляем строки с числов в названии

							// 	меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
								lstOld = List.Buffer(Table.ColumnNames(SelRow)),                // старый список имен столбцов из подтаблиц (все одинаковые)
								lstNew = {"Статья", "Строка_ID", "1", "2", "3", "4", "5", "6"}, // новый список имен столбцов
								rename = Table.RenameColumns(SelRow, List.Zip( {lstOld,lstNew} )),

							// 	чистим (меняем знаки из списка на "") от пробелов, переносов строки и др. Лучше сейчас, до разворота, быстрее будет
							// 	если в ячейке null, то ошибка поэтому try-otherwise
								k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
								transform = Table.TransformColumns(rename,{"Статья", k})
							]  [transform],

						fxHomel= [
							//	выбираем один период или одно предприятие для проверки, по умолчанию все
								from = 	
									let	from  = tbl,
										from1 = Table.SelectRows(tbl, (x)=> x[Date] = #date(2020, 12, 31) )
									in  from,
							//	фильтруем: строки - по номеру ID строки, колонки - по порядковому номеру колонки
								to   =	F_Form[SelectRow_]([
										tbl 	 = tbl,
										nmsOld 	 = F_Form[getColumnNames]({2,3,4,5,6,7,8,9}),
										nmsNew 	 = {"Статья", "Строка_ID", "1", "2", "3", "4", "5", "6"},
										colName  = "Строка_ID",
										kodStart = 10000,
										kodEnd   = 99999
									])
							]  [to],
					
					// 	определяем имя датасета - смотрим первую строку в отсортированной по датасету таблице
						company = tbl[Company]{0},	
					// 	выбираем колонки в подтаблицах с учетом файлов датасета
						//  Применяем функуцию исходя из значений столбца [Company]
							// 	functionMap = [BKS = fxBKS, Homel = fxHomel],
							// 	Return = 
							// 		Table.AddColumn(tbl, "Data", each
							// 			let func = Record.FieldOrDefault(functionMap, [Name], (_)=>_)
							// 			in  func([Content2])
							// 		)
				
						from = 
							if 	    company = "BKS"   then Table.TransformColumns(tbl, {"Data", fxBKS})
							else if company = "Homel" then fxHomel
							else tbl
					
					in  from,

		//	РАЗВОРОТ    Шаг 4. Разворачиваем и типизируем 4 + 4 колонки
			Expand	= 
				let	tbl     = Load,
					fxBKS   =
						let
							// 	разворачиваем и переворачиваем таблицу
								Expand2 = Table.ExpandTableColumn(tbl, "Data", Table.ColumnNames(tbl{0}[Data]),Table.ColumnNames(tbl{0}[Data])),
								nms = {"Company", "DataSet", "form_ID", "Date"},
								lst = {"Предприятие_ID", "Статья", "Строка_ID"},
								Unpivot = Table.UnpivotOtherColumns(Expand2, nms & lst, "Заголовок", "Значение"),

							// меняем тип колонок с трансформацией
							// Статья - чистим текст от пробелов внутри и снаружи и непечатных символов
								TransCol = Table.TransformColumns(
										Unpivot,{
											{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
											{ "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
											{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
											{ "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
											{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
										} )
						in 	tbl,

					fxHomel = //(x)=>	если это делать в подтаблицах
						let
							// 	разворачиваем подтаблицы и переворачиваем таблицу
								Expand2 = Table.ExpandTableColumn(tbl, "Step2", Table.ColumnNames(tbl{0}[Step2])),
								nms = {"Company", "DataSet", "form_ID", "Date"},
								lst = {"Статья", "Строка_ID"},
								Unpivot = Table.UnpivotOtherColumns(Expand2, nms & lst, "Заголовок", "Значение"),

							// 	меняем тип колонок с трансформацией
								TransCol = Table.TransformColumns(
									Unpivot,
									{
										{ "Статья",     (x)=> Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
										{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
										{ "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
										{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
									} )
						in 	TransCol,
					
					from 	= 
						if 	    datasetName = "BKS"   then Table.TransformColumns(tbl, {"Data", fxBKS})
						else if datasetName = "Homel" then fxHomel
						else tbl
				in  from,
		//	ЧИСТКА	    Шаг 5. Трансформируем колонки (чистим, меняем, обогащаем)
			Clean	= 
				let	tbl = Expand,
					// 	удаляем текст внутри скобок + скобки и оставляем текст в скобках по маске
						removeText1 = F[Text.RemoveIntoBracket](tbl, "Статья", {"убыт"}),

					// 	удаляем текст по списку
						removeText2 = F[Text.RemoveMulti](removeText1, "Статья",
							{
								"в т. ч. ",
								"в том числе: ",
								" райпо между ЧУП",
								" от оператора вторичных материальных ресурсов",
								" - всего"
							}
						),

					//  функция добавления неразрывных пробелов
						Q = (x)=> Text.Repeat( Character.FromNumber(160), x),
					
					//  меняем названия статей
						removeText3 =
							Table.ReplaceValue(removeText2, (x)=>x[Статья],
								each Record.Field([
										30100 = Q(3) & "выручка от реализации за - внутрисистемной реализациии",
										38001 = "Налоги и сборы в стоимости товаров, переданных филиалами",
										10300 = "Расходы на аренду",
										10400 = "Расходы на хранение",
										10500 = "Потери товаров",
										10600 = "Расходы на ГСМ в общепите",
										11000 = "Расходы на соцстрахование",
										11600 = "справочно: расходы по завозу товаров включенных в цену",
										11700 = "справочно: проценты за пользование кредитами банка",
										40100 = Q(3) & "транзитные поставки",
										42100 = Q(3) & "продажа пайщикам ",
										42200 = Q(3) & "продажа комиссионно (за - вознаграждения)",
										42300 = Q(3) & "продажа конфиската  (за - вознаграждения)",
										44000 = "Внутрихозяйственный оборот – отпуск собственной продукции",
										46000 = "Внутрихозяйственный оборот – отпуск товаров заготовок",
										60000 = "Остаток товаров по учетным ценам:",
										60100 = Q(3) & "торговая наценка на остаток товаров",
										60200 = Q(3) & "НДС на остаток товарово",
										60300 = Q(3) & "запасы по покупным ценам",
										60310 = Q(6) & "оплаченные запасы",
										60320 = Q(6) & "запасы, полученные на условиях оплаты по мере реализации",
										53000 = "Задолженность на конец отчетного периода",
										52000 = "Получено в погашение задолженности с начала года"
									], Text.From([Строка_ID])),
								Replacer.ReplaceValue, {"Статья"}
							)

				in 	removeText3,
		//	ОБОГАЩЕНИЕ  Шаг 6. Добавляем 4 колонки иерархии: Часть, Раздел, Статья, Подстатья (на основе колонки "Номер строки")
			AddCol	= 
				let	tbl = Clean,
					// 	добавляем колонки с цифровым ID атрибутов
						AddColumns = F[fxAddMultiColumns](tbl, {
							// 	раздел
								{ "Раздел_ID",     each Number.From(Text.Start(Text.From([Строка_ID]), 1)), Int64.Type},

							// 	номер строки + название статьи
								{ "СтатьяСтрока",  each Text.From([Строка_ID]) & " " & Text.From([Статья]) , type text},

							//	уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
								{ "Заголовок_ID",  each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type},

							// 	если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
								{ "Группа_ID", each
									     if Text.End(Text.From([Строка_ID]), 3) = "000" then 1
									else if Text.End(Text.From([Строка_ID]), 2) = "00"  then 2
									else 3, 
								Int64.Type},

							// 	уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (равно 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
								{ "Уровень_ID", each
									if  [Раздел_ID] = 3 then
										if Number.Mod([Строка_ID], 1000) <> 0  then 2
										else 1
									else if [Раздел_ID] = 1 then
										if Number.Mod([Строка_ID], 100) <> 0  then 2
										else 1
									else 2
								, Int64.Type}
							}),

					// 	добавляем колонки из названиями заголовком и разделов из справочников
						JoinBKS =
							let Join_Заголо  = Table.Join(AddColumns, 	"Заголовок_ID",    	F_Form[spr.BKS3][Спр_Заголовки],   "Заголовок_ID"),
								Join_Раздел  = Table.Join(Join_Заголо,  "Раздел_ID",     	F_Form[spr.BKS3][Спр_Разделы],     "Раздел_ID"),
								Join_Предпр  = Table.Join(Join_Раздел, 	"Предприятие_ID", 	F_Form[spr.BKS] [Спр_Предприятия], "Предприятие_ID")
							in  Join_Предпр,
							
						JoinHomel =
							let Join_Заголо  = Table.Join(AddColumns, 	"Заголовок_ID",    	F_Form[spr.BKS3][Спр_Заголовки],   "Заголовок_ID"),
								Join_Раздел  = Table.Join(Join_Заголо,  "Раздел_ID",     	F_Form[spr.BKS3][Спр_Разделы],     "Раздел_ID")
							in  Join_Раздел,	
							
						from = 
								 if	datasetName = "BKS"   then JoinBKS
							else if datasetName = "Homel" then JoinHomel
							else JoinHomel	
				in	from,

		//	ВЫВОД       Шаг 7. Наводим красоту перед выгрузкой
			End = 
				let	tbl = AddCol,
					to 	= Table.SelectRows(tbl, each
							[ЗаголовокИмя] <> "Всего по торговле" and 
							// [ЗаголовокИмя] <> "Итого" and // оставил для сводной таблицы в файла Excel "ПробаБаланс"
							[Строка_ID]    <> 11500 and 
							[Значение]     <> null
						)
				in	to
		
		in 	End,

	rep.BKS7  = // форма 7 БКС "Отчет о себестоимости производства" (готово для BKS и HOMEL)
		(optional options as record) =>	// (optional datasetName as nullable text) => 
		let 
		//	АРГУМЕНТЫ   Шаг 1. Список аргументов 
           	datasetName = 
				let value = Record.FieldOrDefault(options, "ДатаCетИмя", null) 
				in if value = 0 then null else value,
		// 	ПОДКЛЮЧЕНИЕ Шаг 2. Connect 	 - на входе спиок файлов с техколонками в т.ч. колонка Company
			/* let 	from  =	F_Form[Connect](F_Form[GetName](datasetName),"v_7") () 
					 	from1 = Table.SelectRows(from,  (x)=> x[form_ID] = "v_7"),
					 	from2 = Table.SelectRows(from1, (x)=> x[Date] = #date(2021, 12, 1) ),
						from3 = Table.TransformColumns(from2, {"Content", (x)=>Table.SelectRows(x, each [Name]="1")})
				in  from, */
			Connect = 	
				F_Form[Connect_]([
					ДатаCетИмя 		= F_Form[GetName](datasetName), //"Homel" - активное имя датасета сделать проверку на null = ?? 0
					ОтчетИмя 		= "v_7",	//"v_3" = form_ID, если фильтровать отчеты по имени файла в общей папке с файлами отчетов
					МаскаИмениПапки = null,		// маска имени папки, если фильтроватиь отчеты по папкам с отчетами
					boolfilterSheet = false,
					НомерШага 		= null	
				]),
		// 	ЗАГРУЗКА    Шаг 3. SelectRow - обрабатываем подтаблицы с сортировкой строк 
			Load    = // по БКС все нормально, а по Homel надо делать обработку подтаблиц наверное раньше в Connect еще 
				let	tbl = Connect,
					// 	функции обработки отдельных датасетов
						fxBKS=(x)=> [
							// фильтруем колонки (оставляем 10 колонок (макс кол-во с данными для этого отчета), т.к. в некоторых отчетах есть больше колонок и вылазит ошибка
								countCol = 12, // максимальное число колонок со значениями в отчете
								SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=countCol, (x)=>x+1, (x)=>"Column" & Text.From(x) )),
								// RemCol = Table.RemoveColumns(SelCol, {"Column1"}),  // убираем первую колонку с нумерацией статей в этом отчете этого нет

							// фильтруем строки (название листа = число) сделать отдельную колонку с переменной колонки с кодом строки может по названию
								// убрать нули в начале номера строки ли не нажо, может с кодами мин и макс надо. Ошибки появляются! Это важный критерий сортировки!!!
								SelRow = Table.SelectRows(SelCol, each
									try
										Value.Is( Number.From([Column2]), type number) = true and
										Number.From([Column2]) < 100000 and
										Number.From([Column2]) > 9
									otherwise null
								),   // оставляем строки с числов в названии

							// меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
								lstOld = List.Buffer(Table.ColumnNames(SelRow)),                // старый список имен столбцов из подтаблиц (все одинаковые)
								lstNew = {"Статья", "Строка_ID", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"}, // новый список имен столбцов
								rename = Table.RenameColumns(SelRow, List.Zip( {lstOld,lstNew} )),

							// чистим (меняем знаки из списка на "") от пробелов, переносов строки и др. Лучше сейчас, до разворота, быстрее будет
								// если в ячейке null, то ошибка поэтому try-otherwise
								k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
								transform = Table.TransformColumns(rename,{"Статья", k})
							]  [transform],

						fxHomel= [
							//	фильтруем: строки - по номеру ID строки, колонки - по порядковому номеру колонки
								to   =	F_Form[SelectRow_]([
										tbl 	 = tbl,
										nmsOld 	 = F_Form[getColumnNames]({1,2,3,4,5,6,7,8,9,10,11,12}),
										nmsNew 	 = {"Статья", "Строка_ID", "1", "2", "3", "4", "5", "6", "7", "8", "9", "10"},
										colName  = "Строка_ID",
										kodStart = 9,
										kodEnd   = 100000
									])
							]  [to],
					
					// 	выбираем колонки в подтаблицах с учетом файлов датасета и имени датасета
						from = 
							if 	datasetName = "BKS"   then 
								let expand = Table.ExpandTableColumn(tbl, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"})
								in 	Table.TransformColumns(expand, {"Data", fxBKS})
							
							else if datasetName = "Homel" then fxHomel
							else tbl
					
					in  from,

		//	РАЗВОРОТ    Шаг 4. Разворачиваем и типизируем 4 + 4 колонки
			Expand	= 
				let	tbl     = Load,
					fxBKS   =
						let
							// 	разворачиваем и переворачиваем таблицу
								Expand2 = Table.ExpandTableColumn(tbl, "Data", Table.ColumnNames(tbl{0}[Data]),Table.ColumnNames(tbl{0}[Data])),
								nms = {"Company", "DataSet", "form_ID", "Date"},
								lst = {"Предприятие_ID", "Статья", "Строка_ID"},
								Unpivot = Table.UnpivotOtherColumns(Expand2, nms & lst, "Заголовок", "Значение"),

							// меняем тип колонок с трансформацией
							// Статья - чистим текст от пробелов внутри и снаружи и непечатных символов
								TransCol = Table.TransformColumns(
										Unpivot,{
											{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
											{ "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
											{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
											{ "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
											{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
										} )
						in 	tbl,

					fxHomel = //(x)=>	если это делать в подтаблицах
						let
							// 	разворачиваем подтаблицы и переворачиваем таблицу
								Expand2 = Table.ExpandTableColumn(tbl, "Step2", Table.ColumnNames(tbl{0}[Step2])),
								nms = {"Company", "DataSet", "form_ID", "Date"},
								lst = {"Статья", "Строка_ID"},
								Unpivot = Table.UnpivotOtherColumns(Expand2, nms & lst, "Заголовок", "Значение"),

							// 	меняем тип колонок с трансформацией
								TransCol = Table.TransformColumns(
									Unpivot,
									{
										{ "Статья",     (x)=> Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
										{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
										{ "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
										{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
									} )
						in 	TransCol,
					
					from 	= 
						if 	    datasetName = "BKS"   then Table.TransformColumns(tbl, {"Data", fxBKS})
						else if datasetName = "Homel" then fxHomel
						else tbl
				in  from,
		//	ЧИСТКА	    Шаг 5. Трансформируем колонки (чистим, меняем, обогащаем)
			Clean	= 
				let	tbl = Expand,
					//  исходнst параметры
						removeText = tbl,
						nms = "Статья",
						
					// 	удаляем скобки за исключением некоторых
						// removeText1 = F[Text.RemoveIntoBracket](removeText, "Статья", null), // {"убыт","раб"}
						removeText1 = F[Text.DeletBracket](removeText, nms),

					// 	удаляем текст за исключением некоторых
						removeText2 = F[Text.ReplaceMulti](removeText1, nms, {
								{"п/ф-ты","ПФ"},
								{"(работ, услуг)","(Р,У)"}
						}),

					// 	удаляем текст по списку
						removeText3 = F[Text.RemoveMulti](removeText2, nms, { "- всего", "в т. ч. ", "+-"} ),

					//  меняем названия статей
						removeText4 =
							Table.ReplaceValue(removeText3, (x)=>x[Статья],
								each Record.Field([
										20 = "Покупные изделия, ПФ, услуги производственного характера",
										80 = "Отчисления в ФСЗН",
										90 = "Налоги, сборы и пошлины в себестоимости продукции",
										160 = "Фактические затраты, связанные с производством",
										170 = "Произведено продукции в отпускных ценах",
										240 = "Амортизация ОС и НА",
										330 = "Налоги и сборы в выручке от реализации",
										320 = "Отпуск готовой продукции своим предприятиям",

										310 = "Выручка от реализации",
										410 = "Выручка от реализации",

										350 = "Валовая прибыль",
										440 = "Валовая прибыль",

										400 = "Стоимость ГП, переданной между филиалами",
										420 = "Налоги и сборы в выручке от реализации",

										430 = "Производственная себестоимость реализации",
										340 = "Производственная себестоимость реализации",

										380 = "Прибыль от реализации",
										470 = "Прибыль от реализации",

										540 = "Амортизация ОС и НА"

									], Text.From([Строка_ID])),
								Replacer.ReplaceValue, {"Статья"}
							)

				in 	removeText4,

		
		//	ОБОГАЩЕНИЕ  Шаг 6. Добавляем _ колонки: _ (на основе колонки "Номер строки")
			AddCol	= 
				let	tbl = Clean,
					// 	Функция для получения значения Строка_ID
						GetID = each [Строка_ID],

					// 	добавляем колонки с цифровым ID атрибутов
						AddColumns = F[fxAddMultiColumns](tbl, {
							{ "Раздел_ID",  each
								let id = GetID(_)
								in	     if List.Contains ({10..170}, id)  = true then 1
									else if List.Contains ({210..260}, id) = true then 2
									else if List.Contains ({310..406}, id) = true then 3
									else if List.Contains ({410..496}, id) = true then 4
									else if List.Contains ({510..560}, id) = true then 5
									else 6
								, Int64.Type},

						// 	номер строки + название статьи
							{ "СтатьяСтрока",  each Text.From(GetID(_)) & " " & Text.From([Статья]) , type text},

						// 	уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
							{ "Заголовок_ID", each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type},

						// 	если последняя цифра 0, то группа (1), иначе подгруппа (2)
							{ "Группа_ID", each if Text.End(Text.From(GetID(_)), 1) = "0" then 1 else 2, Int64.Type},

						// 	вид деятельности
							{ "Вид_ID", each
										if List.Contains ({1,2,3}, [Раздел_ID]) then "1 Основные виды"
									else if List.Contains ({4,5},   [Раздел_ID]) then "2 Другие виды"
									else "пусто"
								, type text},

						//  уровень иерархии = Группа_ID = > оставить надо одно
							{ "Уровень_ID", each
									if Number.Mod(GetID(_), 10) <> 0  then 2
									else 1
								, Int64.Type},

						// 	группы показателей
							{ "Метрика_ID", each
								let id = GetID(_),
									p1 = {310,350,380, 410,440,470},
									p2 = {340,360,370, 430,450,460, 210,220,510,520 },
									p3 = {40,50,110}
								in		 if List.Contains (p1, id) then "1 Прибыль"
									else if List.Contains (p2, id) then "2 Затраты"
									else if List.Contains (p3, id) then "3 Прочие"
									else "4 пусто"
								, type text}
						}),

						// добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
							JoinBKS   =
								let Join_Заголо  = Table.Join(AddColumns, 	"Заголовок_ID",    	F_Form[spr.BKS7][Спр_Заголовки],   "Заголовок_ID"),
									Join_Предпр  = Table.Join(Join_Заголо, 	"Предприятие_ID", 	F_Form[spr.BKS] [Спр_Предприятия], "Предприятие_ID"),
									Join_Раздел  = Table.Join(Join_Предпр,  "Раздел_ID",     	F_Form[spr.BKS7][Спр_Разделы],     "Раздел_ID")
								in  Join_Раздел,
							
							JoinHomel =
								let Join_Заголо  = Table.Join(AddColumns, 	"Заголовок_ID",    	F_Form[spr.BKS7][Спр_Заголовки],   "Заголовок_ID"),
									Join_Раздел  = Table.Join(Join_Заголо,  "Раздел_ID",     	F_Form[spr.BKS7][Спр_Разделы],     "Раздел_ID")
								in  Join_Раздел,	
							
							from = 
  								     if	datasetName = "BKS"   then JoinBKS
								else if datasetName = "Homel" then JoinHomel
								else JoinHomel

					in	from,

		//	ВЫВОД       Шаг 7. Наводим красоту перед выгрузкой
			StepEnd = 
				let	tbl = 	AddCol,
					to1 = 	Table.SelectRows(tbl, each
								// [ЗаголовокИмя] <> "Всего по промышленности" and 
								// [ЗаголовокИмя] <> "Всего по другим видам деятельности" and
								[Строка_ID]    <> 260 and [Строка_ID] <> 560 and // контрольные суммы
								[Значение]     <> null
							)
				in	to1

		in 	StepEnd,

	rep.1P1   = // !фоpма 1-п (натура) годовая - по ОПС
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			Connect =
					let 
						from   = F_Form[Connect](F_Form[GetName]("BKS"),"s_1p1") () ,

						   f1 = each [Date] 	 = #date(2021, 12, 1),
						   f2 = each [form_ID] = "s_1p",
						   f3 = each Table.SelectRows(_, each [Name]="1"),
						from1 = Table.SelectRows(from,  (x)=> x[form_ID] = "s_1p"),
						from2 = Table.SelectRows(from1, (x)=> x[Date] = #date(2021, 12, 1) ),
						from3 = Table.TransformColumns(from2, {"Content", f3})
					in  from,

			Load = [
				from = Connect,
				Date =
					let f=(x)=> Date.From(Text.Middle(x, 16, 4)),
						trs = Table.TransformColumns(from, {"Name", f, Date.Type}),
						ren = Table.RenameColumns(trs,{"Name", "Year"}),
						to  = Table.TransformColumns(ren, {"Content", (x)=>Excel.Workbook(x)} )
					in  to,

				// оставляем листы с номерами в названии, это коды организаций. Или = Table.AddColumn(Date, "Content2", each f([Content]))
				Transform =
					let f=(x)=> Table.SelectRows(x, each try Value.Is(Number.From([Name]), type number) = true otherwise null),
						sel = Table.TransformColumns(Date,{"Content", f})
					in  sel,

				Expand = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),

				Replace = Table.Buffer(
					Table.ReplaceValue(Expand,
						each [Предприятие_ID],
						each Record.FieldOrDefault([251600080200001 = 5079, 253100000000000 = 5001, 253500000000000 = 5005,
													253900000000000 = 5009, 254100000000000 = 5011, 254900000000000 = 5020,
													227100000000000 = 5555], // 5555 - такого номера нет, временный
								Text.From([Предприятие_ID]),
								[Предприятие_ID]),
						Replacer.ReplaceValue,
						{"Предприятие_ID"})
				)
				] [Replace],

			Expand	=
				let tbl = Load,
					// 	разворачиваем и переворачиваем таблицу
						Expand2 = Table.ExpandTableColumn(tbl, "Data", Table.ColumnNames(tbl{0}[Data]),Table.ColumnNames(tbl{0}[Data])),
						Unpivot = Table.UnpivotOtherColumns(Expand2, {"Year", "ID_Предприятие", "Статья", "ЕдИзм"}, "Заголовок", "Значение"),
						//  меняем тип колонок с трансформацией
						TransCol = Table.TransformColumns(
							Unpivot,{
								{ "Year",       (x)=> Date.From(x), Date.Type},
								{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
								{ "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
								{ "ЕдИзм",      (x)=> x, type text},
								{ "Заголовок",  (x)=> Number.From(x), Int64.Type},
								{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.From(x) otherwise null, Int64.Type}
							} )
				in	TransCol,

			Clean	=
				let tbl = Expand,
					to = tbl
				in	to,

			AddCol  =
				let tbl = Connect,
						// добавляем колонки с цифровым ID атрибутов
							// { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & Text.From([Статья]) , type text},         // номер строки + название статьи
							// { "ID_Группа",    each if Text.End(Text.From([ID_Строка]), 1) = "0" then 1 else 2, Int64.Type}  // если последняя цифра 0, то группа (1), иначе подгруппа (2)
							AddColumns = F[fxAddMultiColumns](tbl, {
								// статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
								{ "ID_Раздел",  each 1, Int64.Type},

								// уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
								{ "Заголовок_ID", each Number.From(Text.From([ID_Раздел]) & Text.From([Заголовок])), Int64.Type}
								}),

						// 	добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
							Join =
								let Join_Заголо  = Table.Join(AddColumns, "Заголовок_ID",    F_Form[spr.1P] [Спр_Заголовки],  "ID_Заголовок"),
									Join_Предпр  = Table.Join(Join_Заголо, "Предприятие_ID", F_Form[spr.BKS][Спр_Предприятия], "Код")
								in  Join_Предпр
				in	Join,

			End	    =
				let tbl = AddCol,
					selRows =
						Table.SelectRows(tbl, each
							(
								[Статья] <> "ИТОГО НЕПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and
								[Статья] <> "ИТОГО ПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ"
							)
						)
				in	selRows
		in 	End,

	rep.1P2   = // форма 1-п (натура) годовая - по райпо 
		let
		//  ФУНКЦИИ
			func = [
				//  функция очистки таблицы
					k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x
			],

		//  ПОДКЛЮЧЕНИЕ
			Connect = [
				path = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\1. Отчетность\БКС\форма 1-п (натура, годовая) райпо",
				from = Table.SelectColumns(Folder.Files(path), {"Name","Content"}) // сделать фильтр по xls или функцию использовать
			]  [from],

		//  ЗАГРУЗКА
			Load = [
				from = Connect,
				Date =
					let f=(x)=> Date.From(Text.Middle(x, 16, 4)),
						trs = Table.TransformColumns(from, {"Name", f, Date.Type}),
						ren = Table.RenameColumns(trs,{"Name", "Year"}),
						to  = Table.TransformColumns(ren, {"Content", (x)=>Excel.Workbook(x)} )
					in  to,

				// оставляем листы с номерами в названии, это коды организаций. Или = Table.AddColumn(Date, "Content2", each f([Content]))
				Transform =
					let f=(x)=> Table.SelectRows(x, each try Value.Is(Number.From([Name]), type number) = true otherwise null),
						sel = Table.TransformColumns(Date,{"Content", f})
					in  sel,

				Expand = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),

				Replace = Table.Buffer(
					Table.ReplaceValue(Expand,
						each [Предприятие_ID],
						each Record.FieldOrDefault([251600080200001 = 5079, 253100000000000 = 5001, 253500000000000 = 5005,
													253900000000000 = 5009, 254100000000000 = 5011, 254900000000000 = 5020,
													227100000000000 = 5555], // 5555 - такого номера нет, временный
								Text.From([Предприятие_ID]),
								[Предприятие_ID]),
						Replacer.ReplaceValue,
						{"Предприятие_ID"})
				)
			] [Replace],


		//  ЧИСТКА И ТРАНСФОРМАЦИЯ
			Clear =
				let from = Load,
					f=(x)=>
						let
							// фильтруем колонки (оставляем countCol (7) колонок = максимальное количество колонок с данными
							countCol = 7,
							SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=countCol, (x)=>x+1, (x)=>"Column" & Text.From(x) )),

							// удаляем (меняем на "") пробелы, переносы, табуляции др. (лучше до extraxt и unpivot)
							// если в ячейке null, то ошибка поэтому try-otherwise
							transform = Table.TransformColumns(SelCol,{"Column1", func[k]}),

							// делаем колонку 2 текстовую
							TransCol   = Table.TransformColumnTypes(transform,{{"Column2", type text}}),

							// оставляем строки с
							lst1 = List.Buffer({"тонн","тыс", "млн", "кал", "дал", "туб", "м.п.","Гкал"}),
							f=(tbl)=> try List.ContainsAny({tbl[Column2]}, lst1, Text.Contains) otherwise null,
							SelRow1 = Table.SelectRows(TransCol, f),

							// оставляем строки с Большой первой буквой
							lst2 = List.Buffer({"А".."Я","A".."Z"}),
							t=(tbl)=> List.Contains(lst2, Text.Start(tbl[Column1],1) ),
							SelRow2 = Table.SelectRows(SelRow1, t),

							// меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
							lstOld = List.Buffer(Table.ColumnNames(SelRow2)),                    // старый список имен столбцов из подтаблиц (все одинаковые)
							lstNew = List.Buffer({"Статья", "ЕдИзм", "1", "2", "3", "4", "5"}),  // новый список имен столбцов = 7
							rename = Table.RenameColumns(SelRow2, List.Zip( {lstOld,lstNew} ))
						in  rename,

					tr = Table.TransformColumns(from,{"Data", f})
				in  tr,


		//  РАЗВОРОТ
			Expand =
				let	from = Clear,

					// 	разворачиваем и переворачиваем таблицу
						lst     = Table.ColumnNames(from{0}[Data]),
						Expand  = Table.ExpandTableColumn(from, "Data", lst, lst),
						Unpivot = Table.UnpivotOtherColumns(Expand,
									{"Year", "Предприятие_ID", "Статья", "ЕдИзм"},
									"Заголовок", "Значение"
								),

					// 	меняем тип колонок с трансформацией
						TransCol =
							Table.TransformColumns(Unpivot,{
								{ "Year",       (x)=> Date.From(x), Date.Type},
								{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
								{ "Статья",     func[k], type text},
								{ "ЕдИзм",      (x)=> x, type text},
								{ "Заголовок",  (x)=> Number.From(x), Int64.Type},
								{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
							}),
							
					to = TransCol
				in 	to,

		//  ОБОГАЩЕНИЕ
			Add =
				let	from = Expand,

				// 	добавляем колонки с цифровым ID атрибутов
					add = F[fxAddMultiColumns](from, {
						// 	статьи
							{ "Раздел_ID",  each 1, Int64.Type},

						// уникальный номер заголовка в пределах отчета
							{ "Заголовок_ID", each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type}

						//
							// { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & Text.From([Статья]) , type text},         // номер строки + название статьи
						//
							// { "ID_Группа",    each if Text.End(Text.From([ID_Строка]), 1) = "0" then 1 else 2, Int64.Type}  // если последняя цифра 0, то группа (1), иначе подгруппа (2)
						}),

					// добавляем значения колонок с ID
					join = let
						j1 = Table.Join(add, "Заголовок_ID",   F_Form[spr.1P][Спр_Заголовки],   "Заголовок_ID"),
						j2 = Table.Join(j1,  "Предприятие_ID", F_Form[spr.BKS][Спр_Предприятия2], "Код")
					in  j2

				in  join,

		//  ВЫВОД
			End =
				let	from = Add,
					tbl1 = Table.SelectRows(from, each ([Статья] <> "ИТОГО НЕПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and [Статья] <> "ИТОГО ПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and [Статья] <> "Справочно: всего переработано плодоовощного сырья всеми методами")),
					tbl2 = Table.SelectRows(tbl1, each ([Значение] <> 0) and ([Значение] <> null))	// F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"})
				in	tbl2
		
		in	End,

	rep.1P3   = // форма 1-п (натура) годовая - по ОПС
		let
		//  ФУНКЦИИ
			func = [
				//  функция очистки таблицы
					k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x
			],

		//  ПОДКЛЮЧЕНИЕ
			Connect = [
				path = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\1. Отчетность\БКС\форма 1-п (натура, годовая) ОПО",
				from = Table.SelectColumns(Folder.Files(path), {"Name","Content"}) // сделать фильтр по xls или функцию использовать
			]  [from],

		//  ЗАГРУЗКА
			Load = [
				from = Connect,
				Date =
					let f=(x)=> Date.From(Text.Middle(x, 16, 4)),
						trs = Table.TransformColumns(from, {"Name", f, Date.Type}),
						ren = Table.RenameColumns(trs,{"Name", "Year"}),
						to  = Table.TransformColumns(ren, {"Content", (x)=>Excel.Workbook(x)} )
					in  to,

				// оставляем листы с номерами в названии, это коды организаций. Или = Table.AddColumn(Date, "Content2", each f([Content]))
				Transform =
					let f=(x)=> Table.SelectRows(x, each try Value.Is(Number.From([Name]), type number) = true otherwise null),
						sel = Table.TransformColumns(Date,{"Content", f})
					in  sel,

				Expand = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),

				Replace = Table.Buffer(
					Table.ReplaceValue(Expand,
						each [Предприятие_ID],
						each Record.FieldOrDefault([251600080200001 = 5079, 253100000000000 = 5001, 253500000000000 = 5005,
													253900000000000 = 5009, 254100000000000 = 5011, 254900000000000 = 5020,
													227100000000000 = 5555], // 5555 - такого номера нет, временный
								Text.From([Предприятие_ID]),
								[Предприятие_ID]),
						Replacer.ReplaceValue,
						{"Предприятие_ID"})
				)
			] [Replace],


		//  ЧИСТКА И ТРАНСФОРМАЦИЯ
			Clear =
				let from = Load,
					f=(x)=>
						let
							// фильтруем колонки (оставляем countCol (7) колонок = максимальное количество колонок с данными
							countCol = 7,
							SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=countCol, (x)=>x+1, (x)=>"Column" & Text.From(x) )),

							// удаляем (меняем на "") пробелы, переносы, табуляции др. (лучше до extraxt и unpivot)
							// если в ячейке null, то ошибка поэтому try-otherwise
							transform = Table.TransformColumns(SelCol,{"Column1", func[k]}),

							// делаем колонку 2 текстовую
							TransCol   = Table.TransformColumnTypes(transform,{{"Column2", type text}}),

							// оставляем строки с
							lst1 = List.Buffer({"тонн","тыс", "млн", "кал", "дал", "туб", "м.п.","Гкал"}),
							f=(tbl)=> try List.ContainsAny({tbl[Column2]}, lst1, Text.Contains) otherwise null,
							SelRow1 = Table.SelectRows(TransCol, f),

							// оставляем строки с Большой первой буквой
							lst2 = List.Buffer({"А".."Я","A".."Z"}),
							t=(tbl)=> List.Contains(lst2, Text.Start(tbl[Column1],1) ),
							SelRow2 = Table.SelectRows(SelRow1, t),

							// меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
							lstOld = List.Buffer(Table.ColumnNames(SelRow2)),                    // старый список имен столбцов из подтаблиц (все одинаковые)
							lstNew = List.Buffer({"Статья", "ЕдИзм", "1", "2", "3", "4", "5"}),  // новый список имен столбцов = 7
							rename = Table.RenameColumns(SelRow2, List.Zip( {lstOld,lstNew} ))
						in  rename,

					tr = Table.TransformColumns(from,{"Data", f})
				in  tr,


		//  РАЗВОРОТ
			Expand =
				let	from = Clear,

					// 	разворачиваем и переворачиваем таблицу
						lst     = Table.ColumnNames(from{0}[Data]),
						Expand  = Table.ExpandTableColumn(from, "Data", lst, lst),
						Unpivot = Table.UnpivotOtherColumns(Expand,
									{"Year", "Предприятие_ID", "Статья", "ЕдИзм"},
									"Заголовок", "Значение"
								),

					// 	меняем тип колонок с трансформацией
						TransCol =
							Table.TransformColumns(Unpivot,{
								{ "Year",       (x)=> Date.From(x), Date.Type},
								{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
								{ "Статья",     func[k], type text},
								{ "ЕдИзм",      (x)=> x, type text},
								{ "Заголовок",  (x)=> Number.From(x), Int64.Type},
								{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
							}),
							
					to = TransCol
				in 	to,

		//  ОБОГАЩЕНИЕ
			Add =
				let	from = Expand,

				// 	добавляем колонки с цифровым ID атрибутов
					add = F[fxAddMultiColumns](from, {
						// 	статьи
							{ "Раздел_ID",  each 1, Int64.Type},

						// уникальный номер заголовка в пределах отчета
							{ "Заголовок_ID", each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type}

						//
							// { "СтатьяСтрока", each Text.From([ID_Строка]) & " " & Text.From([Статья]) , type text},         // номер строки + название статьи
						//
							// { "ID_Группа",    each if Text.End(Text.From([ID_Строка]), 1) = "0" then 1 else 2, Int64.Type}  // если последняя цифра 0, то группа (1), иначе подгруппа (2)
						}),

					// добавляем значения колонок с ID
					join = let
						j1 = Table.Join(add, "Заголовок_ID",   F_Form[spr.1P][Спр_Заголовки],   "Заголовок_ID"),
						j2 = Table.Join(j1,  "Предприятие_ID", F_Form[spr.BKS][Спр_Предприятия2], "Код")
					in  j2

				in  join,

		//  ВЫВОД
			End =
				let	from = Add,
					tbl1 = Table.SelectRows(from, each ([Статья] <> "ИТОГО НЕПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and [Статья] <> "ИТОГО ПРОДОВОЛЬСТВЕННЫЕ ТОВАРЫ" and [Статья] <> "Справочно: всего переработано плодоовощного сырья всеми методами")),
					tbl2 = Table.SelectRows(tbl1, each ([Значение] <> 0) and ([Значение] <> null))	// F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"})
				in	tbl2
		
		in	End,
	
	rep.1TORG = // форма 1-торг (розница) годовая - по ОПО
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			Connect = 
				F_Form[Connect](F_Form[GetName]("BKS"),"s_1torg") () ,

		// 	ЗАГРУЗКА
			Load    =
				let
					ExpandColumns = Table.ExpandTableColumn(Connect, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
					FilteredRows  =
						F_Form[form.SelectRowsID](
							ExpandColumns,
							{"Column1", "Column2", "Column3", "Column4", "Column5", "Column6"},
							{"Статья", "Строка_ID", "1", "2", "3", "4"},
							"Строка_ID", 399, 700
						)
				in
					FilteredRows,

		// 	РАЗВОРОТ
			Expand  =
				let
				// 	разворачиваем и переворачиваем таблицу
					from = Load,

					lst  = List.Buffer(Table.ColumnNames(from[Step2]{0})),		// имена колонок в подтаблицах
					exp  = Table.ExpandTableColumn(from, "Step2", lst, lst),

					nms  = {"Company", "DataSet", "form_ID", "Date"},			// техколонки
					lst2 = {"Предприятие_ID", "Статья", "Строка_ID"},			// оставляемые колонки
					unp  = Table.UnpivotOtherColumns(exp, nms & lst2, "Заголовок", "Значение"),

				// 	меняем тип колонок с трансформацией
					transform = Table.TransformColumns(unp,
						{
							{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
							{ "Статья",     (x)=> Text.Combine(List.RemoveMatchingItems(Text.SplitAny(Text.Clean(x)," #(lf)#(00A0)"), {""}), " "), type text},
							{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
							{ "Заголовок",  (x)=> Number.From(x), Int64.Type},
							{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
						} )
				in	transform,
		// 	ОЧИСТКА
			Clear =
				let
					// 	1 удаляем текст внутри скобок + скобки и оставляем текст в скобках по маске
						removeText1 = F[Text.DeletBracket](Expand, "Статья"),

					//  2 меняем номера строк, т.к. после 2020 года строки 400..499 стали 500..599 с изменениями
						removeText2 = 
							let 
							// 	список  замены
								lst  = {
											{400,500},{401,501},{402,502},{404,503},{406,504},
											{407,505},{408,506},{410,509},{412,512},{414,513},{415,514}
										},
							
							// 	функция замены
								fxReplaceValues=
									(tbl as table, columnName as text, replaceList as list) as table =>
									let	ReplaceValue = (value as any) as any =>
											let	replacement = List.First(List.Select(replaceList, each _{0} = value), null)
											in	if replacement <> null then replacement{1} else value,
											
										to = 
											Table.TransformColumns(tbl, 
												{{columnName, each ReplaceValue(_), type any}}
											)
									in	to
						   	in fxReplaceValues(removeText1, "Строка_ID", lst) ,

					// 	3 удаляем текст по списку
						removeText3 = F[Text.RemoveMulti](removeText2, "Статья",
							{
								" - всего",
								"в том числе: ",
								"из них ",
								"из строки 410",
								"Справочно: из строки 401 ",
								
								"из строки 508",
								"из строки 510",
								"Справочно: из строки 501 "	//после совмещения строк за разные годы
							}
						),

					//  4 меняем названия статей
						removeText4 =
							Table.ReplaceValue(removeText3, (x)=>x[Статья],
								each Record.Field([
										500 = "Торговые объекты",
										600 = "Розничный товароборот фирменной торговли",
										601 = "Количество фирменных магазинов (на конец)",
										604 = "Торговая площадь фирменных магазинов (на конец)",
										410 = "передвижные средства РР торговли",
										412 = "интернет магазины"
									], Text.From([Строка_ID])),
								Replacer.ReplaceValue, {"Статья"}
							)
				in
					removeText4,

		//  ОБОГАЩЕНИЕ
			Add =
				let
					from = Clear,

				// 	Функция для получения значения Строка_ID
					GetID = each [Строка_ID],

				// 	сворачиваем таблицу отбратно, долбавляем две колонки и разворачиваем обратно
					Piv = 
						let 
							Pivot = 
								Table.Pivot(
									Table.TransformColumnTypes(from, 
										{{"Заголовок", type text}}, "ru-RU"), 
										List.Distinct(Table.TransformColumnTypes(from, {{"Заголовок", type text}}, "ru-RU")[Заголовок]), 
									"Заголовок", "Значение"
								),
							AddColumn1 = Table.AddColumn(Pivot,      "5", each [1]-[3], Int64.Type),
							AddColumn2 = Table.AddColumn(AddColumn1, "6", each [2]-[4], Int64.Type),

							nms  = {"Company", "DataSet", "form_ID", "Date"},
							lst = {"Предприятие_ID", "Статья", "Строка_ID"},
							unp  = Table.UnpivotOtherColumns(AddColumn2, nms & lst, "Заголовок", "Значение")
						in	unp,

				// 	добавляем колонки с цифровым ID атрибутов
					add =
						F[fxAddMultiColumns](Piv, {
							// 	номер раздела
								{ "Раздел_ID",  each
										let id = GetID(_)
										in	     if List.Contains ({400..599}, id) = true then 3
											else if List.Contains ({600..700}, id) = true then 4
											else 999
									, Int64.Type},

							// 	номер строки + название статьи
								{ "СтатьяСтрока",  each
										Text.From(GetID(_)) & " " & Text.From([Статья])
									, type text},

							// 	уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
								{ "Заголовок_ID", each
										Number.From( Text.From([Раздел_ID]) & Text.From([Заголовок]) )
									, Int64.Type},

							// 	если последняя цифра 0, то группа (1), иначе подгруппа (2) В ЭТОМ ОТЧЕТЕ ИЗЛИШНЕ, ТУТ НЕ ТАК НУМЕРАЦИЯ ИДЕТ
								{ "Группа_ID", each
										if Text.End(Text.From(GetID(_)), 1) = "00"
										then "Группа"	// только строка 500
										else "Подгруппа"
									, type text},

							// 	вид деятельности
								{ "СелоГород_ID", each
										if 		List.Contains ({31,32}, [Заголовок_ID]) then "1 Всего"
										else if List.Contains ({33,34}, [Заголовок_ID]) then "3 Село"
										else if List.Contains ({35,36}, [Заголовок_ID]) then "2 Город"
										else "пусто"
									, type text},

							// 	уровень иерархии = Группа_ID = > оставить надо одно
								{ "Уровень_ID", each
										if Number.Mod(GetID(_), 10) <> 0  then 2
										else 1
									, Int64.Type}
						}),

				// 	добавляем значения колонок с ID
					join = let
						j1 = Table.Join(add,	"Заголовок_ID",   		F_Form[spr.1TORG][Спр_Заголовки], 		"Заголовок_ID"),
						j2 = Table.Join(j1,  	"Раздел_ID",			F_Form[spr.1TORG][Спр_Разделы], 		"Раздел_ID"),
						j3 = Table.Join(j2,  	"Предприятие_ID", 		F_Form[spr.BKS][Спр_Предприятия], 		"Предприятие_ID"),
						j4 = Table.Join(j3,  	"Предприятие_Группа1", 	F_Form[spr.BKS][Спр_ГруппаПредприятий], "Group_ID")
					in  j4

				in  join,

		//  ВЫВОД
			End =
				let from = Add,
					tbl1 = Table.SelectRows(from, each ([Значение] <> 0) and ([Значение] <> null)),
					tbl2 = Table.SelectRows(tbl1, each [Строка_ID] >= 500 ),
					tbl3 = Table.SelectRows(tbl2, each [СелоГород_ID] <> "1 Всего" )
				in	tbl3
		in	End,

	rep.4TORG = // фоpма 4-торг (продажа) квартальная - по отдельному райпо
		// (optional datasetName as nullable text) => 
		// let 
		// // 	ПОДКЛЮЧЕНИЕ
		// 	Connect =	
		// 		let 
		// 			from  =	F_Form[Connect](F_Form[GetName](datasetName),"v_7") () ,

		let
			F_Torg        = // локальные функции
				[	//  функция очистки таблицы
						k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
					// 	удаляет строки сверху, до строки, в которой находится искомый текст
						Table.SkipUp =
							(txt)=>(row)=> List.Select( Record.ToList(row), (i)=>i is text and Text.Contains(i,txt))={},
					//
						f =(x)=>(y)=>not Text.Contains( Text.Combine( List.Select( Record.ToList(y), (i)=>i is text)), x)
				],

			// Connect = 
			// 	let ConnectFolder = F_Form[Connect][BKS],
			// 		LoadFiles     = F_Form[form.ConnectFiles](ConnectFolder, "s_1torg")
			// 	in 	LoadFiles,

		// 	ПОДКЛЮЧЕНИЕ
			Connect = let 
				path = "G:\=EXCEL BI - мои проекты\2024.01 - Анализ Отчетности\1. Отчетность\Дрогичин\Статотчетность\Фоpма 4-торг (продажа)",
				from = Table.SelectColumns(Folder.Files(path), {"Name","Content"})
			in 	from,

		// 	ЗАГРУЗКА
			Load    =
				[	from = Connect,
					
					// вытягиваем дату из названия файла (4 знака с конца)
					Date =
						let ren = Table.RenameColumns(from, {"Name", "Date"}),
							tr  = Table.TransformColumns(ren, {"Date",    (x)=> Date.From(Text.Start(x, 10)), Date.Type}),
							to  = Table.TransformColumns(tr,  {"Content", (x)=> Excel.Workbook(x)} )
						in  to,

					// оставляем листы с номерами в названии, это коды организаций. Или = Table.AddColumn(Date, "Content2", each f([Content]))
					Transform =
						let f=(x)=> Table.SelectRows(x, each try Value.Is(Number.From([Name]), type number) = true otherwise null),
							sel = Table.TransformColumns(Date,{"Content", f})
						in  sel,

					Expand = Table.ExpandTableColumn(Transform, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),

					// заменяем некоторые номера предприятий на современные
					Replace =
						Table.Buffer(
							Table.ReplaceValue(Expand,
								each [Предприятие_ID],
								each Record.FieldOrDefault(
										[	251600080200001 = 5079,
											253100000000000 = 5001,
											253500000000000 = 5005,
											253900000000000 = 5009,
											254100000000000 = 5011,
											254900000000000 = 5020,
											227100000000000 = 5555], // 5555 - такого номера нет, временный
										Text.From([Предприятие_ID]
									),
								[Предприятие_ID]),
								Replacer.ReplaceValue, {"Предприятие_ID"}
							)
						)
				] [Replace],

		// 	ЧИСТКА подтаблиц
			Clear   =
				let from = Load,
					f=(x)=>
						let
							// список имен колонок
							countCol = 8,
							nms = List.Generate(()=> 1, (x)=> x<=countCol, (x)=>x+1, (x)=>"Column" & Text.From(x) ),

							SelCol = Table.SelectColumns(from, nms),

							// удаляем (меняем на "") пробелы, переносы, табуляции др. (лучше до extraxt и unpivot)
							// если в ячейке null, то ошибка поэтому try-otherwise
							Clean = Table.TransformColumns(x,{"Column1", F_Torg[k]}),

							// делаем колонку 2 текстовую
							TransCol  = Table.TransformColumnTypes(Clean,{{"Column2", type text}}),

							SkipUp   = Table.Skip(TransCol, F_Torg[Table.SkipUp]("Код по СК")),
							SkipDown = Table.Skip(SkipUp,3),
							Remove   = Table.RemoveLastN(Table.RemoveLastN(SkipDown, F_Torg[f]("РАЗДЕЛ II ") ), 1),

							lstOld = List.Buffer(Table.ColumnNames(Remove)), // старый список имен столбцов из подтаблиц (все одинаковые)
							lstNew = List.Buffer({
												"Код",
												"Статья",
												"КодСтроки",
												"ЕдИзм",
												"Продано (всего)",
												"Продано (отеч)",
												"Запасы (всего)",
												"Запасы (отеч)"
											}),
							rename = Table.RenameColumns(Remove, List.Zip( {lstOld,lstNew} ))
						in  rename,

					tr = Table.Sort(
							Table.TransformColumns(from, {"Data", f}),
							{"Date", Order.Ascending}
						)
				in  tr,

		// 	РАЗВОРОТ
			Expand  =
				let lst = Table.ColumnNames(Clear[Data]{0}),
					exp = Table.ExpandTableColumn(Clear, "Data", lst, lst ),
					del = Table.RemoveColumns(exp, {"Предприятие_ID"})
				in  del,

		// 	ОЧИСТКА
			Clean   =
				[	tbl = Expand,
					
					removeText =
						let lst    = {"в том числе: ","  в том числе:","  из них: ","  из нее: ","из него:","из них ","из нее "},
							remove = F[Text.RemoveMulti](tbl, "Статья", lst),
							a = Table.TransformColumns(remove,{{"Статья", Text.Trim, type text}}),
							b = Table.TransformColumns(a, {{"Статья", (x)=>F[Text.ReplaceBetweenDelimiters](x,"(",")"), type text}}),
							c = Table.TransformColumns(b, {{"Статья", (x)=>F_Torg[k](x), type text}})
						in  c,

					replaceText =
						Table.ReplaceValue(removeText,
							(x)=>x[Код],
							each Record.Field([
									Мебель 							 = "47.59",
									#"Ювелирные изделия"			 = "47.77.10.200",
									#"Безалкогольные напитки прочие" = "47.25.12.900",
									#"Печатные издания" 			 = "47.62"
								],
								Text.From([Статья])), Replacer.ReplaceValue, {"Код"}
						),

					typed  =
						Table.TransformColumnTypes(
							replaceText,
							List.Transform(
								List.Skip(Table.ColumnNames(replaceText), 5),
								(x)=>{x, type number}
							)
						)
				][typed],

			addMiltiCol = [ // должно быть именно тут, т.к. на след шаге эти колонки сворачиваются в колонку с показателями!
					func = (table as table, lst as list) as table => List.Accumulate( lst, table, (s, c) => Table.AddColumn(s, c{0}, c{1}, c{2}?)),
					add = func(Clean,{
						{"Продано (заруб)", each [#"Продано (всего)"]-[#"Продано (отеч)"], type number},
						{"Запасы (заруб)",  each [#"Запасы (всего)"]-[#"Запасы (отеч)"], type number}
					})
				][add],

		// 	ДРУГОЕ
			unpivot1 = [
					from    = addMiltiCol,
					lst	    = {"Date", "Код", "Статья", "ЕдИзм", "КодСтроки"}, // колонки на своих местах
					Unpivot = Table.UnpivotOtherColumns(from, lst, "Показатель", "Значение"),
					Replace = Table.ReplaceValue(Unpivot,0,null,Replacer.ReplaceValue,{"Значение"}),
					Select  = Table.SelectRows(Replace, each ([Значение] <> null))
					][Select],

			level = [
				name = "Level",
				a = Table.DuplicateColumn(unpivot1, "Код", name),
				c = Table.ReplaceValue(a,".","",Replacer.ReplaceText,{name}),
				d = Table.TransformColumns(c, {name, Text.Length, Int64.Type}),
				e = Table.TransformColumns(d, {name, each _ - 1, type number})
				]  [e],

			pivot = //
				Table.Pivot(
					Table.TransformColumnTypes(level, {{"Date", type text}}, "ru-RU"),
					List.Distinct(Table.TransformColumnTypes(level, {{"Date", type text}}, "ru-RU")[Date]),
					"Date", "Значение", List.Sum
				),
		
		// 	Нарастающие итоги
			sumNarast = [ // избавляемся от нарастающих итогов, делаем их поквартальными
				f = (lst) => List.Generate(
					()  => [i = 0, r = lst{0} ?? 0, rt = r],
					(x) => x[i] <= 3,
					(x) => [i = x[i] + 1, r = if lst{i} = null then 0 else lst{i} - x[rt], rt = x[rt] + r],
					(x) => x[r]
				),

				to_list =
					Table.ToList(pivot,
						(x) =>
							{x{0}} & {x{1}}& {x{2}} & {x{3}} & {x{4}} & {x{5}}	// поменять на нормальный код
							& List.Combine(List.Transform(List.Split(List.Skip(x, 6), 4), f) ) // если столбец добавился, то изменить число
					),

				to_table = Table.FromList(to_list, (x)=>x, Table.ColumnNames(pivot))
				][to_table],

			unpivot2 = [
				lst = {"Код", "Статья", "ЕдИзм", "Показатель", "Level", "КодСтроки"},
				unp = Table.UnpivotOtherColumns(sumNarast, lst, "Дата", "Значение"),
				tp  = Table.TransformColumnTypes(unp,{{"Дата", type date}, {"Значение", type number}})
				][tp],

			Классификатор = 
				#table(type table [Код=text], {} ), // временный. Чтобы МЫС не ругался что нет идентификатора
			
			join = [
				a      = Table.NestedJoin(unpivot2, {"Код"}, Классификатор, {"Код"}, "Классификатор", JoinKind.LeftOuter),
				b 	   = Table.ExpandTableColumn(a, "Классификатор", {"Наименование"}, {"СтатьяКлас"}),
				nmsTxt = List.FirstN(Table.ColumnNames(b), 5),
				nmsNum = {Table.ColumnNames(b){5}},
				t      = Table.TransformColumnTypes(b,
							List.Transform(nmsTxt,(x)=>{x,type text}) &
							List.Transform(nmsNum,(x)=>{x,type number})
						)
				][t],

			replaceText =
				Table.ReplaceValue(join, (x)=>x[Статья],
					each Record.Field([
						17 = "консервы, пресервы, икра, икорные изделия",
						21 = "шоколад, кондитерские изделия с какао",
						12 = "консервы из мяса и субпродуктов",
						65 = "Парфюм и туалетные принадлежности",
						56 = "Стиральные и сушильны машины",
						57 = "Средства бытовой химии"
					], Text.From([КодСтроки])),
					Replacer.ReplaceValue, {"Статья"}
				),

		// 	НОВЫЕ КОЛОНКИ
			addColumns = F[fxAddMultiColumns](replaceText, {
				{"ГруппаТоваров",         each if Text.Start([Код],4) = "47.2" then "Продтовары" else "Промтовары", type text}
				,{"СтатьяЕдИзм",           each [СтатьяКлас] , type text}
				,{"СтатьяКодСтрокиКлас",   each [КодСтроки] & ", " & [СтатьяКлас], type text}
				,{"СтатьяКодСтрокиИсх",    each [КодСтроки] & " "  & [Статья] & ", " & [ЕдИзм], type text}
				,{"СтатьяКодКодСтрокиИсх", each [Код] & "-" & [КодСтроки] & " " & [Статья] & ", " & [ЕдИзм], type text}
				,{"Группа",                each if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true then "Группа" else "Подгруппа", type text}
				// ,{"Path_1",                each F_Form[TransformText]([Код], {2..8}), type text} // вставляем символы в позиции для разных уровней
				// ,{"Path_2",                each F_Form[TransformText]([Код], {3,5,7}), type text}
			}),
			
			addColumns2 = 
				F_Form[fxAddTechColumn](addColumns,null,1,"Drog","11z"),
		
		// 	ВЫВОД
			End     =
				let	tbl  = addColumns2,
					tbl2 = Table.SelectRows(tbl, each ([Значение] <> 0))	// F[Row.RemoveNullColumns](tbl, {"Текущий период","Прошлый период"})
				in	tbl2
		in	End,
	rep.OCF   = // ОТЧЕТ О ЦЕЛЕВОМ ФИНАНСИРОВАНИИ
		(optional datasetName as nullable text) => 
		/*
			Описание:
		*/
		let
			Connect =	
				F_Form[Connect](F_Form[GetName](datasetName),"ocf") () ,

			Load	=
				let	tbl = Connect,
					// 	выбираем колонки с данными в подтаблицах с особенностей датасета
						company = tbl[Company]{0},
						nmsOld = if company = "Mart" then F_Form[getColumnNames]({1,6,10,14})
							else if company = "..."  then F_Form[getColumnNames]({1..100})
							else null,
				
					//	фильтруем строки в подтаблицах с отчетами по диапазону номеров строк отчета
						to	= 	F_Form[form.SelectRowsID](
									tbl,
									nmsOld,
									{"Показатель", "Строка_ID", "1", "2"},
									"Строка_ID",
									10,
									400
								)
				in 	to,

			Expand  =
				let	tbl = Load,
					// 	список заголовков в подтаблицах отчета - 4 шт.
						lstHeader = {"Статья", "Строка_ID"} & {"На начало", "На конец"},

					// 	разворачиваем таблицу
						Expand =
							Table.ExpandTableColumn(tbl,
								"Step2",
								Table.ColumnNames(Record.Field(tbl{0}, "Step2")),
								lstHeader
							)
				in Expand,

			Clean   =	// на входе 4 технические колонки + 10 колонок отчета + ?1 колонка с индексом
				let	tbl = Expand,
					// 	делаем последние 8 колонок числовыми
						Types =
							Table.TransformColumnTypes(tbl,
								List.Transform(
									List.LastN(Table.ColumnNames(tbl), 2),
									each {_, type number}
								)
							),

					// 	сортируем колонку 5 - это "Строка_ID" (пока не переименовывали колонки)
						Sort = Table.Sort(Types, {Table.ColumnNames(tbl){5}, Order.Ascending} ),

					// 	удаляем лишний текст в подстроке столбца
						MultiRemove = F[Text.RemoveMulti](Sort, "Статья", 
							{"", ""} 
						),

					//  меняем на правильные названия некоторых статей 
						MultiReplace =
							F[Text.MultiReplace](MultiRemove, "Строка_ID", "Статья", 
								[	240 = "безвозмездная помощь",
									325 = "амортизация"
								]
							)

				in MultiReplace,

			Add  =
				let	tbl = Clean,
				
					// делаем отступ, если номер строки не четный
						offsetText =
							F_Form[form.TextOffset](tbl, "Статья", "Строка_ID", 10),

					// 	меняем 0 на null (чтобы в визуализации не появлялся) - можно убрать и оставить фильтр идущий следом + <> 0
						Replace0ToNull =
							Table.ReplaceValue(offsetText,
								0, null, Replacer.ReplaceValue,
								{"На начало","На конец"}
							),

					//  удаляем пустые строки
						Select  = 	Table.SelectRows(Replace0ToNull, each 
							true 
							// [На начало] <> null and [На конец] <> null
						),

					// добавляем новые колонки
						AddColumns =
							F[fxAddMultiColumns](Select, {
								{ 	"Структура_ID",
										each if List.Contains ({100,400},[Строка_ID])  = true then "Остатки"
										else if List.Contains ({200..299},[Строка_ID]) = true then "Поступило"
										else if List.Contains ({300..399},[Строка_ID]) = true then "Использовано"
										else 4 // "Другое"
									, type text
								},

								// 	уровень иерархии: 1 - группа, 2 - подгруппа
								{ 	"Уровень_ID", each
										if Number.Mod([Строка_ID], 10) = 0 = true
										then 1
										else 2
									, Int64.Type
								}
							})
				in AddColumns,

			End = Add		
		in 	End,
		
	rep.11z   = // ОТЧЕТ заготовки
		(optional datasetName as nullable text) => 
		let	// F_Form[GetName](datasetName)
			Connect = F_Form[Connect]("BKS","v_11z","заг") () ,
			Load	= 
				let	exp 	= Table.ExpandTableColumn(Connect, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
					nmsOld 	= F_Form[getColumnNames]({1..17}),
					nmsNew 	= {"Статья", "ЕдИзм","Строка_ID"} & List.Transform({1..14}, (x)=>Text.From(x)),
					to		= if 	List.Count(nmsOld) <> List.Count(nmsNew)
							  then 	"число элементов списков не равно"
							  else 	F_Form[form.SelectRowsID](exp, nmsOld, nmsNew, "Строка_ID", 1, 100) 
				in  to,
			
			Expand  = 
				let	exp = 	Table.ExpandTableColumn(Load, "Step2", Table.ColumnNames(Load[Step2]{0})),
					lst = 	List.FirstN(Table.ColumnNames(exp), 8),
					unp = 	Table.UnpivotOtherColumns(exp, lst, "Header_ID", "Значение"),
					trs = 	Table.TransformColumns(unp,
							{
								{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
								{ "Статья",     (x)=> F[Text.Clear](F[Text.PowerTrim](x," ")), type text},
								{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
								{ "ЕдИзм",  	(x)=> x, type text},
								{ "Header_ID",  (x)=> Number.From(x), Int64.Type},
								{ "Значение",   (x)=> 	try	if Number.From(x) = 0 
															then null 
															else Number.Round(Number.From(x),1) 
														otherwise 
															null
														, Int64.Type}
							})
				
				in  trs,
			
			Clear   = 
				let	del = 	F[Text.DeletBracket](Expand,"Статья"),
					rem = 	F[Text.RemoveMulti](del, "Статья",{
								"в т. ч. : ",
								"в том числе: ",
								"из него ",
								"из них ",
								" - всего"
							}),
					sel = 	Table.SelectRows(rem, each 
								not List.Contains({0,null}, [Значение]) and 
								not List.Contains({36,37,38}, [Строка_ID]))
				in sel,

			Add 	= 
				let	add = 	F[fxAddMultiColumns](Clear, {
						{ "СтатьяСтрока",  each Text.From([Строка_ID]) & " " & Text.From([Статья]) , type text},
						{ "Группа", each
							let id = [Строка_ID]
							in 	 	 if List.Contains ({1..15}, id)  then "1 Продукция растениеводства"
								else if List.Contains ({16..18}, id) then "2 Дикорастущая продукция"
								else if List.Contains ({19..21}, id) then "3 Продукция животноводства"
								else if List.Contains ({22..26}, id) then "4 Кожевенное сырье"
								else if List.Contains ({27..35}, id) then "5 Вторичное сырье"
								else "0",
							type text},

						{ "Уровень", each
									if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true
									then "Статья"
									else "Подстатья"
								, type text}
							
						}),

					Join = F[fxJoinSpr](add, "Header_ID", F_Form[spr.11z][Спр_Заголовки], "Header_ID", {"Header"})
				in	Join,

			End 	= 
				let tbl  = F[Table.TypedField](Add),
					cols = {"Предприятие_ID","Header_ID"},
					to   = Table.TransformColumnTypes(tbl, List.Transform(cols, each {_, Int64.Type}))
				in  to
		in 	End,
	
	rep.12po  = // Фоpма 12-п (месячная, срочная) ОПО // ОТЧЕТ О ПРОИЗВОДСТВЕ ПРОДУКЦИИ И ВЫПОЛHЕHHЫХ РАБОТАХ, УСЛУГАХ ПРОМЫШЛЕHHОГО ХАРАКТЕРА
		/*
			подключаться или так или так
				Connect = F_Form[Connect](F_Form[GetName](datasetName),	null,	 "Фоpма 12-п (месячная, срочная) ОПО" )(),
				Connect = F_Form[Connect](F_Form[GetName](datasetName), "v_12po", null                                )(),
		*/
		
		(optional datasetName as nullable text) => 
		let	Connect = F_Form[Connect](F_Form[GetName](datasetName),null,"Фоpма 12-п (месячная, срочная) ОПО")(),
			Load	= 
				let	exp	 = 	Table.ExpandTableColumn(Connect, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
					//row  =  Table.SelectRows(exp, each ([Предприятие_ID] = "1") and ([Date] = #date(2017, 12, 31))),
					
					f=(x)=> 
						let	flt  = 	Table.SelectRows(x, each ([Column1] <> null)),
							col  = 	Table.SelectColumns(flt, List.FirstN(Table.ColumnNames(flt),15)),
							err  = 	F[Table.ReplaceErr](col, Table.ColumnNames(col)),
							clr  = 	F[Table.ClearAll](err," #(lf)#(00A0)", false, false),
							grp  = 	Table.Group(clr,
										{"Column1"},
										List.Transform({"tmp"},(x)=>{x, (t)=>t, type text}),
										GroupKind.Local,
										(s,c)=>Number.From(Text.StartsWith(c[Column1], "Раздел"))
									),
							ren  = 	Table.RenameColumns(grp, {{"Column1", "Раздел_ID"}}),
							sel  = 	Table.SelectRows(ren, each 
										[Раздел_ID] <> "* Раздел IV отчета заполняется ежеквартально" 
										and Text.Start([Раздел_ID],2) <> "01"
									),
							to   =	Table.TransformColumns(sel, {
										{"Раздел_ID", (x)=>
												 if Text.Contains(x,"IV")   then 4
											else if Text.Contains(x,"III")  then 3
											else if Text.Contains(x,"II") then 2
											else if Text.Contains(x,"I")  then 1
											else 0, type number},
										
										{"tmp", each 
											let 
												skp = Table.PromoteHeaders(Table.Skip(_, F[Table.SkipUp]("А"))),
												rem = F[Column.RemoveNull](skp),
												ren = Table.RenameColumns(rem, {{Table.ColumnNames(rem){0}, "А"}}), // русское А chr(192). Было другое chr(32)!
												col = Table.RemoveColumns(ren,
														List.Select(
															Table.ColumnNames(ren),
															each Text.StartsWith(_, "Column")
														)
													  ),
												typ = [
													tbl  =  col,
													lst  = 	List.Transform({1..12}, each Text.From(_)),
													nlst = 	List.Select(
																Table.ColumnNames(tbl), 
																(i)=>List.Contains(lst, i, (x,y)=>Text.Contains(y,x))
															),
													to   =	Table.TransformColumnTypes(tbl,
																List.Transform(nlst,(i)=>{i,Number.Type})
															// & List.Transform(txtlst,(i)=>{i,Text.Type}
														)
												
													][to],
												err = F[Table.ReplaceErr](typ, Table.ColumnNames(typ)),
												zer = Table.TransformColumns(err, {}, (x)=>if x=0 then null else x),
												ind = Table.AddIndexColumn(zer, "Index", 1, 1, Int64.Type)
											in 
												ind}
									})
						in  to,
					to1  = 	Table.TransformColumns(exp, {"Data", f}),
					g=(x)=>	Table.TransformColumns(x, {"tmp", each 
							if List.Count(Table.ColumnNames(_)) = 6 // раздел 4
							then let a = Table.AddColumn(_, "Б", each "тыс. руб."),
								 	 b = Table.ReorderColumns(a,
											List.InsertRange(
												List.Difference(Table.ColumnNames(a),{"Б"}),
												1, {"Б"}
											)
										)
								in b
							else _
						}),
					to2  = 	Table.TransformColumns(to1, {"Data", each g(_)})
				in  to2,

			Expand 	= 
				let	exp1 = 	Table.ExpandTableColumn(Load, "Data", {"Раздел_ID", "tmp"}),
					nms  =  List.Distinct(List.Combine(
								List.Transform(
									Table.Column(Table.FirstN(exp1, 4), "tmp"), 
									each Table.ColumnNames(_)))
							),
					exp2 = 	Table.ExpandTableColumn(exp1, "tmp", nms),
					unp  = 	Table.UnpivotOtherColumns(exp2, 
								{"Company", "DataSet", "form_ID", "Date", "Предприятие_ID", "Раздел_ID","А","Б","Index"}, 
								"Header_Index", "Значение"
							),
					ren  =  Table.RenameColumns(unp,{{"А","Статья"}}),
					to   =  Table.TransformColumns(ren, {
								{"Предприятие_ID", (x)=>Number.From(x), type number}
							})
				in  to,

			Clear 	= 
				let 
					// tbl  = 	Table.RemoveColumns(Expand,{"Company", "DataSet", "form_ID","Предприятие_ID","Date"}),
					rep  =	F[Text.MultiReplace](Expand, "Б", "Статья", 
								[	101 = {"Объем пpоизводства в фактических отпускных ценах (без налогов из выручки)",
										   "Объем пpоизводства"}{1},
									102 = {"Стоимость перереработанного сыpья (неоплаченного изготовителем) в фактических ценах",
										   "Стоимость перереработанного сыpья"}{1},
									103 = {"Объем отгруженной продукции в фактических отпускных ценах (без налогов из выручки), всего",
										   "Объем отгруженной продукции"}{1},
									105 = {"Стоимость запасов готовой пpодукции в фактических отпускных ценах (без налогов из выручки) на конец периода",
										   "Стоимость запасов готовой пpодукции"}{1},
									1039 =  " переработка и консервирование других фруктов и овощей",
									10710 = "  производство хлеба и х/б изделий; производство мучных конд. изделий недлительного хранения",
									10720 = "  производство сухарей и печенья, производство мучных конд. изделий длительного хранения",
									13960 = "  производство различных текстильных изделий"
								]
							),

					
					rem  = 	F[Text.RemoveMulti](rep, "Статья", {"в т.ч.", ""})
				in	rem,

			Add 	= 
				let
					ren  = 	Table.RenameColumns(Clear, {"Б","Строка"}),
					add1 = 	Table.AddColumn(ren, "Header_ID", each Text.From([Раздел_ID]) & [Header_Index], type text),
					join = 	
						let a = F[fxJoinSpr](add1, "Header_ID", F_Form[spr.12po][Спр_Заголовки], "Header_ID", 
									{"Header"}	),
							b = F[fxJoinSpr](a, "Раздел_ID", F_Form[spr.12po][Спр_Разделы], "Раздел_ID", 
									{"РазделПрефикс","РазделИмя"}),
							c = F[fxJoinSpr](b, "Предприятие_ID", F_Form[spr.BKS] [Спр_Предприятия2], "Код",
									{"Название_исходное","Группа","Область"} )
						in  c,

					sel  = 	Table.SelectRows(join, each 
								[Header] <> null and 
								[Строка] <> "x" and 
								not List.ContainsAny({[Статья]}, {"Всего","Итого"}, Text.Contains)
							),
					add2 = 	F[fxAddMultiColumns](sel, {
								{ "СтатьяСтрока", each 
									if [Раздел_ID] = 1 or [Раздел_ID] = 2 then 
										Text.From([Строка]) & " " & Text.From([Статья]) 
									else if [Раздел_ID] = 3 then 
										Text.From([Статья]) & ", " & Text.From([Строка]) 
									else 
										Text.From([Статья]) & ", тыс. руб."
									, type text},

								{ "Уровень", each
									if [Раздел_ID] <> 2 then 
										if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true
										then "Статья"
										else "Подстатья"
									else "Статья"
									, type text}
							
						}),
					to  = 	add2
				in  to,

			End 	= 
				let tbl  = 	F[Table.TypedField](Add),
					cols = 	{"Предприятие_ID","Header_Index","Раздел_ID","DataSet","Index"},
					to   = 	Table.TransformColumnTypes(tbl, List.Transform(cols, each {_, Int64.Type})),
					sel  = 	Table.SelectRows(to, each 
								not List.ContainsAny({[Header]}, {"Темп"}, Text.Contains)
							),
					lst  =  nmsTech2 & 
								{"Статья","Строка","СтатьяСтрока"} & 
								{"Значение"} &
								{"Раздел_ID","РазделПрефикс","РазделИмя"} & 
								{"Предприятие_ID","Название_исходное","Группа","Область"} & 
								{"Header_Index","Header_ID"} & 
								{"Уровень","Index"}
							,
					reo  =  Table.ReorderColumns(sel, lst, MissingField.UseNull)
				in  reo
		
		in 	End,
					
	rep.12pr  = // Фоpма 12-п (месячная, срочная) райпо // ОТЧЕТ О ПРОИЗВОДСТВЕ ПРОДУКЦИИ И ВЫПОЛHЕHHЫХ РАБОТАХ, УСЛУГАХ ПРОМЫШЛЕHHОГО ХАРАКТЕРА
		/*
			подключаться или так или так
				Connect = F_Form[Connect](F_Form[GetName](datasetName),	null,	 "Фоpма 12-п (месячная, срочная) ОПО" )(),
				Connect = F_Form[Connect](F_Form[GetName](datasetName), "v_12po", null                                )(),
		*/
		
		(optional datasetName as nullable text, optional loadMode as nullable text) => 
		let	
			Connect = F_Form[Connect](F_Form[GetName](datasetName),null,"Фоpма 12-п (месячная, срочная) райпо")(),
			Load	= 
				let	exp	 = 	Table.ExpandTableColumn(Connect, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
					//row  =  Table.SelectRows(exp, each ([Предприятие_ID] = "1") and ([Date] = #date(2017, 12, 31))),
					
					f=(x)=> 
						let	flt  = 	Table.SelectRows(x, each ([Column1] <> null)),
							col  = 	Table.SelectColumns(flt, List.FirstN(Table.ColumnNames(flt),15)),
							err  = 	F[Table.ReplaceErr](col, Table.ColumnNames(col)),
							clr  = 	F[Table.ClearAll](err," #(lf)#(00A0)", false, false),
							grp  = 	Table.Group(clr,
										{"Column1"},
										List.Transform({"tmp"},(x)=>{x, (t)=>t, type text}),
										GroupKind.Local,
										(s,c)=>Number.From(Text.StartsWith(c[Column1], "Раздел"))
									),
							ren  = 	Table.RenameColumns(grp, {{"Column1", "Раздел_ID"}}),
							sel  = 	Table.SelectRows(ren, each 
										[Раздел_ID] <> "* Раздел IV отчета заполняется ежеквартально" 
										and Text.Start([Раздел_ID],2) <> "01"
									),
							to   =	Table.TransformColumns(sel, {
										{"Раздел_ID", (x)=>
												 if Text.Contains(x,"IV")   then 4
											else if Text.Contains(x,"III")  then 3
											else if Text.Contains(x,"II") then 2
											else if Text.Contains(x,"I")  then 1
											else 0, type number},
										
										{"tmp", each 
											let 
												skp = Table.PromoteHeaders(Table.Skip(_, F[Table.SkipUp]("А"))),
												rem = F[Column.RemoveNull](skp),
												ren = Table.RenameColumns(rem, {{Table.ColumnNames(rem){0}, "А"}}), // русское А chr(192). Было другое chr(32)!
												col = Table.RemoveColumns(ren,
														List.Select(
															Table.ColumnNames(ren),
															each Text.StartsWith(_, "Column")
														)
													  ),
												typ = [
													tbl  =  col,
													lst  = 	List.Transform({1..12}, each Text.From(_)),
													nlst = 	List.Select(
																Table.ColumnNames(tbl), 
																(i)=>List.Contains(lst, i, (x,y)=>Text.Contains(y,x))
															),
													to   =	Table.TransformColumnTypes(tbl,
																List.Transform(nlst,(i)=>{i,Number.Type})
															// & List.Transform(txtlst,(i)=>{i,Text.Type}
														)
												
													][to],
												err = F[Table.ReplaceErr](typ, Table.ColumnNames(typ)),
												zer = Table.TransformColumns(err, {}, (x)=>if x=0 then null else x),
												ind = Table.AddIndexColumn(zer, "Index", 1, 1, Int64.Type)
											in 
												ind}
									})
						in  to,
					to1  = 	Table.TransformColumns(exp, {"Data", f}),
					g=(x)=>	Table.TransformColumns(x, {"tmp", each 
							if List.Count(Table.ColumnNames(_)) = 6 // раздел 4
							then let a = Table.AddColumn(_, "Б", each "тыс. руб."),
								 	 b = Table.ReorderColumns(a,
											List.InsertRange(
												List.Difference(Table.ColumnNames(a),{"Б"}),
												1, {"Б"}
											)
										)
								in b
							else _
						}),
					to2  = 	Table.TransformColumns(to1, {"Data", each g(_)})
				in  to2,

			Expand 	= 
				let	exp1 = 	Table.ExpandTableColumn(Load, "Data", {"Раздел_ID", "tmp"}),
					nms  =  List.Distinct(List.Combine(
								List.Transform(
									Table.Column(Table.FirstN(exp1, 4), "tmp"), 
									each Table.ColumnNames(_)))
							),
					exp2 = 	Table.ExpandTableColumn(exp1, "tmp", nms),
					unp  = 	Table.UnpivotOtherColumns(exp2, 
								{"Company", "DataSet", "form_ID", "Date", "Предприятие_ID", "Раздел_ID","А","Б","Index"}, 
								"Header_Index", "Значение"
							),
					ren  =  Table.RenameColumns(unp,{{"А","Статья"}}),
					to   =  Table.TransformColumns(ren, {
								{"Предприятие_ID", (x)=>Number.From(x), type number}
							})
				in  to,

			Clear 	= 
				let 
					// tbl  = 	Table.RemoveColumns(Expand,{"Company", "DataSet", "form_ID","Предприятие_ID","Date"}),
					rep  =	F[Text.MultiReplace](Expand, "Б", "Статья", 
								[	101 = {"Объем пpоизводства в фактических отпускных ценах (без налогов из выручки)",
										   "Объем пpоизводства"}{1},
									102 = {"Стоимость перереработанного сыpья (неоплаченного изготовителем) в фактических ценах",
										   "Стоимость перереработанного сыpья"}{1},
									103 = {"Объем отгруженной продукции в фактических отпускных ценах (без налогов из выручки), всего",
										   "Объем отгруженной продукции"}{1},
									105 = {"Стоимость запасов готовой пpодукции в фактических отпускных ценах (без налогов из выручки) на конец периода",
										   "Стоимость запасов готовой пpодукции"}{1},
									1039 =  " переработка и консервирование других фруктов и овощей",
									10710 = "  производство хлеба и х/б изделий; производство мучных конд. изделий недлительного хранения",
									10720 = "  производство сухарей и печенья, производство мучных конд. изделий длительного хранения",
									13960 = "  производство различных текстильных изделий"
								]
							),

					
					rem  = 	F[Text.RemoveMulti](rep, "Статья", {"в т.ч.", ""})
				in	rem,

			Add 	= 
				let
					ren  = 	Table.RenameColumns(Clear, {"Б","Строка"}),
					add1 = 	Table.AddColumn(ren, "Header_ID", each Text.From([Раздел_ID]) & [Header_Index], type text),
					join = 	
						let a = F[fxJoinSpr](add1, "Header_ID", F_Form[spr.12po][Спр_Заголовки], "Header_ID", 
									{"Header"}	),
							b = F[fxJoinSpr](a, "Раздел_ID", F_Form[spr.12po][Спр_Разделы], "Раздел_ID", 
									{"РазделПрефикс","РазделИмя"}),
							c = F[fxJoinSpr](b, "Предприятие_ID", F_Form[spr.BKS] [Спр_Предприятия2], "Код",
									{"Название_исходное","Группа","Область"} )
						in  c,

					sel  = 	Table.SelectRows(join, each 
								[Header] <> null and 
								[Строка] <> "x" and 
								not List.ContainsAny({[Статья]}, {"Всего","Итого"}, Text.Contains)
							),
					add2 = 	F[fxAddMultiColumns](sel, {
								{ "СтатьяСтрока", each 
									if [Раздел_ID] = 1 or [Раздел_ID] = 2 then 
										Text.From([Строка]) & " " & Text.From([Статья]) 
									else if [Раздел_ID] = 3 then 
										Text.From([Статья]) & ", " & Text.From([Строка]) 
									else 
										Text.From([Статья]) & ", тыс. руб."
									, type text},

								{ "Уровень", each
									if [Раздел_ID] <> 2 then 
										if List.Contains ({"А".."Я","A".."Z"}, Text.Start([Статья],1)) = true
										then "Статья"
										else "Подстатья"
									else "Статья"
									, type text}
							
						}),
					to  = 	add2
				in  to,

			End 	= 
				let tbl  = 	F[Table.TypedField](Add),
					cols = 	{"Предприятие_ID","Header_Index","Раздел_ID","DataSet","Index"},
					to   = 	Table.TransformColumnTypes(tbl, List.Transform(cols, each {_, Int64.Type})),
					sel  = 	Table.SelectRows(to, each 
								not List.ContainsAny({[Header]}, {"Темп"}, Text.Contains)
							),
					lst  =  nmsTech2 & 
								{"Статья","Строка","СтатьяСтрока"} & 
								{"Значение"} &
								{"Раздел_ID","РазделПрефикс","РазделИмя"} & 
								{"Предприятие_ID","Название_исходное","Группа","Область"} & 
								{"Header_Index","Header_ID"} & 
								{"Уровень","Index"}
							,
					reo  =  Table.ReorderColumns(sel, lst, MissingField.UseNull),

					// sel2  = 	Table.SelectRows(reo, each 
					// 			not List.ContainsAny({[Предприятие_ID]}, {1..6}, Text.Contains)
					// 		),
					Result = 
						if   loadMode = "csv" // or "csv"
						then LoadReportCSV("BKS","v_12pr","_v_12pr.csv")
						else reo
				in  Result
		
		in 	End,
				
	
	rep.BKS3svod = 
		(optional datasetName as nullable text) => 
		let 
		// 	ПОДКЛЮЧЕНИЕ
			connect = 	F_Form[Connect](F_Form[GetName](datasetName),"v_3",null,2) () ,
			load    = 	
				let	tbl = connect,
				 	expand = Table.ExpandTableColumn(tbl, "Content", {"Name", "Data"}, {"Предприятие_ID", "Data"}),
					fxBKS  = (x)=> [
						SelCol = Table.SelectColumns(x, List.Generate(()=> 1, (x)=> x<=9, (x)=>x+1, (x)=>"Column" & Text.From(x) )),        // оставляем 9 колонок
						Add     = Table.AddColumn(SelCol, "Статья", each if [Column3] <> null then [Column1] & " " & [Column2] else null),
						Fill    = Table.FillDown(Add,{"Column3","Статья"}),
						SelRow  = Table.SelectRows(Fill, each try Value.Is( Number.From([Column3]), type number) = true otherwise null),   // оставляем строки с числов в названии
						SelRow2 = Table.SelectRows(SelRow, each [Column1] = null),
						SelRow3 = Table.SelectRows(SelRow2, each not List.ContainsAny( {[Column2]},{"СВОД","ИТОГО","в том числе",
								"Рук","под","Глав","20 _","в том ","из них"},Text.Contains) 
							),
						del = Table.RemoveColumns(SelRow3,{"Column1"}),
						reo = Table.ReorderColumns(del, {"Статья", "Column2", "Column3", "Column4", "Column5", "Column6", "Column7", "Column8", "Column9"}),
						// 	меняем имена столбцов, по порядку, потом по номеру название учитывая разделы
							lstOld = List.Buffer(Table.ColumnNames(reo)),
							lstNew = {"Статья","Организация", "Строка_ID", "1", "2", "3", "4", "5", "6"},
							rename = Table.RenameColumns(reo, List.Zip( {lstOld,lstNew} )),

						// 	чистим (меняем знаки из списка на "") от пробелов, переносов строки и др. Лучше сейчас, до разворота, быстрее будет
						// 	если в ячейке null, то ошибка поэтому try-otherwise
							k=(x)=> try Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " ") otherwise x,
							transform = Table.TransformColumns(rename,{"Статья", k})
						]  [transform],
					del2   = Table.RemoveColumns(expand,{"Предприятие_ID"}),
					to     = Table.TransformColumns(del2,{"Data", fxBKS})
				in	to,

			expand	= 
				let	tbl = load,
					// 	разворачиваем и переворачиваем таблицу
						Expand2 = Table.ExpandTableColumn(tbl, "Data", Table.ColumnNames(tbl{0}[Data]),Table.ColumnNames(tbl{0}[Data])),
						nms = {"Company", "DataSet", "form_ID", "Date"},
						lst = {"Организация", "Статья", "Строка_ID"},
						Unpivot = Table.UnpivotOtherColumns(Expand2, nms & lst, "Заголовок", "Значение"),

					// меняем тип колонок с трансформацией
					// Статья - чистим текст от пробелов внутри и снаружи и непечатных символов
						TransCol = Table.TransformColumns(
								Unpivot,{
									{ "Предприятие_ID", (x)=> Number.From(x), Int64.Type},
									{ "Статья",     (x)=>Text.Combine( List.RemoveMatchingItems( Text.SplitAny( Text.Clean(x),  " #(lf)#(00A0)"), {""}), " "), type text},
									{ "Строка_ID",  (x)=> try Number.From(x) otherwise null, Int64.Type},
									{ "Заголовок",  (x)=> Number.From(x), Int64.Type}, // if Number.From(x) = 0 then null else Number.From(x), Int64.Type
									{ "Значение",   (x)=> try if Number.From(x) = 0 then null else Number.Round(Number.From(x),0) otherwise null, Int64.Type}
								} )
				in 	Expand2,

			clean	= 
				let	tbl = expand,
					// 	удаляем текст внутри скобок + скобки и оставляем текст в скобках по маске
						removeText1 = F[Text.RemoveIntoBracket](tbl, "Статья", {"убыт"}),

					// 	удаляем текст по списку
						removeText2 = F[Text.RemoveMulti](removeText1, "Статья",
							{
								"в т. ч. ",
								"в том числе: ",
								" райпо между ЧУП",
								" от оператора вторичных материальных ресурсов",
								" - всего"
							}
						),

					//  функция добавления неразрывных пробелов
						Q = (x)=> Text.Repeat( Character.FromNumber(160), x),
					
					//  меняем названия статей
						removeText3 =
							Table.ReplaceValue(removeText2, (x)=>x[Статья],
								each Record.Field([
										30100 = Q(3) & "выручка от реализации за - внутрисистемной реализациии",
										38001 = "Налоги и сборы в стоимости товаров, переданных филиалами",
										10300 = "Расходы на аренду",
										10400 = "Расходы на хранение",
										10500 = "Потери товаров",
										10600 = "Расходы на ГСМ в общепите",
										11000 = "Расходы на соцстрахование",
										11600 = "справочно: расходы по завозу товаров включенных в цену",
										11700 = "справочно: проценты за пользование кредитами банка",
										40100 = Q(3) & "транзитные поставки",
										42100 = Q(3) & "продажа пайщикам ",
										42200 = Q(3) & "продажа комиссионно (за - вознаграждения)",
										42300 = Q(3) & "продажа конфиската  (за - вознаграждения)",
										44000 = "Внутрихозяйственный оборот – отпуск собственной продукции",
										46000 = "Внутрихозяйственный оборот – отпуск товаров заготовок",
										60000 = "Остаток товаров по учетным ценам:",
										60100 = Q(3) & "торговая наценка на остаток товаров",
										60200 = Q(3) & "НДС на остаток товарово",
										60300 = Q(3) & "запасы по покупным ценам",
										60310 = Q(6) & "оплаченные запасы",
										60320 = Q(6) & "запасы, полученные на условиях оплаты по мере реализации",
										53000 = "Задолженность на конец отчетного периода",
										52000 = "Получено в погашение задолженности с начала года"
									], Text.From([Строка_ID])),
								Replacer.ReplaceValue, {"Статья"}
							)

				in 	removeText3,

			Add	= 
				let	tbl = clean,
					// 	функция мультидобавления колонок

					// 	добавляем колонки с цифровым ID атрибутов
						AddColumns = F[fxAddMultiColumns](tbl, {
							// статьи с номерами строк от 100 до 399 соответствуют Части 1 (Активы), другие Части 2 (Пассивы)
							{ "Раздел_ID",  each Number.From(Text.Start(Text.From([Строка_ID]), 1)), Int64.Type},

							// номер строки + название статьи
							{ "СтатьяСтрока",  each Text.From([Строка_ID]) & " " & Text.From([Статья]) , type text},

							// уникальный номер заголовка в пределах отчета (только столбцы со значениями!): 1 знак - номер раздела, 2 знак - номер заголока в пределах раздела
							{ "Заголовок_ID", each Number.From(Text.From([Раздел_ID]) & Text.From([Заголовок])), Int64.Type},

							// если первая буква большая, то группа (1), если маленькая, то подгруппа (2)
							{ "Группа_ID", each
									if      Text.End(Text.From([Строка_ID]), 3) = "000" then 1
									else if Text.End(Text.From([Строка_ID]), 2) = "00" then 2
									else 3, Int64.Type},

							// уровень иерархии: 1 - Часть, 2 - Раздел, 3 - Статья (равно 0 в конце № строки), 4 - Подстрока (не равно 0 в конце № строки)
							{ "Уровень_ID", each
									if  [Раздел_ID] = 3 then
										if Number.Mod([Строка_ID], 1000) <> 0  then 2
										else 1
									else if [Раздел_ID] = 1 then
										if Number.Mod([Строка_ID], 100) <> 0  then 2
										else 1
									else 2
									, Int64.Type}
						}),

					// добавляем значения колонок с ID, если переключатель IsReportViewGroup = True
					// название на русском справочников раздела
						Join =
							let Join_Заголо  = Table.Join(AddColumns, 	"Заголовок_ID",    	F_Form[spr.BKS3][Спр_Заголовки],   "Заголовок_ID"),
								Join_Предпр  = Table.Join(Join_Заголо, 	"Предприятие_ID", 	F_Form[spr.BKS] [Спр_Предприятия],  "Предприятие_ID"),
								Join_Раздел  = Table.Join(Join_Предпр,  "Раздел_ID",     	F_Form[spr.BKS3][Спр_Разделы],     "Раздел_ID")
							in  Join_Раздел
				in	Join,

			end = 
				let	tbl = Add,
					to = Table.SelectRows(tbl, each
							(
								[ЗаголовокИмя] <> "Всего по торговле" and
								[ЗаголовокИмя] <> "Итого"
							)
							and
								[Строка_ID] <> 11500	//
							and
								[Значение] <> null
						)
				in	to
		
		in 	expand,	
	rep.None  = "",

// 	ДАТАСЕТЫ
	// 	Фабрика 8 Марта
		DataSet_CSV_Mart  = 
			let
				from = Record.ToTable([
						bal  = LoadReportCSV( "Mart", "bal",  "_f_bal.csv"  ),
						opu  = LoadReportCSV( "Mart", "opu",  "_f_opu.csv"  ),
						odds = LoadReportCSV( "Mart", "odds", "_f_odds.csv" ),
						osk  = LoadReportCSV( "Mart", "osk",  "_f_osk.csv" ),
						ocf  = LoadReportCSV( "Mart", "ocf",  "_f_ocf.csv" )	// нет в схеме колонок пока, поэтому ошибка
					]),
				countRow = F_Form[fxRowCount](from)
			in
				countRow,

		DataSet_XLS_Mart  = 
			let
				dataset = "Mart",
				from = Record.ToTable([
					bal  = F_Form[rep.BAL](dataset),
					opu  = F_Form[rep.OPU](dataset),
					odds = F_Form[rep.ODDS](dataset),
					osk  = F_Form[rep.OSK](dataset),
					ocf  = F_Form[rep.OCF](dataset)
				]),
				countRow = F_Form[fxRowCount](from)
			in
				countRow,
		
	// 	Гомельское ОПО
		DataSet_XLS_Homel = 
			let
				dataset = "Homel",
				from     = Record.ToTable([
					bal  = F_Form[rep.BAL] (dataset),
					opu  = F_Form[rep.OPU] (dataset),
					odds = F_Form[rep.ODDS](dataset),
					osk  = F_Form[rep.OSK] (dataset),
					pril = F_Form[rep.PRIL](dataset),
					ofr  = F_Form[rep.OFR] (dataset),
					osr  = F_Form[rep.OSR] (dataset),
					v_3  = F_Form[rep.BKS3] ([ДатаCетИмя = dataset]),
					v_7  = F_Form[rep.BKS7] ([ДатаCетИмя = dataset])
				]),
				countRow = F_Form[fxRowCount](from)
			in
				countRow,
		
		DataSet_CSV_Homel = 
			let
				dataSet = "Homel",
				from    = Record.ToTable([
						bal  = F_Form[LoadReportCSV]( dataSet, "bal",  "_f_bal.csv" ),
						opu  = F_Form[LoadReportCSV]( dataSet, "opu",  "_f_opu.csv" ),
						odds = F_Form[LoadReportCSV]( dataSet, "odds", "_f_odds.csv"),
						osk  = F_Form[LoadReportCSV]( dataSet, "osk",  "_f_osk.csv" ),
						pril = F_Form[LoadReportCSV]( dataSet, "pril", "_f_pril.csv"),
						ofr  = F_Form[LoadReportCSV]( dataSet, "ofr",  "_s_ofr.csv" ),
						osr  = F_Form[LoadReportCSV]( dataSet, "osr",  "_s_osr.csv" )
					]),
				countRow = F_Form[fxRowCount](from)
			in
				countRow,
	
	// 	Белкоопсоюз // ! заготовок нет
		DataSet_XLS_BKS   = 
			let
				dataset = "BKS",
				from = Record.ToTable([
					s_1p1   = F_Form[rep.1P3],
					s_1p2   = F_Form[rep.1P2],
					s_1torg = F_Form[rep.1TORG](dataset),
					v_3     = F_Form[rep.BKS3](dataset),
					v_7     = F_Form[rep.BKS7](dataset),
					v_12po  = F_Form[rep.12po](dataset),
				//	v_12pr	= F_Form[rep.12pr](dataset)
				  v_12pr	= LoadReportCSV("BKS","v_12pr","_v_12pr.csv") // XLS долго грузит // надо через LoadRep грузить а не напрямую
				])
			in	
				F_Form[fxRowCount](from),

		DataSet_CSV_BKS   = // второй аргумент без _, этот знак только для модели нужен
							// ? как скопом загрузить все файлы из папки и раскидать по запросам (может одним файлом)?!
			let
				from 	= Record.ToTable([
					s_1p1  	= LoadReportCSV( "BKS", "s_1p1", 	"_s_1p.csv"),
					s_1p2 	= LoadReportCSV( "BKS", "s_1p2", 	"_s_1p2.csv"),
					s_1torg = LoadReportCSV( "BKS", "s_1torg", 	"_s_1torg.csv"),
					v_3  	= LoadReportCSV( "BKS", "v_3", 		"_v_3.csv"),
					v_7  	= LoadReportCSV( "BKS", "v_7", 		"_v_7.csv"),
					v_12po	= LoadReportCSV( "BKS", "v_12po",	"_v_12po.csv"),
					v_12pr	= LoadReportCSV( "BKS", "v_12pr",	"_v_12pr.csv")
				])
			in	
				F_Form[fxRowCount](from),
				
	// 	Дрогичинское райпо
		DataSet_XLS_Drog  = 
			let
				table1 = 
					F[Table.AddedRows]({
						[ Date = #date(2022, 6, 1), Строка_ID = 104, Значение = 354, Атрибут = "Просроченная"],
						[ Date = #date(2023, 6, 1), Строка_ID = 104, Значение = 256, Атрибут = "Просроченная"]
					}),
				
				table2 = 
					F[Table.AddedRows]({
						[ Date = #date(2022, 6, 1), Строка_ID = 104, Значение = 354, Атрибут = "Просроченная"],
						[ Date = #date(2023, 6, 1), Строка_ID = 104, Значение = 256, Атрибут = "Просроченная"]
					}),

				dataset = "Drog",
				from = 
					Record.ToTable([
						bal   = F_Form[rep.BAL](dataset),
						opu   = F_Form[rep.OPU](dataset),
						osr   = table1,
						ofr   = table2,
						torg4 = F_Form[rep.4TORG]
					]),

				countRow = F_Form[fxRowCount](from)
			in
				countRow,
				
		DataSet_CSV_Drog  = 
			let 
				dataset = "Drog",
				from    = Record.ToTable([
					bal = LoadReportCSV(dataset, "bal", "_f_bal.csv"),
					opu = LoadReportCSV(dataset, "opu", "_f_opu.csv"),
					osr = LoadReportCSV(dataset, "osr", "_s_osr.csv"),
					ofr = F[Table.AddedRows]({
						    [ Company = "Drog", Date = #date(2022, 6, 1), form_ID = "ofr", Строка_ID = 2, Текущий период = 100],
							[ Company = "Drog", Date = #date(2022, 6, 1), form_ID = "ofr", Строка_ID = 1, Текущий период = 100],
							[ Company = "Drog", Date = #date(2023, 6, 1), form_ID = "ofr", Строка_ID = 1, Текущий период = 200]
							})
				]),
				countRow = F_Form[fxRowCount](from)
			in  
				countRow		
	]
in	
	D