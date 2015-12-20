
Namespace Excel

#Region " 列挙型 宣言 "

	''' <summary>
	''' テキスト ファイルのプラットフォームを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPlatform As Integer
		''' <summary>Macintosh</summary>
		xlMacintosh = 1
		''' <summary>MS-DOS</summary>
		xlMSDOS = 3
		''' <summary>Microsoft Windows</summary>
		xlWindows = 2
	End Enum

	''' <summary>
	''' クエリ テーブルにインポートするテキスト ファイルでのデータの列形式を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlTextParsingType As Integer
		''' <summary>既定値。区切り文字によってファイルが区切られます。</summary>
		xlDelimited = 1
		''' <summary>ファイルのデータが固定幅の列に配置されます。</summary>
		xlFixedWidth = 2
	End Enum

	''' <summary>
	''' 文字列を指定するために使用する区切り文字を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlTextQualifier As Integer
		''' <summary>二重引用符 (")</summary>
		xlTextQualifierDoubleQuote = 1
		''' <summary>区切り文字なし</summary>
		xlTextQualifierNone = -4142
		''' <summary>一重引用符 (')</summary>
		xlTextQualifierSingleQuote = 2
	End Enum

	''' <summary>
	''' 列を区切る方法を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlColumnDataType As Integer
		xlNone = 0
		''' <summary>一般</summary>
		xlGeneralFormat = 1
		''' <summary>テキスト</summary>
		xlTextFormat = 2
		''' <summary>MDY (月日年) 形式の日付</summary>
		xlMDYFormat = 3
		''' <summary>DMY (日月年) 形式の日付</summary>
		xlDMYFormat = 4
		''' <summary>YMD (年月日) 形式の日付</summary>
		xlYMDFormat = 5
		''' <summary>MYD (月年日) 形式の日付</summary>
		xlMYDFormat = 6
		''' <summary>DYM (日年月) 形式の日付</summary>
		xlDYMFormat = 7
		''' <summary>YDM (年日月) 形式の日付</summary>
		xlYDMFormat = 8
		''' <summary>列は区切られません。</summary>
		xlSkipColumn = 9
		''' <summary>EMD (台湾年月日) 形式の日付</summary>
		xlEMDFormat = 10
	End Enum

	''' <summary>
	''' 指定した Range オブジェクトのデータ型
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlRangeValueDataType As Integer
		xlNone = 0
		''' <summary>(既定値)  指定した Range オブジェクトが空の場合、Empty 値を取得します。これを調べるには、IsEmpty 関数を使用します。Range オブジェクトに複数のセルが含まれている場合は、値の配列を取得します。これを調べるには、IsArray 関数を使用します。</summary>
		xlRangeValueDefault = 10
		''' <summary>指定した Range オブジェクトを XML 形式で表すレコードセットを取得します。</summary>
		xlRangeValueMSPersistXML = 12
		''' <summary>指定した Range オブジェクトの値、書式、数式、および名前を XML スプレッドシート形式で取得します。</summary>
		xlRangeValueXMLSpreadsheet = 11
	End Enum

	''' <summary>
	''' 移動方向を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlDirection
		''' <summary>下に移動します。</summary>
		xlDown = -4121 ' (&HFFFFEFE7)
		''' <summary>左に移動します。</summary>
		xlToLeft = -4159 ' (&HFFFFEFC1)
		''' <summary>右に移動します。</summary>
		xlToRight = -4161 ' (&HFFFFEFBF)
		''' <summary>上に移動します。</summary>
		xlUp = -4162 ' (&HFFFFEFBE)
	End Enum

	''' <summary>
	''' 数値データをワークシートのコピー先セルでどのように計算するかを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPasteSpecialOperation As Integer
		''' <summary>コピーしたデータにコピー先セルの値を加えます。</summary>
		xlPasteSpecialOperationAdd = 2
		''' <summary>コピーしたデータをコピー先セルの値で割ります。</summary>
		xlPasteSpecialOperationDivide = 5
		''' <summary>コピーしたデータにコピー先セルの値を掛けます。</summary>
		xlPasteSpecialOperationMultiply = 4
		''' <summary>貼り付け操作で計算を実行しません。</summary>
		xlPasteSpecialOperationNone = -4142
		''' <summary>コピーしたデータからコピー先セルの値を引きます。</summary>
		xlPasteSpecialOperationSubtract = 3
	End Enum

	''' <summary>
	''' 貼り付ける範囲の属性を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPasteType As Integer
		''' <summary>すべてのオブジェクトを貼り付けます。</summary>
		xlPasteAll = -4104
		''' <summary>罫線を除くすべてのオブジェクトを貼り付けます。</summary>
		xlPasteAllExceptBorders = 7
		''' <summary>貼り付け元のセルの列幅を貼り付け先のセルに適用します。</summary>
		xlPasteColumnWidths = 8
		''' <summary>コメントを貼り付けます。</summary>
		xlPasteComments = -4144
		''' <summary>書式を貼り付けます。</summary>
		xlPasteFormats = -4122
		''' <summary>数式を貼り付けます。</summary>
		xlPasteFormulas = -4123
		''' <summary>数式と数値書式を貼り付けます。</summary>
		xlPasteFormulasAndNumberFormats = 11
		''' <summary>貼り付け元セルの入力規則を貼り付け先セルに適用します。</summary>
		xlPasteValidation = 6
		''' <summary>値だけを貼り付けます。</summary>
		xlPasteValues = -4163
		''' <summary>数値書式だけを貼り付けます。</summary>
		xlPasteValuesAndNumberFormats = 12
	End Enum

	''' <summary>
	''' Microsoft Excel で使用するグローバル定数を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum Constants As Integer
		xl3DBar = -4099
		xl3DEffects1 = 13
		xl3DEffects2 = 14
		xl3DSurface = -4103
		xlAbove = 0
		xlAccounting1 = 4
		xlAccounting2 = 5
		xlAccounting3 = 6
		xlAccounting4 = 17
		xlAdd = 2
		xlAll = -4104
		xlAllExceptBorders = 7
		xlAutomatic = -4105
		xlBar = 2
		xlBelow = 1
		xlBidi = -5000
		xlBidiCalendar = 3
		xlBoth = 1
		xlBottom = -4107
		xlCascade = 7
		xlCenter = -4108
		xlCenterAcrossSelection = 7
		xlChart4 = 2
		xlChartSeries = 17
		xlChartShort = 6
		xlChartTitles = 18
		xlChecker = 9
		xlCircle = 8
		xlClassic1 = 1
		xlClassic2 = 2
		xlClassic3 = 3
		xlClosed = 3
		xlColor1 = 7
		xlColor2 = 8
		xlColor3 = 9
		xlColumn = 3
		xlCombination = -4111
		xlComplete = 4
		xlConstants = 2
		xlContents = 2
		xlContext = -5002
		xlCorner = 2
		xlCrissCross = 16
		xlCross = 4
		xlCustom = -4114
		xlDebugCodePane = 13
		xlDefaultAutoFormat = -1
		xlDesktop = 9
		xlDiamond = 2
		xlDirect = 1
		xlDistributed = -4117
		xlDivide = 5
		xlDoubleAccounting = 5
		xlDoubleClosed = 5
		xlDoubleOpen = 4
		xlDoubleQuote = 1
		xlDrawingObject = 14
		xlEntireChart = 20
		xlExcelMenus = 1
		xlExtended = 3
		xlFill = 5
		xlFirst = 0
		xlFixedValue = 1
		xlFloating = 5
		xlFormats = -4122
		xlFormula = 5
		xlFullScript = 1
		xlGeneral = 1
		xlGray16 = 17
		xlGray25 = -4124
		xlGray50 = -4125
		xlGray75 = -4126
		xlGray8 = 18
		xlGregorian = 2
		xlGrid = 15
		xlGridline = 22
		xlHigh = -4127
		xlHindiNumerals = 3
		xlIcons = 1
		xlImmediatePane = 12
		xlInside = 2
		xlInteger = 2
		xlJustify = -4130
		xlLast = 1
		xlLastCell = 11
		xlLatin = -5001
		xlLeft = -4131
		xlLeftToRight = 2
		xlLightDown = 13
		xlLightHorizontal = 11
		xlLightUp = 14
		xlLightVertical = 12
		xlList1 = 10
		xlList2 = 11
		xlList3 = 12
		xlLocalFormat1 = 15
		xlLocalFormat2 = 16
		xlLogicalCursor = 1
		xlLong = 3
		xlLotusHelp = 2
		xlLow = -4134
		xlLTR = -5003
		xlMacrosheetCell = 7
		xlManual = -4135
		xlMaximum = 2
		xlMinimum = 4
		xlMinusValues = 3
		xlMixed = 2
		xlMixedAuthorizedScript = 4
		xlMixedScript = 3
		xlModule = -4141
		xlMultiply = 4
		xlNarrow = 1
		xlNextToAxis = 4
		xlNoDocuments = 3
		xlNone = -4142
		xlNotes = -4144
		xlOff = -4146
		xlOn = 1
		xlOpaque = 3
		xlOpen = 2
		xlOutside = 3
		xlPartial = 3
		xlPartialScript = 2
		xlPercent = 2
		xlPlus = 9
		xlPlusValues = 2
		xlReference = 4
		xlRight = -4152
		xlRTL = -5004
		xlScale = 3
		xlSemiautomatic = 2
		xlSemiGray75 = 10
		xlShort = 1
		xlShowLabel = 4
		xlShowLabelAndPercent = 5
		xlShowPercent = 3
		xlShowValue = 2
		xlSimple = -4154
		xlSingle = 2
		xlSingleAccounting = 4
		xlSingleQuote = 2
		xlSolid = 1
		xlSquare = 1
		xlStar = 5
		xlStError = 4
		xlStrict = 2
		xlSubtract = 3
		xlSystem = 1
		xlTextBox = 16
		xlTiled = 1
		xlTitleBar = 8
		xlToolbar = 1
		xlToolbarButton = 2
		xlTop = -4160
		xlTopToBottom = 1
		xlTransparent = 2
		xlTriangle = 3
		xlVeryHidden = 2
		xlVisible = 12
		xlVisualCursor = 2
		xlWatchPane = 11
		xlWide = 3
		xlWorkbookTab = 6
		xlWorksheet4 = 1
		xlWorksheetCell = 3
		xlWorksheetShort = 5
	End Enum

	''' <summary>
	''' 罫線の線の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlLineStyle
		''' <summary>実線</summary>
		xlContinuous = 1
		''' <summary>破線</summary>
		xlDash = -4115 ' (&HFFFFEFED)
		''' <summary>一点鎖線</summary>
		xlDashDot = 4
		''' <summary>二点鎖線</summary>
		xlDashDotDot = 5
		''' <summary>点線</summary>
		xlDot = -4118 ' (&HFFFFEFEA)
		''' <summary>二重線</summary>
		xlDouble = -4119 ' (&HFFFFEFE9)
		''' <summary>線なし</summary>
		xlLineStyleNone = -4142	' (&HFFFFEFD2)
		''' <summary>斜線</summary>
		xlSlantDashDot = 13
	End Enum

	''' <summary>
	''' セル範囲を囲む罫線の太さを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBorderWeight
		''' <summary>極細(最も細い罫線)</summary>
		xlHairline
		''' <summary>中</summary>
		xlMedium
		''' <summary>太い (最も太い罫線)</summary>
		xlThick
		''' <summary>細い</summary>
		xlThin
	End Enum

	''' <summary>
	''' 取得する罫線を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBordersIndex
		''' <summary>セル範囲の各セルの左上隅から右下隅への罫線</summary>
		xlDiagonalDown = 5
		''' <summary>セル範囲の各セルの左下隅から右上隅への罫線</summary>
		xlDiagonalUp = 6
		''' <summary>セル範囲の下側の罫線</summary>
		xlEdgeBottom = 9
		''' <summary>セル範囲の左側の罫線</summary>
		xlEdgeLeft = 7
		''' <summary>セル範囲の右側の罫線</summary>
		xlEdgeRight = 10
		''' <summary>セル範囲の上側の罫線</summary>
		xlEdgeTop = 8
		''' <summary>セル範囲の外枠を除く、すべてのセルの水平方向の罫線</summary>
		xlInsideHorizontal = 12
		''' <summary>セル範囲の外枠を除く、すべてのセルの垂直方向の罫線</summary>
		xlInsideVertical = 11
	End Enum

	''' <summary>
	''' 用紙に合わせてグラフのサイズを調整する方法を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlObjectSize
		''' <summary>画面に表示されているグラフの縦横比を変えずに、用紙の余白が最小になるようグラフを拡大して印刷します。</summary>
		xlFitToPage = 2
		''' <summary>グラフの縦横比を必要に応じて調整し、用紙の余白が最小になるようグラフを拡大して印刷します。</summary>
		xlFullPage = 3
		''' <summary>画面に表示されているサイズでグラフを印刷します。</summary>
		xlScreenSize = 1
	End Enum

	''' <summary>
	''' Macintosh 版 Excel の 32 ビット クリエータ コードを指定します。1480803660 (10 進)、5843454C (16 進)、または XCEL (文字列) を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlCreator
		''' <summary>Macintosh 版 Excel のクリエータ コード</summary>
		xlCreatorCode = 1480803660 ' (&H5843454C)
	End Enum

	Public Enum XlOrder
		xlDownThenOver = 1
		xlOverThenDown = 2
	End Enum

	''' <summary>
	''' ワークシートを印刷するときのページの向きを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageOrientation
		''' <summary>横モードを指定します。</summary>
		xlLandscape = 2
		''' <summary>縦モードを指定します。</summary>
		xlPortrait = 1
	End Enum

	''' <summary>
	''' 用紙のサイズを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPaperSize
		''' <summary>10 x 14 インチ</summary>
		xlPaper10x14 = 16 ' (&H10)
		''' <summary>11 x 17 インチ</summary>
		xlPaper11x17 = 17 ' (&H11)
		''' <summary>A3 (297 mm x 420 mm)</summary>
		xlPaperA3 = 8
		''' <summary>A4 (210 mm x 297 mm)</summary>
		xlPaperA4 = 9
		''' <summary>A4 (小型) (210 mm x 297 mm)</summary>
		xlPaperA4Small = 10
		''' <summary>A5 (148 mm x 210 mm)</summary>
		xlPaperA5 = 11
		''' <summary>B4 (250 mm x 354 mm)</summary>
		xlPaperB4 = 12
		''' <summary>A5 (148 mm x 210 mm)</summary>
		xlPaperB5 = 13
		''' <summary>C (17 インチ x 22 インチ)</summary>
		xlPaperCsheet = 24 ' (&H18)
		''' <summary>D (22 インチ x 34 インチ)</summary>
		xlPaperDsheet = 25 ' (&H19)
		''' <summary>封筒 10 号 (4 1/8 x 9 1/2 インチ)</summary>
		xlPaperEnvelope10 = 20 ' (&H14)
		''' <summary>封筒 11 号 (4 1/2 x 10 3/8 インチ)</summary>
		xlPaperEnvelope11 = 21 ' (&H15)
		''' <summary>封筒 12 号 (4 1/2 x 11 インチ)</summary>
		xlPaperEnvelope12 = 22 ' (&H16)
		''' <summary>封筒 14 号 (5 x 11 1/2 インチ)</summary>
		xlPaperEnvelope14 = 23 ' (&H17)
		''' <summary>封筒 9 号 (3 7/8 x 8 7/8 インチ)</summary>
		xlPaperEnvelope9 = 19 ' (&H13)
		''' <summary>封筒 B4 (250 mm x 353 mm)</summary>
		xlPaperEnvelopeB4 = 33 ' (&H21)
		''' <summary>封筒 B5 (176 mm x 250 mm)</summary>
		xlPaperEnvelopeB5 = 34 ' (&H22)
		''' <summary>封筒 B6 (176 mm x 125 mm)</summary>
		xlPaperEnvelopeB6 = 35 ' (&H23)
		''' <summary>封筒 C3 (324 mm x 458 mm)</summary>
		xlPaperEnvelopeC3 = 29 ' (&H1D)
		''' <summary>封筒 C4 (229 mm x 324 mm)</summary>
		xlPaperEnvelopeC4 = 30 ' (&H1E)
		''' <summary>封筒 C5 (162 mm x 229 mm)</summary>
		xlPaperEnvelopeC5 = 28 ' (&H1C)
		''' <summary>封筒 C6 (114 mm x 162 mm)</summary>
		xlPaperEnvelopeC6 = 31 ' (&H1F)
		''' <summary>封筒 C65 (114 mm x 229 mm)</summary>
		xlPaperEnvelopeC65 = 32	' (&H20)
		''' <summary>封筒 DL (110 x 220 mm)</summary>
		xlPaperEnvelopeDL = 27 ' (&H1B)
		''' <summary>封筒 (110 mm x 230 mm)</summary>
		xlPaperEnvelopeItaly = 36 ' (&H24)
		''' <summary>封筒モナーク (3 7/8 x 7 1/2 インチ)</summary>
		xlPaperEnvelopeMonarch = 37	' (&H25)
		''' <summary>封筒 (3 5/8 x 6 1/2 インチ)</summary>
		xlPaperEnvelopePersonal = 38 ' (&H26)
		''' <summary>E (34 インチ x 44 インチ)</summary>
		xlPaperEsheet = 26 ' (&H1A)
		''' <summary>エグゼクティブ (7 1/2 x 10 1/2 インチ)</summary>
		xlPaperExecutive = 7
		''' <summary>ドイツ リーガル複写紙 (8 1/2 x 13 インチ)</summary>
		xlPaperFanfoldLegalGerman = 41 ' (&H29)
		''' <summary>ドイツ リーガル複写紙 (8 1/2 x 13 インチ)</summary>
		xlPaperFanfoldStdGerman = 40 ' (&H28)
		''' <summary>米国標準ファンフォールド (14 7/8 x 11 インチ)</summary>
		xlPaperFanfoldUS = 39 ' (&H27)
		''' <summary>フォリオ (8 1/2 x 13 インチ)</summary>
		xlPaperFolio = 14
		''' <summary>Ledger (17 x 11 インチ)</summary>
		xlPaperLedger = 4
		''' <summary>リーガル (8 1/2 x 14 インチ)</summary>
		xlPaperLegal = 5
		''' <summary>レター (8 1/2 x 11 インチ)</summary>
		xlPaperLetter = 1
		''' <summary>レター (小型) (8 1/2 x 11 インチ)</summary>
		xlPaperLetterSmall = 2
		''' <summary>ノート (8 1/2 x 11 インチ)</summary>
		xlPaperNote = 18 ' (&H12)
		''' <summary>カート (215 mm x 275 mm)</summary>
		xlPaperQuarto = 15
		''' <summary>ステートメント (5 1/2 x 8 1/2 インチ)</summary>
		xlPaperStatement = 6
		''' <summary>タブロイド (11 x 17 インチ)</summary>
		xlPaperTabloid = 3
		''' <summary>ユーザー設定</summary>
		xlPaperUser = 256 ' (&H100)
	End Enum

	''' <summary>
	''' シートへのコメントの印刷方法を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPrintLocation
		''' <summary>ワークシート内の挿入位置にコメントを印刷します。</summary>
		xlPrintInPlace = 16	' (&H10)
		''' <summary>コメントを印刷しません。</summary>
		xlPrintNoComments = -4142 ' (&HFFFFEFD2)
		''' <summary>ワークシートの最後に、文末脚注としてコメントを印刷します。</summary>
		xlPrintSheetEnd = 1
	End Enum

	''' <summary>
	''' 表示する印刷エラーの種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPrintErrors
		''' <summary>印刷エラーは空白になります。</summary>
		xlPrintErrorsBlank = 1
		''' <summary>印刷エラーはダッシュで表示されます。</summary>
		xlPrintErrorsDash = 2
		''' <summary>印刷エラーをすべて表示します。</summary>
		xlPrintErrorsDisplayed = 0
		''' <summary>印刷エラーを使用不可として表示します。</summary>
		xlPrintErrorsNA = 3
	End Enum

	''' <summary>
	''' 図形に適用する色の変換を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoPictureColorType
		''' <summary>既定</summary>
		msoPictureAutomatic = 1
		''' <summary>白黒</summary>
		msoPictureBlackAndWhite = 3
		''' <summary>グレースケール</summary>
		msoPictureGrayscale = 2
		''' <summary>混合</summary>
		msoPictureMixed = -2 ' (&HFFFFFFFE)
		''' <summary>透かし</summary>
		msoPictureWatermark = 4
	End Enum

	''' <summary>
	''' 3 ステートのブール型 (Boolean) の値を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoTriState
		''' <summary>サポートされていません。</summary>
		msoCTrue = 1
		''' <summary>偽 (False)</summary>
		msoFalse = 0
		''' <summary>サポートされていません。</summary>
		msoTriStateMixed = -2 ' (&HFFFFFFFE)
		''' <summary>サポートされていません。</summary>
		msoTriStateToggle = -3 ' (&HFFFFFFFD)
		''' <summary>真 (True)</summary>
		msoTrue = -1 ' (&HFFFFFFFF)
	End Enum

	''' <summary>
	''' 図中のテキストに適用する背景の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBackground
		''' <summary>Excel で背景を制御します。</summary>
		xlBackgroundAutomatic = -4105 ' (&HFFFFEFF7)
		''' <summary>不透明</summary>
		xlBackgroundOpaque = 3
		''' <summary>透明</summary>
		xlBackgroundTransparent = 2
	End Enum

	''' <summary>
	''' フォントに適用される下線の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlUnderlineStyle
		''' <summary>太い二重下線</summary>
		xlUnderlineStyleDouble = -4119 ' (&HFFFFEFE9)
		''' <summary>互いに近接する細い二重下線</summary>
		xlUnderlineStyleDoubleAccounting = 5
		''' <summary>下線なし</summary>
		xlUnderlineStyleNone = -4142 ' (&HFFFFEFD2)
		''' <summary>一重下線</summary>
		xlUnderlineStyleSingle = 2
		''' <summary>サポートされていません。</summary>
		xlUnderlineStyleSingleAccounting = 4
	End Enum

	''' <summary>
	''' AutoShape オブジェクトの図形の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoAutoShapeType
		''' <summary>星 16</summary>
		msoShape16pointStar = 94
		''' <summary>星 24</summary>
		msoShape24pointStar = 95
		''' <summary>星 32</summary>
		msoShape32pointStar = 96
		''' <summary>星 4</summary>
		msoShape4pointStar = 91
		''' <summary>星 5</summary>
		msoShape5pointStar = 92
		''' <summary>星 8</summary>
		msoShape8pointStar = 93
		''' <summary>[戻る] ボタンまたは [前へ] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonBackorPrevious = 129
		''' <summary>[先頭] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonBeginning = 131
		''' <summary>既定の画像やテキストのないボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonCustom = 125
		''' <summary>[ドキュメント] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonDocument = 134
		''' <summary>[行末] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonEnd = 132
		''' <summary>[進む] ボタンまたは [次へ] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonForwardorNext = 130
		''' <summary>[ヘルプ] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonHelp = 127
		''' <summary>[ホーム] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonHome = 126
		''' <summary>[情報] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonInformation = 128
		''' <summary>[ビデオ] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonMovie = 136
		''' <summary>[戻る] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonReturn = 133
		''' <summary>[サウンド] ボタン。マウスクリックとマウスオーバーの操作をサポートします。</summary>
		msoShapeActionButtonSound = 135
		''' <summary>円弧</summary>
		msoShapeArc = 25
		''' <summary>バルーン</summary>
		msoShapeBalloon = 137
		''' <summary>曲折矢印</summary>
		msoShapeBentArrow = 41
		''' <summary>屈折矢印 (既定では上矢印)</summary>
		msoShapeBentUpArrow = 44
		''' <summary>額縁</summary>
		msoShapeBevel = 15
		''' <summary>アーチ</summary>
		msoShapeBlockArc = 20
		''' <summary>円柱</summary>
		msoShapeCan = 13
		''' <summary>山形</summary>
		msoShapeChevron = 52
		''' <summary>環状矢印</summary>
		msoShapeCircularArrow = 60
		''' <summary>雲形吹き出し</summary>
		msoShapeCloudCallout = 108
		''' <summary>十字形</summary>
		msoShapeCross = 11
		''' <summary>直方体</summary>
		msoShapeCube = 14
		''' <summary>下カーブ矢印</summary>
		msoShapeCurvedDownArrow = 48
		''' <summary>下カーブ リボン</summary>
		msoShapeCurvedDownRibbon = 100
		''' <summary>左カーブ矢印</summary>
		msoShapeCurvedLeftArrow = 46
		''' <summary>右カーブ矢印</summary>
		msoShapeCurvedRightArrow = 45
		''' <summary>上カーブ矢印</summary>
		msoShapeCurvedUpArrow = 47
		''' <summary>上カーブ リボン</summary>
		msoShapeCurvedUpRibbon = 99
		''' <summary>ひし形</summary>
		msoShapeDiamond = 4
		''' <summary>ドーナツ</summary>
		msoShapeDonut = 18
		''' <summary>中かっこ</summary>
		msoShapeDoubleBrace = 27
		''' <summary>大かっこ</summary>
		msoShapeDoubleBracket = 26
		''' <summary>小波</summary>
		msoShapeDoubleWave = 104
		''' <summary>下矢印</summary>
		msoShapeDownArrow = 36
		''' <summary>下矢印吹き出し</summary>
		msoShapeDownArrowCallout = 56
		''' <summary>下リボン</summary>
		msoShapeDownRibbon = 98
		''' <summary>爆発 1</summary>
		msoShapeExplosion1 = 89
		''' <summary>爆発 2</summary>
		msoShapeExplosion2 = 90
		''' <summary>フローチャート : 代替処理</summary>
		msoShapeFlowchartAlternateProcess = 62
		''' <summary>フローチャート : カード</summary>
		msoShapeFlowchartCard = 75
		''' <summary>フローチャート : 照合</summary>
		msoShapeFlowchartCollate = 79
		''' <summary>フローチャート : 結合子</summary>
		msoShapeFlowchartConnector = 73
		''' <summary>フローチャート : データ</summary>
		msoShapeFlowchartData = 64
		''' <summary>フローチャート : 判断</summary>
		msoShapeFlowchartDecision = 63
		''' <summary>フローチャート : 論理積ゲート</summary>
		msoShapeFlowchartDelay = 84
		''' <summary>フローチャート : 直接アクセス記憶</summary>
		msoShapeFlowchartDirectAccessStorage = 87
		''' <summary>フローチャート : 表示</summary>
		msoShapeFlowchartDisplay = 88
		''' <summary>フローチャート : 書類</summary>
		msoShapeFlowchartDocument = 67
		''' <summary>フローチャート : 抜出し</summary>
		msoShapeFlowchartExtract = 81
		''' <summary>フローチャート : 内部記憶</summary>
		msoShapeFlowchartInternalStorage = 66
		''' <summary>フローチャート : 磁気ディスク</summary>
		msoShapeFlowchartMagneticDisk = 86
		''' <summary>フローチャート : 手操作入力</summary>
		msoShapeFlowchartManualInput = 71
		''' <summary>フローチャート : 手作業</summary>
		msoShapeFlowchartManualOperation = 72
		''' <summary>フローチャート : 組合せ</summary>
		msoShapeFlowchartMerge = 82
		''' <summary>フローチャート : 複数書類</summary>
		msoShapeFlowchartMultidocument = 68
		''' <summary>フローチャート : 他ページ結合子</summary>
		msoShapeFlowchartOffpageConnector = 74
		''' <summary>フローチャート : 論理和</summary>
		msoShapeFlowchartOr = 78
		''' <summary>フローチャート : 定義済み処理</summary>
		msoShapeFlowchartPredefinedProcess = 65
		''' <summary>フローチャート : 準備</summary>
		msoShapeFlowchartPreparation = 70
		''' <summary>フローチャート : 処理</summary>
		msoShapeFlowchartProcess = 61
		''' <summary>フローチャート : せん孔テープ</summary>
		msoShapeFlowchartPunchedTape = 76
		''' <summary>フローチャート : 順次アクセス記憶</summary>
		msoShapeFlowchartSequentialAccessStorage = 85
		''' <summary>フローチャート : 分類</summary>
		msoShapeFlowchartSort = 80
		''' <summary>フローチャート : 記憶データ</summary>
		msoShapeFlowchartStoredData = 83
		''' <summary>フローチャート : 和接合</summary>
		msoShapeFlowchartSummingJunction = 77
		''' <summary>フローチャート : 端子</summary>
		msoShapeFlowchartTerminator = 69
		''' <summary>メモ</summary>
		msoShapeFoldedCorner = 16
		''' <summary>ハート</summary>
		msoShapeHeart = 21
		''' <summary>六角形</summary>
		msoShapeHexagon = 10
		''' <summary>横巻き</summary>
		msoShapeHorizontalScroll = 102
		''' <summary>二等辺三角形</summary>
		msoShapeIsoscelesTriangle = 7
		''' <summary>左矢印</summary>
		msoShapeLeftArrow = 34
		''' <summary>左矢印吹き出し</summary>
		msoShapeLeftArrowCallout = 54
		''' <summary>左中かっこ</summary>
		msoShapeLeftBrace = 31
		''' <summary>左大かっこ</summary>
		msoShapeLeftBracket = 29
		''' <summary>左右矢印</summary>
		msoShapeLeftRightArrow = 37
		''' <summary>左右矢印吹き出し</summary>
		msoShapeLeftRightArrowCallout = 57
		''' <summary>三方向矢印</summary>
		msoShapeLeftRightUpArrow = 40
		''' <summary>二方向矢印</summary>
		msoShapeLeftUpArrow = 43
		''' <summary>稲妻</summary>
		msoShapeLightningBolt = 22
		''' <summary>線吹き出し 1 (枠付き)</summary>
		msoShapeLineCallout1 = 109
		''' <summary>強調線吹き出し 1</summary>
		msoShapeLineCallout1AccentBar = 113
		''' <summary>強調線吹き出し 1 (枠付き)</summary>
		msoShapeLineCallout1BorderandAccentBar = 121
		''' <summary>線吹き出し 1</summary>
		msoShapeLineCallout1NoBorder = 117
		''' <summary>線吹き出し 2 (枠付き)</summary>
		msoShapeLineCallout2 = 110
		''' <summary>強調線吹き出し 2</summary>
		msoShapeLineCallout2AccentBar = 114
		''' <summary>強調線吹き出し 2 (枠付き)</summary>
		msoShapeLineCallout2BorderandAccentBar = 122
		''' <summary>線吹き出し 2</summary>
		msoShapeLineCallout2NoBorder = 118
		''' <summary>線吹き出し 3 (枠付き)</summary>
		msoShapeLineCallout3 = 111
		''' <summary>強調線吹き出し 3</summary>
		msoShapeLineCallout3AccentBar = 115
		''' <summary>強調線吹き出し 3 (枠付き)</summary>
		msoShapeLineCallout3BorderandAccentBar = 123
		''' <summary>線吹き出し 3</summary>
		msoShapeLineCallout3NoBorder = 119
		''' <summary>線吹き出し 4 (枠付き)</summary>
		msoShapeLineCallout4 = 112
		''' <summary>強調線吹き出し 4</summary>
		msoShapeLineCallout4AccentBar = 116
		''' <summary>強調線吹き出し 4 (枠付き)</summary>
		msoShapeLineCallout4BorderandAccentBar = 124
		''' <summary>線吹き出し 4</summary>
		msoShapeLineCallout4NoBorder = 120
		''' <summary>値の取得のみ可能です。他の状態の組み合わせを示します。</summary>
		msoShapeMixed = -2
		''' <summary>月</summary>
		msoShapeMoon = 24
		''' <summary>禁止</summary>
		msoShapeNoSymbol = 19
		''' <summary>V 字形矢印</summary>
		msoShapeNotchedRightArrow = 50
		''' <summary>サポートされていません。</summary>
		msoShapeNotPrimitive = 138
		''' <summary>八角形</summary>
		msoShapeOctagon = 6
		''' <summary>楕円</summary>
		msoShapeOval = 9
		''' <summary>円形吹き出し</summary>
		msoShapeOvalCallout = 107
		''' <summary>平行四角形</summary>
		msoShapeParallelogram = 2
		''' <summary>ホームベース</summary>
		msoShapePentagon = 51
		''' <summary>ブローチ</summary>
		msoShapePlaque = 28
		''' <summary>四方向矢印</summary>
		msoShapeQuadArrow = 39
		''' <summary>四方向矢印吹き出し</summary>
		msoShapeQuadArrowCallout = 59
		''' <summary>四角形</summary>
		msoShapeRectangle = 1
		''' <summary>四角形吹き出し</summary>
		msoShapeRectangularCallout = 105
		''' <summary>五角形</summary>
		msoShapeRegularPentagon = 12
		''' <summary>右矢印</summary>
		msoShapeRightArrow = 33
		''' <summary>右矢印吹き出し</summary>
		msoShapeRightArrowCallout = 53
		''' <summary>右中かっこ</summary>
		msoShapeRightBrace = 32
		''' <summary>右大かっこ</summary>
		msoShapeRightBracket = 30
		''' <summary>直角三角形</summary>
		msoShapeRightTriangle = 8
		''' <summary>角丸四角形</summary>
		msoShapeRoundedRectangle = 5
		''' <summary>角丸四角形吹き出し</summary>
		msoShapeRoundedRectangularCallout = 106
		''' <summary>スマイル</summary>
		msoShapeSmileyFace = 17
		''' <summary>ストライプ矢印</summary>
		msoShapeStripedRightArrow = 49
		''' <summary>太陽</summary>
		msoShapeSun = 23
		''' <summary>台形</summary>
		msoShapeTrapezoid = 3
		''' <summary>上矢印</summary>
		msoShapeUpArrow = 35
		''' <summary>上矢印吹き出し</summary>
		msoShapeUpArrowCallout = 55
		''' <summary>上下矢印</summary>
		msoShapeUpDownArrow = 38
		''' <summary>上下矢印吹き出し</summary>
		msoShapeUpDownArrowCallout = 58
		''' <summary>上リボン</summary>
		msoShapeUpRibbon = 97
		''' <summary>U ターン矢印</summary>
		msoShapeUTurnArrow = 42
		''' <summary>縦巻き</summary>
		msoShapeVerticalScroll = 101
		''' <summary>大波</summary>
		msoShapeWave = 103
	End Enum

	''' <summary>
	''' XML データ ファイルを開く方法を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlXmlLoadOption
		''' <summary>XML データ ファイルの内容を XML リストに配置します。</summary>
		xlXmlLoadImportToList = 2
		''' <summary>XML データ ファイルのスキーマを [XML データ構造] 作業ウィンドウに表示します。</summary>
		xlXmlLoadMapXml = 3
		''' <summary>XML データ ファイルを開きます。ファイルの内容はフラット化されます。</summary>
		xlXmlLoadOpenXml = 1
		''' <summary>ファイルを開く方法を選択するよう求めるメッセージが表示されます。	</summary>
		xlXmlLoadPromptUser = 0
	End Enum

	''' <summary>
	''' 全画面表示で改ページするか、印刷領域のみで改ページするかを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageBreakExtent
		''' <summary>全画面</summary>
		xlPageBreakFull = 1
		''' <summary>印刷領域のみ</summary>
		xlPageBreakPartial = 2
	End Enum

	''' <summary>
	''' ワークシートの改ページ位置を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageBreak
		''' <summary>自動的に改ページを追加します。</summary>
		xlPageBreakAutomatic = -4105 ' (&HFFFFEFF7)
		''' <summary>手動で改ページを挿入します。</summary>
		xlPageBreakManual = -4135 ' (&HFFFFEFD9)
		''' <summary>ワークシートに改ページを挿入しません。</summary>
		xlPageBreakNone = -4142 ' (&HFFFFEFD2)
	End Enum

	''' <summary>
	''' ワークシートの種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlSheetType
		''' <summary>グラフ</summary>
		xlChart = -4109	' (&HFFFFEFF3)
		''' <summary>ダイアログ シート</summary>
		xlDialogSheet = -4116 ' (&HFFFFEFEC)
		''' <summary>Excel 4.0 インターナショナル マクロ シート</summary>
		xlExcel4IntlMacroSheet = 4
		''' <summary>Excel 4.0 マクロ シート</summary>
		xlExcel4MacroSheet = 3
		''' <summary>ワークシート</summary>
		xlWorksheet = -4167 ' (&HFFFFEFB9)
	End Enum

	''' <summary>
	''' セル範囲をコピーする方法を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlFillWith
		''' <summary>内容と書式をコピーします。</summary>
		xlFillWithAll = -4104 ' (&HFFFFEFF8)
		''' <summary>内容のみをコピーします。</summary>
		xlFillWithContents = 2
		''' <summary>書式のみをコピーします。</summary>
		xlFillWithFormats = -4122 ' (&HFFFFEFE6)
	End Enum

	''' <summary>
	''' 引き出し線の種類指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoCalloutType
		''' <summary>水平または垂直の単一セグメントの引き出し線</summary>
		msoCalloutOne = 1
		''' <summary>自由に回転する単一セグメントの引き出し線</summary>
		msoCalloutTwo = 2
		''' <summary></summary>
		msoCalloutMixed = -2 ' (&HFFFFFFFE)
		''' <summary>2 つのセグメントから成る引き出し線</summary>
		msoCalloutThree = 3
		''' <summary>3 つのセグメントから成る引き出し線</summary>
		msoCalloutFour = 4
	End Enum

	''' <summary>
	''' コネクタの種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoConnectorType
		msoConnectorElbow = 2
		msoConnectorTypeMixed = -2 ' (&HFFFFFFFE)
		msoConnectorCurve = 3
		msoConnectorStraight = 1
	End Enum

	''' <summary>
	''' 図表の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoDiagramType
		''' <summary>連続したドーナツ型図表を持つ、プロセスを表す図表。</summary>
		msoDiagramCycle = 2
		''' <summary>混合型の図表。</summary>
		msoDiagramMixed = -2 ' (&HFFFFFFFE)
		''' <summary>階層構造の関係を表す図表。</summary>
		msoDiagramOrgChart = 1
		''' <summary>基礎構造的な関係を表す図表。</summary>
		msoDiagramPyramid = 4
		''' <summary>中核となる要素との関係を表す図表。</summary>
		msoDiagramRadial = 3
		''' <summary>ゴールまでのステップを表す図表。</summary>
		msoDiagramTarget = 6
		''' <summary>要素間で重なり合う領域を表す図表。</summary>
		msoDiagramVenn = 5
	End Enum

	''' <summary>
	''' フォーム コントロールの種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlFormControl
		''' <summary>ボタン</summary>
		xlButtonControl = 0
		''' <summary>チェック ボックス</summary>
		xlCheckBox = 1
		''' <summary>コンボ ボックス</summary>
		xlDropDown = 2
		''' <summary>テキスト ボックス</summary>
		xlEditBox = 3
		''' <summary>グループ ボックス</summary>
		xlGroupBox = 4
		''' <summary>ラベル</summary>
		xlLabel = 5
		''' <summary>リスト ボックス</summary>
		xlListBox = 6
		''' <summary>オプション ボタン</summary>
		xlOptionButton = 7
		''' <summary>スクロール バー</summary>
		xlScrollBar = 8
		''' <summary>スピン ボタン</summary>
		xlSpinner = 9
	End Enum

	''' <summary>
	''' 文字列の向きを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoTextOrientation
		''' <summary>右下がり。</summary>
		msoTextOrientationDownward = 3
		''' <summary>横書き。</summary>
		msoTextOrientationHorizontal = 1
		''' <summary>東アジア言語のサポート用の横書きおよび回転。</summary>
		msoTextOrientationHorizontalRotatedFarEast = 6
		''' <summary>サポートされていません。</summary>
		msoTextOrientationMixed = -2 ' (&HFFFFFFFE)
		''' <summary>右上がり。</summary>
		msoTextOrientationUpward = 2
		''' <summary>縦書き。</summary>
		msoTextOrientationVertical = 5
		''' <summary>東アジア言語のサポート用の縦書き。</summary>
		msoTextOrientationVerticalFarEast = 4
	End Enum

	''' <summary>
	''' WordArt オブジェクトで使用する特殊効果テキストを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoPresetTextEffect
		''' <summary>1 番目の特殊効果</summary>
		msoTextEffect1 = 0
		''' <summary>10 番目の特殊効果</summary>
		msoTextEffect10 = 9
		''' <summary>11 番目の特殊効果</summary>
		msoTextEffect11 = 10
		''' <summary>12 番目の特殊効果</summary>
		msoTextEffect12 = 11
		''' <summary>13 番目の特殊効果</summary>
		msoTextEffect13 = 12
		''' <summary>14 番目の特殊効果</summary>
		msoTextEffect14 = 13
		''' <summary>15 番目の特殊効果</summary>
		msoTextEffect15 = 14
		''' <summary>16 番目の特殊効果</summary>
		msoTextEffect16 = 15
		''' <summary>17 番目の特殊効果</summary>
		msoTextEffect17 = 16
		''' <summary>18 番目の特殊効果</summary>
		msoTextEffect18 = 17
		''' <summary>19 番目の特殊効果</summary>
		msoTextEffect19 = 18
		''' <summary>2 番目の特殊効果</summary>
		msoTextEffect2 = 1
		''' <summary>20 番目の特殊効果</summary>
		msoTextEffect20 = 19
		''' <summary>21 番目の特殊効果</summary>
		msoTextEffect21 = 20
		''' <summary>22 番目の特殊効果</summary>
		msoTextEffect22 = 21
		''' <summary>23 番目の特殊効果</summary>
		msoTextEffect23 = 22
		''' <summary>24 番目の特殊効果</summary>
		msoTextEffect24 = 23
		''' <summary>25 番目の特殊効果</summary>
		msoTextEffect25 = 24
		''' <summary>26 番目の特殊効果</summary>
		msoTextEffect26 = 25
		''' <summary>27 番目の特殊効果</summary>
		msoTextEffect27 = 26
		''' <summary>28 番目の特殊効果</summary>
		msoTextEffect28 = 27
		''' <summary>29 番目の特殊効果</summary>
		msoTextEffect29 = 28
		''' <summary>3 番目の特殊効果</summary>
		msoTextEffect3 = 2
		''' <summary>30 番目の特殊効果</summary>
		msoTextEffect30 = 29
		''' <summary>4 番目の特殊効果</summary>
		msoTextEffect4 = 3
		''' <summary>5 番目の特殊効果</summary>
		msoTextEffect5 = 4
		''' <summary>6 番目の特殊効果</summary>
		msoTextEffect6 = 5
		''' <summary>7 番目の特殊効果</summary>
		msoTextEffect7 = 6
		''' <summary>8 番目の特殊効果</summary>
		msoTextEffect8 = 7
		''' <summary>9 番目の特殊効果</summary>
		msoTextEffect9 = 8
		''' <summary>未使用</summary>
		msoTextEffectMixed = -2 ' (&HFFFFFFFE)
	End Enum

	''' <summary>
	''' セグメントの種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoSegmentType
		''' <summary>曲線</summary>
		msoSegmentCurve = 1
		''' <summary>直線</summary>
		msoSegmentLine = 0
	End Enum

	''' <summary>
	''' 節点の編集の種類を指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoEditingType
		''' <summary>編集の種類は、接続しているセグメントの種類に対応します。</summary>
		msoEditingAuto = 0
		''' <summary>コーナーの節点</summary>
		msoEditingCorner = 1
		''' <summary>スムーズな節点</summary>
		msoEditingSmooth = 2
		''' <summary>対称的な節点</summary>
		msoEditingSymmetric = 3
	End Enum

	''' <summary>
	''' グラフの塗りつぶしのパターンまたは塗りつぶしのオブジェクトを指定します。
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPattern
		''' <summary>パターンは Excel によって制御されます。</summary>
		xlPatternAutomatic = -4105
		''' <summary>市松模様のパターンです。</summary>
		xlPatternChecker = 9
		''' <summary>網目模様のパターンです。</summary>
		xlPatternCrissCross = 16
		''' <summary>右下がりの濃い対角線のパターンです。</summary>
		xlPatternDown = -4121
		''' <summary>16% の灰色です。</summary>
		xlPatternGray16 = 17
		''' <summary>25% の灰色です。</summary>
		xlPatternGray25 = -4124
		''' <summary>50% の灰色です。</summary>
		xlPatternGray50 = -4125
		''' <summary>75% の灰色です。</summary>
		xlPatternGray75 = -4126
		''' <summary>8% の灰色です。</summary>
		xlPatternGray8 = 18
		''' <summary>格子模様のパターンです。</summary>
		xlPatternGrid = 15
		''' <summary>濃い横線のパターンです。</summary>
		xlPatternHorizontal = -4128
		''' <summary>右下がりの薄い対角線のパターンです。</summary>
		xlPatternLightDown = 13
		''' <summary>薄い横線のパターンです。</summary>
		xlPatternLightHorizontal = 11
		''' <summary>右上がりの薄い対角線のパターンです。</summary>
		xlPatternLightUp = 14
		''' <summary>薄い縦線のパターンです。</summary>
		xlPatternLightVertical = 12
		''' <summary>パターンはありません。</summary>
		xlPatternNone = -4142
		''' <summary>75% の濃いモアレ パターンです。</summary>
		xlPatternSemiGray75 = 10
		''' <summary>純色です。</summary>
		xlPatternSolid = 1
		''' <summary>右上がりの濃い対角線のパターンです。</summary>
		xlPatternUp = -4162
		''' <summary>濃い縦線のパターンです。</summary>
		xlPatternVertical = -4166
	End Enum

	''' <summary>
	''' 実行するべき自動マクロ
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlRunAutoMacro
		''' <summary>Auto_Activateマクロ</summary>
		xlAutoActivate = 3
		''' <summary>Auto_Closeマクロ</summary>
		xlAutoClose = 2
		''' <summary>Auto_Deactivateマクロ</summary>
		xlAutoDeactivate = 4
		''' <summary>Auto_Openマクロ</summary>
		xlAutoOpen = 1
	End Enum

    ''' <summary>
    ''' 変換するファイルフォーマット
    ''' </summary>
    ''' <remarks>https://msdn.microsoft.com/JA-JP/library/office/ff195006.aspx</remarks>
    Public Enum FixedFormatType As Integer
        ''' <summary>
        ''' PDFファイル
        ''' </summary>
        ''' <remarks></remarks>
        PDF
        ''' <summary>
        ''' XPSファイル
        ''' </summary>
        ''' <remarks></remarks>
        XPS
    End Enum

    ''' <summary>
    ''' 変換品質
    ''' </summary>
    ''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff838396.aspx</remarks>
    Public Enum FixedFormatQuality As Integer
        ''' <summary>
        ''' 標準品質
        ''' </summary>
        ''' <remarks></remarks>
        QualityStandard
        ''' <summary>
        ''' 最小限品質
        ''' </summary>
        ''' <remarks></remarks>
        QualityMinimum
    End Enum

    ''' <summary>
    ''' セルの挿入時にセルをシフトする方向
    ''' </summary>
    ''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff837618.aspx</remarks>
    Public Enum XlInsertShiftDirection As Integer
        ''' <summary>
        ''' セルを下にシフト
        ''' </summary>
        xlShiftDown
        ''' <summary>
        ''' セルを右にシフト
        ''' </summary>
        xlShiftToRight
        ''' <summary>
        ''' 指定なし
        ''' </summary>
        none = 99
    End Enum

#End Region

End Namespace
