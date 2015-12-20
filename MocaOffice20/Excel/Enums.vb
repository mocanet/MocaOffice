
Namespace Excel

#Region " �񋓌^ �錾 "

	''' <summary>
	''' �e�L�X�g �t�@�C���̃v���b�g�t�H�[�����w�肵�܂��B
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
	''' �N�G�� �e�[�u���ɃC���|�[�g����e�L�X�g �t�@�C���ł̃f�[�^�̗�`�����w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlTextParsingType As Integer
		''' <summary>����l�B��؂蕶���ɂ���ăt�@�C������؂��܂��B</summary>
		xlDelimited = 1
		''' <summary>�t�@�C���̃f�[�^���Œ蕝�̗�ɔz�u����܂��B</summary>
		xlFixedWidth = 2
	End Enum

	''' <summary>
	''' ��������w�肷�邽�߂Ɏg�p�����؂蕶�����w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlTextQualifier As Integer
		''' <summary>��d���p�� (")</summary>
		xlTextQualifierDoubleQuote = 1
		''' <summary>��؂蕶���Ȃ�</summary>
		xlTextQualifierNone = -4142
		''' <summary>��d���p�� (')</summary>
		xlTextQualifierSingleQuote = 2
	End Enum

	''' <summary>
	''' �����؂���@���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlColumnDataType As Integer
		xlNone = 0
		''' <summary>���</summary>
		xlGeneralFormat = 1
		''' <summary>�e�L�X�g</summary>
		xlTextFormat = 2
		''' <summary>MDY (�����N) �`���̓��t</summary>
		xlMDYFormat = 3
		''' <summary>DMY (�����N) �`���̓��t</summary>
		xlDMYFormat = 4
		''' <summary>YMD (�N����) �`���̓��t</summary>
		xlYMDFormat = 5
		''' <summary>MYD (���N��) �`���̓��t</summary>
		xlMYDFormat = 6
		''' <summary>DYM (���N��) �`���̓��t</summary>
		xlDYMFormat = 7
		''' <summary>YDM (�N����) �`���̓��t</summary>
		xlYDMFormat = 8
		''' <summary>��͋�؂��܂���B</summary>
		xlSkipColumn = 9
		''' <summary>EMD (��p�N����) �`���̓��t</summary>
		xlEMDFormat = 10
	End Enum

	''' <summary>
	''' �w�肵�� Range �I�u�W�F�N�g�̃f�[�^�^
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlRangeValueDataType As Integer
		xlNone = 0
		''' <summary>(����l)  �w�肵�� Range �I�u�W�F�N�g����̏ꍇ�AEmpty �l���擾���܂��B����𒲂ׂ�ɂ́AIsEmpty �֐����g�p���܂��BRange �I�u�W�F�N�g�ɕ����̃Z�����܂܂�Ă���ꍇ�́A�l�̔z����擾���܂��B����𒲂ׂ�ɂ́AIsArray �֐����g�p���܂��B</summary>
		xlRangeValueDefault = 10
		''' <summary>�w�肵�� Range �I�u�W�F�N�g�� XML �`���ŕ\�����R�[�h�Z�b�g���擾���܂��B</summary>
		xlRangeValueMSPersistXML = 12
		''' <summary>�w�肵�� Range �I�u�W�F�N�g�̒l�A�����A�����A����і��O�� XML �X�v���b�h�V�[�g�`���Ŏ擾���܂��B</summary>
		xlRangeValueXMLSpreadsheet = 11
	End Enum

	''' <summary>
	''' �ړ��������w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlDirection
		''' <summary>���Ɉړ����܂��B</summary>
		xlDown = -4121 ' (&HFFFFEFE7)
		''' <summary>���Ɉړ����܂��B</summary>
		xlToLeft = -4159 ' (&HFFFFEFC1)
		''' <summary>�E�Ɉړ����܂��B</summary>
		xlToRight = -4161 ' (&HFFFFEFBF)
		''' <summary>��Ɉړ����܂��B</summary>
		xlUp = -4162 ' (&HFFFFEFBE)
	End Enum

	''' <summary>
	''' ���l�f�[�^�����[�N�V�[�g�̃R�s�[��Z���łǂ̂悤�Ɍv�Z���邩���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPasteSpecialOperation As Integer
		''' <summary>�R�s�[�����f�[�^�ɃR�s�[��Z���̒l�������܂��B</summary>
		xlPasteSpecialOperationAdd = 2
		''' <summary>�R�s�[�����f�[�^���R�s�[��Z���̒l�Ŋ���܂��B</summary>
		xlPasteSpecialOperationDivide = 5
		''' <summary>�R�s�[�����f�[�^�ɃR�s�[��Z���̒l���|���܂��B</summary>
		xlPasteSpecialOperationMultiply = 4
		''' <summary>�\��t������Ōv�Z�����s���܂���B</summary>
		xlPasteSpecialOperationNone = -4142
		''' <summary>�R�s�[�����f�[�^����R�s�[��Z���̒l�������܂��B</summary>
		xlPasteSpecialOperationSubtract = 3
	End Enum

	''' <summary>
	''' �\��t����͈͂̑������w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPasteType As Integer
		''' <summary>���ׂẴI�u�W�F�N�g��\��t���܂��B</summary>
		xlPasteAll = -4104
		''' <summary>�r�����������ׂẴI�u�W�F�N�g��\��t���܂��B</summary>
		xlPasteAllExceptBorders = 7
		''' <summary>�\��t�����̃Z���̗񕝂�\��t����̃Z���ɓK�p���܂��B</summary>
		xlPasteColumnWidths = 8
		''' <summary>�R�����g��\��t���܂��B</summary>
		xlPasteComments = -4144
		''' <summary>������\��t���܂��B</summary>
		xlPasteFormats = -4122
		''' <summary>������\��t���܂��B</summary>
		xlPasteFormulas = -4123
		''' <summary>�����Ɛ��l������\��t���܂��B</summary>
		xlPasteFormulasAndNumberFormats = 11
		''' <summary>�\��t�����Z���̓��͋K����\��t����Z���ɓK�p���܂��B</summary>
		xlPasteValidation = 6
		''' <summary>�l������\��t���܂��B</summary>
		xlPasteValues = -4163
		''' <summary>���l����������\��t���܂��B</summary>
		xlPasteValuesAndNumberFormats = 12
	End Enum

	''' <summary>
	''' Microsoft Excel �Ŏg�p����O���[�o���萔���w�肵�܂��B
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
	''' �r���̐��̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlLineStyle
		''' <summary>����</summary>
		xlContinuous = 1
		''' <summary>�j��</summary>
		xlDash = -4115 ' (&HFFFFEFED)
		''' <summary>��_����</summary>
		xlDashDot = 4
		''' <summary>��_����</summary>
		xlDashDotDot = 5
		''' <summary>�_��</summary>
		xlDot = -4118 ' (&HFFFFEFEA)
		''' <summary>��d��</summary>
		xlDouble = -4119 ' (&HFFFFEFE9)
		''' <summary>���Ȃ�</summary>
		xlLineStyleNone = -4142	' (&HFFFFEFD2)
		''' <summary>�ΐ�</summary>
		xlSlantDashDot = 13
	End Enum

	''' <summary>
	''' �Z���͈͂��͂ތr���̑������w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBorderWeight
		''' <summary>�ɍ�(�ł��ׂ��r��)</summary>
		xlHairline
		''' <summary>��</summary>
		xlMedium
		''' <summary>���� (�ł������r��)</summary>
		xlThick
		''' <summary>�ׂ�</summary>
		xlThin
	End Enum

	''' <summary>
	''' �擾����r�����w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBordersIndex
		''' <summary>�Z���͈͂̊e�Z���̍��������E�����ւ̌r��</summary>
		xlDiagonalDown = 5
		''' <summary>�Z���͈͂̊e�Z���̍���������E����ւ̌r��</summary>
		xlDiagonalUp = 6
		''' <summary>�Z���͈͂̉����̌r��</summary>
		xlEdgeBottom = 9
		''' <summary>�Z���͈͂̍����̌r��</summary>
		xlEdgeLeft = 7
		''' <summary>�Z���͈͂̉E���̌r��</summary>
		xlEdgeRight = 10
		''' <summary>�Z���͈͂̏㑤�̌r��</summary>
		xlEdgeTop = 8
		''' <summary>�Z���͈͂̊O�g�������A���ׂẴZ���̐��������̌r��</summary>
		xlInsideHorizontal = 12
		''' <summary>�Z���͈͂̊O�g�������A���ׂẴZ���̐��������̌r��</summary>
		xlInsideVertical = 11
	End Enum

	''' <summary>
	''' �p���ɍ��킹�ăO���t�̃T�C�Y�𒲐�������@���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlObjectSize
		''' <summary>��ʂɕ\������Ă���O���t�̏c�����ς����ɁA�p���̗]�����ŏ��ɂȂ�悤�O���t���g�債�Ĉ�����܂��B</summary>
		xlFitToPage = 2
		''' <summary>�O���t�̏c�����K�v�ɉ����Ē������A�p���̗]�����ŏ��ɂȂ�悤�O���t���g�債�Ĉ�����܂��B</summary>
		xlFullPage = 3
		''' <summary>��ʂɕ\������Ă���T�C�Y�ŃO���t��������܂��B</summary>
		xlScreenSize = 1
	End Enum

	''' <summary>
	''' Macintosh �� Excel �� 32 �r�b�g �N���G�[�^ �R�[�h���w�肵�܂��B1480803660 (10 �i)�A5843454C (16 �i)�A�܂��� XCEL (������) ���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlCreator
		''' <summary>Macintosh �� Excel �̃N���G�[�^ �R�[�h</summary>
		xlCreatorCode = 1480803660 ' (&H5843454C)
	End Enum

	Public Enum XlOrder
		xlDownThenOver = 1
		xlOverThenDown = 2
	End Enum

	''' <summary>
	''' ���[�N�V�[�g���������Ƃ��̃y�[�W�̌������w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageOrientation
		''' <summary>�����[�h���w�肵�܂��B</summary>
		xlLandscape = 2
		''' <summary>�c���[�h���w�肵�܂��B</summary>
		xlPortrait = 1
	End Enum

	''' <summary>
	''' �p���̃T�C�Y���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPaperSize
		''' <summary>10 x 14 �C���`</summary>
		xlPaper10x14 = 16 ' (&H10)
		''' <summary>11 x 17 �C���`</summary>
		xlPaper11x17 = 17 ' (&H11)
		''' <summary>A3 (297 mm x 420 mm)</summary>
		xlPaperA3 = 8
		''' <summary>A4 (210 mm x 297 mm)</summary>
		xlPaperA4 = 9
		''' <summary>A4 (���^) (210 mm x 297 mm)</summary>
		xlPaperA4Small = 10
		''' <summary>A5 (148 mm x 210 mm)</summary>
		xlPaperA5 = 11
		''' <summary>B4 (250 mm x 354 mm)</summary>
		xlPaperB4 = 12
		''' <summary>A5 (148 mm x 210 mm)</summary>
		xlPaperB5 = 13
		''' <summary>C (17 �C���` x 22 �C���`)</summary>
		xlPaperCsheet = 24 ' (&H18)
		''' <summary>D (22 �C���` x 34 �C���`)</summary>
		xlPaperDsheet = 25 ' (&H19)
		''' <summary>���� 10 �� (4 1/8 x 9 1/2 �C���`)</summary>
		xlPaperEnvelope10 = 20 ' (&H14)
		''' <summary>���� 11 �� (4 1/2 x 10 3/8 �C���`)</summary>
		xlPaperEnvelope11 = 21 ' (&H15)
		''' <summary>���� 12 �� (4 1/2 x 11 �C���`)</summary>
		xlPaperEnvelope12 = 22 ' (&H16)
		''' <summary>���� 14 �� (5 x 11 1/2 �C���`)</summary>
		xlPaperEnvelope14 = 23 ' (&H17)
		''' <summary>���� 9 �� (3 7/8 x 8 7/8 �C���`)</summary>
		xlPaperEnvelope9 = 19 ' (&H13)
		''' <summary>���� B4 (250 mm x 353 mm)</summary>
		xlPaperEnvelopeB4 = 33 ' (&H21)
		''' <summary>���� B5 (176 mm x 250 mm)</summary>
		xlPaperEnvelopeB5 = 34 ' (&H22)
		''' <summary>���� B6 (176 mm x 125 mm)</summary>
		xlPaperEnvelopeB6 = 35 ' (&H23)
		''' <summary>���� C3 (324 mm x 458 mm)</summary>
		xlPaperEnvelopeC3 = 29 ' (&H1D)
		''' <summary>���� C4 (229 mm x 324 mm)</summary>
		xlPaperEnvelopeC4 = 30 ' (&H1E)
		''' <summary>���� C5 (162 mm x 229 mm)</summary>
		xlPaperEnvelopeC5 = 28 ' (&H1C)
		''' <summary>���� C6 (114 mm x 162 mm)</summary>
		xlPaperEnvelopeC6 = 31 ' (&H1F)
		''' <summary>���� C65 (114 mm x 229 mm)</summary>
		xlPaperEnvelopeC65 = 32	' (&H20)
		''' <summary>���� DL (110 x 220 mm)</summary>
		xlPaperEnvelopeDL = 27 ' (&H1B)
		''' <summary>���� (110 mm x 230 mm)</summary>
		xlPaperEnvelopeItaly = 36 ' (&H24)
		''' <summary>�������i�[�N (3 7/8 x 7 1/2 �C���`)</summary>
		xlPaperEnvelopeMonarch = 37	' (&H25)
		''' <summary>���� (3 5/8 x 6 1/2 �C���`)</summary>
		xlPaperEnvelopePersonal = 38 ' (&H26)
		''' <summary>E (34 �C���` x 44 �C���`)</summary>
		xlPaperEsheet = 26 ' (&H1A)
		''' <summary>�G�O�[�N�e�B�u (7 1/2 x 10 1/2 �C���`)</summary>
		xlPaperExecutive = 7
		''' <summary>�h�C�c ���[�K�����ʎ� (8 1/2 x 13 �C���`)</summary>
		xlPaperFanfoldLegalGerman = 41 ' (&H29)
		''' <summary>�h�C�c ���[�K�����ʎ� (8 1/2 x 13 �C���`)</summary>
		xlPaperFanfoldStdGerman = 40 ' (&H28)
		''' <summary>�č��W���t�@���t�H�[���h (14 7/8 x 11 �C���`)</summary>
		xlPaperFanfoldUS = 39 ' (&H27)
		''' <summary>�t�H���I (8 1/2 x 13 �C���`)</summary>
		xlPaperFolio = 14
		''' <summary>Ledger (17 x 11 �C���`)</summary>
		xlPaperLedger = 4
		''' <summary>���[�K�� (8 1/2 x 14 �C���`)</summary>
		xlPaperLegal = 5
		''' <summary>���^�[ (8 1/2 x 11 �C���`)</summary>
		xlPaperLetter = 1
		''' <summary>���^�[ (���^) (8 1/2 x 11 �C���`)</summary>
		xlPaperLetterSmall = 2
		''' <summary>�m�[�g (8 1/2 x 11 �C���`)</summary>
		xlPaperNote = 18 ' (&H12)
		''' <summary>�J�[�g (215 mm x 275 mm)</summary>
		xlPaperQuarto = 15
		''' <summary>�X�e�[�g�����g (5 1/2 x 8 1/2 �C���`)</summary>
		xlPaperStatement = 6
		''' <summary>�^�u���C�h (11 x 17 �C���`)</summary>
		xlPaperTabloid = 3
		''' <summary>���[�U�[�ݒ�</summary>
		xlPaperUser = 256 ' (&H100)
	End Enum

	''' <summary>
	''' �V�[�g�ւ̃R�����g�̈�����@���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPrintLocation
		''' <summary>���[�N�V�[�g���̑}���ʒu�ɃR�����g��������܂��B</summary>
		xlPrintInPlace = 16	' (&H10)
		''' <summary>�R�����g��������܂���B</summary>
		xlPrintNoComments = -4142 ' (&HFFFFEFD2)
		''' <summary>���[�N�V�[�g�̍Ō�ɁA�����r���Ƃ��ăR�����g��������܂��B</summary>
		xlPrintSheetEnd = 1
	End Enum

	''' <summary>
	''' �\���������G���[�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPrintErrors
		''' <summary>����G���[�͋󔒂ɂȂ�܂��B</summary>
		xlPrintErrorsBlank = 1
		''' <summary>����G���[�̓_�b�V���ŕ\������܂��B</summary>
		xlPrintErrorsDash = 2
		''' <summary>����G���[�����ׂĕ\�����܂��B</summary>
		xlPrintErrorsDisplayed = 0
		''' <summary>����G���[���g�p�s�Ƃ��ĕ\�����܂��B</summary>
		xlPrintErrorsNA = 3
	End Enum

	''' <summary>
	''' �}�`�ɓK�p����F�̕ϊ����w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoPictureColorType
		''' <summary>����</summary>
		msoPictureAutomatic = 1
		''' <summary>����</summary>
		msoPictureBlackAndWhite = 3
		''' <summary>�O���[�X�P�[��</summary>
		msoPictureGrayscale = 2
		''' <summary>����</summary>
		msoPictureMixed = -2 ' (&HFFFFFFFE)
		''' <summary>������</summary>
		msoPictureWatermark = 4
	End Enum

	''' <summary>
	''' 3 �X�e�[�g�̃u�[���^ (Boolean) �̒l���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoTriState
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		msoCTrue = 1
		''' <summary>�U (False)</summary>
		msoFalse = 0
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		msoTriStateMixed = -2 ' (&HFFFFFFFE)
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		msoTriStateToggle = -3 ' (&HFFFFFFFD)
		''' <summary>�^ (True)</summary>
		msoTrue = -1 ' (&HFFFFFFFF)
	End Enum

	''' <summary>
	''' �}���̃e�L�X�g�ɓK�p����w�i�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlBackground
		''' <summary>Excel �Ŕw�i�𐧌䂵�܂��B</summary>
		xlBackgroundAutomatic = -4105 ' (&HFFFFEFF7)
		''' <summary>�s����</summary>
		xlBackgroundOpaque = 3
		''' <summary>����</summary>
		xlBackgroundTransparent = 2
	End Enum

	''' <summary>
	''' �t�H���g�ɓK�p����鉺���̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlUnderlineStyle
		''' <summary>������d����</summary>
		xlUnderlineStyleDouble = -4119 ' (&HFFFFEFE9)
		''' <summary>�݂��ɋߐڂ���ׂ���d����</summary>
		xlUnderlineStyleDoubleAccounting = 5
		''' <summary>�����Ȃ�</summary>
		xlUnderlineStyleNone = -4142 ' (&HFFFFEFD2)
		''' <summary>��d����</summary>
		xlUnderlineStyleSingle = 2
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		xlUnderlineStyleSingleAccounting = 4
	End Enum

	''' <summary>
	''' AutoShape �I�u�W�F�N�g�̐}�`�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoAutoShapeType
		''' <summary>�� 16</summary>
		msoShape16pointStar = 94
		''' <summary>�� 24</summary>
		msoShape24pointStar = 95
		''' <summary>�� 32</summary>
		msoShape32pointStar = 96
		''' <summary>�� 4</summary>
		msoShape4pointStar = 91
		''' <summary>�� 5</summary>
		msoShape5pointStar = 92
		''' <summary>�� 8</summary>
		msoShape8pointStar = 93
		''' <summary>[�߂�] �{�^���܂��� [�O��] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonBackorPrevious = 129
		''' <summary>[�擪] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonBeginning = 131
		''' <summary>����̉摜��e�L�X�g�̂Ȃ��{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonCustom = 125
		''' <summary>[�h�L�������g] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonDocument = 134
		''' <summary>[�s��] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonEnd = 132
		''' <summary>[�i��] �{�^���܂��� [����] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonForwardorNext = 130
		''' <summary>[�w���v] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonHelp = 127
		''' <summary>[�z�[��] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonHome = 126
		''' <summary>[���] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonInformation = 128
		''' <summary>[�r�f�I] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonMovie = 136
		''' <summary>[�߂�] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonReturn = 133
		''' <summary>[�T�E���h] �{�^���B�}�E�X�N���b�N�ƃ}�E�X�I�[�o�[�̑�����T�|�[�g���܂��B</summary>
		msoShapeActionButtonSound = 135
		''' <summary>�~��</summary>
		msoShapeArc = 25
		''' <summary>�o���[��</summary>
		msoShapeBalloon = 137
		''' <summary>�Ȑܖ��</summary>
		msoShapeBentArrow = 41
		''' <summary>���ܖ�� (����ł͏���)</summary>
		msoShapeBentUpArrow = 44
		''' <summary>�z��</summary>
		msoShapeBevel = 15
		''' <summary>�A�[�`</summary>
		msoShapeBlockArc = 20
		''' <summary>�~��</summary>
		msoShapeCan = 13
		''' <summary>�R�`</summary>
		msoShapeChevron = 52
		''' <summary>����</summary>
		msoShapeCircularArrow = 60
		''' <summary>�_�`�����o��</summary>
		msoShapeCloudCallout = 108
		''' <summary>�\���`</summary>
		msoShapeCross = 11
		''' <summary>������</summary>
		msoShapeCube = 14
		''' <summary>���J�[�u���</summary>
		msoShapeCurvedDownArrow = 48
		''' <summary>���J�[�u ���{��</summary>
		msoShapeCurvedDownRibbon = 100
		''' <summary>���J�[�u���</summary>
		msoShapeCurvedLeftArrow = 46
		''' <summary>�E�J�[�u���</summary>
		msoShapeCurvedRightArrow = 45
		''' <summary>��J�[�u���</summary>
		msoShapeCurvedUpArrow = 47
		''' <summary>��J�[�u ���{��</summary>
		msoShapeCurvedUpRibbon = 99
		''' <summary>�Ђ��`</summary>
		msoShapeDiamond = 4
		''' <summary>�h�[�i�c</summary>
		msoShapeDonut = 18
		''' <summary>��������</summary>
		msoShapeDoubleBrace = 27
		''' <summary>�傩����</summary>
		msoShapeDoubleBracket = 26
		''' <summary>���g</summary>
		msoShapeDoubleWave = 104
		''' <summary>�����</summary>
		msoShapeDownArrow = 36
		''' <summary>����󐁂��o��</summary>
		msoShapeDownArrowCallout = 56
		''' <summary>�����{��</summary>
		msoShapeDownRibbon = 98
		''' <summary>���� 1</summary>
		msoShapeExplosion1 = 89
		''' <summary>���� 2</summary>
		msoShapeExplosion2 = 90
		''' <summary>�t���[�`���[�g : ��֏���</summary>
		msoShapeFlowchartAlternateProcess = 62
		''' <summary>�t���[�`���[�g : �J�[�h</summary>
		msoShapeFlowchartCard = 75
		''' <summary>�t���[�`���[�g : �ƍ�</summary>
		msoShapeFlowchartCollate = 79
		''' <summary>�t���[�`���[�g : �����q</summary>
		msoShapeFlowchartConnector = 73
		''' <summary>�t���[�`���[�g : �f�[�^</summary>
		msoShapeFlowchartData = 64
		''' <summary>�t���[�`���[�g : ���f</summary>
		msoShapeFlowchartDecision = 63
		''' <summary>�t���[�`���[�g : �_���σQ�[�g</summary>
		msoShapeFlowchartDelay = 84
		''' <summary>�t���[�`���[�g : ���ڃA�N�Z�X�L��</summary>
		msoShapeFlowchartDirectAccessStorage = 87
		''' <summary>�t���[�`���[�g : �\��</summary>
		msoShapeFlowchartDisplay = 88
		''' <summary>�t���[�`���[�g : ����</summary>
		msoShapeFlowchartDocument = 67
		''' <summary>�t���[�`���[�g : ���o��</summary>
		msoShapeFlowchartExtract = 81
		''' <summary>�t���[�`���[�g : �����L��</summary>
		msoShapeFlowchartInternalStorage = 66
		''' <summary>�t���[�`���[�g : ���C�f�B�X�N</summary>
		msoShapeFlowchartMagneticDisk = 86
		''' <summary>�t���[�`���[�g : �葀�����</summary>
		msoShapeFlowchartManualInput = 71
		''' <summary>�t���[�`���[�g : ����</summary>
		msoShapeFlowchartManualOperation = 72
		''' <summary>�t���[�`���[�g : �g����</summary>
		msoShapeFlowchartMerge = 82
		''' <summary>�t���[�`���[�g : ��������</summary>
		msoShapeFlowchartMultidocument = 68
		''' <summary>�t���[�`���[�g : ���y�[�W�����q</summary>
		msoShapeFlowchartOffpageConnector = 74
		''' <summary>�t���[�`���[�g : �_���a</summary>
		msoShapeFlowchartOr = 78
		''' <summary>�t���[�`���[�g : ��`�ςݏ���</summary>
		msoShapeFlowchartPredefinedProcess = 65
		''' <summary>�t���[�`���[�g : ����</summary>
		msoShapeFlowchartPreparation = 70
		''' <summary>�t���[�`���[�g : ����</summary>
		msoShapeFlowchartProcess = 61
		''' <summary>�t���[�`���[�g : ����E�e�[�v</summary>
		msoShapeFlowchartPunchedTape = 76
		''' <summary>�t���[�`���[�g : �����A�N�Z�X�L��</summary>
		msoShapeFlowchartSequentialAccessStorage = 85
		''' <summary>�t���[�`���[�g : ����</summary>
		msoShapeFlowchartSort = 80
		''' <summary>�t���[�`���[�g : �L���f�[�^</summary>
		msoShapeFlowchartStoredData = 83
		''' <summary>�t���[�`���[�g : �a�ڍ�</summary>
		msoShapeFlowchartSummingJunction = 77
		''' <summary>�t���[�`���[�g : �[�q</summary>
		msoShapeFlowchartTerminator = 69
		''' <summary>����</summary>
		msoShapeFoldedCorner = 16
		''' <summary>�n�[�g</summary>
		msoShapeHeart = 21
		''' <summary>�Z�p�`</summary>
		msoShapeHexagon = 10
		''' <summary>������</summary>
		msoShapeHorizontalScroll = 102
		''' <summary>�񓙕ӎO�p�`</summary>
		msoShapeIsoscelesTriangle = 7
		''' <summary>�����</summary>
		msoShapeLeftArrow = 34
		''' <summary>����󐁂��o��</summary>
		msoShapeLeftArrowCallout = 54
		''' <summary>����������</summary>
		msoShapeLeftBrace = 31
		''' <summary>���傩����</summary>
		msoShapeLeftBracket = 29
		''' <summary>���E���</summary>
		msoShapeLeftRightArrow = 37
		''' <summary>���E��󐁂��o��</summary>
		msoShapeLeftRightArrowCallout = 57
		''' <summary>�O�������</summary>
		msoShapeLeftRightUpArrow = 40
		''' <summary>��������</summary>
		msoShapeLeftUpArrow = 43
		''' <summary>���</summary>
		msoShapeLightningBolt = 22
		''' <summary>�������o�� 1 (�g�t��)</summary>
		msoShapeLineCallout1 = 109
		''' <summary>�����������o�� 1</summary>
		msoShapeLineCallout1AccentBar = 113
		''' <summary>�����������o�� 1 (�g�t��)</summary>
		msoShapeLineCallout1BorderandAccentBar = 121
		''' <summary>�������o�� 1</summary>
		msoShapeLineCallout1NoBorder = 117
		''' <summary>�������o�� 2 (�g�t��)</summary>
		msoShapeLineCallout2 = 110
		''' <summary>�����������o�� 2</summary>
		msoShapeLineCallout2AccentBar = 114
		''' <summary>�����������o�� 2 (�g�t��)</summary>
		msoShapeLineCallout2BorderandAccentBar = 122
		''' <summary>�������o�� 2</summary>
		msoShapeLineCallout2NoBorder = 118
		''' <summary>�������o�� 3 (�g�t��)</summary>
		msoShapeLineCallout3 = 111
		''' <summary>�����������o�� 3</summary>
		msoShapeLineCallout3AccentBar = 115
		''' <summary>�����������o�� 3 (�g�t��)</summary>
		msoShapeLineCallout3BorderandAccentBar = 123
		''' <summary>�������o�� 3</summary>
		msoShapeLineCallout3NoBorder = 119
		''' <summary>�������o�� 4 (�g�t��)</summary>
		msoShapeLineCallout4 = 112
		''' <summary>�����������o�� 4</summary>
		msoShapeLineCallout4AccentBar = 116
		''' <summary>�����������o�� 4 (�g�t��)</summary>
		msoShapeLineCallout4BorderandAccentBar = 124
		''' <summary>�������o�� 4</summary>
		msoShapeLineCallout4NoBorder = 120
		''' <summary>�l�̎擾�̂݉\�ł��B���̏�Ԃ̑g�ݍ��킹�������܂��B</summary>
		msoShapeMixed = -2
		''' <summary>��</summary>
		msoShapeMoon = 24
		''' <summary>�֎~</summary>
		msoShapeNoSymbol = 19
		''' <summary>V ���`���</summary>
		msoShapeNotchedRightArrow = 50
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		msoShapeNotPrimitive = 138
		''' <summary>���p�`</summary>
		msoShapeOctagon = 6
		''' <summary>�ȉ~</summary>
		msoShapeOval = 9
		''' <summary>�~�`�����o��</summary>
		msoShapeOvalCallout = 107
		''' <summary>���s�l�p�`</summary>
		msoShapeParallelogram = 2
		''' <summary>�z�[���x�[�X</summary>
		msoShapePentagon = 51
		''' <summary>�u���[�`</summary>
		msoShapePlaque = 28
		''' <summary>�l�������</summary>
		msoShapeQuadArrow = 39
		''' <summary>�l������󐁂��o��</summary>
		msoShapeQuadArrowCallout = 59
		''' <summary>�l�p�`</summary>
		msoShapeRectangle = 1
		''' <summary>�l�p�`�����o��</summary>
		msoShapeRectangularCallout = 105
		''' <summary>�܊p�`</summary>
		msoShapeRegularPentagon = 12
		''' <summary>�E���</summary>
		msoShapeRightArrow = 33
		''' <summary>�E��󐁂��o��</summary>
		msoShapeRightArrowCallout = 53
		''' <summary>�E��������</summary>
		msoShapeRightBrace = 32
		''' <summary>�E�傩����</summary>
		msoShapeRightBracket = 30
		''' <summary>���p�O�p�`</summary>
		msoShapeRightTriangle = 8
		''' <summary>�p�ێl�p�`</summary>
		msoShapeRoundedRectangle = 5
		''' <summary>�p�ێl�p�`�����o��</summary>
		msoShapeRoundedRectangularCallout = 106
		''' <summary>�X�}�C��</summary>
		msoShapeSmileyFace = 17
		''' <summary>�X�g���C�v���</summary>
		msoShapeStripedRightArrow = 49
		''' <summary>���z</summary>
		msoShapeSun = 23
		''' <summary>��`</summary>
		msoShapeTrapezoid = 3
		''' <summary>����</summary>
		msoShapeUpArrow = 35
		''' <summary>���󐁂��o��</summary>
		msoShapeUpArrowCallout = 55
		''' <summary>�㉺���</summary>
		msoShapeUpDownArrow = 38
		''' <summary>�㉺��󐁂��o��</summary>
		msoShapeUpDownArrowCallout = 58
		''' <summary>�ナ�{��</summary>
		msoShapeUpRibbon = 97
		''' <summary>U �^�[�����</summary>
		msoShapeUTurnArrow = 42
		''' <summary>�c����</summary>
		msoShapeVerticalScroll = 101
		''' <summary>��g</summary>
		msoShapeWave = 103
	End Enum

	''' <summary>
	''' XML �f�[�^ �t�@�C�����J�����@���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlXmlLoadOption
		''' <summary>XML �f�[�^ �t�@�C���̓��e�� XML ���X�g�ɔz�u���܂��B</summary>
		xlXmlLoadImportToList = 2
		''' <summary>XML �f�[�^ �t�@�C���̃X�L�[�}�� [XML �f�[�^�\��] ��ƃE�B���h�E�ɕ\�����܂��B</summary>
		xlXmlLoadMapXml = 3
		''' <summary>XML �f�[�^ �t�@�C�����J���܂��B�t�@�C���̓��e�̓t���b�g������܂��B</summary>
		xlXmlLoadOpenXml = 1
		''' <summary>�t�@�C�����J�����@��I������悤���߂郁�b�Z�[�W���\������܂��B	</summary>
		xlXmlLoadPromptUser = 0
	End Enum

	''' <summary>
	''' �S��ʕ\���ŉ��y�[�W���邩�A����̈�݂̂ŉ��y�[�W���邩���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageBreakExtent
		''' <summary>�S���</summary>
		xlPageBreakFull = 1
		''' <summary>����̈�̂�</summary>
		xlPageBreakPartial = 2
	End Enum

	''' <summary>
	''' ���[�N�V�[�g�̉��y�[�W�ʒu���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPageBreak
		''' <summary>�����I�ɉ��y�[�W��ǉ����܂��B</summary>
		xlPageBreakAutomatic = -4105 ' (&HFFFFEFF7)
		''' <summary>�蓮�ŉ��y�[�W��}�����܂��B</summary>
		xlPageBreakManual = -4135 ' (&HFFFFEFD9)
		''' <summary>���[�N�V�[�g�ɉ��y�[�W��}�����܂���B</summary>
		xlPageBreakNone = -4142 ' (&HFFFFEFD2)
	End Enum

	''' <summary>
	''' ���[�N�V�[�g�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlSheetType
		''' <summary>�O���t</summary>
		xlChart = -4109	' (&HFFFFEFF3)
		''' <summary>�_�C�A���O �V�[�g</summary>
		xlDialogSheet = -4116 ' (&HFFFFEFEC)
		''' <summary>Excel 4.0 �C���^�[�i�V���i�� �}�N�� �V�[�g</summary>
		xlExcel4IntlMacroSheet = 4
		''' <summary>Excel 4.0 �}�N�� �V�[�g</summary>
		xlExcel4MacroSheet = 3
		''' <summary>���[�N�V�[�g</summary>
		xlWorksheet = -4167 ' (&HFFFFEFB9)
	End Enum

	''' <summary>
	''' �Z���͈͂��R�s�[������@���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlFillWith
		''' <summary>���e�Ə������R�s�[���܂��B</summary>
		xlFillWithAll = -4104 ' (&HFFFFEFF8)
		''' <summary>���e�݂̂��R�s�[���܂��B</summary>
		xlFillWithContents = 2
		''' <summary>�����݂̂��R�s�[���܂��B</summary>
		xlFillWithFormats = -4122 ' (&HFFFFEFE6)
	End Enum

	''' <summary>
	''' �����o�����̎�ގw�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoCalloutType
		''' <summary>�����܂��͐����̒P��Z�O�����g�̈����o����</summary>
		msoCalloutOne = 1
		''' <summary>���R�ɉ�]����P��Z�O�����g�̈����o����</summary>
		msoCalloutTwo = 2
		''' <summary></summary>
		msoCalloutMixed = -2 ' (&HFFFFFFFE)
		''' <summary>2 �̃Z�O�����g���琬������o����</summary>
		msoCalloutThree = 3
		''' <summary>3 �̃Z�O�����g���琬������o����</summary>
		msoCalloutFour = 4
	End Enum

	''' <summary>
	''' �R�l�N�^�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoConnectorType
		msoConnectorElbow = 2
		msoConnectorTypeMixed = -2 ' (&HFFFFFFFE)
		msoConnectorCurve = 3
		msoConnectorStraight = 1
	End Enum

	''' <summary>
	''' �}�\�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoDiagramType
		''' <summary>�A�������h�[�i�c�^�}�\�����A�v���Z�X��\���}�\�B</summary>
		msoDiagramCycle = 2
		''' <summary>�����^�̐}�\�B</summary>
		msoDiagramMixed = -2 ' (&HFFFFFFFE)
		''' <summary>�K�w�\���̊֌W��\���}�\�B</summary>
		msoDiagramOrgChart = 1
		''' <summary>��b�\���I�Ȋ֌W��\���}�\�B</summary>
		msoDiagramPyramid = 4
		''' <summary>���j�ƂȂ�v�f�Ƃ̊֌W��\���}�\�B</summary>
		msoDiagramRadial = 3
		''' <summary>�S�[���܂ł̃X�e�b�v��\���}�\�B</summary>
		msoDiagramTarget = 6
		''' <summary>�v�f�Ԃŏd�Ȃ荇���̈��\���}�\�B</summary>
		msoDiagramVenn = 5
	End Enum

	''' <summary>
	''' �t�H�[�� �R���g���[���̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlFormControl
		''' <summary>�{�^��</summary>
		xlButtonControl = 0
		''' <summary>�`�F�b�N �{�b�N�X</summary>
		xlCheckBox = 1
		''' <summary>�R���{ �{�b�N�X</summary>
		xlDropDown = 2
		''' <summary>�e�L�X�g �{�b�N�X</summary>
		xlEditBox = 3
		''' <summary>�O���[�v �{�b�N�X</summary>
		xlGroupBox = 4
		''' <summary>���x��</summary>
		xlLabel = 5
		''' <summary>���X�g �{�b�N�X</summary>
		xlListBox = 6
		''' <summary>�I�v�V���� �{�^��</summary>
		xlOptionButton = 7
		''' <summary>�X�N���[�� �o�[</summary>
		xlScrollBar = 8
		''' <summary>�X�s�� �{�^��</summary>
		xlSpinner = 9
	End Enum

	''' <summary>
	''' ������̌������w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoTextOrientation
		''' <summary>�E������B</summary>
		msoTextOrientationDownward = 3
		''' <summary>�������B</summary>
		msoTextOrientationHorizontal = 1
		''' <summary>���A�W�A����̃T�|�[�g�p�̉���������щ�]�B</summary>
		msoTextOrientationHorizontalRotatedFarEast = 6
		''' <summary>�T�|�[�g����Ă��܂���B</summary>
		msoTextOrientationMixed = -2 ' (&HFFFFFFFE)
		''' <summary>�E�オ��B</summary>
		msoTextOrientationUpward = 2
		''' <summary>�c�����B</summary>
		msoTextOrientationVertical = 5
		''' <summary>���A�W�A����̃T�|�[�g�p�̏c�����B</summary>
		msoTextOrientationVerticalFarEast = 4
	End Enum

	''' <summary>
	''' WordArt �I�u�W�F�N�g�Ŏg�p���������ʃe�L�X�g���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoPresetTextEffect
		''' <summary>1 �Ԗڂ̓������</summary>
		msoTextEffect1 = 0
		''' <summary>10 �Ԗڂ̓������</summary>
		msoTextEffect10 = 9
		''' <summary>11 �Ԗڂ̓������</summary>
		msoTextEffect11 = 10
		''' <summary>12 �Ԗڂ̓������</summary>
		msoTextEffect12 = 11
		''' <summary>13 �Ԗڂ̓������</summary>
		msoTextEffect13 = 12
		''' <summary>14 �Ԗڂ̓������</summary>
		msoTextEffect14 = 13
		''' <summary>15 �Ԗڂ̓������</summary>
		msoTextEffect15 = 14
		''' <summary>16 �Ԗڂ̓������</summary>
		msoTextEffect16 = 15
		''' <summary>17 �Ԗڂ̓������</summary>
		msoTextEffect17 = 16
		''' <summary>18 �Ԗڂ̓������</summary>
		msoTextEffect18 = 17
		''' <summary>19 �Ԗڂ̓������</summary>
		msoTextEffect19 = 18
		''' <summary>2 �Ԗڂ̓������</summary>
		msoTextEffect2 = 1
		''' <summary>20 �Ԗڂ̓������</summary>
		msoTextEffect20 = 19
		''' <summary>21 �Ԗڂ̓������</summary>
		msoTextEffect21 = 20
		''' <summary>22 �Ԗڂ̓������</summary>
		msoTextEffect22 = 21
		''' <summary>23 �Ԗڂ̓������</summary>
		msoTextEffect23 = 22
		''' <summary>24 �Ԗڂ̓������</summary>
		msoTextEffect24 = 23
		''' <summary>25 �Ԗڂ̓������</summary>
		msoTextEffect25 = 24
		''' <summary>26 �Ԗڂ̓������</summary>
		msoTextEffect26 = 25
		''' <summary>27 �Ԗڂ̓������</summary>
		msoTextEffect27 = 26
		''' <summary>28 �Ԗڂ̓������</summary>
		msoTextEffect28 = 27
		''' <summary>29 �Ԗڂ̓������</summary>
		msoTextEffect29 = 28
		''' <summary>3 �Ԗڂ̓������</summary>
		msoTextEffect3 = 2
		''' <summary>30 �Ԗڂ̓������</summary>
		msoTextEffect30 = 29
		''' <summary>4 �Ԗڂ̓������</summary>
		msoTextEffect4 = 3
		''' <summary>5 �Ԗڂ̓������</summary>
		msoTextEffect5 = 4
		''' <summary>6 �Ԗڂ̓������</summary>
		msoTextEffect6 = 5
		''' <summary>7 �Ԗڂ̓������</summary>
		msoTextEffect7 = 6
		''' <summary>8 �Ԗڂ̓������</summary>
		msoTextEffect8 = 7
		''' <summary>9 �Ԗڂ̓������</summary>
		msoTextEffect9 = 8
		''' <summary>���g�p</summary>
		msoTextEffectMixed = -2 ' (&HFFFFFFFE)
	End Enum

	''' <summary>
	''' �Z�O�����g�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoSegmentType
		''' <summary>�Ȑ�</summary>
		msoSegmentCurve = 1
		''' <summary>����</summary>
		msoSegmentLine = 0
	End Enum

	''' <summary>
	''' �ߓ_�̕ҏW�̎�ނ��w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum MsoEditingType
		''' <summary>�ҏW�̎�ނ́A�ڑ����Ă���Z�O�����g�̎�ނɑΉ����܂��B</summary>
		msoEditingAuto = 0
		''' <summary>�R�[�i�[�̐ߓ_</summary>
		msoEditingCorner = 1
		''' <summary>�X���[�Y�Ȑߓ_</summary>
		msoEditingSmooth = 2
		''' <summary>�Ώ̓I�Ȑߓ_</summary>
		msoEditingSymmetric = 3
	End Enum

	''' <summary>
	''' �O���t�̓h��Ԃ��̃p�^�[���܂��͓h��Ԃ��̃I�u�W�F�N�g���w�肵�܂��B
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlPattern
		''' <summary>�p�^�[���� Excel �ɂ���Đ��䂳��܂��B</summary>
		xlPatternAutomatic = -4105
		''' <summary>�s���͗l�̃p�^�[���ł��B</summary>
		xlPatternChecker = 9
		''' <summary>�Ԗږ͗l�̃p�^�[���ł��B</summary>
		xlPatternCrissCross = 16
		''' <summary>�E������̔Z���Ίp���̃p�^�[���ł��B</summary>
		xlPatternDown = -4121
		''' <summary>16% �̊D�F�ł��B</summary>
		xlPatternGray16 = 17
		''' <summary>25% �̊D�F�ł��B</summary>
		xlPatternGray25 = -4124
		''' <summary>50% �̊D�F�ł��B</summary>
		xlPatternGray50 = -4125
		''' <summary>75% �̊D�F�ł��B</summary>
		xlPatternGray75 = -4126
		''' <summary>8% �̊D�F�ł��B</summary>
		xlPatternGray8 = 18
		''' <summary>�i�q�͗l�̃p�^�[���ł��B</summary>
		xlPatternGrid = 15
		''' <summary>�Z�������̃p�^�[���ł��B</summary>
		xlPatternHorizontal = -4128
		''' <summary>�E������̔����Ίp���̃p�^�[���ł��B</summary>
		xlPatternLightDown = 13
		''' <summary>���������̃p�^�[���ł��B</summary>
		xlPatternLightHorizontal = 11
		''' <summary>�E�オ��̔����Ίp���̃p�^�[���ł��B</summary>
		xlPatternLightUp = 14
		''' <summary>�����c���̃p�^�[���ł��B</summary>
		xlPatternLightVertical = 12
		''' <summary>�p�^�[���͂���܂���B</summary>
		xlPatternNone = -4142
		''' <summary>75% �̔Z�����A�� �p�^�[���ł��B</summary>
		xlPatternSemiGray75 = 10
		''' <summary>���F�ł��B</summary>
		xlPatternSolid = 1
		''' <summary>�E�オ��̔Z���Ίp���̃p�^�[���ł��B</summary>
		xlPatternUp = -4162
		''' <summary>�Z���c���̃p�^�[���ł��B</summary>
		xlPatternVertical = -4166
	End Enum

	''' <summary>
	''' ���s����ׂ������}�N��
	''' </summary>
	''' <remarks></remarks>
	Public Enum XlRunAutoMacro
		''' <summary>Auto_Activate�}�N��</summary>
		xlAutoActivate = 3
		''' <summary>Auto_Close�}�N��</summary>
		xlAutoClose = 2
		''' <summary>Auto_Deactivate�}�N��</summary>
		xlAutoDeactivate = 4
		''' <summary>Auto_Open�}�N��</summary>
		xlAutoOpen = 1
	End Enum

    ''' <summary>
    ''' �ϊ�����t�@�C���t�H�[�}�b�g
    ''' </summary>
    ''' <remarks>https://msdn.microsoft.com/JA-JP/library/office/ff195006.aspx</remarks>
    Public Enum FixedFormatType As Integer
        ''' <summary>
        ''' PDF�t�@�C��
        ''' </summary>
        ''' <remarks></remarks>
        PDF
        ''' <summary>
        ''' XPS�t�@�C��
        ''' </summary>
        ''' <remarks></remarks>
        XPS
    End Enum

    ''' <summary>
    ''' �ϊ��i��
    ''' </summary>
    ''' <remarks>https://msdn.microsoft.com/ja-jp/library/office/ff838396.aspx</remarks>
    Public Enum FixedFormatQuality As Integer
        ''' <summary>
        ''' �W���i��
        ''' </summary>
        ''' <remarks></remarks>
        QualityStandard
        ''' <summary>
        ''' �ŏ����i��
        ''' </summary>
        ''' <remarks></remarks>
        QualityMinimum
    End Enum

#End Region

End Namespace
