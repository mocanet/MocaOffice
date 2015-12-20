
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' Excel.Workbooks のラッパークラス
	''' </summary>
	''' <remarks>
	''' Microsoft Excel で現在開いているすべての Workbook オブジェクトのコレクションです。
	''' </remarks>
	Public Class BooksWrapper
		Inherits AbstractExcelWrapper

		''' <summary>Excel.Workbooks インスタンス</summary>
		Private _workbooks As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="xls"></param>
		''' <remarks></remarks>
		Public Sub New(ByVal xls As ExcelWrapper)
			MyBase.New(xls)
			_init(xls)
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_workbooks)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _workbooks
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' Excel.Application のラッパー
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property App() As ExcelWrapper
			Get
				Return DirectCast(MyBase.xlsWrapper, ExcelWrapper)
			End Get
		End Property

		''' <summary>
		''' Excel.Workbooks.Count
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_workbooks, "Count", Nothing), Integer)
			End Get
		End Property

		''' <summary>
		''' 指定したオブジェクトの作成元のアプリケーションを示す 32 ビットの整数値を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_workbooks, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' 指定した調整値を設定します。
		''' </summary>
		''' <param name="Index">必ず指定します。整数型 (Integer) の値を指定します。調整のインデックス番号を指定します。</param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Default Public Property Item( _
		  <InAttribute()> ByVal Index As Integer _
		 ) As BookWrapper
			Get
				Dim obj As Object
				Dim wrapper As BookWrapper
				obj = InvokeGetProperty(_workbooks, "Item", New Object() {Index})
				wrapper = New BookWrapper(Me.App, obj)
				addXlsObject(wrapper)
				Return wrapper
			End Get
			Set(ByVal value As BookWrapper)
				InvokeSetProperty(_workbooks, "Item", New Object() {value})
			End Set
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' 初期化処理
		''' </summary>
		''' <param name="xls">Excel.Applicationラッパー</param>
		''' <remarks></remarks>
		Private Sub _init(ByVal xls As ExcelWrapper)
			' Booksオブジェクトの作成
			_workbooks = InvokeGetProperty(xlsApp, "Workbooks", Nothing)
		End Sub

		''' <summary>
		''' ブックの新規追加
		''' </summary>
		''' <param name="filename">ブックのファイル名</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Add(ByVal filename As String) As BookWrapper
			Dim xlBook As Object
			Dim book As BookWrapper
			xlBook = InvokeMethod(_workbooks, "Add", Nothing)
			book = New BookWrapper(Me.App, xlBook, True)
			book.Filename = filename
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' True を設定すると、サーバーから指定したブックをチェックアウトできます。値の取得および設定が可能です。ブール型 (Boolean) の値を使用します。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を指定します。チェックアウトするファイルの名前を指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function CanCheckOut( _
		 <InAttribute()> ByVal Filename As String _
		) As Boolean
			Return DirectCast(InvokeMethod(_workbooks, "CanCheckOut", New Object() {Filename}), Boolean)
		End Function

		''' <summary>
		''' 指定したブックをサーバーからローカル コンピュータにコピーして、編集できるようにします。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を指定します。チェックアウトするファイルの名前を指定します。</param>
		''' <remarks></remarks>
		Public Sub CheckOut( _
		  <InAttribute()> ByVal Filename As String _
		 )
			InvokeMethod(_workbooks, "CheckOut", New Object() {Filename})
		End Sub

		''' <summary>
		''' 指定したオブジェクトを閉じます。
		''' </summary>
		''' <remarks></remarks>
		Public Sub Close()
			InvokeMethod(_workbooks, "Close", Nothing)
		End Sub

		''' <summary>
		''' コレクション全体での繰り返しをサポートするために、列挙型の値を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator
			Dim bookEnum As IEnumerator
			Dim result As IList(Of BookWrapper)

			result = New List(Of BookWrapper)

			bookEnum = DirectCast(InvokeMethod(_workbooks, "GetEnumerator", Nothing), IEnumerator)
			While bookEnum.MoveNext()
				Dim wrapper As BookWrapper
				wrapper = New BookWrapper(Me.App, bookEnum.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

		''' <summary>
		''' ブックを開きます。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を指定します。開くブックのファイル名を指定します。</param>
		''' <param name="UpdateLinks">省略可能です。オブジェクト型 (Object) の値を指定します。ファイル内のリンクの更新方法を指定します。この引数を省略すると、リンクの更新方法の指定を促すダイアログ ボックスが表示されます。省略しない場合は、次のいずれかの値を指定します。</param>
		''' <param name="ReadOnly">省略可能です。オブジェクト型 (Object) の値を使用します。ブックを読み取り専用モードで開くには、True を指定します。</param>
		''' <param name="Format">省略可能です。オブジェクト型 (Object) の値を使用します。Microsoft Excel がテキスト ファイルを開くときに、この引数に項目の区切り文字を指定します。指定できる区切り文字は次のとおりです。この引数を省略すると、現在指定されている区切り文字が使われます。</param>
		''' <param name="Password">省略可能です。オブジェクト型 (Object) の値を指定します。パスワード保護されたブックを開くために必要なパスワードを指定します。パスワードが必要なときにこの引数を省略すると、パスワードの入力を促すダイアログ ボックスが表示されます。</param>
		''' <param name="WriteResPassword">省略可能です。オブジェクト型 (Object) の値を使用します。書き込み保護されたブックに書き込みをするために必要なパスワードを指定します。パスワードが必要なときにこの引数を省略すると、パスワードの入力を促すダイアログ ボックスが表示されます。</param>
		''' <param name="IgnoreReadOnlyRecommended">省略可能です。オブジェクト型 (Object) の値を指定します。[読み取り専用を推奨する] チェック ボックスをオンにして保存されたブックを開くときでも、読み取り専用を推奨するメッセージを非表示にするには、True を指定します。</param>
		''' <param name="Origin">省略可能です。オブジェクト型 (Object) の値を指定します。指定したファイルがテキスト ファイルのときに、それがどのような形式のテキスト ファイルかを指定します。コード ページと CR/LF を正しく変換するために必要です。使用できる定数は、XlPlatform 列挙型の xlMacintosh、xlWindows、xlMSDOS のいずれかです。この引数を省略すると、現在のオペレーティング システムの形式が使われます。</param>
		''' <param name="Delimiter">省略可能です。オブジェクト型 (Object) の値を指定します。指定したファイルがテキスト ファイルであり、引数 Format に 6 が設定されているときに、区切り記号として使う文字を指定します。たとえば、タブの場合は Chr(9)、コンマの場合は ","、セミコロンの場合は ";" を指定します。任意の文字を指定することもできます。文字列を指定したときは、最初の文字だけが使われます。</param>
		''' <param name="Editable">省略可能です。オブジェクト型 (Object) の値を使用します。指定したファイルが Microsoft Excel 4.0 のアドインの場合、この引数に True を指定すると、アドインをウィンドウとして表示します。この引数に False を指定するか省略すると、アドインは非表示の状態で開かれ、ウィンドウとして表示することはできません。この引数は、Microsoft Excel 5.0 以降のアドインには適用されません。指定したファイルが Excel のテンプレートの場合、True を指定すると、指定されたテンプレートを編集用に開きます。False を指定すると、指定されたテンプレートを基にした、新しいブックを開きます。既定値は False です。</param>
		''' <param name="Notify">省略可能です。オブジェクト型 (Object) の値を使用します。指定したファイルが読み取り/書き込みモードで開けない場合に、ファイルを通知リストに追加するには、True を指定します。ファイルは読み取り専用モードで開かれて通知リストに追加され、ブックを編集できる状態になった時点で、ユーザーにその旨が通知されます。ファイルが開けない場合に、このような通知を行わずにエラーを発生させるには、False を指定するか省略します。</param>
		''' <param name="Converter">省略可能です。オブジェクト型 (Object) の値を使用します。ファイルを開くときに、最初に使うファイル コンバータのインデックス番号を指定します。指定したファイル コンバータでファイルを変換できない場合は、他のすべてのファイル コンバータでの変換が試みられます。指定するインデックス番号は、FileConverters プロパティで取得されるファイル コンバータの行番号です。</param>
		''' <param name="AddToMru">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、最近使用したファイルの一覧にこのブックが追加されます。既定値は False です。</param>
		''' <param name="Local">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、Microsoft Excel で使用されている言語でファイルが保存されます (コントロール パネルの設定を含む)。False (既定値) を設定すると、VBA (Visual Basic for Applications) で使用されている言語でファイルが保存されます (通常はアメリカ英語です。ただし、古い国際版の XL5/95 VBA プロジェクトから Workbooks.Open を実行している場合を除きます)。</param>
		''' <param name="CorruptLoad">省略可能です。オブジェクト型 (Object) の値を使用します。使用できる定数は、xlNormalLoad、xlRepairFile、xlExtractData のいずれかです。この引数を省略したときの既定の動作は、標準の読み込み処理となるのが普通ですが、2 回目以降はセーフ ロードやデータ リカバリとなることがあります。つまり、最初は標準の読み込み処理を試みます。ファイルを開いている途中で処理が停止したときは、次にセーフ ロードを試みます。再び処理が停止したときは、次にデータ リカバリを試みます。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Open(ByVal Filename As String, _
		   Optional ByVal UpdateLinks As Object = Nothing, _
		   Optional ByVal [ReadOnly] As Object = Nothing, _
		   Optional ByVal Format As Object = Nothing, _
		   Optional ByVal Password As Object = Nothing, _
		   Optional ByVal WriteResPassword As Object = Nothing, _
		   Optional ByVal IgnoreReadOnlyRecommended As Object = Nothing, _
		   Optional ByVal Origin As Object = Nothing, _
		   Optional ByVal Delimiter As Object = Nothing, _
		   Optional ByVal Editable As Object = Nothing, _
		   Optional ByVal Notify As Object = Nothing, _
		   Optional ByVal Converter As Object = Nothing, _
		   Optional ByVal AddToMru As Object = Nothing, _
		   Optional ByVal Local As Object = Nothing, _
		   Optional ByVal CorruptLoad As Object = Nothing) As BookWrapper
			Dim xlBook As Object
			Dim book As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If UpdateLinks IsNot Nothing Then
				argsV.Add(UpdateLinks)
				argsN.Add("UpdateLinks")
			End If
			If [ReadOnly] IsNot Nothing Then
				argsV.Add([ReadOnly])
				argsN.Add("ReadOnly")
			End If
			If Format IsNot Nothing Then
				argsV.Add(Format)
				argsN.Add("Format")
			End If
			If Password IsNot Nothing Then
				argsV.Add(Password)
				argsN.Add("Password")
			End If
			If WriteResPassword IsNot Nothing Then
				argsV.Add(WriteResPassword)
				argsN.Add("WriteResPassword")
			End If
			If IgnoreReadOnlyRecommended IsNot Nothing Then
				argsV.Add(IgnoreReadOnlyRecommended)
				argsN.Add("IgnoreReadOnlyRecommended")
			End If
			If Origin IsNot Nothing Then
				argsV.Add(Origin)
				argsN.Add("Origin")
			End If
			If Delimiter IsNot Nothing Then
				argsV.Add(Delimiter)
				argsN.Add("Delimiter")
			End If
			If Editable IsNot Nothing Then
				argsV.Add(Editable)
				argsN.Add("Editable")
			End If
			If Notify IsNot Nothing Then
				argsV.Add(Notify)
				argsN.Add("Notify")
			End If
			If Converter IsNot Nothing Then
				argsV.Add(Converter)
				argsN.Add("Converter")
			End If
			If AddToMru IsNot Nothing Then
				argsV.Add(AddToMru)
				argsN.Add("AddToMru")
			End If
			If Local IsNot Nothing Then
				argsV.Add(Local)
				argsN.Add("Local")
			End If
			If CorruptLoad IsNot Nothing Then
				argsV.Add(CorruptLoad)
				argsN.Add("CorruptLoad")
			End If

			xlBook = InvokeMethod(_workbooks, "Open", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' データベースを表す Workbook オブジェクトを取得します。
		''' </summary>
		''' <param name="Filename"></param>
		''' <param name="CommandText"></param>
		''' <param name="CommandType"></param>
		''' <param name="BackgroundQuery"></param>
		''' <param name="ImportDataAs"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function OpenDatabase( _
		 <InAttribute()> ByVal Filename As String, _
		 <InAttribute()> Optional ByVal CommandText As Object = Nothing, _
		 <InAttribute()> Optional ByVal CommandType As Object = Nothing, _
		 <InAttribute()> Optional ByVal BackgroundQuery As Object = Nothing, _
		 <InAttribute()> Optional ByVal ImportDataAs As Object = Nothing _
		) As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If CommandText IsNot Nothing Then
				argsV.Add(CommandText)
				argsN.Add("CommandText")
			End If
			If CommandType IsNot Nothing Then
				argsV.Add(CommandType)
				argsN.Add("CommandType")
			End If
			If BackgroundQuery IsNot Nothing Then
				argsV.Add(BackgroundQuery)
				argsN.Add("BackgroundQuery")
			End If
			If ImportDataAs IsNot Nothing Then
				argsV.Add(ImportDataAs)
				argsN.Add("ImportDataAs")
			End If

			Dim xlBook As Object
			Dim book As BookWrapper

			xlBook = InvokeMethod(_workbooks, "OpenDatabase", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		''' <summary>
		''' テキスト ファイルを分析して読み込みます。テキスト ファイルを 1 枚のシートとして、それを含む新しいブックを開きます。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を使用します。読み込まれるテキスト ファイルの名前を指定します。</param>
		''' <param name="Origin">省略可能です。オブジェクト型 (Object) の値を指定します。テキスト ファイルが作成された機種を指定します。使用できる定数は、<see cref="XlPlatform"/> 列挙型の xlMacintosh、xlWindows、xlMSDOS のいずれかです。この他に、目的のコード ページのコード ページ番号を表す整数も指定できます。たとえば、"1256" は、ソース テキスト ファイルのエンコードでアラビア語 (Windows) を指定します。この引数を省略すると、Text Import Wizard の [元のファイル] オプションの現在の設定が使用されます。</param>
		''' <param name="StartRow">省略可能です。オブジェクト型 (Object) の値を指定します。取り込み開始行を指定します。最初の行を 1 として数えます。既定値は 1 です。 </param>
		''' <param name="DataType">省略可能です。オブジェクト型 (Object) の値を指定します。ファイルに含まれるデータの形式を指定します。使用できる定数は、<see cref="XlTextParsingType"/> 列挙型の xlDelimited または xlFixedWidth です。この引数を省略すると、ファイルを開いたときにデータの形式が自動的に決定されます。</param>
		''' <param name="TextQualifier">省略可能です。<see cref="XlTextQualifier"/> 列挙型の定数を使用します。文字列の引用符を指定します。使用できる定数は、次に示す XlTextQualifier 列挙型の定数のいずれかです。</param>
		''' <param name="ConsecutiveDelimiter">省略可能です。オブジェクト型 (Object) の値を指定します。連続した区切り文字を 1 文字として扱うときは True を指定します。既定値は False です。</param>
		''' <param name="Tab">省略可能です。オブジェクト型 (Object) の値を指定します。引数 DataType に xlDelimited を指定し、区切り文字にタブを使うときは True を指定します。既定値は False です。</param>
		''' <param name="Semicolon">省略可能です。オブジェクト型 (Object) の値を指定します。引数 DataType に xlDelimited を指定し、区切り文字にセミコロン (;) を使うときは True を指定します。既定値は False です。</param>
		''' <param name="Comma">省略可能です。オブジェクト型 (Object) の値を指定します。引数 DataType に xlDelimited を指定し、区切り文字にコンマ (,) を使うときは True を指定します。既定値は False です。</param>
		''' <param name="Space">省略可能です。オブジェクト型 (Object) の値を指定します。引数 DataType に xlDelimited を指定し、区切り文字にスペースを使うときは True を指定します。既定値は False です。</param>
		''' <param name="Other">省略可能です。オブジェクト型 (Object) の値を指定します。引数 DataType に xlDelimited を指定し、区切り文字に OtherChar で指定した文字を使うときは True を指定します。既定値は False です。</param>
		''' <param name="OtherChar">省略可能です。オブジェクト型 (Object) の値を指定します。引数 Other が True のときは、必ずこの引数に区切り文字を指定します。複数の文字を指定したときは、先頭の文字が区切り文字となり、残りの文字は無視されます。</param>
		''' <param name="FieldInfo">省略可能です。<see cref="XlColumnDataType"/> 列挙型の定数を使用します。各列のデータ形式に関する情報を持つ配列を指定します。データ形式の解釈は、引数 DataType の値によって異なります。データが区切り記号で区切られている場合は、この引数は 2 要素配列の配列で、各 2 要素配列は特定の列の変換オプションを指定します。1 番目の要素には 1 から始まる列番号を指定し、2 番目の要素には列のデータ形式を示す XlColumnDataType 列挙型の定数を指定します。</param>
		''' <param name="TextVisualLayout">省略可能です。オブジェクト型 (Object) の値を指定します。テキストの視覚的な配置を指定します。</param>
		''' <param name="DecimalSeparator">省略可能です。オブジェクト型 (Object) の値を指定します。Microsoft Excel で数値を認識する場合に使う小数点の記号です。既定はシステム設定です。 </param>
		''' <param name="ThousandsSeparator">省略可能です。文字列型 (Object) の値を指定します。数字の認識に使用される桁区切り文字を指定します。既定値は、システム設定です。<br/>さまざまなインポート設定でテキストを Excel にインポートする結果を次に示します。数値の結果は右詰めで表示します。</param>
		''' <param name="TrailingMinusNumbers">省略可能です。</param>
		''' <param name="Local">省略可能です。</param>
		''' <remarks></remarks>
		Public Sub OpenText( _
		 ByVal Filename As String, _
		 Optional ByVal Origin As Object = Nothing, _
		 Optional ByVal StartRow As Object = Nothing, _
		 Optional ByVal DataType As Object = Nothing, _
		 Optional ByVal TextQualifier As XlTextQualifier = XlTextQualifier.xlTextQualifierDoubleQuote, _
		 Optional ByVal ConsecutiveDelimiter As Object = Nothing, _
		 Optional ByVal Tab As Object = Nothing, _
		 Optional ByVal Semicolon As Object = Nothing, _
		 Optional ByVal Comma As Object = Nothing, _
		 Optional ByVal Space As Object = Nothing, _
		 Optional ByVal Other As Object = Nothing, _
		 Optional ByVal OtherChar As Object = Nothing, _
		 Optional ByVal FieldInfo As Object = Nothing, _
		 Optional ByVal TextVisualLayout As Object = Nothing, _
		 Optional ByVal DecimalSeparator As Object = Nothing, _
		 Optional ByVal ThousandsSeparator As Object = Nothing, _
		 Optional ByVal TrailingMinusNumbers As Object = Nothing, _
		 Optional ByVal Local As Object = Nothing)
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If Origin IsNot Nothing Then
				argsV.Add(Origin)
				argsN.Add("Origin")
			End If
			If StartRow IsNot Nothing Then
				argsV.Add(StartRow)
				argsN.Add("StartRow")
			End If
			If DataType IsNot Nothing Then
				argsV.Add(DataType)
				argsN.Add("DataType")
			End If
			argsV.Add(TextQualifier)
			argsN.Add("TextQualifier")
			If ConsecutiveDelimiter IsNot Nothing Then
				argsV.Add(ConsecutiveDelimiter)
				argsN.Add("ConsecutiveDelimiter")
			End If
			If Tab IsNot Nothing Then
				argsV.Add(Tab)
				argsN.Add("Tab")
			End If
			If Semicolon IsNot Nothing Then
				argsV.Add(Semicolon)
				argsN.Add("Semicolon")
			End If
			If Comma IsNot Nothing Then
				argsV.Add(Comma)
				argsN.Add("Comma")
			End If
			If Space IsNot Nothing Then
				argsV.Add(Space)
				argsN.Add("Space")
			End If
			If Other IsNot Nothing Then
				argsV.Add(Other)
				argsN.Add("Other")
			End If
			If OtherChar IsNot Nothing Then
				argsV.Add(OtherChar)
				argsN.Add("OtherChar")
			End If
			If FieldInfo IsNot Nothing Then
				argsV.Add(FieldInfo)
				argsN.Add("FieldInfo")
			End If
			If TextVisualLayout IsNot Nothing Then
				argsV.Add(TextVisualLayout)
				argsN.Add("TextVisualLayout")
			End If
			If DecimalSeparator IsNot Nothing Then
				argsV.Add(DecimalSeparator)
				argsN.Add("DecimalSeparator")
			End If
			If ThousandsSeparator IsNot Nothing Then
				argsV.Add(ThousandsSeparator)
				argsN.Add("ThousandsSeparator")
			End If
			If TrailingMinusNumbers IsNot Nothing Then
				argsV.Add(TrailingMinusNumbers)
				argsN.Add("TrailingMinusNumbers")
			End If
			If Local IsNot Nothing Then
				argsV.Add(Local)
				argsN.Add("Local")
			End If

			InvokeMethod(_workbooks, "OpenText", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' XML データ ファイルを開きます。Workbook オブジェクトを取得します。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を指定します。開くファイル名を指定します。</param>
		''' <param name="Stylesheets">省略可能です。オブジェクト型 (Object) の値を指定します。適用する XSLT (XSL 変換) スタイルシート処理命令を指定する単一の値または値の配列を指定します。</param>
		''' <param name="LoadOption">省略可能です。オブジェクト型 (Object) の値を指定します。Excel が XML データ ファイルを開く方法を指定します。使用できる定数は、次に示す <see cref="XlXmlLoadOption"/> 列挙型の定数のいずれかです。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function OpenXML( _
		  <InAttribute()> ByVal Filename As String, _
		  <InAttribute()> Optional ByVal Stylesheets As Object = Nothing, _
		  <InAttribute()> Optional ByVal LoadOption As Object = Nothing _
		 ) As BookWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(Filename)
			argsN.Add("Filename")
			If Stylesheets IsNot Nothing Then
				argsV.Add(Stylesheets)
				argsN.Add("Stylesheets")
			End If
			If LoadOption IsNot Nothing Then
				argsV.Add(LoadOption)
				argsN.Add("LoadOption")
			End If

			Dim xlBook As Object
			Dim book As BookWrapper

			xlBook = InvokeMethod(_workbooks, "OpenXML", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			book = New BookWrapper(Me.App, xlBook)
			addXlsObject(book)
			Return book
		End Function

		Friend Function GetMyBook(ByVal name As String) As BookWrapper
			For Each item As Object In myXlsObject()
				If Not TypeOf item Is BookWrapper Then
					Continue For
				End If
				Dim book As BookWrapper
				book = DirectCast(item, BookWrapper)
				If book.Name <> name Then
					Continue For
				End If
				Return book
			Next
			Return Nothing
		End Function

#End Region

	End Class

End Namespace
