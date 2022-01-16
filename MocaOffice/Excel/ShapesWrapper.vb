
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 指定したシートのすべての Shape オブジェクトのコレクションです。Shape オブジェクトは、オートシェイプ、フリーフォーム、OLE オブジェクト、図などの描画レイヤのオブジェクトを表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class ShapesWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Sheet のラッパー</summary>
		Private _sheet As SheetWrapper

		''' <summary>Excel.Shapes</summary>
		Private _shapes As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="sheet">親のSheet</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal sheet As SheetWrapper, ByVal shapes As Object)
			MyBase.New(sheet.ApplicationWrapper)
			_sheet = sheet
			_shapes = shapes
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_shapes)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _shapes
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' コレクション内のオブジェクトの数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_shapes, "Count", Nothing), Integer)
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
				Return DirectCast(InvokeGetProperty(_shapes, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' Shapes コレクションに含まれる図形のサブセットを表す ShapeRange オブジェクトを取得します。
		''' </summary>
		''' <param name="Index"></param>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Range(<InAttribute()> ByVal Index As Object) As ShapeRangeWrapper
			Get
				Dim obj As Object
				Dim wrap As ShapeRangeWrapper

				obj = InvokeGetProperty(_shapes, "Range", New Object() {Index})

				wrap = New ShapeRangeWrapper(Me, obj)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' 輪郭なしの線吹き出しを作成します。新しい吹き出しを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Type">必ず指定します。<see cref="MsoCalloutType" /> を指定します。引き出し線の種類を指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、吹き出しのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、吹き出しのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。吹き出しのテキスト ボックスの幅をポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。吹き出しのテキスト ボックスの高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddCallout( _
		  <InAttribute()> ByVal Type As MsoCalloutType, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddCallout", New Object() {Type, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' コネクタを作成します。作成したコネクタを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Type">必ず指定します。<see cref="MsoConnectorType"/> を指定します。追加するコネクタの種類を指定します。</param>
		''' <param name="BeginX">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、コネクタの始点の水平位置をポイント単位で指定します。</param>
		''' <param name="BeginY">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、コネクタの始点の垂直位置をポイント単位で指定します。</param>
		''' <param name="EndX">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、コネクタの終点の水平位置をポイント単位で指定します。</param>
		''' <param name="EndY">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、コネクタの終点の垂直位置をポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddConnector( _
		  <InAttribute()> ByVal Type As MsoConnectorType, _
		  <InAttribute()> ByVal BeginX As Single, _
		  <InAttribute()> ByVal BeginY As Single, _
		  <InAttribute()> ByVal EndX As Single, _
		  <InAttribute()> ByVal EndY As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddConnector", New Object() {Type, BeginX, BeginY, EndX, EndY})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' ワークシートのベジエ曲線を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="SafeArrayOfPoints">必ず指定します。オブジェクト型 (Object) の値を指定します。曲線の両端とコントロール ポイントを指定する座標値の配列を指定します。最初に指定した点が始点となり、次に指定した 2 つの点が最初のベジエ セグメントのコントロール ポイントとなります。次に、曲線に追加されたセグメントごとに、1 つの中継点と 2 つのコントロール ポイントを指定します。最後に指定した点が、曲線の終点となります。必ず 3n + 1 個のコントロール ポイントを指定する必要があります。"n" は、曲線を構成するセグメントの数です。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddCurve( _
		  <InAttribute()> ByVal SafeArrayOfPoints As Object _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddCurve", New Object() {SafeArrayOfPoints})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' 図表を作成します。作成した図表を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Type">必ず指定します。<see cref="MsoDiagramType"/> を指定します。図表の種類を指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。ワークシートの左上隅を基準に、図表の左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。ワークシートの左上隅を基準に、図表の左上端の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。図表の幅をポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。図表の高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddDiagram( _
		  <InAttribute()> ByVal Type As MsoDiagramType, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddDiagram", New Object() {Type, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' Microsoft Excel コントロールを作成します。作成したコントロールを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Type">必ず指定します。XlFormControl を指定します。Microsoft Excel コントロールの種類を指定します。ワークシートに編集ボックスは作成できません。</param>
		''' <param name="Left">必ず指定します。整数型 (Integer) の値を指定します。ワークシートのセル A1 の左上隅またはグラフの左上隅を基準に、新しいオブジェクトの初期座標をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。整数型 (Integer) の値を指定します。ワークシートのセル A1 の左上隅またはグラフの左上隅を基準に、新しいオブジェクトの初期座標をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。整数型 (Integer) の値を指定します。新しいオブジェクトの初期サイズをポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。整数型 (Integer) の値を指定します。新しいオブジェクトの初期サイズをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddFormControl( _
		  <InAttribute()> ByVal Type As XlFormControl, _
		  <InAttribute()> ByVal Left As Integer, _
		  <InAttribute()> ByVal Top As Integer, _
		  <InAttribute()> ByVal Width As Integer, _
		  <InAttribute()> ByVal Height As Integer _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddFormControl", New Object() {Type, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' ラベルを作成します。作成したラベルを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Orientation">必ず指定します。MsoTextOrientation を指定します。ラベル内の文字列の向きを指定します。<br/>選択またはインストールされている言語の設定 (たとえば、日本語) によっては、これらの定数の一部を使用できない場合があります。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、ラベルの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の上端を基準に、ラベルの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。ラベルの幅をポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。ラベルの高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddLabel( _
		  <InAttribute()> ByVal Orientation As MsoTextOrientation, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddLabel", New Object() {Orientation, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' ワークシートの新しい線を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="BeginX">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に線の始点の位置をポイント単位で指定します。</param>
		''' <param name="BeginY">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に線の始点の位置をポイント単位で指定します。</param>
		''' <param name="EndX">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に線の終点の位置をポイント単位で指定します。</param>
		''' <param name="EndY">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に線の終点の位置をポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddLine( _
		  <InAttribute()> ByVal BeginX As Single, _
		  <InAttribute()> ByVal BeginY As Single, _
		  <InAttribute()> ByVal EndX As Single, _
		  <InAttribute()> ByVal EndY As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddLine", New Object() {BeginX, BeginY, EndX, EndY})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' OLE オブジェクトを作成します。作成した OLE オブジェクトを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="ClassType">省略可能です。オブジェクト型 (Object) の値を指定します。引数 ClassType または FileName のいずれかを指定する必要があります。オブジェクトを作成するためのプログラム ID を含む文字列を指定します。引数 ClassType を指定した場合、引数 FileName と Link は無視されます。</param>
		''' <param name="Filename">省略可能です。オブジェクト型 (Object) の値を指定します。オブジェクトの作成元となるファイルを指定します。パスを指定しない場合は、現在作業中のフォルダが使用されます。オブジェクトに対して引数 ClassType または FileName のいずれかを指定する必要があります。両方を指定することはできません。</param>
		''' <param name="Link">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、OLE オブジェクトと作成元のファイルの間にリンクが設定されます。False を設定すると、OLE オブジェクトは、ファイルの独立したコピーとなります。ClassType に値を指定した場合、この引数に False を設定する必要があります。既定値は False です。</param>
		''' <param name="DisplayAsIcon">省略可能です。オブジェクト型 (Object) の値を指定します。True を設定すると、OLE オブジェクトをアイコンで表示します。既定値は False です。</param>
		''' <param name="IconFileName">省略可能です。オブジェクト型 (Object) の値を指定します。表示するアイコンが含まれているファイルを指定します。</param>
		''' <param name="IconIndex">省略可能です。オブジェクト型 (Object) の値を指定します。引数 IconFileName で指定したファイル内でのアイコンのインデックスを指定します。指定したファイル内のアイコンの順序は、[アイコンの変更] ダイアログ ボックス ([オブジェクトの挿入] ダイアログ ボックスで [アイコンで表示] チェック ボックスをオンにしたときにアクセス可能) に表示されるアイコンの順序と同じです。ファイル内の最初のアイコンのインデックス番号は 0 (ゼロ) です。指定したインデックス番号が引数 IconFileName のファイルに存在しない場合、インデックス番号 1 (ファイル内の 2 番目のアイコン) のアイコンが使用されます。既定値は 0 (ゼロ) です。</param>
		''' <param name="IconLabel">省略可能です。オブジェクト型 (Object) の値を指定します。アイコンの下に表示するラベル (標題) を指定します。</param>
		''' <param name="Left">省略可能です。オブジェクト型 (Object) の値を指定します。文書の左上隅を基準に、新しいオブジェクトの左上隅の位置をポイント単位で指定します。既定値は 0 (ゼロ) です。</param>
		''' <param name="Top">省略可能です。オブジェクト型 (Object) の値を指定します。文書の左上隅を基準に、新しいオブジェクトの左上隅の位置をポイント単位で指定します。既定値は 0 (ゼロ) です。</param>
		''' <param name="Width">省略可能です。オブジェクト型 (Object) の値を指定します。OLE オブジェクトの初期サイズをポイント単位で指定します。</param>
		''' <param name="Height">省略可能です。オブジェクト型 (Object) の値を指定します。OLE オブジェクトの初期サイズをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddOLEObject( _
		  <InAttribute()> Optional ByVal ClassType As Object = Nothing, _
		  <InAttribute()> Optional ByVal Filename As Object = Nothing, _
		  <InAttribute()> Optional ByVal Link As Object = Nothing, _
		  <InAttribute()> Optional ByVal DisplayAsIcon As Object = Nothing, _
		  <InAttribute()> Optional ByVal IconFileName As Object = Nothing, _
		  <InAttribute()> Optional ByVal IconIndex As Object = Nothing, _
		  <InAttribute()> Optional ByVal IconLabel As Object = Nothing, _
		  <InAttribute()> Optional ByVal Left As Object = Nothing, _
		  <InAttribute()> Optional ByVal Top As Object = Nothing, _
		  <InAttribute()> Optional ByVal Width As Object = Nothing, _
		  <InAttribute()> Optional ByVal Height As Object = Nothing _
		 ) As ShapeWrapper
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			If ClassType IsNot Nothing Then
				argsV.Add(ClassType)
				argsN.Add("ClassType")
			End If

			If Filename IsNot Nothing Then
				argsV.Add(Filename)
				argsN.Add("Filename")
			End If
			If Link IsNot Nothing Then
				argsV.Add(Link)
				argsN.Add("ClLink")
			End If
			If DisplayAsIcon IsNot Nothing Then
				argsV.Add(DisplayAsIcon)
				argsN.Add("DisplayAsIcon")
			End If
			If IconFileName IsNot Nothing Then
				argsV.Add(IconFileName)
				argsN.Add("IconFileName")
			End If
			If IconIndex IsNot Nothing Then
				argsV.Add(IconIndex)
				argsN.Add("IconIndex")
			End If
			If IconLabel IsNot Nothing Then
				argsV.Add(IconLabel)
				argsN.Add("IconLabel")
			End If
			If Left IsNot Nothing Then
				argsV.Add(Left)
				argsN.Add("Left")
			End If
			If Top IsNot Nothing Then
				argsV.Add(Top)
				argsN.Add("Top")
			End If
			If Width IsNot Nothing Then
				argsV.Add(Width)
				argsN.Add("Width")
			End If
			If Height IsNot Nothing Then
				argsV.Add(Height)
				argsN.Add("Height")
			End If

			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddOLEObject", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' 既存のグラフィック ファイルから図を作成します。作成した図を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Filename">必ず指定します。文字列型 (String) の値を指定します。OLE オブジェクトの作成元となるグラフィック ファイルを指定します。</param>
		''' <param name="LinkToFile">必ず指定します MsoTriState を指定します。リンク先のファイルを指定します。</param>
		''' <param name="SaveWithDocument">必ず指定します。MsoTriState 列挙型を指定します。文書と共に図を保存するかどうかを指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を使用します。文書の左上隅を基準に、図の左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を使用します。文書の上端を基準に、図の左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。図の幅をポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。図の高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddPicture( _
		  <InAttribute()> ByVal Filename As String, _
		  <InAttribute()> ByVal LinkToFile As MsoTriState, _
		  <InAttribute()> ByVal SaveWithDocument As MsoTriState, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddPicture", New Object() {Filename, LinkToFile, SaveWithDocument, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' 開いた曲線または閉じた多角形の描画オブジェクトを作成します。作成した曲線または多角形を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="SafeArrayOfPoints">必ず指定します。オブジェクト型 (Object) の値を指定します。曲線の描画オブジェクトの頂点を指定する座標値の配列を指定します。</param>
		''' <returns></returns>
		''' <remarks>
		''' 閉じた多角形を作成するには、曲線の始点と終点に同じ座標値を割り当てます。
		''' </remarks>
		Public Function AddPolyline( _
		  <InAttribute()> ByVal SafeArrayOfPoints As Object _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddPolyline", New Object() {SafeArrayOfPoints})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' ワークシートの新しいオートシェイプを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Type">必ず指定します。Microsoft.Office.Core.MsoAutoShapeType を指定します。作成するオートシェイプの種類を指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準にオートシェイプのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準にオートシェイプのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単��度浮動小数点数型 (Single) の値を指定します。オートシェイプのテキスト ボックスの幅と高さをポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。オートシェイプのテキスト ボックスの幅と高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddShape( _
		  <InAttribute()> ByVal Type As MsoAutoShapeType, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddShape", New Object() {Type, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' テキスト ボックスを作成します。作成したテキスト ボックスを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="Orientation">必ず指定します。MsoTextOrientation を指定します。テキスト ボックスの向きを指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、テキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を使用します。文書の上端を基準に、テキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Width">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。テキスト ボックスの幅をポイント単位で指定します。</param>
		''' <param name="Height">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。テキスト ボックスの高さをポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddTextbox( _
		  <InAttribute()> ByVal Orientation As MsoTextOrientation, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single, _
		  <InAttribute()> ByVal Width As Single, _
		  <InAttribute()> ByVal Height As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddTextbox", New Object() {Orientation, Left, Top, Width, Height})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' ワードアート オブジェクトを作成します。作成したワードアート オブジェクトを表す Shape オブジェクトを取得します。
		''' </summary>
		''' <param name="PresetTextEffect">必ず指定します。 MsoPresetTextEffect を指定します。既定のワードアート スタイルを指定します。</param>
		''' <param name="Text">必ず指定します。文字列型 (String) の値を指定します。ワードアートの文字列を指定します。</param>
		''' <param name="FontName">必ず指定します。文字列型 (String) の値を指定します。ワードアートで使用するフォント名を指定します。</param>
		''' <param name="FontSize">必ず指定します。単精度浮動小数点数型 (Single) の値を使用します。ワードアートで使用するフォント サイズをポイント単位で指定します。</param>
		''' <param name="FontBold">必ず指定します。 MsoTriState を指定します。ワードアートで使用するフォント スタイルを太字にするかどうかを指定します。</param>
		''' <param name="FontItalic">必ず指定します。MsoTriState 列挙型を指定します。ワードアートで使用するフォント スタイルを斜体にするかどうかを指定します。</param>
		''' <param name="Left">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、ワードアートのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <param name="Top">必ず指定します。単精度浮動小数点数型 (Single) の値を使用します。文書の上端を基準に、ワードアートのテキスト ボックスの左上隅の位置をポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function AddTextEffect( _
		  <InAttribute()> ByVal PresetTextEffect As MsoPresetTextEffect, _
		  <InAttribute()> ByVal Text As String, _
		  <InAttribute()> ByVal FontName As String, _
		  <InAttribute()> ByVal FontSize As Single, _
		  <InAttribute()> ByVal FontBold As MsoTriState, _
		  <InAttribute()> ByVal FontItalic As MsoTriState, _
		  <InAttribute()> ByVal Left As Single, _
		  <InAttribute()> ByVal Top As Single _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "AddTextEffect", New Object() {PresetTextEffect, Text, FontName, FontSize, FontBold, FontItalic, Left, Top})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' フリーフォームを作成します。作成したフリーフォームを表す FreeformBuilder オブジェクトを取得します。
		''' </summary>
		''' <param name="EditingType">必ず指定します。MsoEditingType を指定します。最初の節点の編集プロパティを指定します。</param>
		''' <param name="X1">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準に、フリーフォームの最初の節点の位置をポイント単位で指定します。</param>
		''' <param name="Y1">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。文書の左上隅を基準にフリーフォームの最初の節点の位置をポイント単位で指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function BuildFreeform( _
		  <InAttribute()> ByVal EditingType As MsoEditingType, _
		  <InAttribute()> ByVal X1 As Single, _
		  <InAttribute()> ByVal Y1 As Single _
		 ) As FreeformBuilderWrapper
			Dim obj As Object
			Dim wrapper As FreeformBuilderWrapper
			obj = InvokeMethod(_shapes, "BuildFreeform", New Object() {EditingType, X1, Y1})
			wrapper = New FreeformBuilderWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' コレクション全体での繰り返しをサポートするために、列挙型の値を取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function GetEnumerator() As IEnumerator
			Dim bookEnum As IEnumerator
			Dim result As IList(Of ShapeWrapper)

			result = New List(Of ShapeWrapper)

			bookEnum = DirectCast(InvokeMethod(_shapes, "GetEnumerator", Nothing), IEnumerator)
			While bookEnum.MoveNext()
				Dim wrapper As ShapeWrapper
				wrapper = New ShapeWrapper(Me, bookEnum.Current())
				result.Add(wrapper)
				addXlsObject(wrapper)
			End While

			Return result.GetEnumerator()
		End Function

		''' <summary>
		''' コレクションから単一のオブジェクトを取得します。
		''' </summary>
		''' <param name="Index">必ず指定します。オブジェクト型 (Object) の値を指定します。オブジェクトの名前またはインデックス番号を指定します。</param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Item( _
		  <InAttribute()> ByVal Index As Integer _
		 ) As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_shapes, "Item", New Object() {Index})
			wrapper = New ShapeWrapper(Me, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

		''' <summary>
		''' 指定した Shapes コレクションのすべての図形を選択します。
		''' </summary>
		''' <remarks></remarks>
		Public Sub SelectAll()
			InvokeMethod(_shapes, "SelectAll", Nothing)
		End Sub

#End Region

	End Class

End Namespace
