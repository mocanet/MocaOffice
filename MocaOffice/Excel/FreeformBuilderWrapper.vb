
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 作成中のフリーフォームのジオメトリを表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class FreeformBuilderWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Shapes のラッパー</summary>
		Private _shapes As ShapesWrapper

		''' <summary>Excel.FreeformBuilder</summary>
		Private _freeformBuilder As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="shapes">親のShapes</param>
		''' <param name="freeformBuilder">Excel.FreeformBuilder</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal shapes As ShapesWrapper, ByVal freeformBuilder As Object)
			MyBase.New(shapes.ApplicationWrapper)
			_shapes = shapes
			_freeformBuilder = freeformBuilder
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_freeformBuilder)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _freeformBuilder
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 指定したオブジェクトの作成元のアプリケーションを示す 32 ビットの整数値を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Creator() As XlCreator
			Get
				Return DirectCast(InvokeGetProperty(_freeformBuilder, "Creator", Nothing), XlCreator)
			End Get
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' フリーフォームに節点を追加します。
		''' </summary>
		''' <param name="SegmentType">必ず指定します。MsoSegmentType を指定します。追加するセグメントの種類を指定します。</param>
		''' <param name="EditingType">必ず指定します。MsoEditingType を指定します。頂点の編集プロパティを指定します。<br/>引数 SegmentType に msoSegmentLine を指定した場合、引数 EditingType には msoEditingAuto を指定する必要があります。</param>
		''' <param name="X1">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合、この引数は文書の左上隅から新しいセグメントの終点までの水平距離をポイント単位で指定します。<br/>
		''' 新しい節点の引数 EditingType が msoEditingCorner の場合、この引数は文書の左上隅から新しいセグメントの最初のコントロール ポイントまでの水平距離をポイント単位で指定します。</param>
		''' <param name="Y1">必ず指定します。単精度浮動小数点数型 (Single) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合、この引数は文書の左上隅から新しいセグメントの終点までの水平距離をポイント単位で指定します。<br/>
		''' 新しい節点の引数 EditingType が msoEditingCorner の場合、この引数は文書の左上隅から新しいセグメントの最初のコントロール ポイントまでの水平距離をポイント単位で指定します。</param>
		''' <param name="X2">省略可能です。オブジェクト型 (Object) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingCorner の場合、この引数は文書の左上隅から新しいセグメントの 2 番目のコントロール ポイントまでの水平距離をポイント単位で指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合は、この引数に値を指定しないでください。</param>
		''' <param name="Y2">省略可能です。オブジェクト型 (Object) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingCorner の場合、この引数は文書の左上隅から新しいセグメントの 2 番目のコントロール ポイントまでの水平距離をポイント単位で指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合は、この引数に値を指定しないでください。</param>
		''' <param name="X3">省略可能です。オブジェクト型 (Object) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingCorner の場合、この引数は、文書の左上隅から新しいセグメントの終点までの水平距離をポイント単位で指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合は、この引数に値を指定しないでください。</param>
		''' <param name="Y3">省略可能です。オブジェクト型 (Object) の値を指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingCorner の場合、この引数は、文書の左上隅から新しいセグメントの終点までの垂直距離をポイント単位で指定します。<br/>
		''' 新しいセグメントの引数 EditingType が msoEditingAuto の場合は、この引数に値を指定しないでください。</param>
		''' <remarks></remarks>
		Public Sub AddNodes( _
		  <InAttribute()> ByVal SegmentType As MsoSegmentType, _
		  <InAttribute()> ByVal EditingType As MsoEditingType, _
		  <InAttribute()> ByVal X1 As Single, _
		  <InAttribute()> ByVal Y1 As Single, _
		  <InAttribute()> Optional ByVal X2 As Object = Nothing, _
		  <InAttribute()> Optional ByVal Y2 As Object = Nothing, _
		  <InAttribute()> Optional ByVal X3 As Object = Nothing, _
		  <InAttribute()> Optional ByVal Y3 As Object = Nothing _
		 )
			Dim argsV As ArrayList
			Dim argsN As ArrayList

			argsV = New ArrayList()
			argsN = New ArrayList()

			argsV.Add(SegmentType)
			argsN.Add("SegmentType")
			argsV.Add(EditingType)
			argsN.Add("EditingType")
			argsV.Add(X1)
			argsN.Add("X1")
			argsV.Add(Y1)
			argsN.Add("Y1")
			If X2 IsNot Nothing Then
				argsV.Add(X2)
				argsN.Add("x2")
			End If
			If Y2 IsNot Nothing Then
				argsV.Add(Y2)
				argsN.Add("Y2")
			End If
			If X3 IsNot Nothing Then
				argsV.Add(X3)
				argsN.Add("X3")
			End If
			If Y3 IsNot Nothing Then
				argsV.Add(Y3)
				argsN.Add("Y3")
			End If

			InvokeMethod(_freeformBuilder, "AddNodes", argsV.ToArray(), DirectCast(argsN.ToArray(GetType(String)), String()))
		End Sub

		''' <summary>
		''' 指定した FreeformBuilder オブジェクトの幾何学的な特徴を持つ図形を作成します。新しい図形を表す Shape オブジェクトを取得します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function ConvertToShape() As ShapeWrapper
			Dim obj As Object
			Dim wrapper As ShapeWrapper
			obj = InvokeMethod(_freeformBuilder, "ConvertToShape", Nothing)
			wrapper = New ShapeWrapper(_shapes, obj)
			addXlsObject(wrapper)
			Return wrapper
		End Function

#End Region

	End Class

End Namespace
