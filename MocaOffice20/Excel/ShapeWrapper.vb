
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' オートシェイプ、フリーフォーム、OLE オブジェクト、または図など描画レイヤのオブジェクトを表します。Shape オブジェクトは、Shapes コレクションのメンバです。Shapes コレクションは、文書のすべての図形を表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class ShapeWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Shapes のラッパー</summary>
		Private _shapes As ShapesWrapper

		''' <summary>Excel.Shape</summary>
		Private _shape As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="shapes">親のShapes</param>
		''' <param name="shape">Excel.Shape</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal shapes As ShapesWrapper, ByVal shape As Object)
			MyBase.New(shapes.ApplicationWrapper)
			_shapes = shapes
			_shape = shape
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_shape)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _shape
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 親のExcel.Shapes のラッパー
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Shapes() As ShapesWrapper
			Get
				Return _shapes
			End Get
		End Property

		''' <summary>
		''' 指定した図形のすべての調整値を含む Adjustments オブジェクトを取得します。値の取得のみ可能です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Adjustments() As AdjustmentsWrapper
			Get
				Dim obj As Object
				Dim wrap As AdjustmentsWrapper

				obj = InvokeGetProperty(_shape, "Adjustments", Nothing)

				wrap = New AdjustmentsWrapper(Me, obj)
				addXlsObject(wrap)

				Return wrap
			End Get
		End Property

		''' <summary>
		''' オブジェクトを Web ページとして保存する場合、Shape オブジェクトの説明 (代替) 文字列を設定します。値の取得および設定が可能です。文字列型 (String) の値を使用します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property AlternativeText() As String
			Get
				Return DirectCast(InvokeGetProperty(_shape, "AlternativeText", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_shape, "AlternativeText", New Object() {value})
			End Set
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' オブジェクトを選択します。
		''' </summary>
		''' <param name="Replace"></param>
		''' <remarks></remarks>
		Public Sub [Select]( _
		  <InAttribute()> Optional ByVal Replace As Object = Nothing _
		 )
			InvokeMethod(_shape, "Select", New Object() {Replace})
		End Sub

#End Region

	End Class

End Namespace
