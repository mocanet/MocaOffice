
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 指定したオートシェイプ、ワードアート オブジェクト、またはコネクタの調整値のコレクションが含まれます。
	''' </summary>
	''' <remarks></remarks>
	Public Class AdjustmentsWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Shape のラッパー</summary>
		Private _shape As ShapeWrapper

		''' <summary>Excel.Adjustments</summary>
		Private _adjustments As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="shape">親のShape</param>
		''' <param name="adjustments">Excel.Adjustments</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal shape As ShapeWrapper, ByVal adjustments As Object)
			MyBase.New(shape.ApplicationWrapper)
			_shape = shape
			_adjustments = adjustments
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_adjustments)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _adjustments
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
				Return DirectCast(InvokeGetProperty(_adjustments, "Count", Nothing), Integer)
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
				Return DirectCast(InvokeGetProperty(_adjustments, "Creator", Nothing), XlCreator)
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
		 ) As Single
			Get
				Return DirectCast(InvokeGetProperty(_adjustments, "Item", New Object() {Index}), Single)
			End Get
			Set(ByVal value As Single)
				InvokeSetProperty(_adjustments, "Item", New Object() {value})
			End Set
		End Property

#End Region

	End Class

End Namespace
