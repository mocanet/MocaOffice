
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 
	''' </summary>
	''' <remarks></remarks>
	Public Class ShapeRangeWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Shapes のラッパー</summary>
		Private _shapes As ShapesWrapper

		''' <summary>Excel.ShapeRange</summary>
		Private _shapeRange As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="shapes">親のShapes</param>
		''' <param name="shapeRange">Excel.ShapeRange</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal shapes As ShapesWrapper, ByVal shapeRange As Object)
			MyBase.New(shapes.ApplicationWrapper)
			_shapes = shapes
			_shapeRange = shapeRange
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_shapeRange)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _shapeRange
			End Get
		End Property

#End Region
#Region " プロパティ "

#End Region

	End Class

End Namespace
