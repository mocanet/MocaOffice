
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' 水平方向の改ページを表します。
	''' </summary>
	''' <remarks></remarks>
	Public Class HPageBreakWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.HPageBreaks のラッパー</summary>
		Private _hPageBreaks As HPageBreaksWrapper

		''' <summary>Excel.HPageBreak</summary>
		Private _hPageBreak As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="hPageBreaks">親のHPageBreaks</param>
		''' <param name="hPageBreak">Excel.HPageBreak</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal hPageBreaks As HPageBreaksWrapper, ByVal hPageBreak As Object)
			MyBase.New(hPageBreaks.ApplicationWrapper)
			_hPageBreaks = hPageBreaks
			_hPageBreak = hPageBreak
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_hPageBreak)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _hPageBreak
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
				Return DirectCast(InvokeGetProperty(_hPageBreak, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' 指定した改ページの種類を取得します。画面全体または印刷範囲のみのいずれかを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Extent() As XlPageBreakExtent
			Get
				Return DirectCast(InvokeGetProperty(_hPageBreak, "Extent", Nothing), XlPageBreakExtent)
			End Get
		End Property

		''' <summary>
		''' 改ページの位置を定義するセル (Range オブジェクト) を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Location() As RangeWrapper
			Get
				Dim obj As Object
				Dim wrapper As RangeWrapper
				obj = InvokeGetProperty(_hPageBreak, "Location", Nothing)
				wrapper = New RangeWrapper(_hPageBreaks.Sheet, obj)
				addXlsObject(wrapper)
				Return wrapper
			End Get
			Set(ByVal value As RangeWrapper)

			End Set
		End Property

		''' <summary>
		''' 改ページの種類を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property [Type]() As XlPageBreak
			Get
				Return DirectCast(InvokeGetProperty(_hPageBreak, "Type", Nothing), XlPageBreak)
			End Get
			Set(ByVal value As XlPageBreak)
				InvokeSetProperty(_hPageBreak, "Type", New Object() {value})
			End Set
		End Property

#End Region
#Region " メソッド "

		''' <summary>
		''' オブジェクトを削除します。
		''' </summary>
		''' <remarks></remarks>
		Public Sub Delete()
			InvokeMethod(_hPageBreak, "Delete", Nothing)
		End Sub

		''' <summary>
		''' 改ページ プレビューで表示されている場合、改ページを印刷領域の外にドラッグします。
		''' </summary>
		''' <param name="Direction">必ず指定します。XlDirection 列挙型の定数を使用します。 改ページをドラッグする方向を指定します。</param>
		''' <param name="RegionIndex">必ず指定します。整数型 (Integer) の値を指定します。改ページの印刷範囲領域のインデックスを指定します (ユーザーが改ページをドラッグする場合に、マウス ボタンを押した時点でのマウス ポインタが位置する領域)。印刷範囲が隣接している場合、印刷領域は 1 つだけです。印刷範囲が隣接していない場合、印刷領域は複数あります。</param>
		''' <remarks></remarks>
		Public Sub DragOff( _
		  <InAttribute()> ByVal Direction As XlDirection, _
		  <InAttribute()> ByVal RegionIndex As Integer _
		 )
			InvokeMethod(_hPageBreak, "DragOff", New Object() {Direction, RegionIndex})
		End Sub
#End Region

	End Class

End Namespace
