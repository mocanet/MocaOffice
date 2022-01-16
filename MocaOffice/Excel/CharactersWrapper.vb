
Imports System.Runtime.InteropServices

Namespace Excel

	''' <summary>
	''' オブジェクトに含まれる文字列の文字を表します。Characters オブジェクトを使用すると、文字列のうちの一部だけを修正できます。
	''' </summary>
	''' <remarks></remarks>
	Public Class CharactersWrapper
		Inherits AbstractExcelWrapper

		''' <summary>親のExcel.Sheet のラッパー</summary>
		Private _range As RangeWrapper

		''' <summary>Excel.Characters</summary>
		Private _characters As Object

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " コンストラクタ "

		''' <summary>
		''' コンストラクタ
		''' </summary>
		''' <param name="range">親のRange</param>
		''' <param name="characters">Excel.Characters</param>
		''' <remarks>
		''' </remarks>
		Public Sub New(ByVal range As RangeWrapper, ByVal characters As Object)
			MyBase.New(range.ApplicationWrapper)
			_range = range
			_characters = characters
		End Sub

#End Region
#Region " Overrides "

		''' <summary>
		''' 自分自身で管理しているExcel関係のオブジェクトのメモリ開放
		''' </summary>
		''' <remarks></remarks>
		Public Overrides Sub MyDispose()
			ReleaseExcelObject(_characters)
		End Sub

		''' <summary>
		''' 取得した Excel インスタンス
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Friend Overrides ReadOnly Property OrigianlInstance() As Object
			Get
				Return _characters
			End Get
		End Property

#End Region
#Region " プロパティ "

		''' <summary>
		''' 指定した範囲の文字列です。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Caption() As String
			Get
				Return DirectCast(InvokeGetProperty(_characters, "Caption", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_characters, "Caption", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' コレクション内のオブジェクトの数を取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Count() As Integer
			Get
				Return DirectCast(InvokeGetProperty(_characters, "Count", Nothing), Integer)
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
				Return DirectCast(InvokeGetProperty(_characters, "Creator", Nothing), XlCreator)
			End Get
		End Property

		''' <summary>
		''' 指定したオブジェクトのフォントを表す Font オブジェクトを取得します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public ReadOnly Property Font() As FontWrapper
			Get
				Return DirectCast(InvokeGetProperty(_characters, "Font", Nothing), FontWrapper)
			End Get
		End Property

		''' <summary>
		''' 指定した Characters オブジェクトにふりがなテキストを設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property PhoneticCharacters() As String
			Get
				Return DirectCast(InvokeGetProperty(_characters, "PhoneticCharacters", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_characters, "PhoneticCharacters", New Object() {value})
			End Set
		End Property

		''' <summary>
		''' 指定したオブジェクトの文字列を設定します。
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Property Text() As String
			Get
				Return DirectCast(InvokeGetProperty(_characters, "Text", Nothing), String)
			End Get
			Set(ByVal value As String)
				InvokeSetProperty(_characters, "Text", New Object() {value})
			End Set
		End Property

#End Region

		''' <summary>
		''' オブジェクトを削除します。
		''' </summary>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Delete() As Object
			Return InvokeMethod(_characters, "Delete", Nothing)
		End Function

		''' <summary>
		''' 選択した文字列の前に文字列を挿入します。
		''' </summary>
		''' <param name="str"></param>
		''' <returns></returns>
		''' <remarks></remarks>
		Public Function Insert( _
		  <InAttribute()> ByVal str As String _
		 ) As Object
			Return InvokeMethod(_characters, "Insert", New Object() {str})
		End Function

	End Class

End Namespace
