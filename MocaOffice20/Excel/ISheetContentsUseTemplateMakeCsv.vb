
Imports System.IO

Namespace Excel

	''' <summary>
	''' 一覧形式のテンプレートシートを使用したときに、
	''' データを一度CSVファイルへ出力し、CSVファイルを読込んでExcelへ貼り付ける手法のインタフェース
	''' </summary>
	''' <remarks></remarks>
	Public Interface ISheetContentsUseTemplateMakeCsv
		Inherits ISheetContents

		''' <summary>
		''' 明細部の出力開始行
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' テンプレートシート上の一覧出力される最初の行を返すようにする。
		''' </remarks>
		ReadOnly Property StartRow() As Integer

		''' <summary>
		''' 明細部の出力開始列
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks>
		''' テンプレートシート上の一覧出力される最初の列を返すようにする。
		''' </remarks>
		ReadOnly Property StartCol() As Integer

		''' <summary>
		''' 明細部の出力列数
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property ColumnLength() As Integer

		''' <summary>
		''' 出力するデータ件数
		''' </summary>
		''' <value></value>
		''' <returns></returns>
		''' <remarks></remarks>
		ReadOnly Property DataCount() As Integer

		''' <summary>
		''' データをCSV中間ファイルとして出力する
		''' </summary>
		''' <param name="csv"></param>
		''' <remarks>
		''' 「，」（カンマ）区切りの文字列を出力してください。<br/>
		''' </remarks>
		Sub CsvWrite(ByRef csv As StreamWriter)

		''' <summary>
		''' CSVファイルをExcelにて開く時のフォーマットを指定する
		''' </summary>
		''' <param name="columnIndex"></param>
		''' <remarks>
		''' CSVファイルをExcelにて開く時のフォーマットを指定する場合は指定してください。<br/>
		''' デフォルトでは「一般」のフォーマットにて読み込みます。
		''' </remarks>
		Function SetCsvOpenFormat(ByVal columnIndex As Integer) As XlColumnDataType

	End Interface

End Namespace
