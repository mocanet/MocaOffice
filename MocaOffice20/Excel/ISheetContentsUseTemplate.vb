
Namespace Excel

	''' <summary>
	''' 一覧形式のテンプレートシートを使用したときに、
	''' データを多次元配列に変換し値を設定する手法のインタフェース
	''' </summary>
	''' <remarks>
	''' </remarks>
	Public Interface ISheetContentsUseTemplate
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
		''' 出力するデータ配列を作成する
		''' </summary>
		''' <remarks>
		''' 出力するデータを必要な数の行列に該当する多次元配列を作成し戻り値として返します。<br/>
		''' 作成する配列は、<see cref="DataCount"/>行，<see cref="ColumnLength"/>列としてください。<br/>
		''' 作成された配列を、<see cref="StartRow"/>行：<see cref="StartCol"/>列を元に設定されます。
		''' </remarks>
		Function MakeArrayData() As Array

	End Interface

End Namespace
