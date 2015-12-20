
Namespace Excel

	''' <summary>
	''' �V�[�g���e���\��
	''' </summary>
	''' <remarks></remarks>
	Friend Class SheetContents

		''' <summary></summary>
		Protected contents As ISheetContents

		''' <summary></summary>
		Protected sheet As SheetWrapper

		''' <summary>log4net logger</summary>
		Private ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType)

#Region " �R���X�g���N�^ "

		''' <summary>
		''' �R���X�g���N�^
		''' </summary>
		''' <remarks></remarks>
		Public Sub New(ByVal sheetContents As ISheetContents, ByVal sheet As SheetWrapper)
			Me.contents = sheetContents
			Me.sheet = sheet
		End Sub

#End Region

		''' <summary>
		''' �o�͓��e���Z���֐ݒ肷��
		''' </summary>
		''' <remarks>
		''' </remarks>
		Public Overridable Sub WriteContents()
			' �V�[�g���ݒ�
			If contents.SaveSheetName.Length > 0 Then
				Me.sheet.Name = contents.SaveSheetName
			End If

			contents.WriteContents(sheet)
		End Sub

	End Class

End Namespace
