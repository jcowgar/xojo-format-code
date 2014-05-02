#tag Module
Protected Module FormatCodePreferences
	#tag Constant, Name = AdditionalKeywords, Type = String, Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant

	#tag Constant, Name = AlignAs, Type = String, Dynamic = False, Default = \"No", Scope = Public
	#tag EndConstant

	#tag Constant, Name = AlignEqual, Type = String, Dynamic = False, Default = \"No", Scope = Public
	#tag EndConstant

	#tag Constant, Name = CaseConversion, Type = String, Dynamic = False, Default = \"1", Scope = Public
	#tag EndConstant

	#tag Constant, Name = KeywordsToLowerCase, Type = String, Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant

	#tag Constant, Name = KeywordsToTitleCase, Type = String, Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant

	#tag Constant, Name = KeywordsToUpperCase, Type = String, Dynamic = False, Default = \"", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PadComma, Type = String, Dynamic = False, Default = \"Yes", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PadOperators, Type = String, Dynamic = False, Default = \"Yes", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PadParensInner, Type = String, Dynamic = False, Default = \"No", Scope = Public
	#tag EndConstant

	#tag Constant, Name = PadParensOuter, Type = String, Dynamic = False, Default = \"No", Scope = Public
	#tag EndConstant


	#tag ViewBehavior
		#tag ViewProperty
			Name="Index"
			Visible=true
			Group="ID"
			InitialValue="-2147483648"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Left"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Name"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Super"
			Visible=true
			Group="ID"
			Type="String"
		#tag EndViewProperty
		#tag ViewProperty
			Name="Top"
			Visible=true
			Group="Position"
			InitialValue="0"
			Type="Integer"
		#tag EndViewProperty
	#tag EndViewBehavior
End Module
#tag EndModule
