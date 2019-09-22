<%
'https://github.com/fernando-nishino/ClassicASP.RemoverAcentuacao
'Mude o encoding deste arquivo para: 'UTF-8 WITH BOM'

Function RemoverAcentuacao(palavra)
	Dim letra, i, palavraConvertida
	If IsNull(palavra) Then palavra = ""
	For i = 1 To Len(palavra)
		letra = Mid(palavra, i, 1)
		Select Case (letra)
			Case "Á", "À", "Â", "Ä", "Ã"	letra = "A"
			Case "É", "È", "Ê", "Ë"			letra = "E"
			Case "Í", "Ì", "Î", "Ï"			letra = "I"
			Case "Ó", "Ò", "Ô", "Ö", "Õ"	letra = "O"
			Case "Ú", "Ù", "Û", "Ü"			letra = "U"
			Case "Ç"						letra = "C"
			Case "Ñ"						letra = "N"
			Case "á", "à", "â", "ä", "ã"	letra = "a"
			Case "é", "è", "ê", "ë"			letra = "e"
			Case "í", "ì", "î", "ï"			letra = "i"
			Case "ó", "ò", "ô", "ö", "õ"	letra = "o"
			Case "ú", "ù", "û", "ü"			letra = "u"
			Case "ç"						letra = "c"
			Case "ñ"						letra = "n"
		End Select
		palavraConvertida = palavraConvertida & letra
	Next
	RemoverAcentuacao = palavraConvertida
End Function

Response.Write RemoverAcentuacao("ação, emoção, órgão, copo-d'água")
%>