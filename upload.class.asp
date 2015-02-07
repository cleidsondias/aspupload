<%

'constants:			
Const MAX_UPLOAD_SIZE=200000 'bytes
Const MSG_NO_DATA="não existe dados para carga"
Const MSG_EXCEEDED_MAX_SIZE="O limite máximo de carga de arquivo é de "

Class Upload
	Private m_Request
	Private m_Files
		
	Public Property Get FileCount
		FileCount = m_Files.Count
	End Property
	
	Public Function File(index)
		Dim keys
		keys = m_Files.Keys
		Set File = m_Files(keys(index))
	End Function
	
	Public Function FileItem(strName)
		Dim objFileData
		If m_Request.Exists(strName) Then
			if m_Files.Exists(m_Request(strName)) then
				Set FileItem = m_Files(m_Request(strName))
			else
				Set objFileData = New FileData
				Set FileItem = objFileData
			end if
		Else 
			Set FileItem = ""
		End If
	End Function	
	
	Public Default Property Get Item(strName)
		If m_Request.Exists(strName) Then
			Item = m_Request(strName)
		Else  
			Item = ""
		End If
	End Property
		
	'Construtor da classe
	Private Sub Class_Initialize
		Dim iBytesCount, strBinData
		
		'recupera o total de bytes que foi feita a carga dos dados (upload)
		iBytesCount = Request.TotalBytes
		
		'aborta se não localizar nada
		If iBytesCount=0 Then
			Err.Number = 6555
			Err.Description = MSG_NO_DATA
		End If
		
		'aborta se ultrapassar o limite pré estabelecido para o total da carga
		If iBytesCount>MAX_UPLOAD_SIZE Then
			Err.Number = 6554
			Err.Description = MSG_EXCEEDED_MAX_SIZE & MAX_UPLOAD_SIZE & " bytes caso a soma dos arquivos seja inferior, verifique a variavel AspMaxRequestEntityAllowed no arquivo Metabase "
		End If
		
		'leia a carga de dados
		strBinData = Request.BinaryRead(iBytesCount)
		
		'inicializa as coleções
		Set m_Request = Server.CreateObject("Scripting.Dictionary")
		Set m_Files = Server.CreateObject("Scripting.Dictionary")
     		
		'faz a carga nas coleções
		Call BuildUpload(strBinData)
	End Sub
	
	'Finalizador da classe
	Private Sub Class_Terminate
		Dim fileName
		If IsObject(m_Request) Then
			m_Request.RemoveAll
			Set m_Request = Nothing
		End If
		If IsObject(m_Files) Then
			For Each fileName In m_Files.Keys
				Set m_Files(fileName)=Nothing
			Next
			m_Files.RemoveAll
			Set m_Files = Nothing
		End If
	End Sub
	
	Private Sub BuildUpload(ByVal strBinData)
		Dim strBinQuote, strBinCRLF, iValuePos
		Dim iPosBegin, iPosEnd, strBoundaryData
		Dim strBoundaryEnd, iCurPosition, iBoundaryEndPos
		Dim strElementName, strFileName, objFileData
		Dim strFileType, strFileData, strElementValue
		Dim oBINARYCOMMON
		
		Set oBINARYCOMMON = new BINARYCOMMON
		
		strBinQuote = oBINARYCOMMON.AsciiToBinary(chr(34))
		strBinCRLF = oBINARYCOMMON.AsciiToBinary(chr(13))
		
		'procura as fronteiras
		'RFC1867
		iPosBegin = 1
		iPosEnd = InstrB(iPosBegin, strBinData, strBinCRLF)
		strBoundaryData = MidB(strBinData, iPosBegin, iPosEnd-iPosBegin)
		iCurPosition = InstrB(1, strBinData, strBoundaryData)
		strBoundaryEnd = strBoundaryData & oBINARYCOMMON.AsciiToBinary("--")
		iBoundaryEndPos = InstrB(strBinData, strBoundaryEnd)
		
		'inserindo as cargas nas coleções
		Do until (iCurPosition>=iBoundaryEndPos) Or (iCurPosition=0)
			'ignora dados irrelevantes
			iPosBegin = InstrB(iCurPosition, strBinData, oBINARYCOMMON.AsciiToBinary("Content-Disposition"))
			iPosBegin = InstrB(iPosBegin, strBinData, oBINARYCOMMON.AsciiToBinary("name="))
			iValuePos = iPosBegin
			
			'obtem os nome da tag input file
			iPosBegin = iPosBegin+6
			iPosEnd = InstrB(iPosBegin, strBinData, strBinQuote)
			strElementName = oBINARYCOMMON.BinaryToAscii(MidB(strBinData, iPosBegin, iPosEnd-iPosBegin))
			
			'Fluxo se for arquivo
			iPosBegin = InstrB(iCurPosition, strBinData, oBINARYCOMMON.AsciiToBinary("filename="))
			iPosEnd = InstrB(iPosEnd, strBinData, strBoundaryData)
			If (iPosBegin>0) And (iPosBegin<iPosEnd) Then
				'ignora dados irrelevantes
				iPosBegin = iPosBegin+10
				
				'leia o nome do arquivo
				iPosEnd = InstrB(iPosBegin, strBinData, strBinQuote)
				strFileName = oBINARYCOMMON.BinaryToAscii(MidB(strBinData, iPosBegin, iPosEnd-iPosBegin))
				
				'Criando o arquivo, caso contrario leia de forma ordinaria
				If Len(strFileName)>0 Then
				
					'create file data:
					Set objFileData = New FileData
					objFileData.FileName = strFileName
					
					'leia o tipo de arquivo
					iPosBegin = InstrB(iPosEnd, strBinData, oBINARYCOMMON.AsciiToBinary("Content-Type:"))
					iPosBegin = iPosBegin+14
					iPosEnd = InstrB(iPosBegin, strBinData, strBinCRLF)
					strFileType = oBINARYCOMMON.BinaryToAscii(MidB(strBinData, iPosBegin, iPosEnd-iPosBegin))
					objFileData.ContentType = strFileType
					
					'leia o conteudo do arquivo
					iPosBegin = iPosEnd+4
					iPosEnd = InstrB(iPosBegin, strBinData, strBoundaryData)-2
					strFileData = MidB(strBinData, iPosBegin, iPosEnd-iPosBegin)
					
					'verifica se o arquivo não está vazio
					If LenB(strFileData)>0 Then
						objFileData.Contents = strFileData
						objFileData.Item = strElementName
						'adiciona o arquivo na coleção
						Set m_Files(strFileName) = objFileData
					Else  
						Set objFileData = Nothing
					End If
				End If
				
				strElementValue = strFileName
			Else  
				'valor obtido de forma ordinaria, apenas leia
				iPosBegin = InstrB(iValuePos, strBinData, strBinCRLF)
				iPosBegin = iPosBegin+4
				iPosEnd = InstrB(iPosBegin, strBinData, strBoundaryData)-2
				strElementValue = oBINARYCOMMON.BinaryToAscii(MidB(strBinData, iPosBegin, iPosEnd-iPosBegin))
			End If
			
			'adiciona a requisição na coleção
			m_Request(strElementName) = strElementValue
			
			'vai para o proximo elemento da requisição
			iCurPosition = InstrB(iCurPosition+LenB(strBoundaryData), strBinData, strBoundaryData)
		Loop
	End Sub
	
End Class

Class FileData
	Private m_fileName
	Private m_contentType
	Private m_BinaryContents
	Private m_AsciiContents
	Private m_item
	
	Public Default Property Get BinaryContents
		BinaryContents = m_BinaryContents
	End Property
	
	Public Property Get Extension
		Extension = GetExtension(m_fileName)
	End Property
	
	Public Property Get FileName
		FileName = m_fileName
	End Property
	
	Public Property Get Item
		Item = m_item
	End Property
	
	Public Property Get ContentType
		ContentType = m_contentType
	End Property
	
	Public Property Let FileName(strName)
		Dim arrTemp
		arrTemp = Split(strName, "\")
		m_fileName = arrTemp(UBound(arrTemp))
	End Property
	
	Public Property Let Item(strItem)
		 m_item = strItem
	End Property
		
	Public Property Let ContentType(strType)
		m_contentType = strType
	End Property
	
	Public Property Let Contents(strData)
		Dim oBINARYCOMMON
		Set oBINARYCOMMON = new BINARYCOMMON
		m_BinaryContents = strData
		m_AsciiContents = oBINARYCOMMON.RSBinaryToString(m_BinaryContents)
	End Property
	
	Public Property Get Size
		if LenB(m_BinaryContents) > 0 then
			Size = LenB(m_BinaryContents)
		else
			Size = 0
		end if
	End Property
	
	Private Function GetExtension(strPath)
		Dim arrTemp
		arrTemp = Split(strPath, ".")
		GetExtension = ""
		If UBound(arrTemp)>0 Then
			GetExtension = arrTemp(UBound(arrTemp))
		End If
	End Function
		
	Public Sub SaveToDisk(strFolderPath, ByRef strNewFileName)
		Dim strPath, objFSO, objFile
		Dim i
		Dim objStream, strExtension
		
		strPath = strFolderPath&"\"
		If Len(strNewFileName)=0 Then
			strPath = strPath & m_fileName
		Else  
			strExtension = GetExtension(strNewFileName)
			If Len(strExtension)=0 Then
				strNewFileName = strNewFileName & "." & GetExtension(m_fileName)
			End If
			strPath = strPath & strNewFileName
		End If
		
		Set objFSO = Server.CreateObject("Scripting.FileSystemObject")
		Set objFile = objFSO.CreateTextFile(strPath)
		
		objFile.Write(m_AsciiContents)
		
		'''For i=1 to LenB(m_BinaryContents)
		'''	objFile.Write chr(AscB(MidB(m_BinaryContents, i, 1)))
		'''Next
		
		objFile.Close
		Set objFile=Nothing
		Set objFSO=Nothing
	End Sub

End Class

Class BINARYCOMMON

	Function SimpleBinaryToString(Binary)
	  'SimpleBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
	  'to a string (BSTR) using MultiByte VBS functions
	  Dim I, S
	  For I = 1 To LenB(Binary)
		S = S & Chr(AscB(MidB(Binary, I, 1)))
	  Next
	  SimpleBinaryToString = S
	End Function
	
	Function BinaryToString(Binary)
	  'Antonin Foller, http://www.motobit.com
	  'Optimized version of a simple BinaryToString algorithm.
	  
	  Dim cl1, cl2, cl3, pl1, pl2, pl3
	  Dim L
	  cl1 = 1
	  cl2 = 1
	  cl3 = 1
	  L = LenB(Binary)
	  
	  Do While cl1<=L
		pl3 = pl3 & Chr(AscB(MidB(Binary,cl1,1)))
		cl1 = cl1 + 1
		cl3 = cl3 + 1
		If cl3>300 Then
		  pl2 = pl2 & pl3
		  pl3 = ""
		  cl3 = 1
		  cl2 = cl2 + 1
		  If cl2>200 Then
			pl1 = pl1 & pl2
			pl2 = ""
			cl2 = 1
		  End If
		End If
	  Loop
	  BinaryToString = pl1 & pl2 & pl3
	End Function

	Function RSBinaryToString(xBinary)
	  'Antonin Foller, http://www.motobit.com
	  'RSBinaryToString converts binary data (VT_UI1 | VT_ARRAY Or MultiByte string)
	  'to a string (BSTR) using ADO recordset

	  Dim Binary
	  'MultiByte data must be converted To VT_UI1 | VT_ARRAY first.
	  If vartype(xBinary)=8 Then Binary = MultiByteToBinary(xBinary) Else Binary = xBinary
	  
	  Dim RS, LBinary
	  Const adLongVarChar = 201
	  Set RS = CreateObject("ADODB.Recordset")
	  LBinary = LenB(Binary)
	  
	  If LBinary>0 Then
		RS.Fields.Append "mBinary", adLongVarChar, LBinary
		RS.Open
		RS.AddNew
		  RS("mBinary").AppendChunk Binary 
		RS.Update
		RSBinaryToString = RS("mBinary")
	  Else
		RSBinaryToString = "não foi possivél converter"
	  End If
	End Function

	Function MultiByteToBinary(MultiByte)
	  ' 2000 Antonin Foller, http://www.motobit.com
	  ' MultiByteToBinary converts multibyte string To real binary data (VT_UI1 | VT_ARRAY)
	  ' Using recordset
	  Dim RS, LMultiByte, Binary
	  Const adLongVarBinary = 205
	  Set RS = CreateObject("ADODB.Recordset")
	  LMultiByte = LenB(MultiByte)
	  If LMultiByte>0 Then
		RS.Fields.Append "mBinary", adLongVarBinary, LMultiByte
		RS.Open
		RS.AddNew
		  RS("mBinary").AppendChunk MultiByte & ChrB(0)
		RS.Update
		Binary = RS("mBinary").GetChunk(LMultiByte)
	  End If
	  MultiByteToBinary = Binary
	End Function

	Function Stream_BinaryToString(Binary, CharSet)
		'Stream_BinaryToString Function
		'2003 Antonin Foller, http://www.motobit.com
		'Binary - VT_UI1 | VT_ARRAY data To convert To a string 
		'CharSet - charset of the source binary data - default is "us-ascii"  Const adTypeText = 2
	  Const adTypeBinary = 1
	  
	  'Create Stream object
	  Dim BinaryStream 'As New Stream
	  Set BinaryStream = CreateObject("ADODB.Stream")
	  
	  'Specify stream type - we want To save text/string data.
	  BinaryStream.Type = adTypeBinary
	  
	  'Open the stream And write text/string data To the object
	  BinaryStream.Open
	  BinaryStream.Write Binary
	  
	  
	  'Change stream type To binary
	  BinaryStream.Position = 0
	  BinaryStream.Type = adTypeText
	  
	  'Specify charset For the source text (unicode) data.
	  If Len(CharSet) > 0 Then
		BinaryStream.CharSet = CharSet
	  Else
		BinaryStream.CharSet = "us-ascii"
	  End If
	  
	  'Open the stream And get binary data from the object
	  Stream_BinaryToString = BinaryStream.ReadText
	End Function
	
	Public Function lngConvert(strTemp)
		lngConvert = clng(asc(left(strTemp, 1)) + ((asc(right(strTemp, 1)) * 256)))
	end function
	
	Public Function lngConvert2(strTemp)
		lngConvert2 = clng(asc(right(strTemp, 1)) + ((asc(left(strTemp, 1)) * 256)))
	end function
	
	Public Function AsciiToBinary(strAscii)
		Dim i, char, result
		result = ""
		For i=1 to Len(strAscii)
			char = Mid(strAscii, i, 1)
			result = result & chrB(AscB(char))
		Next
		AsciiToBinary = result
	End Function
	
	Public Function BinaryToAscii(strBinary)
		Dim i, result
		result = ""
		For i=1 to LenB(strBinary)
			result = result & chr(AscB(MidB(strBinary, i, 1))) 
		Next
		BinaryToAscii = result
	End Function
	
	Function BinaryToHex(binary)
		text = ""
		For bytenum = 1 To LenB(binary)
			h = Right( "0" & Hex(AscB(MidB(binary,bytenum,1))), 2 )
			text = text & h
		Next
		BinaryToHex = text
	End Function
		
	Function HexToBinary(hexText)
		binary = ""
		For bytenum = 1 To Len(hexText) Step 2
			oneByte = CInt("&H" & Mid(hexText,bytenum,2) )
			binary = binary & ChrB(oneByte)
		Next
		HexToBinary = binary
	End Function

End Class

%>