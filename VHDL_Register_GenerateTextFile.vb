Sub GenerateTextFile()
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCellE As Object, oCellG As Object, RegAddr As Object, Default_V As Object
    Dim oCellI25 As Object, oCellRDEN As Object, oCellWREN As Object
    Dim oCellRENO As Object, oCellADDR As Object, oCellRST As Object
    Dim oCellCLK As Object, oCellRDO As Object, oCellWRD As Object
    Dim sFilePath As String, sFormattedText As String
    Dim sCellValueE As String, sCellValueG As String, sRegAddr As String, DefaultValueI As String,ZeroPaddedRegSize As String,ZeroPaddedReg As String
    Dim i As Integer, j As Integer, j_temp As Integer, Row As Integer, Rd_Reg_cnt As Integer
    Dim BitWidth As Integer 
    Dim SPLIT_VALUE As Integer, Dec As Double, Ind As Integer
    Dim RDEN As String, WREN As String, RDEO As String
    Dim ADDR As String, RST As String, CLK As String
    Dim RDO As String, WRD As String
    Dim oSimpleFileAccess As Object, oOutputStream As Object
    Dim isWindows As Boolean
    Dim currentDateTime As String
    Dim maxSignalLength As Integer
	Dim formattedSignalName As String

    ' Get active sheet
    oDoc = ThisComponent
    oSheet = oDoc.Sheets(0) ' Assuming first sheet contains the data
    
    START_OF_Row = 10
    ' Column selection
    RegisterName = "E"
    Address      = "F"
    RdORWr       = "G"
    RegSize      = "H"
    ResetValue   = "I"
    
    ' Fetch bit width from L23
    oCellI25 = oSheet.getCellRangeByName("M23")
    BitWidth = IIf(IsNumeric(oCellI25.getValue()) And oCellI25.getValue() > 0, oCellI25.getValue(), 32)
    BitRange_max = "(" & (BitWidth - 1) & " downto 0)"

    ' Fetch FPGA-specific values
    RDEN    = GetCellValueOrDefault(oSheet, "M19", "FPGA_RDEN")
    WREN    = GetCellValueOrDefault(oSheet, "M20", "FPGA_WREN")
    RDEO    = GetCellValueOrDefault(oSheet, "M21", "FPGA_RDEN_OUT")
    ADDR    = GetCellValueOrDefault(oSheet, "M18", "FPGA_ADDR")
    RST     = GetCellValueOrDefault(oSheet, "M14", "RESET")
    CLK     = GetCellValueOrDefault(oSheet, "M15", "CLK")
    RDO     = GetCellValueOrDefault(oSheet, "M16", "READ_DATA")
    WRD     = GetCellValueOrDefault(oSheet, "M17", "WR_DATA")
    RD_SPV  = GetCellValueOrDefault(oSheet, "M24", "16")

    ' Initialize parameters
    SPLIT_VALUE  = RD_SPV
    sFormattedText = ""
    
    ' Get the current date and time
    currentDateTime = Format(Now, "dd/mm/yyyy hh:mm am/pm")
    sFormattedText = "---------------------- ( " & currentDateTime & " ) ----------------------" & Chr(10) & Chr(10)
    
	'-----------------Signal_Declaration-----------------

	' Find the maximum signal name length for alignment
	maxSignalLength = 0
	SepRegMaxLength = 0
	Row = START_OF_Row
	Do While True
	    oCellE = oSheet.getCellRangeByName(RegisterName & Row)
	    Sep_Reg_Size = oSheet.getCellRangeByName(RegSize & Row)
	    sCellValueE = oCellE.getString()
	    sCellValueH = Sep_Reg_Size.getValue()
	    If sCellValueE = "" Then Exit Do
	    If Len(sCellValueE) > maxSignalLength Then
	        maxSignalLength = Len(sCellValueE)
	    End If
	    If Len(PadDataReg(sCellValueH, BitWidth)) > SepRegMaxLen Then
	        SepRegMaxLen = Len(sCellValueE)+Len(PadDataReg(sCellValueH, BitWidth))
	    End If
	    Row = Row + 1
	Loop
	
	' Reset Row and format the output
	Row = START_OF_Row
	Do While True
	    oCellE = oSheet.getCellRangeByName(RegisterName & Row)
	    oCellG = oSheet.getCellRangeByName(RdORWr & Row)
	    Sep_Reg_Size = oSheet.getCellRangeByName(RegSize & Row)
	    sCellValueE = oCellE.getString()
	    sCellValueG = oCellG.getString()
	    sCellValueH = Sep_Reg_Size.getValue()
	    sCellStringH = Sep_Reg_Size.getString()
	    
	    If sCellStringH = "" Then
	    	BitRange =  BitRange_max
	    Else
	    	BitRange =  "(" & (sCellValueH - 1) & " downto 0)"
	    End If 
	    
	
	    If sCellValueE = "" Then Exit Do
	    i = i + 1
	
	    ' Format the signal name with padding
	    formattedSignalName = sCellValueE & "_reg " & String(maxSignalLength - Len(sCellValueE), " ")
	
	    ' Generate the signals with alignment
	    sFormattedText = sFormattedText & "signal " & formattedSignalName & ": std_logic_vector" & BitRange & ":=(others => '0');" & Chr(10)
	    If sCellValueG = "RW" Or sCellValueG = "" Then
	        Rd_Reg_cnt = Rd_Reg_cnt + 1
	        formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
	        sFormattedText = sFormattedText & "signal " & formattedSignalName & " : std_logic:='0';" & Chr(10)
	        formattedSignalName = sCellValueE & "_wr " & String(maxSignalLength - Len(sCellValueE), " ")
	        sFormattedText = sFormattedText & "signal " & formattedSignalName & " : std_logic:='0';" & Chr(10)
	    ElseIf sCellValueG = "R" Then
	        Rd_Reg_cnt = Rd_Reg_cnt + 1
	        formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
	        sFormattedText = sFormattedText & "signal " & formattedSignalName & " : std_logic:='0';" & Chr(10)
	    ElseIf sCellValueG = "W" Then
	    	formattedSignalName = sCellValueE & "_wr " & String(maxSignalLength - Len(sCellValueE), " ")
	        sFormattedText = sFormattedText & "signal " & formattedSignalName & " : std_logic:='0';" & Chr(10)
	    End If
	
	    ' Add a blank line between signal blocks
	    sFormattedText = sFormattedText & Chr(10)
	
	    ' Move to the next row
	    Row = Row + 1
	Loop
    
   	'-----------------Calculating no of Splits if Rd_Reg_cnt > SPLIT_VALUE then -----------------
  	if Rd_Reg_cnt > SPLIT_VALUE then
  		    
	    Do While True
	    	If i = 0 Then 
	    		if k > 0 then
  					j_temp = j_temp +1
  				End if
  				Exit Do
  			end if
  			
	    	k = k + 1
	    	i = i - 1 
  		
  			If K = SPLIT_VALUE Then 
  				j_temp = j_temp +1
  				k = 0
  			end if
    	Loop

	    j = j_temp
	    
	    '-----------------Signal Declaration for Rd_EN and Rd_Data -----------------
  		Do While True
  			If j = 0 Then 
  				sFormattedText = sFormattedText & Chr(10)
  				Exit Do
  			end if
  				formattedSignalName = sCellValueE & j & String(maxSignalLength - Len("Rd_EN"), " ")
    			sFormattedText = sFormattedText & "signal Rd_EN_OUT_" & formattedSignalName & " : std_logic:='0';" & Chr(10)
    			j=j-1
    	Loop
    	
    	j = j_temp
    	
    	Do While True
  			If j = 0 Then 
  				sFormattedText = sFormattedText & Chr(10)
  				Exit Do
  			end if
  				formattedSignalName = sCellValueE & j & String(maxSignalLength - Len("Rd_DA"), " ")
    			sFormattedText = sFormattedText & "signal Rd_DATA_" & formattedSignalName & " : std_logic_vector" & BitRange & ":=(others => '0');" & Chr(10)
    			j=j-1
    	Loop
    	j = j_temp	
    	'sFormattedText = sFormattedText & j_temp &" * j_temp *  " &  Chr(10)
        'sFormattedText = sFormattedText & k &" * k *  " &  Chr(10)
 	End If
     
    sFormattedText = sFormattedText & "------------------------------- Segments after begin -------------------------------" & Chr(10) & Chr(10)
    
    '-----------------Rd_en & Wr_en generation using Mux-----------------
     Row = START_OF_Row
	 Do While True
        ' Get the cell values from column E ,G and F  
        oCellE = oSheet.getCellRangeByName(RegisterName & Row)
        oCellG = oSheet.getCellRangeByName(RdORWr & Row)
        RegAddr = oSheet.getCellRangeByName(Address & Row)
        sCellValueE = oCellE.getString()
        sCellValueG = oCellG.getString()
        sRegAddr    = RegAddr.getString()
		
        ' Stop if the cell in column E is empty (assume end of entries)
         If sCellValueE = "" Then 
        	Row = START_OF_Row
        	j = 0
        	Exit Do
        end if
		
        ' Generate the signals based on the value in column G
        If sCellValueG = "RW" Or sCellValueG = "" Then
        	formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
            sFormattedText =  sFormattedText & formattedSignalName & " <= '1' when "& RDEN &" = '1' and "& ADDR & " = x""" & sRegAddr & """ else '0';" & Chr(10)
            formattedSignalName = sCellValueE & "_wr " & String(maxSignalLength - Len(sCellValueE), " ")
            sFormattedText =  sFormattedText & formattedSignalName & " <= '1' when "& WREN &" = '1' and "& ADDR & " = x""" & sRegAddr & """ else '0';" & Chr(10)
        ElseIf sCellValueG = "R" Then
        	formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
            sFormattedText =  sFormattedText & formattedSignalName & " <= '1' when "& RDEN &" = '1' and "& ADDR & " = x""" & sRegAddr & """ else '0';" & Chr(10)
        ElseIf sCellValueG = "W" Then
       	    formattedSignalName = sCellValueE & "_wr " & String(maxSignalLength - Len(sCellValueE), " ")
            sFormattedText =  sFormattedText & formattedSignalName & " <= '1' when "& WREN &" = '1' and "& ADDR & " = x""" & sRegAddr & """ else '0';" & Chr(10)
        End If
		
        ' Add a blank line between Mux blocks
        sFormattedText = sFormattedText & Chr(10)
		
        ' Move to the next row
        Row = Row + 1
    Loop
    
    '-----------------Read_Data_OR-----------------
        if (Rd_Reg_cnt > SPLIT_VALUE) then
	 		sFormattedText = sFormattedText & RDO &" <= "
	 	Do While True
  			If j = j_temp Then 
  				sFormattedText = sFormattedText & Chr(10) & Chr(10)
  				Exit Do
  			END IF
    			IF j = j_temp-1 THEN
    				sFormattedText = sFormattedText & "Rd_DATA_" & j+1 & ";" 
    			ELSE
    				sFormattedText = sFormattedText & "Rd_DATA_" & j+1 & " or "
    			END IF 
    			j=j+1
    	Loop
    
    	j = 0
    else
    	'sFormattedText = sFormattedText & RDO &" <= "
   	End if
   	
	 '-----------------Read_DATA_Latch-----------------
    
    Row = START_OF_Row
    
	 Do While True
        ' Get the cell values from column E and column G     
        oCellE = oSheet.getCellRangeByName(RegisterName & Row)
        oCellG = oSheet.getCellRangeByName(RdORWr & Row)
        Sep_Reg_Size = oSheet.getCellRangeByName(RegSize & Row)
        sCellValueE = oCellE.getString()
        sCellValueG = oCellG.getString()
        sCellValueH = Sep_Reg_Size.getValue()
	    sCellStringH = Sep_Reg_Size.getString()
	    	    
	    If sCellStringH = "" Then
	    	ZeroPaddedReg = ""	
	    Else	        	
	    	ZeroPaddedReg =  PadDataReg(sCellValueH, BitWidth)        
	    End If 
	    
	    ZeroPaddedRegSize =	ZeroPaddedReg & sCellValueE
       
        ' Stop if the cell in column E is empty (assume end of entries)
        If sCellValueE = "" Then 
        	Row = START_OF_Row
        	C=0
        	sFormattedText = sFormattedText & "	     (others => '0');" & Chr(10) & Chr(10)
			j = 0 
        	Exit Do
        end if	

        ' Generate the signals based on the value in column G
        If sCellValueG = "R"  or sCellValueG = "RW" or sCellValueG = "" Then      	
        	if (Rd_Reg_cnt > SPLIT_VALUE) then
        	'******************************
	        	IF C = (SPLIT_VALUE-1) Then
	        		C=0
	        		'formattedSignalName = sCellValueE & "_reg " & String(maxSignalLength - Len(sCellValueE), " ")
	        		formattedSignalName = ZeroPaddedRegSize & "_reg " & String((SepRegMaxLen - Len(ZeroPaddedRegSize))-10, " ")
	        		sFormattedText =  sFormattedText & "	     " & formattedSignalName & " when "
	        		formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
	        		sFormattedText = sFormattedText & formattedSignalName &" = '1' else "& Chr(10) & "	     (others => '0');" & Chr(10) & Chr(10)
	        	Else
	        		C=C+1
	        		IF C=1 Then
        				sFormattedText = sFormattedText & "Rd_DATA_" & j+1 &" <= "  & Chr(10)
        				j=j+1
        			End If 
        			'formattedSignalName = sCellValueE & "_reg " & String((maxSignalLength) - Len(sCellValueE), " ")
        			formattedSignalName = ZeroPaddedRegSize & "_reg " & String((SepRegMaxLen - Len(ZeroPaddedRegSize))-10, " ")
        			sFormattedText =  sFormattedText & "	     " & formattedSignalName & " when "
        			formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
        			sFormattedText =  sFormattedText & formattedSignalName & " = '1' else " & Chr(10)
        		End if	
        	'******************************
        	Else
        		IF Row = START_OF_Row Then
        			sFormattedText = sFormattedText & RDO & " <= "  & Chr(10)
        		End if
        		formattedSignalName = ZeroPaddedRegSize & "_reg " & String((SepRegMaxLen - Len(ZeroPaddedRegSize))-10, " ")
        		'formattedSignalName = ZeroPaddedRegSize & "_reg " & String(maxSignalLength - Len(sCellValueE), " ")
             	sFormattedText =  sFormattedText & "	     " & formattedSignalName & " when "
             	formattedSignalName = sCellValueE & "_rd " & String(maxSignalLength - Len(sCellValueE), " ")
             	sFormattedText = sFormattedText & formattedSignalName &" = '1' else "& Chr(10)
           	End If
           	
        End if     
  		
        ' Move to the next row
        Row = Row + 1
    Loop

	 '-----------------Read_Enable_OR-----------------
	if (Rd_Reg_cnt > SPLIT_VALUE) then
	 	sFormattedText = sFormattedText & RDEO &" <= "
	 	Do While True
  			If j = j_temp Then 
  				j = 0
  				sFormattedText = sFormattedText &  Chr(10) & Chr(10)
  				Exit Do
  			end if
    			IF j =j_temp-1 THEN
  					sFormattedText = sFormattedText &  "Rd_EN_OUT_" & j+1 & ";"
  				ELSE
  					sFormattedText = sFormattedText &  "Rd_EN_OUT_" & j+1 & " or " 
  				END IF
    			j=j+1
    	Loop  	
    	
    else
    	sFormattedText = sFormattedText & RDEO &" <= "
   	End if
	j = 0 	
	 		 
	 '-----------------Read_Enable_Out-----------------
	 Do While True
        ' Get the cell values from column E and column G     
        oCellE = oSheet.getCellRangeByName(RegisterName & Row)
        oCellG = oSheet.getCellRangeByName(RdORWr & Row)
        sCellValueE = oCellE.getString()
        sCellValueG = oCellG.getString()

        ' Stop if the cell in column E is empty (assume end of entries)
        If sCellValueE = "" Then 
        	j = 0
        	sFormattedText = sFormattedText & Chr(10) 
        	Exit Do
        end if

        ' Generate the signals based on the value in column G
        If sCellValueG = "R"  or sCellValueG = "RW" or sCellValueG = "" Then
        
        	if (Rd_Reg_cnt > SPLIT_VALUE) then
        	'*********************************
            	IF C = (SPLIT_VALUE-1)   Then
	        		C=0
	        		sFormattedText = sFormattedText & sCellValueE & "_rd;" & Chr(10)
	        	Else
	        		C=C+1
	        		IF C=1 Then
        				sFormattedText = sFormattedText & "Rd_EN_OUT_" & j+1 &" <= "
        				j=j+1             					       			
        			End If 
        			If C=K and j=j_temp then
        				C=0
	        			sFormattedText = sFormattedText & sCellValueE & "_rd;" & Chr(10)
        			Else
        				sFormattedText =  sFormattedText & sCellValueE & "_rd or "
        			EndIf
        		End if	
        	'*********************************
            else
            	IF Rd_Reg_cnt =1 THEN
  					sFormattedText = sFormattedText & sCellValueE & "_rd; "  & Chr(10)
  				ELSE
  					sFormattedText = sFormattedText & sCellValueE & "_rd or "  			
  				END IF
    			Rd_Reg_cnt=Rd_Reg_cnt-1
            end if
            
        End If

        ' Move to the next row
        Row = Row + 1
    Loop
     
	'-----------------Write_DATA_LATCH_Process -----------------
	 Row = START_OF_Row
	
	 Do While True
        ' Get the cell values from column E and column G     
        oCellE = oSheet.getCellRangeByName(RegisterName & Row)
        oCellG = oSheet.getCellRangeByName(RdORWr & Row)   
        Sep_Reg_Size = oSheet.getCellRangeByName(RegSize & Row) 
        Default_V = oSheet.getCellRangeByName(ResetValue & Row)
        
        sCellValueE = oCellE.getString()
        sCellValueG = oCellG.getString()
        sCellValueH = Sep_Reg_Size.getValue()
        DefaultValueI = Default_V.getString()
        sCellValueH = Sep_Reg_Size.getValue()
	    sCellStringH = Sep_Reg_Size.getString()
	    
	    If sCellStringH = "" Then
	    	BitRange = ""
	    Else
	    	BitRange =  "(" & (sCellValueH - 1) & " downto 0)"
	    End If 


        ' Stop if the cell in column E is empty (assume end of entries)
        If sCellValueE = "" Then Exit Do

        ' Generate the signals based on the value in column G
        If sCellValueG = "RW" Or sCellValueG = "W" Or sCellValueG = ""Then
        	IF DefaultValueI = "" Then
            	sFormattedText =  sFormattedText &  "process("& RST &","& CLK &") " & Chr(10) & "begin" & Chr(10) & "    if rising_edge("& CLK &") then"& Chr(10) &"        if ("& RST &" = '1') then" & Chr(10) &"            "& sCellValueE &"_reg <= (others => '0');" & Chr(10) & "        elsif("& sCellValueE &"_wr = '1') then" &Chr(10) &"            "& sCellValueE &"_reg <= "& WRD & BitRange &";" & Chr(10) &"        end if;" & Chr(10) &"    end if;" & Chr(10) &"end process;" & Chr(10) & Chr(10)  
            ELSE ' it executes if default value fro write register is provided
            	 If Len(DefaultValueI) < BitWidth/4 Then
	       		 	DefaultValueI_SizeMatched = String((BitWidth/4) - Len(DefaultValueI), "0") & DefaultValueI 
	       		 ELSE
	       		 	DefaultValueI_SizeMatched = DefaultValueI
	   			 End If
            	sFormattedText =  sFormattedText &  "process("& RST &","& CLK &") " & Chr(10) & "begin" & Chr(10) & "    if rising_edge("& CLK &") then"& Chr(10) &"        if ("& RST &" = '1') then" & Chr(10) &"            "& sCellValueE &"_reg <= x""" & DefaultValueI_SizeMatched & """;" & Chr(10) & "        elsif("& sCellValueE &"_wr = '1') then" &Chr(10) &"            "& sCellValueE &"_reg <= "& WRD & BitRange &";" & Chr(10) &"        end if;" & Chr(10) &"    end if;" & Chr(10) &"end process;" & Chr(10) & Chr(10)	
            END IF     
        End If
    
        ' Move to the next row
        Row = Row + 1
    Loop

	'-----------------File Path Extraction, naming and openning the Text file -----------------
	
    ' Get the file path of the Calc file and set the output file name
    sFilePath = ConvertFromURL(oDoc.URL)
    If sFilePath = "" Then
        MsgBox "Error: Unable to determine file path of the Calc file.", 16, "Error"
        Exit Sub
    End If

    sFilePath = Left(sFilePath, Len(sFilePath) - Len(GetFileNameFromPath(sFilePath))) & "GeneratedFile.txt"
    sFilePath = ConvertToURL(sFilePath)  ' Convert to a LibreOffice-compatible URL

    ' Create the text file and write the data
    oSimpleFileAccess = CreateUnoService("com.sun.star.ucb.SimpleFileAccess")
    If oSimpleFileAccess.exists(sFilePath) Then
        oSimpleFileAccess.kill(sFilePath)
    End If

    oOutputStream = oSimpleFileAccess.openFileWrite(sFilePath)  ' Open the file for writing

    ' Write the text directly as bytes
    Dim ByteArray() As Byte
    ByteArray = StringToByteArray(sFormattedText)
    oOutputStream.writeBytes(ByteArray)
   ' oOutputStream.closeOutputStream()  ' Close the stream

    ' Notify the user
    MsgBox "Text file created successfully at " & ConvertFromURL(sFilePath), 64, "Success"
    
   
	sSystemFilePath = ConvertFromURL(sFilePath)
	
    ' Determine the OS
    isWindows = InStr(LCase(Environ("OS")), "windows") > 0

    ' Set the appropriate command for the OS
    If isWindows Then
        sCommand = "cmd /c start " & Chr(34) & sSystemFilePath & Chr(34)
    Else
        sCommand = "xdg-open " & Chr(34) & sSystemFilePath & Chr(34)
    End If

    ' Execute the command to open the file
    Shell(sCommand, 1)
   

End Sub


Function GetFileNameFromPath(FilePath As String) As String
    Dim i As Integer
    Dim FileName As String

    ' Ensure FilePath is not empty
    If FilePath = "" Then
        MsgBox "Error: FilePath is empty.", 16, "Error"
        Exit Function
    End If

    ' Loop through the file path to find the last slash
    For i = Len(FilePath) To 1 Step -1
        If Mid(FilePath, i, 1) = "/" Then
            FileName = Mid(FilePath, i + 1)
            Exit For
        End If
    Next i

    ' Return the file name
    GetFileNameFromPath = FileName
End Function

Function StringToByteArray(Text As String) As Variant
    Dim ByteArray() As Byte
    Dim i As Integer

    ' Convert the string into a byte array
    ReDim ByteArray(Len(Text) - 1)
    For i = 1 To Len(Text)
        ByteArray(i - 1) = Asc(Mid(Text, i, 1))
    Next i

    StringToByteArray = ByteArray
End Function

Function GetCellValueOrDefault(oSheet As Object, CellAddress As String, DefaultValue As String) As String
    On Error Resume Next
    Dim oCell As Object
    Dim CellValue As String
    
    ' Get the cell by address
    oCell = oSheet.getCellRangeByName(CellAddress)
    
    ' Fetch the cell's value as a string
    CellValue = oCell.getString()
    
    ' Check if the cell is empty or invalid
    If Trim(CellValue) = "" Then
        GetCellValueOrDefault = DefaultValue
    Else
        GetCellValueOrDefault = CellValue
    End If
End Function

Function PadDataReg(size As Integer, maxSize As Integer) As String
    Dim remaining As Integer
    Dim groupsOfFour As Integer
    Dim extraZeros As Integer
    Dim padding As String
    
    ' Ensure size and maxSize are provided
    If size <= 0 Or maxSize <= 0 Then
        MsgBox "Error: Invalid input values"
        Exit Function
    End If
    
    ' Calculate the number of bits needed to match max size
    remaining = maxSize - size
    
    ' Validate that maxSize is greater than size
    If remaining < 0 Then
        MsgBox "Error: maxSize must be greater than or equal to size"
        Exit Function
    End If
    
    ' Determine how many full groups of four fit in x"0000" format
    groupsOfFour = remaining \ 4 ' Integer division
    extraZeros = remaining Mod 4 ' Remaining zeros outside full groups
    
    ' Construct the padding using String() instead of Strings.StrDup()
     If groupsOfFour >= 1 Then
    	padding = "x""" & String(groupsOfFour, "0") & """"
     End If
    
    ' Add extra zeros if needed
    If extraZeros > 0 Then
        padding = padding & " & """ & String(extraZeros, "0") & """"
    End If
    
    ' Return only the padding
    PadDataReg = padding
End Function
