# VBA_Learning

    Benefits of VBA: Saving time, Automate tasks, Reduce errors, Interact with other software.
    
    Basic theory:
      Excel Object -> <VBA> -> Excel Object
      
      Excel itself is an object, the largest object in Excel, represented by Application. 
      The Application object contains the Workbook object, the Workbook object contains the Worksheet object, 
      and the Worksheet object contains other child objects.
   
      Object model is a tree structure
      
      Common Excel objects:
        Application, Workbook, Worksheet, Range
  
  
  Rules:
    Notes:
  
      '                               -> notes (equals // in java)
    Variables:
    
     Variables naming rule is similar to Java
     
     Types:
     
      String, Boolean, Date, Object, Integer (Byte, Long, Single [single precision floating point number], 
      Double [double precision floating point number], Currency, Decimal), 
      Variant (any type, but may take up more memory)
      
      Date              While assign value, can use # # OR " " OR NUMBERS
  
     Dim VARIABLE As TYPE             -> define a variable (Dim s As String
                                                            S = "Hello World")
  
     Range("CELL").Value              -> fill the variable in one cell (Range("A1").Value = s 'represents store value s in the A1 cell)
     
     Assign variables is the same as Java or any other language
     
     Const VARIABLES As TYPE = VALUE  -> declare a constant (Same idea with other programming language)
	
  
    Operators:
      
       +
       -   
       *   
       /        divide and answer with decimal
       \        divide and answer without decimal
       Mod    
       ^  
       -        negative
       ->       result
       =
       >
       >=
       <
       <=
       <>       not equal
       And      Both true           return true
       Or       At least one true   return true
       Not      not true            return false
       Xor      Not same equation   return true
       &
       :        put two lines of codes in the same line
       _        separate one line code into two lines
       
    Structures:
      
      Sequential:  Line by line

      Conditional: If - Then
                   Sample:
                       If EXPRESSION Then
                           ........
                       End If

		  If - Then
		  Else
		  Sample:
		      If EXPRESSION Then
			 ........
		      Else
			 ........
		      End If

		  If - ElseIf
		  Else
		  Sample:
		      If EXPRESSION#1 Then
			 ........
		      ElseIf EXPRESSION#2 Then
			 ........
		      ElseIf EXPRESSION#3 Then
			 ........
		      Else
			 ........
		      End If

		  Select Case
		  Sample:
		      Select Case VARIABLE
		      	  Case CONDITION 1
			    .......
			  Case CONDITION 2
			    .......
			Case Else
			  .......
		       End Select
  
      Loop:        Loop

                   For Next 
		  Sample:
                       For VARIABLE = INITIAL VALUE To END VALUE Step ADD VALUE
				(If ADD VALUE is 1, then we can ignore Step)
                          .......
                       Next

		  For Each
		  Sample:
		      For Each ELEMENT In SET
		        .......
		      Next ELEMENT

		  Exit For (Like Break)
		  Sample:
		      If EXPRESSION Then
			Exit For
		      .......
		      
		  Do While:
			Do While ... Loop
			Sample:
			   Do While CONDITION
			    ......
			   Loop

			Do .... While Loop
			Sample:
			   Do
			     .....
			   Loop While EXPRESSION

			Exit Do (same as break)

		  Do Until:
			Do Until ... Loop
			Sample:
			    Do Until EXPRESSION
			      ......
			    Loop

			Do ... Loop Until
			Sample:
			    Do
			      ......
			    Loop Until EXPRESSION

      With (Avoid writing the same object's name, enhance efficiency):

        With OBJECT
	    .ATTRIBUTE = DATA
	    .METHOD
	    .OTHER METHOD / ATTRIBUTES
      
        Nested Structure: In a With structure, if the property of the parent object is another object, 
	  continue to use the With structure for this child object
			
      GoTo (Jump to the specified label to run, 
	so that the code between the GoTo statement and the specified label is not executed):
	
        Sample:
        GoTo SIGN
        ...SKIP CODES...
        SIGN:
        ...RUN CODES...

        Note: GoTo is more like debug tools or dealing with error
           
  
    Arrays:
      
      s(1)     -> similar thing as java (Dims s(1 to 4) As String
                                         S(1) = "Word"
                                         S(2) = "Excel"    ......)
    
    Text Functions:
    
      Format              Format the data and return it as text
      InStr               Returns the position of the specified character
      InStrRev            Return to the specified character position in the reverse direction
      
      Mid                 Returns the text between the specified start and end positions
      Right               Returns the text of the specified length on the right
      Left                Returns the text of the specified length on the left
      Replace             Replace specified characters in text
      
      Space               Returns whitespace text with the specified number of repetitions
      
      String
      Len                 Returns text length
      StrComp             Compare two Strings
      StrConv             Convert text to specified format  
      StrReverse          Reverse the provided string
      
      Trim                Clear text at beginning and end
      Ltrim               Clear leading spaces
      RTrim               Clear spaces at the end
      
      LCase               UpperCase to LowerCase
      UCase               LowerCase to UpperCase
  
    Objects: contain properties that describe static information and methods that can operate on objects.
    
    Procedure Basic Theory:
      
      Sub:

        No Parameter Procedure:
	  Sample:
	  Sub PRECEDURE NAME ()
           .......
	  End Sub

        With Parameter Procedure:
	  Sample:
	  Sub PRECEDURE NAME (VARIABLE 1 As DATA TYPE, .... VARIABLE N As DATA TYPE)
	  ......
	  End Sub

	Call Sub Function:
	  Same as other programming languages, can call in the Main function
	  If a Sub Function have multiple parameters, just use , to seperate:
		Sample:
		Sub Main()
		   Sub1 2022, "Year"
		End Sub
	  
          Otherwise, we can use Call to call sub function
          The only difference will be we need to put parameters in ( )
	        Sample:
	        Sub Main()
 		   Call Sub1(2022, "Year")
	        End Sub
		
	Exit Sub: Like break to quit the current procedure

	End: Kill all current running VBA processes

      Function:
	
	No Parameter Function:
	   Sample:
	   Function FUNCTION NAME() As RETURN VALUE TYPE
	   ......
	   End Function

	With Parameter Function
	   Sample:
	   Function FUNCTION NAME (VARIABLE NAME 1 As DATA TYPE, 
		... VARIABLE NAME N As DATA TYPE N) As RETURN VALUE DATA TYPE
	   ......
	   End Function

	Call Function:
	   If a function not return a value, we can just call like call sub
	   Otherwise, we can call the return value directly.

	Exit Function: Same as other Exit functions
