# VBA_Learning

    Benfits of VBA: Saving time, Automate tasks, Reduce errors, Interact with other software.
    
    Basic theroy:
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
       Or       At least one true   return ture
       Not      not true            return false
       Xor      Not same equation   return true
       &
       :        put two lines of codes in the same line
       _        seperate one line code into two lines
       
    Structures:
      
      Sequential:  Line by line

      Conditional: If - else
                   Sample:
                       If XXXX Then
                           ........
                       End If
  
      Loop:        Loop
                   For loop sample:
                        For i = 1 To 10
                            .......
                        Next i
  
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
    
    Codes happens in Sub structure:
                   Sub MyCode()
                   ........
                   End Sub
