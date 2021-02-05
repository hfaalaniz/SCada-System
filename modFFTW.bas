Attribute VB_Name = "modFFTW"
Option Explicit

'FFTW (más rápida la Transformada de Fourier en el Oeste) contenedor para VB
'Las exportaciones no se puede llamar directamente, ya que la convención cdecl
'
' Arne Elster 2007 - rm_code

'#####################################################################
'#####################################################################

Private Declare Function LoadLibraryA Lib "kernel32" (ByVal strLib As String) As Long

Private Declare Function GetProcAddress Lib "kernel32" (ByVal hLib As Long, ByVal strProc As String) As Long

Private Declare Function GetModuleHandleA Lib "kernel32" (ByVal strMod As String) As Long

Private Declare Function VirtualAlloc Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long

Private Declare Function VirtualFree Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long

Private Declare Function VirtualLock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long

Private Declare Function VirtualUnlock Lib "kernel32" (ByVal lpAddress As Long, ByVal dwSize As Long) As Long

Private Declare Function CallWindowProcA Lib "User32" (ByVal pFnc As Long, ByVal arg1 As Long, ByVal arg2 As Long, ByVal arg3 As Long, ByVal arg4 As Long) As Long

Private Declare Sub CpyMem Lib "kernel32" Alias "RtlMoveMemory" (pDst As Any, pSrc As Any, ByVal cBytes As Long)

Private Const PAGE_EXECUTE_READWRITE    As Long = &H40
Private Const MEM_COMMIT                As Long = &H1000
Private Const MEM_DECOMMIT              As Long = &H4000

'#####################################################################
'#####################################################################

Private Const LIBFFTW                   As String = "libfftw3f.dll"

Public Enum FFTWPlan
    [_]
End Enum

Public Const FFTW_FORWARD               As Long = -1
Public Const FFTW_BACKWARD              As Long = 1

Public Const FFTW_NO_TIMELIMIT          As Long = -1

Public Const FFTW_MEASURE               As Long = &H0
Public Const FFTW_DESTROY_INPUT         As Long = &H1
Public Const FFTW_UNALIGNED             As Long = &H2
Public Const FFTW_CONSERVE_MEMORY       As Long = &H4
Public Const FFTW_EXHAUSTIVE            As Long = &H8
Public Const FFTW_PRESERVE_INPUT        As Long = &H10
Public Const FFTW_PATIENT               As Long = &H20
Public Const FFTW_ESTIMATE              As Long = &H40

Private Type FFTWFunction
    Name                                As String
    Address                             As Long
    Parameters                          As Long
    Size                                As Long
End Type

Private m_fncFFTW_Create_Plan_r2c_1d    As FFTWFunction
Private m_fncFFTW_Create_Plan_c2r_1d    As FFTWFunction
Private m_fncFFTW_Execute               As FFTWFunction
Private m_fncFFTW_Destroy_Plan          As FFTWFunction

'Public Const PI                         As Double = 3.14159265358979

Public Function HanningWindow(ByVal X As Long, ByVal Length As Long) As Double
    HanningWindow = 0.5 * (1 - Cos((2 * PI * X) / Length))
End Function

Public Function FFTW_Create_Plan_c2r_1d(ByVal n As Long, ByVal pRealIn As Long, ByVal pComplexOut As Long, ByVal flags As Long) As FFTWPlan

    With m_fncFFTW_Create_Plan_c2r_1d
        SetParam .Address, .Parameters, 1, n
        SetParam .Address, .Parameters, 2, pRealIn
        SetParam .Address, .Parameters, 3, pComplexOut
        SetParam .Address, .Parameters, 4, flags
        
        FFTW_Create_Plan_c2r_1d = CallWindowProcA(.Address, 0, 0, 0, 0)
    End With
End Function

Public Function FFTW_Create_Plan_r2c_1d(ByVal n As Long, ByVal pRealIn As Long, ByVal pComplexOut As Long, ByVal flags As Long) As FFTWPlan

    With m_fncFFTW_Create_Plan_r2c_1d
        SetParam .Address, .Parameters, 1, n
        SetParam .Address, .Parameters, 2, pRealIn
        SetParam .Address, .Parameters, 3, pComplexOut
        SetParam .Address, .Parameters, 4, flags
        
        FFTW_Create_Plan_r2c_1d = CallWindowProcA(.Address, 0, 0, 0, 0)
    End With
End Function

Public Sub FFTW_Execute(ByVal plan As FFTWPlan)
    With m_fncFFTW_Execute
        SetParam .Address, .Parameters, 1, plan
        CallWindowProcA .Address, 0, 0, 0, 0
    End With
End Sub

Public Sub FFTW_Destroy_Plan(ByVal plan As FFTWPlan)
    With m_fncFFTW_Destroy_Plan
        SetParam .Address, .Parameters, 1, plan
        CallWindowProcA .Address, 0, 0, 0, 0
    End With
End Sub

Public Sub FFTWInit()
    CreateFFTWFunction "fftwf_plan_dft_r2c_1d", 4, m_fncFFTW_Create_Plan_r2c_1d
    CreateFFTWFunction "fftwf_plan_dft_c2r_1d", 4, m_fncFFTW_Create_Plan_c2r_1d
    CreateFFTWFunction "fftwf_execute", 1, m_fncFFTW_Execute
    CreateFFTWFunction "fftwf_destroy_plan", 1, m_fncFFTW_Destroy_Plan
End Sub

Public Sub FFTWTerm()
    DestroyFFTWFunction m_fncFFTW_Create_Plan_r2c_1d
    DestroyFFTWFunction m_fncFFTW_Create_Plan_c2r_1d
    DestroyFFTWFunction m_fncFFTW_Execute
    DestroyFFTWFunction m_fncFFTW_Destroy_Plan
End Sub

Private Sub DestroyFFTWFunction(ByRef fnc As FFTWFunction)
    VirtualUnlock fnc.Address, fnc.Size
    VirtualFree fnc.Address, fnc.Size, MEM_DECOMMIT
End Sub

' Stdcall función de contenedor con la convención de llamada Cdecl
Private Sub CreateFFTWFunction(ByVal Name As String, ByVal params As Long, ByRef fnc As FFTWFunction)

    Dim pMem        As Long
    Dim pFnc        As Long
    Dim lngAsmSize  As Long
    Dim pAsm        As Long
    Dim i           As Long
    
    pFnc = GetProcAddressEx(LIBFFTW, Name)
    If pFnc = 0 Then Err.Raise 600, , "Biblioteca o la exportación no se encuentra!" '
    
    ' Memoria necesaria en bytes
    lngAsmSize = 5 * params + 16
    
    pMem = VirtualAlloc(0, lngAsmSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    VirtualLock pMem, lngAsmSize
    
    pAsm = pMem
    
    AddByte pAsm, &H58                      ' POP EAX
    AddByte pAsm, &H59                      ' POP ECX
    AddByte pAsm, &H59                      ' POP ECX
    AddByte pAsm, &H59                      ' POP ECX
    AddByte pAsm, &H59                      ' POP ECX
    AddByte pAsm, &H50                      ' PUSH EAX
    
    For i = 0 To params - 1
        AddPush pAsm, 0                     ' PUSH 0
    Next
    
    AddCall pAsm, pFnc                      ' CALL pFnc
    AddByte pAsm, &H83                      ' ADD ESP, ArgCount*4
    AddByte pAsm, &HC4
    AddByte pAsm, 4 * params
    AddByte pAsm, &HC3                      ' RET
    AddByte pAsm, &H0
    
    With fnc
        .Name = Name
        .Parameters = params
        .Address = pMem
        .Size = lngAsmSize
    End With
End Sub

Private Sub SetParam(ByVal pAsm As Long, ByVal params As Long, ByVal param As Long, ByVal value As Long)
    CpyMem ByVal pAsm + 7 + (params - param) * 5, value, 4
End Sub

Private Function GetProcAddressEx(ByVal strLib As String, ByVal strFnc As String) As Long
    Dim hMod    As Long
    
    hMod = GetModuleHandleA(strLib)
    If hMod = 0 Then hMod = LoadLibraryA(strLib)
    If hMod = 0 Then Exit Function
    
    GetProcAddressEx = GetProcAddress(hMod, strFnc)
End Function

Private Sub AddPush(pAsm As Long, lng As Long)
    AddByte pAsm, &H68
    AddLong pAsm, lng
End Sub

Private Sub AddCall(pAsm As Long, addr As Long)
    AddByte pAsm, &HE8
    AddLong pAsm, addr - pAsm - 4
End Sub

Private Sub AddLong(pAsm As Long, lng As Long)
    CpyMem ByVal pAsm, lng, 4
    pAsm = pAsm + 4
End Sub

Private Sub AddByte(pAsm As Long, Bt As Byte)
    CpyMem ByVal pAsm, Bt, 1
    pAsm = pAsm + 1
End Sub
