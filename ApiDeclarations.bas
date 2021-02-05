Attribute VB_Name = "ApiDeclarations"
'******************************************************************************
'API constants, listed alphabetically
'******************************************************************************
'from setupapi.h
Public Const DIGCF_PRESENT = &H2
Public Const DIGCF_DEVICEINTERFACE = &H10
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
Public Const GENERIC_READ = &H80000000
Public Const GENERIC_WRITE = &H40000000
'Typedef enum defines a set of integer constants for HidP_Report_Type
'Remember to declare these as integers (16 bits)
Public Const HidP_Input = 0
Public Const HidP_Output = 1
Public Const HidP_Feature = 2

Public Const OPEN_EXISTING = 3
Public Const WAIT_TIMEOUT = &H102&
Public Const WAIT_OBJECT_0 = 0
'******************************************************************************
'User-defined types for API calls, listed alphabetically
'******************************************************************************
Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Type HIDD_ATTRIBUTES
    Size As Long
    VendorID As Integer
    ProductID As Integer
    VersionNumber As Integer
End Type
'Windows 98 DDK documentation is incomplete.
'Use the structure defined in hidpi.h
Public Type HIDP_CAPS
    Usage As Integer
    UsagePage As Integer
    InputReportByteLength As Integer
    OutputReportByteLength As Integer
    FeatureReportByteLength As Integer
    reserved(16) As Integer
    NumberLinkCollectionNodes As Integer
    NumberInputButtonCaps As Integer
    NumberInputValueCaps As Integer
    NumberInputDataIndices As Integer
    NumberOutputButtonCaps As Integer
    NumberOutputValueCaps As Integer
    NumberOutputDataIndices As Integer
    NumberFeatureButtonCaps As Integer
    NumberFeatureValueCaps As Integer
    NumberFeatureDataIndices As Integer
End Type
'If IsRange is false, UsageMin is the Usage and UsageMax is unused.
'If IsStringRange is false, StringMin is the string index and StringMax is unused.
'If IsDesignatorRange is false, DesignatorMin is the designator index and DesignatorMax is unused.
Public Type HidP_Value_Caps
    UsagePage As Integer
    ReportID As Byte
    IsAlias As Long
    BitField As Integer
    LinkCollection As Integer
    LinkUsage As Integer
    LinkUsagePage As Integer
    IsRange As Long
    IsStringRange As Long
    IsDesignatorRange As Long
    IsAbsolute As Long
    HasNull As Long
    reserved As Byte
    BitSize As Integer
    ReportCount As Integer
    Reserved2 As Integer
    Reserved3 As Integer
    Reserved4 As Integer
    Reserved5 As Integer
    Reserved6 As Integer
    LogicalMin As Long
    LogicalMax As Long
    PhysicalMin As Long
    PhysicalMax As Long
    UsageMin As Integer
    UsageMax As Integer
    StringMin As Integer
    StringMax As Integer
    DesignatorMin As Integer
    DesignatorMax As Integer
    DataIndexMin As Integer
    DataIndexMax As Integer
End Type

Public Type OVERLAPPED
    Internal As Long
    InternalHigh As Long
    Offset As Long
    OffsetHigh As Long
    hEvent As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Type SP_DEVICE_INTERFACE_DATA
   cbSize As Long
   InterfaceClassGuid As GUID
   flags As Long
   reserved As Long
End Type

Public Type SP_DEVICE_INTERFACE_DETAIL_DATA
    cbSize As Long
    DevicePath As Byte
End Type

Public Type SP_DEVINFO_DATA
    cbSize As Long
    ClassGuid As GUID
    DevInst As Long
    reserved As Long
End Type
'******************************************************************************
'API functions, listed alphabetically
'******************************************************************************
Public Declare Function CancelIo Lib "kernel32" (ByVal hFile As Long) As Long
Public Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Public Declare Function CreateEvent Lib "kernel32" Alias "CreateEventA" (ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal bManualReset As Long, ByVal bInitialState As Long, ByVal lpName As String) As Long
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByRef lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, ByRef lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageZId As Long, ByVal lpBuffer As String, ByVal nSize As Long, ByVal Arguments As Long) As Long
Public Declare Function HidD_FreePreparsedData Lib "hid.dll" (ByRef PreparsedData As Long) As Long
Public Declare Function HidD_GetAttributes Lib "hid.dll" (ByVal HidDeviceObject As Long, ByRef Attributes As HIDD_ATTRIBUTES) As Long
'Declared as a function for consistency, but returns nothing. (Ignore the returned value.)
Public Declare Function HidD_GetHidGuid Lib "hid.dll" (ByRef HidGuid As GUID) As Long
Public Declare Function HidD_GetPreparsedData Lib "hid.dll" (ByVal HidDeviceObject As Long, ByRef PreparsedData As Long) As Long
Public Declare Function HidP_GetCaps Lib "hid.dll" (ByVal PreparsedData As Long, ByRef Capabilities As HIDP_CAPS) As Long
Public Declare Function HidP_GetValueCaps Lib "hid.dll" (ByVal ReportType As Integer, ByRef ValueCaps As Byte, ByRef ValueCapsLength As Integer, ByVal PreparsedData As Long) As Long
Public Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal dest As String, ByVal Source As Long) As String
Public Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal Source As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As OVERLAPPED) As Long
Public Declare Function ResetEvent Lib "kernel32" (ByVal hEvent As Long) As Long
Public Declare Function RtlMoveMemory Lib "kernel32" (dest As Any, src As Any, ByVal Count As Long) As Long
Public Declare Function SetupDiCreateDeviceInfoList Lib "setupapi.dll" (ByRef ClassGuid As GUID, ByVal hWndParent As Long) As Long
Public Declare Function SetupDiDestroyDeviceInfoList Lib "setupapi.dll" (ByVal DeviceInfoSet As Long) As Long
Public Declare Function SetupDiEnumDeviceInterfaces Lib "setupapi.dll" (ByVal DeviceInfoSet As Long, ByVal DeviceInfoData As Long, ByRef InterfaceClassGuid As GUID, ByVal MemberIndex As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA) As Long
Public Declare Function SetupDiGetClassDevs Lib "setupapi.dll" Alias "SetupDiGetClassDevsA" (ByRef ClassGuid As GUID, ByVal Enumerator As String, ByVal hWndParent As Long, ByVal flags As Long) As Long
Public Declare Function SetupDiGetDeviceInterfaceDetail Lib "setupapi.dll" Alias "SetupDiGetDeviceInterfaceDetailA" (ByVal DeviceInfoSet As Long, ByRef DeviceInterfaceData As SP_DEVICE_INTERFACE_DATA, ByVal DeviceInterfaceDetailData As Long, ByVal DeviceInterfaceDetailDataSize As Long, ByRef RequiredSize As Long, ByVal DeviceInfoData As Long) As Long
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, ByRef lpBuffer As Byte, ByVal nNumberOfBytesToWrite As Long, ByRef lpNumberOfBytesWritten As Long, ByVal lpOverlapped As Long) As Long
'Declaración Api SetParent para incrustar el Progressbar en la barra de estado
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
'-------  Shape Controls
Public Declare Function CreateRoundRectRgn Lib "gdi32.dll" (ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long, ByVal X3 As Long, ByVal Y3 As Long) As Long
' X1,Y1  X2,Y2  Top left & Bottom right coords of rectangle.
' For whole control X1 & Y1 = 0
' X2 & Y2 = Controls width & height
' X3,Y3  width & height of ellipse used to create corners
Public Declare Function SetWindowRgn Lib "user32.dll" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
Public Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
'-------


 
