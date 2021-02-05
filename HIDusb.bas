Attribute VB_Name = "HIDusb"
Option Explicit
'Proposito: genera la comunicacion USB con el dispositivo HID-class.

'Juego de VID, PID para que coincida con los valores en el firmware del dispositivo.

Const MyVendorID = 6017   '&H4D8
Const MyProductID = 2000  '&H3F

Dim bAlertable As Long
Dim Capabilities As HIDP_CAPS
Dim DataString As String
Dim DetailData As Long
Dim DetailDataBuffer() As Byte
Dim DeviceAttributes As HIDD_ATTRIBUTES
Dim DevicePathName As String
Dim DeviceInfoSet As Long
Dim ErrorString As String
Dim EventObject As Long
Public HIDHandle As Long
Dim HIDOverlapped As OVERLAPPED
Dim LastDevice As Boolean
Public MyDeviceDetected As Boolean
Dim MyDeviceInfoData As SP_DEVINFO_DATA
Dim MyDeviceInterfaceDetailData As SP_DEVICE_INTERFACE_DETAIL_DATA
Dim MyDeviceInterfaceData As SP_DEVICE_INTERFACE_DATA
Dim Needed As Long
Public OutputReportData(8) As Byte
Dim PreparsedData As Long
Public ReadHandle As Long
Dim Result As Long
Dim Security As SECURITY_ATTRIBUTES
Public ReadBuffer() As Byte

Function FindTheHid() As Boolean

'Hace una serie de llamadas a la API para localizar el dispositivo de clase HID.
'Devuelve True si el dispositivo es detectado, False si no se detecta.

Dim Count As Integer
Dim GUIDString As String
Dim HidGuid As GUID
Dim MemberIndex As Long
   'On Error GoTo Err_Proc
LastDevice = False
MyDeviceDetected = False
'Los valores de estructura SECURITY_ATTRIBUTES:
Security.lpSecurityDescriptor = 0
Security.bInheritHandle = True
Security.nLength = Len(Security)
'******************************************************************************
'HidD_GetHidGuid
'Obtener el GUID para todos los HID del sistema.
'Devuelve: el GUID en HidGuid.
'La rutina no devuelve un valor en el resultado
'Pero la rutina se declara como una función de coherencia con las llamadas a la API otros.
'******************************************************************************
Result = HidD_GetHidGuid(HidGuid)
'Display the GUID.
GUIDString = _
    Hex$(HidGuid.Data1) & "-" & _
    Hex$(HidGuid.Data2) & "-" & _
    Hex$(HidGuid.Data3) & "-"
For Count = 0 To 7
    'Ensure that each of the 8 bytes in the GUID displays two characters.
    If HidGuid.Data4(Count) >= &H10 Then
        GUIDString = GUIDString & Hex$(HidGuid.Data4(Count)) & " "
    Else
        GUIDString = GUIDString & "0" & Hex$(HidGuid.Data4(Count)) & " "
    End If
Next Count
'******************************************************************************
'SetupDiGetClassDevs
'Devuelve: un identificador a una información del dispositivo establecido para todos los dispositivos instalados.
'Se requiere: La HidGuid devuelto en GetHidGuid.
'******************************************************************************
DeviceInfoSet = SetupDiGetClassDevs(HidGuid, vbNullString, 0, (DIGCF_PRESENT Or DIGCF_DEVICEINTERFACE))
DataString = GetDataString(DeviceInfoSet, 32)
'******************************************************************************
'SetupDiEnumDeviceInterfaces
'En cambio, contiene el identificador de un MyDeviceInterfaceData
'SP_DEVICE_INTERFACE_DATA estructura de un dispositivo detectado.
'Se requiere:
'El DeviceInfoSet devuelto en SetupDiGetClassDevs.
'El HidGuid devuelto en GetHidGuid.
'Un índice para especificar un dispositivo.
'************************************************* *****************************
'Comience con 0 y el incremento hasta que los dispositivos no más se detectan.
MemberIndex = 0
Do
    'El elemento cbSize de la estructura se debe establecer en MyDeviceInterfaceData 'La estructura de tamaño en bytes. El tamaño es de 28 bytes."
    MyDeviceInterfaceData.cbSize = LenB(MyDeviceInterfaceData)
    Result = SetupDiEnumDeviceInterfaces(DeviceInfoSet, 0, HidGuid, MemberIndex, MyDeviceInterfaceData)
    If Result = 0 Then LastDevice = True
        'If a device exists, display the information returned.
    If Result <> 0 Then
            '******************************************************************************
            'SetupDiGetDeviceInterfaceDetail'''Estructura de SP_DEVICE_INTERFACE_DETAIL_DATA: Returns ''que contiene información sobre un dispositivo.
'Para recuperar la información, llamar a esta función dos veces."
'La primera vez que devuelve el tamaño de la estructura en necesaria."
'La segunda vez que devuelve un puntero a los datos de DeviceInfoSet."
'Se requiere:"
'DeviceInfoSet devuelto por SetupDiGetClassDevs y"
'Para la estructura SP_DEVICE_INTERFACE_DATA devuelto por SetupDiEnumDeviceInterfaces."
            '*******************************************************************************
            MyDeviceInfoData.cbSize = Len(MyDeviceInfoData)
            Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, 0, 0, Needed, 0)
            DetailData = Needed
            'Guarde el tamaño de la estructura.
            MyDeviceInterfaceDetailData.cbSize = Len(MyDeviceInterfaceDetailData)
            'Utilice una matriz de bytes para asignar memoria para
            'La estructura de MyDeviceInterfaceDetailData
            ReDim DetailDataBuffer(Needed)
            'Guarde cbSize en los primeros cuatro bytes de la matriz.
            Call RtlMoveMemory(DetailDataBuffer(0), MyDeviceInterfaceDetailData, 4)
            'Llama SetupDiGetDeviceInterfaceDetail de nuevo.
            'Esta vez, pasar la dirección del primer elemento de DetailDataBuffer
            'y devuelve el tamaño requerido en el buffer DetailData.
            Result = SetupDiGetDeviceInterfaceDetail(DeviceInfoSet, MyDeviceInterfaceData, VarPtr(DetailDataBuffer(0)), DetailData, Needed, 0)
            'Convertir la matriz de bytes en una cadena.
            DevicePathName = CStr(DetailDataBuffer())
            'Convert to Unicode.
            DevicePathName = StrConv(DevicePathName, vbUnicode)
            'Pele cbSize (4 bytes) desde el principio.
            DevicePathName = Right$(DevicePathName, Len(DevicePathName) - 4)
            '******************************************************************************
            'CreateFile
            'Returns: a handle that enables reading and writing to the device.
            'Requires:
            'The DevicePathName returned by SetupDiGetDeviceInterfaceDetail.
            '******************************************************************************
            HIDHandle = CreateFile(DevicePathName, GENERIC_READ Or GENERIC_WRITE, (FILE_SHARE_READ Or FILE_SHARE_WRITE), Security, OPEN_EXISTING, 0&, 0)
            'Now we can find out if it's the device we're looking for.
            '******************************************************************************
            'los HidD_GetAttributes
            'Las solicitudes de información desde el dispositivo.
            '"Se requiere: El identificador devuelto por CreateFile.
            'Devuelve: una estructura que contiene HIDD_ATTRIBUTES
            '"el Vendor ID, identificador de producto y número de versión del producto.
            '"Utilice esta información para determinar si el dispositivo detecta
            '"es la que estamos buscando.
            '******************************************************************************
            'Establezca la propiedad Size en el número de bytes en la estructura.
            DeviceAttributes.Size = LenB(DeviceAttributes)
            Result = HidD_GetAttributes(HIDHandle, DeviceAttributes)
            If (DeviceAttributes.VendorID = MyVendorID) And (DeviceAttributes.ProductID = MyProductID) Then
                    'Es el dispositivo que desee.
                    frmPrincipal.Label1.Caption = "Device Detected"
                    MyDeviceDetected = True
            Else
                    MyDeviceDetected = False
                    frmPrincipal.Label1.Caption = "No USB Device"
                    'Detection = "No USB Device"
                    'Si no es el que queremos, cerrar el manejador.
                    Result = CloseHandle(HIDHandle)
                    'DisplayResultOfAPICall ("CloseHandle")
            End If
    End If
    'Sigue buscando hasta que encontremos el dispositivo o no hay más para examinar.
    MemberIndex = MemberIndex + 1
Loop Until (LastDevice = True) Or (MyDeviceDetected = True)
'Liberar la memoria reservada para el DeviceInfoSet devuelto por SetupDiGetClassDevs.
Result = SetupDiDestroyDeviceInfoList(DeviceInfoSet)
If MyDeviceDetected = True Then
    FindTheHid = True
    'Conozca las capacidades del dispositivo
     Call GetDeviceCapabilities
    'Obtener otro manejador para los archivos de lectura superpuestos.
    ReadHandle = CreateFile(DevicePathName, (GENERIC_READ Or GENERIC_WRITE), (FILE_SHARE_READ Or FILE_SHARE_WRITE), Security, OPEN_EXISTING, FILE_FLAG_OVERLAPPED, 0)
    Call PrepareForOverlappedTransfer
Else
    'lstResults.AddItem " Device not found."
End If

Exit_Proc:
   Exit Function
Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "FindTheHid"
   Err.Clear
   Resume Exit_Proc
End Function

Private Function GetDataString(Address As Long, Bytes As Long) As String
'Recupera una cadena de bytes de longitud de la memoria, a partir de Dirección.
'Adaptado de Dan Appleman de" Puzzle Book API Win32 "
Dim Offset As Integer

Dim Result$
Dim ThisByte As Byte
For Offset = 0 To Bytes - 1
    Call RtlMoveMemory(ByVal VarPtr(ThisByte), ByVal Address + Offset, 1)
   'On Error GoTo Err_Proc
    If (ThisByte And &HF0) = 0 Then
        Result$ = Result$ & "0"
    End If
    Result$ = Result$ & Hex$(ThisByte) & " "
Next Offset

GetDataString = Result$

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "GetDataString"
   Err.Clear
   Resume Exit_Proc

End Function

Private Function GetErrorString(ByVal LastError As Long) As String

'Returns the error message for the last error.
'Adapted from Dan Appleman's "Win32 API Puzzle Book"

Dim Bytes As Long
Dim ErrorString As String
ErrorString = String$(129, 0)
   'On Error GoTo Err_Proc
Bytes = FormatMessage(FORMAT_MESSAGE_FROM_SYSTEM, 0&, LastError, 0, ErrorString$, 128, 0)
'Subtract two characters from the message to strip the CR and LF.
If Bytes > 2 Then
    GetErrorString = Left$(ErrorString, Bytes - 2)
End If

Exit_Proc:
   Exit Function

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "GetErrorString"
   Err.Clear
   Resume Exit_Proc

End Function

'Private Sub DisplayResultOfAPICall(FunctionName As String)

'Display the results of an API call.

'Dim ErrorString As String

'ErrorString = GetErrorString(Err.LastDllError)
'Scroll to the bottom of the list box.

'End Sub


Private Sub Form_Unload(Cancel As Integer)
Call Shutdown
End Sub

Private Sub GetDeviceCapabilities()

'******************************************************************************
'HidD_GetPreparsedData
'Returns: a pointer to a buffer containing information about the device's capabilities.
'Requires: A handle returned by CreateFile.
'There's no need to access the buffer directly,
'but HidP_GetCaps and other API functions require a pointer to the buffer.
'******************************************************************************

Dim ppData(29) As Byte
Dim ppDataString As Variant

'Preparsed Data is a pointer to a routine-allocated buffer.

   'On Error GoTo Err_Proc

Result = HidD_GetPreparsedData _
    (HIDHandle, _
    PreparsedData)

'Copy the data at PreparsedData into a byte array.

Result = RtlMoveMemory _
    (ppData(0), _
    PreparsedData, _
    30)
ppDataString = ppData()

'Convert the data to Unicode.

ppDataString = StrConv(ppDataString, vbUnicode)

'******************************************************************************
'HidP_GetCaps
'Find out the device's capabilities.
'For standard devices such as joysticks, you can find out the specific
'capabilities of the device.
'For a custom device, the software will probably know what the device is capable of,
'so this call only verifies the information.
'Requires: The pointer to a buffer containing the information.
'The pointer is returned by HidD_GetPreparsedData.
'Returns: a Capabilites structure containing the information.
'******************************************************************************
Result = HidP_GetCaps _
    (PreparsedData, _
    Capabilities)

'******************************************************************************
'HidP_GetValueCaps
'Returns a buffer containing an array of HidP_ValueCaps structures.
'Each structure defines the capabilities of one value.
'This application doesn't use this data.
'******************************************************************************

'This is a guess. The byte array holds the structures.

Dim ValueCaps(1023) As Byte

Result = HidP_GetValueCaps _
    (HidP_Input, _
    ValueCaps(0), _
    Capabilities.NumberInputValueCaps, _
    PreparsedData)
   
'To use this data, copy the byte array into an array of structures.

'Free the buffer reserved by HidD_GetPreparsedData

Result = HidD_FreePreparsedData _
    (PreparsedData)


Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "GetDeviceCapabilities"
   Err.Clear
   Resume Exit_Proc

End Sub



Private Sub PrepareForOverlappedTransfer()

'******************************************************************************
'CreateEvent
'Creates an event object for the overlapped structure used with ReadFile.
'Requires a security attributes structure or null,
'Manual Reset = True (ResetEvent resets the manual reset object to nonsignaled),
'Initial state = True (signaled),
'and event object name (optional)
'Returns a handle to the event object.
'******************************************************************************

   'On Error GoTo Err_Proc

If EventObject = 0 Then
    EventObject = CreateEvent _
        (Security, _
        True, _
        True, _
        "")
End If
    
'Call DisplayResultOfAPICall("CreateEvent")
    
'Set the members of the overlapped structure.

HIDOverlapped.Offset = 0
HIDOverlapped.OffsetHigh = 0
HIDOverlapped.hEvent = EventObject

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "PrepareForOverlappedTransfer"
   Err.Clear
   Resume Exit_Proc

End Sub

Public Sub ReadAndWriteToDevice()

'Sends two bytes to the device and reads two bytes back.

Dim Count As Integer

   'On Error GoTo Err_Proc

If MyDeviceDetected = False Then
    MyDeviceDetected = FindTheHid
    
End If

If MyDeviceDetected = True Then

    'OutputReportData(0) = 0  'cboByte0.ListIndex
    'OutputReportData(1) = 1  'cboByte1.ListIndex
    Call WriteReport
    Call ReadReport
Else
End If

Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "ReadAndWriteToDevice"
   Err.Clear
   Resume Exit_Proc

End Sub

Private Sub ReadReport()
'On Error Resume Next

'Read data from the device.

'Dim Count
Dim NumberOfBytesRead As Long

'Allocate a buffer for the report.
'Byte 0 is the report ID.

'Dim ReadBuffer() As Byte
Dim UBoundReadBuffer As Integer

'******************************************************************************
'ReadFile
'Returns: the report in ReadBuffer.
'Requires: a device handle returned by CreateFile
'(for overlapped I/O, CreateFile must be called with FILE_FLAG_OVERLAPPED),
'the Input report length in bytes returned by HidP_GetCaps,
'and an overlapped structure whose hEvent member is set to an event object.
'******************************************************************************

Dim ByteValue As String

'The ReadBuffer array begins at 0, so subtract 1 from the number of bytes.

ReDim ReadBuffer(Capabilities.InputReportByteLength - 1)

'Scroll to the bottom of the list box.

'Do an overlapped ReadFile.
'The function returns immediately, even if the data hasn't been received yet.

Result = ReadFile _
    (ReadHandle, _
    ReadBuffer(0), _
    CLng(Capabilities.InputReportByteLength), _
    NumberOfBytesRead, _
    HIDOverlapped)
'Scroll to the bottom of the list box.

bAlertable = True

'******************************************************************************
'WaitForSingleObject
'Used with overlapped ReadFile.
'Returns when ReadFile has received the requested amount of data or on timeout.
'Requires an event object created with CreateEvent
'and a timeout value in milliseconds.
'******************************************************************************
Result = WaitForSingleObject _
    (EventObject, _
    6000)

'Find out if ReadFile completed or timeout.

Select Case Result
    Case WAIT_OBJECT_0
        
        'ReadFile has completed
        
    Case WAIT_TIMEOUT
        
        'Timeout
        'Cancel the operation
        
        '*************************************************************
        'CancelIo
        'Cancels the ReadFile
        'Requires the device handle.
        'Returns non-zero on success.
        '*************************************************************
        Result = CancelIo _
            (ReadHandle)
 
        'The timeout may have been because the device was removed,
        'so close any open handles and
        'set MyDeviceDetected=False to cause the application to
        'look for the device on the next attempt.
        
        CloseHandle (HIDHandle)
        CloseHandle (ReadHandle)
        MyDeviceDetected = False
    Case Else
         MyDeviceDetected = False
End Select
    
'For Count = 1 To UBound(ReadBuffer)

'******************************************************************************
'ResetEvent
'Sets the event object in the overlapped structure to non-signaled.
'Requires a handle to the event object.
'Returns non-zero on success.
'******************************************************************************

Call ResetEvent(EventObject)

End Sub

Private Sub Shutdown()

'Actions that must execute when the program ends.

'Close the open handles to the device.

Result = CloseHandle _
    (HIDHandle)
Result = CloseHandle _
    (ReadHandle)

End Sub


Public Sub WriteReport()

'Send data to the device.

Dim Count As Integer
'Dim NumberOfBytesRead As Long
'Dim NumberOfBytesToSend As Long
Dim NumberOfBytesWritten As Long
'Dim ReadBuffer() As Byte
Dim SendBuffer() As Byte

'The SendBuffer array begins at 0, so subtract 1 from the number of bytes.

   'On Error GoTo Err_Proc

ReDim SendBuffer(Capabilities.OutputReportByteLength - 1)

'******************************************************************************
'WriteFile
'Sends a report to the device.
'Returns: success or failure.
'Requires: the handle returned by CreateFile and
'The output report byte length returned by HidP_GetCaps
'******************************************************************************

'The first byte is the Report ID
'Call device_out
SendBuffer(0) = 0

'The next bytes are data

For Count = 1 To Capabilities.OutputReportByteLength - 1
    SendBuffer(Count) = OutputReportData(Count - 1)
Next Count

NumberOfBytesWritten = 0

Result = WriteFile _
    (HIDHandle, _
    SendBuffer(0), _
    CLng(Capabilities.OutputReportByteLength), _
    NumberOfBytesWritten, _
    0)

'For Count = 1 To UBound(SendBuffer)
   
'Next Count


Exit_Proc:
   Exit Sub

Err_Proc:
   Err_Handler True, Err.Number, Err.Description, "HIDusb", "WriteReport"
   Err.Clear
   Resume Exit_Proc

End Sub



