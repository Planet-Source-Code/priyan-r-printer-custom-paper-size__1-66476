Attribute VB_Name = "modprintercutompapersize"
Option Explicit

Public Declare Function EnumForms Lib "winspool.drv" Alias "EnumFormsA" _
    (ByVal hPrinter As Long, ByVal Level As Long, ByRef pForm As Any, _
    ByVal cbBuf As Long, ByRef pcbNeeded As Long, _
    ByRef pcReturned As Long) As Long

Public Declare Function AddForm Lib "winspool.drv" Alias "AddFormA" _
    (ByVal hPrinter As Long, ByVal Level As Long, pForm As Byte) As Long

Public Declare Function DeleteForm Lib "winspool.drv" Alias "DeleteFormA" _
    (ByVal hPrinter As Long, ByVal pFormName As String) As Long
    
Public Declare Function OpenPrinter Lib "winspool.drv" _
    Alias "OpenPrinterA" (ByVal pPrinterName As String, _
    phPrinter As Long, ByVal pDefault As Long) As Long

Public Declare Function ClosePrinter Lib "winspool.drv" _
    (ByVal hPrinter As Long) As Long

Public Declare Function DocumentProperties Lib "winspool.drv" _
    Alias "DocumentPropertiesA" (ByVal hwnd As Long, _
    ByVal hPrinter As Long, ByVal pDeviceName As String, _
    pDevModeOutput As Any, pDevModeInput As Any, ByVal fMode As Long) _
    As Long

Public Declare Function ResetDC Lib "gdi32" Alias "ResetDCA" _
    (ByVal hdc As Long, lpInitData As Any) As Long

Public Declare Sub CopyMemory Lib "KERNEL32" Alias "RtlMoveMemory" _
    (hpvDest As Any, hpvSource As Any, ByVal cbCopy As Long)

Public Declare Function lstrcpy Lib "KERNEL32" Alias "lstrcpyA" _
    (ByVal lpString1 As String, ByRef lpString2 As Long) As Long
    
' Optional functions not used in this sample, but may be useful.
Public Declare Function GetForm Lib "winspool.drv" Alias "GetFormA" _
    (ByVal hPrinter As Long, ByVal pFormName As String, _
    ByVal Level As Long, pForm As Byte, ByVal cbBuf As Long, _
    pcbNeeded As Long) As Long
    
Public Declare Function SetForm Lib "winspool.drv" Alias "SetFormA" _
    (ByVal hPrinter As Long, ByVal pFormName As String, _
    ByVal Level As Long, pForm As Byte) As Long

' Constants for DEVMODE
Public Const CCHFORMNAME = 32
Public Const CCHDEVICENAME = 32
Public Const DM_FORMNAME As Long = &H10000
Public Const DM_ORIENTATION = &H1&

' Constants for PRINTER_DEFAULTS.DesiredAccess
Public Const PRINTER_ACCESS_ADMINISTER = &H4
Public Const PRINTER_ACCESS_USE = &H8
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const PRINTER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or _
  PRINTER_ACCESS_ADMINISTER Or PRINTER_ACCESS_USE)
  
' Constants for DocumentProperties() call
Public Const DM_MODIFY = 8
Public Const DM_IN_BUFFER = DM_MODIFY
Public Const DM_COPY = 2
Public Const DM_OUT_BUFFER = DM_COPY

' Custom constants for this sample's SelectForm function
Public Const FORM_NOT_SELECTED = 0
Public Const FORM_SELECTED = 1
Public Const FORM_ADDED = 2

Public Type RECTL
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Public Type SIZEL
        cx As Long
        cy As Long
End Type

Public Type SECURITY_DESCRIPTOR
        Revision As Byte
        Sbz1 As Byte
        Control As Long
        Owner As Long
        Group As Long
        Sacl As Long  ' ACL
        Dacl As Long  ' ACL
End Type

' The two definitions for FORM_INFO_1 make the coding easier.
Public Type FORM_INFO_1
        Flags As Long
        pName As Long   ' String
        size As SIZEL
        ImageableArea As RECTL
End Type

Public Type sFORM_INFO_1
        Flags As Long
        pName As String
        size As SIZEL
        ImageableArea As RECTL
End Type

Public Type DEVMODE
        dmDeviceName As String * CCHDEVICENAME
        dmSpecVersion As Integer
        dmDriverVersion As Integer
        dmSize As Integer
        dmDriverExtra As Integer
        dmFields As Long
        dmOrientation As Integer
        dmPaperSize As Integer
        dmPaperLength As Integer
        dmPaperWidth As Integer
        dmScale As Integer
        dmCopies As Integer
        dmDefaultSource As Integer
        dmPrintQuality As Integer
        dmColor As Integer
        dmDuplex As Integer
        dmYResolution As Integer
        dmTTOption As Integer
        dmCollate As Integer
        dmFormName As String * CCHFORMNAME
        dmUnusedPadding As Integer
        dmBitsPerPel As Long
        dmPelsWidth As Long
        dmPelsHeight As Long
        dmDisplayFlags As Long
        dmDisplayFrequency As Long
End Type

Public Type PRINTER_DEFAULTS
        pDatatype As String
        pDevMode As Long    ' DEVMODE
        DesiredAccess As Long
End Type

Public Type PRINTER_INFO_2
        pServerName As String
        pPrinterName As String
        pShareName As String
        pPortName As String
        pDriverName As String
        pComment As String
        pLocation As String
        pDevMode As DEVMODE
        pSepFile As String
        pPrintProcessor As String
        pDatatype As String
        pParameters As String
        pSecurityDescriptor As SECURITY_DESCRIPTOR
        Attributes As Long
        Priority As Long
        DefaultPriority As Long
        StartTime As Long
        UntilTime As Long
        Status As Long
        cJobs As Long
        AveragePPM As Long
End Type
Public Function GetFormNames(ByVal printername$, MyHwnd As Long) As Collection
Set GetFormNames = New Collection
Dim NumForms As Long, I As Long
Dim FI1 As FORM_INFO_1
Dim aFI1() As FORM_INFO_1           ' Working FI1 array
Dim Temp() As Byte                  ' Temp FI1 array
Dim FormIndex As Integer
Dim BytesNeeded As Long
Dim PrinterHandle&
Dim RetVal As Long
If OpenPrinter(printername, PrinterHandle, 0&) Then
 
            FormIndex = 0
            ReDim aFI1(1)
            ' First call retrieves the BytesNeeded.
            RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
            ReDim Temp(BytesNeeded)
            ReDim aFI1(BytesNeeded / Len(FI1))
            ' Second call actually enumerates the supported forms.
            RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, _
                     NumForms)
            Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
            For I = 0 To NumForms - 1
                With aFI1(I)
                  GetFormNames.Add PtrCtoVbString(aFI1(I).pName)
                End With
            Next I
End If
End Function
'Returns Constant Value for The form & form size
Public Function GetFormproperties(ByVal printername$, MyHwnd As Long, _
                          FormSize As SIZEL, FormName As String) As Integer
Dim NumForms As Long, I As Long
Dim FI1 As FORM_INFO_1
Dim aFI1() As FORM_INFO_1           ' Working FI1 array
Dim Temp() As Byte                  ' Temp FI1 array
Dim FormIndex As Integer
Dim BytesNeeded As Long
Dim PrinterHandle&
Dim RetVal As Long
If OpenPrinter(printername, PrinterHandle, 0&) Then
 
            FormIndex = 0
            ReDim aFI1(1)
            ' First call retrieves the BytesNeeded.
            RetVal = EnumForms(PrinterHandle, 1, aFI1(0), 0&, BytesNeeded, NumForms)
            ReDim Temp(BytesNeeded)
            ReDim aFI1(BytesNeeded / Len(FI1))
            ' Second call actually enumerates the supported forms.
            RetVal = EnumForms(PrinterHandle, 1, Temp(0), BytesNeeded, BytesNeeded, _
                     NumForms)
            Call CopyMemory(aFI1(0), Temp(0), BytesNeeded)
            For I = 0 To NumForms - 1
                With aFI1(I)
                    If LCase(PtrCtoVbString(aFI1(I).pName)) = LCase(FormName) Then
                       ' Found the desired form
                        FormSize.cx = aFI1(I).size.cx
                        FormSize.cy = aFI1(I).size.cy
                        GetFormproperties = RetVal - 1
                        Exit For
                    Else
                        RetVal = RetVal + 1
                    End If
                End With
            Next I
End If
End Function

Public Function AddNewForm(ByVal printername$, FormSize As SIZEL, _
                           FormName As String) As Boolean
Dim FI1 As sFORM_INFO_1
Dim aFI1() As Byte
Dim RetVal As Long
Dim PrinterHandle&
With FI1
    .Flags = 0
    .pName = FormName
    With .size
        .cx = FormSize.cx
        .cy = FormSize.cy
    End With
    With .ImageableArea
        .Left = 0
        .Top = 0
        .Right = FI1.size.cx
        .Bottom = FI1.size.cy
    End With
End With
If OpenPrinter(printername, PrinterHandle, 0&) Then
        ReDim aFI1(Len(FI1))
        Call CopyMemory(aFI1(0), FI1, Len(FI1))
        RetVal = AddForm(PrinterHandle, 1, aFI1(0))
        If RetVal = 0 Then
            If Err.LastDllError = 5 Then
                'MsgBox "You do not have permissions to add a form to " & _
                   Printer.DeviceName, vbExclamation, "Access Denied!"
            Else
                'MsgBox "Error: " & Err.LastDllError, "Error Adding Form"
            End If
            
        Else
            AddNewForm = True
        End If
        ClosePrinter PrinterHandle
End If
End Function


Public Function PtrCtoVbString(ByVal Add As Long) As String
Dim sTemp As String * 512, x As Long

x = lstrcpy(sTemp, ByVal Add)
If (InStr(1, sTemp, Chr(0)) = 0) Then
     PtrCtoVbString = ""
Else
     PtrCtoVbString = Left(sTemp, InStr(1, sTemp, Chr(0)) - 1)
End If
End Function

Public Function SetDefaultForm(ByVal printername$, FormName As String, ByVal MyHwnd As Long) _
    As Boolean
Dim nSize As Long           ' Size of DEVMODE
Dim pDevMode As DEVMODE
Dim PrinterHandle As Long   ' Handle to printer
Dim hPrtDC As Long          ' Handle to Printer DC
Dim aDevMode() As Byte      ' Working DEVMODE
Dim FormSize As SIZEL

'printername = Printer.DeviceName  ' Current printer
hPrtDC = Printer.hdc              ' hDC for current Printer
SetDefaultForm = FORM_NOT_SELECTED    ' Set for failure unless reset in code.

' Get a handle to the printer.
If OpenPrinter(printername, PrinterHandle, 0&) Then
    ' Retrieve the size of the DEVMODE.
    nSize = DocumentProperties(MyHwnd, PrinterHandle, printername, 0&, _
            0&, 0&)
    ' Reserve memory for the actual size of the DEVMODE.
    ReDim aDevMode(1 To nSize)

    ' Fill the DEVMODE from the printer.
    nSize = DocumentProperties(MyHwnd, PrinterHandle, printername, _
            aDevMode(1), 0&, DM_OUT_BUFFER)
    ' Copy the Public (predefined) portion of the DEVMODE.
    Call CopyMemory(pDevMode, aDevMode(1), Len(pDevMode))
    
    ' If FormName is "MyCustomForm", we must make sure it exists
    ' before using it. Otherwise, it came from our EnumForms list,
    ' and we do not need to check first. Note that we could have
    ' passed in a Flag instead of checking for a literal name.
    
    
    ' Change the appropriate member in the DevMode.
    ' In this case, you want to change the form name.
    pDevMode.dmFormName = FormName & Chr(0)  ' Must be NULL terminated!
    ' Set the dmFields bit flag to indicate what you are changing.
    pDevMode.dmFields = DM_FORMNAME

    ' Copy your changes back, then update DEVMODE.
    Call CopyMemory(aDevMode(1), pDevMode, Len(pDevMode))
    nSize = DocumentProperties(MyHwnd, PrinterHandle, printername, _
            aDevMode(1), aDevMode(1), DM_IN_BUFFER Or DM_OUT_BUFFER)

    nSize = ResetDC(hPrtDC, aDevMode(1))   ' Reset the DEVMODE for the DC.

    ' Close the handle when you are finished with it.
    ClosePrinter (PrinterHandle)
    ' Selection Succeeded! But was Form Added?
   SetDefaultForm = True
Else
    
End If
End Function


                    



