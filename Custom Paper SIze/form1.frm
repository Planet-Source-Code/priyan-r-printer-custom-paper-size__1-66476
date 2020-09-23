VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Priyan's Custom Paper size"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5400
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   5400
   StartUpPosition =   3  'Windows Default
   Begin VB.Label Label2 
      Caption         =   "Run the program it will automatically create a form priyan1 then set it to the printer object"
      Height          =   1935
      Left            =   600
      TabIndex        =   1
      Top             =   360
      Width           =   3855
   End
   Begin VB.Label Label1 
      Caption         =   "Inside folder Microsoft Article &. To manually add a form goto printers select printer->File->Server properties->Add New Form"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   3480
      Width           =   4455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
Dim psize As modprintercutompapersize.SIZEL, ret&
psize.cx = CLng(10) * 10000 '10cm
psize.cy = CLng(20) * 10000 '20 cm
MsgBox GetFormNames(Printer.DeviceName, Me.hwnd).Count & " Forms Found"
ret = GetFormproperties(Printer.DeviceName, Me.hwnd, psize, "Priyan1")
If ret <> 0 Then
Else
    MsgBox "New Form Added 'Priyan' :" & AddNewForm(Printer.DeviceName, psize, "Priyan1")
End If
ret = GetFormproperties(Printer.DeviceName, Me.hwnd, psize, "Priyan1")
MsgBox "Form 'priyan1' Printer Api method " & vbCrLf & "width:" & psize.cx / 10000 & " CM" & vbCrLf & "Height:" & psize.cy / 10000 & " CM"
SetDefaultForm Printer.DeviceName, "Priyan1", Me.hwnd
MsgBox "Form 'priyan1' From Printer Object " & vbCrLf & "width:" & Printer.ScaleX(Printer.ScaleWidth, Printer.ScaleMode, vbCentimeters) & " CM" & vbCrLf & "Height:" & Printer.ScaleY(Printer.ScaleHeight, Printer.ScaleMode, vbCentimeters) & " CM"
Printer.Print "Test"
Printer.EndDoc

End Sub
