Attribute VB_Name = "Module1"
Public Declare Function OpenUsbDevice Lib "USBIO.DLL" (ByVal MyUsbVendorID As Integer, ByVal MyUsbProductID As Integer) As Boolean
Public Declare Sub OutDataCtrl Lib "USBIO.DLL" (ByVal OutData As Byte, ByVal OutControl As Byte)
Public Declare Sub CloseUsbDevice Lib "USBIO.DLL" ()
Public Const VendorID = &H1234
Public Const ProductID = &H6789
