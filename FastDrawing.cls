VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FastDrawing"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'***************************************************************************
'Real-time drawing class for Visual Basic 5.0/6.0
'©2002-2007 Tanner 'DemonSpectre' Helland
'Created: 10/03/01
'Last updated: 19/July/06
'Last update: Added 2D array conversions, added OriginalWidth/Height requirements
'             for setting image data
'Notes: The OriginalWidth/Height requirements have been added to circumvent errors
'       generated when programmers attempt to capture data from one picture box
'       and assign it to one with a different size.  That should be okay now - the
'       image will just be stretched to the size of the new box, rather than
'       crashing like before.
'
'This class is every graphics programmers dream - it does all the dirty API
'work required for lightning fast graphics, and all you have to do is call a
'couple of routines.  Am I a nice guy or what?  The format is simple; the only
'variables either of the subs require is the picture box you're going to do
'the image processing work on and an array to store the information in.  This is
'about as simple as it gets.  If you can't figure out how to use this module, maybe
'programming isn't the thing for you. ;)
'
'---------------
'TERMS OF USE:
'Feel free to compile this class into any executable of your own, but be
'sure to give me credit somewhere (and please notify me; I love hearing how
'this code helps fellow programmers out, and I find myself more motivated to
'contribute to the online programming community if I get feedback).
'
'DO NOT, however, upload or reproduce this particular class module to any website or
'other form of publicly available media.  If you would like to include this code as
'part of a personal project that will be distributed in code form, you MUST OBTAIN
'PERMISSION FROM ME.  This is for two reasons - one, to ensure that some idiot
'doesn't try to take credit for work he/she didn't do, and - two - to prevent faulty
'code from being distributed with my name attatched to it.  Several years ago I had
'a programmer from a foreign country try to redistribute this very code (with a
'similar brightness routine written by me) except that he had modified it so it
'didn't work correctly.  Stuff like that pisses me off beyond all compare.
'Eventually we were able to sort the problem out, but not before several hundred
'people downloaded faulty code with my name attatched to it.  So CONTACT ME BEFORE
'DISTRIBUTING THIS CODE IN A PUBLIC SETTING.
'
'Also, if you use these routines as part of a dll or ocx project, THEY CANNOT BE
'PART OF A PUBLIC ROUTINE.  They must be declared privately (i.e. not available
'to the end user).
'
'Lastly, I accept NO LEGAL RESPONSIBILITY for any problems that arise from use of
'this code.  Use it at your own risk.  A similar warning appears on my website, so
'you're on your own if this somehow breaks stuff.  It shouldn't, but the API is a
'vast wonderland of errors - so consider yourself warned.
'---------------
'
'Hopefully this code will be of some help to you.  If for any reason you do happen
'to encounter problems with this class, please let me know about them and I'll see
'if I can't help you out.
'
'Best regards,
'
'-Tanner 'DemonSpectre' Helland
'tannerhelland@hotmail.com
'http://www.tannerhelland.com
'http://www.studentsofgamedesign.com
'***************************************************************************

'Stripped down bitmap information
Private Type Bitmap
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type
'Call to transfer an object's properties into a custom variable
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, ByRef lpObject As Any) As Long
'Standard pixel data
Private Type RGBQUAD
        rgbBlue As Byte
        rgbGreen As Byte
        rgbRed As Byte
        rgbAlpha As Byte
End Type
'Full-size bitmap header
Private Type BITMAPINFOHEADER
        bmSize As Long
        bmWidth As Long
        bmHeight As Long
        bmPlanes As Integer
        bmBitCount As Integer
        bmCompression As Long
        bmSizeImage As Long
        bmXPelsPerMeter As Long
        bmYPelsPerMeter As Long
        bmClrUsed As Long
        bmClrImportant As Long
End Type
'Extended header for 8-bit images
Private Type BITMAPINFO
        bmHeader As BITMAPINFOHEADER
        bmColors(0 To 255) As RGBQUAD
End Type

'DIB section interfaces
Private Declare Function GetDIBits Lib "gdi32" (ByVal hdc As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function StretchDIBits Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dX As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal SrcWidth As Long, ByVal SrcHeight As Long, lpBits As Any, lpBitsInfo As BITMAPINFO, ByVal wUsage As Long, ByVal dwRop As Long) As Long

'Get the image width (via API - always accurate, unlike PictureBox.ScaleWidth)
Public Function GetImageWidth(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageWidth = bm.bmWidth
End Function

'Get the image height (via API - always accurate)
Public Function GetImageHeight(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageHeight = bm.bmHeight
End Function

'Get the stream length of an image (via API - always accurate)
Public Function GetImageStreamLength(SrcPictureBox As PictureBox) As Long
    Dim bm As Bitmap
    GetObject SrcPictureBox.Image, Len(bm), bm
    GetImageStreamLength = (bm.bmWidth * (bm.bmHeight + 1)) * 3
End Function

'Added 19/July/2006
'Get an image's pixel information into an array dimensioned (x * 3 + bgr, y)
Public Sub GetImageData2D(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array
    Erase ImageData()
    'Generate a correctly-dimensioned array (for 2-dimensional access)
    Dim ArrayWidth As Long
    ArrayWidth = (bm.bmWidth * 3) - 1
    ArrayWidth = ArrayWidth + (bm.bmWidth Mod 4)  '4-bit alignment
    ReDim ImageData(0 To ArrayWidth, 0 To bm.bmHeight) As Byte
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Get the image data into our array
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0), bmi, 0

End Sub

'Added 19/July/2006
'Set an image's pixel information from an array dimensioned (x * 3 + bgr, y)
Public Sub SetImageData2D(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

'Get an image's pixel information into an array dimensioned (r/g/b, x, y)
Public Sub GetImageData(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array
    Erase ImageData()
    ReDim ImageData(0 To 2, 0 To bm.bmWidth - 1, 0 To bm.bmHeight - 1)
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Get the image data into our array
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0, 0, 0), bmi, 0
End Sub

'Set an image's pixel information from an array dimensioned (r/g/b, x, y)
Public Sub SetImageData(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0, 0, 0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

'Get an image's pixel data into a one-dimesional array (stream)
Public Sub GetImageDataStream(SrcPictureBox As PictureBox, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject SrcPictureBox.Image, Len(bm), bm
    'Build a correctly sized array - in this case, designed as a stream
    Erase ImageData()
    ReDim ImageData(0 To GetImageStreamLength(SrcPictureBox))
    'Create a temporary header to pass to the GetDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = bm.bmWidth
    bmi.bmHeader.bmHeight = bm.bmHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Get the image data into our array
    GetDIBits SrcPictureBox.hdc, SrcPictureBox.Image, 0, bm.bmHeight, ImageData(0), bmi, 0
End Sub

'Set an image's data from a one-dimensional array (stream)
Public Sub SetImageDataStream(DstPictureBox As PictureBox, OriginalWidth As Long, OriginalHeight As Long, ImageData() As Byte)
    Dim bm As Bitmap
    'Get the picture box information
    GetObject DstPictureBox.Image, Len(bm), bm
    'Create a temporary header to pass to the StretchDIBits call
    Dim bmi As BITMAPINFO
    bmi.bmHeader.bmWidth = OriginalWidth
    bmi.bmHeader.bmHeight = OriginalHeight
    bmi.bmHeader.bmSize = 40                'Size, in bytes, of the header
    bmi.bmHeader.bmPlanes = 1               'Number of planes (always one for this instance)
    bmi.bmHeader.bmBitCount = 24            'Bits per pixel (always 24 for this instance)
    bmi.bmHeader.bmCompression = 0          'Compression :standard/none or RLE
    'Send the array to the picture box and draw it accordingly
    StretchDIBits DstPictureBox.hdc, 0, 0, bm.bmWidth, bm.bmHeight, 0, 0, OriginalWidth, OriginalHeight, ImageData(0), bmi, 0, vbSrcCopy
    'Since this doesn't automatically initialize AutoRedraw, we have to do it manually
    If DstPictureBox.AutoRedraw = True Then
        DstPictureBox.Picture = DstPictureBox.Image
        DstPictureBox.Refresh
    End If
    'Always good to manually halt for external processes after heavy API usage
    DoEvents
End Sub

