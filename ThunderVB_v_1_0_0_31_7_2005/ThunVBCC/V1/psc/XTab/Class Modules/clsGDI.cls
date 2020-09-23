VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsGDI"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' On 22nd Oct 2004
' By Neeraj Agrawal nja91@yahoo.com , neeraj_agrawal_ind@rediffmail.com

'clsGDI: GDI class used to store various Device context and Bitmap objects
'        Used for transaparent drawing of images (using mask color)


'The main image which is to be drawn transparantly
Public Property Get Image() As Long
  Image = m_lhBmpImage
End Property

Public Sub DrawImag Let Image(ByVal lhWndNewImage As Long)
  
  

End Property

Public Property Get MaskDC() As Long
  MaskDC = m_lhDCMask
End Property

Public Property Get ImageDC() As Long
  ImageDC = m_lhDCImage
End Property

Public Sub InitMe(hdc As Long, lMaskColor As Long)
  m_lhDCImage = CreateCompatibleDC(hdc)
  m_lhDCMask = CreateCompatibleDC(0)
  m_lMaskColor = lMaskColor
End Sub

Private Sub ReleaseBasicDCs()
  Dim lTmp As Long
  If m_lhBmpImageOld <> 0 Then
    'Select the default bitmap first
    lTmp = SelectObject(m_lhDCImage, m_lhBmpImageOld)
    
    
    'If lTmp <> 0 Then
    '  DeleteObject (lTmp)   'delete existing bitmap
    'End If
  End If
  Call DeleteDC(m_lhDCImage)
  
  
  If m_lhBmpMaskOld <> 0 Then
    'Select the default bitmap first
    lTmp = SelectObject(m_lhDCMask, m_lhBmpMaskOld)
    
    'If lTmp <> 0 Then
    '  DeleteObject (lTmp)   'delete existing bitmap
    'End If
  End If
  Call DeleteDC(m_lhDCMask)
  
End Sub

Public Sub DrawImage(lDestHDC As Long, lhBmp As Long, lTransColor As Long, iLeft As Integer, iTop As Integer, iWidth As Integer, iHeight As Integer)
  Dim lhDCImage As Long
  Dim lhBmpImageOld As Long
  Dim lhDCMask As Long
  Dim lhBmpMask As Long
  Dim lhBmpMaskOld As Long
  Dim lhDCTemp As Long
  Dim lhBmpTemp As Long
  Dim lhBmpTempOld As Long
  
  Dim lTmp As Long
  Dim utBitmap As BITMAP
  Dim lOldColor As Long
  
  
  '----------- For Image DC ------------
  
  lhDCImage = CreateCompatibleDC(lDestHDC)
  
  'Select the bitmap to be drawn into the DC
  m_lhBmpImageOld = SelectObject(lhDCImage, lhBmp)
  
  'Get Bitmap Dimensions
  Call GetObjectA(lhBmp, Len(utBitmap), utBitmap)
  
  '------------ For Mask DC -----------

  lhDCMask = CreateCompatibleDC(0)
  lhBmpMask = CreateCompatibleBitmap(lhDCMask, utBitmap.bmWidth, utBitmap.bmHeight)
    
    'Select the new bitmap into the DC
  lhBmpMaskOld = SelectObject(lhDCMask, lhBmpMask)
  
  '------------ For Temp DC --------------
 
  lhDCTemp = CreateCompatibleDC(0)
  lhBmpTemp = CreateCompatibleBitmap(lhDCTemp, utBitmap.bmWidth, utBitmap.bmHeight)
  
  'Select the new bitmap into the DC
  lhBmpTempOld = SelectObject(lhDCTemp, lTransColor)
  
  '-------------------------------------------
  
  'Now Perform BitOperations
  
  lOldColor = SetBkColor(lhDCImage, lTransColor)
  
  Call BitBlt(lhDCTemp, 0, 0, utBitmap.bmWidth, utBitmap.bmHeight, lhDCImage, 0, 0, SRCCOPY)
  
  Call SetBkColor(lhDCImage, lOldColor)
  
  Call BitBlt(lhDCImage, 0, 0, utBitmap.bmWidth, utBitmap.bmHeight, lhDCTemp, 0, 0, SRCPAINT)
  
  Call BitBlt(lhDCMask, 0, 0, utBitmap.bmWidth, utBitmap.bmHeight, lhDCImage, 0, 0, WHITENESS)
  
  Call BitBlt(lhDCMask, 0, 0, utBitmap.bmWidth, utBitmap.bmHeight, lhDCTemp, 0, 0, SRCINVERT)
     
     
  lOldColor = SetBkColor(lDestHDC, &HFFFFFF)
  
  Call BitBlt(lDestHDC, iLeft, iTop, iWidth, iHeight, lhDCMask, 0, 0, SRCPAINT)
  
  Call BitBlt(lDestHDC, iLeft, iTop, iWidth, iHeight, lhDCImage, 0, 0, SRCAND)

  Call SetBkColor(lDestHDC, lOldColor)
  
  'free the resources
  lTmp = SelectObject(lhDCTemp, lhBmpTempOld)
  DeleteObject (lTmp)
  DeleteDC (lhDCTemp)
  
  'free the resources
  lTmp = SelectObject(lhDCMask, lhBmpMaskOld)
  DeleteObject (lTmp)
  DeleteDC (lhDCMask)
  
  'free the resources
  lTmp = SelectObject(lhDCImage, lhBmpImageOld)
  DeleteDC (lhDCTemp)
End Sub

