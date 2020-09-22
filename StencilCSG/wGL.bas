Attribute VB_Name = "wGL"
Option Explicit



'*****GDI32*****
Public Type PIXELFORMATDESCRIPTOR
 nSize As Integer
 nVersion As Integer
 dwFlags As Long
 iPixelType As Byte
 cColorBits As Byte
 cRedBits As Byte
 cRedShift As Byte
 cGreenBits As Byte
 cGreenShift As Byte
 cBlueBits As Byte
 cBlueShift As Byte
 cAlphaBits As Byte
 cAlphaShift As Byte
 cAccumBits As Byte
 cAccumRedBits As Byte
 cAccumGreenBits As Byte
 cAccumBlueBits As Byte
 cAccumAlphaBits As Byte
 cDepthBits As Byte
 cStencilBits As Byte
 cAuxBuffers As Byte
 iLayerType As Byte
 bReserved As Byte
 dwLayerMask As Long
 dwVisibleMask As Long
 dwDamageMask As Long
End Type
Public Const BITSPIXEL As Long = 12
Public Const LPD_DOUBLEBUFFER As Long = 1
Public Const LPD_SHARE_ACCUM As Long = 256
Public Const LPD_SHARE_DEPTH As Long = 64
Public Const LPD_SHARE_STENCIL As Long = 128
Public Const LPD_STEREO As Long = 2
Public Const LPD_SUPPORT_GDI As Long = 16
Public Const LPD_SUPPORT_OPENGL As Long = 32
Public Const LPD_SWAP_COPY As Long = 1024
Public Const LPD_SWAP_EXCHANGE As Long = 512
Public Const LPD_TRANSPARENT As Long = 4096
Public Const LPD_TYPE_COLORINDEX As Long = 1
Public Const LPD_TYPE_RGBA As Long = 0
Public Const PFD_DEPTH_DONTCARE As Long = 536870912
Public Const PFD_DIRECT3D_ACCELERATED As Long = 16384
Public Const PFD_DOUBLEBUFFER As Long = 1
Public Const PFD_DOUBLEBUFFER_DONTCARE As Long = 1073741824
Public Const PFD_DRAW_TO_BITMAP As Long = 8
Public Const PFD_DRAW_TO_WINDOW As Long = 4
Public Const PFD_GENERIC_ACCELERATED As Long = 4096
Public Const PFD_GENERIC_FORMAT As Long = 64
Public Const PFD_MAIN_PLANE As Long = 0
Public Const PFD_NEED_PALETTE As Long = 128
Public Const PFD_NEED_SYSTEM_PALETTE As Long = 256
Public Const PFD_OVERLAY_PLANE As Long = 1
Public Const PFD_STEREO As Long = 2
Public Const PFD_STEREO_DONTCARE As Long = -2147483648#
Public Const PFD_SUPPORT_COMPOSITION As Long = 32768
Public Const PFD_SUPPORT_DIRECTDRAW As Long = 8192
Public Const PFD_SUPPORT_GDI As Long = 16
Public Const PFD_SUPPORT_OPENGL As Long = 32
Public Const PFD_SWAP_COPY As Long = 1024
Public Const PFD_SWAP_EXCHANGE As Long = 512
Public Const PFD_SWAP_LAYER_BUFFERS As Long = 2048
Public Const PFD_TYPE_COLORINDEX As Long = 1
Public Const PFD_TYPE_RGBA As Long = 0
Public Const PFD_UNDERLAY_PLANE As Long = -1
Public Declare Function ChoosePixelFormat Lib "gdi32" (ByVal hdc As Long, pPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function DescribePixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, ByVal un As Long, lpPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function GetPixelFormat Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SetPixelFormat Lib "gdi32" (ByVal hdc As Long, ByVal n As Long, pcPixelFormatDescriptor As PIXELFORMATDESCRIPTOR) As Long
Public Declare Function SwapBuffers Lib "gdi32" (ByVal hdc As Long) As Long

'****** wGL ******
Public Type POINTFLOAT
 X As Single
 Y As Single
End Type
Public Type GLYPHMETRICSFLOAT
 gmfBlackBoxX As Single
 gmfBlackBoxY As Single
 gmfptGlyphOrigin As POINTFLOAT
 gmfCellIncX As Single
 gmfCellIncY As Single
End Type
Public Type WGLSWAP
 hdc As Long
 uiFlags As Long
End Type
Public Type LAYERPLANEDESCRIPTOR
 nSize As Integer
 nVersion As Long
 dwFlags As Long 'LPDdwFlags
 iPixelType As Byte
 cColorBits As Byte
 cRedBits As Byte
 cRedShift As Byte
 cGreenBits As Byte
 cGreenShift As Byte
 cBlueBits As Byte
 cBlueShift As Byte
 cAlphaBits As Byte
 cAlphaShift As Byte
 cAccumBits As Byte
 cAccumRedBits As Byte
 cAccumGreenBits As Byte
 cAccumBlueBits As Byte
 cAccumAlphaBits As Byte
 cDepthBits As Byte
 cStencilBits As Byte
 cAuxBuffers As Byte
 iLayerPlane As Byte
 bReserved As Byte
 crTransparent As Long
End Type
Public Const WGL_FONT_LINES As Long = 0
Public Const WGL_FONT_POLYGONS As Long = 1
Public Const WGL_SWAP_MAIN_PLANE As Long = 1
Public Const WGL_SWAP_OVERLAY1 As Long = 2
Public Const WGL_SWAP_OVERLAY2 As Long = 4
Public Const WGL_SWAP_OVERLAY3 As Long = 8
Public Const WGL_SWAP_OVERLAY4 As Long = 16
Public Const WGL_SWAP_OVERLAY5 As Long = 32
Public Const WGL_SWAP_OVERLAY6 As Long = 64
Public Const WGL_SWAP_OVERLAY7 As Long = 128
Public Const WGL_SWAP_OVERLAY8 As Long = 256
Public Const WGL_SWAP_OVERLAY9 As Long = 512
Public Const WGL_SWAP_OVERLAY10 As Long = 1024
Public Const WGL_SWAP_OVERLAY11 As Long = 2048
Public Const WGL_SWAP_OVERLAY12 As Long = 4096
Public Const WGL_SWAP_OVERLAY13 As Long = 8192
Public Const WGL_SWAP_OVERLAY14 As Long = 16384
Public Const WGL_SWAP_OVERLAY15 As Long = 32768
Public Const WGL_SWAP_UNDERLAY1 As Long = 65536
Public Const WGL_SWAP_UNDERLAY2 As Long = 131072
Public Const WGL_SWAP_UNDERLAY3 As Long = 262144
Public Const WGL_SWAP_UNDERLAY4 As Long = 524288
Public Const WGL_SWAP_UNDERLAY5 As Long = 1048576
Public Const WGL_SWAP_UNDERLAY6 As Long = 2097152
Public Const WGL_SWAP_UNDERLAY7 As Long = 4194304
Public Const WGL_SWAP_UNDERLAY8 As Long = 8388608
Public Const WGL_SWAP_UNDERLAY9 As Long = 16777216
Public Const WGL_SWAP_UNDERLAY10 As Long = 33554432
Public Const WGL_SWAP_UNDERLAY11 As Long = 67108864
Public Const WGL_SWAP_UNDERLAY12 As Long = 134217728
Public Const WGL_SWAP_UNDERLAY13 As Long = 268435456
Public Const WGL_SWAP_UNDERLAY14 As Long = 536870912
Public Const WGL_SWAP_UNDERLAY15 As Long = 1073741824
Public Declare Function wglCopyContext Lib "opengl32" (ByVal hglrcSrc As Long, ByVal hlglrcDst As Long, ByVal mask As Long) As Long
Public Declare Function wglCreateContext Lib "opengl32" (ByVal hdc As Long) As Long
Public Declare Function wglCreateLayerContext Lib "opengl32" (ByVal hdc As Long, ByVal iLayerPlane As Long) As Long
Public Declare Function wglDeleteContext Lib "opengl32" (ByVal hglrc As Long) As Long
Public Declare Function wglDescribeLayerPlane Lib "opengl32" (ByVal hdc As Long, ByVal iPixelFormat As Long, ByVal iLayerPlane As Long, ByVal nBytes As Long, ByRef plpd As LAYERPLANEDESCRIPTOR) As Long
Public Declare Function wglGetCurrentContext Lib "opengl32" () As Long
Public Declare Function wglGetCurrentDC Lib "opengl32" () As Long
Public Declare Function wglGetLayerPaletteEntries Lib "opengl32" (ByVal hdc As Long, ByVal iLayerPlane As Long, ByVal iStart As Long, ByVal cEntries As Long, ByRef pcr As Long) As Long
Public Declare Function wglGetProcAddress Lib "opengl32" (ByVal lpszProc As String) As Long
Public Declare Function wglGetProcAddressANY Lib "opengl32" Alias "wglGetProcAddress" (ByRef lpStr As Any) As Long
Public Declare Function wglMakeCurrent Lib "opengl32" (ByVal hdc As Long, ByVal hglrc As Long) As Long
Public Declare Function wglRealizeLayerPalette Lib "opengl32" (ByVal hdc As Long, ByVal iLayerPlane As Long, ByVal bRealize As Long) As Long
Public Declare Function wglSetLayerPaletteEntries Lib "opengl32" (ByVal hdc As Long, ByVal iLayerPlane As Long, ByVal iStart As Long, ByVal cEntries As Long, ByRef pcr As Long) As Long
Public Declare Function wglShareLists Lib "opengl32" (ByVal hglrc1 As Long, ByVal hglrc2 As Long) As Long
Public Declare Function wglSwapLayerBuffers Lib "opengl32" (ByVal hdc As Long, ByVal fuPlanes As Long) As Long
Public Declare Function wglSwapMultipleBuffers Lib "opengl32" (ByVal MIDL_0151 As Long, ByRef MIDL_0152 As WGLSWAP) As Long
Public Declare Function wglUseFontBitmaps Lib "opengl32" (ByVal hdc As Long, ByVal first As Long, ByVal count As Long, ByVal listBase As Long) As Long
Public Declare Function wglUseFontOutlines Lib "opengl32" (ByVal hdc As Long, ByVal first As Long, ByVal count As Long, ByVal listBase As Long, ByVal deviation As Single, ByVal extrusion As Single, ByVal format As Long, ByRef lpgmf As GLYPHMETRICSFLOAT) As Long
