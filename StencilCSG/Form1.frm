VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H00000000&
   Caption         =   "Stencil buffer CSG"
   ClientHeight    =   8040
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   13080
   LinkTopic       =   "Form1"
   ScaleHeight     =   536
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   872
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame TPanel 
      BorderStyle     =   0  'None
      Height          =   6060
      Left            =   5205
      TabIndex        =   0
      Top             =   150
      Width           =   2610
      Begin VB.Timer Timer 
         Enabled         =   0   'False
         Interval        =   2000
         Left            =   210
         Top             =   5535
      End
      Begin VB.ListBox lstOp 
         Appearance      =   0  'Flat
         Height          =   1590
         ItemData        =   "Form1.frx":0000
         Left            =   150
         List            =   "Form1.frx":0016
         TabIndex        =   6
         Top             =   3720
         Width           =   2220
      End
      Begin VB.ListBox lstOpB 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "Form1.frx":0049
         Left            =   180
         List            =   "Form1.frx":0059
         TabIndex        =   5
         Top             =   1950
         Width           =   2220
      End
      Begin VB.ListBox lstOpA 
         Appearance      =   0  'Flat
         Height          =   1005
         ItemData        =   "Form1.frx":007B
         Left            =   180
         List            =   "Form1.frx":008B
         TabIndex        =   2
         Top             =   540
         Width           =   2220
      End
      Begin VB.Label Label3 
         Caption         =   "Operación CSG"
         Height          =   240
         Left            =   210
         TabIndex        =   4
         Top             =   3390
         Width           =   1125
      End
      Begin VB.Label Label2 
         Caption         =   "Operando B"
         Height          =   240
         Left            =   150
         TabIndex        =   3
         Top             =   1620
         Width           =   915
      End
      Begin VB.Label Label1 
         Caption         =   "Operando A"
         Height          =   240
         Left            =   105
         TabIndex        =   1
         Top             =   195
         Width           =   915
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long

Private Enum CSGMODE '// Various CSG operations available:
 CSG_ONLY_A 'Draws only operand A
 CSG_ONLY_B ' Draws only operand B
 CSG_UNION ' "A or B"
 CSG_INTERSECTION ' "A and B" 'Default setting.
 CSG_A_MINUS_B '"A - B"
 CSG_B_MINUS_A ' "B - A"
End Enum
Private Const PI As Double = 3.14159265358979
Private Const RADIANS As Single = PI / 180 ' 0.01745329251994 'Multiplicar para obtener Radians
Private Enum progress
 pCONTINUE
 pSTOP
End Enum
Private CSG_MODE As CSGMODE
Private texA As Long, texB As Long   'as TCGTextureObject 'Textures.
Private a As Integer  ' Rotation angle.
Private opA As Long, opB As Long  'as TCSGDrawObjectProc - punteros a que objetos van a dibujar
Dim pnlControls As Long, LP(3) As Single, matAmb(3) As Single, matDiff(3) As Single, matSpec(3) As Single
Private gHDC As Long, gHRC As Long, bDrawing As Boolean, bInFrame As Boolean
Private Type VECTOR3F
 X As Single
 Y As Single
 z As Single
End Type
Private Type Vertex
 PX As Single   'Posicion en X
 PY As Single   'Posicion en Y
 PZ As Single   'Posicion en Z
 nx As Single   'Normal en X
 ny As Single   'Normal en Y
 nz As Single   'Normal en Z
 tU As Single   'Texturado en X
 tV As Single   'texturado en Y
 r As Single    'componente rojo del color  - glColor3f
 G As Single    'componente verde del color - glColor3f
 b As Single    'componente azul del color  - glColor3f
 a As Single    'componente de transparencia - glcolor4f
 'mat As Long    'Número de material del vértice
End Type
Private Type Solid
 nVertex As Long
 modo As Long 'glBeginModeConstants
 Min As VECTOR3F
 max As VECTOR3F
 Cen As VECTOR3F
 v() As Vertex
End Type
Private bResize As Boolean, cR As Single, cG As Single, cB As Single
Private Cubo As Solid, Cilindro As Solid, Cono As Solid, Esfera As Solid ', TmpA As Solid, TmpB As Solid
  
Private CUBE As Long, CILYNDER As Long, CONE As Long, SPHERE As Long, base As Long

  
Private Sub Form_Load()
 a = 0
 LP(0) = -1 '0.1
 LP(1) = 1 '0.1
 LP(2) = 1 '0.1
 LP(3) = 0
 'Intensidad de la luz ambiente
 matAmb(0) = 0.5: matAmb(1) = 0.5: matAmb(2) = 0.5: matAmb(3) = 0.6
 'Color de la luz
 matDiff(0) = 1: matDiff(1) = 1: matDiff(2) = 1: matDiff(3) = 0.6
 'Brillo especular
 matSpec(0) = 1: matSpec(1) = 1: matSpec(2) = 1: matSpec(3) = 0.6

 
 ' Initialize listboxes and default settings:
 lstOpA.ListIndex = 0
 lstOpB.ListIndex = 3
 lstOp.ListIndex = 3
 CSG_MODE = lstOp.ListIndex ' CSG_INTERSECTION
 opA = 0 ' opA := csgDrawCube;
 opB = 3 ' opB := csgDrawSphere;
 pnlControls = 174 'tpanel.width
 Me.Show
 DoEvents
 CreateCube 1.2
 CreateCilynder 0.5, 1.1, 60
 CreateCone 0.9, 1, 60
 CreateSphere 0.4, 20
 
 bResize = True
 gHDC = Me.hdc
 InitGL
 
'  Call glNewList(CILYNDER, GL_COMPILE)
'  CILYNDER = gluNewQuadric()
'  'BASE = gluNewQuadric()
'  'call glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE, cone_mat)
'  'Call gluQuadricOrientation(BASE, GLU_INSIDE)
'  'Call gluDisk(BASE, 0, 15, 64, 1)
'  Call gluCylinder(CILYNDER, 0.5, 0.5, 1.1, 64, 64)
'  Call gluDeleteQuadric(CILYNDER)
'  'Call gluDeleteQuadric(BASE)
' Call glEndList
'
' Call glNewList(CONE, GL_COMPILE)
'  CONE = gluNewQuadric()
'  BASE = gluNewQuadric()
'  'call glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE, cone_mat)
'  Call gluQuadricOrientation(BASE, GLU_INSIDE)
'  Call gluDisk(BASE, 0, 0.9, 64, 1)
'  Call gluCylinder(CONE, 0.9, 0, 1, 64, 64)
'  Call gluDeleteQuadric(CONE)
'  Call gluDeleteQuadric(BASE)
' Call glEndList
' Call glNewList(SPHERE, GL_COMPILE)
'  SPHERE = gluNewQuadric()
'  Call glColor3f(0.1, 0.5, cB = 0.8)
'  ' glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE, sphere_mat)
'  Call gluSphere(SPHERE, 0.4, 64, 64)
'  Call gluDeleteQuadric(SPHERE)
' Call glEndList
 
 bDrawing = True
' Timer.Enabled = True
 Do
  RenderFrame
  DoEvents
 Loop While bDrawing = True
 Erase Cubo.v, Cilindro.v, Cono.v, Esfera.v

 ' Load textures:
 'texA := TCGTextureObject.Create;
 'texA.Image.LoadFromFile('textureA.cgi');
 'texA.Upload;
 'texB := TCGTextureObject.Create;
 'texB.Image.LoadFromFile('textureB.cgi');
 'texB.Upload;
End Sub

Private Sub Form_Paint()
 'RenderFrame
End Sub

Private Sub Form_Resize()
 If WindowState = vbMinimized Then Exit Sub
 TPanel.Move ScaleWidth - pnlControls, 0, pnlControls, ScaleHeight
 'Reset the projection matrix.
 Call glViewport(0, 0, ScaleWidth - pnlControls, ScaleHeight)
 Call glMatrixMode(GL_PROJECTION)
 Call glLoadIdentity
 Call gluPerspective(60, (ScaleWidth - pnlControls) / ScaleHeight, 0.1, 3200)
 Call glMatrixMode(GL_MODELVIEW)
 Call glLoadIdentity
' Call glMatrixMode(GL_PROJECTION)
' Call glOrtho(-50#, 50#, -50#, 50#, -50#, 50#)
' Call glMatrixMode(GL_MODELVIEW)
End Sub

Private Sub Form_Unload(Cancel As Integer)
 bDrawing = False
 Timer.Enabled = False
 If gHRC <> 0 Then
  Call wglMakeCurrent(gHDC, 0)
  Call wglDeleteContext(gHRC)
 End If
 Call ReleaseDC(hwnd, gHDC)
 End
 'The textures are all we have to delete.
 'texA.Free;
 'texB.Free;
End Sub



                                              


Private Sub lstOp_Click()
 CSG_MODE = lstOp.ListIndex
End Sub
Private Sub lstOpA_Click()
 opA = lstOpA.ListIndex
End Sub
Private Sub lstOpB_Click()
 opB = lstOpB.ListIndex
End Sub


Private Sub Timer_Timer()
 a = (a + 1) Mod 360
 RenderFrame  'Form_Paint
End Sub
Private Sub InitGL()
 On Local Error GoTo NoIniciado
 Dim pfd As PIXELFORMATDESCRIPTOR, PixelFormat As Long, CT As Long, Lt As Long
 With pfd
  .nSize = Len(pfd)
  .nVersion = 1
  .dwFlags = PFD_SUPPORT_OPENGL Or PFD_DRAW_TO_WINDOW Or PFD_DOUBLEBUFFER Or PFD_TYPE_RGBA
  .iPixelType = PFD_TYPE_RGBA
  .cColorBits = CByte(GetDeviceCaps(GetDesktopWindow, BITSPIXEL))  'xBpp
  .cAlphaBits = 8 'mglAlphaBits
  .cDepthBits = 32 'mglDepthBits
  .cStencilBits = 8 'mglStencilBits
  .iLayerType = PFD_MAIN_PLANE
 End With
 PixelFormat = ChoosePixelFormat(gHDC, pfd)
 If PixelFormat = 0 Then GoTo NoIniciado
 Call SetPixelFormat(gHDC, PixelFormat, pfd)
 gHRC = wglCreateContext(gHDC)
 Call wglMakeCurrent(gHDC, gHRC)
 Call glClearColor(0#, 0#, 0#, 0#) '0.1, 0.1, 0.5, 1)    ' Black Background
 Call glShadeModel(GL_SMOOTH) ' Enables Smooth Color Shading
 Call glClearDepth(1#) ' Depth Buffer Setup
 Call glEnable(GL_DEPTH_TEST) ' Enable Depth Buffer
 Call glDepthFunc(GL_LESS) ' The Type Of Depth Test To Do
 Call glHint(GL_PERSPECTIVE_CORRECTION_HINT, GL_NICEST) 'Realy Nice perspective calculations
 Call glEnable(GL_BLEND)    'enable alpha blending
 Call glBlendFunc(GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA)    'set alpha blending options
 Call glEnable(GL_COLOR_MATERIAL)
 'Call glEnable(GL_NORMALIZE)
 Call glFrontFace(GL_CCW)
 Call glEnable(GL_CULL_FACE)
 Call glLightModeli(GL_LIGHT_MODEL_TWO_SIDE, 1)
 Call glLightModeli(GL_LIGHT_MODEL_LOCAL_VIEWER, 0)
 Call glEnable(GL_LIGHTING)
 Call glEnable(GL_LIGHT0)
 Call glLightfv(GL_LIGHT0, GL_POSITION, LP(0))
 Call glLightfv(GL_LIGHT0, GL_AMBIENT, matAmb(0))
 Call glLightfv(GL_LIGHT0, GL_DIFFUSE, matDiff(0))
 Call glLightfv(GL_LIGHT0, GL_SPECULAR, matSpec(0))
 Exit Sub
NoIniciado:
 MsgBox "Error iniciando"
End Sub
Private Sub PageFlip()
' Call glFinish 'Call glFlush
 Call SwapBuffers(gHDC)
' Call wglMakeCurrent(gHDC, 0)
End Sub
'Dibuja el onjeto especificado - Sustituye a los punteros a funciones
Private Sub Render(ByVal Numero As Long)
 Select Case Numero 'TCSGDrawObjectProc
  Case 0 'Cubo
   DrawSolid Cubo
  Case 1 'Cono
   DrawSolid Cono
   'Call glCallList(CONE)
  Case 2 'Esfera
   DrawSolid Esfera
   'Call glCallList(SPHERE)
  Case 3 'Cilindro
   DrawSolid Cilindro
   'Call glCallList(CILYNDER)
 End Select
End Sub


Private Sub RenderFrame()
 If bInFrame Then Exit Sub
 bInFrame = True
' Call wglMakeCurrent(gHDC, gHRC)
 Call glClear(GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT Or GL_STENCIL_BUFFER_BIT)
 a = (a + 1) ' Mod 360'  A = -90 * RADIANS
 If a > 360 Then a = 0
  
' Call glViewport(0, 0, ScaleWidth - pnlControls, ScaleHeight)
' Call glMatrixMode(GL_MODELVIEW)
' glLoadIdentity
 
  Call glViewport(0, 0, ScaleWidth - pnlControls, ScaleHeight)
 Call glMatrixMode(GL_PROJECTION)
 Call glLoadIdentity
 Call gluPerspective(60, (ScaleWidth - pnlControls) / ScaleHeight, 0.001, 3200)
 Call glMatrixMode(GL_MODELVIEW)
 Call glLoadIdentity
 Call glPushMatrix
 Call glTranslatef(0, 0, -3)
 Call glRotatef(a, 1, 1, 1) ' Sin(a * RADIANS), 1, Cos(a * RADIANS))


  'Just call one of the CSG routines, passing it the current operands.
  Select Case CSG_MODE
   Case CSG_ONLY_A: Call one(opA)
   Case CSG_ONLY_B: Call one(opB)
   Case CSG_UNION: Call oru(opA, opB)
   Case CSG_INTERSECTION: Call andI(opA, opB)
   Case CSG_A_MINUS_B: Call subs(opB, opA)
   Case CSG_B_MINUS_A: Call subs(opA, opB)
  End Select
  Call glPopMatrix
  PageFlip
  bInFrame = False
End Sub


'crea un cubo cuadrado - funcion probada
Public Sub CreateCube(Optional ByVal SizeFaces As Single = 1)
 ReDim Cubo.v(0)
 Cubo.nVertex = 0
 Cubo.modo = GL_QUADS
 Dim lado As Single
 lado = SizeFaces * 0.5
 cR = 0: cG = 1: cB = 0
 'cara 1 - Y+ (eje z-)
 Call AddVertexXYZ(lado, -lado, -lado, Cubo)
 Call AddVertexXYZ(-lado, -lado, -lado, Cubo)
 Call AddVertexXYZ(-lado, lado, -lado, Cubo)
 Call AddVertexXYZ(lado, lado, -lado, Cubo)
 'cara 2 - Y- (eje z+)
 Call AddVertexXYZ(lado, -lado, lado, Cubo)
 Call AddVertexXYZ(lado, lado, lado, Cubo)
 Call AddVertexXYZ(-lado, lado, lado, Cubo)
 Call AddVertexXYZ(-lado, -lado, lado, Cubo)
 'cara 3 - Z+ (eje y+)
 Call AddVertexXYZ(lado, lado, -lado, Cubo)
 Call AddVertexXYZ(-lado, lado, -lado, Cubo)
 Call AddVertexXYZ(-lado, lado, lado, Cubo)
 Call AddVertexXYZ(lado, lado, lado, Cubo)
 'cara 4 - Z- (eje y-)
 Call AddVertexXYZ(lado, -lado, -lado, Cubo)
 Call AddVertexXYZ(lado, -lado, lado, Cubo)
 Call AddVertexXYZ(-lado, -lado, lado, Cubo)
 Call AddVertexXYZ(-lado, -lado, -lado, Cubo)
 'cara 5 - X+ (eje x+)
 Call AddVertexXYZ(lado, -lado, lado, Cubo)
 Call AddVertexXYZ(lado, -lado, -lado, Cubo)
 Call AddVertexXYZ(lado, lado, -lado, Cubo)
 Call AddVertexXYZ(lado, lado, lado, Cubo)
 'cara 6 - X- (eje x-)
 Call AddVertexXYZ(-lado, -lado, lado, Cubo)
 Call AddVertexXYZ(-lado, lado, lado, Cubo)
 Call AddVertexXYZ(-lado, lado, -lado, Cubo)
 Call AddVertexXYZ(-lado, -lado, -lado, Cubo)
End Sub
Private Sub CreateCilynder(Optional ByVal Radio As Single = 1, Optional ByVal Longitud As Single = 3, Optional ByVal resolucion As Integer = 72)
 ReDim Cilindro.v(0)
 Cilindro.nVertex = 0
 Cilindro.modo = GL_QUADS
 cR = 1: cG = 1: cB = 0
 Dim p As Integer, a As Integer, Ar As Single
 Dim X As Single, Y As Single, z As Single
 Dim x2 As Single, y2 As Single, Z2 As Single
 Dim AR2 As Single
 If resolucion < 4 Then resolucion = 4
COMPROBAR:
 p = (360 Mod resolucion)
 If p <> 0 Then
  resolucion = resolucion + 1
  GoTo COMPROBAR
 End If
 p = 360 / resolucion 'Numero de puntos
 'Crea el cuerpo
 z = 0
 Z2 = Longitud
 For a = 0 To 360 - p Step p
  Ar = a * RADIANS
  AR2 = (a + p) * RADIANS
  X = Radio * Cos(Ar)
  Y = Radio * Sin(Ar)
  x2 = Radio * Cos(AR2)
  y2 = Radio * Sin(AR2)
  Call AddVertexXYZ(x2, z, y2, Cilindro)
  Call AddVertexXYZ(X, z, Y, Cilindro)
  Call AddVertexXYZ(X, Z2, Y, Cilindro)
  Call AddVertexXYZ(x2, Z2, y2, Cilindro)
'  'Cara trasera
'  Call AddVertexXYZ(X2, Z, Y2, Cilindro)
'  Call AddVertexXYZ(X2, Z2, Y2, Cilindro)
'  Call AddVertexXYZ(X, Z2, Y, Cilindro)
'  Call AddVertexXYZ(X, Z, Y, Cilindro)
 Next a
 'Crea la tapa superior
 z = Longitud
 For a = 0 To 360 - p Step p
  Ar = a * RADIANS
  AR2 = (a + p) * RADIANS
  X = Radio * Cos(Ar)
  Y = Radio * Sin(Ar)
  x2 = Radio * Cos(AR2)
  y2 = Radio * Sin(AR2)
  Call AddVertexXYZ(0, z, 0, Cilindro)
  Call AddVertexXYZ(0, z, 0, Cilindro)
  Call AddVertexXYZ(x2, z, y2, Cilindro)
  Call AddVertexXYZ(X, z, Y, Cilindro)
'  'Cara trasera
'  Call AddVertexXYZ(0, Z, 0, Cilindro)
'  Call AddVertexXYZ(X, Z, Y, Cilindro)
'  Call AddVertexXYZ(X2, Z, Y2, Cilindro)
'  Call AddVertexXYZ(0, Z, 0, Cilindro)
 Next a
 'Crea la tapa inferior
 z = 0
 For a = 0 To 360 - p Step p
  Ar = a * RADIANS
  AR2 = (a + p) * RADIANS
  X = Radio * Cos(Ar)
  Y = Radio * Sin(Ar)
  x2 = Radio * Cos(AR2)
  y2 = Radio * Sin(AR2)
  Call AddVertexXYZ(0, z, 0, Cilindro)
  Call AddVertexXYZ(0, z, 0, Cilindro)
  Call AddVertexXYZ(X, z, Y, Cilindro)
  Call AddVertexXYZ(x2, z, y2, Cilindro)
 Next a
End Sub
Private Sub CreateCone(Optional ByVal Radio As Single = 1, Optional ByVal Longitud As Single = 3, Optional ByVal resolucion As Integer = 72)
 ReDim Cono.v(0)
 Cono.nVertex = 0
 Cono.modo = GL_TRIANGLES
 cR = 0: cG = 1: cB = 1
 Dim p As Integer, a As Integer, Ar As Single
 Dim X As Single, Y As Single, z As Single
 Dim x2 As Single, y2 As Single, Z2 As Single
 Dim AR2 As Single
 If resolucion < 4 Then resolucion = 4
COMPROBAR:
 p = (360 Mod resolucion)
 If p <> 0 Then
  resolucion = resolucion + 1
  GoTo COMPROBAR
 End If
 p = 360 / resolucion 'Numero de puntos
 'Crea el cuerpo
 z = 0
 Z2 = Longitud
 For a = 0 To 360 - p Step p
  Ar = a * RADIANS
  AR2 = (a + p) * RADIANS
  X = Radio * Cos(Ar)
  Y = Radio * Sin(Ar)
  x2 = Radio * Cos(AR2)
  y2 = Radio * Sin(AR2)
  Call AddVertexXYZ(0, Z2, 0, Cono)
  Call AddVertexXYZ(x2, z, y2, Cono)
  Call AddVertexXYZ(X, z, Y, Cono)
 Next a
 'Crea la tapa inferior
 z = 0
 For a = 0 To 360 - p Step p
  Ar = a * RADIANS
  AR2 = (a + p) * RADIANS
  X = Radio * Cos(Ar)
  Y = Radio * Sin(Ar)
  x2 = Radio * Cos(AR2)
  y2 = Radio * Sin(AR2)
  Call AddVertexXYZ(0, z, 0, Cono)
  Call AddVertexXYZ(X, z, Y, Cono)
  Call AddVertexXYZ(x2, z, y2, Cono)
 Next a
End Sub
Private Sub CreateSphere(Optional ByVal radius As Single = 1, Optional nDivisions As Integer = 20)
 ReDim Esfera.v(0)
 Esfera.nVertex = 0
 Esfera.modo = GL_TRIANGLES
 cR = 0.1: cG = 0.5: cB = 0.8
 Dim Y As Long, K As Long, angle As Double, angleincr As Double, matShine As Single
 Dim mat() As Single, p(0 To 3) As VECTOR3F
 Dim n As Integer, m As Long
 angle = 0
 n = nDivisions '40
 m = n * 4
 angleincr = PI / n
 ReDim mat(0 To m + 1)
 For K = 0 To m / 2 '40
  mat(K * 2) = Cos(angle)
  mat(K * 2 + 1) = Sin(angle)
  angle = angle + angleincr
 Next K
 matShine = 40
 For Y = 1 To n
  For K = 1 To n * 2
   p(0).X = radius * (mat(K * 2) * mat(Y * 2 + 1))
   p(0).Y = radius * (mat(Y * 2))
   p(0).z = radius * (mat(K * 2 + 1) * mat(Y * 2 + 1))
   p(1).X = radius * (mat((K - 1) * 2) * mat(Y * 2 + 1))
   p(1).Y = radius * (mat(Y * 2))
   p(1).z = radius * (mat((K - 1) * 2 + 1) * mat(Y * 2 + 1))
   p(2).X = radius * (mat((K - 1) * 2) * mat((Y - 1) * 2 + 1))
   p(2).Y = radius * (mat((Y - 1) * 2))
   p(2).z = radius * (mat((K - 1) * 2 + 1) * mat((Y - 1) * 2 + 1))
   p(3).X = radius * (mat(K * 2) * mat((Y - 1) * 2 + 1))
   p(3).Y = radius * (mat((Y - 1) * 2))
   p(3).z = radius * (mat(K * 2 + 1) * mat((Y - 1) * 2 + 1))
   Call AddVertexXYZ(p(0).X, p(0).Y, p(0).z, Esfera)
   Call AddVertexXYZ(p(1).X, p(1).Y, p(1).z, Esfera)
   Call AddVertexXYZ(p(2).X, p(2).Y, p(2).z, Esfera)
   Call AddVertexXYZ(p(2).X, p(2).Y, p(2).z, Esfera)
   Call AddVertexXYZ(p(3).X, p(3).Y, p(3).z, Esfera)
   Call AddVertexXYZ(p(0).X, p(0).Y, p(0).z, Esfera)
  Next K
 Next Y
 Erase mat, p
End Sub


'Añadir vertice por componentes X,Y,Z
Private Sub AddVertexXYZ(ByVal X As Single, ByVal Y As Single, ByVal z As Single, O As Solid)
 Dim temp As Long
 temp = O.nVertex
 ReDim Preserve O.v(temp)
 With O
  .v(temp).PX = X
  .v(temp).PY = Y
  .v(temp).PZ = z
  'utilizamos color del objeto por defecto
  .v(temp).r = cR
  .v(temp).G = cG
  .v(temp).b = cB
  .v(temp).a = 1
  'Calcula la normal del punto
  Dim Magnitud As Single, fMult As Single
  fMult = 0
  Magnitud = CSng(Sqr((X * X) + (Y * Y) + (z * z)))
  If Magnitud <> 0 Then fMult = 1 / Magnitud
  .v(temp).nx = X * fMult
  .v(temp).ny = Y * fMult
  .v(temp).nz = z * fMult
  'Calcula puntos caja limite
  If X < .Min.X Then .Min.X = X - 0.1
  If Y < .Min.Y Then .Min.Y = Y - 0.1
  If z < .Min.z Then .Min.z = z - 0.1
  If X > .max.X Then .max.X = X + 0.1
  If Y > .max.Y Then .max.Y = Y + 0.1
  If z > .max.z Then .max.z = z + 0.1
  .Cen.X = (.max.X - .Min.X) * 0.5
  .Cen.Y = (.max.Y - .Min.Y) * 0.5
  .Cen.z = (.max.z - .Min.z) * 0.5
  'añade numero de puntos
  temp = temp + 1
  O.nVertex = temp
 End With
End Sub


Private Sub DrawSolid(O As Solid)
 If O.nVertex = 0 Then Exit Sub
 Call glPushMatrix
 'realiza Escala,Rotacion y Traslacion si es necesario
 Call glVertex3f(O.Cen.X, O.Cen.Y, O.Cen.z) 'Define el vertice de rotacion
 'If bScale Then Call glScalef(Sca.X, Sca.Y, Sca.Z)
 'If bRotate Then
 ' Call glRotatef(Rot.Y, 0, 1, 0)
 ' Call glRotatef(Rot.Z, 0, 0, 1)
 ' Call glRotatef(Rot.X, 1, 0, 0)
 'End If
 'If bMove Then Call glTranslatef(pos.X, pos.Y, pos.Z)
 'If bCompiled Then
 'Call glCallList(nCompiled)
 'Else
 
 'Ahora dibujamos
 'Call glPushName(oIndex) 'para modo pick
 
 'Seleccionamos cantidad de vertices segun modo
 Dim TVERTEX As Long, MinP As Integer, vFor As Long, md As Long 'glBeginModeConstants
 Dim n As Long, n1 As Long, n2 As Long, n3 As Long, n4 As Long
 TVERTEX = O.nVertex - 1
 md = O.modo
 Select Case md
  Case GL_POINTS, GL_QUAD_STRIP, GL_TRIANGLE_FAN, GL_TRIANGLE_STRIP, GL_LINE_LOOP, GL_LINE_STRIP
   MinP = 1
  Case GL_LINES
   MinP = 2
  Case GL_TRIANGLES
   MinP = 3
  Case GL_QUADS
   MinP = 4
  Case GL_POLYGON    'bmLineStrip,¿bmpolygon - 1 punto?
   'MinP = 3 'Se necesitan al menos 3 puntos para dibujar
  'Case gl_trianglestrip
   'MinP = 4 'Se necesitan al menos 4 puntos para dibujar
 End Select
 'Ahora dibuja los vértices del objeto
 Call glBegin(md)
  For vFor = 0 To TVERTEX Step MinP
   Select Case O.modo
    Case GL_POINTS, GL_QUAD_STRIP, GL_TRIANGLE_FAN, GL_TRIANGLE_STRIP, GL_LINE_LOOP, GL_LINE_STRIP
     Call glColor4f(O.v(vFor).r, O.v(vFor).G, O.v(vFor).b, O.v(vFor).a)
     'If bTextured Then Call glTexCoord2f(O.V(vfor).tU, O.V(vfor).tV)
     Call glNormal3f(O.v(vFor).nx, O.v(vFor).ny, O.v(vFor).nz)
     Call glVertex3f(O.v(vFor).PX, O.v(vFor).PY, O.v(vFor).PZ)
    Case GL_LINES
     If vFor <> 0 Then
      n1 = vFor - 1
      n2 = vFor
     Else
      n1 = vFor
      n2 = vFor + 1
     End If
     'Primer punto
     Call glColor4f(O.v(n1).r, O.v(n1).G, O.v(n1).b, O.v(n1).a)
     'If bTextured Then Call glTexCoord2f(O.V(n1).tU, O.V(n1).tV)
     Call glNormal3f(O.v(n1).nx, O.v(n1).ny, O.v(n1).nz)
     Call glVertex3f(O.v(n1).PX, O.v(n1).PY, O.v(n1).PZ)
     'Segundo punto
     Call glColor4f(O.v(n2).r, O.v(n2).G, O.v(n2).b, O.v(n2).a)
     'If bTextured Then Call glTexCoord2f(O.V(n2).tU, O.V(n2).tV)
     Call glNormal3f(O.v(n2).nx, O.v(n2).ny, O.v(n2).nz)
     Call glVertex3f(O.v(n2).PX, O.v(n2).PY, O.v(n2).PZ)
    Case GL_TRIANGLES
     n = vFor
     n2 = n + 1
     n3 = n + 2
     'Primer punto
     Call glColor4f(O.v(n).r, O.v(n).G, O.v(n).b, O.v(n).a)
     'If bTextured Then Call glTexCoord2f(O.V(n).tU, O.V(n).tV)
     Call glNormal3f(O.v(n).nx, O.v(n).ny, O.v(n).nz)
     Call glVertex3f(O.v(n).PX, O.v(n).PY, O.v(n).PZ)
     'Segundo punto
     Call glColor4f(O.v(n2).r, O.v(n2).G, O.v(n2).b, O.v(n2).a)
     'If bTextured Then Call glTexCoord2f(O.V(n2).tU, O.V(n2).tV)
     Call glNormal3f(O.v(n2).nx, O.v(n2).ny, O.v(n2).nz)
     Call glVertex3f(O.v(n2).PX, O.v(n2).PY, O.v(n2).PZ)
     'Tercer punto
     Call glColor4f(O.v(n3).r, O.v(n3).G, O.v(n3).b, O.v(n3).a)
     'If bTextured Then Call glTexCoord2f(O.V(n3).tU, O.V(n3).tV)
     Call glNormal3f(O.v(n3).nx, O.v(n3).ny, O.v(n3).nz)
     Call glVertex3f(O.v(n3).PX, O.v(n3).PY, O.v(n3).PZ)
    Case GL_QUADS
     n = vFor
     n2 = n + 1
     n3 = n + 2
     n4 = n + 3
     'Primer punto
     Call glColor4f(O.v(n).r, O.v(n).G, O.v(n).b, O.v(n).a)
     'If bTextured Then Call glTexCoord2f(O.V(n).tU, O.V(n).tV)
     Call glNormal3f(O.v(n).nx, O.v(n).ny, O.v(n).nz)
     Call glVertex3f(O.v(n).PX, O.v(n).PY, O.v(n).PZ)
     'Segundo punto
     Call glColor4f(O.v(n2).r, O.v(n2).G, O.v(n2).b, O.v(n2).a)
     'If bTextured Then Call glTexCoord2f(O.V(n2).tU, O.V(n2).tV)
     Call glNormal3f(O.v(n2).nx, O.v(n2).ny, O.v(n2).nz)
     Call glVertex3f(O.v(n2).PX, O.v(n2).PY, O.v(n2).PZ)
     'Tercer punto
     Call glColor4f(O.v(n3).r, O.v(n3).G, O.v(n3).b, O.v(n3).a)
     'If bTextured Then Call glTexCoord2f(O.V(n3).tU, O.V(n3).tV)
     Call glNormal3f(O.v(n3).nx, O.v(n3).ny, O.v(n3).nz)
     Call glVertex3f(O.v(n3).PX, O.v(n3).PY, O.v(n3).PZ)
     'Cuarto punto
     Call glColor4f(O.v(n4).r, O.v(n4).G, O.v(n4).b, O.v(n4).a)
     'If bTextured Then Call glTexCoord2f(O.V(n4).tU, O.V(n4).tV)
     Call glNormal3f(O.v(n4).nx, O.v(n4).ny, O.v(n4).nz)
     Call glVertex3f(O.v(n4).PX, O.v(n4).PY, O.v(n4).PZ)
    Case GL_POLYGON    'bmLineStrip,¿bmpolygon - 1 punto?
     'MinP = 3 'Se necesitan al menos 3 puntos para dibujar
    'Case gl_trianglestrip
     'MinP = 4 'Se necesitan al menos 4 puntos para dibujar
   End Select
  Next
 Call glEnd
 Call glPopMatrix
 'Call glPopName 'para modo pick
End Sub













'just draw single object
Private Sub one(ByVal a As Long)
 Call glEnable(GL_DEPTH_TEST)
 Render a
 Call glDisable(GL_DEPTH_TEST)
End Sub

'"or" is easy; simply draw both objects with depth buffering on */
Private Sub oru(ByVal a As Long, ByVal b As Long)
 Call glEnable(GL_DEPTH_TEST)
 Render a
 Render b
 Call glDisable(GL_DEPTH_TEST)
End Sub

'Set stencil buffer to show the part of a (front or back face) that's inside b's volume.
'Requirements: GL_CULL_FACE enabled, depth func GL_LESS
'Side effects: depth test, stencil func, stencil op
Private Sub firstInsideSecond(ByVal a As Long, ByVal b As Long, ByVal face As Long, ByVal test As Long)
  Call glEnable(GL_DEPTH_TEST)
  Call glColorMask(GL_FALSE, GL_FALSE, GL_FALSE, GL_FALSE)
  Call glCullFace(face)  'controls which face of a to use*/
  Render a 'draw a face of a into depth buffer */
  'use stencil plane to find parts of a in b */
  Call glDepthMask(GL_FALSE)
  Call glEnable(GL_STENCIL_TEST)
  Call glStencilFunc(GL_ALWAYS, 0, 0)
  Call glStencilOp(GL_KEEP, GL_KEEP, GL_INCR)
  Call glCullFace(GL_BACK)
  Render b 'increment the stencil where the front face of b is drawn */
  Call glStencilOp(GL_KEEP, GL_KEEP, GL_DECR)
  Call glCullFace(GL_FRONT)
  Render b 'decrement the stencil buffer where the back face of b is drawn */
  Call glDepthMask(GL_TRUE)
  Call glColorMask(GL_TRUE, GL_TRUE, GL_TRUE, GL_TRUE)
  Call glStencilFunc(test, 0, 1)
  Call glDisable(GL_DEPTH_TEST)
  Call glCullFace(face)
  Render a 'draw the part of a that's in b */
End Sub

Private Sub fixDepth(ByVal a As Long)
 Call glColorMask(GL_FALSE, GL_FALSE, GL_FALSE, GL_FALSE)
 Call glEnable(GL_DEPTH_TEST)
 Call glDisable(GL_STENCIL_TEST)
 Call glDepthFunc(GL_ALWAYS)
 Render a  'draw the front face of a, fixing the depth buffer */
 Call glDepthFunc(GL_LESS)
End Sub

'"and" two objects together */
Private Sub andI(ByVal a As Long, ByVal b As Long)
 Call firstInsideSecond(a, b, GL_BACK, GL_NOTEQUAL)
 Call fixDepth(b)
 Call firstInsideSecond(b, a, GL_BACK, GL_NOTEQUAL)
 Call glDisable(GL_STENCIL_TEST)  'reset things */
End Sub

'subtract b from a */
Private Sub subs(ByVal a As Long, ByVal b As Long)
 Call firstInsideSecond(a, b, GL_FRONT, GL_NOTEQUAL)
 Call fixDepth(b)
 Call firstInsideSecond(b, a, GL_BACK, GL_EQUAL)
 Call glDisable(GL_STENCIL_TEST)  'reset things */
End Sub

''animate scene by rotating */
'enum {ANIM_LEFT, ANIM_RIGHT}
'int animDirection = ANIM_LEFT
'
'Void Anim(Void)
'{
'  if(animDirection == ANIM_LEFT)
'    viewangle -= 3.f
'  Else
'    viewangle += 3.f
'  glutPostRedisplay()
'}
'
''special keys, like array and F keys */
'void special(int key, int x, int y)
'{
'  switch(key) {
'  Case GLUT_KEY_LEFT:
'    glutIdleFunc(anim)
'    animDirection = ANIM_LEFT
'    break
'  Case GLUT_KEY_RIGHT:
'    glutIdleFunc(anim)
'    animDirection = ANIM_RIGHT
'    break
'  Case GLUT_KEY_UP:
'  Case GLUT_KEY_DOWN:
'    glutIdleFunc(0)
'    break
'  }
'}
'
'void key(unsigned char key, int x, int y)
'{
'  switch(key) {
'  case 'a':
'    viewangle -= 10.f
'    glutPostRedisplay()
'    break
'  case 's':
'    viewangle += 10.f
'    glutPostRedisplay()
'    break
'  case '\033':
'    exit(0)
'  }
'}
'
'
'int picked_object
'int xpos = 0, ypos = 0
'int newxpos, newypos
'int startx, starty
'
'Void
'mouse(int button, int state, int x, int y)
'{
'  if(state == GLUT_UP) {
'      picked_object = button
'      xpos += newxpos
'      ypos += newypos
'      newxpos = 0
'      newypos = 0
'  } else { 'GLUT_DOWN */
'    startx = x
'    starty = y
'  }
'}
'
'void motion(int x, int y){
'  GLfloat r, objx, objy, objz
'  newxpos = x - startx
'  newypos = starty - y
'  r = (newxpos + xpos) * 50.f/512.f
'  objx = r * (float)cos(viewangle * DEGTORAD)
'  objy = (newypos + ypos) * 50.f/512.f
'  objz = r * (float)sin(viewangle * DEGTORAD)
'  switch(picked_object) {
'  Case CSG_A:
'    coneX = objx
'    coneY = objy
'    coneZ = objz
'    break
'  Case CSG_B:
'    sphereX = objx
'    sphereY = objy
'    sphereZ = objz
'    break
'  }
'  glutPostRedisplay()
'}

'Void
'main(int argc, char **argv)
'{
'    static GLfloat lightpos[] = {25.f, 50.f, -50.f, 1.f}
'    static GLfloat sphere_mat[] = {1.f, .5f, 0.f, 1.f}
'    static GLfloat cone_mat[] = {0.f, .5f, 1.f, 1.f}
'    GLUquadricObj *sphere, *cone, *base
'
'    glutInit(&argc, argv)
'    glutInitWindowSize(512, 512)
'    glutInitDisplayMode(GLUT_STENCIL|GLUT_DEPTH|GLUT_DOUBLE)
'    (void)glutCreateWindow("csg")
'    glutDisplayFunc(redraw)
'    glutKeyboardFunc(key)
'    glutSpecialFunc(special)
'    glutMouseFunc(mouse)
'    glutMotionFunc(motion)
'
'    glutCreateMenu(menu)
'    glutAddMenuEntry("A only", CSG_A)
'    glutAddMenuEntry("B only", CSG_B)
'    glutAddMenuEntry("A or B", CSG_A_OR_B)
'    glutAddMenuEntry("A and B", CSG_A_AND_B)
'    glutAddMenuEntry("A sub B", CSG_A_SUB_B)
'    glutAddMenuEntry("B sub A", CSG_B_SUB_A)
'    glutAttachMenu(GLUT_RIGHT_BUTTON)
'
'
'    glEnable(GL_CULL_FACE)
'    glEnable(GL_LIGHTING)
'    glEnable(GL_LIGHT0)
'
'    glLightfv(GL_LIGHT0, GL_POSITION, lightpos)
'    glLightModeli(GL_LIGHT_MODEL_TWO_SIDE, GL_TRUE)
'
'
'    'make display lists for sphere and cone for efficiency */
'
'    glNewList(SPHERE, GL_COMPILE)
'    sphere = gluNewQuadric()
'    glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE, sphere_mat)
'    gluSphere(sphere, 20.f, 64, 64)
'    gluDeleteQuadric(sphere)
'    glEndList()
'
'    glNewList(CONE, GL_COMPILE)
'    cone = gluNewQuadric()
'    base = gluNewQuadric()
'    glMaterialfv(GL_FRONT_AND_BACK, GL_AMBIENT_AND_DIFFUSE, cone_mat)
'    gluQuadricOrientation(base, GLU_INSIDE)
'    gluDisk(base, 0., 15., 64, 1)
'    gluCylinder(cone, 15., 0., 60., 64, 64)
'    gluDeleteQuadric(cone)
'    gluDeleteQuadric(base)
'    glEndList()
'
'    glMatrixMode(GL_PROJECTION)
'    glOrtho(-50., 50., -50., 50., -50., 50.)
'    glMatrixMode(GL_MODELVIEW)
'    glutMainLoop()
'}



