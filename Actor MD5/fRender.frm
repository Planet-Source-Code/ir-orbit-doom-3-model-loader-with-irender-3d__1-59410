VERSION 5.00
Begin VB.Form fRender 
   Caption         =   "Form1"
   ClientHeight    =   3900
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6000
   LinkTopic       =   "Form1"
   ScaleHeight     =   260
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "fRender"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'--------------------------------------------------------------'
' Tutorial 32 - Doom III Model Loder                           '
'--------------------------------------------------------------'
' iRender 3D supports md5 models with animation (md5mesh,      '
' md5anim). This models have very good mesh quality and        '
' real-looking animation. To make rendering faster code is     '
' written with C++ and comes with iRender with separate plugin '
'--------------------------------------------------------------'
' Note: Check that plugin ActorMD5Plug.dll is in the same      '
' folder that program is or in the windows system folder.      '
'--------------------------------------------------------------'
' OrBit, AKiR (c) 2004-2005                                    '
'--------------------------------------------------------------'

Option Explicit


' Main engine class. Provides access to graphic functions.
Dim iR          As New iR_Engine
' Control class for keyboard and mouse
Dim Control     As New iR_Control
' Camera class needs for every 3D aplication
Dim Camera      As New iR_Camera

' MD5 Actor Class
Dim Model       As New iR_ActorMD5


Private Sub Form_Load()

    ' Show initialization dialog
    If Not iR.InitWithDialog(Me.hWnd) Then End
    ' Show rendering form
    Me.Show
    
    ' Tells the engine that we will write log file
    iR.SetLogging 1
    ' Show fps
    iR.SetDisplayFPS 1
    
    ' Setups view frustum
    iR.SetViewFrustum 5000, 90, 1
    
    ' Set gray color for back ground
    iR.SetBackGroundColor RGB(55, 55, 55)
    
    ' Set texture filtering
    iR.SetTextureFilter iR_Filter_Bilinear
    
    ' Load model
    Model.LoadActor "md5 imp\imp.md5mesh"
    ' Load texture and set its index to the model
    Model.SetTexture iR.LoadTexture("md5 imp\imp.jpg")

    
    ' Put camera so we can see the model
    Camera.SetPosition iR.CreateVec3(0, 50, -100)
    
    'Main rendering loop
    Do
        DoEvents
        
        ' Begin our 3D scene and clear viewport
        iR.BeginScene
        iR.Clear
        
        ' End program if Espace is pressed
        If Control.CheckKBKeyPressed(iR_Key_Escape) Then End
        
        ' Update camera
        Camera.Update
        
        ' Turn model
        Model.SetRotation iR.CreateVec3(0, iR.GetTickPassed / 1000, 0)
        ' Animate model
        Model.SetFrame CLng(iR.GetTickPassed / 30)
        
        ' Render model
        Model.Render
    
        ' Flip buffers and finish 3D scene
        iR.Present
        iR.EndScene
    Loop
    
End Sub



