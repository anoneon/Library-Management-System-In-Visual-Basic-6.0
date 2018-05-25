VERSION 5.00
Begin VB.Form listUser 
   Caption         =   "ADMIN"
   ClientHeight    =   7170
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11715
   LinkTopic       =   "Form1"
   ScaleHeight     =   7170
   ScaleWidth      =   11715
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "listUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    Me.Icon = LoadPicture(App.Path & "\images\lbm_ico.ico")
End Sub
