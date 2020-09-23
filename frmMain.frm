VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   BackColor       =   &H00404040&
   BorderStyle     =   0  'None
   ClientHeight    =   9000
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12000
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "frmMain"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   600
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   800
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Background 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   14400
      Left            =   2100
      ScaleHeight     =   960
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   520
      TabIndex        =   0
      Top             =   -5400
      Width           =   7800
      Begin MSComctlLib.ImageList BadShip 
         Index           =   1
         Left            =   960
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   148
         ImageHeight     =   170
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":0000
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   2
         Left            =   1560
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   119
         ImageHeight     =   206
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1272C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   3
         Left            =   2160
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   50
         ImageHeight     =   87
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":24930
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   4
         Left            =   2760
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   88
         ImageHeight     =   87
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":27D2C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   5
         Left            =   3360
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   36
         ImageHeight     =   80
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2D738
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":2F94C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":31B60
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList GoodShip 
         Index           =   0
         Left            =   360
         Top             =   12120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   87
         ImageHeight     =   89
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":33D74
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":39990
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList GoodShip 
         Index           =   1
         Left            =   960
         Top             =   12120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   87
         ImageHeight     =   89
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":3F5AC
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":451C8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   6
         Left            =   3960
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   53
         ImageHeight     =   145
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":4ADE4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   7
         Left            =   4560
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   198
         ImageHeight     =   87
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":508D8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   8
         Left            =   5160
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   85
         ImageHeight     =   82
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":5D3B8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   9
         Left            =   5760
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   92
         ImageHeight     =   65
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   2
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6260C
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":66C74
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   11
         Left            =   6960
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   74
         ImageHeight     =   82
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6B2DC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Effect 
         Index           =   0
         Left            =   360
         Top             =   10680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   82
         ImageHeight     =   82
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":6FAF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":74AB4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":79A78
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":7EA3C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":83A00
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":889C4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Effect 
         Index           =   1
         Left            =   960
         Top             =   10680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   107
         ImageHeight     =   107
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":8D988
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":96148
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":9E908
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":A70C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":AF888
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":B8048
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   0
         Left            =   360
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   115
         ImageHeight     =   132
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":C0808
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   0
         Left            =   360
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   3
         ImageHeight     =   3
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CBBCC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   10
         Left            =   6360
         Top             =   12840
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   69
         ImageHeight     =   69
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CBC44
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":CF4A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D2D0C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   1
         Left            =   960
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   30
         ImageHeight     =   30
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D6570
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   2
         Left            =   1560
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   8
         ImageHeight     =   25
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D708C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   3
         Left            =   2160
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   8
         ImageHeight     =   29
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D7338
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   4
         Left            =   2760
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   4
         ImageHeight     =   4
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D7644
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Effect 
         Index           =   2
         Left            =   1560
         Top             =   10680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   122
         ImageHeight     =   122
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":D76C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":E267C
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":ED630
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":F85E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":103598
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":10E54C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   12
         Left            =   360
         Top             =   13560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   112
         ImageHeight     =   103
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":119500
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Effect 
         Index           =   3
         Left            =   2160
         Top             =   10680
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   24
         ImageHeight     =   24
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   6
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":121C84
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":122398
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":122AAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1231C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1238D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":123FE8
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   13
         Left            =   960
         Top             =   13560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   67
         ImageHeight     =   58
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1246FC
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   5
         Left            =   3360
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   8
         ImageHeight     =   8
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":127588
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   6
         Left            =   3960
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   3
         ImageHeight     =   25
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12769C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   7
         Left            =   4560
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   5
         ImageHeight     =   19
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":12781C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   14
         Left            =   1560
         Top             =   13560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   118
         ImageHeight     =   118
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":1279A0
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   15
         Left            =   2160
         Top             =   13560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   81
         ImageHeight     =   81
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":131E0C
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList BadShip 
         Index           =   16
         Left            =   2760
         Top             =   13560
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   40
         ImageHeight     =   40
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":136B94
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ImageList Shot 
         Index           =   8
         Left            =   5160
         Top             =   11400
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   0
         ImageWidth      =   40
         ImageHeight     =   10
         MaskColor       =   16777215
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmMain.frx":137EA8
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Shape Border 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   3
      Left            =   120
      Top             =   600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Shape Border 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   375
      Index           =   2
      Left            =   9960
      Top             =   600
      Width           =   1935
   End
   Begin VB.Label lblBlock 
      Alignment       =   2  'Center
      BackColor       =   &H008D5327&
      Caption         =   "Press F1 to Start"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblBlock 
      Alignment       =   2  'Center
      BackColor       =   &H008D5327&
      Caption         =   "Press F2 to Start"
      BeginProperty Font 
         Name            =   "Franklin Gothic Medium Cond"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Border 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      Height          =   375
      Index           =   0
      Left            =   9960
      Top             =   120
      Width           =   1935
   End
   Begin VB.Shape Border 
      BorderColor     =   &H000080FF&
      BorderWidth     =   2
      Height          =   375
      Index           =   1
      Left            =   120
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Player 2"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblPlayer 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "Player 1"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   5
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label lblC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   9960
      TabIndex        =   2
      Top             =   600
      Width           =   375
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   375
      Index           =   0
      Left            =   10320
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00FF0000&
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   480
      TabIndex        =   4
      Top             =   600
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label lblC 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FF0000&
      Caption         =   "$"
      BeginProperty Font 
         Name            =   "OCR A Extended"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000080FF&
      Height          =   375
      Index           =   1
      Left            =   120
      TabIndex        =   3
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetTickCount Lib "kernel32" () As Long
Dim time1 As Long
Dim time2 As Long
Dim timer As Long

Private Type PlayerShip
    ShipType As Integer
    PosX As Integer
    PosY As Integer
    Speed As Integer
    HitPoints As Integer
    Money As Long
    Age As Long
    IsLeft As Boolean
    IsUp As Boolean
    IsRight As Boolean
    IsDown As Boolean
    IsShooting As Boolean
End Type

Private Type BadShip
    ShipType As Integer
    PosX As Integer
    PosY As Integer
    Speed As Integer
    SpeedX As Integer
    SpeedY As Integer
    HitPoints As Integer
    Cash As Integer
    Age As Long
    Frames As Integer
    CurFrame As Integer
    Weapon(1) As Integer
    MoveX(3) As Integer
    MoveY(3) As Integer
End Type

Private Type Projectile
    Owner As Integer
    ShotType As Integer
    PosX As Integer
    PosY As Integer
    SpeedX As Integer
    SpeedY As Integer
    Attack As Integer
End Type

Private Type Effect
    EffectType As Integer
    PosX As Integer
    PosY As Integer
    Age As Long
End Type

Private Type Bg
    Speed As Double
    PosX As Integer
    PosY As Integer
End Type

Const screenTop = 360
Const numPlayers = 1
Const numBad = 15
Const numShots = 30
Const numEffects = 60
Const numBg = 200

Dim paused As Boolean
Dim num As Integer
Dim num2 As Integer
Dim num3 As Integer

Dim p(numPlayers) As PlayerShip
Dim enemy(numBad) As BadShip
Dim bullet(numShots) As Projectile
Dim explosion(numEffects) As Effect
Dim star(numBg) As Bg

Private Sub Background_KeyDown(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyLeft
    p(0).IsLeft = True
    p(0).IsRight = False
Case vbKeyUp
    p(0).IsUp = True
    p(0).IsDown = False
Case vbKeyRight
    p(0).IsLeft = False
    p(0).IsRight = True
Case vbKeyDown
    p(0).IsUp = False
    p(0).IsDown = True
Case vbKeyReturn
    p(0).IsShooting = True

Case 65
    p(1).IsLeft = True
Case 87
    p(1).IsUp = True
Case 68
    p(1).IsRight = True
Case 83
    p(1).IsDown = True
Case vbKeySpace
    p(1).IsShooting = True

Case 19
    If paused = False Then
        paused = True
    Else
        paused = False
    End If
Case vbKeyF1
    If p(0).Speed = 0 Then Call NewPlayer(0, 7, 100, -1)
Case vbKeyF2
    If p(1).Speed = 0 Then Call NewPlayer(1, 7, 100, -1)
Case vbKeyEscape
    End
End Select

End Sub

Private Sub Background_KeyUp(KeyCode As Integer, Shift As Integer)

Select Case KeyCode
Case vbKeyLeft
    p(0).IsLeft = False
Case vbKeyUp
    p(0).IsUp = False
Case vbKeyRight
    p(0).IsRight = False
Case vbKeyDown
    p(0).IsDown = False
Case vbKeyReturn
    p(0).IsShooting = False

Case vbKeyA
    p(1).IsLeft = False
Case vbKeyW
    p(1).IsUp = False
Case vbKeyD
    p(1).IsRight = False
Case vbKeyS
    p(1).IsDown = False
Case vbKeySpace
    p(1).IsShooting = False
End Select

End Sub

Private Sub Form_Activate()

paused = False

Call NewPlayer(0, 7, 100, -1)

Do
    DoEvents
    If paused = False Then
        Do
            time2 = GetTickCount
            If time2 - time1 >= 1 Then GameLoop
        Loop Until paused = True
    End If
Loop

End Sub

Private Sub Form_Unload(Cancel As Integer)

End

End Sub

Private Sub GameLoop()

DoEvents
Background.Cls
Randomize

'----------------------------------------
'   Background
'----------------------------------------
For num = 0 To numBg
    If star(num).PosY = 0 Then
        star(num).Speed = Rnd * 5 + 1
        star(num).PosX = Int(Rnd * Background.Width)
        star(num).PosY = screenTop + Int(Rnd * Background.Height)
    End If

    star(num).PosY = star(num).PosY + star(num).Speed
    If star(num).PosY > Background.Height Then star(num).PosY = Int(Rnd * 360)
    Background.PSet (star(num).PosX, star(num).PosY), vbWhite
Next num

'----------------------------------------
'   New objects
'----------------------------------------
timer = timer + 1
Select Case timer
Case 31: Call NewBad(4, 16, 0, 0, 0, 0, 0, 0, 0, 0)
Case 43: Call NewBad(4, 116, 0, 0, 0, 0, 0, 0, 0, 0)
Case 55: Call NewBad(4, 216, 0, 0, 0, 0, 0, 0, 0, 0)
Case 67: Call NewBad(4, 316, 0, 0, 0, 0, 0, 0, 0, 0)
Case 79: Call NewBad(4, 416, 0, 0, 0, 0, 0, 0, 0, 0)

Case 150
    Call NewBad(10, 16, 0, 0, 0, 0, 0, 0, 0, 0)
    Call NewBad(10, 116, 0, 0, 0, 0, 0, 0, 0, 0)
    Call NewBad(10, 216, 0, 0, 0, 0, 0, 0, 0, 0)
    Call NewBad(10, 316, 0, 0, 0, 0, 0, 0, 0, 0)
    Call NewBad(10, 416, 0, 0, 0, 0, 0, 0, 0, 0)

Case Is < 600
    If timer > 250 And timer Mod 15 = 0 Then
        Call NewBad(Int(Rnd * 3) + 14, Int(Rnd * 500), 0, 0, 0, 0, 0, 0, 0, 0)
    End If

Case Is > 700
    If timer Mod 30 = 0 Then
        Call NewBad(Int(Rnd * 14), Int(Rnd * 500), 10 * (Int(Rnd * 2) - 1), _
        31 * Int(Rnd * 2), 50 * (Int(Rnd * 2) - 1), _
        71 * Int(Rnd * 2), 10 * Int(Rnd * 2), 24 * Int(Rnd * 2), _
        37 * Int(Rnd * 2), 48 * Int(Rnd * 2))
    End If
Case Is > 500000: End
End Select

'----------------------------------------
'   Enemy ships
'----------------------------------------
For num = 0 To numBad
    If enemy(num).PosY > 0 Then
        'Shooting
        enemy(num).Age = enemy(num).Age + 1
        For num2 = 0 To 1
            Select Case enemy(num).Weapon(num2)
            Case 4
                If enemy(num).Age Mod 50 = 0 Then Call NewShot(num, 4, 1 / 4)
                If enemy(num).Age Mod 50 = 0 Then Call NewShot(num, 4, 3 / 4)
            Case 5
                If enemy(num).Age Mod 60 = 0 Then Call NewShot(num, 5, 1 / 2)
            Case 6
                If enemy(num).Age Mod 80 = 0 Then
                    Call NewShot(num, 6, 1 / 6)
                    Call NewShot(num, 6, 2 / 6)
                    Call NewShot(num, 6, 3 / 6)
                    Call NewShot(num, 6, 4 / 6)
                    Call NewShot(num, 6, 5 / 6)
                End If
            Case 7
                If enemy(num).Age Mod 70 = 0 Then
                    Call NewShot(num, 7, 1 / 3)
                    Call NewShot(num, 7, 2 / 3)
                End If
            Case 8
                If enemy(num).Age Mod 50 = 0 Then Call NewShot(num, 8, 1 / 2)
            End Select
        Next num2

        'Animates ship
        If timer Mod 4 = 0 Then
            enemy(num).CurFrame = enemy(num).CurFrame + 1
            If enemy(num).CurFrame > enemy(num).Frames Then enemy(num).CurFrame = 1
        End If

        'Draws ship
        BadShip(enemy(num).ShipType).ListImages(enemy(num).CurFrame).Draw Background.hDC, enemy(num).PosX, enemy(num).PosY, 1

        'Sets player ship's new coordinates
        enemy(num).PosX = enemy(num).PosX + enemy(num).SpeedX
        enemy(num).PosY = enemy(num).PosY + enemy(num).SpeedY

        For num2 = 0 To 3
            If enemy(num).MoveX(num2) = enemy(num).Age Then
                enemy(num).SpeedX = enemy(num).Speed
                If enemy(num).MoveX(num2) Mod 10 = 1 Then enemy(num).SpeedX = 0
            End If
            If -enemy(num).MoveX(num2) = enemy(num).Age Then
                enemy(num).SpeedX = -enemy(num).Speed
                If enemy(num).MoveX(num2) Mod 10 = 1 Then enemy(num).SpeedX = 0
            End If

            If enemy(num).MoveY(num2) = enemy(num).Age Then
                enemy(num).SpeedY = enemy(num).Speed
                If enemy(num).MoveY(num2) Mod 10 = 1 Then enemy(num).SpeedY = 0
            End If
            If -enemy(num).MoveY(num2) = enemy(num).Age Then
                enemy(num).SpeedY = -enemy(num).Speed
                If enemy(num).MoveY(num2) Mod 10 = 1 Then enemy(num).SpeedY = 0
            End If
        Next num2

        'Keeps ship on screen horizontally
        Select Case enemy(num).PosX
        Case Is < 0
            enemy(num).PosX = 0
            enemy(num).SpeedX = -enemy(num).SpeedX
        Case Is > Background.Width - BadShip(enemy(num).ShipType).ImageWidth
            enemy(num).PosX = Background.Width - BadShip(enemy(num).ShipType).ImageWidth
            enemy(num).SpeedX = -enemy(num).SpeedX
        End Select

        'Resets ship off screen
        If enemy(num).PosY > Background.Height Then KillBad (num)
    End If
Next num

'----------------------------------------
'   Player ships
'----------------------------------------
For num = 0 To numPlayers
    If p(num).Speed > 0 Then
        'Kills ship
        If p(num).HitPoints <= 0 And p(num).Speed > 0 Then
            Call NewEffect(2, p(num).PosX, p(num).PosY)
            p(num).Speed = 0
            p(num).IsShooting = False

            lblBlock(num).Visible = True
            lblPlayer(num).Visible = False
            lblCash(num).Visible = False
            lblC(num).Visible = False
            Border(num).Visible = False
            Border(num + 2).Visible = False
        End If

        'Shooting
        p(num).Age = p(num).Age + 1
        If p(num).IsShooting = True Then
            If p(num).Age Mod 3 = 0 Then
                Call NewShot(num, 0, 1 / 2)
            End If
            If p(num).Age Mod 12 = 0 Then
                Call NewShot(num, 3, 1 / 5)
                Call NewShot(num, 3, 4 / 5)
            End If
        End If

        'Draws ship
        If p(num).Age >= 50 Then
            GoodShip(p(num).ShipType).ListImages(1).Draw Background.hDC, p(num).PosX, p(num).PosY, 1
        Else
            If p(num).Age Mod 5 > 1 And p(num).Age Mod 5 < 5 Then
                GoodShip(p(num).ShipType).ListImages(1).Draw Background.hDC, p(num).PosX, p(num).PosY, 1
            Else
                GoodShip(p(num).ShipType).ListImages(2).Draw Background.hDC, p(num).PosX, p(num).PosY, 1
            End If
        End If

        'Sets player ship's new coordinates
        If p(num).IsLeft = True Then p(num).PosX = p(num).PosX - p(num).Speed
        If p(num).IsUp = True Then p(num).PosY = p(num).PosY - p(num).Speed
        If p(num).IsRight = True Then p(num).PosX = p(num).PosX + p(num).Speed
        If p(num).IsDown = True Then p(num).PosY = p(num).PosY + p(num).Speed

        'Keeps ship on screen horizontally
        Select Case p(num).PosX
        Case Is < 0
            p(num).PosX = 0
        Case Is > Background.Width - GoodShip(p(num).ShipType).ImageWidth
            p(num).PosX = Background.Width - GoodShip(p(num).ShipType).ImageWidth
        End Select

        'Keeps ship on screen vertically
        Select Case p(num).PosY
        Case Is < screenTop
            If p(num).HitPoints Then p(num).PosY = screenTop
        Case Is > Background.Height - GoodShip(p(num).ShipType).ImageHeight
            p(num).PosY = Background.Height - GoodShip(p(num).ShipType).ImageHeight
        End Select
    End If
Next num

'----------------------------------------
'   Collisions
'----------------------------------------
For num = 0 To numPlayers
    If p(num).Speed > 0 And p(num).Age > 50 Then
        For num2 = 0 To numBad
            If p(num).PosX + GoodShip(p(num).ShipType).ImageWidth > enemy(num2).PosX And _
            p(num).PosX < enemy(num2).PosX + BadShip(enemy(num2).ShipType).ImageWidth And _
            p(num).PosY + GoodShip(p(num).ShipType).ImageHeight > enemy(num2).PosY And _
            p(num).PosY < enemy(num2).PosY + BadShip(enemy(num2).ShipType).ImageHeight Then
                p(num).HitPoints = p(num).HitPoints - 30
                enemy(num2).HitPoints = enemy(num2).HitPoints - 30
                Call NewEffect(1, (p(num).PosX + enemy(num2).PosX) / 2, (p(num).PosY + enemy(num2).PosY) / 2)

                If enemy(num2).HitPoints <= 0 Then
                    Call NewEffect(2, enemy(num2).PosX, enemy(num2).PosY)
                    KillBad (num2)
                    p(num).Money = p(num).Money + enemy(num2).Cash
                    lblCash(num).Caption = p(num).Money
                End If
            End If
        Next num2
    End If
Next num

'----------------------------------------
'   Bullets
'----------------------------------------
For num = 0 To numShots
    If bullet(num).PosY > 0 Then
        If bullet(num).SpeedY < 0 Then
            'Bullet damage to enemies
            For num2 = 0 To numBad
                If bullet(num).PosY <= enemy(num2).PosY + BadShip(enemy(num2).ShipType).ImageHeight And _
                bullet(num).PosY + Shot(bullet(num).ShotType).ImageHeight >= enemy(num2).PosY And _
                bullet(num).PosX + Shot(bullet(num).ShotType).ImageWidth >= enemy(num2).PosX And _
                bullet(num).PosX <= enemy(num2).PosX + BadShip(enemy(num2).ShipType).ImageWidth Then
                    enemy(num2).HitPoints = enemy(num2).HitPoints - bullet(num).Attack
                    Call NewEffect(3, Abs(bullet(num).PosX - (Effect(0).ImageWidth / 2)), bullet(num).PosY)
                    KillShot (num)
                End If
                If enemy(num2).HitPoints <= 0 And enemy(num2).PosY > 0 Then
                    Call NewEffect(2, enemy(num2).PosX, enemy(num2).PosY)
                    KillBad (num2)
                    p(bullet(num).Owner).Money = p(bullet(num).Owner).Money + enemy(num2).Cash
                    lblCash(bullet(num).Owner).Caption = p(bullet(num).Owner).Money
                End If
            Next num2
        Else
            'Bullet damage to players
            For num2 = 0 To numPlayers
                If p(num2).Speed > 0 And p(num2).Age > 50 And bullet(num).PosY > 0 Then
                    If bullet(num).PosY <= p(num2).PosY + GoodShip(p(num2).ShipType).ImageHeight And _
                    bullet(num).PosY + Shot(bullet(num).ShotType).ImageHeight >= p(num2).PosY And _
                    bullet(num).PosX + Shot(bullet(num).ShotType).ImageWidth >= p(num2).PosX And _
                    bullet(num).PosX <= p(num2).PosX + GoodShip(p(num2).ShipType).ImageWidth Then
                        p(num2).HitPoints = p(num2).HitPoints - bullet(num).Attack
                        Call NewEffect(0, Abs(bullet(num).PosX - (Effect(0).ImageWidth / 2)), bullet(num).PosY)
                        KillShot (num)
                    End If
                End If
            Next num2
        End If

        'Draws bullet and sets new coordinates
        Shot(bullet(num).ShotType).ListImages(1).Draw Background.hDC, bullet(num).PosX, bullet(num).PosY, 1

        bullet(num).PosX = bullet(num).PosX + bullet(num).SpeedX
        bullet(num).PosY = bullet(num).PosY + bullet(num).SpeedY

        If bullet(num).PosY < screenTop - Shot(bullet(num).ShotType).ImageHeight Or _
        bullet(num).PosY > Background.Height Then
            KillShot (num)
        End If
    End If
Next num

'Player HP bars
If p(0).HitPoints > 0 Then
    frmMain.Line (715, 350)-(745, 540), &HC0C0C0, BF
    frmMain.Line (715, 540 - (p(0).HitPoints * 1.9))-(745, 540), vbYellow, BF
Else
    frmMain.Line (715, 350)-(745, 540), &H404040, BF
End If
If p(1).HitPoints > 0 Then
    frmMain.Line (55, 350)-(85, 540), &HC0C0C0, BF
    frmMain.Line (55, 540 - (p(1).HitPoints * 1.9))-(85, 540), vbYellow, BF
Else
    frmMain.Line (55, 350)-(85, 540), &H404040, BF
End If

'----------------------------------------
'   Effects
'----------------------------------------
For num = 0 To numEffects
    If explosion(num).PosY > 0 Then
        explosion(num).Age = explosion(num).Age + 1
        If explosion(num).Age >= 6 Then KillEffect (num)
        Effect(explosion(num).EffectType).ListImages(explosion(num).Age).Draw Background.hDC, explosion(num).PosX, explosion(num).PosY, 1
    End If
Next num

time1 = GetTickCount

End Sub

Private Sub KillBad(shipNum As Integer)

enemy(shipNum).PosY = 0
enemy(shipNum).Speed = 0
enemy(shipNum).SpeedX = 0
enemy(shipNum).SpeedY = 0
enemy(shipNum).HitPoints = 0
enemy(shipNum).Age = 0

End Sub

Private Sub KillEffect(effectNum As Integer)

explosion(effectNum).PosY = 0

End Sub

Private Sub KillShot(shotNum As Integer)

bullet(shotNum).PosX = Background.Width
bullet(shotNum).PosY = 0
bullet(shotNum).SpeedX = 0
bullet(shotNum).SpeedY = 0

End Sub

Private Sub NewBad(newShipType As Integer, newPosX As Integer, _
x1 As Integer, x2 As Integer, x3 As Integer, x4 As Integer, _
y1 As Integer, y2 As Integer, y3 As Integer, y4 As Integer)

Dim shipNum As Integer

shipNum = -1

Do
    shipNum = shipNum + 1
Loop Until shipNum = numBad Or enemy(shipNum).PosY = 0

enemy(shipNum).ShipType = newShipType

Select Case enemy(shipNum).ShipType
Case 0
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 60
    enemy(shipNum).Cash = 700: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 4: enemy(shipNum).Weapon(1) = 7
Case 1
    enemy(shipNum).Speed = 4: enemy(shipNum).HitPoints = 50
    enemy(shipNum).Cash = 650: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 2
    enemy(shipNum).Speed = 4: enemy(shipNum).HitPoints = 50
    enemy(shipNum).Cash = 600: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 3
    enemy(shipNum).Speed = 6: enemy(shipNum).HitPoints = 10
    enemy(shipNum).Cash = 100: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 4: enemy(shipNum).Weapon(1) = -1
Case 4
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 25
    enemy(shipNum).Cash = 450: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 5
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 15
    enemy(shipNum).Cash = 150: enemy(shipNum).Frames = 3
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 6
    enemy(shipNum).Speed = 2: enemy(shipNum).HitPoints = 25
    enemy(shipNum).Cash = 400: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 7
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 40
    enemy(shipNum).Cash = 650: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 8: enemy(shipNum).Weapon(1) = -1
Case 8
    enemy(shipNum).Speed = 6: enemy(shipNum).HitPoints = 20
    enemy(shipNum).Cash = 500: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 9
    enemy(shipNum).Speed = 5: enemy(shipNum).HitPoints = 18
    enemy(shipNum).Cash = 300: enemy(shipNum).Frames = 2
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 10
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 25
    enemy(shipNum).Cash = 500: enemy(shipNum).Frames = 3
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 11
    enemy(shipNum).Speed = 5: enemy(shipNum).HitPoints = 20
    enemy(shipNum).Cash = 400: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 12
    enemy(shipNum).Speed = 3: enemy(shipNum).HitPoints = 45
    enemy(shipNum).Cash = 550: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 5: enemy(shipNum).Weapon(1) = -1
Case 13
    enemy(shipNum).Speed = 2: enemy(shipNum).HitPoints = 300
    enemy(shipNum).Cash = 5000: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = 6: enemy(shipNum).Weapon(1) = -1
Case 14
    enemy(shipNum).Speed = Int(Rnd * 4) + 2: enemy(shipNum).HitPoints = 85
    enemy(shipNum).Cash = 200: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = -1: enemy(shipNum).Weapon(1) = -1
Case 15
    enemy(shipNum).Speed = Int(Rnd * 5) + 2: enemy(shipNum).HitPoints = 65
    enemy(shipNum).Cash = 150: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = -1: enemy(shipNum).Weapon(1) = -1
Case 16
    enemy(shipNum).Speed = Int(Rnd * 5) + 3: enemy(shipNum).HitPoints = 45
    enemy(shipNum).Cash = 100: enemy(shipNum).Frames = 1
    enemy(shipNum).Weapon(0) = -1: enemy(shipNum).Weapon(1) = -1
End Select

enemy(shipNum).SpeedX = 0
enemy(shipNum).SpeedY = enemy(shipNum).Speed

Select Case newPosX
Case Is < 0
    enemy(shipNum).PosX = Background.Width / 2 - BadShip(enemy(shipNum).ShipType).ImageWidth / 2
Case Else
    enemy(shipNum).PosX = newPosX
End Select
enemy(shipNum).PosY = screenTop - BadShip(enemy(shipNum).ShipType).ImageHeight

enemy(shipNum).MoveX(0) = x1
enemy(shipNum).MoveX(1) = x2
enemy(shipNum).MoveX(2) = x3
enemy(shipNum).MoveX(3) = x4

enemy(shipNum).MoveY(0) = y1
enemy(shipNum).MoveY(1) = y2
enemy(shipNum).MoveY(2) = y3
enemy(shipNum).MoveY(3) = y4

enemy(shipNum).Age = Int(Rnd * 100)
enemy(shipNum).CurFrame = Int(Rnd * enemy(shipNum).Frames) + 1

End Sub

Private Sub NewEffect(newEffectType As Integer, newPosX As Integer, newPosY As Integer)

Dim effectNum As Integer

effectNum = -1
Do
    effectNum = effectNum + 1
Loop Until effectNum = numEffects Or explosion(effectNum).PosY = 0

explosion(effectNum).EffectType = newEffectType
explosion(effectNum).PosX = newPosX
explosion(effectNum).PosY = newPosY
explosion(effectNum).Age = 0

End Sub

Private Sub NewShot(newOwner As Integer, newShotType As Integer, allignment As Double)

Dim shotNum As Integer

shotNum = -1
Do
    shotNum = shotNum + 1
Loop Until shotNum = numShots Or bullet(shotNum).PosY = 0

bullet(shotNum).Owner = newOwner
bullet(shotNum).ShotType = newShotType

Select Case bullet(shotNum).ShotType
Case Is < 4
    bullet(shotNum).PosY = p(newOwner).PosY - Shot(bullet(shotNum).ShotType).ImageHeight - 5
Case Is >= 4
    bullet(shotNum).PosY = enemy(newOwner).PosY + BadShip(enemy(newOwner).ShipType).ImageHeight + Shot(bullet(shotNum).ShotType).ImageHeight + 5
End Select

Select Case bullet(shotNum).ShotType
'Machine gun
Case 0
    bullet(shotNum).Attack = 3
    bullet(shotNum).SpeedX = Int(Rnd * 4) - 2
    bullet(shotNum).SpeedY = -30
    bullet(shotNum).PosX = p(newOwner).PosX + ((GoodShip(p(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Red box
Case 1
    bullet(shotNum).Attack = 6
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = -21
    bullet(shotNum).PosX = p(newOwner).PosX + ((GoodShip(p(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Gray and red missile
Case 2
    bullet(shotNum).Attack = 5
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = -22
    bullet(shotNum).PosX = p(newOwner).PosX + ((GoodShip(p(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Blue and yellow missile
Case 3
    bullet(shotNum).Attack = 6
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = -22
    bullet(shotNum).PosX = p(newOwner).PosX + ((GoodShip(p(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)

'Yellow bullet
Case 4
    bullet(shotNum).Attack = 3
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = 11
    bullet(shotNum).PosX = enemy(newOwner).PosX + ((BadShip(enemy(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Fireball
Case 5
    bullet(shotNum).Attack = 5
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = 9
    bullet(shotNum).PosX = enemy(newOwner).PosX + ((BadShip(enemy(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Gray and red missile
Case 6
    bullet(shotNum).Attack = 7
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = 9
    bullet(shotNum).PosX = enemy(newOwner).PosX + ((BadShip(enemy(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Yellow and red missile
Case 7
    bullet(shotNum).Attack = 6
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = 9
    bullet(shotNum).PosX = enemy(newOwner).PosX + ((BadShip(enemy(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
'Yellow and red arc
Case 8
    bullet(shotNum).Attack = 8
    bullet(shotNum).SpeedX = 0
    bullet(shotNum).SpeedY = 5
    bullet(shotNum).PosX = enemy(newOwner).PosX + ((BadShip(enemy(newOwner).ShipType).ImageWidth - _
    Shot(bullet(shotNum).ShotType).ImageWidth) * allignment)
End Select

End Sub

Private Sub NewPlayer(shipNum As Integer, newSpeed As Integer, newHP As Integer, newPosX As Integer)

p(shipNum).ShipType = shipNum
p(shipNum).Speed = newSpeed
p(shipNum).HitPoints = newHP
p(shipNum).Money = 0
Select Case newPosX
Case Is < 0
    p(shipNum).PosX = Background.Width / 2 - GoodShip(p(shipNum).ShipType).ImageWidth / 2
Case Else
    p(shipNum).PosX = newPosX
End Select
p(shipNum).PosY = Background.Height - GoodShip(p(shipNum).ShipType).ImageHeight
p(shipNum).Age = 0

lblBlock(shipNum).Visible = False
lblPlayer(shipNum).Visible = True
lblCash(shipNum).Visible = True
lblC(shipNum).Visible = True
Border(shipNum).Visible = True
Border(shipNum + 2).Visible = True

lblCash(shipNum).Caption = p(shipNum).Money

End Sub
