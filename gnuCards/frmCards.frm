VERSION 5.00
Begin VB.Form frmCards 
   Caption         =   "Form1"
   ClientHeight    =   7965
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7365
   LinkTopic       =   "Form1"
   ScaleHeight     =   7965
   ScaleWidth      =   7365
   StartUpPosition =   3  'Windows Default
   Begin VB.Image imgBack 
      Height          =   1950
      Index           =   1
      Left            =   5880
      Picture         =   "frmCards.frx":0000
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgBack 
      Height          =   1950
      Index           =   0
      Left            =   4380
      Picture         =   "frmCards.frx":205A
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   51
      Left            =   2910
      Picture         =   "frmCards.frx":43FB
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   50
      Left            =   2670
      Picture         =   "frmCards.frx":5CD7
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   49
      Left            =   2430
      Picture         =   "frmCards.frx":75E6
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   48
      Left            =   2190
      Picture         =   "frmCards.frx":8E1C
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   47
      Left            =   1950
      Picture         =   "frmCards.frx":A768
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   46
      Left            =   1710
      Picture         =   "frmCards.frx":BFDE
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   45
      Left            =   1470
      Picture         =   "frmCards.frx":D856
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   44
      Left            =   1230
      Picture         =   "frmCards.frx":F082
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   43
      Left            =   990
      Picture         =   "frmCards.frx":10949
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   42
      Left            =   750
      Picture         =   "frmCards.frx":121F7
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   41
      Left            =   510
      Picture         =   "frmCards.frx":13A36
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   40
      Left            =   270
      Picture         =   "frmCards.frx":152DF
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   39
      Left            =   30
      Picture         =   "frmCards.frx":16B7B
      Top             =   5970
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   38
      Left            =   2910
      Picture         =   "frmCards.frx":18458
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   37
      Left            =   2670
      Picture         =   "frmCards.frx":19CE6
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   36
      Left            =   2430
      Picture         =   "frmCards.frx":1B58C
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   35
      Left            =   2190
      Picture         =   "frmCards.frx":1CD82
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   34
      Left            =   1950
      Picture         =   "frmCards.frx":1E65E
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   33
      Left            =   1710
      Picture         =   "frmCards.frx":1FE88
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   32
      Left            =   1470
      Picture         =   "frmCards.frx":216C8
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   31
      Left            =   1230
      Picture         =   "frmCards.frx":22EB3
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   30
      Left            =   990
      Picture         =   "frmCards.frx":246E6
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   29
      Left            =   750
      Picture         =   "frmCards.frx":25EF5
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   28
      Left            =   510
      Picture         =   "frmCards.frx":276E3
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   27
      Left            =   270
      Picture         =   "frmCards.frx":28F20
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   26
      Left            =   30
      Picture         =   "frmCards.frx":2A746
      Top             =   3990
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   25
      Left            =   2910
      Picture         =   "frmCards.frx":2BFAD
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   24
      Left            =   2670
      Picture         =   "frmCards.frx":2D93A
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   23
      Left            =   2430
      Picture         =   "frmCards.frx":2F2E6
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   22
      Left            =   2190
      Picture         =   "frmCards.frx":30BE9
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   21
      Left            =   1950
      Picture         =   "frmCards.frx":325BA
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   20
      Left            =   1710
      Picture         =   "frmCards.frx":33EDA
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   19
      Left            =   1470
      Picture         =   "frmCards.frx":35829
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   18
      Left            =   1230
      Picture         =   "frmCards.frx":370F9
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   17
      Left            =   990
      Picture         =   "frmCards.frx":38A32
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   16
      Left            =   750
      Picture         =   "frmCards.frx":3A35D
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   15
      Left            =   510
      Picture         =   "frmCards.frx":3BC33
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   14
      Left            =   270
      Picture         =   "frmCards.frx":3D561
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   13
      Left            =   30
      Picture         =   "frmCards.frx":3EE7A
      Top             =   2010
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   12
      Left            =   2910
      Picture         =   "frmCards.frx":407C2
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   11
      Left            =   2670
      Picture         =   "frmCards.frx":420D5
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   10
      Left            =   2430
      Picture         =   "frmCards.frx":43A20
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   9
      Left            =   2190
      Picture         =   "frmCards.frx":45277
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   8
      Left            =   1950
      Picture         =   "frmCards.frx":46BEE
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   7
      Left            =   1710
      Picture         =   "frmCards.frx":4847C
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   6
      Left            =   1470
      Picture         =   "frmCards.frx":49D52
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   5
      Left            =   1230
      Picture         =   "frmCards.frx":4B588
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   4
      Left            =   990
      Picture         =   "frmCards.frx":4CE33
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   3
      Left            =   750
      Picture         =   "frmCards.frx":4E6CF
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   2
      Left            =   510
      Picture         =   "frmCards.frx":4FF11
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   1
      Left            =   270
      Picture         =   "frmCards.frx":517B1
      Top             =   30
      Width           =   1425
   End
   Begin VB.Image imgCard 
      Height          =   1950
      Index           =   0
      Left            =   30
      Picture         =   "frmCards.frx":53029
      Top             =   30
      Width           =   1425
   End
End
Attribute VB_Name = "frmCards"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

