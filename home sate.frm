VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8595
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12870
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   12870
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text11 
      Alignment       =   2  'Center
      BeginProperty Font 
         Name            =   "Segoe Script"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   600
      TabIndex        =   31
      Top             =   7560
      Width           =   10815
   End
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   9720
      Top             =   240
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   6120
      TabIndex        =   30
      Top             =   6720
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   6120
      TabIndex        =   28
      Top             =   6120
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6120
      TabIndex        =   26
      Top             =   5520
      Width           =   1695
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   6120
      TabIndex        =   25
      Top             =   4800
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6120
      TabIndex        =   24
      Top             =   4080
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6120
      TabIndex        =   23
      Top             =   3360
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   2040
      TabIndex        =   17
      Top             =   6960
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   2040
      TabIndex        =   16
      Top             =   6360
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   2040
      TabIndex        =   15
      Top             =   5760
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2040
      TabIndex        =   14
      Top             =   5160
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kembalian"
      Height          =   615
      Left            =   9720
      TabIndex        =   8
      Top             =   3360
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proses"
      Height          =   615
      Left            =   9720
      TabIndex        =   7
      Top             =   2400
      Width           =   2175
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Keluar"
      Height          =   615
      Left            =   9720
      TabIndex        =   6
      Top             =   4320
      Width           =   2175
   End
   Begin VB.ComboBox Combo3 
      Height          =   315
      Left            =   7680
      TabIndex        =   5
      Text            =   "Harga Nasi"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   5760
      TabIndex        =   4
      Text            =   "Harga Minuman"
      Top             =   1440
      Width           =   1575
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   4080
      TabIndex        =   3
      Text            =   "Harga Sate"
      Top             =   1440
      Width           =   1335
   End
   Begin VB.ListBox List2 
      Height          =   2790
      Left            =   2280
      TabIndex        =   2
      Top             =   1320
      Width           =   1455
   End
   Begin VB.ListBox List1 
      Height          =   2790
      Left            =   240
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Label Label15 
      Caption         =   "Tanggal"
      Height          =   495
      Left            =   10560
      TabIndex        =   33
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label14 
      Caption         =   "Jam"
      Height          =   375
      Left            =   10560
      TabIndex        =   32
      Top             =   240
      Width           =   1695
   End
   Begin VB.Label Label13 
      Caption         =   "Kembalian"
      Height          =   255
      Left            =   4200
      TabIndex        =   29
      Top             =   6840
      Width           =   1695
   End
   Begin VB.Label Label12 
      Caption         =   "Uang yang dibayarkan"
      Height          =   255
      Left            =   4200
      TabIndex        =   27
      Top             =   6240
      Width           =   1695
   End
   Begin VB.Label Label11 
      Caption         =   "Total Bayar"
      Height          =   255
      Left            =   4440
      TabIndex        =   22
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label10 
      Caption         =   "Nasi"
      Height          =   255
      Left            =   4440
      TabIndex        =   21
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label9 
      Caption         =   "Minuman"
      Height          =   255
      Left            =   4440
      TabIndex        =   20
      Top             =   4200
      Width           =   1335
   End
   Begin VB.Label Label8 
      Caption         =   "Total Harga Sate"
      Height          =   375
      Left            =   4440
      TabIndex        =   19
      Top             =   3360
      Width           =   1335
   End
   Begin VB.Label Label7 
      Alignment       =   2  'Center
      Caption         =   "Pembayaran"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4800
      TabIndex        =   18
      Top             =   2520
      Width           =   2415
   End
   Begin VB.Label Label6 
      Caption         =   "Harga nasi"
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   6960
      Width           =   1455
   End
   Begin VB.Label Label5 
      Caption         =   "Harga Minuman"
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   " Harga Sate"
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   5760
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah Pesanan"
      Height          =   255
      Left            =   240
      TabIndex        =   10
      Top             =   5160
      Width           =   1455
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Pesanan"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   9
      Top             =   4320
      Width           =   1935
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Home Sate"
      BeginProperty Font 
         Name            =   "Segoe Print"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
End

End Sub

Private Sub Command2_Click()
Text2 = Combo1.Text
Text3 = Combo2.Text
Text4 = Combo3.Text
Text5 = Text1 * Val(Text2)
Text6 = Text3
Text7 = Text4
Text8 = Val(Text5) + Val(Text6) + Val(Text7)

End Sub

Private Sub Command3_Click()
Text10 = Text9 - Text8
Text11 = "Terima Kasih Telah Mengunjungi Home Sate"
End Sub

Private Sub form_load()
List1.AddItem "Sate Ayam"
List1.AddItem "Sate Sapi"
List1.AddItem "Sate Domba"
List1.AddItem "Sate Kambing"
List1.AddItem "Sate Kelenci"
List1.AddItem "Sate Kuda"

List2.AddItem "Jus Alpukat"
List2.AddItem "Jus Melon"
List2.AddItem "Jus Jeruk"
List2.AddItem "Jus Sirsak"
List2.AddItem "Jus Mangga"
List2.AddItem "Es Teh Manis"
List2.AddItem "Es Teh Tawar"
List2.AddItem "Teh Anget"

Combo1.AddItem 2500
Combo1.AddItem 4000
Combo1.AddItem 4500
Combo1.AddItem 4750
Combo1.AddItem 5000
Combo1.AddItem 5500

Combo2.AddItem 10000
Combo2.AddItem 10000
Combo2.AddItem 10000
Combo2.AddItem 10000
Combo2.AddItem 10000
Combo2.AddItem 3000
Combo2.AddItem 3000
Combo2.AddItem 3000
Combo2.AddItem 2000
Combo2.AddItem 0

Combo3.AddItem 6000
Combo3.AddItem 5000
Combo3.AddItem 0
End Sub
Private Sub Timer1_Timer()
Label14.Caption = Format(Time, "hh:mm:ss AM/PM")
Label15.Caption = Format(Date, "dd/mm/yyyy")
End Sub
