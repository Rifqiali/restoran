VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   11355
   LinkTopic       =   "Form1"
   ScaleHeight     =   8415
   ScaleWidth      =   11355
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   6600
      TabIndex        =   36
      Text            =   " "
      Top             =   5760
      Width           =   1695
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   8880
      TabIndex        =   34
      Top             =   1200
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Height          =   495
      Left            =   6600
      TabIndex        =   33
      Top             =   5040
      Width           =   1695
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   6600
      TabIndex        =   32
      Top             =   4440
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   6600
      TabIndex        =   29
      Top             =   3840
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Kembalian"
      Height          =   615
      Left            =   8760
      TabIndex        =   27
      Top             =   5760
      Width           =   2295
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Proses"
      Height          =   615
      Left            =   8760
      TabIndex        =   26
      Top             =   4920
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   6600
      TabIndex        =   25
      Top             =   3240
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Exit"
      Height          =   615
      Left            =   8760
      TabIndex        =   23
      Top             =   6600
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   6600
      TabIndex        =   22
      Top             =   2520
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   6600
      TabIndex        =   20
      Top             =   1920
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   6600
      TabIndex        =   19
      Top             =   1320
      Width           =   1695
   End
   Begin VB.Frame Frame2 
      Caption         =   "Minuman"
      Height          =   3615
      Left            =   2880
      TabIndex        =   11
      Top             =   1080
      Width           =   1695
      Begin VB.OptionButton Option15 
         Caption         =   "Jus Alpukat"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   2640
         Width           =   1455
      End
      Begin VB.OptionButton Option14 
         Caption         =   "Es Campur"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2160
         Width           =   1335
      End
      Begin VB.OptionButton Option13 
         Caption         =   "Jus Naga"
         Height          =   375
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   1095
      End
      Begin VB.OptionButton Option12 
         Caption         =   "Jus Melon"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   1080
         Width           =   1095
      End
      Begin VB.OptionButton Option11 
         Caption         =   "Jus Apel"
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Makanan"
      Height          =   6135
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   2535
      Begin VB.OptionButton Option10 
         Caption         =   "Sapi Lada HItam"
         Height          =   375
         Left            =   120
         TabIndex        =   10
         Top             =   5640
         Width           =   1575
      End
      Begin VB.OptionButton Option9 
         Caption         =   "Tomyum"
         Height          =   615
         Left            =   120
         TabIndex        =   9
         Top             =   4920
         Width           =   1215
      End
      Begin VB.OptionButton Option8 
         Caption         =   "Fuyunghai"
         Height          =   495
         Left            =   120
         TabIndex        =   8
         Top             =   4320
         Width           =   1695
      End
      Begin VB.OptionButton Option7 
         Caption         =   "Bihun Goreng Seafood"
         Height          =   495
         Left            =   120
         TabIndex        =   7
         Top             =   3720
         Width           =   1935
      End
      Begin VB.OptionButton Option6 
         Caption         =   "Kwetiaw Goreng Seafood"
         Height          =   495
         Left            =   120
         TabIndex        =   6
         Top             =   3120
         Width           =   2175
      End
      Begin VB.OptionButton Option5 
         Caption         =   "Udang Krispi"
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   2520
         Width           =   1335
      End
      Begin VB.OptionButton Option4 
         Caption         =   "Mie Goreng Seafood"
         Height          =   615
         Left            =   120
         TabIndex        =   4
         Top             =   1920
         Width           =   1815
      End
      Begin VB.OptionButton Option3 
         Caption         =   "Kakap Asam Manis"
         Height          =   495
         Left            =   120
         TabIndex        =   3
         Top             =   1440
         Width           =   2055
      End
      Begin VB.OptionButton Option2 
         Caption         =   "Cumi Asam Manis"
         Height          =   375
         Left            =   120
         TabIndex        =   2
         Top             =   960
         Width           =   1815
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Nasi Goreng Seafood"
         Height          =   495
         Left            =   120
         TabIndex        =   1
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Label Label8 
      Caption         =   "Jumlah Pesanan"
      Height          =   255
      Left            =   4800
      TabIndex        =   35
      Top             =   5760
      Width           =   1575
   End
   Begin VB.Label Label7 
      Caption         =   "Kembalian"
      Height          =   375
      Left            =   4800
      TabIndex        =   31
      Top             =   5160
      Width           =   1575
   End
   Begin VB.Label Label6 
      Caption         =   "Total Bayar"
      Height          =   255
      Left            =   4800
      TabIndex        =   30
      Top             =   4560
      Width           =   1575
   End
   Begin VB.Label Label5 
      Caption         =   "Jumlah keseluruhan"
      Height          =   375
      Left            =   4920
      TabIndex        =   28
      Top             =   3840
      Width           =   1455
   End
   Begin VB.Label Label4 
      Caption         =   "Diskon"
      Height          =   255
      Left            =   4920
      TabIndex        =   24
      Top             =   3360
      Width           =   1455
   End
   Begin VB.Label Label3 
      Caption         =   "Jumlah"
      Height          =   255
      Left            =   4920
      TabIndex        =   21
      Top             =   2760
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Harga Minuman"
      Height          =   255
      Left            =   4920
      TabIndex        =   18
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Harga Makanan"
      Height          =   375
      Left            =   4920
      TabIndex        =   17
      Top             =   1320
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()
Text3.Text = Val(Text1) + Val(Text2)
Text4.Text = Combo1.Text
Text5.Text = Text3 - Text4 / 10 * 10000

End Sub

Private Sub Command3_Click()
Text7.Text = Text6 - Text5
End Sub

Private Sub form_load()
Combo1.AddItem "10"
Combo1.AddItem "0"
End Sub

Private Sub Command1_Click()
End

End Sub



Private Sub Option1_Click()
Text1.Text = 30000
End Sub

Private Sub Option10_Click()
Text1.Text = 80000
End Sub

Private Sub Option11_Click()
Text2.Text = 10000
End Sub

Private Sub Option12_Click()
Text2.Text = 10000
End Sub

Private Sub Option13_Click()
Text2.Text = 15000
End Sub

Private Sub Option14_Click()
Text2.Text = 25000
End Sub

Private Sub Option15_Click()
Text2.Text = 10000
End Sub

Private Sub Option2_Click()
Text1.Text = 40000
End Sub

Private Sub Option3_Click()
Text1.Text = 90000
End Sub

Private Sub Option4_Click()
Text1.Text = 30000
End Sub

Private Sub Option5_Click()
Text1.Text = 60000
End Sub

Private Sub Option6_Click()
Text1.Text = 30000
End Sub

Private Sub Option7_Click()
Text1.Text = 25000
End Sub

Private Sub Option8_Click()
Text1.Text = 70000
End Sub

Private Sub Option9_Click()
Text1.Text = 80000
End Sub

