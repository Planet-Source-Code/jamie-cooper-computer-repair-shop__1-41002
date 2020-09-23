VERSION 5.00
Begin VB.Form frmJDC 
   Caption         =   "Jamie's Computer Repair Shop"
   ClientHeight    =   5235
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5925
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   5925
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   1200
      TabIndex        =   6
      Top             =   1080
      Width           =   2415
   End
   Begin VB.ListBox lstCustomers 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2910
      Left            =   0
      TabIndex        =   15
      Top             =   2280
      Width           =   5895
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   4320
      MaxLength       =   5
      TabIndex        =   5
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   3
      Left            =   840
      TabIndex        =   4
      Top             =   720
      Width           =   1575
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   2
      Left            =   5040
      MaxLength       =   2
      TabIndex        =   3
      Top             =   360
      Width           =   855
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   1920
      TabIndex        =   2
      Top             =   360
      Width           =   2295
   End
   Begin VB.TextBox txtCustomer 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   2415
   End
   Begin VB.CommandButton CmdProcess 
      Caption         =   "Create Invoice"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   3
      Left            =   4320
      TabIndex        =   10
      Top             =   1560
      Width           =   1695
   End
   Begin VB.CommandButton CmdProcess 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   2
      Left            =   3000
      TabIndex        =   9
      Top             =   1560
      Width           =   1335
   End
   Begin VB.CommandButton CmdProcess 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   1
      Left            =   1560
      TabIndex        =   8
      Top             =   1560
      Width           =   1455
   End
   Begin VB.CommandButton CmdProcess 
      Caption         =   "Add Data"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Index           =   0
      Left            =   0
      TabIndex        =   7
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label lblCity 
      Caption         =   "City:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   16
      Top             =   720
      Width           =   735
   End
   Begin VB.Label lblPhoneN 
      Caption         =   "Phone #:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   14
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label lblState 
      Caption         =   "State:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4200
      TabIndex        =   13
      Top             =   360
      Width           =   735
   End
   Begin VB.Label lblZip 
      Caption         =   "Zip Code:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3000
      TabIndex        =   12
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label lblAddress 
      Caption         =   "Street Address:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   360
      Width           =   2055
   End
   Begin VB.Label lblName 
      Caption         =   "Customer Name:"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2055
   End
End
Attribute VB_Name = "frmJDC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Form where you can add people to the file and create an invoice
Dim i As Integer
Dim customer(6), custo As String

Private Sub CmdProcess_Click(Index As Integer)
    'sub routine that looks at the command buttons that are on form

    'loop that puts the names from the text boxes into an array known as customer
    For i = 0 To 5
        customer(i) = txtCustomer(i).Text
    Next i

    'select case that decides which button was clicked and what to do
    Select Case Index
        Case 0
            Open "A:\Customers.txt" For Append As #1
            For i = 0 To 5
                Write #1, customer(i)
            Next i
            Close #1
            lstCustomers.AddItem (customer(0))
            lstCustomers.AddItem (customer(1))
            lstCustomers.AddItem (customer(3) & " " & customer(2) & ", " & customer(4))
            lstCustomers.AddItem (customer(5))
        Case 1
            For i = 0 To 5
                txtCustomer(i).Text = ""
            Next i
        Case 2
            Unload Me
            End
        Case 3
            frmInvoice.Show
    End Select
End Sub
