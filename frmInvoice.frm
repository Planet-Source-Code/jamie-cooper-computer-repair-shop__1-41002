VERSION 5.00
Begin VB.Form frmInvoice 
   Caption         =   "Invoice"
   ClientHeight    =   5730
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6480
   LinkTopic       =   "Form1"
   ScaleHeight     =   5730
   ScaleWidth      =   6480
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
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
      Height          =   735
      Left            =   3600
      TabIndex        =   5
      Top             =   960
      Width           =   2775
   End
   Begin VB.CommandButton cmdTotal 
      Caption         =   "Total"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1680
      TabIndex        =   4
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton cmdAddItem 
      Caption         =   "Add Item to Invoice"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   2775
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "Print Invoice"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3120
      TabIndex        =   2
      Top             =   5040
      Width           =   1695
   End
   Begin VB.ComboBox cboCode 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      ItemData        =   "frmInvoice.frx":0000
      Left            =   120
      List            =   "frmInvoice.frx":002E
      TabIndex        =   1
      Text            =   "Job Code"
      Top             =   120
      Width           =   3375
   End
   Begin VB.ListBox lstInvoice 
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3195
      Left            =   0
      TabIndex        =   0
      Top             =   1800
      Width           =   6495
   End
End
Attribute VB_Name = "frmInvoice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'This is the form that displays the invoice
'dimming the variables
Dim code, custom(6), c, n As String
Dim price, ctr As Single

Private Sub cmdAddItem_Click()
    'Routine that adds an item from a combo box
    code = UCase(Left(cboCode.Text, 2))
    Select Case code
        Case "HD"
            price = price + 39.99
            lstInvoice.AddItem ("Hard Drive Instal: $39.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Hard Drive Instal: $39.99"
        Case "MB"
            price = price + 69.99
            lstInvoice.AddItem ("MotherBoard Instal: $69.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "MotherBoard Instal: $69.99"
        Case "PU"
            price = price + 59.99
            lstInvoice.AddItem ("Processor Upgrade: $59.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Processor Upgrade: $59.99"
        Case "MY"
            price = price + 19.99
            lstInvoice.AddItem ("Memory Instal: $19.99")
           Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Memory Instal: $19.99"
        Case "SC"
            price = price + 39.99
            lstInvoice.AddItem ("Sound Card Instal: $39.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Sound Card Instal: $39.99"
        Case "VC"
            price = price + 39.99
            lstInvoice.AddItem ("Video Card Instal: $39.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Video Card Instal: $39.99"
        Case "MM"
            price = price + 29.99
            lstInvoice.AddItem ("Modem Instal: $29.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Modem Instal: $29.99"
        Case "NC"
            price = price + 29.99
            lstInvoice.AddItem ("Network Card Instal: $29.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Network Card Instal: $29.99"
        Case "PS"
            price = price + 49.99
            lstInvoice.AddItem ("Power Supply Instal: $49.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Power Supply Instal: $49.99"
        Case "FD"
            price = price + 19.99
            lstInvoice.AddItem ("Floppy Drive Instal: $19.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Floppy Drive Instal: $19.99"
        Case "ZD"
            price = price + 19.99
            lstInvoice.AddItem ("Zip Drive Instal: $19.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "Zip Drive Instal: $19.99"
        Case "CD"
            price = price + 39.99
            lstInvoice.AddItem ("CD, or CDR Drive: $39.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "CD or CDR Drive: $39.99"
        Case "DD"
            price = price + 39.99
            lstInvoice.AddItem ("DVD or DVDRW Drive: $39.99")
            Printer.FontBold = True
            Printer.FontSize = 14
            Printer.Print Spc(10); "DVD or DVDRW Drive: $39.99"
    End Select
End Sub

Private Sub cmdClear_Click()
    'Clear the boxes
    lstInvoice.Clear
    price = 0
    Call Form_Load
End Sub

Private Sub cmdPrint_Click()
    'print to the printer
    Printer.EndDoc
End Sub

Private Sub cmdTotal_Click()
    'routine that when clicked displays the total and sends the total to the printer
    lstInvoice.AddItem ("Total: " & FormatCurrency(price))
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print Spc(10); "Total: " & "$" & price
End Sub

Private Sub Form_Load()
    'sub called when the from is loaded
    c = InputBox("Please enter the name of the customer: ")
    Open "A:\Customers.txt" For Input As #1
    Do Until EOF(1)
        Input #1, n
        If c = n Then
            custom(0) = c
            For ctr = 1 To 5
                Input #1, n
                custom(ctr) = n
            Next ctr
        End If
    Loop
    Close #1
    
    'Displaying the header in list box
    lstInvoice.AddItem (Space(25) & custom(0) & " Repair Invoice")
    lstInvoice.AddItem (Space(20) & custom(1))
    lstInvoice.AddItem (Space(20) & custom(2) & " " & custom(3) & ", " & custom(4))
    lstInvoice.AddItem (Space(20) & custom(6))
    lstInvoice.AddItem ("")
    lstInvoice.AddItem ("")
    
    'Displaying the header on the printed invoice
    Printer.FontBold = True
    Printer.FontSize = 14
    Printer.Print Spc(20), custom(0) & " Repair Invoice"
    Printer.Print Spc(15), custom(1)
    Printer.Print Spc(15), custom(2) & " " & custom(3) & ", " & custom(4)
    Printer.Print Spc(15), custom(6)
End Sub
