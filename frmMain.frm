VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Main Menu"
   ClientHeight    =   1485
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   ScaleHeight     =   1485
   ScaleWidth      =   4080
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   4095
   End
   Begin VB.CommandButton cmdInvoice 
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
      Height          =   495
      Left            =   0
      TabIndex        =   1
      Top             =   480
      Width           =   4095
   End
   Begin VB.CommandButton cmdAddNew 
      Caption         =   "Add New Customer"
      BeginProperty Font 
         Name            =   "Bookman Old Style"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Programmer: Jamie Cooper
'Date: 11/25/02
'Program: Final VB6 project
'A program that can print invoices for a computer repair shop

Private Sub cmdAddNew_Click()
    'sub to call the invoice form
    frmJDC.Show
End Sub

Private Sub cmdExit_Click()
    'End the program
    End
End Sub

Private Sub cmdInvoice_Click()
    'Call up the invoice form
    frmInvoice.Show
End Sub
