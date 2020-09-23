VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmToolBox 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Tool Box"
   ClientHeight    =   4965
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3975
   Icon            =   "ToolBox.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4965
   ScaleWidth      =   3975
   StartUpPosition =   1  'CenterOwner
   Begin MSComctlLib.TabStrip TabStrip1 
      Height          =   735
      Left            =   0
      TabIndex        =   10
      Top             =   0
      Width           =   3975
      _ExtentX        =   7011
      _ExtentY        =   1296
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   8
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Temp"
            Object.Tag             =   "1"
            Object.ToolTipText     =   "Convert Farenheit, Celcius, Rankine, Kelvin"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Cylinder"
            Object.Tag             =   "2"
            Object.ToolTipText     =   "Calculate the Volume of a Vertical Cylinder"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Sphere"
            Object.Tag             =   "3"
            Object.ToolTipText     =   "Calculate the Volume of a Sphere"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Loan"
            Object.Tag             =   "4"
            Object.ToolTipText     =   "Amortize a loan"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ASCII"
            Object.Tag             =   "5"
            Object.ToolTipText     =   "Find the ASCII Value of any keyboard key"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ohms Law"
            Object.ToolTipText     =   "Find any of the missing factors for Ohms Law"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab7 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "System"
            Object.ToolTipText     =   "Physical RAM and Page File Information"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab8 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Timer"
            Object.ToolTipText     =   "Clock and  Event Timer with Pause Feature"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraTemp 
      Height          =   4095
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Width           =   3975
      Begin VB.TextBox txtFar 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   8
         ToolTipText     =   "Enter a known temperature here"
         Top             =   1680
         Width           =   1215
      End
      Begin VB.TextBox txtKel 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   7
         ToolTipText     =   "Enter a known temperature here"
         Top             =   1200
         Width           =   1215
      End
      Begin VB.TextBox txtRan 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   6
         ToolTipText     =   "Enter a known temperature here"
         Top             =   720
         Width           =   1215
      End
      Begin VB.TextBox txtCel 
         Alignment       =   1  'Right Justify
         Height          =   375
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Enter a known temperature here"
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label1 
         Caption         =   "Deg Farenheit"
         Height          =   255
         Index           =   3
         Left            =   1440
         TabIndex        =   4
         Top             =   1800
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Deg Rankine"
         Height          =   255
         Index           =   1
         Left            =   1440
         TabIndex        =   2
         Top             =   840
         Width           =   1455
      End
      Begin VB.Label Label1 
         Caption         =   "Deg Celsius"
         Height          =   255
         Index           =   0
         Left            =   1440
         TabIndex        =   1
         Top             =   360
         Width           =   1455
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter a known value in one of the cells.  Double click on any of the other blank cells to convert.  1Billion MAX."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   3135
      End
      Begin VB.Label Label1 
         Caption         =   "Deg Kelvin"
         Height          =   255
         Index           =   2
         Left            =   1440
         TabIndex        =   3
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label Label8 
         Height          =   255
         Left            =   2520
         TabIndex        =   24
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame fraLoan 
      Height          =   4095
      Left            =   0
      TabIndex        =   25
      Top             =   600
      Width           =   3975
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         Caption         =   "Frame2"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   38
         Top             =   2400
         Width           =   3735
         Begin VB.Label Label12 
            BackColor       =   &H80000005&
            Caption         =   "Pmt"
            Height          =   255
            Left            =   0
            TabIndex        =   42
            Top             =   0
            Width           =   375
         End
         Begin VB.Label Label13 
            BackColor       =   &H80000005&
            Caption         =   "Principal"
            Height          =   255
            Left            =   480
            TabIndex        =   41
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label14 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "Interest"
            Height          =   255
            Left            =   1200
            TabIndex        =   40
            Top             =   0
            Width           =   735
         End
         Begin VB.Label Label15 
            Alignment       =   1  'Right Justify
            BackColor       =   &H80000005&
            Caption         =   "Balance"
            Height          =   255
            Left            =   2280
            TabIndex        =   39
            Top             =   0
            Width           =   855
         End
      End
      Begin VB.TextBox txtTotInt 
         Alignment       =   2  'Center
         ForeColor       =   &H000000C0&
         Height          =   285
         Left            =   2400
         Locked          =   -1  'True
         TabIndex        =   36
         ToolTipText     =   "This is the total interest you will pay over the life of the loan"
         Top             =   2040
         Width           =   1455
      End
      Begin VB.TextBox txtPayment 
         Alignment       =   2  'Center
         ForeColor       =   &H00FF0000&
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   34
         ToolTipText     =   "This is the amount of your monthly payment"
         Top             =   2040
         Width           =   1575
      End
      Begin VB.ListBox lstAmortize 
         Appearance      =   0  'Flat
         Height          =   1005
         Left            =   120
         TabIndex        =   33
         ToolTipText     =   "Amortization schedule shows repayment progress each month"
         Top             =   2760
         Width           =   3735
      End
      Begin VB.CommandButton cmdUpdate 
         Caption         =   "Update"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         ToolTipText     =   "Calculate the payment, total interest, and build an amoritzation schedule"
         Top             =   1320
         Width           =   2055
      End
      Begin VB.ComboBox cmbInterest 
         Height          =   315
         ItemData        =   "ToolBox.frx":0442
         Left            =   120
         List            =   "ToolBox.frx":0444
         TabIndex        =   28
         Text            =   "Combo1"
         ToolTipText     =   "Choose an interest rate from the list"
         Top             =   960
         Width           =   1095
      End
      Begin VB.TextBox txtTerm 
         Height          =   285
         Left            =   240
         TabIndex        =   27
         ToolTipText     =   "Enter the number of monthly payments"
         Top             =   600
         Width           =   855
      End
      Begin VB.TextBox txtPV 
         Height          =   285
         Left            =   240
         TabIndex        =   26
         ToolTipText     =   "Enter the Original amount of the loan"
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label51 
         Alignment       =   1  'Right Justify
         Caption         =   "$"
         Height          =   255
         Left            =   0
         TabIndex        =   101
         Top             =   240
         Width           =   135
      End
      Begin VB.Label Label18 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter the Amount Borrowed, the number or months, and choose an interest rate from the drop down list."
         Height          =   1455
         Left            =   2400
         TabIndex        =   43
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label Label17 
         Alignment       =   2  'Center
         Caption         =   "Total Interest"
         Height          =   255
         Left            =   2520
         TabIndex        =   37
         Top             =   1800
         Width           =   1095
      End
      Begin VB.Label Label16 
         Caption         =   "Monthly Payment"
         Height          =   375
         Left            =   360
         TabIndex        =   35
         Top             =   1800
         Width           =   1335
      End
      Begin VB.Label Label11 
         Caption         =   "Interest Rate"
         Height          =   375
         Left            =   1320
         TabIndex        =   32
         Top             =   960
         Width           =   1215
      End
      Begin VB.Label Label10 
         Caption         =   "Num of Months"
         Height          =   255
         Left            =   1200
         TabIndex        =   31
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label Label9 
         Caption         =   "Amt Borrowed"
         Height          =   255
         Left            =   1200
         TabIndex        =   30
         Top             =   240
         Width           =   1215
      End
      Begin VB.Label Label19 
         Alignment       =   2  'Center
         Caption         =   "Amortization of a loan"
         Height          =   255
         Left            =   240
         TabIndex        =   44
         Top             =   3840
         Width           =   3495
      End
   End
   Begin VB.Frame fraSphere 
      Height          =   3975
      Left            =   0
      TabIndex        =   47
      Top             =   720
      Width           =   3975
      Begin VB.TextBox Text3 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   54
         ToolTipText     =   "Cu Ft Results Displayed here"
         Top             =   3600
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "Calculate"
         Height          =   495
         Left            =   240
         TabIndex        =   53
         ToolTipText     =   "Press to calculate after entering Diameter"
         Top             =   1320
         Width           =   3495
      End
      Begin VB.TextBox Text2 
         Height          =   375
         Left            =   240
         Locked          =   -1  'True
         TabIndex        =   51
         ToolTipText     =   "Cu In Results Displayed here"
         Top             =   2880
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Height          =   375
         Left            =   240
         TabIndex        =   49
         ToolTipText     =   "Enter the Diameter in inches here - 1500 MAX"
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label26 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "Find the Volume of a 3Dimentional Sphere.  Enter the Diameter of the Sphere at it's widest point.  Use inches."
         Height          =   855
         Left            =   240
         TabIndex        =   56
         Top             =   360
         Width           =   3375
      End
      Begin VB.Label Label25 
         Alignment       =   2  'Center
         Caption         =   "Volume, Cu Feet"
         Height          =   375
         Left            =   240
         TabIndex        =   55
         Top             =   3360
         Width           =   1335
      End
      Begin VB.Label Label24 
         Alignment       =   2  'Center
         Caption         =   "Volume, Cubic In."
         Height          =   255
         Left            =   120
         TabIndex        =   52
         Top             =   2640
         Width           =   1455
      End
      Begin VB.Label Label22 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         Caption         =   "SPHERE"
         Height          =   255
         Left            =   2160
         TabIndex        =   48
         Top             =   3720
         Width           =   1335
      End
      Begin VB.Image Image2 
         Height          =   1905
         Left            =   1800
         Picture         =   "ToolBox.frx":0446
         Top             =   2040
         Width           =   1995
      End
      Begin VB.Label Label23 
         Alignment       =   2  'Center
         Caption         =   "Diameter, Inches"
         Height          =   375
         Left            =   120
         TabIndex        =   50
         Top             =   1920
         Width           =   1455
      End
   End
   Begin VB.Frame fraAscii 
      Height          =   3975
      Left            =   0
      TabIndex        =   57
      Top             =   600
      Width           =   3975
      Begin VB.TextBox Text4 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   480
         TabIndex        =   58
         ToolTipText     =   "Press any key"
         Top             =   840
         Width           =   375
      End
      Begin VB.Label Label27 
         Caption         =   "Press a key on the keyboard"
         Height          =   255
         Left            =   1320
         TabIndex        =   63
         Top             =   840
         Width           =   2295
      End
      Begin VB.Label Label28 
         Alignment       =   2  'Center
         Caption         =   "ASCII Value Returned"
         Height          =   495
         Left            =   1320
         TabIndex        =   62
         Top             =   1080
         Width           =   2055
      End
      Begin VB.Label Label29 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   360
         TabIndex        =   61
         ToolTipText     =   "ASCII value shown here"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label31 
         Alignment       =   2  'Center
         Caption         =   "Find the ASCII Value of any key input"
         Height          =   255
         Left            =   120
         TabIndex        =   60
         Top             =   240
         Width           =   3375
      End
      Begin VB.Label Label32 
         Caption         =   "ASCII value of key pressed"
         Height          =   375
         Left            =   1200
         TabIndex        =   59
         Top             =   1560
         Width           =   2295
      End
      Begin VB.Line Line1 
         BorderWidth     =   4
         X1              =   240
         X2              =   3720
         Y1              =   2160
         Y2              =   2160
      End
   End
   Begin VB.Frame fraMem 
      Height          =   3975
      Left            =   0
      TabIndex        =   77
      Top             =   720
      Width           =   3975
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   333
         Left            =   240
         Top             =   3240
      End
      Begin VB.Label Label47 
         Alignment       =   2  'Center
         Caption         =   "GlobalMemoryStatus"
         Height          =   375
         Left            =   360
         TabIndex        =   89
         ToolTipText     =   "Windows API call"
         Top             =   3240
         Width           =   3255
      End
      Begin VB.Label Label49 
         Alignment       =   2  'Center
         Caption         =   "Percent Page File Free:"
         Height          =   255
         Left            =   360
         TabIndex        =   93
         Top             =   2640
         Width           =   2055
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   2520
         TabIndex        =   92
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label48 
         Alignment       =   1  'Right Justify
         Caption         =   "Percent Free Memory:"
         Height          =   375
         Left            =   360
         TabIndex        =   91
         Top             =   1560
         Width           =   1935
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   5
         Left            =   2520
         TabIndex        =   90
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Line Line2 
         BorderWidth     =   3
         X1              =   600
         X2              =   3360
         Y1              =   3120
         Y2              =   3120
      End
      Begin VB.Label Label46 
         Alignment       =   1  'Right Justify
         Caption         =   "Available Page File Bytes:"
         Height          =   375
         Left            =   360
         TabIndex        =   88
         Top             =   2280
         Width           =   1935
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   4
         Left            =   2520
         TabIndex        =   87
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   3
         Left            =   2520
         TabIndex        =   86
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   2
         Left            =   2520
         TabIndex        =   85
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   1
         Left            =   2520
         TabIndex        =   84
         Top             =   840
         Width           =   1095
      End
      Begin VB.Label Label45 
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   0
         Left            =   2520
         TabIndex        =   83
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label44 
         Height          =   495
         Left            =   240
         TabIndex        =   82
         Top             =   2880
         Width           =   1695
      End
      Begin VB.Label Label43 
         Alignment       =   1  'Right Justify
         Caption         =   "Page File Size:"
         Height          =   375
         Left            =   120
         TabIndex        =   81
         Top             =   1920
         Width           =   2175
      End
      Begin VB.Label Label42 
         Alignment       =   1  'Right Justify
         Caption         =   "Curren Memory Load:"
         Height          =   375
         Left            =   120
         TabIndex        =   80
         Top             =   1200
         Width           =   2175
      End
      Begin VB.Label Label41 
         Alignment       =   1  'Right Justify
         Caption         =   "Available Physical Memory:"
         Height          =   495
         Left            =   120
         TabIndex        =   79
         Top             =   840
         Width           =   2175
      End
      Begin VB.Label Label40 
         Alignment       =   1  'Right Justify
         Caption         =   "Total Physical Memory: "
         Height          =   375
         Left            =   120
         TabIndex        =   78
         Top             =   480
         Width           =   2175
      End
   End
   Begin VB.Frame fraCylinder 
      Height          =   3975
      Left            =   0
      TabIndex        =   11
      Top             =   600
      Width           =   3975
      Begin VB.CommandButton Command2 
         Caption         =   "Reset"
         Height          =   375
         Left            =   120
         TabIndex        =   23
         ToolTipText     =   "Clear old numbers"
         Top             =   960
         Width           =   1815
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Calculate"
         Height          =   375
         Left            =   120
         TabIndex        =   22
         ToolTipText     =   "Calculate you new entries"
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox txtToCuFt 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "0"
         ToolTipText     =   "Results show here"
         Top             =   2880
         Width           =   735
      End
      Begin VB.TextBox txtToCuIn 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "0"
         ToolTipText     =   "Results show here"
         Top             =   2520
         Width           =   735
      End
      Begin VB.TextBox txtCuInIn 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "0"
         ToolTipText     =   "Results show here"
         Top             =   2160
         Width           =   735
      End
      Begin VB.TextBox txtLen 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   13
         Text            =   "1"
         ToolTipText     =   "Enter the length here - In Inches"
         Top             =   1800
         Width           =   735
      End
      Begin VB.TextBox txtDia 
         Alignment       =   1  'Right Justify
         Height          =   285
         Left            =   120
         TabIndex        =   12
         ToolTipText     =   "Enter the diameter here - In Inches"
         Top             =   1440
         Width           =   735
      End
      Begin VB.Label Label21 
         Alignment       =   2  'Center
         Caption         =   "Find the Volume"
         Height          =   255
         Left            =   120
         TabIndex        =   46
         Top             =   3720
         Width           =   1815
      End
      Begin VB.Label Label20 
         Alignment       =   2  'Center
         BackColor       =   &H80000018&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Enter the diameter inches of the cylinder.  Enter a length or leave at 1 inch."
         Height          =   735
         Left            =   120
         TabIndex        =   45
         Top             =   240
         Width           =   2055
      End
      Begin VB.Image Image1 
         Height          =   3525
         Left            =   2160
         Picture         =   "ToolBox.frx":CAF8
         Top             =   480
         Width           =   1605
      End
      Begin VB.Label Label7 
         Caption         =   "Total Cu Ft"
         Height          =   255
         Left            =   960
         TabIndex        =   21
         Top             =   2880
         Width           =   975
      End
      Begin VB.Label Label6 
         Caption         =   "Total Cu In"
         Height          =   255
         Left            =   960
         TabIndex        =   19
         Top             =   2520
         Width           =   1095
      End
      Begin VB.Label Label5 
         Caption         =   "Cu In per Inch"
         Height          =   255
         Left            =   960
         TabIndex        =   18
         Top             =   2160
         Width           =   1095
      End
      Begin VB.Label Label3 
         Caption         =   "Diameter"
         Height          =   255
         Left            =   960
         TabIndex        =   14
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label4 
         Caption         =   "Lenth "
         Height          =   255
         Left            =   960
         TabIndex        =   15
         Top             =   1800
         Width           =   855
      End
   End
   Begin VB.Frame fraOhm 
      Height          =   3975
      Left            =   0
      TabIndex        =   64
      Top             =   720
      Width           =   3975
      Begin VB.CommandButton Command4 
         Caption         =   "Clear Entries"
         Height          =   375
         Left            =   240
         TabIndex        =   76
         ToolTipText     =   "Press Button to clear old numbers"
         Top             =   2760
         Width           =   1335
      End
      Begin VB.TextBox txtOhm 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   70
         ToolTipText     =   "Enter Ohms, or double click for result"
         Top             =   2280
         Width           =   1335
      End
      Begin VB.TextBox txtAmp 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   69
         ToolTipText     =   "Enter Ohms, or double click for result"
         Top             =   1800
         Width           =   1335
      End
      Begin VB.TextBox txtVolt 
         Alignment       =   2  'Center
         Height          =   375
         Left            =   240
         TabIndex        =   68
         ToolTipText     =   "Enter Volts, or double click for result"
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label39 
         Alignment       =   2  'Center
         Caption         =   "Enter any two of the known, double click in a cell to solve for the unknown"
         Height          =   495
         Left            =   360
         TabIndex        =   75
         Top             =   3120
         Width           =   2895
      End
      Begin VB.Label Label38 
         Alignment       =   2  'Center
         Caption         =   "Also, Remember Amps = Watts / Volts "
         Height          =   255
         Left            =   120
         TabIndex        =   74
         Top             =   960
         Width           =   3735
      End
      Begin VB.Label Label37 
         Caption         =   "Ohms"
         Height          =   375
         Left            =   1680
         TabIndex        =   73
         Top             =   2400
         Width           =   1335
      End
      Begin VB.Label Label35 
         Caption         =   "Volts"
         Height          =   255
         Left            =   1680
         TabIndex        =   71
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label34 
         Alignment       =   2  'Center
         Caption         =   "( Voltage = Amperes x Ohms )"
         Height          =   375
         Left            =   240
         TabIndex        =   67
         Top             =   720
         Width           =   3495
      End
      Begin VB.Label Label33 
         Alignment       =   2  'Center
         Caption         =   "Potential Difference = Current   x   Resistance"
         Height          =   375
         Left            =   120
         TabIndex        =   66
         Top             =   480
         Width           =   3735
      End
      Begin VB.Label Label30 
         Alignment       =   2  'Center
         Caption         =   "OHMS LAW"
         Height          =   255
         Left            =   120
         TabIndex        =   65
         Top             =   240
         Width           =   3735
      End
      Begin VB.Label Label36 
         Caption         =   "Amperes"
         Height          =   375
         Left            =   1680
         TabIndex        =   72
         Top             =   1920
         Width           =   1215
      End
   End
   Begin VB.Frame fraTimer 
      Height          =   3975
      Left            =   0
      TabIndex        =   94
      Top             =   720
      Width           =   3975
      Begin VB.CommandButton Command6 
         BackColor       =   &H80000014&
         Caption         =   "Pause Timer"
         Enabled         =   0   'False
         Height          =   375
         Left            =   720
         TabIndex        =   100
         Top             =   2640
         Width           =   2775
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   2160
         Top             =   3240
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H80000014&
         Caption         =   "Start Timer"
         Height          =   375
         Left            =   1080
         TabIndex        =   99
         Top             =   1320
         Width           =   1815
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   240
         Top             =   3360
      End
      Begin VB.Label Label52 
         Alignment       =   2  'Center
         Caption         =   "Elapsed Time"
         Height          =   255
         Left            =   1200
         TabIndex        =   98
         Top             =   1800
         Width           =   1575
      End
      Begin VB.Label lblTimer 
         Alignment       =   2  'Center
         Caption         =   "00:00:00"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   720
         TabIndex        =   97
         Top             =   2040
         Width           =   2655
      End
      Begin VB.Line Line3 
         X1              =   240
         X2              =   3840
         Y1              =   1200
         Y2              =   1200
      End
      Begin VB.Shape Shape1 
         Height          =   975
         Left            =   360
         Top             =   120
         Width           =   3255
      End
      Begin VB.Label Label50 
         Alignment       =   2  'Center
         Caption         =   "Current System Time:"
         Height          =   255
         Left            =   720
         TabIndex        =   96
         Top             =   240
         Width           =   2655
      End
      Begin VB.Label lblTime 
         Alignment       =   2  'Center
         Caption         =   "Label50"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   360
         TabIndex        =   95
         Top             =   480
         Width           =   3255
      End
   End
End
Attribute VB_Name = "frmToolBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim StartTime
Dim stoptime

Private Declare Function SendMessage Lib _
"user32" Alias "SendMessageA" (ByVal hwnd As Long, _
ByVal wMsg As Long, ByVal wParam As Long, _
lParam As Any) As Long

'---- Constant used by the above function for a listbox -----'
Private Const LB_SETTABSTOPS = &H192




Private Function ClrCylinder()
txtLen = "1"
txtDia = ""
txtCuInIn = "0"
txtToCuIn = "0"
txtToCuFt = "0"
End Function
Private Function ClrCel()
txtKel = ""
txtRan = ""
txtFar = ""
End Function
Private Function ClrRan()
txtKel = ""
txtCel = ""
txtFar = ""
End Function
Private Function ClrKel()
txtCel = ""
txtRan = ""
txtFar = ""
End Function
Private Function ClrFar()
txtKel = ""
txtRan = ""
txtCel = ""
End Function
Private Function ClrTxt()
txtCel = ""
txtKel = ""
txtRan = ""
txtFar = ""
End Function

Private Sub cmdUpdate_Click()
On Error GoTo StopHere
If txtPV = "" Or txtTerm = "" Then
MsgBox ("Get Real, I can't calculate empty boxes")
GoTo StopHere
End If

Dim Month As Integer
Dim MonthlyIntrest As Currency
Dim MonthlyPrincipal As Currency
Dim TotInt As Currency
Dim Amount As Currency
Dim Term As Double
Dim Rate As Double
Dim MonthlyPmt As Currency
Dim Balance As Currency
Call ClrList
Amount = Val(txtPV.Text)
Term = Val(txtTerm.Text)
Rate = Format(cmbInterest.Text, "General Number")
MonthlyPmt = -Pmt(Rate / 12, Term, Amount, 0, 0) 'Payment
txtPayment = Format(MonthlyPmt, "Currency")
Balance = Amount

SetLBTabStops lstAmortize, 20, 60, 105

For Month = 1 To Term Step 1

    MonthlyIntrest = Balance * Rate / 12
    MonthlyPrincipal = MonthlyPmt - MonthlyIntrest
    Balance = Balance - MonthlyPrincipal
    
    lstAmortize.AddItem Format(Month, "000") & vbTab & _
                         Format(MonthlyPrincipal, "Currency") & vbTab & _
                         Format(MonthlyIntrest, "Currency") & vbTab & _
                         Format(Balance, "Currency")
Next Month
TotInt = (MonthlyPmt * Term) - Amount
txtTotInt = Format(TotInt, "Currency")

StopHere:
End Sub

Private Sub Command1_Click()
If txtDia = "" Then
GoTo StopHere
End If
Length = txtLen
Dia = txtDia
Radias = Dia * 0.5
CuInIn = (Radias * Radias) * Pi
ToCuIn = CuInIn * Length
ToCuFt = ToCuIn / StdCuFt
txtCuInIn = CuInIn
txtToCuIn = ToCuIn
If ToCuFt >= 1 Then
    txtToCuFt = ToCuFt
    Else
    txtToCuFt = "0"
End If
StopHere:
End Sub

Private Sub Command2_Click()
Call ClrCylinder

End Sub

Private Sub Command3_Click()
On Error GoTo StopHere

Dim a As Long
Dim b As Long
Dim c As Single
Dim r As Long
Dim vol As Long
Dim vol2 As Double

c = 4 / 3
a = Val(Text1)
r = (a * 0.5)
b = (r ^ 3)
If a > 1500 Then
MsgBox ("Diameter too large, try 1500 or less !")
GoTo StopHere:
End If

vol = (c * Pi * b)
vol2 = vol / 1728
Text2 = Val(vol)
If vol > 1728 Then
Text3 = vol2
End If
StopHere:
End Sub

Private Sub Command4_Click()
txtVolt = ""
txtAmp = ""
txtOhm = ""
txtVolt.SetFocus
End Sub

Private Sub Command5_Click()
If Command5.Caption = "Start Timer" Then
Command5.Caption = "Stop Timer"
StartTime = Now
Timer3.Enabled = True
Command6.Enabled = True
Else
Command5.Caption = "Start Timer"
Timer3.Enabled = False
Command6.Enabled = False
End If


End Sub

Private Sub Command6_Click()
If Command6.Caption = "Pause Timer" Then
stoptime = Now - StartTime
Command5.Enabled = False

Command6.Caption = "Resume Timer"
Timer3.Enabled = False
Else
StartTime = Now - stoptime
Command6.Caption = "Pause Timer"
Timer3.Enabled = True
Command5.Enabled = True
End If

End Sub

Private Sub Command7_Click()
Dim a As Double
Dim b As Double
Dim c As Double
a = txtBottom
b = txtTop
Dim i As Double
Dim j As Double

For i = a To b
    c = Sqr(i)
    For j = 2 To c
    If j Mod i = 0 Then
    GoTo BailOut
    End If
    If j Mod i <> 0 And j >= c Then
    Text5 = j & "" & c & "" & j Mod i
    List1.AddItem (i)
    End If
    Next j
BailOut:
Next i



End Sub

Private Sub Form_Load()
If App.PrevInstance = True Then
MsgBox ("Program Already Running !!")
End
End If

lblTime.Caption = Format$(Now, "HH:MM:SS")
' set the Borderstyle of all frames to none
fraTimer.BorderStyle = 0
fraTemp.BorderStyle = 0
fraSphere.BorderStyle = 0
fraLoan.BorderStyle = 0
fraOhm.BorderStyle = 0
fraCylinder.BorderStyle = 0
fraAscii.BorderStyle = 0
fraMem.BorderStyle = 0

'Declarations for the Loan application
Dim X As Single
Dim i As Single

'---- For loop to populate the combobox with .5% to 25% ------'
For i = 0.005 To 0.3 Step 0.0001
    X = i
    cmbInterest.AddItem Format(X, "Percent")
Next i

'----- set the listbox to show 10% by default ------'
cmbInterest.ListIndex = 950

Call ClrTxt
Billion = 1000000000


End Sub
Public Function SetLBTabStops(LB As Object, _
ParamArray TabStops()) As Boolean

'----- Local Variables used in this function -----'
Dim alTabStops() As Long
Dim lCtr As Long
Dim lColumns As Long
Dim lRet As Long

'PURPOSE: Set TabSTops for a list box using the hwnd property.
'This creates columns separated by a tab character

'USAGE:
'Pass ListBox Object and a comma delimited
'list of tab stops.  Tab stops are expressed
'in dialog units which approximately equal
'1/4 the width of a character

On Error GoTo errorhandler:

ReDim alTabStops(UBound(TabStops)) As Long

For lCtr = 0 To UBound(TabStops)
    alTabStops(lCtr) = TabStops(lCtr)
Next

lColumns = UBound(alTabStops) + 1


lRet = SendMessage(LB.hwnd, LB_SETTABSTOPS, _
lColumns, alTabStops(0))

SetLBTabStops = (lRet = 0)
Exit Function

errorhandler:
    SetLBTabStops = False

End Function
Private Sub txtTemp_Change(Index As Integer)

End Sub

Private Sub txtTemp_DblClick(Index As Integer)


End Sub

Private Sub TabStrip1_Click()
' A little brute force is used in making the tab strip work

If TabStrip1.SelectedItem = "Timer" Then
Timer2.Enabled = True
fraTimer.Visible = True
Timer1.Enabled = False
fraMem.Visible = False
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = False
fraSphere.Visible = False
fraAscii.Visible = False
fraOhm.Visible = False
End If

If TabStrip1.SelectedItem = "System" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = True
fraMem.Visible = True
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = False
fraSphere.Visible = False
fraAscii.Visible = False
fraOhm.Visible = False
End If

If TabStrip1.SelectedItem = "Ohms Law" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = False
fraSphere.Visible = False
fraAscii.Visible = False
fraOhm.Visible = True
fraMem.Visible = False
txtVolt.SetFocus
End If

If TabStrip1.SelectedItem = "ASCII" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = False
fraSphere.Visible = False
fraAscii.Visible = True
fraOhm.Visible = False
fraMem.Visible = False
Text4.SetFocus

End If

If TabStrip1.SelectedItem = "Temp" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = True
fraSphere.Visible = False
fraAscii.Visible = False
fraOhm.Visible = False
fraMem.Visible = False
txtCel.SetFocus
End If

If TabStrip1.SelectedItem = "Sphere" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraAscii.Visible = False
fraSphere.Visible = True
fraCylinder.Visible = False
fraLoan.Visible = False
fraTemp.Visible = False
fraOhm.Visible = False
fraMem.Visible = False
Text1.SetFocus
End If

If TabStrip1.SelectedItem = "Cylinder" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraAscii.Visible = False
fraCylinder.Visible = True
fraLoan.Visible = False
fraTemp.Visible = False
fraSphere.Visible = False
fraOhm.Visible = False
fraMem.Visible = False
txtDia.SetFocus
End If
 
If TabStrip1.SelectedItem = "Loan" Then
Timer2.Enabled = False
fraTimer.Visible = False
Timer1.Enabled = False
fraCylinder.Visible = False
fraLoan.Visible = True
fraTemp.Visible = False
fraSphere.Visible = False
fraAscii.Visible = False
fraOhm.Visible = False
fraMem.Visible = False
txtPV.SetFocus
End If





End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
Label29.Caption = KeyAscii
'Label30.Caption = Hex
'Text5 = Text4

'Label31.Caption = dec(Label29.Caption)
End Sub

Private Sub Text4_KeyUp(KeyCode As Integer, Shift As Integer)
Text4 = ""
End Sub

Private Sub Timer1_Timer()
Dim MStat As MEMORYSTATUS
    MStat.dwLength = Len(MStat)
    GlobalMemoryStatus MStat
Dim AvailPg As Integer
Label45(0).Caption = MStat.dwTotalPhys
Label45(1).Caption = MStat.dwAvailPhys
Label45(2).Caption = MStat.dwMemoryLoad & "% Used"
Label45(3).Caption = MStat.dwTotalPageFile
Label45(4).Caption = MStat.dwAvailPageFile
Label45(5).Caption = (100 - MStat.dwMemoryLoad) & "% Free"
AvailPg = (MStat.dwAvailPageFile / MStat.dwTotalPageFile) * 100
Label45(6).Caption = AvailPg & "% Avail."
End Sub

Private Sub Timer2_Timer()
lblTime.Caption = Format$(Now, "hh:mm:ss")

End Sub

Private Sub Timer3_Timer()
lblTimer.Caption = Format$(Now - StartTime, "hh:mm:ss")

End Sub

Private Sub txtAmp_DblClick()
Dim volt As Single
Dim ohm As Single
Dim amp As Single
If txtAmp = "" And (txtOhm = "" Or txtVolt = "") Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere:
End If

If txtOhm = "" And txtVolt = "" Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere
End If

If txtVolt = "" Then
amp = txtAmp
ohm = txtOhm
volt = ohm * amp
txtVolt = volt
Else
If txtAmp = "" Then
ohm = txtOhm
volt = txtVolt
amp = volt / ohm
txtAmp = amp
Else
If txtOhm = "" Then
volt = txtVolt
amp = txtAmp
ohm = volt / amp
txtOhm = ohm
End If
End If
End If

StopHere:
End Sub

Private Sub txtAmp_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 46) Then
KeyAscii = 0
End If
End Sub

Private Sub txtCel_DblClick()
Dim X As Long
If txtCel <> "" Then
    X = txtCel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Cel = txtCel
Call ConvCel
txtRan = Ran
txtKel = Kel
txtFar = Far
Else
If txtRan <> "" Then
    X = txtRan.Text
    If X > Billion Then
    GoTo StopHere
    End If
Ran = txtRan
Call ConvRan
txtCel = Cel
txtKel = Kel
txtFar = Far
Else
If txtKel <> "" Then
X = txtKel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Kel = txtKel
Call ConvKel
txtCel = Cel
txtRan = Ran
txtFar = Far
Else
If txtFar <> "" Then
X = txtFar.Text
    If X > Billion Then
    GoTo StopHere
    End If
Far = txtFar
Call ConvFar
txtCel = Cel
txtRan = Ran
txtKel = Kel
'End If

End If
End If
End If
End If
StopHere:
End Sub

Private Sub txtCel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 45) Then
KeyAscii = 0
GoTo StopHere
End If
Call ClrCel
StopHere:
End Sub

Private Sub txtDia_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 46) Then
KeyAscii = 0
GoTo StopHere
End If

StopHere:
End Sub

Private Sub txtFar_DblClick()
On Error GoTo StopHere

Dim X As Long
If txtCel <> "" Then
    X = txtCel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Cel = txtCel
Call ConvCel
txtRan = Ran
txtKel = Kel
txtFar = Far
Else
If txtRan <> "" Then
    X = txtRan.Text
    If X > Billion Then
    GoTo StopHere
    End If
Ran = txtRan
Call ConvRan
txtCel = Cel
txtKel = Kel
txtFar = Far
Else
If txtKel <> "" Then
X = txtKel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Kel = txtKel
Call ConvKel
txtCel = Cel
txtRan = Ran
txtFar = Far
Else
If txtFar <> "" Then
X = txtFar.Text
    If X > Billion Then
    GoTo StopHere
    End If
Far = txtFar
Call ConvFar
txtCel = Cel
txtRan = Ran
txtKel = Kel
'End If

End If
End If
End If
End If
StopHere:
End Sub

Private Sub txtFar_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 45) Then
KeyAscii = 0
GoTo StopHere
End If
Call ClrFar
StopHere:
End Sub

Private Sub txtKel_DblClick()
On Error GoTo StopHere


Dim X As Long
If txtCel <> "" Then
    X = txtCel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Cel = txtCel
Call ConvCel
txtRan = Ran
txtKel = Kel
txtFar = Far
Else
If txtRan <> "" Then
    X = txtRan.Text
    If X > Billion Then
    GoTo StopHere
    End If
Ran = txtRan
Call ConvRan
txtCel = Cel
txtKel = Kel
txtFar = Far
Else
If txtKel <> "" Then
X = txtKel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Kel = txtKel
Call ConvKel
txtCel = Cel
txtRan = Ran
txtFar = Far
Else
If txtFar <> "" Then
X = txtFar.Text
    If X > Billion Then
    GoTo StopHere
    End If
Far = txtFar
Call ConvFar
txtCel = Cel
txtRan = Ran
txtKel = Kel
'End If

End If
End If
End If
End If
StopHere:
End Sub

Private Sub txtKel_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 45) Then
KeyAscii = 0
GoTo StopHere
End If
Call ClrKel
StopHere:
End Sub

Private Sub txtLen_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 46) Then
KeyAscii = 0
GoTo StopHere
End If

StopHere:
End Sub

Private Sub txtOhm_DblClick()
Dim volt As Single
Dim ohm As Single
Dim amp As Single
If txtAmp = "" And (txtOhm = "" Or txtVolt = "") Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere:
End If

If txtOhm = "" And txtVolt = "" Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere
End If

If txtVolt = "" Then
amp = txtAmp
ohm = txtOhm
volt = ohm * amp
txtVolt = volt
Else
If txtAmp = "" Then
ohm = txtOhm
volt = txtVolt
amp = volt / ohm
txtAmp = amp
Else
If txtOhm = "" Then
volt = txtVolt
amp = txtAmp
ohm = volt / amp
txtOhm = ohm
End If
End If
End If

StopHere:
End Sub

Private Sub txtOhm_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 46) Then
KeyAscii = 0
End If
End Sub

Private Sub txtPV_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
KeyAscii = 0
GoTo StopHere
End If
StopHere:
End Sub

Private Sub txtRan_DblClick()
On Error GoTo StopHere
Dim X As Long
If txtCel <> "" Then
    X = txtCel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Cel = txtCel
Call ConvCel
txtRan = Ran
txtKel = Kel
txtFar = Far
Else
If txtRan <> "" Then
    X = txtRan.Text
    If X > Billion Then
    GoTo StopHere
    End If
Ran = txtRan
Call ConvRan
txtCel = Cel
txtKel = Kel
txtFar = Far
Else
If txtKel <> "" Then
X = txtKel.Text
    If X > Billion Then
    GoTo StopHere
    End If
Kel = txtKel
Call ConvKel
txtCel = Cel
txtRan = Ran
txtFar = Far
Else
If txtFar <> "" Then
X = txtFar.Text
    If X > Billion Then
    GoTo StopHere
    End If
Far = txtFar
Call ConvFar
txtCel = Cel
txtRan = Ran
txtKel = Kel
'End If

End If
End If
End If
End If
StopHere:
End Sub

Private Sub txtRan_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8 And KeyAscii <> 45) Then
KeyAscii = 0
GoTo StopHere
End If
Call ClrRan
StopHere:

End Sub

Private Sub txtTerm_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
KeyAscii = 0
GoTo StopHere
End If
StopHere:
End Sub

Private Sub txtVolt_DblClick()
Dim volt As Single
Dim ohm As Single
Dim amp As Single
If txtAmp = "" And (txtOhm = "" Or txtVolt = "") Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere:
End If

If txtOhm = "" And txtVolt = "" Then
MsgBox ("Mission Control, We hav a problem - need two data fields to continue !")
GoTo StopHere
End If

If txtVolt = "" Then
amp = txtAmp
ohm = txtOhm
volt = ohm * amp
txtVolt = volt
Else
If txtAmp = "" Then
ohm = txtOhm
volt = txtVolt
amp = volt / ohm
txtAmp = amp
Else
If txtOhm = "" Then
volt = txtVolt
amp = txtAmp
ohm = volt / amp
txtOhm = ohm
End If
End If
End If

StopHere:
End Sub

Private Sub txtVolt_KeyPress(KeyAscii As Integer)
If (KeyAscii < 48 Or KeyAscii > 57) And (KeyAscii <> 8) Then
KeyAscii = 0
End If
End Sub
