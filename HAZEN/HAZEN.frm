VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "HAZEN'S FORMULA"
   ClientHeight    =   5820
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9285
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5820
   ScaleWidth      =   9285
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtOutput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4845
      Left            =   5640
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "HAZEN.frx":0000
      Top             =   600
      Width           =   3375
   End
   Begin VB.TextBox txttrial 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   5
      Text            =   "15"
      Top             =   2040
      Width           =   1815
   End
   Begin VB.TextBox txtsg 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3360
      TabIndex        =   4
      Text            =   "2.65"
      Top             =   3480
      Width           =   1815
   End
   Begin VB.TextBox txtnu 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   3
      Text            =   "0.9"
      Top             =   1320
      Width           =   1815
   End
   Begin VB.TextBox txttemp 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   3360
      TabIndex        =   2
      Text            =   "25"
      Top             =   2760
      Width           =   1815
   End
   Begin VB.TextBox txtdia 
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   3360
      TabIndex        =   0
      Text            =   "0.14"
      Top             =   600
      Width           =   1815
   End
   Begin VB.Label lblTRIAL 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "NUMBER OF TRIAL"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   840
      TabIndex        =   12
      Top             =   2040
      Width           =   2340
   End
   Begin VB.Label lblSpGr 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SP. GRAVITY"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   11
      Top             =   3480
      Width           =   1770
   End
   Begin VB.Label lblDynVisc 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "dyn visc(mm/s2)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   10
      Top             =   1320
      Width           =   2160
   End
   Begin VB.Label lblTDegree 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "T(degree c)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   840
      TabIndex        =   9
      Top             =   2760
      Width           =   1545
   End
   Begin VB.Label lblDiameterMm 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "diameter(mm)"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   435
      Left            =   840
      TabIndex        =   8
      Top             =   600
      Width           =   2040
   End
   Begin VB.Label lblEnd 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "end"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   960
      TabIndex        =   7
      Top             =   4680
      Width           =   645
   End
   Begin VB.Label lblsolve 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "SOLVE"
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   18
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      Left            =   3720
      TabIndex        =   1
      Top             =   4680
      Width           =   1065
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim d As Double
Dim v As Double
Dim re As Double
Dim cd As Double
Dim iv As Double
Dim s As Double
Dim nu As Double
Dim t As Double
Dim check As Integer
Dim i As Integer
Dim trial As Integer



Private Sub Form_Load()
check = 1
End Sub

Private Sub lblEnd_Click()
End
End Sub

Private Sub lblsolve_Click()
d = Val(txtdia.Text)
trial = Val(txttrial.Text)
s = Val(txtsg.Text)
t = Val(txttemp.Text)
nu = Val(txtnu.Text)
txtOutput.Text = ""
If d > 0.1 And d <= 1 Then
If check = 1 Then
    Call initialize
End If
    For i = 1 To trial
        re = iv * d / nu
        cd = 24 / re + 3 / (re ^ 0.5) + 0.34
        v = 4 * 9.81 * 1000 * d * (s - 1) / (3 * cd)
        v = v ^ 0.5
         Call outputprint
        iv = v
       
    Next
    
Else
    
End If
End Sub

Public Sub initialize()
iv = 418 * (s - 1) * d * d * ((3 * t + 70) / 100)
check = 0
End Sub

Public Sub outputprint()
    txtOutput.Text = txtOutput.Text & "TRIAL # " & i & vbCrLf
    txtOutput.Text = txtOutput.Text & "-----------------" & vbCrLf
    txtOutput.Text = txtOutput.Text & "v =  " & Format(iv, "000.0000") & "  mm/s" & vbCrLf
    txtOutput.Text = txtOutput.Text & "Re = " & Format(re, "000.0000") & vbCrLf
    txtOutput.Text = txtOutput.Text & "CD = " & Format(cd, "000.0000") & vbCrLf
    txtOutput.Text = txtOutput.Text & "     ***     " & vbCrLf
End Sub

Private Sub txtdia_Change()
check = 1
End Sub

Private Sub txtnu_Change()
check = 1
End Sub

Private Sub txtsg_Change()
check = 1
End Sub

Private Sub txttemp_Change()
check = 1
End Sub

Private Sub txttrial_Change()
check = 1
End Sub
