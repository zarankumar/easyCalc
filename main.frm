VERSION 5.00
Begin VB.Form frmmain 
   BackColor       =   &H00C0FFC0&
   Caption         =   "EasyCalc Ver: 0.1"
   ClientHeight    =   10830
   ClientLeft      =   2685
   ClientTop       =   600
   ClientWidth     =   18165
   LinkTopic       =   "Form1"
   ScaleHeight     =   10830
   ScaleWidth      =   18165
   WindowState     =   2  'Maximized
   Begin VB.ComboBox cmbCalc 
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7560
      TabIndex        =   12
      Text            =   "Area"
      Top             =   2040
      Width           =   2895
   End
   Begin VB.ComboBox cmbShape 
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   12
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   420
      Left            =   7440
      TabIndex        =   11
      Text            =   "Circle"
      Top             =   2760
      Width           =   2895
   End
   Begin VB.TextBox txtradius 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   10
      Top             =   3600
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.TextBox txtlength 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   4440
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.TextBox txtwidth 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   8
      Top             =   5885
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtheight 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   7
      Top             =   6600
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtangle 
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   11.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7440
      TabIndex        =   6
      Top             =   7320
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtBreadth 
      Height          =   375
      Left            =   7440
      TabIndex        =   5
      Top             =   5161
      Visible         =   0   'False
      Width           =   2655
   End
   Begin VB.TextBox txtresult 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   27.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1575
      Left            =   4920
      TabIndex        =   2
      Top             =   8880
      Width           =   4215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Myriad Pro"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   7680
      TabIndex        =   1
      Top             =   8040
      Width           =   1935
   End
   Begin VB.Label labelunit 
      BackStyle       =   0  'Transparent
      Caption         =   "Unit"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   20
      Top             =   9480
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape1 
      Height          =   7215
      Left            =   3960
      Top             =   1440
      Width           =   9375
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Calculate"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   19
      Top             =   2040
      Width           =   2175
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Geometry"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   18
      Top             =   2725
      Width           =   2055
   End
   Begin VB.Label labelRadius 
      BackStyle       =   0  'Transparent
      Caption         =   "Radius"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   17
      Top             =   3410
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labelLength 
      BackStyle       =   0  'Transparent
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   16
      Top             =   4335
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label labelWidth 
      BackStyle       =   0  'Transparent
      Caption         =   "Width"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   15
      Top             =   5585
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Label LabelHeight 
      BackStyle       =   0  'Transparent
      Caption         =   "Height"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5040
      TabIndex        =   14
      Top             =   6270
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.Label labelAngle 
      BackStyle       =   0  'Transparent
      Caption         =   "Angle"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5040
      TabIndex        =   13
      Top             =   7200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label labelResult 
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   15.75
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   1800
      TabIndex        =   4
      Top             =   9000
      Width           =   2895
   End
   Begin VB.Label labelBreadth 
      BackStyle       =   0  'Transparent
      Caption         =   "Breadth"
      BeginProperty Font 
         Name            =   "Myriad Pro Light"
         Size            =   14.25
         Charset         =   0
         Weight          =   600
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5040
      TabIndex        =   3
      Top             =   5020
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Easy Calc"
      BeginProperty Font 
         Name            =   "Myriad Pro Cond"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Index           =   0
      Left            =   6600
      TabIndex        =   0
      Top             =   0
      Width           =   5295
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub desable()

labelRadius.Visible = False
txtradius.Visible = False
labelLength.Visible = False
txtlength.Visible = False
txtwidth.Visible = False
labelWidth.Visible = False
labelBreadth.Visible = False
txtBreadth.Visible = False
txtangle.Visible = False
labelAngle.Visible = False
LabelHeight.Visible = False
txtheight.Visible = False

End Sub






Private Sub cmbCalc_Click()
cmbShape.Clear

If cmbCalc.Text = "Area" Then

        'adding values to cmbShape
        cmbShape.Text = "Circle"
        cmbShape.AddItem "Circle"
        cmbShape.AddItem "Square"
        cmbShape.AddItem "Rectangle"
        cmbShape.AddItem "Parallelogram"
        cmbShape.AddItem "Triangle"
        cmbShape.AddItem "Ellipse"
        cmbShape.AddItem "Trapezoind"
End If
If cmbCalc.Text = "Volume" Then
            cmbShape.Text = "Cube"
          cmbShape.AddItem "Cube"
          cmbShape.AddItem "Rectangle Prism"
          cmbShape.AddItem "Irregular Prism"
          cmbShape.AddItem "Cylinder"
          cmbShape.AddItem "Pyramid"
          cmbShape.AddItem "Cone"
          cmbShape.AddItem "Sphere"
          cmbShape.AddItem "Ellipsoid"
          
End If



End Sub

Private Sub cmbShape_Click()
Call desable
'...................AREA CALCULATIONS....................

 If cmbCalc.Text = "Area" Then
            If cmbShape.Text = "Circle" Then
                Call desable
                 labelRadius.Visible = True
                 txtradius.Visible = True
                 txtradius.Text = ""
            ElseIf cmbShape.Text = "Triangle" Then
                Call desable
                labelBreadth.Visible = True
                LabelHeight.Visible = True
                txtBreadth.Visible = True
                txtheight.Visible = True
                txtBreadth.Text = ""
                txtheight.Text = ""
          ElseIf cmbShape.Text = "Rectangle" Then
                Call desable
                labelLength.Visible = True
                labelWidth.Visible = True
                txtlength.Visible = True
                txtwidth.Visible = True
                txtlength.Text = ""
                txtwidth.Text = ""
                
                
          ElseIf cmbShape.Text = "Square" Then
                Call desable
               labelLength.Visible = True
                txtlength.Visible = True
                txtlength.Text = ""
        ElseIf cmbShape.Text = "Parallelogram" Then
                Call desable
                labelBreadth.Visible = True
                LabelHeight.Visible = True
                txtBreadth.Visible = True
                txtheight.Visible = True
                txtBreadth.Text = ""
                 txtheight.Text = ""
         ElseIf cmbShape.Text = "Ellipse" Then
                Call desable
                labelRadius.Visible = True
                txtradius.Visible = True
                labelLength.Visible = True
                labelLength.Caption = "Radius-2"
                txtlength.Visible = True
                txtradius.Text = ""
                txtlength.Text = ""
          ElseIf cmbShape.Text = "Trapezoind" Then
                Call desable
                LabelHeight.Visible = True
                labelWidth.Visible = True
                labelBreadth.Visible = True
                txtheight.Visible = True
                txtBreadth.Visible = True
                txtwidth.Visible = True
                txtheight.Text = ""
                txtBreadth.Text = ""
                txtwidth.Text = ""
                
                
        End If
        
        
                
     
                

 End If
 
  ' if user clicked volume
If cmbCalc.Text = "Volume" Then
            If cmbShape.Text = "Cube" Then
            Call desable
            labelLength.Visible = True
            txtlength.Visible = True
            txtlength.Text = ""
            ElseIf cmbShape.Text = "Rectangle Prism" Then
            Call desable
            labelLength.Visible = True
            labelWidth.Visible = True
            LabelHeight.Visible = True
            txtlength.Visible = True
            txtwidth.Visible = True
            txtheight.Visible = True
            txtlength.Text = ""
            txtwidth.Text = ""
            txtheight.Text = ""
            ElseIf cmbShape.Text = "Irregular Prism" Then
            Call desable
            labelBreadth.Visible = True
            LabelHeight.Visible = True
            txtBreadth.Visible = True
            txtheight.Visible = True
            txtBreadth.Text = ""
            txtheight.Text = ""
            ElseIf cmbShape.Text = "Cylinder" Then
            Call desable
            labelRadius.Visible = True
            LabelHeight.Visible = True
            txtradius.Visible = True
            txtheight.Visible = True
            txtradius.Text = ""
            txtheight.Text = ""
            ElseIf cmbShape.Text = "Pyramid" Then
            Call desable
            labelBreadth.Visible = True
            LabelHeight.Visible = True
            txtBreadth.Visible = True
            txtheight.Visible = True
            txtBreadth.Text = ""
            txtheight.Text = ""
            ElseIf cmbShape.Text = "Cone" Then
            Call desable
            labelRadius.Visible = True
            LabelHeight.Visible = True
            txtradius.Visible = True
            txtheight.Visible = True
            txtradius.Text = ""
            txtheight.Text = ""
            ElseIf cmbShape.Text = "Sphere" Then
            Call desable
            labelRadius.Visible = True
                   txtradius.Visible = True
                        txtradius.Text = ""
            ElseIf cmbShape.Text = "Ellipsoid" Then
            Call desable
            labelLength.Visible = True
            labelWidth.Visible = True
            LabelHeight.Visible = True
            txtlength.Visible = True
            txtwidth.Visible = True
            txtheight.Visible = True
            txtlength.Text = ""
            txtwidth.Text = ""
            txtheight.Text = ""
            
            End If
            

End If




End Sub

Private Sub Command1_Click()
labelResult.Caption = cmbCalc.Text
labelunit.Visible = True

txtresult.Text = ""



'finding Area
If cmbCalc.Text = "Area" Then
    If cmbShape.Text = "Circle" Then
      txtresult.Text = 3.14 * Val(txtradius.Text) * Val(txtradius.Text)
    ElseIf cmbShape.Text = "Triangle" Then
      txtresult.Text = (Val(txtheight.Text) * Val(txtBreadth.Text)) / 2
   ElseIf cmbShape.Text = "Rectangle" Then
        txtresult.Text = Val(txtlength.Text) * Val(txtwidth.Text)
  ElseIf cmbShape.Text = "Square" Then
          txtresult.Text = Val(txtlength.Text) * Val(txtlength.Text)
    ElseIf cmbShape.Text = "Parallelogram" Then
    txtresult.Text = Val(txtBreadth.Text) * Val(txtheight.Text)
    ElseIf cmbShape.Text = "Ellipse" Then
    txtresult.Text = Val(txtradius.Text) * Val(txtlength.Text) * 3.14
   ElseIf cmbShape.Text = "Trapezoind" Then
   txtresult.Text = (Val(txtwidth.Text) + Val(txtBreadth.Text)) / (Val(txtheight.Text) / 2)
   End If
   
    
        
      



End If
' finding Volume
If cmbCalc.Text = "Volume" Then
      If cmbShape.Text = "Cube" Then
      txtresult.Text = Val(txtlength.Text) * Val(txtlength.Text) * Val(txtlength.Text)
      ElseIf cmbShape.Text = "Rectangle Prism" Then
      txtresult.Text = Val(txtlength.Text) * Val(txtwidth.Text) * Val(txtheight.Text)
      ElseIf cmbShape.Text = "Irregular Prism" Then
      txtresult.Text = Val(txtBreadth.Text) * Val(txtheight.Text)
      ElseIf cmbShape.Text = "Cylinder" Then
      txtresult.Text = 3.14 * Val(txtradius.Text) * Val(txtradius.Text) * Val(txtheight.Text)
      ElseIf cmbShape.Text = "Pyramid" Then
      txtresult.Text = 0.33333 * Val(txtBreadth.Text) * Val(txtheight.Text)
       ElseIf cmbShape.Text = "Cone" Then
       txtresult.Text = 0.3333 * 3.14 * Val(txtradius.Text) * Val(txtradius.Text) * Val(txtheight.Text)
             ElseIf cmbShape.Text = "Sphere" Then
             txtresult.Text = 1.333 * Val(txtradius.Text) * Val(txtradius.Text) * Val(txtradius.Text)
      ElseIf cmbShape.Text = "Ellipsoid" Then
      txtresult.Text = 1.333 * 3.14 * Val(txtlength.Text) * Val(txtwidth.Text) * Val(txtheight.Text)
      End If
      
End If





End Sub

Private Sub Form_Load()
'adding values to cmbcalc
labelunit.Visible = False

cmbCalc.AddItem "Area"
cmbCalc.AddItem "Volume"







End Sub
