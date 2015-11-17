VERSION 5.00
Begin VB.Form frm1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   9285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12495
   DrawStyle       =   6  'Inside Solid
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "Courier New"
      Size            =   12
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H8000000E&
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   PaletteMode     =   2  'Custom
   ScaleHeight     =   9285
   ScaleWidth      =   12495
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   9135
      Left            =   120
      ScaleHeight     =   9135
      ScaleWidth      =   12015
      TabIndex        =   0
      Top             =   120
      Width           =   12015
      Begin VB.PictureBox picS 
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         Height          =   8205
         Left            =   3480
         ScaleHeight     =   8205
         ScaleWidth      =   8055
         TabIndex        =   2
         Top             =   390
         Width           =   8055
         Begin VB.PictureBox Picture3 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   7050
            Left            =   -30
            ScaleHeight     =   7050
            ScaleWidth      =   3405
            TabIndex        =   236
            Top             =   -120
            Width           =   3405
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sikap / Sifat Anda"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Index           =   1
               Left            =   300
               TabIndex        =   257
               Top             =   165
               Width           =   2700
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Sikap Positif"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   18
               Left            =   975
               TabIndex        =   256
               Top             =   1755
               Width           =   2055
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Hangat"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   17
               Left            =   2055
               TabIndex        =   255
               Top             =   2055
               Width           =   975
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Cerewet"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   16
               Left            =   1935
               TabIndex        =   254
               Top             =   2340
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Bersemangat"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   15
               Left            =   1335
               TabIndex        =   253
               Top             =   2655
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Jarang Cemas"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   14
               Left            =   1215
               TabIndex        =   252
               Top             =   2940
               Width           =   1815
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Berbelaskasihan"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   13
               Left            =   735
               TabIndex        =   251
               Top             =   3270
               Width           =   2295
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Dermawan"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   12
               Left            =   1815
               TabIndex        =   250
               Top             =   3570
               Width           =   1215
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tidak Disiplin"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   11
               Left            =   855
               TabIndex        =   249
               Top             =   3870
               Width           =   2175
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mudah Terpengaruh"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   10
               Left            =   375
               TabIndex        =   248
               Top             =   4170
               Width           =   2655
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Gelisah"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   9
               Left            =   1935
               TabIndex        =   247
               Top             =   4485
               Width           =   1095
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tidak Teratur"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   8
               Left            =   975
               TabIndex        =   246
               Top             =   4770
               Width           =   2055
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tak Bertanggungjawab"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   7
               Left            =   15
               TabIndex        =   245
               Top             =   5085
               Width           =   3015
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Terus Terang"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   6
               Left            =   1215
               TabIndex        =   244
               Top             =   5400
               Width           =   1815
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mudah Patuh"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   20
               Left            =   1335
               TabIndex        =   243
               Top             =   1185
               Width           =   1695
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Tulus"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   19
               Left            =   2175
               TabIndex        =   242
               Top             =   1485
               Width           =   855
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Periang, Ramah"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   0
               Left            =   855
               TabIndex        =   241
               Top             =   855
               Width           =   2175
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Mengajukan Diri"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   5
               Left            =   735
               TabIndex        =   240
               Top             =   5685
               Width           =   2295
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Membesar-besarkan"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   4
               Left            =   375
               TabIndex        =   239
               Top             =   5985
               Width           =   2655
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Penakut, Kurang Aman"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   3
               Left            =   15
               TabIndex        =   238
               Top             =   6285
               Width           =   3015
            End
            Begin VB.Label Label2 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "Modern"
               ForeColor       =   &H0080FF80&
               Height          =   375
               Index           =   2
               Left            =   2055
               TabIndex        =   237
               Top             =   6615
               Width           =   975
            End
         End
         Begin VB.PictureBox Picture2 
            BackColor       =   &H00000000&
            BorderStyle     =   0  'None
            Height          =   6765
            Left            =   3375
            ScaleHeight     =   6765
            ScaleWidth      =   4065
            TabIndex        =   3
            Top             =   -30
            Width           =   4065
            Begin VB.PictureBox Picture1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   270
               Left            =   75
               ScaleHeight     =   270
               ScaleWidth      =   4515
               TabIndex        =   224
               Top             =   450
               Width           =   4515
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "10"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   10
                  Left            =   3495
                  TabIndex        =   234
                  Top             =   0
                  Width           =   300
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "1"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   11
                  Left            =   0
                  TabIndex        =   233
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "2"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   12
                  Left            =   400
                  TabIndex        =   232
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "3"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   13
                  Left            =   800
                  TabIndex        =   231
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "4"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   14
                  Left            =   1200
                  TabIndex        =   230
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "5"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   15
                  Left            =   1600
                  TabIndex        =   229
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "6"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   16
                  Left            =   2000
                  TabIndex        =   228
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "7"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   17
                  Left            =   2400
                  TabIndex        =   227
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "8"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   18
                  Left            =   2800
                  TabIndex        =   226
                  Top             =   0
                  Width           =   150
               End
               Begin VB.Label Label1 
                  AutoSize        =   -1  'True
                  BackStyle       =   0  'Transparent
                  Caption         =   "9"
                  ForeColor       =   &H0000FFFF&
                  Height          =   270
                  Index           =   19
                  Left            =   3200
                  TabIndex        =   225
                  Top             =   0
                  Width           =   150
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   315
               Index           =   19
               Left            =   0
               ScaleHeight     =   315
               ScaleWidth      =   3930
               TabIndex        =   213
               Top             =   6465
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   190
                  Left            =   0
                  TabIndex        =   223
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   191
                  Left            =   400
                  TabIndex        =   222
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   192
                  Left            =   800
                  TabIndex        =   221
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   193
                  Left            =   1200
                  TabIndex        =   220
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   194
                  Left            =   1600
                  TabIndex        =   219
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   195
                  Left            =   2000
                  TabIndex        =   218
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   196
                  Left            =   2400
                  TabIndex        =   217
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   197
                  Left            =   2800
                  TabIndex        =   216
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   198
                  Left            =   3200
                  TabIndex        =   215
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   199
                  Left            =   3600
                  TabIndex        =   214
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   18
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   202
               Top             =   6165
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   180
                  Left            =   0
                  TabIndex        =   212
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   181
                  Left            =   400
                  TabIndex        =   211
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   182
                  Left            =   800
                  TabIndex        =   210
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   183
                  Left            =   1200
                  TabIndex        =   209
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   184
                  Left            =   1600
                  TabIndex        =   208
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   185
                  Left            =   2000
                  TabIndex        =   207
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   186
                  Left            =   2400
                  TabIndex        =   206
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   187
                  Left            =   2800
                  TabIndex        =   205
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   188
                  Left            =   3200
                  TabIndex        =   204
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   189
                  Left            =   3600
                  TabIndex        =   203
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   17
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   191
               Top             =   5865
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   170
                  Left            =   0
                  TabIndex        =   201
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   171
                  Left            =   400
                  TabIndex        =   200
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   172
                  Left            =   800
                  TabIndex        =   199
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   173
                  Left            =   1200
                  TabIndex        =   198
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   174
                  Left            =   1600
                  TabIndex        =   197
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   175
                  Left            =   2000
                  TabIndex        =   196
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   176
                  Left            =   2400
                  TabIndex        =   195
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   177
                  Left            =   2800
                  TabIndex        =   194
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   178
                  Left            =   3200
                  TabIndex        =   193
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   179
                  Left            =   3600
                  TabIndex        =   192
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   16
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   180
               Top             =   5565
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   160
                  Left            =   0
                  TabIndex        =   190
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   161
                  Left            =   400
                  TabIndex        =   189
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   162
                  Left            =   800
                  TabIndex        =   188
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   163
                  Left            =   1200
                  TabIndex        =   187
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   164
                  Left            =   1600
                  TabIndex        =   186
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   165
                  Left            =   2000
                  TabIndex        =   185
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   166
                  Left            =   2400
                  TabIndex        =   184
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   167
                  Left            =   2800
                  TabIndex        =   183
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   168
                  Left            =   3200
                  TabIndex        =   182
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   169
                  Left            =   3600
                  TabIndex        =   181
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   15
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   169
               Top             =   5265
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   150
                  Left            =   0
                  TabIndex        =   179
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   151
                  Left            =   400
                  TabIndex        =   178
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   152
                  Left            =   800
                  TabIndex        =   177
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   153
                  Left            =   1200
                  TabIndex        =   176
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   154
                  Left            =   1600
                  TabIndex        =   175
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   155
                  Left            =   2000
                  TabIndex        =   174
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   156
                  Left            =   2400
                  TabIndex        =   173
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   157
                  Left            =   2800
                  TabIndex        =   172
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   158
                  Left            =   3200
                  TabIndex        =   171
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   159
                  Left            =   3600
                  TabIndex        =   170
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   14
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   158
               Top             =   4965
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   140
                  Left            =   0
                  TabIndex        =   168
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   141
                  Left            =   400
                  TabIndex        =   167
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   142
                  Left            =   800
                  TabIndex        =   166
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   143
                  Left            =   1200
                  TabIndex        =   165
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   144
                  Left            =   1600
                  TabIndex        =   164
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   145
                  Left            =   2000
                  TabIndex        =   163
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   146
                  Left            =   2400
                  TabIndex        =   162
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   147
                  Left            =   2800
                  TabIndex        =   161
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   148
                  Left            =   3200
                  TabIndex        =   160
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   149
                  Left            =   3600
                  TabIndex        =   159
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   13
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   147
               Top             =   4665
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   130
                  Left            =   0
                  TabIndex        =   157
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   131
                  Left            =   400
                  TabIndex        =   156
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   132
                  Left            =   800
                  TabIndex        =   155
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   133
                  Left            =   1200
                  TabIndex        =   154
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   134
                  Left            =   1600
                  TabIndex        =   153
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   135
                  Left            =   2000
                  TabIndex        =   152
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   136
                  Left            =   2400
                  TabIndex        =   151
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   137
                  Left            =   2800
                  TabIndex        =   150
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   138
                  Left            =   3200
                  TabIndex        =   149
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   139
                  Left            =   3600
                  TabIndex        =   148
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   12
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   136
               Top             =   4365
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   120
                  Left            =   0
                  TabIndex        =   146
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   121
                  Left            =   400
                  TabIndex        =   145
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   122
                  Left            =   800
                  TabIndex        =   144
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   123
                  Left            =   1200
                  TabIndex        =   143
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   124
                  Left            =   1600
                  TabIndex        =   142
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   125
                  Left            =   2000
                  TabIndex        =   141
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   126
                  Left            =   2400
                  TabIndex        =   140
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   127
                  Left            =   2800
                  TabIndex        =   139
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   128
                  Left            =   3200
                  TabIndex        =   138
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   129
                  Left            =   3600
                  TabIndex        =   137
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   11
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   125
               Top             =   4065
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   110
                  Left            =   0
                  TabIndex        =   135
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   111
                  Left            =   400
                  TabIndex        =   134
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   112
                  Left            =   800
                  TabIndex        =   133
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   113
                  Left            =   1200
                  TabIndex        =   132
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   114
                  Left            =   1600
                  TabIndex        =   131
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   115
                  Left            =   2000
                  TabIndex        =   130
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   116
                  Left            =   2400
                  TabIndex        =   129
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   117
                  Left            =   2800
                  TabIndex        =   128
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   118
                  Left            =   3200
                  TabIndex        =   127
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   119
                  Left            =   3600
                  TabIndex        =   126
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   10
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   114
               Top             =   3765
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   100
                  Left            =   0
                  TabIndex        =   124
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   101
                  Left            =   400
                  TabIndex        =   123
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   102
                  Left            =   800
                  TabIndex        =   122
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   103
                  Left            =   1200
                  TabIndex        =   121
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   104
                  Left            =   1600
                  TabIndex        =   120
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   105
                  Left            =   2000
                  TabIndex        =   119
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   106
                  Left            =   2400
                  TabIndex        =   118
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   107
                  Left            =   2800
                  TabIndex        =   117
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   108
                  Left            =   3200
                  TabIndex        =   116
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   109
                  Left            =   3600
                  TabIndex        =   115
                  Top             =   0
                  Width           =   255
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   9
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   103
               Top             =   3465
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   90
                  Left            =   0
                  TabIndex        =   113
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   91
                  Left            =   400
                  TabIndex        =   112
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   92
                  Left            =   800
                  TabIndex        =   111
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   93
                  Left            =   1200
                  TabIndex        =   110
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   94
                  Left            =   1600
                  TabIndex        =   109
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   95
                  Left            =   2000
                  TabIndex        =   108
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   96
                  Left            =   2400
                  TabIndex        =   107
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   97
                  Left            =   2800
                  TabIndex        =   106
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   98
                  Left            =   3200
                  TabIndex        =   105
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   99
                  Left            =   3600
                  TabIndex        =   104
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   8
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   92
               Top             =   3165
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   80
                  Left            =   0
                  TabIndex        =   102
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   81
                  Left            =   400
                  TabIndex        =   101
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   82
                  Left            =   800
                  TabIndex        =   100
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   83
                  Left            =   1200
                  TabIndex        =   99
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   84
                  Left            =   1600
                  TabIndex        =   98
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   85
                  Left            =   2000
                  TabIndex        =   97
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   86
                  Left            =   2400
                  TabIndex        =   96
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   87
                  Left            =   2800
                  TabIndex        =   95
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   88
                  Left            =   3200
                  TabIndex        =   94
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   89
                  Left            =   3600
                  TabIndex        =   93
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   7
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   81
               Top             =   2865
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   70
                  Left            =   0
                  TabIndex        =   91
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   71
                  Left            =   400
                  TabIndex        =   90
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   72
                  Left            =   800
                  TabIndex        =   89
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   73
                  Left            =   1200
                  TabIndex        =   88
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   74
                  Left            =   1600
                  TabIndex        =   87
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   75
                  Left            =   2000
                  TabIndex        =   86
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   76
                  Left            =   2400
                  TabIndex        =   85
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   77
                  Left            =   2800
                  TabIndex        =   84
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   78
                  Left            =   3200
                  TabIndex        =   83
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   79
                  Left            =   3600
                  TabIndex        =   82
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   6
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   70
               Top             =   2565
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   60
                  Left            =   0
                  TabIndex        =   80
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   61
                  Left            =   400
                  TabIndex        =   79
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   62
                  Left            =   800
                  TabIndex        =   78
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   63
                  Left            =   1200
                  TabIndex        =   77
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   64
                  Left            =   1600
                  TabIndex        =   76
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   65
                  Left            =   2000
                  TabIndex        =   75
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   66
                  Left            =   2400
                  TabIndex        =   74
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   67
                  Left            =   2800
                  TabIndex        =   73
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   68
                  Left            =   3200
                  TabIndex        =   72
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   69
                  Left            =   3600
                  TabIndex        =   71
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   5
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   59
               Top             =   2265
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   50
                  Left            =   0
                  TabIndex        =   69
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   51
                  Left            =   400
                  TabIndex        =   68
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   52
                  Left            =   800
                  TabIndex        =   67
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   53
                  Left            =   1200
                  TabIndex        =   66
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   54
                  Left            =   1600
                  TabIndex        =   65
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   55
                  Left            =   2000
                  TabIndex        =   64
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   56
                  Left            =   2400
                  TabIndex        =   63
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   57
                  Left            =   2800
                  TabIndex        =   62
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   58
                  Left            =   3200
                  TabIndex        =   61
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   59
                  Left            =   3600
                  TabIndex        =   60
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   4
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   48
               Top             =   1965
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   40
                  Left            =   0
                  TabIndex        =   58
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   41
                  Left            =   400
                  TabIndex        =   57
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   42
                  Left            =   800
                  TabIndex        =   56
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   43
                  Left            =   1200
                  TabIndex        =   55
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   44
                  Left            =   1600
                  TabIndex        =   54
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   45
                  Left            =   2000
                  TabIndex        =   53
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   46
                  Left            =   2400
                  TabIndex        =   52
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   47
                  Left            =   2800
                  TabIndex        =   51
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   48
                  Left            =   3200
                  TabIndex        =   50
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   49
                  Left            =   3600
                  TabIndex        =   49
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   3
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   37
               Top             =   1665
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   30
                  Left            =   0
                  TabIndex        =   47
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   31
                  Left            =   400
                  TabIndex        =   46
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   32
                  Left            =   800
                  TabIndex        =   45
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   33
                  Left            =   1200
                  TabIndex        =   44
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   34
                  Left            =   1600
                  TabIndex        =   43
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   35
                  Left            =   2000
                  TabIndex        =   42
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   36
                  Left            =   2400
                  TabIndex        =   41
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   37
                  Left            =   2800
                  TabIndex        =   40
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   38
                  Left            =   3200
                  TabIndex        =   39
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   39
                  Left            =   3600
                  TabIndex        =   38
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   2
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   26
               Top             =   1365
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   20
                  Left            =   0
                  TabIndex        =   36
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   21
                  Left            =   400
                  TabIndex        =   35
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   22
                  Left            =   800
                  TabIndex        =   34
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   23
                  Left            =   1200
                  TabIndex        =   33
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   24
                  Left            =   1600
                  TabIndex        =   32
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   25
                  Left            =   2000
                  TabIndex        =   31
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   26
                  Left            =   2400
                  TabIndex        =   30
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   27
                  Left            =   2800
                  TabIndex        =   29
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   28
                  Left            =   3200
                  TabIndex        =   28
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   29
                  Left            =   3600
                  TabIndex        =   27
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   0
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   15
               Top             =   765
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   9
                  Left            =   3600
                  TabIndex        =   25
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   8
                  Left            =   3200
                  TabIndex        =   24
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   7
                  Left            =   2800
                  TabIndex        =   23
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   6
                  Left            =   2400
                  TabIndex        =   22
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   5
                  Left            =   2000
                  TabIndex        =   21
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   4
                  Left            =   1600
                  TabIndex        =   20
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   3
                  Left            =   1200
                  TabIndex        =   19
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   2
                  Left            =   800
                  TabIndex        =   18
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   1
                  Left            =   400
                  TabIndex        =   17
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   0
                  Left            =   15
                  TabIndex        =   16
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.PictureBox pic1 
               BackColor       =   &H00000000&
               BorderStyle     =   0  'None
               Height          =   300
               Index           =   1
               Left            =   0
               ScaleHeight     =   300
               ScaleWidth      =   3930
               TabIndex        =   4
               Top             =   1065
               Width           =   3930
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   10
                  Left            =   0
                  TabIndex        =   14
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   11
                  Left            =   400
                  TabIndex        =   13
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   12
                  Left            =   800
                  TabIndex        =   12
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   13
                  Left            =   1200
                  TabIndex        =   11
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   14
                  Left            =   1600
                  TabIndex        =   10
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   15
                  Left            =   2000
                  TabIndex        =   9
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   16
                  Left            =   2400
                  TabIndex        =   8
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   17
                  Left            =   2800
                  TabIndex        =   7
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   18
                  Left            =   3200
                  TabIndex        =   6
                  Top             =   0
                  Width           =   270
               End
               Begin VB.OptionButton optB 
                  BackColor       =   &H00000000&
                  Caption         =   "Option1"
                  Height          =   270
                  Index           =   19
                  Left            =   3600
                  TabIndex        =   5
                  Top             =   0
                  Width           =   270
               End
            End
            Begin VB.Label Label2 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "NILAI"
               ForeColor       =   &H00FFFFFF&
               Height          =   270
               Index           =   21
               Left            =   1560
               TabIndex        =   235
               Top             =   45
               Width           =   750
            End
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "<ESC>  -------> KELUAR"
            ForeColor       =   &H0080FFFF&
            Height          =   270
            Index           =   0
            Left            =   3360
            TabIndex        =   260
            Top             =   7800
            Width           =   3300
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "==========================="
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Index           =   3
            Left            =   3360
            TabIndex        =   259
            Top             =   7560
            Width           =   4065
         End
         Begin VB.Label Label3 
            BackStyle       =   0  'Transparent
            Caption         =   "==========================="
            ForeColor       =   &H0080FFFF&
            Height          =   375
            Index           =   2
            Left            =   3360
            TabIndex        =   258
            Top             =   8040
            Width           =   4065
         End
      End
      Begin VB.CommandButton cmdProsen 
         Caption         =   "&Prosentasi"
         Height          =   435
         Left            =   480
         TabIndex        =   1
         Top             =   8175
         Width           =   1755
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   8565
         Index           =   1
         Left            =   240
         Top             =   210
         Width           =   11580
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FFFFFF&
         BorderWidth     =   2
         Height          =   8760
         Index           =   2
         Left            =   120
         Top             =   120
         Width           =   11805
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "1"
         BeginProperty Font 
            Name            =   "Bookman Old Style"
            Size            =   150
            Charset         =   0
            Weight          =   600
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   2865
         Left            =   240
         TabIndex        =   261
         Top             =   240
         Width           =   1695
      End
   End
End
Attribute VB_Name = "frm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit: Dim K

Private Sub cmdProsen_Click(): ' On Error Resume Next
Dim I, K, L, Bpoint(1 To 20)
If optB(0).Value = True Then Bpoint(1) = 1
If optB(1).Value = True Then Bpoint(1) = 2
If optB(2).Value = True Then Bpoint(1) = 3
If optB(3).Value = True Then Bpoint(1) = 4
If optB(4).Value = True Then Bpoint(1) = 5
If optB(5).Value = True Then Bpoint(1) = 6
If optB(6).Value = True Then Bpoint(1) = 7
If optB(7).Value = True Then Bpoint(1) = 8
If optB(8).Value = True Then Bpoint(1) = 9
If optB(9).Value = True Then Bpoint(1) = 10
   If optB(10).Value = True Then Bpoint(2) = 1
   If optB(11).Value = True Then Bpoint(2) = 2
   If optB(12).Value = True Then Bpoint(2) = 3
   If optB(13).Value = True Then Bpoint(2) = 4
   If optB(14).Value = True Then Bpoint(2) = 5
   If optB(15).Value = True Then Bpoint(2) = 6
   If optB(16).Value = True Then Bpoint(2) = 7
   If optB(17).Value = True Then Bpoint(2) = 8
   If optB(18).Value = True Then Bpoint(2) = 9
   If optB(19).Value = True Then Bpoint(2) = 10
      If optB(20).Value = True Then Bpoint(3) = 1
      If optB(21).Value = True Then Bpoint(3) = 2
      If optB(22).Value = True Then Bpoint(3) = 3
      If optB(23).Value = True Then Bpoint(3) = 4
      If optB(24).Value = True Then Bpoint(3) = 5
      If optB(25).Value = True Then Bpoint(3) = 6
      If optB(26).Value = True Then Bpoint(3) = 7
      If optB(27).Value = True Then Bpoint(3) = 8
      If optB(28).Value = True Then Bpoint(3) = 9
      If optB(29).Value = True Then Bpoint(3) = 10
If optB(30).Value = True Then Bpoint(4) = 1
If optB(31).Value = True Then Bpoint(4) = 2
If optB(32).Value = True Then Bpoint(4) = 3
If optB(33).Value = True Then Bpoint(4) = 4
If optB(34).Value = True Then Bpoint(4) = 5
If optB(35).Value = True Then Bpoint(4) = 6
If optB(36).Value = True Then Bpoint(4) = 7
If optB(37).Value = True Then Bpoint(4) = 8
If optB(38).Value = True Then Bpoint(4) = 9
If optB(39).Value = True Then Bpoint(4) = 10
   If optB(40).Value = True Then Bpoint(5) = 1
   If optB(41).Value = True Then Bpoint(5) = 2
   If optB(42).Value = True Then Bpoint(5) = 3
   If optB(43).Value = True Then Bpoint(5) = 4
   If optB(44).Value = True Then Bpoint(5) = 5
   If optB(45).Value = True Then Bpoint(5) = 6
   If optB(46).Value = True Then Bpoint(5) = 7
   If optB(47).Value = True Then Bpoint(5) = 8
   If optB(48).Value = True Then Bpoint(5) = 9
   If optB(49).Value = True Then Bpoint(5) = 10
      If optB(50).Value = True Then Bpoint(6) = 1
      If optB(51).Value = True Then Bpoint(6) = 2
      If optB(52).Value = True Then Bpoint(6) = 3
      If optB(53).Value = True Then Bpoint(6) = 4
      If optB(54).Value = True Then Bpoint(6) = 5
      If optB(55).Value = True Then Bpoint(6) = 6
      If optB(56).Value = True Then Bpoint(6) = 7
      If optB(57).Value = True Then Bpoint(6) = 8
      If optB(58).Value = True Then Bpoint(6) = 9
      If optB(59).Value = True Then Bpoint(6) = 10
         If optB(60).Value = True Then Bpoint(7) = 1
         If optB(61).Value = True Then Bpoint(7) = 2
         If optB(62).Value = True Then Bpoint(7) = 3
         If optB(63).Value = True Then Bpoint(7) = 4
         If optB(64).Value = True Then Bpoint(7) = 5
         If optB(65).Value = True Then Bpoint(7) = 6
         If optB(66).Value = True Then Bpoint(7) = 7
         If optB(67).Value = True Then Bpoint(7) = 8
         If optB(68).Value = True Then Bpoint(7) = 9
         If optB(69).Value = True Then Bpoint(7) = 10
If optB(70).Value = True Then Bpoint(8) = 1
If optB(71).Value = True Then Bpoint(8) = 2
If optB(72).Value = True Then Bpoint(8) = 3
If optB(73).Value = True Then Bpoint(8) = 4
If optB(74).Value = True Then Bpoint(8) = 5
If optB(75).Value = True Then Bpoint(8) = 6
If optB(76).Value = True Then Bpoint(8) = 7
If optB(77).Value = True Then Bpoint(8) = 8
If optB(78).Value = True Then Bpoint(8) = 9
If optB(79).Value = True Then Bpoint(8) = 10
   If optB(80).Value = True Then Bpoint(9) = 1
   If optB(81).Value = True Then Bpoint(9) = 2
   If optB(82).Value = True Then Bpoint(9) = 3
   If optB(83).Value = True Then Bpoint(9) = 4
   If optB(84).Value = True Then Bpoint(9) = 5
   If optB(85).Value = True Then Bpoint(9) = 6
   If optB(86).Value = True Then Bpoint(9) = 7
   If optB(87).Value = True Then Bpoint(9) = 8
   If optB(88).Value = True Then Bpoint(9) = 9
   If optB(89).Value = True Then Bpoint(9) = 10
      If optB(90).Value = True Then Bpoint(10) = 1
      If optB(91).Value = True Then Bpoint(10) = 2
      If optB(92).Value = True Then Bpoint(10) = 3
      If optB(93).Value = True Then Bpoint(10) = 4
      If optB(94).Value = True Then Bpoint(10) = 5
      If optB(95).Value = True Then Bpoint(10) = 6
      If optB(96).Value = True Then Bpoint(10) = 7
      If optB(97).Value = True Then Bpoint(10) = 8
      If optB(98).Value = True Then Bpoint(10) = 9
      If optB(99).Value = True Then Bpoint(10) = 10
         If optB(100).Value = True Then Bpoint(11) = 1
         If optB(101).Value = True Then Bpoint(11) = 2
         If optB(102).Value = True Then Bpoint(11) = 3
         If optB(103).Value = True Then Bpoint(11) = 4
         If optB(104).Value = True Then Bpoint(11) = 5
         If optB(105).Value = True Then Bpoint(11) = 6
         If optB(106).Value = True Then Bpoint(11) = 7
         If optB(107).Value = True Then Bpoint(11) = 8
         If optB(108).Value = True Then Bpoint(11) = 9
         If optB(109).Value = True Then Bpoint(11) = 10
If optB(110).Value = True Then Bpoint(12) = 1
If optB(111).Value = True Then Bpoint(12) = 2
If optB(112).Value = True Then Bpoint(12) = 3
If optB(113).Value = True Then Bpoint(12) = 4
If optB(114).Value = True Then Bpoint(12) = 5
If optB(115).Value = True Then Bpoint(12) = 6
If optB(116).Value = True Then Bpoint(12) = 7
If optB(117).Value = True Then Bpoint(12) = 8
If optB(118).Value = True Then Bpoint(12) = 9
If optB(119).Value = True Then Bpoint(12) = 10
   If optB(120).Value = True Then Bpoint(13) = 1
   If optB(121).Value = True Then Bpoint(13) = 2
   If optB(122).Value = True Then Bpoint(13) = 3
   If optB(123).Value = True Then Bpoint(13) = 4
   If optB(124).Value = True Then Bpoint(13) = 5
   If optB(125).Value = True Then Bpoint(13) = 6
   If optB(126).Value = True Then Bpoint(13) = 7
   If optB(127).Value = True Then Bpoint(13) = 8
   If optB(128).Value = True Then Bpoint(13) = 9
   If optB(129).Value = True Then Bpoint(13) = 10
      If optB(130).Value = True Then Bpoint(14) = 1
      If optB(131).Value = True Then Bpoint(14) = 2
      If optB(132).Value = True Then Bpoint(14) = 3
      If optB(133).Value = True Then Bpoint(14) = 4
      If optB(134).Value = True Then Bpoint(14) = 5
      If optB(135).Value = True Then Bpoint(14) = 6
      If optB(136).Value = True Then Bpoint(14) = 7
      If optB(137).Value = True Then Bpoint(14) = 8
      If optB(138).Value = True Then Bpoint(14) = 9
      If optB(139).Value = True Then Bpoint(14) = 10
         If optB(140).Value = True Then Bpoint(15) = 1
         If optB(141).Value = True Then Bpoint(15) = 2
         If optB(142).Value = True Then Bpoint(15) = 3
         If optB(143).Value = True Then Bpoint(15) = 4
         If optB(144).Value = True Then Bpoint(15) = 5
         If optB(145).Value = True Then Bpoint(15) = 6
         If optB(146).Value = True Then Bpoint(15) = 7
         If optB(147).Value = True Then Bpoint(15) = 8
         If optB(148).Value = True Then Bpoint(15) = 9
         If optB(149).Value = True Then Bpoint(15) = 10
If optB(150).Value = True Then Bpoint(16) = 1
If optB(151).Value = True Then Bpoint(16) = 2
If optB(152).Value = True Then Bpoint(16) = 3
If optB(153).Value = True Then Bpoint(16) = 4
If optB(154).Value = True Then Bpoint(16) = 5
If optB(155).Value = True Then Bpoint(16) = 6
If optB(156).Value = True Then Bpoint(16) = 7
If optB(157).Value = True Then Bpoint(16) = 8
If optB(158).Value = True Then Bpoint(16) = 9
If optB(159).Value = True Then Bpoint(16) = 10
   If optB(160).Value = True Then Bpoint(17) = 1
   If optB(161).Value = True Then Bpoint(17) = 2
   If optB(162).Value = True Then Bpoint(17) = 3
   If optB(163).Value = True Then Bpoint(17) = 4
   If optB(164).Value = True Then Bpoint(17) = 5
   If optB(165).Value = True Then Bpoint(17) = 6
   If optB(166).Value = True Then Bpoint(17) = 7
   If optB(167).Value = True Then Bpoint(17) = 8
   If optB(168).Value = True Then Bpoint(17) = 9
   If optB(169).Value = True Then Bpoint(17) = 10
      If optB(170).Value = True Then Bpoint(18) = 1
      If optB(171).Value = True Then Bpoint(18) = 2
      If optB(172).Value = True Then Bpoint(18) = 3
      If optB(173).Value = True Then Bpoint(18) = 4
      If optB(174).Value = True Then Bpoint(18) = 5
      If optB(175).Value = True Then Bpoint(18) = 6
      If optB(176).Value = True Then Bpoint(18) = 7
      If optB(177).Value = True Then Bpoint(18) = 8
      If optB(178).Value = True Then Bpoint(18) = 9
      If optB(179).Value = True Then Bpoint(18) = 10
         If optB(180).Value = True Then Bpoint(19) = 1
         If optB(181).Value = True Then Bpoint(19) = 2
         If optB(182).Value = True Then Bpoint(19) = 3
         If optB(183).Value = True Then Bpoint(19) = 4
         If optB(184).Value = True Then Bpoint(19) = 5
         If optB(185).Value = True Then Bpoint(19) = 6
         If optB(186).Value = True Then Bpoint(19) = 7
         If optB(187).Value = True Then Bpoint(19) = 8
         If optB(188).Value = True Then Bpoint(19) = 9
         If optB(189).Value = True Then Bpoint(19) = 10
If optB(190).Value = True Then Bpoint(20) = 1
If optB(191).Value = True Then Bpoint(20) = 2
If optB(192).Value = True Then Bpoint(20) = 3
If optB(193).Value = True Then Bpoint(20) = 4
If optB(194).Value = True Then Bpoint(20) = 5
If optB(195).Value = True Then Bpoint(20) = 6
If optB(196).Value = True Then Bpoint(20) = 7
If optB(197).Value = True Then Bpoint(20) = 8
If optB(198).Value = True Then Bpoint(20) = 9
If optB(199).Value = True Then Bpoint(20) = 10

K = Bpoint(1) + Bpoint(2) + Bpoint(3) + Bpoint(4) + Bpoint(5) + Bpoint(6) + _
Bpoint(7) + Bpoint(8) + Bpoint(9) + Bpoint(10) + Bpoint(11) + Bpoint(12) + _
Bpoint(13) + Bpoint(14) + Bpoint(15) + Bpoint(16) + Bpoint(17) + Bpoint(18) + _
Bpoint(19) + Bpoint(20)
MsgBox "Total Nilai Anda = " & K & vbCrLf & "Prosentasi Nilai Anda = " & K / 20
TotalKolom1 = K: Unload Me

End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer): If KeyCode = vbKeyEscape Then Unload Me
End Sub
Private Sub Form_Load(): On Error Resume Next: Dim I: For I = optB.lBound To optB.UBound: optB(I).BackColor = vbBlack: Next: End Sub
Private Sub Form_Resize(): picS.Left = (Me.Width - Me.picS.Width) / 2
    Me.Move 0, 0, Screen.Width, Screen.Height
    Me.Picture4.Move (Me.ScaleWidth - Me.Picture4.ScaleWidth) \ 2, _
    (Me.ScaleHeight - Me.Picture4.ScaleHeight) / 2
End Sub
Private Sub Form_Unload(Cancel As Integer)
frm2.Show: End Sub


