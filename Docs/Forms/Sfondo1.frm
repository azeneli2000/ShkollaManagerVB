VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{572FF236-2066-11D4-8ED4-00E07D815373}#1.0#0"; "MBMsgEx.ocx"
Begin VB.Form Sfondo 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   $"Sfondo1.frx":0000
   ClientHeight    =   12585
   ClientLeft      =   0
   ClientTop       =   345
   ClientWidth     =   17190
   Enabled         =   0   'False
   FillColor       =   &H00800000&
   Icon            =   "Sfondo1.frx":00DC
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   12585
   ScaleWidth      =   17190
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar2 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      ToolTipText     =   "Questo Status bar appare qui solamente in esempio, ma non deve apparire all'ACQUIRENTE! E' Chiaro?"
      Top             =   12210
      Width           =   17190
      _ExtentX        =   30321
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3870
      Top             =   6600
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.FileListBox File1 
      Height          =   2625
      Left            =   1680
      Pattern         =   "*.frm;*.cls;*.bas;*.Vbp"
      TabIndex        =   0
      Top             =   3840
      Visible         =   0   'False
      Width           =   1755
   End
   Begin MSComctlLib.ImageList Iml 
      Left            =   4530
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   131
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":03E6
            Key             =   "new"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":0542
            Key             =   "open"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":069E
            Key             =   "close"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":07FA
            Key             =   "exit"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":0956
            Key             =   "save"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":0EF2
            Key             =   "preview"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":104E
            Key             =   "print"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11AA
            Key             =   "property"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1306
            Key             =   "copy"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1462
            Key             =   "paste"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":15BE
            Key             =   "clipboard"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":171A
            Key             =   "filter"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1CB6
            Key             =   "find"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":2252
            Key             =   "findnext"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":27EE
            Key             =   "findprev"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":2D8A
            Key             =   "zoomin"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":2EE6
            Key             =   "zoomout"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":3042
            Key             =   "toolbar"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":319E
            Key             =   "front"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":32FA
            Key             =   "back"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":3456
            Key             =   "plus"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":35B2
            Key             =   "meno"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":370E
            Key             =   "eyedropper"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":3CAA
            Key             =   "len"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":3E06
            Key             =   "mail"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":3F62
            Key             =   "update"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":44FE
            Key             =   "help"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":465A
            Key             =   "helpcontents"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":4836
            Key             =   "helpfind"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":4992
            Key             =   "helptips"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":4F2E
            Key             =   "helptopic"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":54CA
            Key             =   "a1"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":591C
            Key             =   "a2"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":5D6E
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":61C0
            Key             =   "a4"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":6612
            Key             =   "a5"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":6A64
            Key             =   "a6"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":6EB6
            Key             =   "a7"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":7308
            Key             =   "a8"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":775A
            Key             =   "a9"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":7BAC
            Key             =   "a10"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":7FFE
            Key             =   "a11"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":8450
            Key             =   "a12"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":85AA
            Key             =   "a13"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":8704
            Key             =   "a14"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":885E
            Key             =   "a15"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":8CB0
            Key             =   "a16"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9102
            Key             =   "a17"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":925C
            Key             =   "a19"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":93B6
            Key             =   "a20"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9510
            Key             =   "a21"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":966A
            Key             =   "a22"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":97C4
            Key             =   "a23"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":991E
            Key             =   "a25"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9A78
            Key             =   "a26"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9BD2
            Key             =   "a27"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9D2C
            Key             =   "a28"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9E86
            Key             =   "a29"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":9FE0
            Key             =   "a30"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A13A
            Key             =   "a31"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A294
            Key             =   "a32"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A3EE
            Key             =   "a33"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A548
            Key             =   "a34"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A6A2
            Key             =   "a35"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A7FC
            Key             =   "a36"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":A956
            Key             =   "a37"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":AAB0
            Key             =   "a38"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":AC0A
            Key             =   "a39"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":AD64
            Key             =   "a40"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":AEBE
            Key             =   "a41"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B018
            Key             =   "a42"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B172
            Key             =   "a43"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B2CC
            Key             =   "a44"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B426
            Key             =   "a45"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B580
            Key             =   "a46"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B6DA
            Key             =   "a47"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B834
            Key             =   "a48"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":B98E
            Key             =   "a49"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":BAE8
            Key             =   "a50"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":BC42
            Key             =   "a51"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":BD9C
            Key             =   "a52"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":BEF6
            Key             =   "a53"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C050
            Key             =   "a54"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C1AA
            Key             =   "a55"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C304
            Key             =   "a56"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C45E
            Key             =   "a57"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C5B8
            Key             =   "a58"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C712
            Key             =   "a59"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C86C
            Key             =   "a60"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":C9C6
            Key             =   "a61"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":CB20
            Key             =   "a62"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":CC7A
            Key             =   "a63"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":CDD4
            Key             =   "a64"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":CF2E
            Key             =   "a65"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D088
            Key             =   "a66"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D1E2
            Key             =   "a67"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D33C
            Key             =   "a68"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D496
            Key             =   "a69"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D5F0
            Key             =   "a70"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D74A
            Key             =   "a71"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D8A4
            Key             =   "a72"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":D9FE
            Key             =   "a73"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":DB58
            Key             =   "a74"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":DCB2
            Key             =   "a75"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":DE0C
            Key             =   "a76"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":DF66
            Key             =   "a77"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E0C0
            Key             =   "a78"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E21A
            Key             =   "a79"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E374
            Key             =   "a80"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E4CE
            Key             =   "a81"
         EndProperty
         BeginProperty ListImage111 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E628
            Key             =   "a82"
         EndProperty
         BeginProperty ListImage112 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E782
            Key             =   "a83"
         EndProperty
         BeginProperty ListImage113 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":E8DC
            Key             =   "a84"
         EndProperty
         BeginProperty ListImage114 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":EA36
            Key             =   "a85"
         EndProperty
         BeginProperty ListImage115 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":EB90
            Key             =   "a86"
         EndProperty
         BeginProperty ListImage116 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":ECEA
            Key             =   "a87"
         EndProperty
         BeginProperty ListImage117 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":EE44
            Key             =   "a88"
         EndProperty
         BeginProperty ListImage118 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":EF9E
            Key             =   "a89"
         EndProperty
         BeginProperty ListImage119 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F0F8
            Key             =   "a90"
         EndProperty
         BeginProperty ListImage120 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F252
            Key             =   "a91"
         EndProperty
         BeginProperty ListImage121 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F3AC
            Key             =   "a92"
         EndProperty
         BeginProperty ListImage122 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F506
            Key             =   "a93"
         EndProperty
         BeginProperty ListImage123 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F660
            Key             =   "a94"
         EndProperty
         BeginProperty ListImage124 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F7BA
            Key             =   "a95"
         EndProperty
         BeginProperty ListImage125 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":F914
            Key             =   "a96"
         EndProperty
         BeginProperty ListImage126 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":FA6E
            Key             =   "a97"
         EndProperty
         BeginProperty ListImage127 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":FBC8
            Key             =   "a98"
         EndProperty
         BeginProperty ListImage128 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":FD22
            Key             =   "a99"
         EndProperty
         BeginProperty ListImage129 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":FE7C
            Key             =   "a100"
         EndProperty
         BeginProperty ListImage130 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":FFD6
            Key             =   "a101"
         EndProperty
         BeginProperty ListImage131 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10130
            Key             =   "a102"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList Img16x16 
      Left            =   4530
      Top             =   7320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   110
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1028A
            Key             =   "a1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":103E4
            Key             =   "a2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1053E
            Key             =   "a3"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10698
            Key             =   "a4"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":107F2
            Key             =   "a5"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1094C
            Key             =   "a6"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10AA6
            Key             =   "a7"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10C00
            Key             =   "a8"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10D5A
            Key             =   "a9"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":10EB4
            Key             =   "a10"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1100E
            Key             =   "a11"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11168
            Key             =   "a12"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":112C2
            Key             =   "a13"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1141C
            Key             =   "a14"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11576
            Key             =   "a15"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":116D0
            Key             =   "a16"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1182A
            Key             =   "a17"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11984
            Key             =   "a18"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11ADE
            Key             =   "a19"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11C38
            Key             =   "a20"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11D92
            Key             =   "a21"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":11EEC
            Key             =   "a22"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":12046
            Key             =   "a23"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":121A0
            Key             =   "a24"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":122FA
            Key             =   "a25"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":12454
            Key             =   "a26"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":125AE
            Key             =   "a27"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":12708
            Key             =   "a28"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":12EDA
            Key             =   "a29"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13034
            Key             =   "a30"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1318E
            Key             =   "a31"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":132E8
            Key             =   "a32"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13442
            Key             =   "a33"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1359C
            Key             =   "a34"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":136F6
            Key             =   "a35"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13850
            Key             =   "a36"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":139AA
            Key             =   "a37"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13B04
            Key             =   "a38"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13C5E
            Key             =   "a39"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13DB8
            Key             =   "a40"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":13F12
            Key             =   "a41"
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1406C
            Key             =   "a42"
         EndProperty
         BeginProperty ListImage43 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":141C6
            Key             =   "a43"
         EndProperty
         BeginProperty ListImage44 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14320
            Key             =   "a44"
         EndProperty
         BeginProperty ListImage45 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1447A
            Key             =   "a45"
         EndProperty
         BeginProperty ListImage46 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":145D4
            Key             =   "a46"
         EndProperty
         BeginProperty ListImage47 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1472E
            Key             =   "a47"
         EndProperty
         BeginProperty ListImage48 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14888
            Key             =   "a48"
         EndProperty
         BeginProperty ListImage49 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":149E2
            Key             =   "a49"
         EndProperty
         BeginProperty ListImage50 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14B3C
            Key             =   "a50"
         EndProperty
         BeginProperty ListImage51 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14C96
            Key             =   "a51"
         EndProperty
         BeginProperty ListImage52 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14DF0
            Key             =   "a52"
         EndProperty
         BeginProperty ListImage53 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":14F4A
            Key             =   "a53"
         EndProperty
         BeginProperty ListImage54 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":176FC
            Key             =   "a54"
         EndProperty
         BeginProperty ListImage55 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":17856
            Key             =   "a55"
         EndProperty
         BeginProperty ListImage56 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":179B0
            Key             =   "a56"
         EndProperty
         BeginProperty ListImage57 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":17B0A
            Key             =   "a57"
         EndProperty
         BeginProperty ListImage58 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":17C64
            Key             =   "a58"
         EndProperty
         BeginProperty ListImage59 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":17DBE
            Key             =   "a59"
         EndProperty
         BeginProperty ListImage60 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":17F18
            Key             =   "a60"
         EndProperty
         BeginProperty ListImage61 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18072
            Key             =   "a61"
         EndProperty
         BeginProperty ListImage62 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":181CC
            Key             =   "a62"
         EndProperty
         BeginProperty ListImage63 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18326
            Key             =   "a63"
         EndProperty
         BeginProperty ListImage64 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18480
            Key             =   "a64"
         EndProperty
         BeginProperty ListImage65 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":185DA
            Key             =   "a65"
         EndProperty
         BeginProperty ListImage66 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18734
            Key             =   "a66"
         EndProperty
         BeginProperty ListImage67 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1888E
            Key             =   "a67"
         EndProperty
         BeginProperty ListImage68 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":189E8
            Key             =   "a68"
         EndProperty
         BeginProperty ListImage69 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18B42
            Key             =   "a69"
         EndProperty
         BeginProperty ListImage70 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18C9C
            Key             =   "a70"
         EndProperty
         BeginProperty ListImage71 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18DF6
            Key             =   "a71"
         EndProperty
         BeginProperty ListImage72 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":18F50
            Key             =   "a72"
         EndProperty
         BeginProperty ListImage73 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":190AA
            Key             =   "a73"
         EndProperty
         BeginProperty ListImage74 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19204
            Key             =   "a74"
         EndProperty
         BeginProperty ListImage75 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1935E
            Key             =   "a75"
         EndProperty
         BeginProperty ListImage76 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":194B8
            Key             =   "a76"
         EndProperty
         BeginProperty ListImage77 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19612
            Key             =   "a77"
         EndProperty
         BeginProperty ListImage78 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1976C
            Key             =   "a78"
         EndProperty
         BeginProperty ListImage79 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":198C6
            Key             =   "a79"
         EndProperty
         BeginProperty ListImage80 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19A20
            Key             =   "a80"
         EndProperty
         BeginProperty ListImage81 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19B7A
            Key             =   "a81"
         EndProperty
         BeginProperty ListImage82 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19CD4
            Key             =   "a82"
         EndProperty
         BeginProperty ListImage83 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19E2E
            Key             =   "a83"
         EndProperty
         BeginProperty ListImage84 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":19F88
            Key             =   "a84"
         EndProperty
         BeginProperty ListImage85 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A0E2
            Key             =   "a85"
         EndProperty
         BeginProperty ListImage86 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A23C
            Key             =   "a86"
         EndProperty
         BeginProperty ListImage87 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A396
            Key             =   "a87"
         EndProperty
         BeginProperty ListImage88 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A4F0
            Key             =   "a88"
         EndProperty
         BeginProperty ListImage89 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A64A
            Key             =   "a89"
         EndProperty
         BeginProperty ListImage90 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A7A4
            Key             =   "a90"
         EndProperty
         BeginProperty ListImage91 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1A8FE
            Key             =   "a91"
         EndProperty
         BeginProperty ListImage92 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1AA58
            Key             =   "a92"
         EndProperty
         BeginProperty ListImage93 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1ABB2
            Key             =   "a93"
         EndProperty
         BeginProperty ListImage94 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1AD0C
            Key             =   "a94"
         EndProperty
         BeginProperty ListImage95 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1AE66
            Key             =   "a95"
         EndProperty
         BeginProperty ListImage96 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1AFC0
            Key             =   "a96"
         EndProperty
         BeginProperty ListImage97 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1B11A
            Key             =   "a97"
         EndProperty
         BeginProperty ListImage98 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1B274
            Key             =   "a98"
         EndProperty
         BeginProperty ListImage99 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1B3CE
            Key             =   "a99"
         EndProperty
         BeginProperty ListImage100 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1B528
            Key             =   "a100"
         EndProperty
         BeginProperty ListImage101 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1B682
            Key             =   "a101"
         EndProperty
         BeginProperty ListImage102 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C03C
            Key             =   "a102"
         EndProperty
         BeginProperty ListImage103 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C196
            Key             =   "a103"
         EndProperty
         BeginProperty ListImage104 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C2F0
            Key             =   "a104"
         EndProperty
         BeginProperty ListImage105 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C742
            Key             =   "a105"
         EndProperty
         BeginProperty ListImage106 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C89C
            Key             =   "a106"
         EndProperty
         BeginProperty ListImage107 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1C9F6
            Key             =   "a107"
         EndProperty
         BeginProperty ListImage108 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1CB50
            Key             =   "a108"
         EndProperty
         BeginProperty ListImage109 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1CCAA
            Key             =   "a109"
         EndProperty
         BeginProperty ListImage110 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Sfondo1.frx":1CE04
            Key             =   "a110"
         EndProperty
      EndProperty
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFF80&
      Caption         =   "VisionInfoSolution"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   36
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   975
      Left            =   9360
      TabIndex        =   4
      Top             =   9600
      Width           =   6015
   End
   Begin VB.Label Label2 
      Caption         =   "InfoShkolla"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   120
      TabIndex        =   3
      Top             =   960
      Width           =   6615
   End
   Begin VB.Label Label1 
      Caption         =   "REGJISTRIMI I PROGRAMIT "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   615
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   6615
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   47
      Left            =   8940
      Picture         =   "Sfondo1.frx":1CF5E
      Top             =   5310
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   210
      Index           =   43
      Left            =   6000
      Picture         =   "Sfondo1.frx":1D268
      Top             =   5940
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   195
      Index           =   44
      Left            =   6660
      Picture         =   "Sfondo1.frx":1D512
      Top             =   5940
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   45
      Left            =   7305
      Picture         =   "Sfondo1.frx":1D790
      Top             =   5940
      Visible         =   0   'False
      Width           =   225
   End
   Begin VB.Image Image1 
      Height          =   225
      Index           =   46
      Left            =   7950
      Picture         =   "Sfondo1.frx":1DAA2
      Top             =   5940
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   12
      Left            =   5730
      Picture         =   "Sfondo1.frx":1DD78
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   11
      Left            =   5130
      Picture         =   "Sfondo1.frx":1E082
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   10
      Left            =   4530
      Picture         =   "Sfondo1.frx":1E38C
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   9
      Left            =   3870
      Picture         =   "Sfondo1.frx":1E696
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   0
      Left            =   3870
      Picture         =   "Sfondo1.frx":1E9A0
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   13
      Left            =   6390
      Picture         =   "Sfondo1.frx":1ECAA
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   8
      Left            =   8280
      Picture         =   "Sfondo1.frx":1EFB4
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   7
      Left            =   7680
      Picture         =   "Sfondo1.frx":1F2BE
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   6
      Left            =   7050
      Picture         =   "Sfondo1.frx":1F5C8
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   5
      Left            =   6390
      Picture         =   "Sfondo1.frx":1F8D2
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   4
      Left            =   5730
      Picture         =   "Sfondo1.frx":1FBDC
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   3
      Left            =   5130
      Picture         =   "Sfondo1.frx":1FEE6
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   1
      Left            =   4530
      Picture         =   "Sfondo1.frx":201F0
      Top             =   2850
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   2
      Left            =   7050
      Picture         =   "Sfondo1.frx":204FA
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin MBMsgBoxEx.MsgBoxEx Msg1 
      Left            =   8280
      Top             =   5940
      _ExtentX        =   847
      _ExtentY        =   847
      Buttons         =   2
      CustomIcon      =   "Sfondo1.frx":20804
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   14
      Left            =   7680
      Picture         =   "Sfondo1.frx":20820
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   15
      Left            =   8280
      Picture         =   "Sfondo1.frx":210EA
      Stretch         =   -1  'True
      Top             =   3480
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   16
      Left            =   3870
      Picture         =   "Sfondo1.frx":213F4
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   17
      Left            =   4530
      Picture         =   "Sfondo1.frx":21706
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   18
      Left            =   5130
      Picture         =   "Sfondo1.frx":21A10
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   19
      Left            =   5730
      Picture         =   "Sfondo1.frx":222DA
      Stretch         =   -1  'True
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   20
      Left            =   6390
      Picture         =   "Sfondo1.frx":225E4
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   21
      Left            =   7080
      Picture         =   "Sfondo1.frx":228EE
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   22
      Left            =   7680
      Picture         =   "Sfondo1.frx":22BF8
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   23
      Left            =   8280
      Picture         =   "Sfondo1.frx":22F02
      Top             =   4050
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   24
      Left            =   3870
      Picture         =   "Sfondo1.frx":2320C
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   25
      Left            =   4530
      Picture         =   "Sfondo1.frx":23516
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   26
      Left            =   5130
      Picture         =   "Sfondo1.frx":23820
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   27
      Left            =   5730
      Picture         =   "Sfondo1.frx":23B2A
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   28
      Left            =   6390
      Picture         =   "Sfondo1.frx":262CC
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   29
      Left            =   7050
      Picture         =   "Sfondo1.frx":265D6
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   30
      Left            =   7680
      Picture         =   "Sfondo1.frx":268E0
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   31
      Left            =   8280
      Picture         =   "Sfondo1.frx":26BEA
      Stretch         =   -1  'True
      Top             =   4680
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   32
      Left            =   3900
      Picture         =   "Sfondo1.frx":26EF4
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   33
      Left            =   4530
      Picture         =   "Sfondo1.frx":271FE
      Stretch         =   -1  'True
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   34
      Left            =   5130
      Picture         =   "Sfondo1.frx":28040
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   35
      Left            =   5730
      Picture         =   "Sfondo1.frx":2834A
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   36
      Left            =   6390
      Picture         =   "Sfondo1.frx":28654
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   37
      Left            =   7050
      Picture         =   "Sfondo1.frx":28F1E
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   38
      Left            =   7680
      Picture         =   "Sfondo1.frx":297E8
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   39
      Left            =   8280
      Picture         =   "Sfondo1.frx":2A0B2
      Top             =   5280
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   40
      Left            =   3870
      Picture         =   "Sfondo1.frx":2A97C
      Top             =   5940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   41
      Left            =   4530
      Picture         =   "Sfondo1.frx":2AC86
      Top             =   5940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image Image1 
      Height          =   480
      Index           =   42
      Left            =   5130
      Picture         =   "Sfondo1.frx":2AF90
      Top             =   5940
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgErrore 
      Height          =   480
      Left            =   14730
      Picture         =   "Sfondo1.frx":2B85A
      ToolTipText     =   "La procedura ... Ha causato Errore"
      Top             =   60
      Visible         =   0   'False
      Width           =   480
   End
   Begin VB.Image ImgGibra 
      Height          =   6780
      Left            =   1440
      Picture         =   "Sfondo1.frx":2BB64
      Top             =   2280
      Visible         =   0   'False
      Width           =   9420
   End
End
Attribute VB_Name = "Sfondo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
     Option Explicit
     

'*******************  Programmazione Visual basic  ******************
'*                           Programmatore                          *
'*                          by Paolo Puglisi                        *
'*                        vb6access@hotmail.com                     *
'*                  ************************************            *
'*                  * Creazione del                    *            *
'*                  ************************************            *
'*                  *    Modifica del 28 giugno 2003   *            *
'*                  ************************************            *
'********************************************************************
       
     Dim Para1 As Long
     Dim TitoloApplicazione, SottoTitoloApplicazione As String


'*********************************************************
'*                                                       *
Private Sub Form_Resize()
'*                                                       *
'*********************************************************
     Screen.MousePointer = 0
     On Error Resume Next
     Dim a1, b1, c1 As Long
     a1 = 255
     b1 = 175

     With Me
           .AutoRedraw = True
           .DrawStyle = vbInsideSolid
           .DrawMode = vbCopyPen
           .ScaleMode = vbPixels
           .DrawWidth = 3
           .ScaleHeight = 256
     End With

      Dim intLoop As Long

      For intLoop = 255 To 0 Step -1
            Me.Line (0, intLoop)-(Screen.Width, intLoop - 1), RGB(a1, b1, intLoop), B
      Next

 End Sub

