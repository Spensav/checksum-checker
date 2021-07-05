VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   Caption         =   "AutoCeksumer :D, Minimalis...."
   ClientHeight    =   4815
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6735
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   6735
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   120
      TabIndex        =   6
      Text            =   "W32."
      Top             =   600
      Width           =   1695
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Sampai Sampai Kedalam Folder...."
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   4200
      Width           =   3495
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Browse"
      Height          =   375
      Left            =   5640
      TabIndex        =   4
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   3840
      TabIndex        =   3
      Top             =   4200
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   5415
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Tampilkan File"
      BeginProperty Font 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   5160
      TabIndex        =   1
      Top             =   4200
      Width           =   1455
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   3135
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   6495
      _ExtentX        =   11456
      _ExtentY        =   5530
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Calibri"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   2
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Lokasi"
         Object.Width           =   7832
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Ceksum"
         Object.Width           =   3246
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'SOURCE CODE SIMPEL, UNTUk MENENTUKAN CEKSUM BUATAN SAYA LHO...HEHEHEHE
'BY.SPENSAV ANTIVIRUS | 2009 - 2012
'Isfahani Ghiyath
'======================================================================
'PENJELASAN :
'- Masuk referensi
'- tambahkan Referensi "Microsoft Scripting Runtime"
'.....
Dim Gitu As ListItem 'Untuk Item ListViewnya :D
Dim isfahani As Boolean 'Ini Fungsi Berhenti ya...
Sub Spensav(ByVal Fol As Scripting.Folder)
If Berhenti = True Then Exit Sub
Dim C As New cCommonDialog
   Dim sFileName As String
   Dim ceksum As String
   Dim m_CRC As ClsCRC
   Dim namavirus As String
   Set m_CRC = New ClsCRC
On Error Resume Next
Dim fi As Scripting.File
Dim fo As Scripting.Folder
  
For Each fi In Fol.Files
DoEvents
If Berhenti = True Then Exit Sub
Set Gitu = ListView1.ListItems.Add(, , fi.path)
Gitu.SubItems(1) = Hex(m_CRC.CalculateFile(fi.path))
Next

If Check1.Value = 1 Then
For Each fo In Fol.SubFolders  '|
Spensav fo                     '|> INI FUNGSI MENCARI FILE SAMPAI KEDALAM-
Next                           '|> DALAM FOLDER YANG ADA DI PATH....
Else
End If
End Sub
Private Sub Command1_Click()
If Command1.Caption = "Tampilkan File" Then
ListView1.ListItems.Clear
isfahani = False
Command1.Caption = "Berhenti"
Dim fso As New FileSystemObject
Spensav fso.GetFolder(Text1.Text) 'Text1.Text adalah Pathnya :D
Command1.Caption = "Tampilkan File"
Else 'stop
Berhenti = True
Command1.Caption = "Tampilkan File"
End If
End Sub
'WARNING...!
Private Sub Command2_Click()
On Error Resume Next
Dim i As Integer
Open App.path & "\" & "asyike.txt" For Output As #1
For i = 1 To ListView1.ListItems.Count
Print #1, "Master Ari, Nih Sampelnya, Emang Ada Spasinya di Kolom Nomor Urut,"
Print #1, "Tetapi Ini AUtoCheksumer Buat SC Adesinichi.."
Print #1, "Hiraukan saja Spasinya, Kalo Udah di Cheksumer"
Print #1, "Copy aja langsung ke Database"
Print #1, "By.Isfahani Master Newbie"
Print #1, ""
Print #1, ListView1.ListItems(i).SubItems(1) & ":" & GetFileName(ListView1.ListItems(i))
Next i
Close #i
End Sub
Private Sub Command3_Click()
Text1.Text = BrowseFolder("Select Path To Scan", Me)
End Sub
Private Function GetFileName(path As String) As String
GetFileName = Right$(path, Len(path) - InStrRev(path, "\"))
End Function
Private Sub Form_Load()
Combo1.AddItem "W32."
Combo1.AddItem "Worm."
Combo1.AddItem "Item.Virut."
Combo1.AddItem "Trojan."
Combo1.AddItem "Spy."
Combo1.AddItem "Conficker."
End Sub
