VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   7470
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12300
   LinkTopic       =   "Form2"
   ScaleHeight     =   7470
   ScaleWidth      =   12300
   StartUpPosition =   3  'Windows Default
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2175
      Left            =   840
      TabIndex        =   10
      Top             =   5040
      Width           =   7575
      _ExtentX        =   13361
      _ExtentY        =   3836
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Comkeluar 
      Caption         =   "KELUAR"
      Height          =   495
      Left            =   5760
      TabIndex        =   9
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Comhapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   3120
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Comsimpan 
      Caption         =   "SIMPAN"
      Height          =   615
      Left            =   480
      TabIndex        =   7
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txtberatemas 
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   2520
      Width           =   2895
   End
   Begin VB.TextBox txtemas 
      Height          =   375
      Left            =   3840
      TabIndex        =   5
      Top             =   1920
      Width           =   3015
   End
   Begin VB.TextBox txtidemas 
      Height          =   375
      Left            =   3840
      TabIndex        =   4
      Top             =   1320
      Width           =   3015
   End
   Begin VB.Label Label4 
      Caption         =   "BERAT EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2520
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1920
      Width           =   2535
   End
   Begin VB.Label Label2 
      BorderStyle     =   1  'Fixed Single
      Caption         =   "ID EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1320
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "DATA EMAS PEGADAIAN"
      Height          =   615
      Left            =   3240
      TabIndex        =   0
      Top             =   360
      Width           =   4455
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   -240
      Picture         =   "Form2.frx":0000
      Top             =   -960
      Width           =   23400
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim emaspegadaian As New ADODB.Recordset
Private Sub Comhapus_Click()
koneksidb.Execute "delete from tabel_emaspegadaian where idemas='" & txtidemas & "'"
Call refreshh
Call kosong
txtidemas.SetFocus
End Sub

Private Sub Comkeluar_Click()
x = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If x = vbYes Then End
End Sub

Private Sub Comsimpan_Click()
If txtidemas = "" Then
MsgBox "Id Emas Kosong", vbExclamation, "pesan"
txtidemas.SetFocus
Exit Sub
End If
    If txtemas = "" Then
    MsgBox "Emas Kosong", vbExclamation, "pesan"
    txtemas.SetFocus
    Exit Sub
    End If
If txtberatemas = "" Then
MsgBox "Berat Emas Kosong", vbExclamation, "pesan"
txtberatemas.SetFocus
Exit Sub
End If
Set emaspegadaian = New ADODB.Recordset
emaspegadaian.Open "select*from tabel_emaspegadaian where idemas='" & txtidemas & "'", koneksidb
If Not emaspegadaian.EOF Then
MsgBox "Id Emas sudah ada", vbCritical, "pesan"
txtidemas = ""
txtidemas.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tabel_emaspegadaian(idemas,emas,beratemas) value ('" & txtidemas & "','" & txtemas & "','" & txtberatemas & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = emaspegadaian
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub DataGrid1_Click()
txtidemas.Text = emaspegadaian!idemas
txtemas.Text = emaspegadaian!emas
txtberatemas.Text = emaspegadaian!beratemas
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = emaspegadaian
With emaspegadaian
End With
Call edit_grid
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Id emas"
    .Columns(1).Caption = "Emas"
    .Columns(2).Caption = "Berat Emas"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
End With
End Sub

Sub tampil_data()
Set emaspegadaian = New ADODB.Recordset
emaspegadaian.ActiveConnection = koneksidb
emaspegadaian.CursorLocation = adUseClient
emaspegadaian.LockType = adLockOptimistic
emaspegadaian.Source = "select * from tabel_emaspegadaian"
emaspegadaian.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = emaspegadaian
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = emaspegadaian
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txtidemas = ""
txtemas = ""
txtberatemas = ""
txtidemas.SetFocus
End Sub

