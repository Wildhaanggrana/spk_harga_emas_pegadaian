VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form3 
   Caption         =   "Form3"
   ClientHeight    =   7785
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12255
   LinkTopic       =   "Form3"
   ScaleHeight     =   7785
   ScaleWidth      =   12255
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox txtemas 
      Height          =   405
      Left            =   3960
      TabIndex        =   13
      Top             =   1680
      Width           =   3135
   End
   Begin VB.TextBox txthargapasaran 
      Height          =   375
      Left            =   3960
      TabIndex        =   12
      Top             =   2880
      Width           =   3135
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2295
      Left            =   360
      TabIndex        =   10
      Top             =   5160
      Width           =   8775
      _ExtentX        =   15478
      _ExtentY        =   4048
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
      Left            =   7080
      TabIndex        =   9
      Top             =   4320
      Width           =   1815
   End
   Begin VB.CommandButton Comhapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4800
      TabIndex        =   8
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Comedit 
      Caption         =   "EDIT"
      Height          =   495
      Left            =   2520
      TabIndex        =   7
      Top             =   4320
      Width           =   1695
   End
   Begin VB.CommandButton Comsimpan 
      Caption         =   "SIMPAN"
      Height          =   495
      Left            =   240
      TabIndex        =   6
      Top             =   4320
      Width           =   1575
   End
   Begin VB.TextBox txtberatemas 
      Height          =   375
      Left            =   3960
      TabIndex        =   5
      Top             =   2280
      Width           =   3135
   End
   Begin VB.TextBox txtidemas 
      Height          =   375
      Left            =   3960
      TabIndex        =   4
      Top             =   1080
      Width           =   3135
   End
   Begin VB.Label HARG 
      Caption         =   "HARGA PASARAN"
      Height          =   375
      Left            =   600
      TabIndex        =   11
      Top             =   2880
      Width           =   2535
   End
   Begin VB.Label Label4 
      Caption         =   "BERAT EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   2280
      Width           =   2535
   End
   Begin VB.Label Label3 
      Caption         =   "EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   2
      Top             =   1680
      Width           =   2535
   End
   Begin VB.Label Label2 
      Caption         =   "ID EMAS"
      Height          =   375
      Left            =   600
      TabIndex        =   1
      Top             =   1080
      Width           =   2535
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PERAMALAN DAN PENAPSIRAN PEGADAIAN"
      Height          =   495
      Left            =   2880
      TabIndex        =   0
      Top             =   240
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   -240
      Picture         =   "Form3.frx":0000
      Top             =   -1080
      Width           =   23400
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim peramalandanpenapsiranpegadaian As New ADODB.Recordset

Private Sub Comedit_Click()
koneksidb.Execute "update tabel_peramalandanpenapsiranpegadaian set idemas='" & txtidemas & "',emas='" & txtemas & "',beratemas='" & txtberatemas & "',hargapasaran='" & txthargapasaran & "'"
Call update
Call edit_grid
Call kosong
End Sub
Private Sub Comhapus_Click()
koneksidb.Execute "delete from tabel_peramalandanpenapsiranpegadaian where idemas='" & txtidemas & "'"
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
    If txthargapasaran = "" Then
    MsgBox "Harga Pasaran Kosong", vbExclamation, "pesan"
    txthargapasaran.SetFocus
    Exit Sub
    End If
Set peramalandanpenapsiranpegadaian = New ADODB.Recordset
peramalandanpenapsiranpegadaian.Open "select*from tabel_peramalandanpenapsiranpegadaian where idemas='" & txtidemas & "'", koneksidb
If Not peramalandanpenapsiranpegadaian.EOF Then
MsgBox "Id Emas sudah ada", vbCritical, "pesan"
txtidemas = ""
txtidemas.SetFocus
Exit Sub
Else
koneksidb.Execute "insert into tabel_peramalandanpenapsiranpegadaian(idemas,emas,beratemas,hargapasaran) value ('" & txtidemas & "','" & txtemas & "','" & txtberatemas & "','" & txthargapasaran & "')"
MsgBox "data tersimpan"
Call tampil_data
Set DataGrid1.DataSource = peramalandanpenapsiranpegadaian
With DataGrid1
End With
Call edit_grid
End If
End Sub

Private Sub DataGrid1_Click()
txtidemas.Text = peramalandanpenapsiranpegadaian!idemas
txtemas.Text = peramalandanpenapsiranpegadaian!emas
txtberatemas.Text = peramalandanpenapsiranpegadaian!beratemas
txthargapasaran.Text = peramalandanpenapsiranpegadaian!hargapasaran
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = peramalandanpenapsiranpegadaian
With peramalandanpenapsiranpegadaian
Call edit_grid
End With
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Id Emas"
    .Columns(1).Caption = "Emas"
    .Columns(2).Caption = "Berat Emas"
    .Columns(3).Caption = "Harga Pasaran"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
End With
End Sub

Sub tampil_data()
Set peramalandanpenapsiranpegadaian = New ADODB.Recordset
peramalandanpenapsiranpegadaian.ActiveConnection = koneksidb
peramalandanpenapsiranpegadaian.CursorLocation = adUseClient
peramalandanpenapsiranpegadaian.LockType = adLockOptimistic
peramalandanpenapsiranpegadaian.Source = "select * from tabel_peramalandanpenapsiranpegadaian"
peramalandanpenapsiranpegadaian.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = peramalandanpenapsiranpegadaian
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = peramalandanpenapsiranpegadaian
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txtidemas = ""
txtemas = ""
txtberatemas = ""
txthargapasaran = ""
txtidemas.SetFocus
End Sub

