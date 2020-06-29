VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form Form4 
   Caption         =   "Form4"
   ClientHeight    =   7560
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12525
   LinkTopic       =   "Form4"
   ScaleHeight     =   7560
   ScaleWidth      =   12525
   StartUpPosition =   3  'Windows Default
   Begin VB.ComboBox Combulansekarang 
      Height          =   315
      Left            =   4320
      TabIndex        =   13
      Text            =   "Combo1"
      Top             =   2640
      Width           =   3375
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   2535
      Left            =   360
      TabIndex        =   12
      Top             =   4920
      Width           =   9495
      _ExtentX        =   16748
      _ExtentY        =   4471
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
      TabIndex        =   11
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Comhapus 
      Caption         =   "HAPUS"
      Height          =   495
      Left            =   4560
      TabIndex        =   10
      Top             =   4200
      Width           =   1455
   End
   Begin VB.CommandButton Comedit 
      Caption         =   "TRANSFER"
      Height          =   495
      Left            =   2280
      TabIndex        =   9
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Comproses 
      Caption         =   "PROSES"
      Height          =   495
      Left            =   240
      TabIndex        =   8
      Top             =   4200
      Width           =   1455
   End
   Begin VB.TextBox txthasilprediksi 
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   3360
      Width           =   3375
   End
   Begin VB.TextBox txthargapasaran 
      Height          =   495
      Left            =   4320
      TabIndex        =   6
      Top             =   1920
      Width           =   3375
   End
   Begin VB.TextBox txthargaemaspegadaian 
      Height          =   495
      Left            =   4320
      TabIndex        =   5
      Top             =   1200
      Width           =   3375
   End
   Begin VB.Label Label5 
      Caption         =   "HASIL PREDIKSI"
      Height          =   495
      Left            =   480
      TabIndex        =   4
      Top             =   3360
      Width           =   2895
   End
   Begin VB.Label Label4 
      Caption         =   "BULAN SEKARANG"
      Height          =   495
      Left            =   480
      TabIndex        =   3
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label3 
      Caption         =   "HARGA PASARAN"
      Height          =   495
      Left            =   480
      TabIndex        =   2
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label2 
      Caption         =   "HARGA EMAS PEGADAIAN"
      Height          =   495
      Left            =   480
      TabIndex        =   1
      Top             =   1200
      Width           =   2895
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "PREDIKSI HARGA EMAS PEGADAIAN"
      Height          =   495
      Left            =   3480
      TabIndex        =   0
      Top             =   240
      Width           =   5295
   End
   Begin VB.Image Image1 
      Height          =   10800
      Left            =   0
      Picture         =   "Form4.frx":0000
      Top             =   -1080
      Width           =   23400
   End
End
Attribute VB_Name = "Form4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim prediksihargaemaspegadaian As New ADODB.Recordset
Private Sub Comproses_Click()
txthasilprediksi.Text = Val(txthargaemaspegadaian.Text) + Val(txthargapasaran.Text) / Val(txtbulansekarang.Text)
Call PROSES
End Sub

Private Sub Comedit_Click()
koneksidb.Execute "update tabel_prediksihargaemaspegadaian set hargaemaspegadaian='" & txthargaemaspegadaian.Text & "',hargapasaran='" & txthargapasaran.Text & "',bulansekarang='" & Combulansekarang.Text & "',hasilprediksi='" & txthasilprediksi.Text & "'"
Call update
Call edit_grid
Call kosong
End Sub

Private Sub Comhapus_Click()
koneksidb.Execute "delete from tabel_prediksihargaemaspegadaian where hargaemaspegadaian='" & txthargaemaspegadaian & "'"
Call refreshh
Call kosong
txthargaemaspegadaian.SetFocus
End Sub

Private Sub Comkeluar_Click()
x = MsgBox("Yakin Keluar?", vbQuestion + vbYesNo, "informasi")
If x = vbYes Then End
End Sub


Private Sub DataGrid1_Click()
txthargaemaspegadaian.Text = prediksihargaemaspegadaian!hargaemaspegadaian
txthargapasaran.Text = prediksihargaemaspegadaian!hargapasaran
Combulansekarang.Text = prediksihargaemaspegadaian!bulansekarang
txthasilprediksi.Text = prediksihargaemaspegadaian!hasilprediksi
End Sub

Private Sub Form_Load()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = prediksihargaemaspegadaian
With prediksihargaemaspegadaian
With Combulansekarang
    .AddItem " 1 "
    .AddItem " 2 "
    .AddItem " 3 "
    .AddItem " 4 "
    .AddItem " 5 "
    .AddItem " 6 "
    .AddItem " 7 "
    .AddItem " 8 "
    .AddItem " 9 "
    .AddItem " 10 "
    .AddItem " 11 "
    .AddItem " 12 "
End With
Call edit_grid
End With
End Sub

Sub edit_grid()
With DataGrid1
    .Columns(0).Caption = "Harga Emas Pegadaian"
    .Columns(1).Caption = "Harga Pasaran"
    .Columns(2).Caption = "Bulan Sekrang"
    .Columns(3).Caption = "Hasil Prediksi"
    .Columns(0).Width = 1200
    .Columns(1).Width = 1200
    .Columns(2).Width = 1200
    .Columns(3).Width = 1200
End With
End Sub

Sub tampil_data()
Set prediksihargaemaspegadaian = New ADODB.Recordset
prediksihargaemaspegadaian.ActiveConnection = koneksidb
prediksihargaemaspegadaian.CursorLocation = adUseClient
prediksihargaemaspegadaian.LockType = adLockOptimistic
prediksihargaemaspegadaian.Source = "select * from tabel_prediksihargaemaspegadaian"
prediksihargaemaspegadaian.Open
End Sub

Sub update()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = prediksihargaemaspegadaian
With DataGrid1
End With
End Sub

Sub refreshh()
Call bukadb
Call tampil_data
Set DataGrid1.DataSource = prediksihargaemaspegadaian
With DataGrid1
End With
Call edit_grid
End Sub

Sub kosong()
txthargaemaspegadaian = ""
txthargapasaran = ""
Combulansekarang = ""
txthasilprediksi = ""
txthargaemaspegadaian.SetFocus
End Sub

