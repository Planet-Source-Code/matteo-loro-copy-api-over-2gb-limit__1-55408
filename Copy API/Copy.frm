VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Copy File (Over 2 GB)"
   ClientHeight    =   2460
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2460
   ScaleWidth      =   6540
   StartUpPosition =   2  'CenterScreen
   Begin MSComDlg.CommonDialog CmnDlg 
      Left            =   6000
      Top             =   1920
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   6255
      Begin VB.CommandButton CmdBrowse 
         Caption         =   "Apri"
         Height          =   255
         Left            =   5400
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.TextBox TxtSrc 
         Height          =   285
         Left            =   1200
         TabIndex        =   0
         Top             =   240
         Width           =   4095
      End
      Begin VB.TextBox TxtDest 
         Height          =   285
         Left            =   1200
         TabIndex        =   2
         Text            =   "c:\temp.tmp"
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton CmdCopy 
         Caption         =   "Copia"
         Default         =   -1  'True
         Height          =   615
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   6015
      End
      Begin MSComctlLib.ProgressBar ProgressBar1 
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   1680
         Width           =   5025
         _ExtentX        =   8864
         _ExtentY        =   556
         _Version        =   393216
         Appearance      =   0
         Scrolling       =   1
      End
      Begin VB.Shape Shape1 
         Height          =   375
         Left            =   150
         Top             =   1650
         Width           =   5085
      End
      Begin VB.Label Label3 
         Caption         =   "%"
         Height          =   255
         Left            =   5880
         TabIndex        =   9
         Top             =   1755
         Width           =   255
      End
      Begin VB.Label LblPercent 
         Alignment       =   2  'Center
         Caption         =   "0"
         Height          =   255
         Left            =   5280
         TabIndex        =   8
         Top             =   1755
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Sorgente:"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Destinazione:"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   975
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private F As Random
Private FDest As Random

Private Sub CmdBrowse_Click()

  'Imposto il filtro a tutti i file
  CmnDlg.Filter = "Tutti i file|*.*"
  'Mostro la finestra di apertura file
  CmnDlg.ShowOpen
  TxtSrc.Text = CmnDlg.FileName

End Sub

Private Sub CmdCopy_Click()

 Dim SrcFileLen As Double
 Dim rest As String
 Dim BytesToGet As Integer
 Dim BytesCopied As Double
 Dim Chunk() As Byte
 Dim srcFile As String
 Dim DestFile As String

 Dim fs As New FileSystemObject
 Dim fl As File

  srcFile = TxtSrc
  DestFile = TxtDest

  Set fl = fs.GetFile(srcFile)
  'Ricavo le dimensioni del file
  SrcFileLen = fl.Size

  'Li svuoto altrimenti va in conflitto con le funzioni della classe Random
  Set fl = Nothing
  Set fs = Nothing

  'Inizializzazione ProgressBar
  ProgressBar1.Min = 0
  ProgressBar1.Max = SrcFileLen
  LblPercent.Caption = "0"
  ProgressBar1.Value = 0

  'Byte da copiare ad ogni ciclo (4kb)
  BytesToGet = 4096
  BytesCopied = 0

  'Apro il file da copiare e quello di destinazione
  F.OpenFileRead TxtSrc.Text
  FDest.OpenFile TxtDest.Text

  Do While BytesCopied < SrcFileLen
    'Controllo quanto manca alla fine
    If BytesToGet < (SrcFileLen - BytesCopied) Then
      'Leggo 4 KBytes
      rest = Space(BytesToGet)
      Chunk = F.ReadBytes(Len(rest))
     Else 'NOT BYTESTOGET...
      'Leggo i Bytes rimanenti
      rest = Space(SrcFileLen - BytesCopied)
      Chunk = F.ReadBytes(Len(rest))
    End If
    'Aggiorno la posizione del file raggiunta
    BytesCopied = BytesCopied + Len(rest)

    'Aggiorno ProgressBar
    ProgressBar1.Value = BytesCopied
    'Mostro Percentuale
    LblPercent.Caption = Int(BytesCopied / SrcFileLen * 100)
    LblPercent.Refresh

    'Scrivo i dati letti nel file di destinazione
    FDest.WriteBytes Chunk
    DoEvents
  Loop

  'Chiudo i due file
  F.CloseFile
  FDest.CloseFile

  'Avviso
  MsgBox "Copia terminata con successo!"

End Sub

Private Sub Form_Load()

  'Inizializzo le variabili
  Set F = New Random
  Set FDest = New Random

End Sub

Private Sub Form_Unload(Cancel As Integer)

  'Svuoto la memoria
  Set F = Nothing
  Set FDest = Nothing

End Sub

':) Ulli's VB Code Formatter V2.16.6 (2004-ago-06 09:43) 4 + 100 = 104 Lines
