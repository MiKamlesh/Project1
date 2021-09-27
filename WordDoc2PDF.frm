VERSION 5.00
Begin VB.Form WordToPDF 
   BackColor       =   &H00FFFFFF&
   Caption         =   "SHELL THIS TO CONVERT A WORD DOC TO A PDF"
   ClientHeight    =   4275
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6855
   Icon            =   "WordDoc2PDF.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4275
   ScaleWidth      =   6855
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   3360
      TabIndex        =   3
      Top             =   2760
      Width           =   1335
   End
   Begin VB.FileListBox File1 
      Appearance      =   0  'Flat
      Height          =   1785
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   2295
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "This is NOT one of their sample programs - Kevin Ritch wrote this from scratch"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   2040
      Width           =   6690
   End
   Begin VB.Image Image1 
      Height          =   660
      Left            =   360
      Picture         =   "WordDoc2PDF.frx":08CA
      Top             =   240
      Width           =   2025
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Only requires the FREE version of Win2PDF (The Printer Software that converts stuff to a PDF file)  Get yours at www.Win2PDF.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   360
      TabIndex        =   0
      Top             =   960
      Width           =   6135
   End
End
Attribute VB_Name = "WordToPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    
    Dim NotifyCallingAppThatProgramHasCompleted As String
    Dim CommaPosition As Integer
    Dim WordDocument As String
    Dim PDFDocument As String
    Dim StorePrinter As String
    Dim oApp As Word.Application
    Dim oDoc As Word.Document
    Dim StrInputName As String
    Dim StrOutputName As String
    Dim Dloop As Double
    Dim RequestStr As String
    
    On Error GoTo CloseWinWordObject:
    Me.Visible = False
    
    
    RequestStr = Command$
    
    File1.Path = App.Path & "\IN"
    File1.Refresh
    For Dloop = 0 To File1.ListCount - 1
        
        StrInputName = App.Path & "\IN\" & File1.List(Dloop)
        StrOutputName = App.Path & "\OUT\" & Replace(File1.List(Dloop), ".doc", ".pdf")
        
        WordDocument = Trim$(StrInputName)
        PDFDocument = Trim$(StrOutputName)
        
        Set oApp = CreateObject("Word.Application")
        oApp.Documents.Open (WordDocument)
        StorePrinter = oApp.ActivePrinter
        oApp.ActivePrinter = "Win2PDF"
        
        SaveSetting "Dane Prairie Systems", "Win2PDF", "PDFFileName", PDFDocument
        oApp.PrintOut
        oApp.ActivePrinter = StorePrinter
CloseWinWordObject:
        oApp.ActiveDocument.Close
        oApp.Quit
        Set oApp = Nothing
        
'        Whoops:
'        MsgBox "Sorry, you forgot to submit kosher filenames in your COMMAND line!" & String$(2, 10) & "SHELL SYNTAX: WordDoc2PDF.exe SourceDoc,TargetPDF", vbApplicationModal + vbExclamation, "Whoops! - SHELL SYNTAX ERROR"
'        GoTo CloseWinWordObject
    Next
    
    On Error Resume Next
'    CloseWinWordObject:
'    oApp.Quit
'    Set oApp = Nothing
    
'    MkDir "c:\Program Files\V8Software.com"
    Open "c:\Program Files\V8Software.com\AppCompleted.dat" For Binary Shared As #1
    NotifyCallingAppThatProgramHasCompleted = "TRUE"
    Put #1, 1, NotifyCallingAppThatProgramHasCompleted
'    End
    
    Command1.Value = True
End Sub

Private Sub Form_Load()
    Me.Visible = False
    Command1.Value = True
End Sub
