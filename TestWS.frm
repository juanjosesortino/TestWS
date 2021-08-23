VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form1 
   Caption         =   "Test AFIP TLS"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11895
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   11895
   StartUpPosition =   3  'Windows Default
   Begin MSComDlg.CommonDialog cdg1 
      Left            =   5700
      Top             =   2490
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command3 
      Caption         =   "..."
      Height          =   315
      Left            =   11190
      TabIndex        =   18
      Top             =   2520
      Width           =   525
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2370
      TabIndex        =   17
      Text            =   "D:\Algoritmo\Factura Electronica\LoginTicketResponse.xml"
      Top             =   2520
      Width           =   8745
   End
   Begin VB.Frame Frame3 
      Height          =   2445
      Left            =   0
      TabIndex        =   11
      Top             =   30
      Width           =   11865
      Begin VB.Frame Frame1 
         Caption         =   "Ambiente"
         Height          =   1125
         Left            =   210
         TabIndex        =   13
         Top             =   1020
         Width           =   1545
         Begin VB.OptionButton Option1 
            Caption         =   "Homologacion"
            Height          =   525
            Left            =   150
            TabIndex        =   15
            Top             =   180
            Width           =   1335
         End
         Begin VB.OptionButton Option2 
            Caption         =   "Produccion"
            Height          =   525
            Left            =   150
            TabIndex        =   14
            Top             =   540
            Width           =   1245
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Test Dummy"
         Height          =   525
         Left            =   180
         TabIndex        =   12
         Top             =   300
         Width           =   1575
      End
      Begin VB.Label Label6 
         BorderStyle     =   1  'Fixed Single
         Height          =   1875
         Left            =   2220
         TabIndex        =   20
         Top             =   270
         Width           =   9525
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "RequestSecureProtocols"
      Height          =   1485
      Left            =   120
      TabIndex        =   7
      Top             =   3990
      Width           =   2055
      Begin VB.OptionButton Option5 
         Caption         =   "TLS1.2"
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   900
         Width           =   1035
      End
      Begin VB.OptionButton Option4 
         Caption         =   "TLS1.1"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1335
      End
      Begin VB.OptionButton Option3 
         Caption         =   "TLS1"
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   1005
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Test Alicuotas"
      Height          =   525
      Left            =   30
      TabIndex        =   6
      Top             =   2910
      Width           =   1575
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   2370
      TabIndex        =   5
      Top             =   3600
      Width           =   9315
   End
   Begin VB.TextBox Text3 
      Height          =   285
      Left            =   2370
      TabIndex        =   3
      Top             =   3240
      Width           =   9315
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   2370
      TabIndex        =   0
      Top             =   2880
      Width           =   9315
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  'Fixed Single
      Height          =   1875
      Left            =   2250
      TabIndex        =   19
      Top             =   3960
      Width           =   9585
   End
   Begin VB.Label Label4 
      Caption         =   "LoginTicketResponse.xml"
      Height          =   255
      Left            =   60
      TabIndex        =   16
      Top             =   2580
      Width           =   2235
   End
   Begin VB.Label Label3 
      Caption         =   "<Cuit>"
      Height          =   255
      Left            =   1650
      TabIndex        =   4
      Top             =   3630
      Width           =   645
   End
   Begin VB.Label Label2 
      Caption         =   "<sign>"
      Height          =   255
      Left            =   1650
      TabIndex        =   2
      Top             =   3270
      Width           =   645
   End
   Begin VB.Label Label1 
      Caption         =   "<token>"
      Height          =   255
      Left            =   1650
      TabIndex        =   1
      Top             =   2910
      Width           =   645
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private strLoginTicketResponse As String

Private fs                             As Object

Private strToken As String
Private strSign  As String
Private strCuit  As String

Enum WinHttpRequestSecureProtocols
   SecureProtocol_SSL2 = 8
   SecureProtocol_SSL3 = 32
   SecureProtocol_TLS = 128
   SecureProtocol_TLS1 = 512
   SecureProtocol_TLS12 = 2048
   SecureProtocol_ALL = 168
End Enum

Private Sub Command1_Click()
Dim strAmbiente As String
Dim WinOption_Secure As Variant

'configurado a opción segura
   WinOption_Secure = 9

   Set WinHttpReq = New WinHttpRequest
  
   WinHttpReq.SetTimeouts 0, 60000, 60000, 60000
   
   If Option1.Value = True Then
      WinHttpReq.Open "POST", "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?op=FEDummy", False
   Else
      WinHttpReq.Open "POST", "https://servicios1.afip.gov.ar/wsfev1/service.asmx?op=FEDummy", False
   End If
   
   WinHttpReq.SetRequestHeader "Content-Type", "text/xml"
   WinHttpReq.SetRequestHeader "SOAPAction", "http://ar.gov.afip.dif.FEV1/FEDummy"

   WinHttpReq.Option(WinOption_Secure) = SecureProtocol_TLS12

   ' Send the HTTP Request.
   WinHttpReq.Send GenerarFEDummy
   
   Label6.Caption = WinHttpReq.ResponseText
End Sub
Private Function GenerarFEDummy() As String

       GenerarFEDummy = _
               "<?xml version=""1.0"" encoding=""utf-8""?>" & _
               "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
                 "<soap:Body>" & _
                   "<FEDummy xmlns=""http://ar.gov.afip.dif.facturaelectronica/"" />" & _
                 "</soap:Body>" & _
               "</soap:Envelope>"
               
End Function

Private Sub Command3_Click()
   On Error Resume Next
   
   cdg1.InitDir = "D:\Algoritmo\Factura Electronica"
   cdg1.DialogTitle = "Busqueda de LoginTicketResponse.xml"
   cdg1.CancelError = True
   cdg1.Filter = "Documentos *.xml|*.xml"
   cdg1.FilterIndex = 1
   cdg1.Flags = cdlOFNHideReadOnly + cdlOFNExtensionDifferent + cdlOFNOverwritePrompt
   cdg1.DefaultExt = "LoginTicketResponse.xml"
   cdg1.ShowOpen
   
   If cdg1.FileName <> "" Then
      Text6.Text = cdg1.FileName
      CompletarDatos
   End If
End Sub

Private Sub Form_Load()
   Option1.Value = True
   Option5.Value = True
   
   
   CompletarDatos
End Sub

Private Sub CompletarDatos()
   Set fs = CreateObject("Scripting.FileSystemObject")
   
   If fs.fileexists(Text6.Text) Then
      Set objReadFile = fs.OpenTextFile(Text6.Text, 1)
      strLoginTicketResponse = objReadFile.ReadAll
      objReadFile.Close
      Set objReadFile = Nothing
   End If
   On Error Resume Next
   Text2.Text = Mid(strLoginTicketResponse, InStr(strLoginTicketResponse, "&lt;token&gt;") + 13, InStr(strLoginTicketResponse, "&lt;/token&gt;") - InStr(strLoginTicketResponse, "&lt;token&gt;") - 13)
   Text3.Text = Mid(strLoginTicketResponse, InStr(strLoginTicketResponse, "&lt;sign&gt;") + 12, InStr(strLoginTicketResponse, "&lt;/sign&gt;") - InStr(strLoginTicketResponse, "&lt;sign&gt;") - 12)
   
   Text4.Text = Mid(strLoginTicketResponse, InStr(strLoginTicketResponse, "&lt;destination&gt") + 18, InStr(strLoginTicketResponse, "&lt;/destination&gt;") - InStr(strLoginTicketResponse, "&lt;destination&gt") - 18)
   Text4.Text = Mid(Text4.Text, InStr(Text4.Text, "CUIT ") + 5, InStr(Text4.Text, ", CN") - InStr(Text4.Text, "CUIT ") - 5)

End Sub

Private Sub Command2_Click()
Dim strAmbiente As String
Dim WinOption_Secure As Variant

   On Error GoTo ErrorHandler
   
   Label5.Caption = ""
   
   'configurado a opción segura
   WinOption_Secure = 9

   Set WinHttpReq = New WinHttpRequest
  
   WinHttpReq.SetTimeouts 0, 60000, 60000, 60000
   
   If Option1.Value = True Then
      WinHttpReq.Open "POST", "https://wswhomo.afip.gov.ar/wsfev1/service.asmx?op=FEParamGetTiposIva", False
   Else
      WinHttpReq.Open "POST", "https://servicios1.afip.gov.ar/wsfev1/service.asmx?op=FEParamGetTiposIva", False
   End If
   
   WinHttpReq.SetRequestHeader "Content-Type", "text/xml"
   WinHttpReq.SetRequestHeader "SOAPAction", "http://ar.gov.afip.dif.FEV1/FEParamGetTiposIva"

   If Option3.Value = True Then
      WinHttpReq.Option(WinOption_Secure) = SecureProtocol_TLS
   Else
      If Option4.Value = True Then
         WinHttpReq.Option(WinOption_Secure) = SecureProtocol_TLS1
      Else
         WinHttpReq.Option(WinOption_Secure) = SecureProtocol_TLS12
      End If
   End If
    
   ' Send the HTTP Request.
   WinHttpReq.Send FEParamGetTiposIva
   
   Label5.Caption = WinHttpReq.ResponseText
   
   Exit Sub

ErrorHandler:
   MsgBox Err.Description
End Sub


Private Function FEParamGetTiposIva() As String

   FEParamGetTiposIva = _
               "<?xml version=""1.0"" encoding=""utf-8:""?>" & _
               "<soap:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap=""http://schemas.xmlsoap.org/soap/envelope/"">" & _
               "<soap:Body> <FEParamGetTiposIva xmlns=""http://ar.gov.afip.dif.FEV1/"">" & _
                  "<Auth>" & _
                     "<Token>" & Text2.Text & "</Token>" & _
                     "<Sign>" & Text3.Text & "</Sign>" & _
                     "<Cuit>" & Text4.Text & "</Cuit>" & _
                  "</Auth>" & _
               "</FEParamGetTiposIva>" & _
               "</soap:Body></soap:Envelope>"
End Function


