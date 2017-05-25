Attribute VB_Name = "Module11"
Sub CSV2XML()
'
' CSV2XML Macro
' Converteer OUTLOOK contacten CSV naar XML voor Fritzbox
'
'
' Variabelen
'
    Dim Directory As String
    Dim OutloookCSV As String
    Dim OutlookCSVcompl As String
    Dim Nrofrows As Integer
    Dim Today As Date
    Dim Converter As String
    Dim Jaar As String
    Dim Maand As String
    Dim Dag As String
    Dim XMLfilename As String
    Dim XLSXfilename As String
    Dim Rij As Integer
    Dim XMLrecord As String
    Dim LastRow As Long
    Dim Rng As String
    Dim Teller As Integer
    Dim Answer As String
    Dim MyNote As String
    
       
'
' Zet prompts uit
'

    Application.DisplayAlerts = False

'
' Zet de waarden die vast zijn
'
    Directory = "D:\Data\Sync Gedeeld\Agenda & Mail\Outlook back-up\"
    OutlookCSV = "OUTLOOK.CSV"
    Converter = "Converter.xlsm"
'
' Lees CSV file en plak deze in het tabblad contacts
'
    ChDir Directory
    OutlookCSVcompl = Directory + OutlookCSV
    Workbooks.Open Filename:=OutlookCSVcompl
        
      
'
' Kopieer naar tabblad "Contacts"
'
    Windows(OutlookCSV).Activate
    Columns("A:I").Select
    Application.CutCopyMode = False
    Selection.Copy
    Windows(Converter).Activate
    Sheets("Contacts").Select
    Columns("A:I").Select
    ActiveSheet.Paste
    
'
' Bepaal aantal rijen
'
    LastRow = ActiveSheet.UsedRange.Rows(ActiveSheet.UsedRange.Rows.Count).Row
    
'
' Sluit input CSV
'
    Windows(OutlookCSV).Activate
    ActiveWorkbook.Close
    Application.WindowState = xlNormal
            
'
' Kolombreedte goed zetten
'
    Columns("A:I").Select
    Selection.Columns.AutoFit
    
'
' Vervang de 5 XML special characters door XML keywords
'
    
     Selection.Replace What:="&", Replacement:="&amp;", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="'", Replacement:="&apos;", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="""", Replacement:="&quot;", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:=">", Replacement:="&gt;", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Selection.Replace What:="<", Replacement:="&lt;", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
        
'
' Vul XML tabblad met juiste aantal regels
'
    Sheets("XML").Select
    Rows("2:2").Select
    Selection.Copy
    Rng = "3:" + CStr(LastRow)
    Rows(Rng).Select
    ActiveSheet.Paste
    Range("B1").Select
    Application.CutCopyMode = False
    Selection.Cut
    Rng = "A" + CStr(LastRow + 1)
    Range(Rng).Select
    ActiveSheet.Paste
    
'
' Creëer dataset name voor XML file en XLSX file
'

Jaar = CStr(Year(Date))

If Month(Date) > 9 Then
Maand = Month(Date)
Else
Maand = CStr(0) + CStr(Month(Date))
End If


If Day(Date) > 9 Then
Dag = Day(Date)
Else
Dag = CStr(0) + CStr(Day(Date))
End If

XLSXfilename = "Genereer XML " + Jaar + Maand + Dag + ".xlsx"
XMLfilename = "Upload XML " + Jaar + Maand + Dag + ".txt"

'
' Save de XLSX file met de XML erin (for debugging purposes only)
'

ActiveWorkbook.SaveAs Filename:=Directory + XLSXfilename, FileFormat:=51

'
' Schrijf XML weg
'
  
Worksheets("XML").Activate

' Create Stream object
   Set fsT = CreateObject("ADODB.Stream")

 
' Specify stream type - we want To save text/string data.
  fsT.Type = 2
' Specify charset For the source text data.
  fsT.Charset = "utf-8"
' Open the stream
  fsT.Open
  
Teller = 0
Rij = 1
For Rij = 1 To LastRow + 1

    XMLrecord = Cells(Rij, 1).Text + Cells(Rij, 2).Text + Cells(Rij, 3).Text
    
    If XMLrecord <> "" Then
        Teller = Teller + 1
    End If
    
    fsT.writetext XMLrecord

Next Rij

 fsT.SaveToFile Directory + XMLfilename, 2
 
'
' Maximize window and set cursor om 1,1
'
Application.WindowState = xlMaximized
Range("A1").Select
'
' Display aantal weggeschreven contacten
'
Application.DisplayAlerts = True

    MyNote = CStr(Teller - 2) + " Contacten in XML formaat weggeschreven naar " + Chr(10) + Chr(13) + Directory + XMLfilename + Chr(10) + Chr(13) + Chr(10) + Chr(13) + "Wilt u direct door naar Fritzbox??"
    Answer = MsgBox(MyNote, vbQuestion + vbYesNo, "XML Generated")
 
    If Answer = vbYes Then
'
' Goto FritzBox
'
    ActiveWorkbook.FollowHyperlink Address:="http://192.168.178.1", NewWindow:=True
    
    Else
    
    MsgBox ("You can upload your XML later on")

    End If
 
'
' Quit
'
MsgBox ("XML generation completed")

ActiveWorkbook.Close

Application.Quit



End Sub
       
