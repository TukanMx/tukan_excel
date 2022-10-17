![Image text](https://tukanmx.com/static/media/tukan_logo_trim.fa49c6d7.png)

## Table of Contents
1. [General Info](#general-info)
2. [Technologies](#technologies)
3. [Installation](#installation)
4. [Coding](#Coding)
5. [Files](#Files)
### General Info
***
In this repository where you can find the VBA code to Query to Tukan with Excel.

## Technologies
***
* Excel
* TUKAN APP
## Installation
***
### File Download

Use this code alongside the TUKAN app to get TUKAN queries into your excel. Your Query Id should be in the C3 cell, Token on the C2 and Language on C4. With that you can execute the Macro to get the data.

The easiest way to get the VBA code working is downloading the macro file in this folder enabling permissions and simply filling out the cells. Alternatively you can copy the VBA code and write the code yourself.

Remember to get the References enabled on your excel to have the code execute correctly.


<center>
<img src = "./img/Excel_reference_requirements.png" width = "80%"/>
</center>

### VBA CODE

```
Public Sub ReadWeb()
 Dim Sheet As Worksheet
 Set Sheet = Application.ActiveSheet
 Dim Link As String, QueryId As String, Token As String, Language As String
 QueryId = Sheet.Range("C3").Value
 Token = Sheet.Range("C2").Value
 Language = Sheet.Range("C4").Value
 Link = "http://client.tukanmx.com/visualizations/retrieve_query_csv/" & Language & "/" & QueryId & "/" & Token & "/"
    Dim html As HTMLDocument
    Set html = New HTMLDocument
    With CreateObject("MSXML2.XMLHTTP")
      .Open "GET", Link, False
      .send
      ReadPSV (.responseText)
    End With
End Sub
Private Sub ReadPSV(PSV As String)
Dim Rows() As String
Rows = Split(PSV, vbLf)
Dim Sht As Worksheet
Set Sht = Application.ActiveSheet
Dim Cell As Range
Set Cell = Sht.Range("A7")
Dim Titles() As String

Dim i As Integer
Dim j As Integer
For j = 0 To UBound(Rows)
    Titles = Split(Rows(j), "|")
For i = 0 To UBound(Titles, 1)
 Cell.Offset(j, i).Value = Titles(i)
Next
Next
End Sub

```

