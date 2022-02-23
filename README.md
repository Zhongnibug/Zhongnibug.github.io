## There is Zhongnibug's Personal Blog
```
Sub 合同系列word生成()
'
' 合同系列word生成 Macro
' 生成合同的申请审批表，询价函会签单，询价函等一系列word文档
'

Dim nowPath As String
nowPath = ActiveWorkbook.Path


Dim titleToName As String
titleToName = ActiveSheet.Cells(2, "c").Value
Dim bidtype As String
bidtype = ActiveSheet.Cells(7, "c").Value
Dim preamount As String
preamount = ActiveSheet.Cells(7, "e").Value
Dim reason As String
reason = ActiveSheet.Cells(6, "c").Value
Dim work As String
work = ActiveSheet.Cells(3, "c").Value
Dim company1 As String
company1 = ActiveSheet.Cells(14, "c").Value
Dim company2 As String
company2 = ActiveSheet.Cells(15, "c").Value
Dim company3 As String
company3 = ActiveSheet.Cells(16, "c").Value
Dim price1 As String
price1 = ActiveSheet.Cells(14, "e").Value
Dim price2 As String
price2 = ActiveSheet.Cells(15, "e").Value
Dim price3 As String
price3 = ActiveSheet.Cells(16, "e").Value
Dim requirement As String
requirement = ActiveSheet.Cells(4, "c").Value
Dim requirement2 As String
requirement2 = ActiveSheet.Cells(12, "e").Value

Dim deadlinesum As String
deadlinesum = ActiveSheet.Cells(19, "c").Value

deadlinesplit = Split(deadlinesum, "-")
Dim dealineyear As String
Dim dealine As String
dealineyear = deadlinesplit(0)
dealine = deadlinesplit(1) & "月" & deadlinesplit(2)

Dim asktimesum As String
asktimesum = ActiveSheet.Cells(18, "c").Value
asktimesumsplit = Split(asktimesum, "-")
Dim askyear As String
Dim asktime As String
askyear = asktimesumsplit(0)
asktime = asktimesumsplit(1) & "月" & asktimesumsplit(2)
Dim znaskyear As String
Dim znasktime As String

znaskyear = intToZnint(Mid(asktimesumsplit(0), 1, 1)) & intToZnint(Mid(asktimesumsplit(0), 2, 1)) & intToZnint(Mid(asktimesumsplit(0), 3, 1)) & intToZnint(Mid(asktimesumsplit(0), 4, 1))

znasktime = intToZnint(asktimesumsplit(1)) & "月" & intToZnint(asktimesumsplit(2))

Dim meetingsum As String
meetingsum = ActiveSheet.Cells(20, "c").Value

meetingsumsplit = Split(meetingsum, "-")
Dim meetingyear As String
Dim meetingtime As String
meetingyear = meetingsumsplit(0)
meetingtime = meetingsumsplit(1) & "月" & meetingsumsplit(2)

Dim asignsum As String
asignsum = ActiveSheet.Cells(21, "c").Value

asignsumsplit = Split(asignsum, "-")
Dim asignyear As String
asignyear = asignsumsplit(0)

Dim winnum As Integer
winnum = ActiveSheet.Cells(17, "c").Value

Dim wincompany As String
wincompany = ActiveSheet.Cells(13 + winnum, "c").Value
Dim winprice As String
winprice = ActiveSheet.Cells(13 + winnum, "e").Value

Dim bidoption As String
bidoption = ActiveSheet.Cells(8, "e").Value

Dim title As String
If InStr(1, titleToName, Chr(10), vbBinaryCompare) > 0 Then
    titleToName = Application.WorksheetFunction.Clean(titleToName)
End If
title = titleToName
If InStr(1, titleToName, Chr(34), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(34), "")
End If
If InStr(1, titleToName, Chr(42), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(42), "")
End If
If InStr(1, titleToName, Chr(47), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(47), "")
End If
If InStr(1, titleToName, Chr(58), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(58), "")
End If
If InStr(1, titleToName, Chr(60), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(60), "")
End If
If InStr(1, titleToName, Chr(62), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(62), "")
End If
If InStr(1, titleToName, Chr(63), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(63), "")
End If
If InStr(1, titleToName, Chr(92), vbBinaryCompare) > 0 Then
    titleToName = Replace(titleToName, Chr(92), "")
End If

Dim signsum As String
signsum = ActiveSheet.Cells(21, "c").Value
signsplit = Split(signsum, "-")
Dim signyear As String
Dim signtime As String
signyear = signsplit(0)
signtime = signsplit(1) & "月" & signsplit(2) & "日"

If titleToName = "零星材料采购" Then
titleToName = "零星材料采购-" & signtime
End If

Dim host As String
host = ActiveSheet.Cells(23, "c").Value

Dim participants As String
participants = ActiveSheet.Cells(23, "e").Value

Dim haveSecurity As Boolean
If InStr(1, participants, "武罗佳") > 0 Then
haveSecurity = True
Else
haveSecurity = False
End If

Dim haveReq As String
haveReq = ActiveSheet.Cells(11, "c").Value

Dim worktime As String
worktime = ActiveSheet.Cells(9, "e").Value
Dim worktimereason As String
worktimereason = ActiveSheet.Cells(10, "c").Value
Dim payreason As String
payreason = ActiveSheet.Cells(10, "e").Value
Dim qualification As String
qualification = ActiveSheet.Cells(13, "c").Value
Dim deadlinehm As String
deadlinehm = ActiveSheet.Cells(19, "e").Value
Dim tecconsulting As String
tecconsulting = ActiveSheet.Cells(22, "c").Value
Dim meetingtimeap As String
meetingtimeap = ActiveSheet.Cells(20, "e").Value
Dim infoType As String
infoType = ActiveSheet.Cells(11, "e").Value
Dim goodType As String
goodType = ActiveSheet.Cells(4, "e").Value
Dim makeUpType As String
makeUpType = ActiveSheet.Cells(3, "e").Value
Dim anotherComan As String
anotherComan = ActiveSheet.Cells(21, "e").Value
Dim meetreason As String
meetreason = ActiveSheet.Cells(13, "e").Value
Dim deadprice As String
deadprice = ActiveSheet.Cells(6, "e").Value
Dim projtype As String
projtype = ActiveSheet.Cells(5, "e").Value
Dim amounttype As String
amounttype = ActiveSheet.Cells(22, "e").Value

Dim idate As Date
idate = Format(Now, "yyyy-m-d")





'创文件夹
On Error Resume Next
VBA.MkDir (nowPath & "\工作目录")
VBA.MkDir (nowPath & "\工作目录" & "\" & idate & "-" & titleToName)
On Error GoTo 0


Dim wrd As Word.Application
Set wrd = CreateObject("word.Application")
Application.ScreenUpdating = False

Dim titlelen As Integer
titlelen = Len(title)
Dim worktimelen As Integer
worktimelen = Len(worktimereason)
'1申请审批表---j、x

wrdPath = nowPath & "\" & bidoption & "\1申请审批表.docx"
With wrd
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 16 Then
.Selection = title & Chr(10)
Else
.Selection = title
End If
.Selection.Goto Name:="name"
.Selection = title
.Selection.Goto Name:="type"
.Selection = bidtype
.Selection.Goto Name:="preamount"
.Selection = preamount
.Selection.Goto Name:="reason"
.Selection = reason
.Selection.Goto Name:="work"
.Selection = work

Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\1-申请审批表.docx"
Application.DisplayAlerts = True
wrd.Documents.Close
End With
'2,3询价函---j/x


If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\2询价函.docx"
Else
wrdPath = nowPath & "\" & bidoption & "\3询价函.docx"
End If


''company1
With wrd
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen < 18 Then
.Selection = title & Chr(10)
ElseIf titlelen < 23 Then
.Selection = Chr(10) & title & Chr(10)
Else
.Selection = Chr(10) & title
End If
.Selection.Goto Name:="company"
.Selection = company1
.Selection.Goto Name:="theme"
.Selection = title
If bidoption <> "竞争性比选" Then
.Selection.Goto Name:="work"
.Selection = work
End If
If bidoption = "竞争性比选" Then
.Selection.Goto Name:="theme2"
.Selection = title
Else
If haveReq = "有" Then
.Selection.Goto Name:="requirement"
.Selection = "详见《鸡冠石污水处理厂" & requirement & "要求》"
Else
.Selection.Goto Name:="requirement"
.Selection = requirement2
End If
End If
.Selection.Goto Name:="worktime"
.Selection = worktime
If worktimelen > 45 And bidoption = "竞争性比选" Then
.Selection.Goto Name:="worktimemakeup"
.Selection = "详见《鸡冠石污水处理厂" & title & "竞争性比选》"
Else
.Selection.Goto Name:="worktimereason"
.Selection = worktimereason
End If
.Selection.Goto Name:="price"
.Selection = projtype
.Selection.Goto Name:="payreason"
.Selection = payreason
.Selection.Goto Name:="qualification1"
.Selection = qualification
.Selection.Goto Name:="deadlineyear"
.Selection = dealineyear
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="askyear2"
.Selection = askyear
.Selection.Goto Name:="asktime2"
.Selection = asktime
.Selection.Goto Name:="deadprice"
.Selection = deadprice
.Selection.Goto Name:="amounttype"
.Selection = amounttype

If bidoption = "竞争性比选" Then
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\2-询价函-1.docx"
Application.DisplayAlerts = True
ElseIf bidoption = "询价比价" Then
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\3-询价函-1.docx"
Application.DisplayAlerts = True
Else
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\3-询价函.docx"
Application.DisplayAlerts = True
End If
.Documents.Close
End With

''company2
If bidoption <> "议价方式" Then
With wrd
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen < 18 Then
.Selection = title & Chr(10)
ElseIf titlelen < 23 Then
.Selection = Chr(10) & title & Chr(10)
Else
.Selection = Chr(10) & title
End If
.Selection.Goto Name:="company"
.Selection = company2
.Selection.Goto Name:="theme"
.Selection = title
If bidoption <> "竞争性比选" Then
.Selection.Goto Name:="work"
.Selection = work
End If
If bidoption = "竞争性比选" Then
.Selection.Goto Name:="theme2"
.Selection = title
Else
If haveReq = "有" Then
.Selection.Goto Name:="requirement"
.Selection = "详见《鸡冠石污水处理厂" & requirement & "要求》"
Else
.Selection.Goto Name:="requirement"
.Selection = requirement2
End If
End If
.Selection.Goto Name:="worktime"
.Selection = worktime
If worktimelen > 45 And bidoption = "竞争性比选" Then
.Selection.Goto Name:="worktimemakeup"
.Selection = "详见《鸡冠石污水处理厂" & title & "竞争性比选》"
Else
.Selection.Goto Name:="worktimereason"
.Selection = worktimereason
End If
.Selection.Goto Name:="price"
.Selection = projtype
.Selection.Goto Name:="payreason"
.Selection = payreason
.Selection.Goto Name:="qualification1"
.Selection = qualification
.Selection.Goto Name:="deadlineyear"
.Selection = dealineyear
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="askyear2"
.Selection = askyear
.Selection.Goto Name:="asktime2"
.Selection = asktime
.Selection.Goto Name:="deadprice"
.Selection = deadprice
.Selection.Goto Name:="amounttype"
.Selection = amounttype

If bidoption = "竞争性比选" Then
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\2-询价函-2.docx"
Application.DisplayAlerts = True
Else
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\3-询价函-2.docx"
Application.DisplayAlerts = True
End If
.Documents.Close
End With

''company3
With wrd
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen < 18 Then
.Selection = title & Chr(10)
ElseIf titlelen < 23 Then
.Selection = Chr(10) & title & Chr(10)
Else
.Selection = Chr(10) & title
End If
.Selection.Goto Name:="company"
.Selection = company3
.Selection.Goto Name:="theme"
.Selection = title
If bidoption <> "竞争性比选" Then
.Selection.Goto Name:="work"
.Selection = work
End If
If bidoption = "竞争性比选" Then
.Selection.Goto Name:="theme2"
.Selection = title
Else
If haveReq = "有" Then
.Selection.Goto Name:="requirement"
.Selection = "详见《鸡冠石污水处理厂" & requirement & "要求》"
Else
.Selection.Goto Name:="requirement"
.Selection = requirement2
End If
End If
.Selection.Goto Name:="worktime"
.Selection = worktime
If worktimelen > 45 And bidoption = "竞争性比选" Then
.Selection.Goto Name:="worktimemakeup"
.Selection = "详见《鸡冠石污水处理厂" & title & "竞争性比选》"
Else
.Selection.Goto Name:="worktimereason"
.Selection = worktimereason
End If
.Selection.Goto Name:="price"
.Selection = projtype
.Selection.Goto Name:="payreason"
.Selection = payreason
.Selection.Goto Name:="qualification1"
.Selection = qualification
.Selection.Goto Name:="deadlineyear"
.Selection = dealineyear
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="askyear2"
.Selection = askyear
.Selection.Goto Name:="asktime2"
.Selection = asktime
.Selection.Goto Name:="deadprice"
.Selection = deadprice
.Selection.Goto Name:="amounttype"
.Selection = amounttype

If bidoption = "竞争性比选" Then
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\2-询价函-3.docx"
Application.DisplayAlerts = True
Else
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\3-询价函-3.docx"
Application.DisplayAlerts = True
End If
.Documents.Close
End With
End If

'2询价函会签单---x
With wrd
If bidoption = "询价比价" Then
wrdPath = nowPath & "\" & bidoption & "\2询价函会签单.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
.Selection = title
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="company1"
.Selection = company1
.Selection.Goto Name:="company2"
.Selection = company2
.Selection.Goto Name:="company3"
.Selection = company3
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\2-询价函会签单.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'2询价函会签单---y
With wrd
If bidoption = "议价方式" Then
wrdPath = nowPath & "\" & bidoption & "\2询价函会签单.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
.Selection = title
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="company1"
.Selection = company1
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\2-询价函会签单.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'3比选文件---j
If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\3比选文件.docx"
With wrd
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 14 Then
.Selection = Mid(title, 1, titlelen - 2) & Chr(10) & Mid(title, titlelen - 1, 2)
Else
.Selection = title
End If
.Selection.Goto Name:="znaskyear"
.Selection = znaskyear
.Selection.Goto Name:="znasktime"
.Selection = znasktime
.Selection.Goto Name:="theme"
.Selection = title
If bidtype = "采   购" Then
.Selection.Goto Name:="requirement"
.Selection = requirement
Else
If haveReq = "有" Then
.Selection.Goto Name:="requirement"
.Selection = requirement
Else
.Selection.Goto Name:="requirement2"
.Selection = requirement2
End If
End If
.Selection.Goto Name:="worktimereason"
.Selection = worktimereason
.Selection.Goto Name:="qualification"
.Selection = qualification
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="payreason"
.Selection = payreason
.Selection.Goto Name:="askyear"
.Selection = askyear
.Selection.Goto Name:="asktime"
.Selection = asktime
.Selection.Goto Name:="askyear2"
.Selection = askyear
.Selection.Goto Name:="asktime2"
.Selection = asktime
.Selection.Goto Name:="tecconsulting"
.Selection = tecconsulting
.Selection.Goto Name:="deadlineyear"
.Selection = dealineyear
.Selection.Goto Name:="deadline2"
.Selection = dealine
.Selection.Goto Name:="deadlinehm2"
.Selection = deadlinehm
.Selection.Goto Name:="znaskyear2"
.Selection = znaskyear
.Selection.Goto Name:="znasktime2"
.Selection = znasktime
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\3-比选文件.docx"
Application.DisplayAlerts = True
.Documents.Close
End With
End If

'4会议纪要---j
With wrd
If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\4会议纪要.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 20 Then
.Selection = title & Chr(10)
Else
.Selection = title
End If
.Selection.Goto Name:="meetingyear"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime"
.Selection = meetingtime
.Selection.Goto Name:="host"
.Selection = host
.Selection.Goto Name:="participants"
.Selection = participants
.Selection.Goto Name:="meetingyear2"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime2"
.Selection = meetingtime
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="reason"
.Selection = reason
.Selection.Goto Name:="theme2"
.Selection = title
.Selection.Goto Name:="askyear"
.Selection = askyear
.Selection.Goto Name:="asktime"
.Selection = asktime
.Selection.Goto Name:="company1"
.Selection = company1
.Selection.Goto Name:="company2"
.Selection = company2
.Selection.Goto Name:="company3"
.Selection = company3
.Selection.Goto Name:="deadlineyear"
.Selection = dealineyear
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="ccompany1"
.Selection = company1
.Selection.Goto Name:="ccompany2"
.Selection = company2
.Selection.Goto Name:="ccompany3"
.Selection = company3
.Selection.Goto Name:="theme3"
.Selection = title
.Selection.Goto Name:="cccompany1"
.Selection = company1
.Selection.Goto Name:="cccompany2"
.Selection = company2
.Selection.Goto Name:="cccompany3"
.Selection = company3
.Selection.Goto Name:="price1"
.Selection = price1
.Selection.Goto Name:="price2"
.Selection = price2
.Selection.Goto Name:="price3"
.Selection = price3
.Selection.Goto Name:="meetingyear3"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime3"
.Selection = meetingtime
.Selection.Goto Name:="meetingtimeap"
.Selection = meetingtimeap
.Selection.Goto Name:="theme4"
.Selection = title
.Selection.Goto Name:="winnercompany"
.Selection = wincompany
.Selection.Goto Name:="winnerprice"
.Selection = winprice
.Selection.Goto Name:="meetingyear4"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime4"
.Selection = meetingtime
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\4-会议纪要.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'5中标通知书---j
With wrd
If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\5中标通知书.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 15 Then
.Selection = title & Chr(10)
Else
.Selection = title
End If
.Selection.Goto Name:="asignyear"
.Selection = asignyear
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="winnercompany"
.Selection = wincompany
.Selection.Goto Name:="winnerprice"
.Selection = winprice
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\5-中标通知书.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'6唱标记录表（备用）---j
'7唱标记录表（备用）---x
If bidoption <> "议价方式" Then
With wrd
If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\6唱标记录表（备用）.docx"
Else
wrdPath = nowPath & "\" & bidoption & "\7唱标记录表（备用）.docx"
End If
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 16 Then
.Selection = title & Chr(10)
Else
.Selection = title
End If
.Selection.Goto Name:="company1"
.Selection = company1
.Selection.Goto Name:="company2"
.Selection = company2
.Selection.Goto Name:="company3"
.Selection = company3
.Selection.Goto Name:="price1"
.Selection = price1
.Selection.Goto Name:="price2"
.Selection = price2
.Selection.Goto Name:="price3"
.Selection = price3
If bidoption = "竞争性比选" Then
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\6-唱标记录表（备用）.docx"
Application.DisplayAlerts = True
Else
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\7-唱标记录表（备用）.docx"
Application.DisplayAlerts = True
End If
.Documents.Close
End With
End If

'6会议纪要---x
With wrd
If bidoption = "询价比价" Then
wrdPath = nowPath & "\" & bidoption & "\6会议纪要.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
If titlelen > 19 Then
.Selection = Chr(10) & title
Else
.Selection = title
End If
.Selection.Goto Name:="meetingyear"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime"
.Selection = meetingtime
.Selection.Goto Name:="host"
.Selection = host
.Selection.Goto Name:="participants"
.Selection = participants
.Selection.Goto Name:="reason"
.Selection = reason
.Selection.Goto Name:="asktime"
.Selection = asktime
.Selection.Goto Name:="company1"
.Selection = company1
.Selection.Goto Name:="company2"
.Selection = company2
.Selection.Goto Name:="company3"
.Selection = company3
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="deadline"
.Selection = dealine
.Selection.Goto Name:="deadlinehm"
.Selection = deadlinehm
.Selection.Goto Name:="meetingtime2"
.Selection = meetingtime
.Selection.Goto Name:="ccompany1"
.Selection = company1
.Selection.Goto Name:="ccompany2"
.Selection = company2
.Selection.Goto Name:="ccompany3"
.Selection = company3
.Selection.Goto Name:="price1"
.Selection = price1
.Selection.Goto Name:="price2"
.Selection = price2
.Selection.Goto Name:="price3"
.Selection = price3
.Selection.Goto Name:="winnercompany"
.Selection = wincompany
.Selection.Goto Name:="winnerprice"
.Selection = winprice
'Dim asdf As String
'asdf = Winprice
'Dim winnerpriceint As Integer
'winnerpriceint = Int(Winprice)
'.Selection.Goto Name:="nyinteger"
'If winnerpriceint = Winprice Then
'.Selection = "整"
'Else
'.Selection = ""
'End If
.Selection.Goto Name:="nyinteger"
.Selection = ""
.Selection.Goto Name:="meetingyear2"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime3"
.Selection = meetingtime
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\6-会议纪要.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With


'6会议纪要---y
With wrd
If bidoption = "议价方式" Then
wrdPath = nowPath & "\" & bidoption & "\6会议纪要.doc"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
.Selection = title
.Selection.Goto Name:="meetingyear1"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime1"
.Selection = meetingtime
.Selection.Goto Name:="host1"
.Selection = host
.Selection.Goto Name:="theme"
.Selection = title
.Selection.Goto Name:="reason"
.Selection = meetreason
.Selection.Goto Name:="work"
.Selection = work
.Selection.Goto Name:="company1"
.Selection = company1
.Selection.Goto Name:="ttype"
.Selection = makeUpType
.Selection.Goto Name:="company11"
.Selection = company1
.Selection.Goto Name:="price1"
.Selection = price1
.Selection.Goto Name:="company12"
.Selection = company1
.Selection.Goto Name:="price2"
.Selection = price2
.Selection.Goto Name:="company13"
.Selection = company1
.Selection.Goto Name:="meetingyear2"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime2"
.Selection = meetingtime
.Selection.Goto Name:="host2"
.Selection = host
.Selection.Goto Name:="company14"
.Selection = company1
.Selection.Goto Name:="anothercom"
.Selection = anotherComan
.Selection.Goto Name:="anticipates"
.Selection = participants
.Selection.Goto Name:="meetingyear3"
.Selection = meetingyear
.Selection.Goto Name:="meetingtime3"
.Selection = meetingtime
If bidtype = "维   修" Then
.Selection.Goto Name:="modify"
.Selection = "进行" & work
End If

Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\6-会议纪要.doc"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'7比选会签单---j
With wrd
If bidoption = "竞争性比选" Then
wrdPath = nowPath & "\" & bidoption & "\7比选会签单.docx"
.Documents.Open wrdPath
.Selection.Goto Name:="title"
.Selection = title
.Selection.Goto Name:="theme"
.Selection = title
If haveSecurity = False Then
.Selection.Goto Name:="security"
.Selection = ""
End If
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\7-比选会签单.docx"
Application.DisplayAlerts = True
.Documents.Close
End If
End With

'8合同审核---j、x

With wrd
If haveSecurity = True Then
wrdPath = nowPath & "\" & bidoption & "\8合同审核-1.docx"
Else
wrdPath = nowPath & "\" & bidoption & "\8合同审核-2.docx"
End If
.Documents.Open wrdPath
.Selection.Goto Name:="title"
.Selection = title
.Selection.Goto Name:="theme"
.Selection = title
If bidoption = "议价方式" Then
.Selection.Goto Name:="winnercompany"
.Selection = company1
.Selection.Goto Name:="winnerprice"
.Selection = price2
Else
.Selection.Goto Name:="winnercompany"
.Selection = wincompany
.Selection.Goto Name:="winnerprice"
.Selection = winprice
End If
Application.DisplayAlerts = False
.ActiveDocument.SaveAs nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\8-合同审核.docx"
Application.DisplayAlerts = True
.Documents.Close
End With

wrd.Quit
Set wrd = Nothing
'============================================================================================================
Dim nowwork As Workbook
Set noework = ActiveWorkbook
Dim acts As Worksheet
Set acts = ActiveWorkbook.Sheets(1)
Dim actreq As Worksheet
Set actreq = ActiveWorkbook.Sheets(2)
Dim actreqs As Integer
actreqs = 3
Do While actreq.Cells(actreqs, "a").Value <> ""
actreqs = actreqs + 1
Loop

'9要求


Dim wb1 As Workbook
Dim reqName As String
Dim reqPath As String
If bidtype = "采   购" Then
reqName = "9-要求.xlsx"
Else
If haveReq = "有" Then
reqName = "9-要求.xlsx"
Else
reqName = "9-要求（无）.xlsx"
End If
End If
reqPath = nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\" & reqName
Dim fs1, f1, f11, fc1

Dim haveFile1 As Boolean

Set fs1 = CreateObject("scripting.filesystemobject")
Set f1 = fs1.GetFolder(nowPath & "\工作目录" & "\" & idate & "-" & titleToName)
Set fc1 = f1.Files
haveFile1 = False
For Each f11 In fc1
If f11.Name = reqName Then
haveFile1 = True
Exit For
End If
Next

Set fs1 = Nothing
Set f1 = Nothing
Set f11 = Nothing
Set fc1 = Nothing

If haveFile1 = True Then
Kill reqPath
End If


Set wb1 = Workbooks.Open(nowPath & "\" & bidoption & "\9要求.xlsx")
wb1.SaveAs reqPath
wb1.Close True

Dim reqlen As Integer
reqlen = Len(requirement)
Set wb1 = Workbooks.Open(reqPath)
If reqlen > 11 And reqlen < 16 Then
wb1.Sheets(1).Cells(1, 1).Value = "鸡冠石污水处理厂" & Mid(requirement, 1, reqlen - 2) & Chr(10) & Mid(requirement, reqlen - 1, 2) & "要求"
Else
wb1.Sheets(1).Cells(1, 1).Value = "鸡冠石污水处理厂" & requirement & "要求"
End If
Dim reqss As Integer
Dim reqsss As Integer

If bidoption = "询价比价" Or bidoption = "议价方式" Then

If actreqs > 3 Then
For reqsss = 3 To (actreqs - 1)
wb1.Sheets(1).Cells(reqsss, "a").Value = reqsss - 2
wb1.Sheets(1).Cells(reqsss, "b").Value = actreq.Cells(reqsss, "a").Value
wb1.Sheets(1).Cells(reqsss, "c").Value = actreq.Cells(reqsss, "b").Value
wb1.Sheets(1).Cells(reqsss, "e").Value = actreq.Cells(reqsss, "c").Value
wb1.Sheets(1).Cells(reqsss, "d").Value = actreq.Cells(reqsss, "d").Value
Next reqsss
    With Range("a3:e" & (actreqs - 1)).Borders(xlEdgeLeft)                                  '设定边框
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("a3:e" & (actreqs - 1)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("a3:e" & (actreqs - 1)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("a3:e" & (actreqs - 1)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("a3:e" & (actreqs - 1)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Range("a3:e" & (actreqs - 1)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
End If

Else
If actreqs > 3 Then
For reqsss = 6 To (actreqs + 2)
wb1.Sheets(1).Cells(reqsss, "a").Value = reqsss - 5
wb1.Sheets(1).Cells(reqsss, "b").Value = actreq.Cells(reqsss - 3, "a").Value
wb1.Sheets(1).Cells(reqsss, "c").Value = actreq.Cells(reqsss - 3, "b").Value
wb1.Sheets(1).Cells(reqsss, "d").Value = actreq.Cells(reqsss - 3, "c").Value
wb1.Sheets(1).Cells(reqsss, "e").Value = actreq.Cells(reqsss - 3, "d").Value
Next reqsss
    With Range("a6:e" & (actreqs + 2)).Borders(xlEdgeLeft)                                  '设定边框
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range("a6:e" & (actreqs + 2)).Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range("a6:e" & (actreqs + 2)).Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range("a6:e" & (actreqs + 2)).Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range("a6:e" & (actreqs + 2)).Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
    With Range("a6:e" & (actreqs + 2)).Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlMedium
    End With
End If
End If

wb1.Save
wb1.Close True
Set wb1 = Nothing




'总结表

Dim savePath2, saveName2 As String
savePath2 = nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\"         '新建文件保存的路径
saveName2 = "总结表-" & titleToName & ".xlsx"                         '新建文件的名称

Dim fs2, f2, f12, fc2
Dim haveFile2 As Boolean

Set fs2 = CreateObject("scripting.filesystemobject")
Set f2 = fs2.GetFolder(nowPath & "\工作目录" & "\" & idate & "-" & titleToName)
Set fc2 = f2.Files
haveFile2 = False
For Each f12 In fc2
If f12.Name = saveName2 Then
haveFile2 = True
Exit For
End If
Next

Set fs2 = Nothing
Set f2 = Nothing
Set f12 = Nothing
Set fc2 = Nothing


Dim excelApp2, excelWB2 As Object
If haveFile2 = False Then
Set excelApp2 = CreateObject("Excel.Application")
Set excelWB2 = excelApp2.Workbooks.Add
excelWB2.SaveAs savePath2 & saveName2
excelWB2.Close True
excelApp2.Quit
End If
Set excelApp2 = Nothing
Set excelWB2 = Nothing

Dim wb2 As Workbook
Set wb2 = Workbooks.Open(savePath2 & saveName2)
acts.Range("a1:e23").Copy wb2.Sheets(1).Range("a1:e23")
wb2.Sheets(1).DrawingObjects.Delete
If wb2.Sheets(1).Cells(23, "e").Value = "" Then
MsgBox "总结表未成功导出，请点击《总结表导出》，并到..\工作目录\0-导入的工作表中将此总结表文件复制替换此文件夹中的总结表文件！", vbCritical, "警告"
End If
If wb2.Sheets(1).Cells(1, 1).Value = "标的：副本" Then
MsgBox "总结表未成功导出，请点击《总结表导出》，并到..\工作目录\0-导入的工作表中将此总结表文件复制替换此文件夹中的总结表文件！", vbCritical, "警告"
End If

wb2.Sheets(1).Cells(1, 1).Value = "标的：副本"

If wb2.Sheets(1).Cells(1, 1).Value = "标的：总结表" Then
MsgBox "总结表导出发生未知错误，建议仔细核对导出的总结表，如果未正确导出总结表，请点击《总结表导出》，并到..\工作目录\0-导入的工作表中将此总结表文件复制替换此文件夹中的总结表文件！", vbCritical, "警告"
End If

Dim wb2reqs As Integer
wb2reqs = 3
If haveFile2 = True Then
Do While wb2.Sheets(2).Cells(wb2reqs, "a").Value <> ""
wb2reqs = wb2reqs + 1
Loop
End If

If haveFile2 = True And actreqs > 3 Then
wb2.Sheets(2).Range("a1:g" & wb2reqs).ClearContents
End If

actreq.Range("a1:g2").Copy wb2.Sheets(2).Range("a1:g2")
Dim summary As Integer
If actreqs > 3 Then
For summary = 3 To (actreqs - 1)
wb2.Sheets(2).Cells(summary, "a").Value = actreq.Cells(summary, "a").Value
wb2.Sheets(2).Cells(summary, "b").Value = actreq.Cells(summary, "b").Value
wb2.Sheets(2).Cells(summary, "c").Value = actreq.Cells(summary, "c").Value
wb2.Sheets(2).Cells(summary, "d").Value = actreq.Cells(summary, "d").Value
wb2.Sheets(2).Cells(summary, "e").Value = actreq.Cells(summary, "e").Value
wb2.Sheets(2).Cells(summary, "f").Value = actreq.Cells(summary, "f").Value
wb2.Sheets(2).Cells(summary, "g").Value = actreq.Cells(summary, "g").Value
Next summary
End If
wb2.Sheets(2).DrawingObjects.Delete
If wb2.Sheets(2).Cells(actreqs - 1, "g").Value = "" Then
MsgBox "要求表未成功导出或要求表未填写交货时间！", vbCritical, "警告"
End If
If wb2.Sheets(2).Cells(1, 1).Value = "标的：副本" Then
MsgBox "要求表未成功导出，请点击《总结表导出》，并到..\工作目录\0-导入的工作表中将此要求表文件复制替换此文件夹中的要求表文件！", vbCritical, "警告"
End If

wb2.Sheets(2).Cells(1, 1).Value = "标的：副本"

If wb2.Sheets(2).Cells(1, 1).Value = "标的：要求表" Then
MsgBox "要求表导出发生未知错误，建议仔细核对导出的要求表！", vbCritical, "警告"
End If


wb2.Save
wb2.Close True
Set wb2 = Nothing

'明细信息表

Dim savePath3, saveName3 As String
savePath3 = nowPath & "\工作目录" & "\" & idate & "-" & titleToName & "\"         '新建文件保存的路径
saveName3 = "明细信息表.xls"                         '新建文件的名称
Dim fs3, f3, f13, fc3
Dim haveFile3 As Boolean

Set fs3 = CreateObject("scripting.filesystemobject")
Set f3 = fs3.GetFolder(nowPath & "\工作目录" & "\" & idate & "-" & titleToName)
Set fc3 = f3.Files
haveFile3 = False
For Each f13 In fc3
If f13.Name = saveName3 Then
haveFile3 = True
Exit For
End If
Next

Set fs3 = Nothing
Set f3 = Nothing
Set f13 = Nothing
Set fc3 = Nothing

If haveFile3 = True Then
Kill savePath3 & saveName3
End If

Dim infoPath As String
Dim infoName As String
infoPath = nowPath & "\电商明细模版" & "\"
If infoType = "物资" Then
infoName = "明细信息表-物资.xls"
ElseIf infoType = "工程" Then
infoName = "明细信息表-工程.xls"
Else
infoName = "明细信息表-服务.xls"
End If

Dim wb3 As Workbook
Set wb3 = Workbooks.Open(infoPath & infoName)

Dim typenum As Integer

If infoType = "物资" Then                               '开始填写明细表
If actreqs > 3 Then
For typenum = 3 To actreqs - 1
wb3.Sheets(1).Cells(typenum + 1, "b").Value = actreq.Cells(typenum, "a").Value
If actreq.Cells(typenum, "b") <> "" Then
wb3.Sheets(1).Cells(typenum + 1, "d").Value = actreq.Cells(typenum, "b").Value
Else
wb3.Sheets(1).Cells(typenum + 1, "d").Value = makeUpType
End If
wb3.Sheets(1).Cells(typenum + 1, "f").Value = actreq.Cells(typenum, "c").Value
wb3.Sheets(1).Cells(typenum + 1, "e").Value = actreq.Cells(typenum, "d").Value
wb3.Sheets(1).Cells(typenum + 1, "c").Value = actreq.Cells(typenum, "e").Value
wb3.Sheets(1).Cells(typenum + 1, "i").Value = actreq.Cells(typenum, "f").Value
wb3.Sheets(1).Cells(typenum + 1, "j").Value = actreq.Cells(typenum, "g").Value
Next typenum
End If

ElseIf infoType = "工程" Then
wb3.Sheets(1).Cells(4, "a").Value = goodType
wb3.Sheets(1).Cells(4, "b").Value = title
wb3.Sheets(1).Cells(4, "c").Value = work
wb3.Sheets(1).Cells(4, "e").Value = preamount
wb3.Sheets(1).Cells(4, "f").Value = worktimereason

Else
wb3.Sheets(1).Cells(4, "a").Value = goodType
wb3.Sheets(1).Cells(4, "b").Value = title
wb3.Sheets(1).Cells(4, "c").Value = work
wb3.Sheets(1).Cells(4, "d").Value = preamount
wb3.Sheets(1).Cells(4, "e").Value = worktimereason
End If


wb3.SaveAs savePath3 & saveName3
wb3.Close True
Set wb3 = Nothing

'======================================================================================

Set actreq = Nothing
Set acts = Nothing
Application.ScreenUpdating = True

'提示
MsgBox "已经生成所有文件！", vbOKOnly, "成功"
End Sub

Function intToZnint(temp As Variant) As String
Dim midd As String

If temp = 0 Then
midd = "〇"
ElseIf temp = 1 Then
midd = "一"
ElseIf temp = 2 Then
midd = "二"
ElseIf temp = 3 Then
midd = "三"
ElseIf temp = 4 Then
midd = "四"
ElseIf temp = 5 Then
midd = "五"
ElseIf temp = 6 Then
midd = "六"
ElseIf temp = 7 Then
midd = "七"
ElseIf temp = 8 Then
midd = "八"
ElseIf temp = 9 Then
midd = "九"
ElseIf temp = 10 Then
midd = "十"
ElseIf temp = 11 Then
midd = "十一"
ElseIf temp = 12 Then
midd = "十二"
ElseIf temp = 13 Then
midd = "十三"
ElseIf temp = 14 Then
midd = "十四"
ElseIf temp = 15 Then
midd = "十五"
ElseIf temp = 16 Then
midd = "十六"
ElseIf temp = 17 Then
midd = "十七"
ElseIf temp = 18 Then
midd = "十八"
ElseIf temp = 19 Then
midd = "十九"
ElseIf temp = 20 Then
midd = "二十"
ElseIf temp = 21 Then
midd = "二十一"
ElseIf temp = 22 Then
midd = "二十二"
ElseIf temp = 23 Then
midd = "二十三"
ElseIf temp = 24 Then
midd = "二十四"
ElseIf temp = 25 Then
midd = "二十五"
ElseIf temp = 26 Then
midd = "二十六"
ElseIf temp = 27 Then
midd = "二十七"
ElseIf temp = 28 Then
midd = "二十八"
ElseIf temp = 29 Then
midd = "二十九"
ElseIf temp = 30 Then
midd = "三十"
ElseIf temp = 31 Then
midd = "三十一"
Else
MsgBox "含有非法字符！"
End If
intToZnint = midd
End Function
```


You can go to [https://github.com/Zhongnibug/Zhongnibug.github.io](https://github.com/Zhongnibug/Zhongnibug.github.io).
