<div align="center">

## ADO Header\-Detail Relational Tables \(without using datacontrols\)


</div>

### Description

This is a demonstration to show you how Header Tables relate with Detail Tables. Don't forget to vote for it.
 
### More Info
 
Save this text to "ADOHeaderDetail.FRM" and use the NWIND.MDB from Northwind.


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Walter Narvasa](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/walter-narvasa.md)
**Level**          |Advanced
**User Rating**    |4.7 (14 globes from 3 users)
**Compatibility**  |VB 5\.0, VB 6\.0
**Category**       |[Databases/ Data Access/ DAO/ ADO](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/databases-data-access-dao-ado__1-6.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/walter-narvasa-ado-header-detail-relational-tables-without-using-datacontrols__1-8771/archive/master.zip)





### Source Code

```
VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form ADOHeaderDetail
  BorderStyle   =  1 'Fixed Single
  Caption     =  "Order Entry - ADO Header-Detail Sample by Walter A. Narvasa"
  ClientHeight  =  6495
  ClientLeft   =  1095
  ClientTop    =  390
  ClientWidth   =  9735
  KeyPreview   =  -1 'True
  LinkTopic    =  "Form1"
  MaxButton    =  0  'False
  MinButton    =  0  'False
  ScaleHeight   =  6495
  ScaleWidth   =  9735
  StartUpPosition =  2 'CenterScreen
  Begin VB.PictureBox picButtons
   Align      =  2 'Align Bottom
   Appearance   =  0 'Flat
   BorderStyle   =  0 'None
   ForeColor    =  &H80000008&
   Height     =  300
   Left      =  0
   ScaleHeight   =  300
   ScaleWidth   =  9735
   TabIndex    =  34
   Top       =  5895
   Width      =  9735
   Begin VB.CommandButton cmdCancel
     Caption     =  "&Undo"
     Height     =  300
     Left      =  3600
     TabIndex    =  41
     Top       =  0
     Visible     =  0  'False
     Width      =  1095
   End
   Begin VB.CommandButton cmdClose
     Caption     =  "E&xit"
     Height     =  300
     Left      =  4800
     TabIndex    =  39
     Top       =  0
     Width      =  1095
   End
   Begin VB.CommandButton cmdRefresh
     Caption     =  "&Refresh"
     Height     =  300
     Left      =  3600
     TabIndex    =  38
     Top       =  0
     Width      =  1095
   End
   Begin VB.CommandButton cmdAdd
     Caption     =  "&New"
     Height     =  300
     Left      =  0
     TabIndex    =  35
     Top       =  0
     Width      =  1095
   End
   Begin VB.CommandButton cmdEdit
     Caption     =  "&Edit"
     Height     =  300
     Left      =  1200
     TabIndex    =  36
     Top       =  0
     Width      =  1095
   End
   Begin VB.CommandButton cmdUpdate
     Caption     =  "&Save"
     Height     =  300
     Left      =  2400
     TabIndex    =  40
     Top       =  0
     Visible     =  0  'False
     Width      =  1095
   End
   Begin VB.CommandButton cmdDelete
     Caption     =  "&Delete"
     Height     =  300
     Left      =  2400
     TabIndex    =  37
     Top       =  0
     Width      =  1095
   End
  End
  Begin VB.PictureBox picStatBox
   Align      =  2 'Align Bottom
   Appearance   =  0 'Flat
   BorderStyle   =  0 'None
   ForeColor    =  &H80000008&
   Height     =  300
   Left      =  0
   ScaleHeight   =  300
   ScaleWidth   =  9735
   TabIndex    =  28
   Top       =  6195
   Width      =  9735
   Begin VB.CommandButton cmdLast
     Height     =  300
     Left      =  4545
     Picture     =  "ADOHeaderDetail.frx":0000
     Style      =  1 'Graphical
     TabIndex    =  32
     Top       =  0
     UseMaskColor  =  -1 'True
     Width      =  345
   End
   Begin VB.CommandButton cmdNext
     Height     =  300
     Left      =  4200
     Picture     =  "ADOHeaderDetail.frx":0342
     Style      =  1 'Graphical
     TabIndex    =  31
     Top       =  0
     UseMaskColor  =  -1 'True
     Width      =  345
   End
   Begin VB.CommandButton cmdPrevious
     Height     =  300
     Left      =  345
     Picture     =  "ADOHeaderDetail.frx":0684
     Style      =  1 'Graphical
     TabIndex    =  30
     Top       =  0
     UseMaskColor  =  -1 'True
     Width      =  345
   End
   Begin VB.CommandButton cmdFirst
     Height     =  300
     Left      =  0
     Picture     =  "ADOHeaderDetail.frx":09C6
     Style      =  1 'Graphical
     TabIndex    =  29
     Top       =  0
     UseMaskColor  =  -1 'True
     Width      =  345
   End
   Begin VB.Label lblStatus
     BackColor    =  &H00FFFFFF&
     BorderStyle   =  1 'Fixed Single
     Height     =  285
     Left      =  690
     TabIndex    =  33
     Top       =  0
     Width      =  3360
   End
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipVia"
   Height     =  285
   Index      =  13
   Left      =  5640
   TabIndex    =  27
   Top       =  2415
   Width      =  3375
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipRegion"
   Height     =  285
   Index      =  12
   Left      =  5640
   TabIndex    =  25
   Top       =  2100
   Width      =  3375
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipPostalCode"
   Height     =  285
   Index      =  11
   Left      =  5640
   TabIndex    =  23
   Top       =  1785
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShippedDate"
   Height     =  285
   Index      =  10
   Left      =  5640
   TabIndex    =  21
   Top       =  1455
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipName"
   Height     =  285
   Index      =  9
   Left      =  5640
   TabIndex    =  19
   Top       =  1140
   Width      =  3855
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipCountry"
   Height     =  285
   Index      =  8
   Left      =  5640
   TabIndex    =  17
   Top       =  825
   Width      =  3855
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipCity"
   Height     =  285
   Index      =  7
   Left      =  5640
   TabIndex    =  15
   Top       =  495
   Width      =  3855
  End
  Begin VB.TextBox txtFields
   DataField    =  "ShipAddress"
   Height     =  285
   Index      =  6
   Left      =  5640
   TabIndex    =  13
   Top       =  180
   Width      =  3855
  End
  Begin VB.TextBox txtFields
   DataField    =  "RequiredDate"
   Height     =  285
   Index      =  5
   Left      =  2040
   TabIndex    =  11
   Top       =  1785
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "Freight"
   Height     =  285
   Index      =  4
   Left      =  2040
   TabIndex    =  9
   Top       =  1455
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "CustomerID"
   Height     =  285
   Index      =  3
   Left      =  2040
   TabIndex    =  7
   Top       =  1140
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "EmployeeID"
   Height     =  285
   Index      =  2
   Left      =  2040
   TabIndex    =  5
   Top       =  825
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "OrderDate"
   Height     =  285
   Index      =  1
   Left      =  2040
   TabIndex    =  3
   Top       =  495
   Width      =  1455
  End
  Begin VB.TextBox txtFields
   DataField    =  "OrderID"
   Height     =  285
   Index      =  0
   Left      =  2040
   TabIndex    =  1
   Top       =  180
   Width      =  1455
  End
  Begin MSDataGridLib.DataGrid grdDataGrid
   Height     =  2745
   Left      =  120
   TabIndex    =  42
   Top       =  3000
   Width      =  9360
   _ExtentX    =  16510
   _ExtentY    =  4842
   _Version    =  393216
   AllowUpdate   =  0  'False
   HeadLines    =  1
   RowHeight    =  15
   BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851}
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  400
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851}
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  400
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   ColumnCount   =  2
   BeginProperty Column00
     DataField    =  ""
     Caption     =  ""
     BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}
      Type      =  0
      Format     =  ""
      HaveTrueFalseNull=  0
      FirstDayOfWeek =  0
      FirstWeekOfYear =  0
      LCID      =  1033
      SubFormatType  =  0
     EndProperty
   EndProperty
   BeginProperty Column01
     DataField    =  ""
     Caption     =  ""
     BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED}
      Type      =  0
      Format     =  ""
      HaveTrueFalseNull=  0
      FirstDayOfWeek =  0
      FirstWeekOfYear =  0
      LCID      =  1033
      SubFormatType  =  0
     EndProperty
   EndProperty
   SplitCount   =  1
   BeginProperty Split0
     BeginProperty Column00
     EndProperty
     BeginProperty Column01
     EndProperty
   EndProperty
  End
  Begin VB.Label lblLabels
   Caption     =  "Detail Information:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  14
   Left      =  120
   TabIndex    =  43
   Top       =  2760
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipVia:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  13
   Left      =  3720
   TabIndex    =  26
   Top       =  2415
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipRegion:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  12
   Left      =  3720
   TabIndex    =  24
   Top       =  2100
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipPostalCode:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  11
   Left      =  3720
   TabIndex    =  22
   Top       =  1785
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShippedDate:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  10
   Left      =  3720
   TabIndex    =  20
   Top       =  1455
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipName:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  9
   Left      =  3720
   TabIndex    =  18
   Top       =  1140
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipCountry:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  8
   Left      =  3720
   TabIndex    =  16
   Top       =  825
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipCity:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  7
   Left      =  3720
   TabIndex    =  14
   Top       =  495
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "ShipAddress:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  6
   Left      =  3720
   TabIndex    =  12
   Top       =  180
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "RequiredDate:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  5
   Left      =  120
   TabIndex    =  10
   Top       =  1785
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "Freight:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  4
   Left      =  120
   TabIndex    =  8
   Top       =  1455
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "CustomerID:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  3
   Left      =  120
   TabIndex    =  6
   Top       =  1140
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "EmployeeID:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  2
   Left      =  120
   TabIndex    =  4
   Top       =  825
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "OrderDate:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  1
   Left      =  120
   TabIndex    =  2
   Top       =  495
   Width      =  1815
  End
  Begin VB.Label lblLabels
   Alignment    =  1 'Right Justify
   Caption     =  "OrderID:"
   BeginProperty Font
     Name      =  "MS Sans Serif"
     Size      =  8.25
     Charset     =  0
     Weight     =  700
     Underline    =  0  'False
     Italic     =  0  'False
     Strikethrough  =  0  'False
   EndProperty
   Height     =  255
   Index      =  0
   Left      =  120
   TabIndex    =  0
   Top       =  180
   Width      =  1815
  End
End
Attribute VB_Name = "ADOHeaderDetail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Program Sample by: Walter A. Narvasa
'Country: Philippines
'Experience: 6 years in Database Programming
'Email: walter@wancom.8k.com
'Website: wancom.8k.com
Dim WithEvents adoPrimaryRS As Recordset
Attribute adoPrimaryRS.VB_VarHelpID = -1
Dim mbChangedByCode As Boolean
Dim mvBookMark As Variant
Dim mbEditFlag As Boolean
Dim mbAddNewFlag As Boolean
Dim mbDataChanged As Boolean
Private Sub Form_Load()
 Dim db As Connection
 Set db = New Connection
 db.CursorLocation = adUseClient
 db.Open "PROVIDER=MSDataShape;Data PROVIDER=Microsoft.Jet.OLEDB.3.51;Data Source=D:\Program Language\Microsoft Visual Studio\VB98\NWIND.MDB;"
 Set adoPrimaryRS = New Recordset
 adoPrimaryRS.Open "SHAPE {select OrderID,OrderDate,EmployeeID,CustomerID,Freight,RequiredDate,ShipAddress,ShipCity,ShipCountry,ShipName,ShippedDate,ShipPostalCode,ShipRegion,ShipVia from Orders Order by OrderID} AS ParentCMD APPEND ({select OrderID,ProductID,Quantity,UnitPrice,Discount from [Order Details] Order by ProductID } AS ChildCMD RELATE OrderID TO OrderID) AS ChildCMD", db, adOpenStatic, adLockOptimistic
 Dim oText As TextBox
 'Bind the text boxes to the data provider
 For Each oText In Me.txtFields
  Set oText.DataSource = adoPrimaryRS
 Next
 Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
 mbDataChanged = False
End Sub
Private Sub Form_Resize()
 On Error Resume Next
 'This will resize the grid when the form is resized
 grdDataGrid.Width = Me.ScaleWidth
 grdDataGrid.Height = Me.ScaleHeight - grdDataGrid.Top - 30 - picButtons.Height - picStatBox.Height
 lblStatus.Width = Me.Width - 1500
 cmdNext.Left = lblStatus.Width + 700
 cmdLast.Left = cmdNext.Left + 340
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
 If mbEditFlag Or mbAddNewFlag Then Exit Sub
 Select Case KeyCode
  Case vbKeyEscape
   cmdClose_Click
  Case vbKeyEnd
   cmdLast_Click
  Case vbKeyHome
   cmdFirst_Click
  Case vbKeyUp, vbKeyPageUp
   If Shift = vbCtrlMask Then
    cmdFirst_Click
   Else
    cmdPrevious_Click
   End If
  Case vbKeyDown, vbKeyPageDown
   If Shift = vbCtrlMask Then
    cmdLast_Click
   Else
    cmdNext_Click
   End If
 End Select
End Sub
Private Sub Form_Unload(Cancel As Integer)
 Screen.MousePointer = vbDefault
End Sub
Private Sub adoPrimaryRS_MoveComplete(ByVal adReason As ADODB.EventReasonEnum, ByVal pError As ADODB.Error, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This will display the current record position for this recordset
 lblStatus.Caption = "Record: " & CStr(adoPrimaryRS.AbsolutePosition)
End Sub
Private Sub adoPrimaryRS_WillChangeRecord(ByVal adReason As ADODB.EventReasonEnum, ByVal cRecords As Long, adStatus As ADODB.EventStatusEnum, ByVal pRecordset As ADODB.Recordset)
 'This is where you put validation code
 'This event gets called when the following actions occur
 Dim bCancel As Boolean
 Select Case adReason
 Case adRsnAddNew
 Case adRsnClose
 Case adRsnDelete
 Case adRsnFirstChange
 Case adRsnMove
 Case adRsnRequery
 Case adRsnResynch
 Case adRsnUndoAddNew
 Case adRsnUndoDelete
 Case adRsnUndoUpdate
 Case adRsnUpdate
 End Select
 If bCancel Then adStatus = adStatusCancel
End Sub
Private Sub cmdAdd_Click()
 On Error GoTo AddErr
 With adoPrimaryRS
  If Not (.BOF And .EOF) Then
   mvBookMark = .Bookmark
  End If
  .AddNew
  lblStatus.Caption = "Add record"
  mbAddNewFlag = True
  SetButtons False
 End With
 Exit Sub
AddErr:
 MsgBox Err.Description
End Sub
Private Sub cmdDelete_Click()
 On Error GoTo DeleteErr
 With adoPrimaryRS
  .Delete
  .MoveNext
  If .EOF Then .MoveLast
 End With
 Exit Sub
DeleteErr:
 MsgBox Err.Description
End Sub
Private Sub cmdRefresh_Click()
 'This is only needed for multi user apps
 On Error GoTo RefreshErr
 Set grdDataGrid.DataSource = Nothing
 adoPrimaryRS.Requery
 Set grdDataGrid.DataSource = adoPrimaryRS("ChildCMD").UnderlyingValue
 Exit Sub
RefreshErr:
 MsgBox Err.Description
End Sub
Private Sub cmdEdit_Click()
 On Error GoTo EditErr
 lblStatus.Caption = "Edit record"
 mbEditFlag = True
 SetButtons False
 Exit Sub
EditErr:
 MsgBox Err.Description
End Sub
Private Sub cmdCancel_Click()
 On Error Resume Next
 SetButtons True
 mbEditFlag = False
 mbAddNewFlag = False
 adoPrimaryRS.CancelUpdate
 If mvBookMark > 0 Then
  adoPrimaryRS.Bookmark = mvBookMark
 Else
  adoPrimaryRS.MoveFirst
 End If
 mbDataChanged = False
End Sub
Private Sub cmdUpdate_Click()
 On Error GoTo UpdateErr
 adoPrimaryRS.UpdateBatch adAffectAll
 If mbAddNewFlag Then
  adoPrimaryRS.MoveLast       'move to the new record
 End If
 mbEditFlag = False
 mbAddNewFlag = False
 SetButtons True
 mbDataChanged = False
 Exit Sub
UpdateErr:
 MsgBox Err.Description
End Sub
Private Sub cmdClose_Click()
 Unload Me
End Sub
Private Sub cmdFirst_Click()
 On Error GoTo GoFirstError
 adoPrimaryRS.MoveFirst
 mbDataChanged = False
 Exit Sub
GoFirstError:
 MsgBox Err.Description
End Sub
Private Sub cmdLast_Click()
 On Error GoTo GoLastError
 adoPrimaryRS.MoveLast
 mbDataChanged = False
 Exit Sub
GoLastError:
 MsgBox Err.Description
End Sub
Private Sub cmdNext_Click()
 On Error GoTo GoNextError
 If Not adoPrimaryRS.EOF Then adoPrimaryRS.MoveNext
 If adoPrimaryRS.EOF And adoPrimaryRS.RecordCount > 0 Then
  Beep
   'moved off the end so go back
  adoPrimaryRS.MoveLast
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoNextError:
 MsgBox Err.Description
End Sub
Private Sub cmdPrevious_Click()
 On Error GoTo GoPrevError
 If Not adoPrimaryRS.BOF Then adoPrimaryRS.MovePrevious
 If adoPrimaryRS.BOF And adoPrimaryRS.RecordCount > 0 Then
  Beep
  'moved off the end so go back
  adoPrimaryRS.MoveFirst
 End If
 'show the current record
 mbDataChanged = False
 Exit Sub
GoPrevError:
 MsgBox Err.Description
End Sub
Private Sub SetButtons(bVal As Boolean)
 cmdAdd.Visible = bVal
 cmdEdit.Visible = bVal
 cmdUpdate.Visible = Not bVal
 cmdCancel.Visible = Not bVal
 cmdDelete.Visible = bVal
 cmdClose.Visible = bVal
 cmdRefresh.Visible = bVal
 cmdNext.Enabled = bVal
 cmdFirst.Enabled = bVal
 cmdLast.Enabled = bVal
 cmdPrevious.Enabled = bVal
End Sub
```

