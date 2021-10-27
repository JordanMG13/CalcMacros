sub TM_Helper
rem ----------------------------------------------------------------------
rem define variables
dim document   as object
dim dispatcher as object
rem ----------------------------------------------------------------------
rem get access to the document
document   = ThisComponent.CurrentController.Frame
dispatcher = createUnoService("com.sun.star.frame.DispatchHelper")

rem ----------------------------------------------------------------------
dim args1(0) as new com.sun.star.beans.PropertyValue
args1(0).Name = "ToPoint"
args1(0).Value = "$A$1"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args1())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:Paste", "", 0, Array())

rem ----------------------------------------------------------------------
Const lColumnIndex   As Long = 0
Const lNumberOfCells As Long = 250

Dim oDoc    As Object : oDoc    = ThisComponent
Dim oSheet  As Object : oSheet  = oDoc.CurrentController.ActiveSheet
Dim oColumn As Object : oColumn = oSheet.Columns.getByIndex( lColumnIndex )
Dim oRanges As Object : oRanges = oColumn.queryVisibleCells()
Dim oRange  As Object : oRange  = oRanges.getByIndex(0)
   
Dim oCursor As Object : oCursor = oSheet.createCursorByRange( oRange )
oCursor.collapseToSize ( 1, lNumberOfCells )
oDoc.CurrentController.select( oCursor )

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:TextToColumns", "", 0, Array())

rem ----------------------------------------------------------------------
dim oFilterField(3) As New com.sun.star.sheet.TableFilterField   

oSheet = thisComponent.Sheets.getByName("Sheet1") 
oRange = oSheet.getCellRangeByName("A1:A250")

oFilter= oRange.createFilterDescriptor(true)
'oRange.filter(oFilter)' 

with oFilter									
	.ContainsHeader = true						
	.CopyOutputData = false						
	.IsCaseSensitive = false				
	.UseRegularExpressions = false
end with

oFilterField(0).Field = 1
oFilterField(0).Operator = com.sun.star.sheet.FilterOperator2.DOES_NOT_CONTAIN
oFilterField(0).NumericValue = 20
			
oFilter.setFilterFields(oFilterField())
oRange.filter(oFilter)

rem ----------------------------------------------------------------------
dim args5(0) as new com.sun.star.beans.PropertyValue
args5(0).Name = "ToPoint"
args5(0).Value = "A2:AMJ250"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:DeleteRows", "", 0, Array())

rem ----------------------------------------------------------------------
oSheet = ThisComponent.getSheets().getByIndex(0)
oFilterDesc = oSheet.createFilterDescriptor(True)
oSheet.filter(oFilterDesc)

rem ----------------------------------------------------------------------
dim oFilterField2(3) As New com.sun.star.sheet.TableFilterField   

oSheet2 = thisComponent.Sheets.getByName("Sheet1") 
oRange2 = oSheet2.getCellRangeByName("A1:A250")

oFilter2= oRange2.createFilterDescriptor(true)
'oRange.filter(oFilter)' 

with oFilter2									
	.ContainsHeader = true						
	.CopyOutputData = false						
	.IsCaseSensitive = false				
	.UseRegularExpressions = false
end with

oFilterField2(0).Field = 1
oFilterField2(0).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
oFilterField2(0).StringValue = " "
			
oFilter2.setFilterFields(oFilterField2())
oRange2.filter(oFilter2)

rem ----------------------------------------------------------------------
dim args6(0) as new com.sun.star.beans.PropertyValue
args6(0).Name = "ToPoint"
args6(0).Value = "A2:AMJ250"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args6())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:DeleteRows", "", 0, Array())

rem ----------------------------------------------------------------------
oSheet = ThisComponent.getSheets().getByIndex(0)
oFilterDesc = oSheet.createFilterDescriptor(True)
oSheet.filter(oFilterDesc)

rem ----------------------------------------------------------------------
Dim oSheet3
Dim oRange3
Dim oSortFields(0) as new com.sun.star.util.SortField
Dim oSortDesc(0) as new com.sun.star.beans.PropertyValue
oSheet3 = ThisComponent.Sheets(0)

REM Set the range on which to sort
oRange3 = oSheet3.getCellRangeByName("A2:Z250")

REM Sort by the Average grade field in the range in descending order
oSortFields(0).Field = 0
oSortFields(0).SortAscending = TRUE

REM Set the sort fields to use
oSortDesc(0).Name = "SortFields"
oSortDesc(0).Value = oSortFields()

REM Now sort the range!
oRange3.Sort(oSortDesc())

end sub
