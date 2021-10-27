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
Const lNumberOfCells As Long = 200

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
oRange = oSheet.getCellRangeByName("A1:A100")

oFilter= oRange.createFilterDescriptor(true)
'oRange.filter(oFilter)' 

with oFilter									
	.ContainsHeader = true						
	.CopyOutputData = false						
	.IsCaseSensitive = false				
	.UseRegularExpressions = false
end with

oFilterField(0).Field = 1
oFilterField(0).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
oFilterField(0).StringValue = ""
oFilterField(1).Connection = com.sun.star.sheet.FilterConnection.OR
oFilterField(1).Field = 1
oFilterField(1).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
oFilterField(1).StringValue = "Department"
oFilterField(2).Connection = com.sun.star.sheet.FilterConnection.OR
oFilterField(2).Field = 1
oFilterField(2).Operator = com.sun.star.sheet.FilterOperator2.CONTAINS
oFilterField(2).StringValue = " "
			
oFilter.setFilterFields(oFilterField())
oRange.filter(oFilter)

rem ----------------------------------------------------------------------
dim args5(0) as new com.sun.star.beans.PropertyValue
args5(0).Name = "ToPoint"
args5(0).Value = "A2:AMJ80"

dispatcher.executeDispatch(document, ".uno:GoToCell", "", 0, args5())

rem ----------------------------------------------------------------------
dispatcher.executeDispatch(document, ".uno:DeleteRows", "", 0, Array())

rem ----------------------------------------------------------------------
oSheet = ThisComponent.getSheets().getByIndex(0)
oFilterDesc = oSheet.createFilterDescriptor(True)
oSheet.filter(oFilterDesc)

end sub
