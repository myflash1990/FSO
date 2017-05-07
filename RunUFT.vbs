Dim obj

Set obj = CreateObject("QuickTest.Application")

obj.launch 
obj.Visible = True



'Open the test 

obj.Open "C:\Users\tahsin\Documents\Unified Functional Testing\Search" , False,False

Set qtTest = obj.Test

qtTest.Run





obj.quit