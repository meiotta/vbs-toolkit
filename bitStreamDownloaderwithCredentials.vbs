dim xHttp: Set xHttp = createobject("Microsoft.XMLHTTP")
dim bStrm: Set bStrm = createobject("Adodb.Stream")
xHttp.Open "GET", "https://thatserver.com/hosted/theFileYouWant.txt", False , "USERNAME", "PASSWORD"
xHttp.Send



with bStrm
    .type = 1 '//binary
    .open
    .write xHttp.responseBody
    .savetofile "\\netdrive\netfolder\netsubfolder\downloadedFile.txt", 2 '//overwrite
end with

