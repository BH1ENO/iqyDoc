# iqyDoc
Asynchronously grab information of iqy documentary using VBA from Excel. 

Basic skill:
1.DOM

2.WinHttpRequest(Event handler, finnal version, can handle URL redirect)

3.XMLHTTP60(callback javascript, alpha version, cannot handle URL redirect)

4.URL redirect (301, 302) cause acess denied problem.

5.VBA callback tip: put Attribute OnReadyStateChange.VB_UserMemId = 0 just below the event callback handler sub name in export class file, then import the export file again. 
For example:

Public Sub OnReadyStateChange()
Attribute OnReadyStateChange.VB_UserMemId = 0
    Dim strDoc As String
    Dim lngBegin As Long
    Dim strTmp As String
    Dim strToFind As String
    


Grab 3 kinds of information.
1.static webpage elements by DOM.

2.in-page javascript variables by Instr function.

3.Ajax data by find data querying URL and build new URL to query.

