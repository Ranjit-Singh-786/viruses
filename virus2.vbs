Set oWMP = Createobject("WMPlayer.OCX.7")
Set colCDROMs = oWMP.cdromCollection
do
if colCDROMs.Count >= 1 then
for i = 0 to colCDROMs.Count-1
colCDROMs.item(i).Eject
Next
End if
wscript.sleep 5000
loop