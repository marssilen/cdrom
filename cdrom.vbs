Set TI = CreateObject("WMPlayer.OCX.7" )
Set CDROM = TI.cdromCollection
if CDROM.Count >= 1 then

For f=0 to 10

For i = 0 to CDROM.Count - 1
CDROM.Item(i).Eject
Next ' CDTRAY

Next

End If