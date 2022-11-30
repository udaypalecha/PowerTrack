set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
  iFull = oResult.FullChargedCapacity
next

Dim sapi
Set sapi = createObject("sapi.spVoice")
Set sapi.Voice = sapi.GetVoices.Item(2)

while(1)
  for each oResult in oServices.ExecQuery("select * from batterystatus")
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining/iFull)*100) mod 100
  
  If bCharging and (iPercent >= 79) Then
    sapi.Speak("Battery 80% charged, please disconnect the charger!")
    msgbox "80% battery charged!"
  
  ElseIf NOT(bCharging) and (iPercent<=22) Then
    sapi.Speak("Battery at critical levels. Please connect the charger!")
    msgbox "20% Battery Remaining!"

  ElseIf NOT(bCharging) and (iPercent<=32) Then
    sapi.Speak("Battery at 30%")
  End If
  wscript.sleep 30000 ' 30seconds
wend