set oLocator = CreateObject("WbemScripting.SWbemLocator")
set oServices = oLocator.ConnectServer(".","root\wmi")
set oResults = oServices.ExecQuery("select * from batteryfullchargedcapacity")
for each oResult in oResults
  iFull = oResult.FullChargedCapacity
next

Dim sapi
set sapi = createObject("sapi.spVoice")
set sapi.Voice = sapi.GetVoices.Item(1)

while(1)
  for each oResult in oServices.ExecQuery("select * from batterystatus")
    iRemaining = oResult.RemainingCapacity
    bCharging = oResult.Charging
  next
  iPercent = ((iRemaining/iFull)*100) mod 100

  If bCharging and (iPercent >= 79) Then
    sapi.Speak("Battery fully charged, please disconnect the charger!")
  ElseIf bCharging and (iPercent = 80) Then
    msgbox "80% battery charged!"

  ElseIf NOT(bCharging) and (iPercent<=22) Then
    sapi.Speak("Battery at critical levels. Please connect the charger!")
    msgbox "20% Battery Remaining!"

  ElseIf NOT(bCharging) and (iPercent<=31) Then
    sapi.Speak("Battery at 30%")
  End If
  
  If (iPercent>=40) and (iPercent<=70) Then
    wscript.sleep 420000  ' 7minutes
  Else
    wscript.sleep 30000 ' 30seconds
  End If
wend
