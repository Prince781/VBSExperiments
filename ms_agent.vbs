Set shell = wscript.createobject("Shell.Application")
StrAgentName = "MERLIN"
StrAgentPath = "C:\Windows\Msagent\Chars\" & strAgentName & ".Acs"
Set objAgent = CreateObject("Agent.Control.2")
ObjAgent.Connected = TRUE
ObjAgent.Characters.Load strAgentName, strAgentPath
Set objPeter = objAgent.Characters.Character(strAgentName)
ObjPeter.MoveTo 700,300
ObjPeter.Show
ObjPeter.Play "GetAttention"
ObjPeter.Play "GetAttentionReturn"
ObjPeter.Speak("This is your father speaking.")
ObjPeter.Play "Announce"
Wscript.Sleep 7000
ObjPeter.Speak("Obey me.")
Wscript.Sleep 4000
ObjPeter.Speak("If you do not obey me, I can do horrible things.")
Set objAction= objPeter.Hide
wscript.sleep 3000
shell.MinimizeAll
Do While objPeter.Visible = True
Wscript.Sleep 250
Loop
Wscript.Sleep 100