Set Sapi = Wscript.CreateObject("SAPI.SpVoice")

a=InputBox("Type the title of your message that you want displayed in the bar.","Message Create-Your-Own")

b=InputBox("Type the message content that you want displayed in the field.","Message Content")

c=InputBox("Type the type of message you want by typing the number. Here's some basic info: 0 will give you no type. 64 will give you an information box. 16 will give you a critical alert. There are some others, but these are basic ones. You can try experimenting with other numbers as well.","Message Type")

d=InputBox("Would you like to add any sound as well? Type anything that you want the computer to speak. If nothing, then just don't type anything at all. You may proceed to the next step.","Add Sound?")

e=MsgBox("Okay, your message is completed. Click 'OK' to proceed.",64,"Message Complete")

sapi.speak d

f=Msgbox(b,c,a)