''' 
write a program to pronounce list of names using win32 api for mac
if you are given a list l as follows:
    #Python
l = ["Rahul", "Nishat", "Harry"]

your program should pronounce

shoutout to rahul
shoutout to nisant

'''

import win32com.client as wincl

l = ["Rahul", "Nishat", "Harry","Gourab", "Subham"]

def sappi(name):
    speaker_number = 1
    spk = wincl.Dispatch("SAPI.SpVoice")
    vcs = spk.GetVoices()
    SVSFlag = 11
    print(vcs.Item (speaker_number) .GetAttribute ("Name")) # speaker name
    spk.Voice
    spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)
    spk.Speak(f"Shout out to {name}")

for names in l:
    sappi(names)
