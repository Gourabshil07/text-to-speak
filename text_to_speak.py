''' 
write a program to pronounce list of names 
if you are given a list l as follows:
    #Python
l = ["Gourab", "Souvik", "Subham"]

your program should pronounce

shoutout to Gourab
shoutout to Souvik

'''

import win32com.client as wincl

list = ["Gourab", "Souvik","Subham", "Manash"]

def sappi(name):
    speaker_number = 1
    spk = wincl.Dispatch("SAPI.SpVoice")
    vcs = spk.GetVoices()
    SVSFlag = 11
    print(vcs.Item (speaker_number) .GetAttribute ("Name")) # speaker name
    spk.Voice
    spk.SetVoice(vcs.Item(speaker_number)) # set voice (see Windows Text-to-Speech settings)
    spk.Speak(f"Shout out to {name}")

for names in list:
    sappi(names)
