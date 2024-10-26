import win32com.client as wincl
import os

with open('Sample.txt', 'r') as f:
    file = f.read()

speaker_number = 1
spk = wincl.Dispatch("SAPI.SpVoice")
vcs = spk.GetVoices()
SVSFlag = 11
print(vcs.Item (speaker_number) .GetAttribute ("Name")) 
spk.Voice
spk.SetVoice(vcs.Item(speaker_number))

spk.Speak(file)


