# Import the win32com.client module to use the Windows SAPI interface
import win32com.client as wincl

# Create a speech synthesis object
speak = wincl.Dispatch("SAPI.SpVoice")

# List of strings to be spoken
l = ["Rahul", "Nishant", "Harry"]

# Use the Speak method to vocalize each string in the list
for item in l:
  speak.Speak(f"Hello {item}")
