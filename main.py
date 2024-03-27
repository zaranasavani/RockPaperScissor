import os
import random
import win32com.client as wincom


speak = wincom.Dispatch("SAPI.spVoice")
print("Let's Play Rock Paper Scissor")
s1 = "Lets play rock paper scissor"
# command1 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s1}');"
# os.system(command1)
speak.Speak(s1)

l = ['rock','paper','scissor']
while True:
    ucount = 0
    ccount = 0

    s2 = "Game Start enter 1 for start and 2 for exit"
    # command2 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s2}');"
    # os.system(command2)
    speak.Speak(s2)

    uchoice = int(input('''Game Start....
    1. Yes
    2. No/Exit'''))

    if uchoice==1:
        for i in range(1,6):
            s3 = " enter 1 for rock  2 for paper 3 for scissor"
            # command3 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s3}');"
            # os.system(command3)
            speak.Speak(s3)

            userIP = int(input('''
            1.Rock
            2.Paper
            3.Scissor'''))

            if userIP==1:
                userchoice = "rock"
            elif userIP==2:
                userchoice = "paper"
            elif userIP==3:
                userchoice = "scissor"

            s4 = f"your choice {userchoice}"
            # command4 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s4}');"
            # os.system(command4)
            speak.Speak(s4)

            print("Your Choice" , userchoice)

            compchoice = random.choice(l)
            s5 = f"computers choice {compchoice}"
            # command5 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s5}');"
            # os.system(command5)
            speak.Speak(s5)

            print("Computer Choice", compchoice)

            if(userchoice==compchoice):
                print("game draw")
                s6 = "game draw"
                # command6 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s6}');"
                # os.system(command6)
                speak.Speak(s6)

                ucount=ucount+1
                ccount=ccount+1

            elif((userchoice=="rock" and compchoice=="scissor") or (userchoice=="paper" and compchoice=="rock") or (userchoice=="scissor" and compchoice=="paper")):
                print(" You Win!!")
                s7 = " You Win!!"
                # command7 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s7}');"
                # os.system(command7)
                speak.Speak(s7)

                ucount = ucount + 1

            else:
                print("Computer Win!!")
                s8 = "Computer Win!!"
                # command8 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s8}');"
                # os.system(command8)
                speak.Speak(s8)

                ccount = ccount + 1

        print("Your Score is ",ucount)
        s9 = f"Your Score is {ucount}"
        # command9 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s9}');"
        # os.system(command9)
        speak.Speak(s9)

        print("Computer Score is " ,ccount)
        s10 = f"Computer Score is {ccount}"
        # command10 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s10}');"
        # os.system(command10)
        speak.Speak(s10)

        if(ucount==ccount):
            print("Fimal Game is Draw..")
            s11 = "Fimal Game is Draw.."
            # command11 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s11}');"
            # os.system(command11)
            speak.Speak(s11)

        elif(ucount>ccount):
            print("Congratulations You win the final game..")
            s12 = "Congratulations You win the final game.."
            # command12 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s12}');"
            # os.system(command12)
            speak.Speak(s12)

        else:
            print("Computer win the final game")
            s13 = "Computer win the final game"
            # command13 = f"PowerShell -Command "f"Add-Type -AssemblyName System.Speech;(New-Object System.Speech.Synthesis.SpeechSynthesizer).Speak('{s13}');"
            # os.system(command13)
            speak.Speak(s13)

