import os
import random
import win32com.client as wincom


speak = wincom.Dispatch("SAPI.spVoice")
print("Let's Play Rock Paper Scissor")
s1 = "Lets play rock paper scissor"
speak.Speak(s1)

l = ['rock','paper','scissor']
while True:
    ucount = 0
    ccount = 0 

    s2 = "Game Start enter 1 for start and 2 for exit"
    speak.Speak(s2)

    uchoice = int(input('''Game Start....   
    1. Yes
    2. No/Exit'''))

    if uchoice==1:
        for i in range(1,6):
            s3 = " enter 1 for rock  2 for paper 3 for scissor"
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
            speak.Speak(s4)

            print("Your Choice" , userchoice)

            compchoice = random.choice(l)
            s5 = f"computers choice {compchoice}"
            speak.Speak(s5)

            print("Computer Choice", compchoice)

            if(userchoice==compchoice):
                print("game draw")
                s6 = "game draw"
                speak.Speak(s6)

                ucount=ucount+1
                ccount=ccount+1

            elif((userchoice=="rock" and compchoice=="scissor") or (userchoice=="paper" and compchoice=="rock") or (userchoice=="scissor" and compchoice=="paper")):
                print(" You Win!!")
                s7 = " You Win!!"
                speak.Speak(s7)

                ucount = ucount + 1

            else:
                print("Computer Win!!")
                s8 = "Computer Win!!"
                speak.Speak(s8)

                ccount = ccount + 1

        print("Your Score is ",ucount)
        s9 = f"Your Score is {ucount}"
        speak.Speak(s9)

        print("Computer Score is " ,ccount)
        s10 = f"Computer Score is {ccount}"
        speak.Speak(s10)

        if(ucount==ccount):
            print("Fimal Game is Draw..")
            s11 = "Fimal Game is Draw.."
            speak.Speak(s11)

        elif(ucount>ccount):
            print("Congratulations You win the final game..")
            s12 = "Congratulations You win the final game.."
            speak.Speak(s12)

        else:
            print("Computer win the final game")
            s13 = "Computer win the final game"
            speak.Speak(s13)

print("from dev branch git")
