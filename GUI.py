import gui
import os
window = gui.Window("Dental Doc Convertor", layout="pack", master = None)
LSDbtn = window.add_button("          Lostus Smile Dental            ")
JMWRFbtn = window.add_button("         Jvon Clinic Walking Registration Form            ")
PASbtn = window.add_button("         Jvon Clinic Patient A Rules           ")

def LSD():
    os.system('python LSD.py')

def JM_WRF():
    os.system('python JM-WRF.py')

def PAS():
    os.system('python PAS.py')

gui.on("btnPress", LSDbtn, LSD)
gui.on("btnPress", JMWRFbtn, JM_WRF)
gui.on("btnPress", PASbtn, PAS)
#If you don't know what lambda is, it just returns an anonymous, single expression function.
window.start()