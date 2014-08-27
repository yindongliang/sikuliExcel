import sys,os,re
import shutil
import time as systime
import traceback
import java.lang
import string
import javax.management.Attribute
from java.util import *
from javax.management import *
from sikuli.Sikuli import getImagePath, addImagePath, setShowActions, setAutoWaitTimeout, getNumberScreens
from sikuli import Screen, Region
from sikuli.Sikuli import *
from org.apache.log4j import Logger
import UserInterruption

imgFolder=''
screenX=Screen(0)
enginlog=Logger.getLogger("testEngin")
def createImageFolder(foldername):
    imgPath = makeImgFolder(foldername)
    if os.path.isdir(imgPath):
        global imgFolder
        imgFolder = imgPath
    else:
        print "image folder created failed"
        exit(0)
def makeImgFolder(custom_foldername):
    
    currentTime = systime.strftime('%Y%m%d%H%M%S',systime.localtime(systime.time()))
    if custom_foldername =='':
        imgPath = ".\\img\\" + currentTime
    else:
        imgPath = ".\\img\\" +custom_foldername+"\\" +currentTime
    
    os.makedirs(r''+imgPath)
    enginlog.info(imgPath+" created")
    return imgPath


def saveScreen():
    currentTime = systime.strftime('%Y%m%d%H%M%S',systime.localtime(systime.time()))
    file = screenX.capture(screenX.getBounds())
    shutil.move(file, imgFolder+ "\\" + currentTime + ".png")

def extendExists(image,similarity=1):
    
    if similarity>0.8:
        similarity = similarity-0.01
    else:
        return False
    if not screenX.exists(Pattern(image).similar(similarity)):
        
        return extendExists(image,similarity)
    else:
        enginlog.info(image+" found ")
        return True

def extendWaitEvent(image,offsetDirect="left",offsetValue=0,event="click",inputContent="",timeOut=10,hotKey="",timeOutContinue=""):
    
    offsetValue=float(offsetValue);
    timeOut=int(float(timeOut))
    time = 0
    findFlg = 0
    if timeOut > 0:
        timeover = timeOut
    else:
        timeover = 10
    enginlog.info(image +" still in searching...,please wait paitiently")
    while time < timeover:
        if extendExists(image,1):
            findFlg = 1
            break
        else:
            time = time+1
            
    if findFlg != 1:
        if timeOutContinue=="":
            enginlog.info(image +" can not be detected until timeout,please check your steps.process stopped!") 
            exit(0)
        else:
            enginlog.info(image +" can not be detected until timeout,but continue")
            
    else:
        saveScreen()
        if not hotKey == "":
            eval("screenX.type("+hotKey+")")
        else:
            if event=="drag" or event=="dropAt":
                dragAndDropElement(event,offsetDirect,offsetValue)
                
            else:
                clickOrtypeContent(event,offsetDirect,offsetValue,inputContent)
                hover(screenX)
            
            

def clickOrtypeContent(eventType="click",offsetDirect="",offsetValue=0,inputContent=""):
    
    lastMatch = screenX.getLastMatch()
    if offsetValue == 0 or offsetDirect == "":
        click(lastMatch)
        
    else:
        offsetValue=int(offsetValue)
        click(eval("lastMatch."+offsetDirect+"(int('"+str(offsetValue)+"'))"))
        
    if eventType == "clicktypeClear":
        wait(2)
        clearContent()
        screenX.type(inputContent)
    if eventType == "clicktypeNoclear":
        wait(2)
        screenX.type(inputContent)
        
def dragAndDropElement(eventType="drag",offsetDirect="",offsetValue=0):
    
    lastMatch = screenX.getLastMatch()
    if offsetValue == 0 or offsetDirect == "":
        eval(eventType+"(lastMatch)")
        
    else:
        offsetValue=int(offsetValue)
        eval(eventType+"("+"lastMatch."+offsetDirect+"(int('"+str(offsetValue)+"'))"+")")
        
    
    
def clearContent():
    
    screenX.type("a",KEY_CTRL)
    screenX.type(Key.BACKSPACE)


def loadImages(path):
    
    current_files = os.listdir(path)
    
    for file_name in current_files:
        full_file_name = os.path.join(path, file_name)
        
        if os.path.isdir(full_file_name):
            addImagePath(full_file_name)
            loadImages(full_file_name)
            
def main():
    # A small value such as 6 makes the matching algorithm be faster.
    #Vision.setParameter("MinTargetSize",6)
    # A large value such as 18 makes the matching algorithm be more robust.
    #Vision.setParameter("MinTargetSize",18)
    #Load configuration
    #currentDir =  str(os.path.abspath(sys.argv[0]))
    #currentDir = os.getcwd()
    loadImages(".\\image\\")
main()