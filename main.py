#import libraries
import cv2
import mediapipe as mp
import numpy as np
from mediapipe.framework.formats import landmark_pb2
import time
from math import sqrt
import win32api, win32con
import math
import time
import pyautogui
from ctypes import cast, POINTER
from comtypes import CLSCTX_ALL
from pycaw.pycaw import AudioUtilities, IAudioEndpointVolume
import speech_recognition as sr


#Initializing speech and voice recognition modules
listener = sr.Recognizer()

#GUI settings
pyautogui.FAILSAFE = False

#initialize and detect speakers/audio devices
devices = AudioUtilities.GetSpeakers()
interface = devices.Activate(IAudioEndpointVolume._iid_, CLSCTX_ALL, None)
asdf = cast(interface, POINTER(IAudioEndpointVolume))

#get volume levels and set a scale from 0 to 1
volume_level = asdf.GetMasterVolumeLevel()
volRange = volume = asdf.GetVolumeRange()
minvol = volRange[0]
maxvol = volRange[1]


#mask/drawing settings for mediapipe landmarks
landmark_draw = mp.solutions.drawing_utils
mp_hands = mp.solutions.hands
click = 0
show_im = False
drawSpec1 = landmark_draw.DrawingSpec(thickness = 2, circle_radius = 1)

#webcam input
video = cv2.VideoCapture(0)



#get hand position with landmarks (each joint on each finger)
def findPosition(image, handNo=0):
    landmark_list = []
    if hndlms.multi_hand_landmarks:
        Hand = hndlms.multi_hand_landmarks[handNo]
        for id, lm in enumerate(Hand.landmark):  
            h, w, c = image.shape
            cx, cy = int(lm.x * w), int(lm.y * h)
            landmark_list.append([id, cx, cy])
    return landmark_list


#left click function with win32 api        
def leftClick():
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTDOWN,0,0)
        time.sleep(0.3)
        win32api.mouse_event(win32con.MOUSEEVENTF_LEFTUP,0,0)

#right click with win32 api
def rightClick():
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTDOWN,0,0)
    time.sleep(0.3)
    win32api.mouse_event(win32con.MOUSEEVENTF_RIGHTUP,0,0)
    
    
#speak to type function
def speak_to_type():
    data = None
    with sr.Microphone() as source:
            try:
                voice = listener.listen(source)
                data = listener.recognize_google(voice)
                data.lower()
                data.strip()
                pyautogui.write(data, interval=0.1)
            except: 
                pass
            
    


#switching tabs with pyautogui
def alt_tab():
    pyautogui.keyDown("alt")
    time.sleep(0.2)
    pyautogui.keyDown("tab")    
    time.sleep(0.4)
    pyautogui.keyUp("alt")
    pyautogui.keyUp("tab")

#get hands
with mp_hands.Hands() as hands:
    #video frame preprocessing
    while video.isOpened():
        _, frame = video.read()
        image = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        image = cv2.flip(image, 1)
        imageHeight, image_Width, _ = image.shape
        
        #Get landmarks/positions for each hand per frame
        hndlms = hands.process(image)
        landmark_list = findPosition(image)
        
        #draw handlandmarks
        if hndlms.multi_hand_landmarks:
            for num, hand in enumerate(hndlms.multi_hand_landmarks):
                landmark_draw.draw_landmarks(image, hand, mp_hands.HAND_CONNECTIONS, landmark_draw.DrawingSpec(color = (50, 255, 50), thickness = 2, circle_radius = 2))
        
        #Check if hand is left or right
        try:
            for hand in hndlms.multi_handedness:
                handType=hand.classification[0].label
        except: 
            continue
        
        #iterate through each point/joint/landmark
        if hndlms.multi_hand_landmarks != None:
            for handLandmarks in hndlms.multi_hand_landmarks:
                for point in mp_hands.HandLandmark:
                    normalizedLandmark = handLandmarks.landmark[point]
                    pixelCoordinatesLandmark = landmark_draw._normalized_to_pixel_coordinates(normalizedLandmark.x, normalizedLandmark.y, image_Width, imageHeight)
                    point = str(point)
                    
                    #control cursor with right hand index finger tip   
                    if handType == "Right":
                        if point == 'HandLandmark.INDEX_FINGER_TIP':
                            try:
                                indexfingertip_x = pixelCoordinatesLandmark[0]
                                indexfingertip_y =  pixelCoordinatesLandmark[1]
                                win32api.SetCursorPos((indexfingertip_x*6, indexfingertip_y*7))
                                          
                            except:
                                pass
                        #set coordinates for the right hand thumb to use later
                        elif point == 'HandLandmark.THUMB_TIP':
                            try:
                                thumbfingertip_x = pixelCoordinatesLandmark[0]
                                thumbfingertip_y =  pixelCoordinatesLandmark[1]
                            except:
                                pass
                        
                                   
                        
                        
                if handType == "Right": 
                    for handLms in hndlms.multi_hand_landmarks:
                        for id, lm in enumerate(handLms.landmark):

                            h, w, c,  = image.shape
                            cx, cy = int(lm.x*w), int(lm.y*h)

                            
                            #set line values to be used for index finger thumb middle finger pinky ring finger
                            if id == 4:
                                line_val1 = (cx, cy)
                                cv2.circle(image, (cx, cy), 3, (255, 0, 255), cv2.FILLED)

                            if id == 12:
                                line_val2 = (cx, cy)

                            if id == 20:
                                line_val3 = (cx, cy)
                                cv2.circle(image, (cx, cy), 3, (255, 0, 255), cv2.FILLED)
                            
                            if id == 16:
                                line_val4 = (cx,cy)
                            

                    cv2.line(image, line_val1, line_val2, (190,190,190), 2)
                    cv2.line(image, line_val1, line_val3, (190,190,190), 2)
                    cv2.line(image, line_val1, line_val4, (190,190,190), 2)
                    
            if handType == "Right":
                #get line lengths in relavance to the current frame
                lx, ly = int((line_val1[0] + line_val2[0])/2), int((line_val1[1] + line_val2[1])/2)
                cv2.circle(image, (lx, ly), 3, (255, 0, 255), cv2.FILLED)

                length_line = math.hypot( line_val2[0] - line_val1[0],line_val2[1] - line_val1[1])
                length_line1 = math.hypot( line_val3[0] - line_val1[0],line_val3[1] - line_val1[1])
                length_line2 = math.hypot( line_val4[0] - line_val1[0],line_val4[1] - line_val1[1])

                #left click by pinching middle finger and thumb
                if int(length_line) < 15:
                        leftClick()
                #right click by pinching pinky and thumb
                if int(length_line1) < 15:
                        rightClick()
                        
                if int(length_line2) < 15:
                        speak_to_type()

                #switch tabs on swipe right
                if landmark_list[8][1] > 600:
                       alt_tab()
                

        #left hand
        if handType == "Left":
         if hndlms.multi_hand_landmarks:
          for handLms in hndlms.multi_hand_landmarks:
            for id, lm in enumerate(handLms.landmark):
                #set line variables to use alter
                lineval_1 = None
                lineval_2 = None
                h, w, c,  = frame.shape
                cx, cy = int(lm.x*w), int(lm.y*h)
                #get positions of index and thumb fingers
                if id == 4:
                    line_val1 = (cx, cy)
                    cv2.circle(frame, (cx, cy), 3, (255, 0, 255), cv2.FILLED)
                if id == 8:
                    line_val2 = (cx, cy)
                    cv2.circle(frame, (cx, cy), 3, (255, 0, 255), cv2.FILLED)
            
            #draw line between index and thumb finger
            cv2.line(frame, line_val1, line_val2, (190, 190, 190), 2)
            
            
            #get the length of the line in relevance to the frame
            lx, ly = int((line_val1[0] + line_val2[0])/2), int((line_val1[1] + line_val2[1])/2)
            cv2.circle(frame, (lx, ly), 3, (255, 0, 255), cv2.FILLED)
            
            length_line = math.hypot( line_val2[0] - line_val1[0],line_val2[1] - line_val1[1])
            
            #change volume based on the length of the line
            vol_1 = np.interp(length_line, [11, 150], [minvol, maxvol])
            asdf.SetMasterVolumeLevel(vol_1, None)
            
            #draw landmarks
            landmark_draw.draw_landmarks(frame, handLms, mp_hands.HAND_CONNECTIONS, drawSpec1, drawSpec1)
            
            
        if show_im:
            cv2.imshow("img", image)
            cv2.waitKey(1)


 

 
