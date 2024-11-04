
# -*- coding: utf-8 -*-
"""
Created on Thu Jun 30 10:07:59 2022

@author: gmartinsrodrigues
"""

import pyautogui
import time


def cordenadas_do_mouse():
    
    cordenadas_atuais = [0, 0]
    cordenadas_anteriores = [0, 0]
    
    while(True):
        
        time.sleep(0.1)
        cordenadas_atuais[0], cordenadas_atuais[1] = pyautogui.position()
        
        if cordenadas_anteriores[0] != cordenadas_atuais[0] or cordenadas_anteriores[1] != cordenadas_atuais[1]:
            
            print("x: ", cordenadas_atuais[0], " , y: ", cordenadas_atuais[1])
            
            cordenadas_anteriores[0] = cordenadas_atuais[0]
            cordenadas_anteriores[1] = cordenadas_atuais[1]


if __name__ == '__main__':
    cordenadas_do_mouse()
