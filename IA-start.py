from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
import glob
import cv2
import numpy as np
import os
from matplotlib import pyplot as plt
from pathlib import Path
import time

def go(img):
    x,y = img.shape
    if x<=4 or y<=4:
        return ""
    cnt = [0]*100
    val0 = [0]*100
    val1 = [0]*100
    mylist = [f for f in glob.glob('templates\\*\\*.png')]
    for i in mylist:
        img2 = img.copy()
        template = cv2.imread(i,0)
        w, h = template.shape[::-1]
        meth = 'cv2.TM_CCOEFF_NORMED'
        method = eval(meth)
        max_val = -1.0
        if img2.shape[0]<template.shape[0] or img2.shape[1]<template.shape[1]:
            template = cv2.resize(template,(min(template.shape[1],img2.shape[1]),min(template.shape[0],img2.shape[0])))
            w,h= template.shape[::-1]

        res = cv2.matchTemplate(img2,template,method)
        min_val, max_val, min_loc, max_loc = cv2.minMaxLoc(res)
        top_left = max_loc
        bottom_right = (top_left[0] + w, top_left[1] + h)

        h = int(ord(i[10]))
        if int(ord(i[10]))<58:
            if cnt[h-22] < max_val:
                cnt[h-22] = max_val
                val0[h-22] = (top_left[0])
                val1[h-22] = (bottom_right[0])
        else:
            if cnt[h-97] < max_val:
                cnt[h-97] = max_val
                val0[h-97] = (top_left[0])
                val1[h-97] = (bottom_right[0])
    rr = max(cnt)
    if rr<0.6666 :
        return ""
    ct=0
    point0=0
    point1=y-1
    for i in cnt:
        if rr==i:
            if ct<26:
                res=chr(ct+97)
                point0=val0[ct]
                point1=val1[ct]
            else:
                res=chr(ct+22)
                point0=val0[ct]
                point1=val1[ct]
            break;
        ct=ct+1
    R=""
    L=""
    if point0>0:
        imgL=img[:,:min(y-1,point0+1)]
        L=go(imgL)
    if point1<y-1:
        imgR=img[:,max(0,point1-1):]
        R=go(imgR)
    return L+res+R;

def descargar(driver,intentos):
    action = ActionChains(driver)
    #driver.find_element_by_xpath("//button[@id='tbBuscador:idFormBuscarProceso:btnExportar']").click()
    exp = driver.find_element_by_xpath("//button[@id='tbBuscador:idFormBuscarProceso:btnExportar']")
    action.move_to_element(exp)
    action.pause(1)
    exp2 = driver.find_element_by_xpath("//button[@id='tbBuscador:idFormBuscarProceso:btnExportar']")
    action.move_to_element(exp2)
    action.pause(1)
    exp3 = driver.find_element_by_xpath("//button[@id='tbBuscador:idFormBuscarProceso:btnExportar']")
    action.move_to_element(exp3).click()
    action.pause(1)
    action.perform()

    time.sleep(4)
    num_list = [f for f in glob.glob("C:\\Users\\DIEGO\\Downloads\\*.xls")]
    count = 0
    for i in num_list:
        if i[25:39] == "Lista-Procesos":
            count+=1
    if count==0 and intentos<3:
        time.sleep(2)
        descargar(driver,intentos+1)


def binarizar(imgC, io):
    imgR = imgC[227:270,450:695]
    cv2.imwrite(io,imgR)

    img1 = cv2.imread(io,0)
    imgT = cv2.resize(img1,(201,35))
    ret,thresh1 = cv2.threshold(imgT,127,255,cv2.THRESH_TRUNC)
    plt.imsave(io,thresh1, 0, 45,'gray')
    img2 = cv2.imread(io,0)
    ret,th1 = cv2.threshold(img2,120,255,cv2.THRESH_BINARY_INV)
    plt.imsave(io,th1, 0, 255,'gray')
    img3 = cv2.imread(io,0)

    xx,yy = imgT.shape
    by1=-1
    by2=-1
    for ii in range(yy):
        for jj in range(xx):
            if img3[jj][ii]==255 and by1<0:
                by1=ii
            if img3[jj][ii]==255:
                by2=ii
    img4=img3[:,by1:by2]
    cv2.imwrite(io,img4)

def refresh_captcha(driver):
    action2 = ActionChains(driver)
    refresh_captcha = driver.find_element_by_xpath("//button[@id='tbBuscador:idFormBuscarProceso:btnrefreshcaptcha']")
    action2.move_to_element(refresh_captcha).click()
    action2.pause(4)
    action2.perform()

def deletefiles():
    num_list = [f for f in glob.glob("C:\\Users\\DIEGO\\Downloads\\*.xls")]
    for i in num_list:
        if i[25:39] == "Lista-Procesos":
            os.remove(i)

driver = webdriver.Chrome()
driver.get("http://prodapp2.seace.gob.pe/seacebus-uiwd-pub/buscadorPublico/buscadorPublico.xhtml")
driver.maximize_window()

time.sleep(2)
refresh_captcha(driver)

driver.save_screenshot('a.png')
imgC = cv2.imread('a.png',0)
io = 'b.png'

binarizar(imgC,io)

imgSeace = cv2.imread(io,0)
captcha = go(imgSeace)

elem = driver.find_element_by_id("tbBuscador:idFormBuscarProceso:codigoCaptcha")
elem.send_keys(captcha)
elem.send_keys(Keys.RETURN)

driver.save_screenshot('c.png')
time.sleep(6)
driver.save_screenshot('d.png')
deletefiles()
descargar(driver,0)

#driver.close()
