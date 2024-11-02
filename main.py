#프로그램명 : pdf변환기
#작성자 : 최윤호
#최종 수정일 : 2024_11_02

import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(2)
import PIL
import tkinter as tk
from tkinter import filedialog
import time
from img2pdf import convert
import random
import os
import shutil
import keyboard
import pydirectinput as pyd
import pyautogui
from tkinter import messagebox

window= tk.Tk()
window.title("PDF 변환기")
window.resizable(False,False)

intro_label=tk.Label(window, text="PDF변환기 입니다.",font=("Arial",20,"bold"),height=2)
intro_label.grid(row=0,column=0,columnspan=4,sticky='news')


#예제 메세지 삭제
def temp_text(e):
    if e.widget == left_input:
        left_input.delete(0, "end")
    elif e.widget == right_input:
        right_input.delete(0, "end")
    elif e.widget== input:
        input.delete(0, "end")


#마우스 좌표 저장 변수
global x1
global y1
global x2
global y2
x1=x2=y1=y2=-1
def get_pointer_coo(n):
    
    global x1
    global y1
    global x2
    global y2
    while(True):
        if(keyboard.is_pressed("s")):
            break
    if(n==0):
        x1,y1=pyd.position()
        adress_label1.configure(text=("왼쪽 상단 좌표 : "+str(x1)+" , "+str(y1)))
    else :
        x2,y2=pyd.position()
        adress_label2.configure(text=("오른쪽 하단 좌표 : "+str(x2)+" , "+str(y2)))   
        
def get_pointer_pass(n,input):
    global x1
    global y1
    global x2
    global y2
    if(n==0):
        x1,y1=(str(input).split(","))
        adress_label1.configure(text=("왼쪽 상단 좌표 : "+x1+" , "+y1))
    else:
        x2,y2=(str(input).split(","))
        adress_label2.configure(text=("오른쪽 하단 좌표 : "+x2+" , "+y2))
    
    x1=int(x1)
    x2=int(x2)
    y1=int(y1)
    y1=int(y2)

def fullscreen_adress():
    global x1
    global y1
    global x2
    global y2
    if(full_screen.get()):
        width, height = pyautogui.size()
        x1=y1=0
        x2=width
        y2=height    
    adress_label1.configure(text=("왼쪽 상단 좌표 : "+str(x1)+" , "+str(y1)))
    adress_label2.configure(text=("오른쪽 하단 : "+str(x2)+" , "+str(y2)))


#오른쪽 좌표
right_coordinate=tk.Button(window,text="오른쪽 하단 좌표 감지",command=lambda:get_pointer_coo(1))
right_g=tk.Label(window,text="좌표입력 :")
right_input=tk.Entry(window,width=10)
right_input_b=tk.Button(window,text="확인",command=lambda:get_pointer_pass(1,right_input.get()))

#왼쪽 좌표        
left_coordinate=tk.Button(window,text="왼쪽 상단 좌표 감지",command=lambda:get_pointer_coo(0))
left_g=tk.Label(window,text="좌표입력 :")
left_input=tk.Entry(window,width=10)
left_input_b=tk.Button(window,text="확인",command=lambda:get_pointer_pass(0,left_input.get()))


#예제 텍스트 삽입
left_input.insert(0, "ex)10,10")
right_input.insert(0, "ex)10,10")

#예제 텍스트 삭제
left_input.bind("<FocusIn>", temp_text)
right_input.bind("<FocusIn>", temp_text)

#좌표값 출력
adress_label1=tk.Label(window,text="왼쪽 상단 좌표 : ",pady=5)
adress_label2=tk.Label(window,text="오른쪽 하단 좌표 : ",pady=5)

######파일 주소 저장 GUI
def fileClick():
    global file_name
    file_name=window.file=filedialog.askdirectory(
        title="파일 선택창",
        )
    file_label.configure(text=file_name)


file_bu=tk.Button(window,text="저장할 파일 선택",width=20,command=lambda:fileClick())
file_label = tk.Label(window,wraplength=480)

########################


######페이지 입력 GUI
def confirm():
    global page
    page=int(input.get())
    input_label.configure(text=("page:"+str(page)))

input= tk.Entry(window,width=10)
input.insert(0,"자동모드시 입력")
input.bind("<FocusIn>", temp_text)
input_label=tk.Label(window,text="page:")
input_button=tk.Button(window,text="확인",command=lambda:confirm())
###########

#######모드 체크박스
auto=tk.IntVar()
auto_check=tk.Checkbutton(window,text="자동 화면 전환",variable=auto,font=("Arial",10),height=2)
#############

######전체화면 캡쳐 모드 
full_screen=tk.IntVar()
full_screen_check=tk.Checkbutton(window,text="전체 화면 캡쳐",variable=full_screen,font=("Arial",10),height=2,command=lambda:fullscreen_adress())

######캡쳐 장수 모드 설정 
capchar_page=tk.IntVar()
capchar_page_check=tk.Checkbutton(window,text="한장 캡쳐",variable=capchar_page,font=("Arial",10),height=2)


###캡쳐 함수
def capchar(i,mode):#mode 0은 두장모드
    if(mode==0):#두장모드
        screenshot1 = PIL.ImageGrab.grab(bbox=(x1,y1,x1+((x2-x1)//2),y2))
        screenshot1 = PIL.ImageGrab.grab(bbox=(x1,y1,x1+((x2-x1)//2),y2))
        screenshot1.save(file_name+"/png"+f"/screenshot{i}.png")
        screenshot2 = PIL.ImageGrab.grab(bbox=(x1+((x2-x1)//2),y1,x2,y2))
        screenshot2 = PIL.ImageGrab.grab(bbox=(x1+((x2-x1)//2),y1,x2,y2))
        screenshot2.save(file_name+"/png"+f"/screenshot{i+1}.png")
        #insert
        image_list.append(file_name+"/png"+f"/screenshot{i}.png")
        image_list.append(file_name+"/png"+f"/screenshot{i+1}.png")
    else:
        screenshot1 = PIL.ImageGrab.grab(bbox=(x1,y1,x2,y2))
        screenshot1.save(file_name+"/png"+f"/screenshot{i}.png")
        #insert
        image_list.append(file_name+"/png"+f"/screenshot{i}.png")

#mode = 1 은 수동 
def action():

    if( x1==-1 or x2==-1 or y1 ==-1 or y2==-1 or y1>=y2 or x1>=x2):
        messagebox.showerror("오류","좌표설정을 다시해주세요")
        return
        
    try:
        
        if(file_name==None):
            tk.messagebox.showinfo("제목", "여기에 메시지를 입력하세요.")
        
        if(os.path.exists(file_name+"/png")):
            shutil.rmtree(file_name+"/png")
        global image_list
        image_list=[]
    
    
        print("변환시작")

        os.mkdir(file_name+"/png")
        capchar_page_mode=capchar_page.get()
        if(auto.get()==0):
            print("수동모드 진입")
            i=0
            while(True):
                if(keyboard.is_pressed("right")):
                    time.sleep(1)
                    if(capchar_page_mode):
                        capchar(i,capchar_page_mode) 
                        i+=2
                    else:
                        capchar(i,capchar_page_mode) 
                        i+=1

                elif(keyboard.is_pressed("Escape")):
                    break
        else:
            print("자동모드 진입")
            while(True):
                if(keyboard.is_pressed("s")):
                    print("캡쳐시작")
                    break
            if(capchar_page_mode):
                for j in range(0,page+1,2):
                    if(keyboard.is_pressed("Escape")):
                        break
                    time.sleep(random.uniform(1,1.8))
                    pyd.press("right")
                    capchar(j, capchar_page_mode)
            else:
                for j in range(0,page+1,1):
                    if(keyboard.is_pressed("Escape")):
                        break
                    time.sleep(random.uniform(1,1.8))
                    pyd.press("right")
                    capchar(j, capchar_page_mode)

        print("캡쳐 종료")
        with open(file_name+"/out.pdf","wb") as f:
            pdf = convert(image_list)
            f.write(pdf)
        shutil.rmtree(file_name+"/png")
    except NameError:
        messagebox.showerror("오류","저장할 파일을 지정해주세요")

        

######변환 시작 버튼
action_button=tk.Button(window,text="변환 시작",bg="gray",height=3,command=lambda:action())

#왼쪽 좌표 설정 버튼 및 라벨
left_coordinate.grid(row=1,column=0,sticky='news')
left_g.grid(row=1,column=1,sticky='news')
left_input.grid(row=1,column=2,sticky='news')
left_input_b.grid(row=1,column=3,sticky='news')
adress_label1.grid(row=2,column=0,columnspan=3,sticky='news')

#오른쪽 좌표 설정 버튼 및 라벨
right_coordinate.grid(row=3,column=0,sticky='news')
right_g.grid(row=3,column=1,sticky='news')
right_input.grid(row=3,column=2,sticky='news')
right_input_b.grid(row=3,column=3,sticky='news')
adress_label2.grid(row=4,column=0,columnspan=3)

#저장 경로 설정 버튼 및 라벨
file_bu.grid(row=5,column=0,sticky='news')
file_label.grid(row=5,column=1,columnspan=4,sticky='news')

#전체화면 캡쳐 모드 체크박스
full_screen_check.grid(row=6,column=0,sticky='news')

#화면 전환 모드 체크박스
auto_check.grid(row=6,column=1,sticky='news')

#캡쳐 한장 모드 체크박스 
capchar_page_check.grid(row=6,column=2,sticky='news')

#페이지 수 설정 버튼 및 라벨
input_label.grid(row=7,column=0,sticky='news')
input.grid(row=7,column=1,sticky='news')
input_button.grid(row=7,column=2,sticky='news')

#변환 시작 버튼
action_button.grid(row=8,column=0,columnspan=4,sticky='news')

window.mainloop()
