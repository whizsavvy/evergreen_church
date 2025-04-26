from pptx import Presentation
from pptx.util import Pt, Cm
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_AUTO_SIZE  # Correct import location for MSO_AUTO_SIZE
from pptx.enum.text import MSO_ANCHOR
import re
import datetime

today = datetime.datetime.now().strftime('%Y-%m-%d')


exec(open("EvergreenSlideMaker/setting.py").read())

hymn_list = ['오늘 숨을 쉬는 것 감사', '예수를 나의 구주 삼고', '주님은 나의 힘이요', '하나님이 세상을 사랑하사', '우리 오늘 눈물로', '이 땅의 동과 서 남과 북']

def create_presentation(hymn_list=[]):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)
    directory = folder_path+"/bible"
    pic_dic = folder_path+"/image/"
    add_image_slide(prs, pic_dic+'2025.jpg', text='주일 1부 예배')
    add_image_slide(prs, pic_dic+'2025.jpg', text='주일 2부 예배')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[0])
    add_hymn_slide(prs, hymn_list[1])
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)


    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "요한복음", "3:16")
    add_subtitle_slide(prs, input_text='가장 큰 사랑의 선물 (요 3:16)')
    add_blank_slide(prs)
    add_subtitle_slide(prs, input_text='1. "하나님이 세상을 이처럼 사랑하사" – 조건 없는 사랑의 선언')
    add_bible_slide(prs, directory, "로마서", "5:8")
    add_bible_slide(prs, directory, "요한일서", "4:10")
    add_subtitle_slide(prs, input_text='2. "독생자를 주셨으니" – 생명을 내어준 선물')
    
    add_bible_slide(prs, directory, "요한일서", "5:11")
    add_bible_slide(prs, directory, "로마서", "6:23")

    add_subtitle_slide(prs, input_text='3. "믿는 자마다 멸망하지 않고 영생을 얻게 하려 하심이라" – 모든 사람에게 열려 있는 구원의 길')
    add_bible_slide(prs, directory, "베드로후서", "3:9")
    add_bible_slide(prs, directory, "로마서", "10:13")
    add_bible_slide(prs, directory, "마태복음", "28:19", "28:20")
  
    # add_hymn_slide(prs, '나는 주를 섬기는 것에 후회가 없습니다')
    # add_image_slide(prs)
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_hymn_slide(prs, hymn_list[5])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
