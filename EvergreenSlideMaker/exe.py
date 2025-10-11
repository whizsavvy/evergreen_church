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

hymn_list = ['주가 일하시네', '영광의 주님 찬양하세', '만세 반석', '주 임재 안에서', '나의 한숨을 바꾸셨네', '나의 갈 길 다 가도록', '너는 내 아들이라']

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

    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "8:1", "8:4")
    add_subtitle_slide(prs, input_text="고난의 역할 (사도행전 8:1~4)")
    
    add_bible_slide(prs, directory, "사도행전", "8:1")
    add_bible_slide(prs, directory, "사도행전", "8:2")
    add_bible_slide(prs, directory, "사도행전", "8:3")
    add_bible_slide(prs, directory, "사도행전", "8:4")
    add_bible_slide(prs, directory, "사도행전", "8:8")
    
    add_bible_slide(prs, directory, "시편", "23:4")
    add_bible_slide(prs, directory, "로마서", "5:3", "5:4")
    add_bible_slide(prs, directory, "시편", "119:71")
    add_bible_slide(prs, directory, "욥기", "23:10")
    add_bible_slide(prs, directory, "신명기", "32:11")
    add_bible_slide(prs, directory, "에베소서", "5:14")
    add_bible_slide(prs, directory, "이사야", "40:31")
    

    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    # add_card_slide(prs, input_text= '성찬')

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '오늘 숨을 쉬는 것 감사')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
