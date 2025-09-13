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

hymn_list = ['부흥', 'Again 1907', '일어나라 주의 백성', '예수 십자가에 흘린 피로써', '인애하신 구세주여', '무화과 나무 잎이 마르고', '오 신실하신 주', '나 같은 죄인 살리신']

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
    # add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "6:1", "6:4")
    add_subtitle_slide(prs, input_text="섬김으로 세워지는 공동체 (사도행전 6:1~4)")
    
    add_bible_slide(prs, directory, "사도행전", "6:1")
    add_bible_slide(prs, directory, "사도행전", "6:2")
    add_bible_slide(prs, directory, "디모데후서", "4:2")
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_bible_slide(prs, directory, "갈라디아서", "5:13")
    add_bible_slide(prs, directory, "사도행전", "6:3")
    add_bible_slide(prs, directory, "에베소서", "4:12")
    add_bible_slide(prs, directory, "사도행전", "6:4")
    add_bible_slide(prs, directory, "사도행전", "4:31")
    add_bible_slide(prs, directory, "사도행전", "6:7")
    add_bible_slide(prs, directory, "마가복음", "10:45")
    add_bible_slide(prs, directory, "느헤미야", "11:2")
    add_bible_slide(prs, directory, "느헤미야", "7:4")


    add_card_slide(prs, input_text= '성찬') 
    add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '우리 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
