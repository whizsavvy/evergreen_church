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

hymn_list = ['예수 예수','하늘 위에 주님 밖에', '주 앙망하는자', '주님은 나의 힘이요', '아름답고 놀라운 주 예수']

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
    # add_hymn_slide(prs, hymn_list[5])
    # add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "7:54", "7:60")
    add_subtitle_slide(prs, input_text="돌을 던지는 세상, 무릎 꿇는 성도 (사도행전 7:54~60)")
    
    add_bible_slide(prs, directory, "디모데후서", "4:7")
    add_bible_slide(prs, directory, "이사야", "55:10")
    add_bible_slide(prs, directory, "여호수아", "24:2")
    add_bible_slide(prs, directory, "창세기", "12:1", "12:4")
    add_bible_slide(prs, directory, "창세기", "15:2", "15:5")
    add_bible_slide(prs, directory, "창세기", "17:4", "17:5")
    add_bible_slide(prs, directory, "창세기", "17:17", "17:19")
    add_bible_slide(prs, directory, "창세기", "21:5")
    add_bible_slide(prs, directory, "누가복음", "11:9")
    add_bible_slide(prs, directory, "시편", "46:5")
    add_bible_slide(prs, directory, "고린도전서", "1:21")
    add_bible_slide(prs, directory, "창세기", "39:2")
    add_bible_slide(prs, directory, "마태복음", "7:22")
    add_bible_slide(prs, directory, "잠언", "3:6")
    add_bible_slide(prs, directory, "시편", "91:14")
    add_bible_slide(prs, directory, "창세기", "45:4", "45:5")
    add_bible_slide(prs, directory, "창세기", "45:7", "45:8")
    add_bible_slide(prs, directory, "창세기", "45:14", "45:15")
    add_bible_slide(prs, directory, "마태복음", "10:17", "10:19")
    add_bible_slide(prs, directory, "요한복음", "2:19", "2:21")
    add_bible_slide(prs, directory, "사도행전", "7:56")
    add_bible_slide(prs, directory, "디모데전서", "4:4")
    add_hymn_slide(prs, '내 구주 예수를 더욱 사랑')


    # add_card_slide(prs, input_text= '성찬') 
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '우리 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
