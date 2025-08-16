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

hymn_list = ['나의 한숨을 바꾸셨네', '문들아 머리 들어라', '예수 열방의 소망', '영광의 주님 찬양하세', '목마른 사슴', '사랑하는 나의 아버지', '내 맘이 낙심되며']

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
    add_bible_slide(prs, directory, "사도행전", "5:1", "5:5")
    add_subtitle_slide(prs, input_text="성령 앞에 신실하라 (사도행전 5:1~5)")
    
    add_bible_slide(prs, directory, "사도행전", "4:32", "4:37")
    add_bible_slide(prs, directory, "사도행전", "4:36", "4:37")
    add_bible_slide(prs, directory, "사도행전", "5:3")
    add_bible_slide(prs, directory, "사도행전", "5:4")
    add_bible_slide(prs, directory, "요한복음", "8:44")
    add_bible_slide(prs, directory, "잠언", "12:22")
    add_bible_slide(prs, directory, "사도행전", "5:5")  # 다시 등장함
    add_bible_slide(prs, directory, "사도행전", "5:11")
    add_bible_slide(prs, directory, "여호수아", "7:1")
    add_bible_slide(prs, directory, "고린도전서", "3:17")
    add_bible_slide(prs, directory, "에베소서", "2:2")
    add_bible_slide(prs, directory, "사도행전", "5:7")
    add_bible_slide(prs, directory, "사도행전", "5:8")
    add_bible_slide(prs, directory, "시편", "101:6")
    add_bible_slide(prs, directory, "시편", "139:23", "139:24")


    
    
    add_hymn_slide(prs, hymn_list[6])

    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
