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

hymn_list = ['저 높은 곳을 향하여', '슬픈 마음 있는 사람', '예수 열방의 소망', '오 나의 자비로운 주여', '말씀 앞에서']

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
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "출애굽기", "16:4")
    add_subtitle_slide(prs, input_text='광야에서 역사하시는 하나님말씀(출애굽기 16:4)')
    
    add_bible_slide(prs, directory, "아모스", "7:8")
    add_bible_slide(prs, directory, "고린도후서", "13:5")
    add_bible_slide(prs, directory, "시편", "119:105")
    add_bible_slide(prs, directory, "마태복음", "4:12")
    add_bible_slide(prs, directory, "출애굽기", "16:3")
    add_bible_slide(prs, directory, "출애굽기", "16:4")
    add_bible_slide(prs, directory, "요한복음", "6:33")
    add_bible_slide(prs, directory, "출애굽기", "16:4")
    add_bible_slide(prs, directory, "창세기", "6:22")
    add_bible_slide(prs, directory, "히브리서", "4:12")
    
    add_hymn_slide(prs, hymn_list[4])
                    
    


    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
