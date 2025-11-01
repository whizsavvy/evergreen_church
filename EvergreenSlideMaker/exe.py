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

hymn_list = ['내 구주 예수를 더욱 사랑', '나의 죄를 씻기는', '예수 십자가에 흘린 피로써', '예수는 내 힘이요', '송축해 내 영혼', '나 같은 죄인 살리신', '하나님의 약속']

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
    add_hymn_slide(prs, hymn_list[2]) 
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
        # 제목/본문 (필수)
    add_bible_slide(prs, directory, "사도행전", "10:28", "10:35")
    add_subtitle_slide(prs, input_text="복음의 경계를 넘어 – 베드로와 고넬료의 만남 (사도행전 10:28~35)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "사도행전", "10:34", "10:35")
    add_bible_slide(prs, directory, "신명기", "14:2")
    add_bible_slide(prs, directory, "사도행전", "10:15")
    add_bible_slide(prs, directory, "사도행전", "10:2")
    add_bible_slide(prs, directory, "사도행전", "10:4")
    add_bible_slide(prs, directory, "사도행전", "10:11", "10:12")
    add_bible_slide(prs, directory, "사도행전", "10:15")  # 재등장(강조)
    add_bible_slide(prs, directory, "로마서", "3:22")
    add_bible_slide(prs, directory, "사도행전", "10:24")
    add_bible_slide(prs, directory, "사도행전", "10:44", "10:45")
    add_bible_slide(prs, directory, "고린도전서", "12:13")
    add_bible_slide(prs, directory, "디모데전서", "2:4")
    

    # add_hymn_slide(prs, hymn_list[6])
    
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '오늘 숨을 쉬는 것 감사')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
