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

hymn_list = ['복의 근원 강림하사', '꽃들도', '구주 예수 의지함이', '예수 십자가에 흘린 피로써', '주 예수의 이름 높이세', '충만', '부름 받아 나선 이 몸']

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
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "사도행전", "8:34", "8:40")
    add_subtitle_slide(prs, input_text="광야에서 만난 빌립과 에디오피아 내시 (사도행전 8:34~40)")
    
    # RED only, 원고 순서
    add_bible_slide(prs, directory, "마태복음", "24:12")
    add_bible_slide(prs, directory, "마태복음", "24:6", "24:8")
    add_bible_slide(prs, directory, "마태복음", "24:14")
    
    add_bible_slide(prs, directory, "사도행전", "8:26")
    add_bible_slide(prs, directory, "사도행전", "8:29")
    
    add_bible_slide(prs, directory, "창세기", "6:22")
    add_bible_slide(prs, directory, "누가복음", "5:5", "5:6")
    
    add_bible_slide(prs, directory, "히브리서", "11:6")
    add_bible_slide(prs, directory, "요한계시록", "3:15")
    add_bible_slide(prs, directory, "잠언", "18:12")
    add_bible_slide(prs, directory, "야고보서", "4:6")
    
    add_bible_slide(prs, directory, "사도행전", "8:30", "8:31")
    add_bible_slide(prs, directory, "사도행전", "8:36")
    
    add_bible_slide(prs, directory, "전도서", "3:1")
    
    add_bible_slide(prs, directory, "사도행전", "8:35")
    add_bible_slide(prs, directory, "사도행전", "8:39")
    

    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs, hymn_list[6])
    # add_card_slide(prs, input_text= '성찬')

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '오늘 숨을 쉬는 것 감사')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
