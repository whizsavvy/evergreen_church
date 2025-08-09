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

hymn_list = ['하나님의 은혜', '생명 주께 있네', '다와서 찬양해', '주의 이름 송축하리', '왕의 왕 주의 주', '선한 능력으로', '능력의 이름 예수', '세상의 유혹 시험이']

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
    add_bible_slide(prs, directory, "사도행전", "4:5", "4:12")
    add_subtitle_slide(prs, input_text='구원과 능력의 이름, 예수 그리스도 (사도행전 4:5~12)')
    
    add_bible_slide(prs, directory, "사도행전", "4:12")
    add_bible_slide(prs, directory, "에베소서", "2:8", "2:9")
    add_bible_slide(prs, directory, "요한복음", "14:6")
    add_bible_slide(prs, directory, "사도행전", "4:10")
    add_bible_slide(prs, directory, "마가복음", "16:17", "16:18")
    add_bible_slide(prs, directory, "히브리서", "13:8")
    add_bible_slide(prs, directory, "사도행전", "4:11")
    add_bible_slide(prs, directory, "에베소서", "2:20")
    add_bible_slide(prs, directory, "사도행전", "4:20")

    
    
    add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs, hymn_list[7])
    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
