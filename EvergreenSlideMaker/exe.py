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

hymn_list = ['세상의 유혹 시험이', '슬픈 마음 있는 사람', '내 영혼이 은총 입어', '태산을 넘어 험곡에 가도', '크신 내 주님']

def create_presentation(hymn_list=[]):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)
    directory = folder_path+"/bible"
    pic_dic = folder_path+"/image/"
    add_image_slide(prs, pic_dic+'2026.png', text='주일 1부 예배')
    add_image_slide(prs, pic_dic+'2026.png', text='주일 2부 예배')

    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[0])
    add_hymn_slide(prs, hymn_list[1])
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[6])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "고린도후서", "2:15")
    add_subtitle_slide(prs, input_text="당신은 그리스도의 향기입니다 (고린도후서 2:15)")
    
    add_bible_slide(prs, directory, "고린도전서", "1:21")
    add_bible_slide(prs, directory, "고린도후서", "2:15")
    add_bible_slide(prs, directory, "고린도후서", "3:3")
    add_bible_slide(prs, directory, "로마서", "9:2")
    add_bible_slide(prs, directory, "로마서", "9:3")
    add_bible_slide(prs, directory, "고린도후서", "5:14")
    add_bible_slide(prs, directory, "누가복음", "15:4")
    add_bible_slide(prs, directory, "마태복음", "16:26")
    add_bible_slide(prs, directory, "요한복음", "1:45")
    add_bible_slide(prs, directory, "사도행전", "1:8")

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '그 날')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
