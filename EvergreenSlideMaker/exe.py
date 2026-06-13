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

hymn_list = ['입례', '나 같은 죄인 살리신', '슬픈 마음 있는 사람', '주의 약속하신 말씀 위에서(B)', '빈 들에 마른 풀같이 / 성령님 오셔서']

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
    # add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "로마서", "14:17", "14:23")
    add_subtitle_slide(prs, input_text="하나님의 나라는 (로마서 14:17~23)")
    
    add_bible_slide(prs, directory, "사도행전", "18:2")
    add_bible_slide(prs, directory, "로마서", "14:2")
    add_bible_slide(prs, directory, "로마서", "14:5")
    add_bible_slide(prs, directory, "로마서", "14:17")
    add_bible_slide(prs, directory, "로마서", "14:18")
    add_bible_slide(prs, directory, "로마서", "14:19")
    add_bible_slide(prs, directory, "고린도전서", "1:12")
    add_bible_slide(prs, directory, "로마서", "14:20")
    add_bible_slide(prs, directory, "에베소서", "4:29")
    add_bible_slide(prs, directory, "로마서", "14:23")
    add_bible_slide(prs, directory, "에베소서", "1:23")
    add_bible_slide(prs, directory, "고린도전서", "10:32", "10:33")
    add_bible_slide(prs, directory, "요한복음", "13:14", "13:15")
    add_bible_slide(prs, directory, "요한복음", "13:17")

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '빛을 들고 세상으로')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
