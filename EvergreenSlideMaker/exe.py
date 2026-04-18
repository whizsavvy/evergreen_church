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

hymn_list = ['예수 이름 높이세', '태산을 넘어 험곡에 가도', '나의 하나님', '주님의 임재 앞에서']

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
    # add_hymn_slide(prs, hymn_list[1])
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_blank_slide(prs)
    

    # add_card_slide(prs, input_text= '성가대 찬양')
    # add_bible_slide(prs, directory, "시편", "46:1", "46:11")
    # add_subtitle_slide(prs, input_text="하나님께 집중하다 – 가만히 있으라 (시편 46:1~11)")
    
    add_bible_slide(prs, directory, "마태복음", "5:13", "5:16")
    add_subtitle_slide(prs, input_text="복음은 살아내는 것입니다 (마태복음 5:13~16)")
    
    add_bible_slide(prs, directory, "마태복음", "28:19", "28:20")
    add_bible_slide(prs, directory, "베드로전서", "2:9")
    add_bible_slide(prs, directory, "마태복음", "5:13")
    add_bible_slide(prs, directory, "레위기", "2:13")
    add_bible_slide(prs, directory, "에베소서", "5:8")
    add_bible_slide(prs, directory, "고린도후서", "3:2", "3:3")
    add_bible_slide(prs, directory, "마태복음", "5:16")
    add_bible_slide(prs, directory, "베드로전서", "3:15")
    add_bible_slide(prs, directory, "요한복음", "8:12")

    # add_hymn_slide(prs, hymn_list[5])
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '주 하나님 지으신 모든 세계')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
