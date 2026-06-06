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

hymn_list = ['주를 찾는 모든 자들이', '주 하나님 지으신 (아이자야)', '내 마음 다해', '주님은 나의 힘이요', '나 주님의 기쁨 되기 원하네']

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
    # add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "로마서", "14:1", "14:12")
    add_subtitle_slide(prs, input_text="주를 위하여 (로마서 14:1~12)")
    
    add_bible_slide(prs, directory, "로마서", "15:7")
    add_bible_slide(prs, directory, "사도행전", "2:44", "2:47")
    add_bible_slide(prs, directory, "로마서", "14:6")
    add_bible_slide(prs, directory, "갈라디아서", "1:10")
    add_bible_slide(prs, directory, "고린도전서", "10:31")
    add_bible_slide(prs, directory, "로마서", "14:7", "14:8")
    add_bible_slide(prs, directory, "빌립보서", "2:4")
    add_bible_slide(prs, directory, "로마서", "14:9")
    add_bible_slide(prs, directory, "로마서", "14:10")
    add_bible_slide(prs, directory, "이사야", "45:23")

    add_hymn_slide(prs, hymn_list[4])

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
