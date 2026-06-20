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

hymn_list = ['부르신 곳에서', '그 참혹한 십자가에', ' 오 주여 나의 마음이', '살아계신 주', '하나님의 부르심']

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
    
    
    add_bible_slide(prs, directory, "에베소서", "4:12")
    add_subtitle_slide(prs, input_text="나를 향한 하나님의 부르심 (에베소서 4:12)")
    
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "마태복음", "4:21")
    add_bible_slide(prs, directory, "시편", "147:3")
    add_bible_slide(prs, directory, "마가복음", "10:45")
    add_bible_slide(prs, directory, "베드로전서", "4:10")
    add_bible_slide(prs, directory, "베드로전서", "2:5")
    add_bible_slide(prs, directory, "고린도전서", "12:26")

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
