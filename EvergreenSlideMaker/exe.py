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

hymn_list = ['주 날 위해 버림 받음으로', '빈들에 마른 풀같이', '내 영혼이 은총 입어', '나는 주의 친구']

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
    
    add_bible_slide(prs, directory, "요한복음", "1:40", "1:42")
    add_subtitle_slide(prs, input_text="복음은 관계를 통해 흐릅니다 (요한복음 1:40~42)")
    
    add_bible_slide(prs, directory, "고린도전서", "1:21")
    add_bible_slide(prs, directory, "로마서", "10:14")
    add_bible_slide(prs, directory, "요한일서", "1:3")
    add_bible_slide(prs, directory, "요한복음", "4:28", "4:30")
    add_bible_slide(prs, directory, "누가복음", "14:23")
    add_bible_slide(prs, directory, "요한복음", "3:16")

    add_hymn_slide(prs, '하나님 아버지의 마음')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '주 하나님 지으신 모든 세계')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
