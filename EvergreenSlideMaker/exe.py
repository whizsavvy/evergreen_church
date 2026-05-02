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

hymn_list = ['거친 길 위를 걸어갈 때도', '이 험한 세상 나 살아 갈 동안', '마귀들과 싸울지라', '내가 매일 기쁘게', '나로부터 시작되리']

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
    add_blank_slide(prs)
    

    # add_card_slide(prs, input_text= '성가대 찬양')
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_subtitle_slide(prs, input_text="성령께서 하시니, 내가 간다 (사도행전 1:8)")
    
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_bible_slide(prs, directory, "스가랴", "4:6")
    add_bible_slide(prs, directory, "요한복음", "6:44")
    add_bible_slide(prs, directory, "고린도전서", "2:4")
    add_bible_slide(prs, directory, "요한복음", "9:25")
    add_bible_slide(prs, directory, "요한복음", "4:35")
    add_bible_slide(prs, directory, "고린도후서", "6:2")
    add_bible_slide(prs, directory, "로마서", "10:14")
    add_bible_slide(prs, directory, "누가복음", "11:13")

    # add_hymn_slide(prs, '하나님 아버지의 마음')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '그 날')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
