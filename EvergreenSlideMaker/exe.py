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

hymn_list = ['예수 우리들의 밝은 빛', '날마다', '내 마음 다해', '주께 가오니', '나의 하나님', '예수의 이름으로 나는 일어서리라', '충만' ]

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
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", 3:10", "3:10")
    add_subtitle_slide(prs, input_text='예수의 이름으로 일어나 걸으라 (행 3:1~10)')
    
    add_bible_slide(prs, directory, "히브리서", "4:12")
    add_bible_slide(prs, directory, "사도행전", "3:1")
    add_bible_slide(prs, directory, "사도행전", "3:2")
    add_bible_slide(prs, directory, "사도행전", "3:4")
    add_bible_slide(prs, directory, "열왕기하", "6:16", "6:17")
    add_bible_slide(prs, directory, "시편", "34:7")
    add_bible_slide(prs, directory, "요한복음", "14:13", "14:14")
    add_bible_slide(prs, directory, "사도행전", "3:5")
    add_bible_slide(prs, directory, "사도행전", "3:6")
    add_bible_slide(prs, directory, "에스겔", "37:1", "37:10")
    add_bible_slide(prs, directory, "사도행전", "3:7", "3:8")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "요한복음", "5:24")
    add_bible_slide(prs, directory, "사도행전", "3:10")
    add_bible_slide(prs, directory, "역대하", "16:9")
    add_bible_slide(prs, directory, "사도행전", "3:12")
    add_bible_slide(prs, directory, "사도행전", "3:16")
    
    
    add_hymn_slide(prs,  '예수의 이름으로 나는 일어서리라')
    add_hymn_slide(prs,  '충만')
    add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
