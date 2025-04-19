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

hymn_list = ['어둔 밤 물리치고', '무덤에 머물러', '이것이 나의 간증이요', '할렐루야 살아계신 주', '이 땅위에 오신']

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
    add_bible_slide(prs, directory, "누가복음", "24:4", "24:8")
    add_subtitle_slide(prs, input_text='왜 산 자를 죽은 자 가운데서 찾느냐(눅 24:4~8)')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "누가복음", "24:4", "24:8")
    add_bible_slide(prs, directory, "누가복음", "23:55", "23:56")
    add_bible_slide(prs, directory, "누가복음", "24:5", "24:5")  # 5절)
    add_bible_slide(prs, directory, "고린도전서", "15:3", "15:6")
    add_bible_slide(prs, directory, "사도행전", "4:20", "4:20")
    add_bible_slide(prs, directory, "사도행전", "2:32", "2:32")
    add_bible_slide(prs, directory, "고린도전서", "15:17", "15:17")
    add_bible_slide(prs, directory, "요한복음", "20:19", "20:19")
    add_bible_slide(prs, directory, "사도행전", "4:13", "4:13")
    add_bible_slide(prs, directory, "사도행전", "2:42", "2:42")
    add_bible_slide(prs, directory, "고린도전서", "15:26", "15:26")

    # add_hymn_slide(prs, '나는 주를 섬기는 것에 후회가 없습니다')
    # add_image_slide(prs)
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    # add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
