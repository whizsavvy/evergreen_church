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

hymn_list = ['예수로 나의 구주 삼고', '선하신 목자',  '축복의 통로', '주의 보좌로 나아갈 때에', '하나님 한번도 나를']

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
    # add_hymn_slide(prs, hymn_list[4])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "시편", "23:1", "23:6")
    add_subtitle_slide(prs, input_text='뜻밖의 희망(시 23:1-6)')
    
    add_bible_slide(prs, directory, "시편", "23:1")
    add_bible_slide(prs, directory, "시편", "23:2")
    add_bible_slide(prs, directory, "시편", "23:3")
    add_bible_slide(prs, directory, "시편", "19:7")
    add_bible_slide(prs, directory, "시편", "25:11")
    add_bible_slide(prs, directory, "시편", "23:4")
    add_bible_slide(prs, directory, "예레미야", "2:6")
    add_bible_slide(prs, directory, "이사야", "41:10")
    add_bible_slide(prs, directory, "사무엘상", "17:43")
    add_bible_slide(prs, directory, "시편", "23:5")
    add_bible_slide(prs, directory, "누가복음", "7:44", "7:46")
    add_bible_slide(prs, directory, "시편", "23:6")
    add_bible_slide(prs, directory, "호세아", "8:3")
    add_bible_slide(prs, directory, "시편", "25:7")
    add_bible_slide(prs, directory, "에베소서", "2:19")
    add_bible_slide(prs, directory, "시편", "23:1")
    add_bible_slide(prs, directory, "로마서", "8:18")
    add_bible_slide(prs, directory, "히브리서", "11:1")
    
    
    add_hymn_slide(prs, hymn_list[4])
    add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
