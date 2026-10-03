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

hymn_list = ['하늘문 여소서', '주의 보좌로 나아갈 때에', '송축해 내 영혼', '죄에서 자유를 얻게 함은', '나의 죄를 씻기는', '주 임재 안에서']
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
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    
    add_bible_slide(prs, directory, "사무엘상", "11:6", "11:13")
    add_subtitle_slide(prs, input_text="위기 앞에서 일어서는 사람 (사무엘상 11:6~13)")
    
    add_bible_slide(prs, directory, "사무엘상", "10:7")
    add_bible_slide(prs, directory, "사무엘상", "11:1")
    add_bible_slide(prs, directory, "사무엘상", "11:2")
    add_bible_slide(prs, directory, "사무엘상", "11:5")
    add_bible_slide(prs, directory, "사무엘상", "11:7")
    add_bible_slide(prs, directory, "사무엘상", "11:6")
    add_bible_slide(prs, directory, "사무엘상", "10:6")
    add_bible_slide(prs, directory, "사무엘상", "11:7")
    add_bible_slide(prs, directory, "사무엘상", "11:11")
    add_bible_slide(prs, directory, "사무엘상", "10:27")
    add_bible_slide(prs, directory, "사무엘상", "11:12")
    add_bible_slide(prs, directory, "사무엘상", "11:13")
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '내 구주 예수를 더욱 사랑')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '주님 말씀하시면')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
