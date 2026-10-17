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

hymn_list = ["은혜", "비 준비하시니", "주의 친절한 팔에 안기세", "내 영혼이 은총 입어", "주님의 은혜 내게 넘치네", "나는 믿네"]
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
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')

    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    
    add_bible_slide(prs, directory, "사무엘상", "13:8", "13:14")
    add_subtitle_slide(prs, input_text="부득이하여 그리하였나이다 (사무엘상 13:8~14)")
    
    add_bible_slide(prs, directory, "사무엘상", "13:12")
    add_bible_slide(prs, directory, "잠언", "3:5", "3:6")
    add_bible_slide(prs, directory, "사무엘상", "13:9")
    add_bible_slide(prs, directory, "시편", "56:3")
    add_bible_slide(prs, directory, "출애굽기", "14:13")
    add_bible_slide(prs, directory, "사무엘상", "13:11", "13:12")
    add_bible_slide(prs, directory, "사무엘상", "15:22")
    add_bible_slide(prs, directory, "창세기", "3:12")
    add_bible_slide(prs, directory, "사무엘상", "13:13")
    add_bible_slide(prs, directory, "시편", "46:10")
    add_bible_slide(prs, directory, "마태복음", "26:39")
    add_bible_slide(prs, directory, "빌립보서", "4:6", "4:7")
    
    # add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, "하나님의 부르심")
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '주님 말씀하시면')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
