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

hymn_list = ['나의 영혼이 잠잠히', '임재', '이 눈에 아무 증거 아니 뵈어도', '구주 예수 의지함이', '주 이름 찬양', '주 임재 안에서', '나는 믿네']
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
    add_hymn_slide(prs, hymn_list[6])
    add_blank_slide(prs)
    
    add_bible_slide(prs, directory, "사무엘상", "12:19", "12:25")
    add_subtitle_slide(prs, input_text="자기 백성을 버리지 아니하실 것이요 (사무엘상 12:19~25)")
    
    add_bible_slide(prs, directory, "사무엘상", "11:13")
    add_bible_slide(prs, directory, "사무엘상", "12:19")
    add_bible_slide(prs, directory, "사무엘상", "8:20")
    add_bible_slide(prs, directory, "사무엘상", "8:7")
    add_bible_slide(prs, directory, "사무엘상", "12:17")
    add_bible_slide(prs, directory, "사무엘상", "12:20")
    add_bible_slide(prs, directory, "에스겔", "18:23")
    add_bible_slide(prs, directory, "사무엘상", "12:21")
    add_bible_slide(prs, directory, "잠언", "3:5", "3:6")
    add_bible_slide(prs, directory, "사무엘상", "12:22")
    add_bible_slide(prs, directory, "신명기", "7:7", "7:8")
    add_bible_slide(prs, directory, "로마서", "5:8")
    add_bible_slide(prs, directory, "사무엘상", "12:25")
    add_bible_slide(prs, directory, "로마서", "5:20")
    add_bible_slide(prs, directory, "로마서", "2:4")
    add_bible_slide(prs, directory, "디도서", "2:11", "2:12")
    add_bible_slide(prs, directory, "사무엘상", "12:23")
    add_bible_slide(prs, directory, "에베소서", "6:18")
    add_bible_slide(prs, directory, "사무엘상", "12:24")
    add_bible_slide(prs, directory, "신명기", "8:2")
    add_bible_slide(prs, directory, "사무엘상", "12:22")
    # add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '오 놀라운 구세주')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '주님 말씀하시면')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
