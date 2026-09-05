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

hymn_list = ['나 주를 멀리 떠났다', '변찮는 주님의 사랑과', '주 이름 찬양', '엘리야의 날', '살아계신 주', '비 준비하시니']
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
    
    add_bible_slide(prs, directory, "사무엘상", "8:4", "8:22")
    add_subtitle_slide(prs, input_text="누가 나의 왕인가요? (사무엘상 8:4~22)")
    
    add_bible_slide(prs, directory, "사무엘상", "7:12")
    add_bible_slide(prs, directory, "사무엘상", "8:5")
    add_bible_slide(prs, directory, "사무엘상", "8:20")
    add_bible_slide(prs, directory, "사무엘상", "8:4", "8:5")
    add_bible_slide(prs, directory, "신명기", "17:14")
    add_bible_slide(prs, directory, "신명기", "17:15")
    add_bible_slide(prs, directory, "출애굽기", "19:5", "19:6")
    add_bible_slide(prs, directory, "로마서", "12:2")
    add_bible_slide(prs, directory, "시편", "20:7")
    add_bible_slide(prs, directory, "사무엘상", "8:6")
    add_bible_slide(prs, directory, "사무엘상", "8:7")
    add_bible_slide(prs, directory, "로마서", "10:9")
    add_bible_slide(prs, directory, "마태복음", "26:39")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")
    add_bible_slide(prs, directory, "사무엘상", "8:22")
    add_bible_slide(prs, directory, "사무엘상", "8:17")
    add_bible_slide(prs, directory, "마가복음", "10:45")
    add_bible_slide(prs, directory, "빌립보서", "2:8", "2:11")

    # add_hymn_slide(prs, '지금까지 지내온 것')
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '내 구주 예수를 더욱 사랑')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '주님의 영광 나타나셨네')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
