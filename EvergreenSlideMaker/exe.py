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

hymn_list = ['부르신 곳에서', '허락하신 새 땅에', '갈 길을 밝히 보이시니', '나로부터 시작되리', '세상의 유혹 시험이', '나는 예배자입니다', '예배자']

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
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864")
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
       
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "시편", "95:1", "95:7")
    add_subtitle_slide(prs, input_text="예배는 선택이 아니라 생명이다 (시편 95:1~7 / 요한복음 4:23)")
    
    # RED only — 원고 순서
    add_bible_slide(prs, directory, "요한계시록", "3:1")
    add_bible_slide(prs, directory, "요한계시록", "3:2")
    
    # 1. 예배는 하나님을 만나는 자리
    add_bible_slide(prs, directory, "시편", "95:6", "95:7")
    add_bible_slide(prs, directory, "출애굽기", "20:24")
    add_bible_slide(prs, directory, "출애굽기", "8:27")
    add_bible_slide(prs, directory, "출애굽기", "8:25")
    add_bible_slide(prs, directory, "신명기", "12:5")
    add_bible_slide(prs, directory, "신명기", "12:11")
    
    # 2. 예배가 살아나면 시선이 바뀝니다
    add_bible_slide(prs, directory, "이사야", "6:1")
    add_bible_slide(prs, directory, "이사야", "6:5")
    add_bible_slide(prs, directory, "이사야", "6:8")
    add_bible_slide(prs, directory, "시편", "73:17")
    add_bible_slide(prs, directory, "고린도후서", "4:18")
    
    # 3. 하나님이 찾으시는 예배자
    add_bible_slide(prs, directory, "요한복음", "4:23")
    add_bible_slide(prs, directory, "요한복음", "14:6")
    add_bible_slide(prs, directory, "아모스", "5:24")
    
    # 결단부
    add_bible_slide(prs, directory, "시편", "84:10")
    add_bible_slide(prs, directory, "마태복음", "4:4")


    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '예수를 나의 구주 삼고')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '보라 새 일을 행하시리니')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
