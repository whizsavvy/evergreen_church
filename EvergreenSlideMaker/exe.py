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

hymn_list = ['모든 능력과 모든 권세', '죄에서 자유를 얻게 함은', '예수 열방의 소망', 'Winning All', '나 주님의 기쁨 되기 원하네']
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
    # add_hymn_slide(prs, hymn_list[1])

    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[1]) 
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "사도행전", "12:1", "12:5")
    add_subtitle_slide(prs, input_text="위기 앞에 무릎 꿇는 교회 (사도행전 12:1~5)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "사도행전", "12:5")          # “교회는 그를 위하여 간절히…” (핵심 구절)
    add_bible_slide(prs, directory, "시편", "40:2")
    add_bible_slide(prs, directory, "마태복음", "11:28", "11:30")
    add_bible_slide(prs, directory, "히브리서", "11:6")
    add_bible_slide(prs, directory, "사도행전", "12:13", "12:16") # 로데 장면
    add_bible_slide(prs, directory, "사도행전", "12:7")           # “그 밤에 주의 사자가…”
    add_bible_slide(prs, directory, "빌립보서", "4:6", "4:7")

    
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '하나님의 약속')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
