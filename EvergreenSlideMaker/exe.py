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

hymn_list = ['나의 피난처 예수', '확정되었네', '오 주여 나의 마음이', '주의 인자하심이', '오직 예수 뿐이네', '아 하나님의 은혜로']
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
    add_hymn_slide(prs, hymn_list[2]) 
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "시편", "23:1")
    add_subtitle_slide(prs, input_text="부족함 속에 피는 믿음의 고백 (시편 23:1)")
    
    # RED only — 원고 순서 그대로
    add_bible_slide(prs, directory, "요한복음", "1:12")
    add_bible_slide(prs, directory, "출애굽기", "6:2", "6:4")
    add_bible_slide(prs, directory, "예레미야", "33:3")
    add_bible_slide(prs, directory, "시편", "34:10")
    add_bible_slide(prs, directory, "빌립보서", "4:11", "4:12")
    add_bible_slide(prs, directory, "빌립보서", "4:6", "4:7")
    add_bible_slide(prs, directory, "창세기", "13:15")
    
    # (1절) — 본문 절 표기
    add_bible_slide(prs, directory, "시편", "23:1")
    
    # (2절)
    add_bible_slide(prs, directory, "시편", "23:2")
    add_bible_slide(prs, directory, "마태복음", "6:31")
    add_bible_slide(prs, directory, "마태복음", "6:34")
    
    # (3절)
    add_bible_slide(prs, directory, "시편", "23:3")
    add_bible_slide(prs, directory, "이사야", "54:10")
    add_bible_slide(prs, directory, "예레미야애가", "3:22", "3:23")
    add_bible_slide(prs, directory, "창세기", "2:7")
    
    # (4절)
    add_bible_slide(prs, directory, "시편", "23:4")
    add_bible_slide(prs, directory, "에베소서", "3:18", "3:19")
    
    # (5절)
    add_bible_slide(prs, directory, "시편", "23:5")
    add_bible_slide(prs, directory, "시편", "37:5")
    add_bible_slide(prs, directory, "에베소서", "3:20")
    
    # (6절)
    add_bible_slide(prs, directory, "시편", "23:6")
    

    add_hymn_slide(prs, '여호와 나의 목자 내게 부족없네')
    
    add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '하나님의 약속')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
