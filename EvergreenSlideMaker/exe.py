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

hymn_list = ['찬양하라 내 영혼아', '사랑하는 나의 아버지', '약할 때 강함되시네', '내 영혼이 은총 입어', '태산을 넘어 험곡에 가도', '나 같은 죄인 살리신', '나의 갈 길 다 가도록']

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
    add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "시편", "15:1", "15:5")
    add_subtitle_slide(prs, input_text="주의 성산에 사는 자 (시편 15:1~5)")
    
    add_bible_slide(prs, directory, "시편", "5:4")
    add_bible_slide(prs, directory, "이사야", "33:14")
    add_bible_slide(prs, directory, "히브리서", "10:19")
    add_bible_slide(prs, directory, "히브리서", "4:16")
    add_bible_slide(prs, directory, "시편", "61:3", "61:4")
    add_bible_slide(prs, directory, "시편", "27:4", "27:5")
    add_bible_slide(prs, directory, "시편", "73:28")
    add_bible_slide(prs, directory, "창세기", "6:9")
    add_bible_slide(prs, directory, "창세기", "17:1")
    add_bible_slide(prs, directory, "에스겔", "18:5", "18:9")
    add_bible_slide(prs, directory, "이사야", "29:13")
    add_bible_slide(prs, directory, "사무엘상", "15:22")
    add_bible_slide(prs, directory, "창세기", "4:4", "4:5")
    add_bible_slide(prs, directory, "마태복음", "5:14", "5:16")
    add_bible_slide(prs, directory, "레위기", "19:16")
    add_bible_slide(prs, directory, "마태복음", "5:23", "5:24")
    add_bible_slide(prs, directory, "출애굽기", "22:25")
    add_bible_slide(prs, directory, "잠언", "17:23")
    add_bible_slide(prs, directory, "아모스", "5:12")

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '그 날')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
