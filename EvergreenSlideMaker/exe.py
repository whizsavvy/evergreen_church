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

hymn_list = ['약할 때 강함되시네', '우리 보좌 앞에 모였네'
, '곤한 내 영혼 편히 쉴 곳과', '나의 죄를 씻기는', '태산을 넘어 험곡에 가도']
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
    # add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
    
    add_bible_slide(prs, directory, "사무엘상", "9:14", "9:17")
    add_subtitle_slide(prs, input_text="뜻밖의 길에서 (사무엘상 9:14~17)")
    
    add_bible_slide(prs, directory, "사무엘상", "9:4")
    add_bible_slide(prs, directory, "사무엘상", "9:5")
    add_bible_slide(prs, directory, "사무엘상", "9:14")
    add_bible_slide(prs, directory, "사무엘상", "9:16")
    add_bible_slide(prs, directory, "욥기", "23:10")
    add_bible_slide(prs, directory, "로마서", "8:28")
    add_bible_slide(prs, directory, "사무엘상", "9:15")
    add_bible_slide(prs, directory, "출애굽기", "23:20")
    add_bible_slide(prs, directory, "시편", "139:16")
    add_bible_slide(prs, directory, "사무엘상", "9:17")
    add_bible_slide(prs, directory, "사무엘상", "9:20")
    add_bible_slide(prs, directory, "사무엘상", "9:16")
    add_bible_slide(prs, directory, "누가복음", "19:10")
    add_bible_slide(prs, directory, "요한복음", "14:6")

    add_hymn_slide(prs, '은혜')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '내 구주 예수를 더욱 사랑')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '주님의 영광 나타나셨네')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
