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

hymn_list = ['우리 보좌 앞에 모였네', '피난처 되시는 주 예수', '죄에서 자유를 얻게 함은', '예수 열방의 소망', '하나님의 나라 온 땅 흔드네', '하나님의 부르심', '주님 다시 오실 때 까지' , '파송의 노래']

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
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    

    add_hymn_slide(prs, hymn_list[6])
    add_bible_slide(prs, directory, "마태복음", "6:25", "6:33")
    add_subtitle_slide(prs, input_text="먼저 그 나라와 의를 구하는 사람 (마태복음 6:25~33)")
    
    add_bible_slide(prs, directory, "창세기", "12:1")
    add_bible_slide(prs, directory, "열왕기상", "3:11")
    add_bible_slide(prs, directory, "빌립보서", "4:19")
    add_bible_slide(prs, directory, "마태복음", "6:25")
    add_bible_slide(prs, directory, "마태복음", "6:26")
    add_bible_slide(prs, directory, "마태복음", "6:28")
    add_bible_slide(prs, directory, "로마서", "10:14")
    add_bible_slide(prs, directory, "요한복음", "20:21")
    add_bible_slide(prs, directory, "마태복음", "6:33")
    add_bible_slide(prs, directory, "빌립보서", "4:19")

    # add_hymn_slide(prs, hymn_list[4])

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    add_card_slide(prs, input_text= '파송기도 및 축복')
    add_hymn_slide(prs, hymn_list[7])
    # add_hymn_slide(prs,  '부흥 2000')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
