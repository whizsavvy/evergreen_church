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

hymn_list = ['나의 갈 길 다 가도록', '나는 주의 친구', '주 안에서 기뻐해', '만왕의 왕 주께서', '예수 피를 힘입어', '유월절 어린 양의 피로']

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
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864", title="주 예수 나의 산  소망")
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')

    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "갈라디아서", "2:20")
    add_subtitle_slide(prs, input_text="복음은 정체성을 바꾼다 (갈라디아서 2:20)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "로마서", "6:6")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")   # “이제 내가 육체 가운데 사는 것은…”
    add_bible_slide(prs, directory, "로마서", "8:1")
    add_bible_slide(prs, directory, "갈라디아서", "6:14")
    add_bible_slide(prs, directory, "로마서", "5:8")
    add_bible_slide(prs, directory, "사도행전", "2:38")
    add_bible_slide(prs, directory, "누가복음", "19:8")
    add_bible_slide(prs, directory, "빌립보서", "1:21")
    add_bible_slide(prs, directory, "로마서", "5:1")
    add_bible_slide(prs, directory, "로마서", "1:16")




    add_hymn_slide(prs, hymn_list[5])
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '우린 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
