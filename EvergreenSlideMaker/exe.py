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

hymn_list = ['원하고 바라고 기도합니다'
, '주 이름 큰 능력 있도다', '주 안에서 기뻐해', '예수님 목마릅니다', '나의 한숨을 바꾸셨네'
, '나의 죄를 씻기는', '주 하나님 독생자 예수', '기뻐하며 왕께 노래 부르리']

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
    add_bible_slide(prs, directory, "창세기", "1:26", "1:31")
    add_subtitle_slide(prs, input_text="하나님이 보시기에 심히 좋았더라 (창세기 1:26~31)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "창세기", "1:26", "1:27")
    add_bible_slide(prs, directory, "창세기", "1:28")
    add_bible_slide(prs, directory, "베드로전서", "2:9")
    add_bible_slide(prs, directory, "이사야", "43:4")
    add_bible_slide(prs, directory, "에베소서", "2:10")
    add_bible_slide(prs, directory, "로마서", "5:8")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")


    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs, hymn_list[7])
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '예수를 나의 구주 삼고')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '보라 새 일을 행하시리니')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
