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

hymn_list = ['Winning All', '영광의 이름 예수', '모든 열방 주 볼 때까지', '비 준비하시니']

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
    # add_hymn_slide(prs, hymn_list[3])
    # add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864")
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
       
        # 제목/본문 (필수)
    add_bible_slide(prs, directory, "로마서", "12:1", "12:2")
    add_subtitle_slide(prs, input_text="창조신앙 – 예배는 삶의 방향을 바꾼다 (로마서 12:1~2)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "로마서", "12:1")          # (롬 12:1)
    add_bible_slide(prs, directory, "로마서", "12:2")          # 2절>
    
    add_bible_slide(prs, directory, "창세기", "1:1")           # 태초에 하나님이 천지를 창조하시니라
    add_bible_slide(prs, directory, "창세기", "1:2")           # 혼돈·공허·흑암… 성령 운행
    
    add_bible_slide(prs, directory, "히브리서", "11:3")        # 믿음으로 모든 세계가…
    add_bible_slide(prs, directory, "시편", "33:6")            # 여호와의 말씀으로 하늘이 지음
    add_bible_slide(prs, directory, "시편", "33:9")            # 말씀하시매 이루어졌으며…
    
    add_bible_slide(prs, directory, "시편", "90:2")            # 영원부터 영원까지 주는 하나님
    add_bible_slide(prs, directory, "시편", "139:13", "139:14")# 모태에서 나를 지으심
    add_bible_slide(prs, directory, "이사야", "43:21")         # 나를 찬송하게 하려 함이니라
    add_bible_slide(prs, directory, "시편", "147:4", "147:5")  # 별들의 수효를 세시고…
    
    # 결론부 재강조(원고 순서상 재등장)
    add_bible_slide(prs, directory, "로마서", "12:2")          # 본문 재인용



    add_hymn_slide(prs, hymn_list[3])
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '예수를 나의 구주 삼고')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '보라 새 일을 행하시리니')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
