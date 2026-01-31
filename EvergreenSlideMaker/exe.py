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

hymn_list = ['내 주 되신 주를 참 사랑하고', '완전하신 나의 주', '내가 매일 기쁘게', '온 세상 위하여', 내 영혼이 은총 입어', '주 예수 나의 참 소망', '충만', '보혈을 지나']

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
    add_hymn_slide(prs, hymn_list[5])
    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864", title="주 예수 나의 산  소망")
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
       
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "요한복음", "3:3", "3:7")
    add_subtitle_slide(prs, input_text="복음은 나를 다시 태어나게 한다 (요한복음 3:3~7)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "요한복음", "3:3")            # “거듭나지 아니하면…”
    add_bible_slide(prs, directory, "요한복음", "3:1", "3:2")      # 니고데모 소개
    add_bible_slide(prs, directory, "요한복음", "3:5")            # 물과 성령으로 나지 아니하면
    add_bible_slide(prs, directory, "요한복음", "3:6")            # 육으로 난 것은 육…
    add_bible_slide(prs, directory, "고린도후서", "5:17")          # 새로운 피조물
    add_bible_slide(prs, directory, "요한복음", "3:4")            # “어떻게 날 수 있사옵나이까”
    add_bible_slide(prs, directory, "요한복음", "3:8")            # 바람처럼 임하시는 성령
    add_bible_slide(prs, directory, "요한복음", "3:14", "3:15")    # 놋뱀/인자 들림
    add_bible_slide(prs, directory, "로마서", "5:8")               # 우리가 아직 죄인 되었을 때
    add_bible_slide(prs, directory, "갈라디아서", "2:20")          # 내가 그리스도와 함께…
    add_bible_slide(prs, directory, "요한일서", "1:9")             # 우리가 우리 죄를 자백하면
    add_bible_slide(prs, directory, "빌립보서", "3:8")             # 가장 고상한 지식
    add_bible_slide(prs, directory, "베드로전서", "2:2")           # 신령한 젖을 사모하라
    add_bible_slide(prs, directory, "요한복음", "3:7")            # 너희는 거듭나야 하겠다



    add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '우린 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
