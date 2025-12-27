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

hymn_list = ['곤한 내 영혼 편히 쉴 곳과', '주님은 나의 힘이요', '마음속에 근심 있는 사람', '하나님의 부르심', '찬송하며 살리라', '은혜', '아 하나님의 은혜로', '하나님의 부르심']
def create_presentation(hymn_list=[]):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)
    directory = folder_path+"/bible"
    pic_dic = folder_path+"/image/"
    # add_image_slide(prs, pic_dic+'2025.jpg', text='주일 1부 예배')
    # add_image_slide(prs, pic_dic+'2025.jpg', text='주일 2부 예배')
    add_image_slide(prs, pic_dic+'2025.jpg', text='성탄감사예배')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[0])
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_choir_slides_from_file(prs, box_color="203864")
    
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    # add_hymn_slide(prs, hymn_list[1]) 
    # add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
    
   
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "빌립보서", "3:13", "3:14")
    add_subtitle_slide(prs, input_text="사명을 품고 다시 달려갑시다 (빌립보서 3:13~14)")
    
    # RED only — 원고 순서 그대로
    add_bible_slide(prs, directory, "빌립보서", "3:13", "3:14")   # 서두 인용
    
    # 13절> 관련 전개
    add_bible_slide(prs, directory, "빌립보서", "3:13")
    
    # ‘달려간다’(12,14절 표기 전개)
    add_bible_slide(prs, directory, "빌립보서", "3:12")
    add_bible_slide(prs, directory, "빌립보서", "3:14")
    
    # 1) 뒤에 것을 잊고 앞을 잡다
    add_bible_slide(prs, directory, "히브리서", "12:1")
    add_bible_slide(prs, directory, "이사야", "1:18")
    add_bible_slide(prs, directory, "시편", "103:4", "103:5")
    add_bible_slide(prs, directory, "히브리서", "10:30")
    add_bible_slide(prs, directory, "빌립보서", "3:12")           # “이미 얻었다 함도 아니요…”
    add_bible_slide(prs, directory, "빌립보서", "3:13")           # “뒤에 있는 것은 잊어버리고…”
    
    # 2) 푯대를 향하여 달리자
    add_bible_slide(prs, directory, "빌립보서", "3:14")           # “부름의 상을 위하여…”
    add_bible_slide(prs, directory, "빌립보서", "1:6")
    add_bible_slide(prs, directory, "마태복음", "13:31", "13:33")
    add_bible_slide(prs, directory, "히브리서", "11:1")
    add_bible_slide(prs, directory, "갈라디아서", "6:9")

    add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs, hymn_list[7])
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '창조의 아버지')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
