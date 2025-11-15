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

hymn_list = ['비 준비하시니', '하늘 위에 주님 밖에', '성도여 다 함께', '곤한 내 영혼 편히 쉴 곳과', '슬픈 마음 있는 사람', '은혜']
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
    add_bible_slide(prs, directory, "빌립보서", "1:3", "1:4")
    add_subtitle_slide(prs, input_text="감사와 기쁨으로 (빌립보서 1:3~4)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "로마서", "1:21")
    add_bible_slide(prs, directory, "골로새서", "3:15", "3:17")
    add_bible_slide(prs, directory, "시편", "100:4")
    add_bible_slide(prs, directory, "데살로니가전서", "5:18")
    add_bible_slide(prs, directory, "다니엘서", "6:10")
    add_bible_slide(prs, directory, "욥기", "1:21")
    
    add_bible_slide(prs, directory, "빌립보서", "1:4")
    add_bible_slide(prs, directory, "빌립보서", "4:4")
    add_bible_slide(prs, directory, "시편", "16:11")
    add_bible_slide(prs, directory, "골로새서", "3:23")

    
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '하나님의 약속')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
