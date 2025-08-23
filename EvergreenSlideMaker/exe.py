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

hymn_list = ['갈 길을 밝히 보이시니', '내게 강 같은 평화', '비 준비하시니', '곤한 내 영혼 편히 쉴 곳과', '주님은 나의 힘이요', '마음속에 근심 있는 사람', '하나님의 부르심', '부름 받아 나선 이 몸']

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
    add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "5:42")
    add_subtitle_slide(prs, input_text="날마다 일어나는 교회 (사도행전 5:42)")
    
    add_bible_slide(prs, directory, "사도행전", "5:14", "5:16")
    add_bible_slide(prs, directory, "사도행전", "5:17", "5:18")
    add_bible_slide(prs, directory, "사도행전", "5:29")
    add_bible_slide(prs, directory, "사도행전", "5:32")
    add_bible_slide(prs, directory, "갈라디아서", "1:10")
    add_bible_slide(prs, directory, "사도행전", "5:33")
    add_bible_slide(prs, directory, "사도행전", "5:34")
    add_bible_slide(prs, directory, "사도행전", "5:36")
    add_bible_slide(prs, directory, "호세아", "11:8")
    add_bible_slide(prs, directory, "고린도전서", "10:13")
    add_bible_slide(prs, directory, "다니엘", "6:10")  # 언급은 없지만 "하루 세 번 기도" 직접적 연결 절 (선택사항)
    add_bible_slide(prs, directory, "다니엘", "6:22")  # 사자 입 막힘
    add_bible_slide(prs, directory, "다니엘", "6:23")  # 구출 장면
    add_bible_slide(prs, directory, "다니엘", "6:24")  # 반전과 형벌
    add_bible_slide(prs, directory, "다니엘", "6:26")  # 왕의 선포
    add_bible_slide(prs, directory, "다니엘", "6:27")  # 구원하시는 하나님
    add_bible_slide(prs, directory, "마태복음", "6:33")
    add_bible_slide(prs, directory, "다니엘", "12:3")


    
    
    add_hymn_slide(prs, hymn_list[7])

    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
