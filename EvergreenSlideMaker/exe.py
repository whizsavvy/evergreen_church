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

hymn_list = ['그 크신 하나님의 사랑', '높은 산들 흔들리고', '변찮는 주님의 사랑과', '이 땅 위에 오신', '모든 상황 속에서', '이 땅의 동과 서 남과 북']

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
    
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)

    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])





    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "마태복음", "6:31", "6:33")
    add_bible_slide(prs, directory, "히브리서", "11:6")
    add_subtitle_slide(prs, input_text='믿음은 오늘을 이기게 합니다')
    add_blank_slide(prs)
    add_subtitle_slide(prs, input_text='1) 믿음은 하나님과의 관계의 시작입니다.')
    add_bible_slide(prs, directory, "히브리서", "11:6")
    add_subtitle_slide(prs, input_text='2) 믿음은 구원의 길입니다.')
    add_bible_slide(prs, directory, "에베소서", "2:8")
    add_bible_slide(prs, directory, "로마서", "5:1")

    add_subtitle_slide(prs, input_text='3) 믿음은 신자의 삶의 방식입니다.')
    add_bible_slide(prs, directory, "로마서", "1:17")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")

    add_subtitle_slide(prs, input_text='4) 믿음은 세상을 이기는 능력입니다.')
    add_bible_slide(prs, directory, "요한일서", "5:4")


    add_subtitle_slide(prs, input_text='1. 믿음은 평안을 줍니다')
    add_bible_slide(prs, directory, "시편", "23:1", "23:3")
    add_bible_slide(prs, directory, "요한복음", "14:27")
    add_bible_slide(prs, directory, "빌립보서", "4:6", "4:7")
    add_bible_slide(prs, directory, "이사야", "26:3")
    add_bible_slide(prs, directory, "마가복음", "5:25", "5:34")


    add_subtitle_slide(prs, input_text='2. 믿음은 하나님의 인도하심을 경험하게 합니다')
    add_bible_slide(prs, directory, "시편", "23:4")

    
    add_bible_slide(prs, directory, "잠언", "3:5", "3:6")
    add_bible_slide(prs, directory, "시편", "32:8")

    add_subtitle_slide(prs, input_text='3. 믿음은 공급과 채우심의 은혜를 경험하게 합니다')
    add_bible_slide(prs, directory, "마태복음", "6:31", "6:33")

    

    add_bible_slide(prs, directory, "열왕기상", "17:8", "17:16")
    add_bible_slide(prs, directory, "빌립보서", "4:19")
    add_bible_slide(prs, directory, "말라기", "3:10")
    add_bible_slide(prs, directory, "출애굽기", "19:4")

    add_bible_slide(prs, directory, "요한일서", "5:4")


  
    # add_hymn_slide(prs, '나는 주를 섬기는 것에 후회가 없습니다')
    # add_image_slide(prs)
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_hymn_slide(prs, hymn_list[4])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[5])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
