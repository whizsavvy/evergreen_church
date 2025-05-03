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

hymn_list = ['주 하나님 지으신 모든 세계', '주 이름 찬양','선교사', '십자가를 질 수 있나', '변찮는 주님의 사랑과', '내가 매일 기쁘게', '보혈을 지나', '이 땅의 동과 서 남과 북']

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
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)


    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "에베소서", "5:15", "5:17")
    add_subtitle_slide(prs, input_text='인생 후반전, 어떻게 살 것인가 (엡 5:15~17)')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "에베소서", "5:15")
    add_bible_slide(prs, directory, "시편", "90:12")
    add_bible_slide(prs, directory, "고린도전서", "10:31")
    add_bible_slide(prs, directory, "마태복음", "5:16")

    add_bible_slide(prs, directory, "에베소서", "5:16")
    add_bible_slide(prs, directory, "전도서", "3:1")

    add_bible_slide(prs, directory, "에베소서", "5:16")
    add_bible_slide(prs, directory, "로마서", "13:11", "13:12")
    add_bible_slide(prs, directory, "디모데후서", "3:1", "3:2")
    add_bible_slide(prs, directory, "마태복음", "24:44")

    add_bible_slide(prs, directory, "에베소서", "5:17")
    add_bible_slide(prs, directory, "요한복음", "15:7")
    add_bible_slide(prs, directory, "이사야", "40:31")
    add_bible_slide(prs, directory, "마태복음", "27:5")
    add_bible_slide(prs, directory, "에베소서", "5:16")
    add_bible_slide(prs, directory, "요한복음", "13:27")
    add_bible_slide(prs, directory, "출애굽기", "7:7")
    add_bible_slide(prs, directory, "사도행전", "20:24")
    add_bible_slide(prs, directory, "고린도전서", "15:58")
    add_bible_slide(prs, directory, "마태복음", "5:23", "5:24")
    add_bible_slide(prs, directory, "골로새서", "3:13")
    add_bible_slide(prs, directory, "히브리서", "12:14")
  
    # add_hymn_slide(prs, '나는 주를 섬기는 것에 후회가 없습니다')
    # add_image_slide(prs)
    add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[7])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
