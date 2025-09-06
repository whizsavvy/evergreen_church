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

hymn_list = ['하나님의 부르심', '주의 진리 위해 십자가 군기', '흑암에 사는 백성들을 보라', '이 땅의 동과 서 남과 북']

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
    # add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])
    # add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "시편", "8:1", "8:9")
    add_subtitle_slide(prs, input_text="주의 이름이 어찌 그리 아름다운지요 (시편 8:1~9)")
    
    add_bible_slide(prs, directory, "이사야", "40:6")
    add_bible_slide(prs, directory, "고린도후서", "4:7")
    add_bible_slide(prs, directory, "야고보서", "4:4")
    add_bible_slide(prs, directory, "로마서", "5:6", "5:8")
    add_bible_slide(prs, directory, "에베소서", "2:1")
    add_bible_slide(prs, directory, "에베소서", "2:8")
    add_bible_slide(prs, directory, "시편", "8:4")
    add_bible_slide(prs, directory, "출애굽기", "3:14")
    add_bible_slide(prs, directory, "살전", "5:17", "5:18")
    add_bible_slide(prs, directory, "누가복음", "18:1")
    add_bible_slide(prs, directory, "시편", "8:2")
    add_bible_slide(prs, directory, "고린도전서", "1:27")
    add_bible_slide(prs, directory, "창세기", "1:26")
    add_bible_slide(prs, directory, "시편", "8:5")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "시편", "8:6")
    add_bible_slide(prs, directory, "요한계시록", "5:10")
    add_bible_slide(prs, directory, "로마서", "1:20")
    add_bible_slide(prs, directory, "시편", "8:9")


    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '우리 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
