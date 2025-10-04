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

hymn_list = ['우리 보좌 앞에 모였네', '슬픈 마음 있는 사람', '내게 강 같은 평화', '주 이름 찬양', '주가 일하시네', '물 위를 걷는 자', ''주 임재 안에서', '하나님의 약속']

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
    add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "시편", "128:1", "128:2")
    add_subtitle_slide(prs, input_text="여호와께서 너희 가정에 복을 주시리라 (시편 128:1–2)")
    
    add_bible_slide(prs, directory, "여호수아", "24:15")
    add_bible_slide(prs, directory, "잠언", "1:7")
    add_bible_slide(prs, directory, "잠언", "14:26")
    add_bible_slide(prs, directory, "잠언", "14:27")
    add_bible_slide(prs, directory, "창세기", "22:12")
    add_bible_slide(prs, directory, "잠언", "8:13")
    add_bible_slide(prs, directory, "잠언", "22:4")
    add_bible_slide(prs, directory, "창세기", "31:42")
    add_bible_slide(prs, directory, "시편", "127:1")
    add_bible_slide(prs, directory, "잠언", "3:6")
    add_bible_slide(prs, directory, "고린도전서", "10:31")
    add_bible_slide(prs, directory, "시편", "128:3")  # (3절)
    add_bible_slide(prs, directory, "창세기", "45:5")
    add_bible_slide(prs, directory, "골로새서", "3:13")
    add_bible_slide(prs, directory, "마태복음", "18:21", "18:22")
    add_bible_slide(prs, directory, "시편", "128:6")  # (6절)
    add_bible_slide(prs, directory, "디모데후서", "1:5")
    


    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_hymn_slide(prs, '하나님의 약속')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '오늘 숨을 쉬는 것 감사')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
