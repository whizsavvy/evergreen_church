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

hymn_list = ['그 사랑', '더 원합니다', '주님 큰 영광 받으소서'
, '주 이름 큰 능력 있도다', '할렐루야 살아계신 주', '무덤에 머물러'

'나 같은 죄인 살리신' , '예수 나의 산 소망']

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
    add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    

    # add_card_slide(prs, input_text= '성가대 찬양')
    # add_bible_slide(prs, directory, "시편", "46:1", "46:11")
    # add_subtitle_slide(prs, input_text="하나님께 집중하다 – 가만히 있으라 (시편 46:1~11)")
    
    add_bible_slide(prs, directory, "고린도전서", "15:3", "15:8")
    add_subtitle_slide(prs, input_text="믿을 것인가, 부인할 것인가 (고린도전서 15:3~8)")
    
    add_bible_slide(prs, directory, "고린도전서", "15:3", "15:8")
    add_bible_slide(prs, directory, "로마서", "6:23")
    add_bible_slide(prs, directory, "히브리서", "9:22")
    add_bible_slide(prs, directory, "레위기", "17:11")
    add_bible_slide(prs, directory, "히브리서", "10:4")
    add_bible_slide(prs, directory, "로마서", "3:25")
    add_bible_slide(prs, directory, "요한일서", "2:2")
    add_bible_slide(prs, directory, "로마서", "5:8")
    add_bible_slide(prs, directory, "로마서", "4:25")
    add_bible_slide(prs, directory, "고린도전서", "15:54", "15:55")
    add_bible_slide(prs, directory, "히브리서", "2:14")
    add_bible_slide(prs, directory, "로마서", "1:4")
    add_bible_slide(prs, directory, "고린도전서", "15:20")
    add_bible_slide(prs, directory, "요한복음", "14:19")
    add_bible_slide(prs, directory, "히브리서", "7:25")
    add_bible_slide(prs, directory, "고린도전서", "15:6")
    add_bible_slide(prs, directory, "고린도전서", "15:5", "15:6")
    add_bible_slide(prs, directory, "사도행전", "4:20")
    add_bible_slide(prs, directory, "요한일서", "1:1")
    add_bible_slide(prs, directory, "고린도전서", "15:55")


    
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs,  '예수 나의 산 소망')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
