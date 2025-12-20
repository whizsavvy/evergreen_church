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

hymn_list = ['천사들의 노래가', '참 반가운 성도여' , '주님 큰 영광 받으소서', '크신 주님', '영광의 이름 예수', '빛나고 높은 보좌와']
def create_presentation(hymn_list=[]):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)
    directory = folder_path+"/bible"
    pic_dic = folder_path+"/image/"
    add_image_slide(prs, pic_dic+'2025.jpg', text='주일 1부 예배')
    add_image_slide(prs, pic_dic+'2025.jpg', text='주일 2부 예배')
    add_blank_slide(prs)
    # add_hymn_slide(prs, hymn_list[0])
    # add_hymn_slide(prs, hymn_list[1])
    # add_hymn_slide(prs, hymn_list[2])
    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864")
    
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    # add_hymn_slide(prs, hymn_list[1]) 
    # add_hymn_slide(prs, hymn_list[3])
    # add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
    
   
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "누가복음", "2:10", "2:11")
    add_subtitle_slide(prs, input_text="너희를 위하여 구주가 나셨으니 (누가복음 2:10~11)")
    
    # RED only — 원고 등장 순서 그대로
    add_bible_slide(prs, directory, "누가복음", "2:10")
    add_bible_slide(prs, directory, "이사야", "41:10")
    add_bible_slide(prs, directory, "마태복음", "1:23")
    add_bible_slide(prs, directory, "요한복음", "14:27")
    
    add_bible_slide(prs, directory, "요한복음", "3:16")
    add_bible_slide(prs, directory, "디도서", "2:11")
    add_bible_slide(prs, directory, "로마서", "10:13")
    
    add_bible_slide(prs, directory, "누가복음", "2:11")
    add_bible_slide(prs, directory, "마태복음", "1:21")
    add_bible_slide(prs, directory, "요한복음", "1:12")
    add_bible_slide(prs, directory, "사도행전", "4:12")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "요한복음", "10:10")
    
    add_bible_slide(prs, directory, "로마서", "15:13")
    add_bible_slide(prs, directory, "이사야", "9:6")
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '창조의 아버지')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
