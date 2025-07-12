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

hymn_list = ['당신을 향한 노래'
,'내 마음을 가득 채운'
,'구주의 십자가 보혈로'
,'주 십자가를 지심으로'
,'우리는 주의 백성이오니' ]

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
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "1:12", "1:14")
    add_subtitle_slide(prs, input_text='순종과 기도로 탄생하는 교회 (행 1:12-14)')
    
    add_bible_slide(prs, directory, "사도행전", "1:6")
    add_bible_slide(prs, directory, "사도행전", "1:1")
    add_bible_slide(prs, directory, "사도행전", "1:3")
    add_bible_slide(prs, directory, "요한복음", "20:28")
    add_bible_slide(prs, directory, "사도행전", "2:36")
    add_bible_slide(prs, directory, "사도행전", "1:4")
    add_bible_slide(prs, directory, "시편", "40:1")
    add_bible_slide(prs, directory, "사도행전", "1:12")
    add_bible_slide(prs, directory, "고린도전서", "15:6")
    add_bible_slide(prs, directory, "사도행전", "1:14", "1:15")
    add_bible_slide(prs, directory, "마태복음", "7:24")
    add_bible_slide(prs, directory, "사도행전", "1:13")
    add_bible_slide(prs, directory, "히브리서", "10:25")
    add_bible_slide(prs, directory, "사도행전", "1:14")
    add_bible_slide(prs, directory, "누가복음", "8:6")
    add_bible_slide(prs, directory, "에베소서", "6:18")
    add_bible_slide(prs, directory, "사도행전", "1:15")
    add_bible_slide(prs, directory, "요한복음", "1:42")
    
    
                    
    
    
    add_hymn_slide(prs,  '부름받아 나선이 몸')
    add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '말씀 앞에서')
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
