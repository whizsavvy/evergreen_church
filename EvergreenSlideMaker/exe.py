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

hymn_list = ['주님을 보게하소서', '주 하나님 지으신 모든 세계', '내 마음 다해', '하늘 위에 주님 밖에', '주 여호와는 광대하시도다', '나 같은 죄인 살리신', '말씀 앞에서']

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
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_subtitle_slide(prs, input_text='성령과 동행하라 (사도행전 1:8)')
    
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "요한복음", "14:16", "14:17")
    add_bible_slide(prs, directory, "요한일서", "4:4")
    add_bible_slide(prs, directory, "갈라디아서", "5:22", "5:23")
    add_bible_slide(prs, directory, "골로새서", "3:16")
    add_bible_slide(prs, directory, "데살로니가전서", "5:17")
    add_bible_slide(prs, directory, "사도행전", "5:12", "5:13")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")
    add_bible_slide(prs, directory, "디모데후서", "3:12")
    add_bible_slide(prs, directory, "사도행전", "28:31")
    add_bible_slide(prs, directory, "시편", "143:10")
    add_bible_slide(prs, directory, "잠언", "3:5", "3:6")
    add_bible_slide(prs, directory, "여호수아", "1:8")
    add_bible_slide(prs, directory, "히브리서", "10:24", "10:25")
    add_bible_slide(prs, directory, "디모데후서", "3:12")
    
    
                    
    
    
    add_hymn_slide(prs,  '살아계신 성령님 날 붙드소서')
    add_card_slide(prs, input_text= '성찬')    
    add_hymn_slide(prs, hymn_list[5])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
