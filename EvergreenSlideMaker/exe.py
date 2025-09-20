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

hymn_list = ['내 영혼의 그윽히 깊은데서', '송축해 내 영혼', '주 앙망하는 자', '살아계신 주', '비 준비하시니', '세상 흔들리고']

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
    # add_hymn_slide(prs, hymn_list[5])
    # add_hymn_slide(prs, hymn_list[6])
    



    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    add_bible_slide(prs, directory, "사도행전", "6:15")
    add_subtitle_slide(prs, input_text="그 얼굴이 천사의 얼굴과 같더라 (사도행전 6:15)")
    
    add_bible_slide(prs, directory, "사도행전", "6:8")
    add_bible_slide(prs, directory, "출애굽기", "17:1")
    add_bible_slide(prs, directory, "출애굽기", "17:7")
    add_bible_slide(prs, directory, "출애굽기", "17:8")
    add_bible_slide(prs, directory, "호세아", "6:3")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")
    add_bible_slide(prs, directory, "사도행전", "6:9")
    add_bible_slide(prs, directory, "사도행전", "6:10")
    add_bible_slide(prs, directory, "마태복음", "10:17", "10:19")
    add_bible_slide(prs, directory, "사도행전", "6:11", "6:13")
    add_bible_slide(prs, directory, "사도행전", "6:14")
    add_bible_slide(prs, directory, "요한복음", "2:19", "2:21")
    add_bible_slide(prs, directory, "사도행전", "6:15")
    add_bible_slide(prs, directory, "고린도후서", "4:7", "4:8")
    add_bible_slide(prs, directory, "사도행전", "7:54", "7:56")
    add_bible_slide(prs, directory, "빌립보서", "3:8")
    add_hymn_slide(prs, hymn_list[5])


    # add_card_slide(prs, input_text= '성찬') 
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '우리 오늘 눈물로')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
