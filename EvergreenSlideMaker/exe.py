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

hymn_list = ['사랑한다 말하시네', '하나님의 부르심', '불을 내려주소서', '성령이여 임하소서', '나로부터 시작되리']

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
    # add_hymn_slide(prs, hymn_list[1])
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "에베소서", "4:11", "4:16")
    add_subtitle_slide(prs, input_text="부르심을 따라 사는 교회 (에베소서 4:11~16)")
    
    add_bible_slide(prs, directory, "출애굽기", "3:10")
    add_bible_slide(prs, directory, "에베소서", "4:1")
    add_bible_slide(prs, directory, "베드로전서", "2:9")
    add_bible_slide(prs, directory, "사사기", "6:12")
    add_bible_slide(prs, directory, "에베소서", "4:4")
    add_bible_slide(prs, directory, "에베소서", "4:4", "4:6")
    add_bible_slide(prs, directory, "에베소서", "4:7")
    add_bible_slide(prs, directory, "에베소서", "4:12")
    add_bible_slide(prs, directory, "에베소서", "4:13")
    add_bible_slide(prs, directory, "사도행전", "2:46", "2:47")
    add_bible_slide(prs, directory, "요한복음", "13:35")

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '나의 기도 하는 것보다')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
