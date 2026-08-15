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

hymn_list = ['거친 길 위를 걸어갈 때도', '감사함으로', '왕 되신 주께 감사하세', '내가 늘 의지하는 예수', '예수 사랑하심을']

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
    

    add_bible_slide(prs, directory, "사무엘상", "3:1", "3:10")
    add_subtitle_slide(prs, input_text="말씀하옵소서, 주의 종이 듣겠나이다 (사무엘상 3:1~10)")
    
    add_bible_slide(prs, directory, "사무엘상", "3:1")
    add_bible_slide(prs, directory, "사무엘상", "2:30")
    add_bible_slide(prs, directory, "사무엘상", "2:17")
    add_bible_slide(prs, directory, "요한계시록", "2:7")
    add_bible_slide(prs, directory, "사무엘상", "3:3")
    add_bible_slide(prs, directory, "이사야", "40:8")
    add_bible_slide(prs, directory, "사무엘상", "3:4")
    add_bible_slide(prs, directory, "사무엘상", "3:5")
    add_bible_slide(prs, directory, "사무엘상", "3:7")
    add_bible_slide(prs, directory, "이사야", "30:21")
    add_bible_slide(prs, directory, "디모데후서", "3:16")
    add_bible_slide(prs, directory, "요한복음", "16:13")
    add_bible_slide(prs, directory, "시편", "119:105")
    add_bible_slide(prs, directory, "요한복음", "10:27")
    add_bible_slide(prs, directory, "사무엘상", "3:9", "3:10")
    add_bible_slide(prs, directory, "사무엘상", "3:9")
    add_bible_slide(prs, directory, "야고보서", "1:22")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")
    add_bible_slide(prs, directory, "디모데후서", "4:7")
    add_bible_slide(prs, directory, "요한복음", "6:38")
    add_bible_slide(prs, directory, "시편", "119:105")
    add_bible_slide(prs, directory, "요한복음", "16:13")

    add_hymn_slide(prs, '말씀 앞에서')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '나의 기도 하는 것보다')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
