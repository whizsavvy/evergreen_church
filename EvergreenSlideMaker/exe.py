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

hymn_list = ['거룩하신 하나님 주께 감사드리세', '나의 등 뒤에서', '성도여 다 함께', '내가 매일 기쁘게', '부흥']

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
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "사무엘상", "1:1", "1:11")
    add_subtitle_slide(prs, input_text="기도가 역사를 시작합니다 (사무엘상 1:1~11)")
    
    add_bible_slide(prs, directory, "사사기", "21:25")
    add_bible_slide(prs, directory, "사무엘상", "3:1")
    add_bible_slide(prs, directory, "사무엘상", "2:12")
    add_bible_slide(prs, directory, "사무엘상", "1:5")
    add_bible_slide(prs, directory, "사무엘상", "1:6")
    add_bible_slide(prs, directory, "사무엘상", "1:7")
    add_bible_slide(prs, directory, "사무엘상", "1:7")
    add_bible_slide(prs, directory, "사무엘상", "1:5", "1:6")
    add_bible_slide(prs, directory, "로마서", "8:28")
    add_bible_slide(prs, directory, "사무엘상", "1:9")
    add_bible_slide(prs, directory, "베드로전서", "5:7")
    add_bible_slide(prs, directory, "사무엘상", "1:10")
    add_bible_slide(prs, directory, "마태복음", "11:28")
    add_bible_slide(prs, directory, "사무엘상", "1:11")
    add_bible_slide(prs, directory, "마태복음", "26:39")
    add_bible_slide(prs, directory, "사무엘상", "1:15")
    add_bible_slide(prs, directory, "빌립보서", "4:6", "4:7")
    add_bible_slide(prs, directory, "사무엘상", "1:19", "1:20")
    add_bible_slide(prs, directory, "창세기", "8:1")
    add_bible_slide(prs, directory, "창세기", "19:29")
    add_bible_slide(prs, directory, "출애굽기", "2:24")
    add_bible_slide(prs, directory, "베드로후서", "3:9")
    add_bible_slide(prs, directory, "야고보서", "5:16")
    add_bible_slide(prs, directory, "역대하", "7:14")
    add_bible_slide(prs, directory, "이사야", "40:31")

    add_hymn_slide(prs, '내 기도하는 그 시간')
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
