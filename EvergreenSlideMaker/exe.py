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

hymn_list = ['은혜', '마음 속에 근심 있는 사람', '주 안에 있는 나에게', '멈출 수 없네', '주의 나라가 임할 때']
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
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
    
    add_bible_slide(prs, directory, "사무엘상", "7:3", "7:12")
    add_subtitle_slide(prs, input_text="여기까지, 그리고 다시 시작 (사무엘상 7:3~12)")
    
    add_bible_slide(prs, directory, "사무엘상", "7:12")
    add_bible_slide(prs, directory, "사무엘상", "7:2")
    add_bible_slide(prs, directory, "사무엘상", "7:3")
    add_bible_slide(prs, directory, "누가복음", "15:20")
    add_bible_slide(prs, directory, "시편", "51:10")
    add_bible_slide(prs, directory, "사무엘상", "7:5")
    add_bible_slide(prs, directory, "사무엘상", "7:6")
    add_bible_slide(prs, directory, "사무엘상", "7:8")
    add_bible_slide(prs, directory, "사무엘상", "7:9")
    add_bible_slide(prs, directory, "사무엘상", "7:10")
    add_bible_slide(prs, directory, "잠언", "3:5", "3:6")
    add_bible_slide(prs, directory, "사도행전", "1:14")
    add_bible_slide(prs, directory, "사도행전", "12:5")
    add_bible_slide(prs, directory, "사무엘상", "7:12")
    add_bible_slide(prs, directory, "신명기", "8:2")
    add_bible_slide(prs, directory, "시편", "103:2")
    add_bible_slide(prs, directory, "사무엘상", "17:37")
    add_bible_slide(prs, directory, "고린도전서", "15:10")
    add_bible_slide(prs, directory, "빌립보서", "1:6")
    add_bible_slide(prs, directory, "고린도후서", "1:4")

    add_hymn_slide(prs, '지금까지 지내온 것')
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
