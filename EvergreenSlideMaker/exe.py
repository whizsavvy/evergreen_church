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

hymn_list = ['우리 보좌 앞에 모였네', '죄에서 자유를 얻게 함은', '나의 등 뒤에서', '주와 같이 길 가는 것', '영광의 이름 예수', '엘리야의 날', '마라나타', '하나님 한번도 나를']

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
    add_hymn_slide(prs, hymn_list[1    ])
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "사무엘상", "4:1", "4:11")
    add_subtitle_slide(prs, input_text="그것으로 우리를 구원하게 하자 (사무엘상 4:1~11)")
    
    add_bible_slide(prs, directory, "사무엘상", "4:3")
    add_bible_slide(prs, directory, "시편", "13:1")
    add_bible_slide(prs, directory, "사무엘상", "4:3")
    add_bible_slide(prs, directory, "하박국", "3:17", "3:18")
    add_bible_slide(prs, directory, "사무엘상", "4:5")
    add_bible_slide(prs, directory, "사무엘상", "4:11")
    add_bible_slide(prs, directory, "사무엘상", "5:3")
    add_bible_slide(prs, directory, "사무엘상", "5:4")
    add_bible_slide(prs, directory, "마태복음", "5:14")
    add_bible_slide(prs, directory, "마태복음", "16:18")
    add_bible_slide(prs, directory, "시편", "27:4")
    add_bible_slide(prs, directory, "요한복음", "15:5")
    add_bible_slide(prs, directory, "베드로전서", "3:18")
    add_bible_slide(prs, directory, "시편", "23:4")
    add_bible_slide(prs, directory, "빌립보서", "1:21")
    add_bible_slide(prs, directory, "마태복음", "5:16")

    add_hymn_slide(prs, '하나님 한번도 나를')
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
