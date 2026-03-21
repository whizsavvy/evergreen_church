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

hymn_list = ['하나님이 세상을 사랑하사', '주 임재 안에서', '임재', '베드로의 고백', '내 영혼이 은총 입어', '예수 만물의 주', '주 하나님 지으신 모든 세계', '예수 나의 산 소망']

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
    add_blank_slide(prs)
    
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs, hymn_list[7])

    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864", title="내 맘에 한 노래 있어")
    add_bible_slide(prs, directory, "데살로니가전서", "5:23")
    add_subtitle_slide(prs, input_text="행복한 삶 (살전 5:23)")

    add_bible_slide(prs, directory, "요한복음", "12:31", "12:32")
    add_bible_slide(prs, directory, "베드로전서", "3:22")
    add_bible_slide(prs, directory, "창세기", "3:8")
    add_bible_slide(prs, directory, "창세기", "3:12")
    add_bible_slide(prs, directory, "창세기", "5:5")
    add_bible_slide(prs, directory, "요한복음", "5:24")
    add_bible_slide(prs, directory, "빌립보서", "2:12")
    add_bible_slide(prs, directory, "로마서", "13:11")
    add_bible_slide(prs, directory, "욥기", "5:18", "5:20")
    add_bible_slide(prs, directory, "베드로전서", "2:2")
    add_bible_slide(prs, directory, "빌립보서", "2:13", "2:16")
    add_bible_slide(prs, directory, "디모데후서", "2:10")



    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '부흥 2000')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
