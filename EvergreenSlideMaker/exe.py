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

hymn_list = ['내 안에 가장 귀한 것', '태산을 넘어 험곡에 가도', '옳은 길 따르라 의의 길을', '주님의 선하심']

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

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])

    add_card_slide(prs, input_text= '성가대 찬양')
    # add_choir_slides_from_file(prs, box_color="203864", title="주 예수 나의 산  소망")
    add_bible_slide(prs, directory, "여호수아", "23:8", "23:11")
    add_subtitle_slide(prs, input_text="하나님께 딱 붙어 있으라 (여호수아 23:8~11)")
    
    add_bible_slide(prs, directory, "여호수아", "23:8")
    add_bible_slide(prs, directory, "창세기", "2:24")
    add_bible_slide(prs, directory, "요한계시록", "2:4")
    add_bible_slide(prs, directory, "시편", "73:28")
    add_bible_slide(prs, directory, "룻기", "1:16")
    add_bible_slide(prs, directory, "여호수아", "23:9", "23:10")
    add_bible_slide(prs, directory, "신명기", "8:17", "8:18")
    add_bible_slide(prs, directory, "사무엘상", "17:47")
    add_bible_slide(prs, directory, "출애굽기", "14:14")
    add_bible_slide(prs, directory, "요한복음", "15:5")
    add_bible_slide(prs, directory, "여호수아", "23:11")
    add_bible_slide(prs, directory, "마태복음", "22:37")
    add_bible_slide(prs, directory, "요한복음", "14:15")
    add_bible_slide(prs, directory, "시편", "1:2")
    add_bible_slide(prs, directory, "마가복음", "1:35")
    add_bible_slide(prs, directory, "시편", "84:10")
    add_bible_slide(prs, directory, "시편", "73:28")
    add_hymn_slide(prs, '주만 바라볼찌라' )


    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '부흥 2000')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
