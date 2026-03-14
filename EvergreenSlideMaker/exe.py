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
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')

    add_bible_slide(prs, directory, "출애굽기", "14:13", "14:14")
    add_subtitle_slide(prs, input_text="하나님이 우리를 당황하게 하실 때 (출애굽기 14:13–14)")
    
    add_bible_slide(prs, directory, "출애굽기", "14:13")
    add_bible_slide(prs, directory, "출애굽기", "14:2")
    add_bible_slide(prs, directory, "출애굽기", "14:4")
    add_bible_slide(prs, directory, "시편", "37:5")
    add_bible_slide(prs, directory, "욥기", "23:10")
    add_bible_slide(prs, directory, "출애굽기", "14:11")
    add_bible_slide(prs, directory, "열왕기상", "19:4")
    add_bible_slide(prs, directory, "출애굽기", "14:13")
    add_bible_slide(prs, directory, "출애굽기", "14:21")
    add_bible_slide(prs, directory, "히브리서", "11:29")
    add_bible_slide(prs, directory, "여호수아", "3:13")
    add_bible_slide(prs, directory, "마태복음", "14:29")
    add_bible_slide(prs, directory, "출애굽기", "14:14")
    add_bible_slide(prs, directory, "예레미야", "32:17")
    add_bible_slide(prs, directory, "말라기", "3:6")
    add_bible_slide(prs, directory, "출애굽기", "14:15")
    add_hymn_slide(prs, '주만 바라볼찌라' )


    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '보혈을 지나')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '부흥 2000')    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
