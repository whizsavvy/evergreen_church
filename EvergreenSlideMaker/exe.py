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

hymn_list = ['나의 사랑하는 책', '. Again 1907', '불을 내려주소서', '마음이 상한 자를', '이런교회 되게 하소서']

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
    
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)

    add_hymn_slide(prs, hymn_list[1])
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])




    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "예레미야", "20:9")
    add_subtitle_slide(prs, input_text='불붙은 사명자')
    
    add_bible_slide(prs, directory, "예레미야", "20:9")
    add_bible_slide(prs, directory, "예레미야", "1:5")
    add_bible_slide(prs, directory, "예레미야", "7:3")
    add_bible_slide(prs, directory, "예레미야", "7:23")
    add_bible_slide(prs, directory, "예레미야", "18:11")
    add_bible_slide(prs, directory, "예레미야", "20:7", "20:8")
    add_subtitle_slide(prs, input_text='1) 하나님은 모든 사람이 구원받기를 원하십니다')
    add_bible_slide(prs, directory, "디모데전서", "2:4")
    add_subtitle_slide(prs, input_text='2) 예수님은 우리에게 사명을 위임하셨습니다')
    add_bible_slide(prs, directory, "마태복음", "28:19")
    add_subtitle_slide(prs, input_text='3) 복음을 들을 기회를 주지 않으면 영혼은 잃어버립니다')
    add_bible_slide(prs, directory, "누가복음", "19:10")
    add_bible_slide(prs, directory, "로마서", "10:14")
    add_subtitle_slide(prs, input_text='1) 내 삶이 먼저 복음이 되어야 합니다')
    add_bible_slide(prs, directory, "마태복음", "5:14", "5:16")
    add_subtitle_slide(prs, input_text='2) 관계전도를 실천합시다')
    add_bible_slide(prs, directory, "마태복음", "5:14", "5:16")
    add_bible_slide(prs, directory, "마태복음", "10:32")
    add_bible_slide(prs, directory, "사도행전", "1:8")
                    
    


    # add_card_slide(prs, input_text= '성찬')    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[5])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
