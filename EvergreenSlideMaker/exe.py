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

hymn_list = ['내 이름 아시죠', '하늘 위에 주님 밖에', '주님의 은혜 넘치네', '아무 것도 두려워 말라', '나의 갈 길 다 가도록', '나의 가는 길', '이 땅의 동과 서 남과 북']

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
    add_bible_slide(prs, directory, "이사야", "43:19")
    add_subtitle_slide(prs, input_text='광야에 길이 있습니다. (사 43:19)')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "이사야", "39:6")
    add_bible_slide(prs, directory, "예레미야애가", "1:1")
    add_bible_slide(prs, directory, "이사야", "43:18", "43:19")
    add_bible_slide(prs, directory, "갈라디아서", "2:20")

    add_subtitle_slide(prs, input_text='1. 광야는 끝이 아니라 과정입니다.')
    add_bible_slide(prs, directory, "신명기", "8:2")
    add_bible_slide(prs, directory, "민수기", "14:34")
    add_subtitle_slide(prs, input_text='2. 하나님은 광야에서 새 일을 행하십니다.')
    add_bible_slide(prs, directory, "이사야", "43:18")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_subtitle_slide(prs, input_text='3. 하나님은 길을 없는 광야에 길을 만드십니다.')
    add_bible_slide(prs, directory, "이사야", "40:4")
    add_bible_slide(prs, directory, "호세아", "6:1")
    add_bible_slide(prs, directory, "신명기", "8:16")
    

    # add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
