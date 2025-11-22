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

hymn_list = ['주님 다시 오실 때 까지', '우리 주 하나님', '내게 강같은 평화', '죄에서 자유를 얻게 함은', '엘리야의 날', '주 여호와는 광대하시도다']
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
    add_hymn_slide(prs, hymn_list[1])

    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[2]) 
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    # add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "요한계시록", "1:9", "1:10")
    add_subtitle_slide(prs, input_text="주일성수(계 1:9~10)")
    
    add_bible_slide(prs, directory, "사도행전", "16:6", "16:10")
    add_subtitle_slide(prs, input_text="닫힌문과 열린문(행 16:6~10)")
    
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '하나님의 약속')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
