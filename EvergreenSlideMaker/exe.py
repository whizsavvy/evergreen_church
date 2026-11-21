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

hymn_list = ['놀라우신 은혜', '하늘에 계신 아버지' ,'태산을 넘어 험곡에 가도', '내 영혼이 은총 입어', '모든 열방 주 볼 때까지'
, '내가 주인 삼은 모든 것', '나의 기도하는 것보다']
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

    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
   
    add_bible_slide(prs, directory, "사무엘상", "15:1", "15:9")
    add_subtitle_slide(prs, input_text="기름진 것과 좋은 것은 남기고 (사무엘상 15:1~9)")
    
    add_bible_slide(prs, directory, "사무엘상", "15:3")
    add_bible_slide(prs, directory, "사무엘상", "15:9")
    add_bible_slide(prs, directory, "사무엘상", "15:15")
    add_bible_slide(prs, directory, "사무엘상", "15:21")
    add_bible_slide(prs, directory, "사무엘상", "15:12")
    add_bible_slide(prs, directory, "사무엘상", "15:24")
    add_bible_slide(prs, directory, "사무엘상", "15:22")
    add_bible_slide(prs, directory, "사무엘상", "15:15")
    add_bible_slide(prs, directory, "사무엘상", "15:21")
    add_bible_slide(prs, directory, "사무엘상", "15:24")
    add_bible_slide(prs, directory, "사무엘상", "15:30")

    add_hymn_slide(prs, hymn_list[5])
    
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '나의 기도하는 것보다')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
