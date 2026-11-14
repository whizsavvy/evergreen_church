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

hymn_list = ['예수는 내 힘이요',
'하늘에 계신 아버지', '마귀들과 싸울지라'
, '주의 아름다움은 말로 다'
, '거룩 영원히']
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

    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_blank_slide(prs)
   
    add_bible_slide(prs, directory, "사무엘상", "14:24", "14:30")
    add_subtitle_slide(prs, input_text="누구를 위한 열심입니까? (사무엘상 14:24~30)")
    
    add_bible_slide(prs, directory, "고린도전서", "15:58")
    add_bible_slide(prs, directory, "사무엘상", "14:24")
    add_bible_slide(prs, directory, "사무엘상", "14:24")
    add_bible_slide(prs, directory, "사무엘상", "14:25", "14:26")
    add_bible_slide(prs, directory, "사무엘상", "14:28")
    add_bible_slide(prs, directory, "베드로전서", "4:10")
    add_bible_slide(prs, directory, "마가복음", "10:45")
    add_bible_slide(prs, directory, "사무엘상", "14:27")
    add_bible_slide(prs, directory, "사무엘상", "14:29")
    add_bible_slide(prs, directory, "사무엘상", "14:30")
    add_bible_slide(prs, directory, "갈라디아서", "6:9")
    add_bible_slide(prs, directory, "히브리서", "6:10")
    
    
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '나의 기도하는 것 보다')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'2026/{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
