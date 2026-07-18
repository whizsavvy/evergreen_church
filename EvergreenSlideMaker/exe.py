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

hymn_list = ['내 구주 예수를 더욱 사랑', '태산을 넘어 험곡에 가도', '주 이름 찬양', '저 바다보다도 더 넓고', '주의 진리 위해 십자가 군기']

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
    # add_hymn_slide(prs, hymn_list[2])
    

   
    # add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'2026_신앙고백1.JPG')
    add_image_slide(prs, pic_dic+'2026_신앙고백2.JPG')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])
    add_blank_slide(prs)
    

    add_bible_slide(prs, directory, "로마서", "1:16", "1:17")
    add_subtitle_slide(prs, input_text="복음을 부끄러워하지 않는 교회 (로마서 1:16~17)")
    
    add_bible_slide(prs, directory, "마가복음", "1:15")
    add_bible_slide(prs, directory, "마가복음", "8:38")
    add_bible_slide(prs, directory, "사도행전", "4:20")
    add_bible_slide(prs, directory, "고린도전서", "1:23", "1:24")
    add_bible_slide(prs, directory, "갈라디아서", "6:14")
    add_bible_slide(prs, directory, "빌립보서", "3:8")
    add_bible_slide(prs, directory, "디모데후서", "1:14")
    add_bible_slide(prs, directory, "에베소서", "2:1")
    add_bible_slide(prs, directory, "고린도후서", "5:17")
    add_bible_slide(prs, directory, "누가복음", "19:8")
    add_bible_slide(prs, directory, "사도행전", "1:8")
    add_bible_slide(prs, directory, "빌립보서", "1:21")
    add_bible_slide(prs, directory, "디모데후서", "4:7")
    add_bible_slide(prs, directory, "요한복음", "21:15")
    add_bible_slide(prs, directory, "요한계시록", "2:4")
    add_bible_slide(prs, directory, "요한일서", "4:19")
    add_hymn_slide(prs, '말씀 앞에서')

    # add_hymn_slide(prs, '부름 받아 나선 이 몸')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도')
    add_card_slide(prs, input_text= '광고')
    # add_card_slide(prs, input_text= '파송기도 및 축복')
    # add_hymn_slide(prs, hymn_list[7])
    add_hymn_slide(prs,  '부흥 2000')
    
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
