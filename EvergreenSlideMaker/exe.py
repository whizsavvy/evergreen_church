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

hymn_list = ['우리 때문에', '죄 짐 맡은 우리 구주', '주의 이름 높이며', '주님 한 분만으로', '엘리야의 날']
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
    # add_hymn_slide(prs, hymn_list[1]) 
    add_hymn_slide(prs, hymn_list[2])
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864")
    
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "사도행전", "13:1", "13:3")
    add_subtitle_slide(prs, input_text="닮고 싶은 교회 (사도행전 13:1~3)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "사도행전", "13:1")        # 13:1
    add_bible_slide(prs, directory, "사도행전", "13:2")        # 13:2
    
    add_bible_slide(prs, directory, "출애굽기", "4:20")         # 출 4:20
    add_bible_slide(prs, directory, "갈라디아서", "2:20")       # 갈 2:20
    
    add_bible_slide(prs, directory, "사도행전", "13:2")        # 13:2 (재등장)
    add_bible_slide(prs, directory, "사도행전", "13:3")        # 13:3
    
    add_bible_slide(prs, directory, "빌립보서", "1:18")        # 빌 1:18 (본문 인용 구절)
    add_bible_slide(prs, directory, "마태복음", "28:19", "28:20")  # 마 28:19~20
    add_bible_slide(prs, directory, "마가복음", "16:17", "16:18")  # 막 16:17~18

    
    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, '나 같은 죄인 살리신')

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '창조의 아버지')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
