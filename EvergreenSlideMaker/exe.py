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

hymn_list = ['주님과 같이', '사랑합니다 나의 예수님', '주 안에서 기뻐해', '십자가 군병들아', '불을 내려주소서', '이 땅의 동과 서 남과 북']

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
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '대표기도')
    add_blank_slide(prs)
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_blank_slide(prs)
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "사도행전", "9:15")
    add_subtitle_slide(prs, input_text="택한 나의 그릇이라 (사도행전 9:15)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "사도행전", "9:6")        # 6절>
    add_bible_slide(prs, directory, "창세기", "50:20")
    add_bible_slide(prs, directory, "열왕기상", "19:9")
    add_bible_slide(prs, directory, "요한복음", "21:17")
    add_bible_slide(prs, directory, "로마서", "8:28")
    
    add_bible_slide(prs, directory, "사도행전", "9:15")       # 15절>
    add_bible_slide(prs, directory, "로마서", "9:15", "9:16")
    add_bible_slide(prs, directory, "디모데후서", "2:20")
    add_bible_slide(prs, directory, "이사야", "6:7", "6:8")
    add_bible_slide(prs, directory, "고린도전서", "15:10")
    
    add_bible_slide(prs, directory, "사도행전", "9:18")       # 18절>
    add_bible_slide(prs, directory, "고린도전서", "15:8")
    add_bible_slide(prs, directory, "에베소서", "3:8")
    add_bible_slide(prs, directory, "디모데전서", "1:15")
    
    add_bible_slide(prs, directory, "사도행전", "9:15")       # 15절> (다시)
    add_bible_slide(prs, directory, "사도행전", "9:17")       # 17절>
    
    add_bible_slide(prs, directory, "야고보서", "2:17")
    add_bible_slide(prs, directory, "누가복음", "5:5")
    
    add_bible_slide(prs, directory, "고린도전서", "9:22")
    add_bible_slide(prs, directory, "빌립보서", "3:13")
    

    # add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs, hymn_list[6])
    # add_card_slide(prs, input_text= '성찬')

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '오늘 숨을 쉬는 것 감사')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
