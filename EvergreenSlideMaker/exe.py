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

hymn_list = ['천사들의 노래가 하늘에서 울리네', '예수 열방의 소망']
def create_presentation(hymn_list=[]):
    prs = Presentation()
    prs.slide_width = Cm(33.867)
    prs.slide_height = Cm(19.05)
    directory = folder_path+"/bible"
    pic_dic = folder_path+"/image/"
    # add_image_slide(prs, pic_dic+'2025.jpg', text='주일 1부 예배')
    # add_image_slide(prs, pic_dic+'2025.jpg', text='주일 2부 예배')
    add_image_slide(prs, pic_dic+'2025.jpg', text='성탄감사예배')
    add_blank_slide(prs)
    add_hymn_slide(prs, hymn_list[0])
    add_hymn_slide(prs, hymn_list[1])
    # add_hymn_slide(prs, hymn_list[2])
    add_card_slide(prs, input_text= '성가대 찬양')
    # add_choir_slides_from_file(prs, box_color="203864")
    
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    # add_hymn_slide(prs, hymn_list[1]) 
    # add_hymn_slide(prs, hymn_list[3])
    # add_hymn_slide(prs, hymn_list[4])
    # add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
    
   
        # 제목/본문 (필수)
    add_bible_slide(prs, directory, "마태복음", "1:21", "1:23")
    add_subtitle_slide(prs, input_text="임마누엘로 오신 예수님 (마태복음 1:21~23)")
    
    # RED only — 원고 순서 그대로
    add_bible_slide(prs, directory, "누가복음", "2:12")      # “강보에 싸여 구유에…”
    add_bible_slide(prs, directory, "마태복음", "1:23")      # “임마누엘이라 하리라”
    
    # (①~③ 단락 전개 순서)
    add_bible_slide(prs, directory, "마태복음", "1:21")      # 자기 백성을 죄에서 구원
    add_bible_slide(prs, directory, "마태복음", "1:22")      # 말씀의 성취
    add_bible_slide(prs, directory, "이사야", "7:14")        # 처녀 잉태 예언
    add_bible_slide(prs, directory, "마태복음", "1:23")      # 임마누엘 재강조
    add_bible_slide(prs, directory, "마태복음", "28:20")     # “세상 끝날까지…”
    
    # 결론부 인용 순서
    add_bible_slide(prs, directory, "히브리서", "2:18")
    add_bible_slide(prs, directory, "히브리서", "13:5")
    add_bible_slide(prs, directory, "마태복음", "28:20")     # 결론에서 한 번 더
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '창조의 아버지')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
