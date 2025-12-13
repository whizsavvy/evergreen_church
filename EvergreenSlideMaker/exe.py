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

hymn_list = ['천사들의 노래가', '참 반가운 성도여' , '주님 큰 영광 받으소서', '크신 주님', '영광의 이름 예수', '빛나고 높은 보좌와']
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
    add_hymn_slide(prs, hymn_list[2])
    add_image_slide(prs, pic_dic+'신앙고백.png')
    add_image_slide(prs, pic_dic+'신앙고백_1.png')
    add_image_slide(prs, pic_dic+'신앙고백_2.png')
    # add_card_slide(prs, input_text= '신앙고백', background_color='000000')
    add_blank_slide(prs)
    # add_hymn_slide(prs, hymn_list[1]) 
    add_hymn_slide(prs, hymn_list[3])
    add_hymn_slide(prs, hymn_list[4])
    add_hymn_slide(prs, hymn_list[5])

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '대표기도')
    add_card_slide(prs, input_text= '성가대 찬양')
    add_choir_slides_from_file(prs, box_color="203864")
   
    # 제목/본문 (필수)
    add_bible_slide(prs, directory, "창세기", "47:7", "47:10")
    add_subtitle_slide(prs, input_text="네 나이가 얼마냐 (창세기 47:7~10)")
    
    # RED only — 원고 등장 순서
    add_bible_slide(prs, directory, "창세기", "45:5", "45:8")   # 요셉의 고백
    add_bible_slide(prs, directory, "창세기", "47:8")           # “네 나이가 얼마냐?”
    add_bible_slide(prs, directory, "창세기", "47:9")           # “나그네 길의 세월…”
    
    add_bible_slide(prs, directory, "창세기", "47:7")
    add_bible_slide(prs, directory, "창세기", "47:10")          # 야곱이 바로를 축복
    
    add_bible_slide(prs, directory, "히브리서", "7:7")
    add_bible_slide(prs, directory, "창세기", "27:29")
    add_bible_slide(prs, directory, "빌립보서", "3:7", "3:9")
    
    add_bible_slide(prs, directory, "고린도전서", "15:10")
    add_bible_slide(prs, directory, "에베소서", "5:16")
    add_bible_slide(prs, directory, "시편", "90:10")
    add_bible_slide(prs, directory, "시편", "90:9")



    add_hymn_slide(prs, '왜 나만 겪는 고냔이냐고')
    add_hymn_slide(prs, '내가 어둠속에서')
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs, '나 같은 죄인 살리신')

    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    # add_hymn_slide(prs, hymn_list[6])
    add_hymn_slide(prs,  '창조의 아버지')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
