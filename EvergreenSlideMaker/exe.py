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

hymn_list = ['시선', '아 하나님의 은혜로', '난 예수가 좋다오', '은혜', '말씀 앞에서', '보혈을 지나', '비 준비하시니']

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
    add_bible_slide(prs, directory, "하박국", "3:1", "3:2")
    add_subtitle_slide(prs, input_text='진노 중에라도 긍휼을 잊지 마옵소서(하박국 3:1–2)')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "하박국", "1:2")
    add_bible_slide(prs, directory, "하박국", "1:13")
    
    
    add_subtitle_slide(prs, input_text='1. 하나님의 답변 앞에 놀라는 하박국')
    add_bible_slide(prs, directory, "하박국", "1:5", "1:11")
    
    add_bible_slide(prs, directory, "하박국", "1:6")
    add_bible_slide(prs, directory, "하박국", "1:11")
    add_bible_slide(prs, directory, "하박국", "1:13")

    add_subtitle_slide(prs, input_text='2. 내가 파수하는 곳에 서리라')
    
    add_bible_slide(prs, directory, "하박국", "2:1")
    add_bible_slide(prs, directory, "로마서", "4:18", "4:20")
    add_bible_slide(prs, directory, "시편", "73:16", "73:17")
    add_bible_slide(prs, directory, "잠언", "16:9")

    add_subtitle_slide(prs, input_text='3. 하나님의 응답하심. 의인은 믿음으로 산다')
    add_bible_slide(prs, directory, "하박국", "2:4")
    add_bible_slide(prs, directory, "하박국", "2:14")

    add_subtitle_slide(prs, input_text='4.여호와여, 내가 주의 소문을 듣고 놀랐나이다')
    add_bible_slide(prs, directory, "하박국", "3:2")


    add_subtitle_slide(prs, input_text='5. 주의 일을 이 수년 내에 부흥하게 하옵소서')
    add_bible_slide(prs, directory, "에스겔", "36:26")


    add_subtitle_slide(prs, input_text='6. 진노 중에라도 긍휼을 잊지 마옵소서')
    add_bible_slide(prs, directory, "호세아", "6:1")
    add_bible_slide(prs, directory, "출애굽기", "32:13", "32:14")
    add_bible_slide(prs, directory, "요나", "3:10")
    add_bible_slide(prs, directory, "하박국", "3:17", "3:18")
    add_bible_slide(prs, directory, "하박국", "3:19")
    
    add_hymn_slide(prs, hymn_list[4])

    add_card_slide(prs, input_text= '성찬')
    add_hymn_slide(prs, hymn_list[5])
    
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs, hymn_list[6])
    # add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
