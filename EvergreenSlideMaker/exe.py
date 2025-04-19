exec(open("code/setting.py").read())

hymn_list = ['십자가를 참으신', '더 원합니다', '하늘 위에 주님 밖에', '우리를 죄에서 구하시려', '주 앙망하는 자', '마라나타', '주님은 나의 사랑']

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

    add_card_slide(prs, input_text= '신앙고백', background_color='000000')
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
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "요한복음", "10:17", "10:18")
    add_subtitle_slide(prs, input_text='내가 내 목숨을 버리는 것은(요 10:17~18)')
    add_blank_slide(prs)
    add_bible_slide(prs, directory, "요한복음", "1:5", "1:6")
    add_bible_slide(prs, directory, "요한복음", "10:41")
    add_bible_slide(prs, directory, "요한복음", "10:22")

    add_bible_slide(prs, directory, "요한복음", "2:19", "2:21")
    add_bible_slide(prs, directory, "요한복음", "10:31")

    add_bible_slide(prs, directory, "요한복음", "10:23")
    add_bible_slide(prs, directory, "요한복음", "10:15")
    add_bible_slide(prs, directory, "요한복음", "10:17")
    add_bible_slide(prs, directory, "요한복음", "10:36")
    add_bible_slide(prs, directory, "요한복음", "10:40")
    add_bible_slide(prs, directory, "출애굽기", "12:5", "12:6")

    add_subtitle_slide(prs, input_text='가나의 혼인잔치에서 물을 포도주로 변화시키심 (요 2:1-11)')
    add_subtitle_slide(prs, input_text='왕의 신하의 아들을 말씀으로 고치심 (요 4:46-54)')
    add_subtitle_slide(prs, input_text='베데스다 연못에서 38년 된 병자를 고치심 (요 5:1-15)')
    add_subtitle_slide(prs, input_text='오병이어로 5천 명을 먹이심 (요 6:1-15)')
    add_subtitle_slide(prs, input_text='물위를 걸으시고')
    add_subtitle_slide(prs, input_text='날 때부터 맹인의 눈을 뜨게 하심 (요 9:1-41)')
    add_subtitle_slide(prs, input_text='나사로를 죽음에서 살리심 (요 11:1-44)')

    add_bible_slide(prs, directory, "요한복음", "10:28", "10:29")
    add_bible_slide(prs, directory, "요한복음", "5:39")
    add_bible_slide(prs, directory, "요한복음", "20:31")
    add_bible_slide(prs, directory, "고린도후서", "4:7")
    add_bible_slide(prs, directory, "고린도후서", "4:10")
    add_bible_slide(prs, directory, "요한복음", "6:40")

    # add_hymn_slide(prs, '나는 주를 섬기는 것에 후회가 없습니다')
    # add_image_slide(prs)
    # add_card_slide(prs, input_text= '성찬')
    # add_hymn_slide(prs,  '나 같은 죄인 살리신')
    add_hymn_slide(prs, hymn_list[6])
    add_card_slide(prs, input_text= '통성기도', background_color='000000')
    add_card_slide(prs, input_text= '광고')
    add_hymn_slide(prs,  '하늘에 계신(주기도문)')
    add_card_slide(prs, input_text= '축도')

    prs.save(F'{today}_늘푸른교회_.pptx')

create_presentation(hymn_list)
