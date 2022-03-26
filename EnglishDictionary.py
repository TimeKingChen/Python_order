import xlrd,time,random,os
from aip import AipSpeech
from pydub import AudioSegment
from pydub.playback import play
AudioSegment.converter = os.getcwd() + '/ffmpeg.exe'

APP_ID = '25322432'
API_KEY = 'fUneEWdu6CtTdv8qgbnoCUjH'
SECRET_KEY = 'WRW4A3rBV9iUl0x1jvGbiaPqPbEu6koz'
client = AipSpeech(APP_ID, API_KEY, SECRET_KEY)
word_all_copy = []
word_all_lower = []
def crate():
    global word_all_copy
    global word_all_lower
    global word_all
    read_book = xlrd.open_workbook("人教版初中英语.xls")
    word_all = []
    word_all_lower = []
    sheet = read_book.sheets()
    for sheet_num in range(0,len(sheet)):
        sheet = read_book.sheets()[sheet_num]
        for word_list in range(1,int(sheet.nrows)):
            is_phrase = False
            word_every = sheet.cell_value(word_list, 0)
            if not '.' in sheet.cell_value(word_list, 1):
                is_phrase = True
                word_all.append({'单词':word_every,'词性':'','中文':sheet.cell_value(word_list, 1)})
                word_all_lower.append({'单词':word_every.lower(),'词性':'','中文':sheet.cell_value(word_list, 1)})
            if is_phrase == False:
                a = sheet.cell_value(word_list, 1).split('.')
                word_all.append({'单词':word_every,'词性':a[0]+'.','中文':a[1]})
                word_all_lower.append({'单词':word_every.lower(),'词性':a[0]+'.','中文':a[1]})
    word_all_copy = word_all
    random.shuffle(word_all_copy)
def main_ask():
    crate()
    while True:
        try:
            home_main = input('''
            ==============================================
                      可以用 ‘退出’ 二字退出哦
            ==============================================
            | 1 排序        |           |2 wordle        |
            |               |  welcome! |                |
            |   (顺序排列)   |    to     |   (长度选择)   |
            |      or       |    the    |      and       |
            |   (倒序排列)   |   Enlish  |   (完美复刻)   |
            |===============| dictionary|================|
            |3 查找         |    and    | 4 朗读功能      |
            |               |  wordle   |                |
            |    (字母查找)  |           |      (百度)    |
            |       and     |   enjoy   |                |
            |    (长度查找)  |    it     |   (不要乱输哦)  |
            ==============================================
            |your input: ''')
            home_main_int = int(home_main)
            if home_main_int < 0 or home_main_int > 4:
                print('opps,try again')
                time.sleep(1)
            else:
                if home_main_int == 1:
                    order()
                if home_main_int == 2:
                    wordle()
                if home_main_int == 3:
                    find()
                if home_main_int == 4:
                    read()
        except:
            if home_main == '退出':
                print('bye!')
                time.sleep(1)
                break
            time.sleep(1)
def wordle():
    while 1:
        global word_all
        print('''
        猜字符串游戏
        规则：输入你要猜的字符串长度，计算机会从初中英语词库中选一个单词。你可以输入符合这个长度初中英文单词来猜
        1. 如果你输入的字符串中的字母在目标字符串中并且位置也正确，计算机会在这一位显示√
        2. 如果字母正确，位置不对，计算机会在该字母位置显示-
        3. 如果目标字符串没有这个字母，计算机会显示x
        ''')
        word_length = int(input('请输入你要猜的字符串长度：'))
        answer = ''
        fuhetiaojiandanci = [i['单词'] for i in word_all if len(i['单词']) == word_length]
        current_notice = []
        guess_frequency = 0
        guess_shengyucishu=0
        for i in range(word_length):
            current_notice.append('x')
        answer += random.choice(fuhetiaojiandanci)
        print(answer + '参考')
        while 'x' in current_notice or '-' in current_notice:
            guess_word = input('请输入你猜的字符串: ')
            if len(guess_word) != word_length:
                print('输入的字符串长度不正确')
            elif guess_word not in fuhetiaojiandanci:
                print("请重新输入一个初中单词")
            else:
                for i, guess_letter in enumerate(guess_word):
                    if guess_letter == answer[i]:
                        current_notice[i] = '√'
                    elif guess_letter in answer:
                        current_notice[i] = '-'
                    else:
                        current_notice[i] = 'x'
                print(guess_word)
                print(' '.join(current_notice))
                guess_frequency +=1
                guess_shengyucishu =6-guess_frequency
                if guess_frequency == 6:
                    print("次数已用完，你输了")
                    break
                else:
                    print(f'您已输入{guess_frequency}次，还有{guess_shengyucishu}次机会')
        print(f"""                                 
                               ============================
                                          游戏结束
                                        本次共用了{guess_frequency}次
                                    已超越全球99%的玩家
                               ============================
                        """)
        time.sleep(1)
        while True:
            jixu=input('''
                               ============================
                                     请问要再玩一局吗？
                               1.继续游戏           2.我不玩了
                               ============================
                               |your input:''')
            if jixu == "1":
                break
            else:
                print('''
                               ============================
                                   谢谢游玩，欢迎下次继续
                               ============================
                ''')
                break
        if jixu == 2:
            break
def order():
    global word_all_copy
    global word_all_lower
    while 1:
        order_request = input('''
                =============================
                |  1.顺序排列（区分大小写）   |
                |---------------------------|
                |  2.倒序排列（区分大小写）   |
                |---------------------------|
                |  3.顺序排列（不区分大小写） |
                |---------------------------|
                |  4.倒序排列（不区分大小写） |
                |---------------------------|
                |  5.返回                   |
                ============================
                |your input: ''')
        if order_request == '1':
            order_reverse = sorted(word_all_copy, key = lambda e:e.__getitem__('单词'))
            for number_of_pages in range(0,len(order_reverse),10):
                for order_show in range(number_of_pages,number_of_pages+10):
                    print(order_reverse[order_show])
                    time.sleep(0.001)
                ask_order = input('''
                        ======================
                        |下一页？（yes or no）|
                        ======================
                        | your input: ''')
                if ask_order == 'yes':
                    continue
                else:
                    break
        elif order_request == '2':
            order_reverse = sorted(word_all_copy, key = lambda e:e.__getitem__('单词'),reverse = True)
            for number_of_pages in range(0,len(order_reverse),10):
                for order_show in range(number_of_pages,number_of_pages+10):
                    print(order_reverse[order_show])
                    time.sleep(0.001)
                ask_order = input('''
                        ======================
                        |下一页？（yes or no）|
                        ======================
                        | your input: ''')
                if ask_order == 'yes':
                    continue
                else:
                    break
        elif order_request == '3':
            order_reverse = sorted(word_all_lower, key = lambda e:e.__getitem__('单词'))
            for number_of_pages in range(0,len(order_reverse),10):
                for order_show in range(number_of_pages,number_of_pages+10):
                    print(order_reverse[order_show])
                    time.sleep(0.001)
                ask_order = input('''
                        ======================
                        |下一页？（yes or no）|
                        ======================
                        | your input: ''')
                if ask_order == 'yes':
                    continue
                else:
                    break
        elif order_request == '4':
            order_reverse = sorted(word_all_lower, key = lambda e:e.__getitem__('单词'),reverse = True)
            for number_of_pages in range(0,len(order_reverse),10):
                for order_show in range(number_of_pages,number_of_pages+10):
                    print(order_reverse[order_show])
                    time.sleep(0.001)
                ask_order = input('''
                        ======================
                        |下一页？（yes or no）|
                        ======================
                        | your input: ''')
                if ask_order == 'yes':
                    continue
                else:
                    break
        elif order_request == '5':
            break
        else:
            print('oops,may be it is finish,so u can try again?')
def find():
    answer = input('''
        =======================
        |   1.按首字母查询     |
        =======================
        |   2.包含字母就查询   |
        =======================
        |   3.长度查找         |
        =======================
        |   4.退出            |
        =======================
        |your input: ''')
    if answer == '1':
        word_first_letter = input('''
                  输入你想要的首字母
            ===========================
            |your input: ''')
        word_list_right = []
        for every_word in word_all:
            find_word_head_letter = every_word['单词']
            if find_word_head_letter[0] == word_first_letter:
                word_list_right.append(every_word)
        for numbers in range(0,len(word_list_right),10):
            for order_show in range(numbers,numbers+10):
                    print(word_list_right[order_show])
                    time.sleep(0.001)
            answer_second = input('''
                ====================
                |      下一页?      |
                |    (yes or no)   |
                |===================
                |your input: ''')
            if answer_second == 'yes':
                continue
            else:
                break
    if answer == '2':
        word_first_letter = input('''
                  输入你想要的包含字母
            ===========================
            |your input: ''')
        word_list_right = []
        for every_word in word_all:
            find_word_head_letter = every_word['单词']
            for every_letter in range(0,len(find_word_head_letter)):
                if find_word_head_letter[every_letter] == word_first_letter:
                    word_list_right.append(every_word)
                    break
        for numbers in range(0,len(word_list_right),10):
            for order_show in range(numbers,numbers+10):
                    print(word_list_right[order_show])
                    time.sleep(0.001)
            answer_second = input('''
                ====================
                |      下一页?      |
                |    (yes or no)   |
                |===================
                |your input: ''')
            if answer_second == 'yes':
                continue
            else:
                break
    if answer == '3':
        letter_length = input('''
            =============================
                输入你想要查找的单词长度
            =============================
            |your input: ''')
        if letter_length.isdigit():
            letter_length = int(letter_length)
            filter_words = [w for w in word_all if len(w['单词']) == letter_length]
            for numbers in range(0,len(filter_words),10):
                for order_show in range(numbers,numbers+10):
                        print(filter_words[order_show])
                        time.sleep(0.001)
                answer_second = input('''
                    ====================
                    |      下一页?      |
                    |    (yes or no)   |
                    |===================
                    |your input: ''')
                if answer_second == 'yes':
                    continue
                else:
                    break
        else:
            print('Your input is not a number')
def read():
    text_to_voice(input('''
                =============================
                |      输入一个英文单词       |
                =============================
                | your input: '''))
def text_to_voice(text):
    result  = client.synthesis(text, 'zh', 1, {
        'vol': 5,
    })
    #text = text.replace(' ', '')
    if not isinstance(result, dict):
        with open('{}-result.mp3'.format(text), 'wb') as f:
            f.write(result)
        return play_voice('{}-result.mp3'.format(text))
def play_voice(file_input_path):
    #print(file_input_path)
    if file_input_path.endswith('.wav'):
        song = AudioSegment.from_wav(file_input_path)
    elif file_input_path.endswith('.mp3'):
        song = AudioSegment.from_mp3(file_input_path)
    #print(dir(song))
    if song != None:
        play(song)
    else:
        print('may be have something wrong?~')
if __name__ == '__main__':
    main_ask()