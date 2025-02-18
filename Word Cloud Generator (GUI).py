# 데이터 처리 및 분석
import pandas as pd

# GUI 애플리케이션
import tkinter as tk
from tkinter import Canvas, Label, Scrollbar, Toplevel, colorchooser, filedialog, messagebox

# 텍스트 시각화
import matplotlib.pyplot as plt
from matplotlib import font_manager
from wordcloud import WordCloud

# 텍스트 패턴 매칭
import re

# 운영 체제 및 파일 시스템 관리
import os

# 이미지 처리
from PIL import Image, ImageDraw, ImageFont, ImageTk

# 웹 브라우저 실행
import webbrowser



# 폰트 경로
font_path = None
font_prop = font_manager.FontProperties(fname=font_path)

# tkinter 윈도우 설정
root = tk.Tk()
root.title("워드 클라우드 생성기")

print("\n [정보] 워드 클라우드 생성기 실행")

# 초기화 변수
file_path = None
df = None
character_dialogues = {}
name_column = None
text_column = None
header_text_color_code = "#000000"
header_background_color_code = "#ffffff"

# 워드 클라우드 기본 설정 변수
cloud_width = 800
cloud_height = 600
wordcloud_background_color_code = None
wordcloud_background_color_var = tk.StringVar(value="white")

# 현재 생성된 워드 클라우드를 저장할 변수
current_wordcloud = None
current_selected_character = None

# 제목 배경색 투명 체크박스용 변수
header_transparent_bg_var = tk.BooleanVar(value=False)

# 워드 클라우드 배경색 투명 체크박스용 변수
wordcloud_transparent_bg_var = tk.BooleanVar(value=False)



# 폰트 파일 선택 함수
def select_font_file():
    global font_path

    font_path = filedialog.askopenfilename(
        title="폰트 파일 선택", 
        filetypes=[("Font File", "*.ttf;*.otf;*.fon;*.pcf")]
    )
    
    if font_path:
        try:
            font_prop = font_manager.FontProperties(fname=font_path)
            print(f"\n [정보] 폰트 로딩: {os.path.basename(font_path)}")
            messagebox.showinfo("폰트 선택", f"선택한 폰트: {os.path.basename(font_path)}")

        except Exception as e:
            print(f"\n [오류] 폰트 파일을 읽을 수 없습니다: {e}")
            messagebox.showerror("오류", f"폰트 파일을 읽을 수 없습니다: {e}")

# 프로그램 정보 표시
def show_program_about():
    show_program_about_text = """워드 클라우드 생성기 (Version 1.0)

Created by (Github) IZH318 in 2025.
"""

    messagebox.showinfo("정보", show_program_about_text)

# 개발자 웹사이트 이동
def open_developer_website():
    # 개발사 웹사이트 URL
    url = "https://github.com/IZH318/"

    # 기본 웹 브라우저에서 URL 열기
    webbrowser.open(url)



# *.csv 파일 선택 함수
def select_file():
    global file_path, df
    
    file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
    
    if file_path:
        try:
            df = pd.read_csv(file_path)
            print(f"\n [정보] *.csv 파일 로딩: {file_path}")
            messagebox.showinfo("안내", f"파일이 성공적으로 로드되었습니다:\n\n{file_path}")
            update_columns()
            
        except Exception as e:
            print(f"\n [오류] 파일을 읽는 데 문제가 발생했습니다: {e}")
            messagebox.showerror("오류", f"파일을 읽는 데 문제가 발생했습니다: {e}")



# *.csv 파일 데이터 전처리 함수
def preprocess_text(text):
    # '\n'을 ' '(띄어쓰기 한 칸)로 바꾸기
    text = text.replace("\\n", " ")

    # 'FFFFFFFF'을 ''(공백)로 바꾸기
    text = text.replace("FFFFFFFF", "")
    
    # 특수 문자 목록 정의
    special_chars = ['…', '"', '「', '」', '『', '』', '【', '】', '~', '♪']
    
    # 특수 문자 제거
    for char in special_chars:
        text = text.replace(char, '')
    
    return text



# 열 선택 메뉴 업데이트
def update_columns():
    if df is not None:
        name_column_menu["menu"].delete(0, "end")
        text_column_menu["menu"].delete(0, "end")
        
        for column in df.columns:
            name_column_menu["menu"].add_command(label=column, command=tk._setit(name_column_var, column))
            text_column_menu["menu"].add_command(label=column, command=tk._setit(text_column_var, column))



# 항목별 데이터 수집 및 리스트박스 업데이트
def set_wordcloud_settings():
    global name_column, text_column
    
    name_column = name_column_var.get()
    text_column = text_column_var.get()

    # 폰트 지정 여부 확인
    if not font_path:
        print(f"\n [오류] 폰트가 지정되지 않음.")
        messagebox.showerror("오류", "폰트가 지정되지 않았습니다.")
        return

    # *.csv 파일 로딩 검사
    if not file_path:
        print(f"\n [오류] *.csv 파일 할당되지 않음.")
        messagebox.showerror("오류", "*.csv 파일을 선택해주세요.")
        return

    # 구분 열 검사
    if not name_column or not text_column:
        print(f"\n [오류] *.csv 파일에서 사용 할 열이 할당되지 않음.")
        messagebox.showerror("오류", "*.csv 에서 사용 할 열을 모두 설정해주세요.")
        return
    
    # 입력 값 불러오기
    cloud_width = width_entry.get()
    cloud_height = height_entry.get()
    max_words = max_words_entry.get()
    header_offset = header_offset_entry.get()
    header_fontsize = header_fontsize_entry.get()
    header_text_y = header_text_y_entry.get()

    # 숫자 입력 여부 확인
    if not (cloud_width.isdigit() and cloud_height.isdigit() and max_words.isdigit() and header_offset.isdigit() and header_fontsize.isdigit() and header_text_y.isdigit()):
        print(f"\n [오류] 가로 크기, 세로 크기, 최대 단어 표시, 머릿글 표시 범위, 머릿글 글자 크기, 제목 세로축 조정 중 숫자로 입력되지 않은 항목이 있음.")
        messagebox.showerror("오류", "다음 항목 중 숫자(정수)값이 아닌 항목이 있습니다.\n\n가로 크기\n세로 크기\n최대 단어 표시\n머릿글 표시 범위\n머릿글 글자 크기\n제목 세로축 조정")
        return

    # 빈칸 확인
    if cloud_width == "" or cloud_height == "" or max_words == "" or header_offset == "" or header_fontsize == "" or header_text_y == "":
        print(f"\n [오류] 가로 크기, 세로 크기, 최대 단어 표시, 머릿글 표시 범위, 머릿글 글자 크기, 제목 세로축 조정 중 빈칸이 입력됨.")
        messagebox.showerror("오류", "다음 항목 중 빈칸이 입력되었습니다.\n\n가로 크기\n세로 크기\n최대 단어 표시\n머릿글 표시 범위\n머릿글 글자 크기\n제목 세로축 조정")
        return
    
    character_dialogues.clear()
    
    for _, row in df.iterrows():
        name = row[name_column]
        text = row[text_column]

        # 텍스트 전처리 적용
        text = preprocess_text(text)

        # 데이터를 딕셔너리에 추가
        character_dialogues.setdefault(name, "")
        character_dialogues[name] += text + " "
        
    update_character_listbox()

# 생성 된 워드 클라우드 항목 리스트로 추가
def update_character_listbox():
    character_listbox.delete(0, tk.END)

    for character in character_dialogues.keys():
        character_listbox.insert(tk.END, character)



# 제목 글자색 선택 함수
def choose_header_text_color():
    global header_text_color_code  # 전역 변수로 사용
    header_text_color_code = colorchooser.askcolor(title="제목 글자색 선택")[1]  # 선택된 색을 header_text_color_code에 저장

    print(f"\n [정보] 제목 글자색: {header_text_color_code}")

# 제목 배경색 선택 함수
def choose_header_background_color():
    global header_background_color_code  # 전역 변수로 사용
    header_background_color_code = colorchooser.askcolor(title="제목 배경색 선택")[1]  # 선택된 색을 header_background_color_code에 저장
        
    print(f"\n [정보] 제목 배경색: {header_background_color_code}")

# 제목 배경색 투명 체크박스 상태에 따라 배경색 선택 버튼 활성화/비활성화
def toggle_header_background_color_button():
    if header_transparent_bg_var.get():  # 체크박스가 체크되면
        print(f"\n [정보] 제목 배경색 투명 활성화")
        header_background_color_button.config(state="disabled")  # 배경색 선택 버튼 비활성화

    else:  # 체크박스가 체크 해제되면
        print(f"\n [정보] 제목 배경색 투명 비활성화")
        header_background_color_button.config(state="normal")  # 배경색 선택 버튼 활성화



# 워드 클라우드 배경색 선택 함수
def choose_wordcloud_background_color():
    global wordcloud_background_color_code  # 전역 변수로 사용

    if not wordcloud_transparent_bg_var.get():  # 투명 체크박스가 선택되지 않았다면
        wordcloud_background_color_code = colorchooser.askcolor(title="워드 클라우드 배경색 선택")[1]  # 선택된 색을 wordcloud_background_color_code에 저장

        if wordcloud_background_color_code:
            wordcloud_background_color_var.set(wordcloud_background_color_code)
        
        print(f"\n [정보] 워드 클라우드 배경색: {wordcloud_background_color_code}")

    else:
        # 투명 배경이 선택된 경우, 배경색을 투명으로 설정
        wordcloud_background_color_code = None  # 배경색을 투명으로 설정
        wordcloud_background_color_var.set(None)

# 워드 클라우드 배경색 투명 체크박스 상태에 따라 배경색 선택 버튼 활성화/비활성화
def toggle_worldcloud_background_color_button():
    if wordcloud_transparent_bg_var.get():  # 체크박스가 체크되면
        print(f"\n [정보] 워드 클라우드 배경색 투명 활성화")
        wordcloud_background_color_button.config(state="disabled")  # 배경색 선택 버튼 비활성화

    else:  # 체크박스가 체크 해제되면
        print(f"\n [정보] 워드 클라우드 배경색 투명 비활성화")
        wordcloud_background_color_button.config(state="normal")  # 배경색 선택 버튼 활성화



# 워드 클라우드 미리보기 이미지 생성
def generate_wordcloud_preview():
    global current_wordcloud, header_text_color_code, header_background_color_code, wordcloud_background_color_code, current_wordcloud, current_selected_character

    # 입력 값 불러오기
    cloud_width = int(width_entry.get())
    cloud_height = int(height_entry.get())
    max_words = int(max_words_entry.get())
    header_offset = int(header_offset_entry.get())
    header_fontsize = int(header_fontsize_entry.get())
    header_text_y = int(header_text_y_entry.get())
    
    # 워드 클라우드 생성
    selected_character = character_listbox.get(tk.ACTIVE)
    if selected_character not in character_dialogues:
        print(f"\n [오류] 선택된 항목의 데이터 없음.")
        messagebox.showerror("오류", "선택된 항목의 데이터가 없습니다.")
        return
    
    dialogues = character_dialogues[selected_character]
    current_selected_character = selected_character

    # 데이터가 비어있는지 확인
    if not dialogues or dialogues.strip() == "":
        print(f"\n [오류] {selected_character}의 데이터가 비어 있음.")
        messagebox.showerror("오류", f"{selected_character}의 데이터가 비어 있습니다.")
        return

    # 워드 클라우드 생성
    try:
        wordcloud = WordCloud(
            font_path=font_path,
            width=cloud_width,
            height=cloud_height - header_offset,
            background_color=wordcloud_background_color_var.get() if not wordcloud_transparent_bg_var.get() else None,
            mode="RGBA",  # RGBA 모드를 사용하여 투명 배경을 처리
            max_words=max_words  # 최대 단어 수 설정
        ).generate(dialogues)

    except ValueError as e:
        messagebox.showerror("오류", f"{e}")
        return
    
    # 워드 클라우드 이미지로 변환
    wordcloud_image = wordcloud.to_image()

    # 머릿글 범위에 맞춰 이미지 크기 조정
    if header_transparent_bg_var.get():  # 투명 체크박스가 선택되었으면
        extended_image = Image.new("RGBA", (cloud_width, cloud_height), (0, 0, 0, 0))  # 투명 배경
    else:
        # 사용자 지정 배경색으로 설정 (wordcloud_background_color_code가 RGB 형태라면 RGBA로 변환)
        if isinstance(header_background_color_code, tuple) and len(header_background_color_code) == 3:
            header_background_color_code = (*header_background_color_code, 255)  # RGBA로 변환 (A=255)
        extended_image = Image.new("RGBA", (cloud_width, cloud_height), header_background_color_code)  # 사용자 지정 배경색
    
    extended_image.paste(wordcloud_image, (0, header_offset))

    # 이미지 상단에 텍스트 추가
    draw = ImageDraw.Draw(extended_image)
    font = ImageFont.truetype(font_path, size=header_fontsize)
    text = selected_character

    # 텍스트 경계 상자를 계산하여 크기 가져오기
    bbox = draw.textbbox((0, 0), text, font=font)
    text_width = bbox[2] - bbox[0]
    text_height = bbox[3] - bbox[1]

    # 텍스트의 중앙 좌표 계산
    text_position = ((cloud_width - text_width) // 2, header_text_y)

    # 텍스트를 이미지에 그리기
    draw.text(text_position, text, font=font, fill=header_text_color_code)

    # 새 창을 열어서 이미지를 표시
    wordcloud_preview_window(extended_image)

# 워드 클라우드 이미지를 새 창에서 미리보기 형태로 표시하는 함수
def wordcloud_preview_window(wordcloud_image):
    # 새로운 Toplevel 창을 생성
    wordcloud_window = tk.Toplevel()
    wordcloud_window.title("Word Cloud Preview")
    
    # 캔버스와 스크롤바를 위한 Frame 생성
    frame = tk.Frame(wordcloud_window)
    frame.pack(fill="both", expand=True)

    # 캔버스 생성
    canvas = Canvas(frame)
    
    # 세로 스크롤바 생성 (이미지 오른쪽에 배치)
    v_scrollbar = Scrollbar(frame, orient="vertical", command=canvas.yview)
    v_scrollbar.pack(side="right", fill="y")

    # 가로 스크롤바 생성 (이미지 하단에 배치)
    h_scrollbar = Scrollbar(frame, orient="horizontal", command=canvas.xview)
    h_scrollbar.pack(side="bottom", fill="x")
    
    # 캔버스 설정
    canvas.config(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
    
    # Tkinter에서 사용할 수 있도록 변환
    wordcloud_image_tk = ImageTk.PhotoImage(wordcloud_image)

    # 캔버스에 이미지 추가
    canvas.create_image(0, 0, image=wordcloud_image_tk, anchor="nw")
    canvas.pack(side="left", expand=True, fill="both")

    # 캔버스의 스크롤 영역 설정
    canvas.config(scrollregion=canvas.bbox("all"))

    # 캔버스의 크기 및 스크롤바를 포함한 창 크기 설정
    canvas.update_idletasks()  # 캔버스에 모든 항목을 그리도록 강제 수행
    bbox = canvas.bbox("all")  # 이미지의 경계 박스 가져오기
    image_width = bbox[2]  # 이미지의 폭
    image_height = bbox[3]  # 이미지의 높이

    # 창 크기 설정 (스크롤바 크기 포함)
    wordcloud_window.geometry(f"{image_width}x{image_height}")

    # 마우스 휠을 사용하여 세로 스크롤할 수 있도록 설정
    def on_mouse_wheel(event):
        if event.delta:
            if event.state & 0x0001:  # 마우스 휠을 좌우로 스크롤하려면 Ctrl 키가 눌렸을 때
                canvas.xview_scroll(-1 * (event.delta // 120), "units")  # 가로로 스크롤
            else:
                canvas.yview_scroll(-1 * (event.delta // 120), "units")  # 세로로 스크롤

    # 키보드 화살표 키로 세로/가로 스크롤하기 위한 설정
    def on_key_press(event):
        if event.keysym == 'Down':
            canvas.yview_scroll(1, 'units')  # 아래로 세로 스크롤
        elif event.keysym == 'Up':
            canvas.yview_scroll(-1, 'units')  # 위로 세로 스크롤
        elif event.keysym == 'Right':
            canvas.xview_scroll(1, 'units')  # 오른쪽으로 가로 스크롤
        elif event.keysym == 'Left':
            canvas.xview_scroll(-1, 'units')  # 왼쪽으로 가로 스크롤

    # 마우스 휠 및 키보드 이벤트 바인딩
    wordcloud_window.bind_all("<MouseWheel>", on_mouse_wheel)  # 마우스 휠 이벤트
    wordcloud_window.bind_all("<KeyPress-Up>", on_key_press)  # 위 화살표 키 이벤트
    wordcloud_window.bind_all("<KeyPress-Down>", on_key_press)  # 아래 화살표 키 이벤트
    wordcloud_window.bind_all("<KeyPress-Left>", on_key_press)  # 왼쪽 화살표 키 이벤트
    wordcloud_window.bind_all("<KeyPress-Right>", on_key_press)  # 오른쪽 화살표 키 이벤트

    # 새 창을 표시
    wordcloud_window.mainloop()




# 워드 클라우드 이미지 파일 저장 시 반각 문자로 할당 된 특수 기호를 전각 문자로 특수 기호 변환
def convert_to_fullwidth_symbol(character):
    fullwidth_map = {
        '\\': '＼',
        '/': '／',
        '*': '＊',
        '?': '？',
        ':': '：',
        '"': '＂',
        '<': '＜',
        '>': '＞',
        '|': '｜'
    }
    
    for symbol, fullwidth_symbol in fullwidth_map.items():
        character = character.replace(symbol, fullwidth_symbol)
    
    return character



# 워드 클라우드 저장 함수
def save_wordcloud():
    global current_wordcloud, header_text_color_code, header_background_color_code, wordcloud_background_color_code, header_text_y
    
    # 리스트박스에서 선택된 항목
    selected_character = character_listbox.get(tk.ACTIVE)

    # 입력 값 불러오기
    cloud_width = int(width_entry.get())
    cloud_height = int(height_entry.get())
    max_words = int(max_words_entry.get())
    header_offset = int(header_offset_entry.get())
    header_fontsize = int(header_fontsize_entry.get())
    header_text_y = int(header_text_y_entry.get())

    # 빈칸 확인
    if cloud_width == "" or cloud_height == "" or max_words == "" or header_offset == "" or header_fontsize == "" or header_text_y == "":
        print(f"\n [오류] 가로 크기, 세로 크기, 최대 단어 표시, 머릿글 표시 범위, 머릿글 글자 크기, 제목 세로축 조정 중 빈칸이 입력됨.")
        messagebox.showerror("오류", "다음 항목 중 빈칸이 입력되었습니다.\n\n가로 크기\n세로 크기\n최대 단어 표시\n머릿글 표시 범위\n머릿글 글자 크기\n제목 세로축 조정")
        return
    
    # 선택된 항목이 없다면 메시지 박스
    if selected_character == '':
        print(f"\n [오류] 저장 대상을 선택하지 않음.")
        messagebox.showerror("오류", "저장할 대상을 선택해주세요.")
        return

    # 항목 데이터 가져오기
    if selected_character in character_dialogues:
        dialogues = character_dialogues[selected_character]
    else:
        print(f"\n [오류] {selected_character}의 데이터를 찾을 수 없음.")
        messagebox.showerror("오류", f"{selected_character}의 데이터를 찾을 수 없습니다.")
        return
    
    # 확인 다이얼로그
    response = messagebox.askquestion("저장 옵션", "선택한 항목만 저장 => [Y]\n전체 항목 모두 저장 => [N]")

    # 오류 발생 여부를 추적하는 변수
    some_error_occurred = False
    
    if response == 'yes':
        # 선택한 항목만 저장
        print(f"\n [정보] 선택한 항목만 저장")
        
        # 전처리된 항목 이름
        valid_character_name = convert_to_fullwidth_symbol(selected_character)
        
        save_path = filedialog.asksaveasfilename(
            defaultextension=".png", 
            filetypes=[("PNG Files", "*.png")],
            initialfile=f"{valid_character_name}_({max_words}_{header_offset}_{header_fontsize}_{header_text_y}).png"  # 전처리된 항목 이름으로 파일 이름 기본 설정
        )
        
        if save_path:
            # 파일 경로에 확장자가 없을 경우 처리
            if not save_path.lower().endswith(".png"):
                save_path += ".png"

            try:
                # 워드 클라우드 생성
                wordcloud = WordCloud(
                    font_path=font_path,
                    width=cloud_width,
                    height=cloud_height - header_offset,
                    background_color=wordcloud_background_color_var.get() if not wordcloud_transparent_bg_var.get() else None,
                    mode="RGBA",  # RGBA 모드를 사용하여 투명 배경을 처리
                    max_words=max_words  # 최대 단어 수 설정
                ).generate(dialogues)
                
                # 워드 클라우드 저장
                wordcloud_image = wordcloud.to_image()

                # 머릿글 범위에 맞춰 이미지 크기 조정
                if header_transparent_bg_var.get():  # 투명 체크박스가 선택되었으면
                    extended_image = Image.new("RGBA", (cloud_width, cloud_height), (0, 0, 0, 0))  # 투명 배경

                else:
                    # 사용자 지정 배경색으로 설정 (wordcloud_background_color_code가 RGB 형태라면 RGBA로 변환)
                    if isinstance(header_background_color_code, tuple) and len(header_background_color_code) == 3:
                        header_background_color_code = (*header_background_color_code, 255)  # RGBA로 변환 (A=255)
                    extended_image = Image.new("RGBA", (cloud_width, cloud_height), header_background_color_code)  # 사용자 지정 배경색
                    
                extended_image.paste(wordcloud_image, (0, header_offset))

                # 이미지 상단에 텍스트 추가
                draw = ImageDraw.Draw(extended_image)
                font = ImageFont.truetype(font_path, size=header_fontsize)
                text = selected_character

                # 텍스트 경계 상자를 계산하여 크기 가져오기
                bbox = draw.textbbox((0, 0), text, font=font)
                text_width = bbox[2] - bbox[0]
                text_height = bbox[3] - bbox[1]

                # 텍스트의 중앙 좌표 계산
                text_position = ((cloud_width - text_width) // 2, header_text_y)

                # 텍스트를 이미지에 그리기
                draw.text(text_position, text, font=font, fill=header_text_color_code)

                # 워드 클라우드 이미지 저장
                extended_image.save(save_path)
                print(f"\n [정보] 워드 클라우드 저장: {save_path}")
                messagebox.showinfo("안내", f"워드 클라우드가 성공적으로 저장되었습니다:\n\n{save_path}")
                
            except ValueError:  # 단어가 1개도 없을 때 발생하는 오류 처리
                # 0KB 더미 파일 생성
                with open(save_path, 'wb') as f:
                    pass  # 아무 내용도 없는 파일을 생성
                
                print(f"\n [경고] 데이터 부족으로 워드 클라우드를 생성할 수 없으므로 0KB 더미 파일 생성.")
                messagebox.showwarning("경고", f"데이터 부족으로 워드 클라우드를 생성할 수 없으므로 0KB 더미 파일이 생성되었습니다.")
                
                return  # 더 이상 진행하지 않음

    elif response == 'no':
        # 전체 항목을 저장
        print(f"\n [정보] 전체 항목 저장")
        
        save_dir = filedialog.askdirectory(title="저장할 폴더 선택")
        
        if save_dir:
            for character, dialogues in character_dialogues.items():
                # 단어가 비어 있지 않은 경우에만 저장
                if not dialogues or dialogues.strip() == "":
                    # 단어가 비어 있는 경우 0KB 더미 파일 생성
                    valid_character_name = convert_to_fullwidth_symbol(character)
                    save_path = f"{save_dir}/{valid_character_name}_({max_words}_{header_offset}_{header_fontsize}).png"
                    
                    # 0KB 더미 파일 생성
                    with open(save_path, 'wb') as f:
                        pass  # 아무 내용도 없는 파일을 생성
                    
                    print(f" [경고] 데이터 부족으로 워드 클라우드를 생성할 수 없으므로 0KB 더미 파일 생성.: {save_path}")
                    continue  # 더미 파일을 생성하고 다음 항목으로 넘어감

                # 단어가 있는 경우 워드 클라우드 생성
                valid_character_name = convert_to_fullwidth_symbol(character)
                save_path = f"{save_dir}/{valid_character_name}_({max_words}_{header_offset}_{header_fontsize}_{header_text_y}).png"  # 파일 이름을 항목 이름으로 지정
                
                # 파일 경로에 확장자가 없을 경우 처리
                if not save_path.lower().endswith(".png"):
                    save_path += ".png"
                
                try:
                    # 워드 클라우드 생성
                    wordcloud = WordCloud(
                        font_path=font_path,
                        width=cloud_width,
                        height=cloud_height - header_offset,
                        background_color=wordcloud_background_color_var.get() if not wordcloud_transparent_bg_var.get() else None,
                        mode="RGBA",  # RGBA 모드를 사용하여 투명 배경을 처리
                        max_words=max_words  # 최대 단어 수 설정
                    ).generate(dialogues)

                    # 워드 클라우드 저장
                    wordcloud_image = wordcloud.to_image()

                    # 머릿글 범위에 맞춰 이미지 크기 조정
                    if header_transparent_bg_var.get():  # 투명 체크박스가 선택되었으면
                        extended_image = Image.new("RGBA", (cloud_width, cloud_height), (0, 0, 0, 0))  # 투명 배경

                    else:
                        # 사용자 지정 배경색으로 설정 (wordcloud_background_color_code가 RGB 형태라면 RGBA로 변환)
                        if isinstance(header_background_color_code, tuple) and len(header_background_color_code) == 3:
                            header_background_color_code = (*header_background_color_code, 255)  # RGBA로 변환 (A=255)
                        extended_image = Image.new("RGBA", (cloud_width, cloud_height), header_background_color_code)  # 사용자 지정 배경색
                        
                    extended_image.paste(wordcloud_image, (0, header_offset))

                    # 이미지 상단에 텍스트 추가
                    draw = ImageDraw.Draw(extended_image)
                    font = ImageFont.truetype(font_path, size=header_fontsize)
                    text = character

                    # 텍스트 경계 상자를 계산하여 크기 가져오기
                    bbox = draw.textbbox((0, 0), text, font=font)
                    text_width = bbox[2] - bbox[0]
                    text_height = bbox[3] - bbox[1]

                    # 텍스트의 중앙 좌표 계산
                    text_position = ((cloud_width - text_width) // 2, header_text_y)

                    # 텍스트를 이미지에 그리기
                    draw.text(text_position, text, font=font, fill=header_text_color_code)

                    # 워드 클라우드 이미지 저장
                    extended_image.save(save_path)
                    print(f"\n [정보] 워드 클라우드 저장: {save_path}")
                    
                except ValueError:  # 단어가 1개도 없을 때 발생하는 오류 처리
                    # 0KB 더미 파일 생성
                    valid_character_name = convert_to_fullwidth_symbol(character)
                    save_path = f"{save_dir}/{valid_character_name}.png"
                    
                    with open(save_path, 'wb') as f:
                        pass  # 아무 내용도 없는 파일을 생성

                    print(f" [경고] 데이터 부족으로 워드 클라우드를 생성할 수 없으므로 0KB 더미 파일 생성.: {save_path}")
                    some_error_occurred = True  # 오류 발생 내역 갱신
                    continue  # 오류가 발생하면 계속해서 다른 항목을 처리합니다.

            if some_error_occurred:
                print(f"\n [경고] 일부 항목을 제외한 워드 클라우드 저장 (문제가 발생한 항목은 0KB 더미 파일로 생성.)")
                messagebox.showwarning("경고", "일부 항목을 제외하고 워드 클라우드가 성공적으로 저장되었습니다.\n\n문제가 발생한 항목은 0KB 더미 파일로 생성되었습니다.")

            else:
                print(f"\n [정보] 모든 워드 클라우드 저장.")
                messagebox.showinfo("안내", "모든 워드 클라우드가 성공적으로 저장되었습니다.")



# 창 크기 및 위치 설정
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
window_width = 600
window_height = 570
position_x = (screen_width - window_width) // 2
position_y = (screen_height - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# 반응형 그리드 설정
root.rowconfigure(0, weight=1)
root.columnconfigure(0, weight=1)

main_frame = tk.Frame(root)
main_frame.grid(row=0, column=0, sticky="nsew")

for i in range(15):  # 15개의 열에 대해 설정
    main_frame.rowconfigure(i, weight=1)
    
for i in range(5):  # 5개의 열에 대해 설정
    main_frame.columnconfigure(i, weight=1)



# 메뉴 바 설정
menubar = tk.Menu(root)
root.config(menu=menubar)

# '프로그램 정보' 메뉴 추가
font_menu = tk.Menu(menubar, tearoff=0)
menubar.add_command(label="프로그램 정보", command=show_program_about)

# '개발자 사이트 이동' 메뉴 추가
font_menu = tk.Menu(menubar, tearoff=0)
menubar.add_command(label="개발자 웹사이트 이동", command=open_developer_website)



# 상단 라벨
setup_label = tk.Label(main_frame, text="설정:")
setup_label.grid(row=0, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

character_listbox_label = tk.Label(main_frame, text="목록:")
character_listbox_label.grid(row=0, column=3, columnspan=2, sticky="nsew", padx=5, pady=5)

character_listbox_label = tk.Label(main_frame, text="목록 필터:")
character_listbox_label.grid(row=1, column=3, sticky="nsew", padx=5, pady=5)



# 위젯 구성
tk.Button(main_frame, text="폰트 파일 불러오기", command=select_font_file).grid(row=1, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

tk.Button(main_frame, text="*.csv 파일 불러오기", command=select_file).grid(row=2, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)



tk.Label(main_frame, text="구분 열 (Ex. 이름)").grid(row=3, column=0, sticky="w", padx=5, pady=5)
name_column_var = tk.StringVar()
name_column_menu = tk.OptionMenu(main_frame, name_column_var, "")
name_column_menu.grid(row=3, column=1, columnspan=2, sticky="e", padx=5, pady=5)

tk.Label(main_frame, text="데이터 열 (Ex. 대사)").grid(row=4, column=0, sticky="w", padx=5, pady=5)
text_column_var = tk.StringVar()
text_column_menu = tk.OptionMenu(main_frame, text_column_var, "")
text_column_menu.grid(row=4, column=1, columnspan=2, sticky="e", padx=5, pady=5)

tk.Button(main_frame, text="워드 클라우드 리스트 생성", command=set_wordcloud_settings).grid(row=5, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)



tk.Label(main_frame, text="가로 크기 (px)").grid(row=6, column=0, sticky="w", padx=5, pady=5)
width_entry = tk.Entry(main_frame)
width_entry.insert(0, "800")
width_entry.grid(row=6, column=1, columnspan=2, sticky="e", padx=5, pady=5)

tk.Label(main_frame, text="세로 크기 (px)").grid(row=7, column=0, sticky="w", padx=5, pady=5)
height_entry = tk.Entry(main_frame)
height_entry.insert(0, "600")
height_entry.grid(row=7, column=1, columnspan=2, sticky="e", padx=5, pady=5)



# 위젯 구성
tk.Label(main_frame, text="최대 단어 표시").grid(row=8, column=0, sticky="w", padx=5, pady=5)
max_words_entry = tk.Entry(main_frame)
max_words_entry.insert(0, "200")  # 기본값을 200으로 설정
max_words_entry.grid(row=8, column=1, columnspan=2, sticky="e", padx=5, pady=5)



# 머릿글 표시 범위 입력란
tk.Label(main_frame, text="머릿글 표시 범위 (px)").grid(row=9, column=0, sticky="w", padx=5, pady=5)
header_offset_entry = tk.Entry(main_frame)
header_offset_entry.insert(0, "50")  # 기본값은 0px
header_offset_entry.grid(row=9, column=1, columnspan=2, sticky="e", padx=5, pady=5)

# 머릿글 글자 크기 입력란
tk.Label(main_frame, text="머릿글 글자 크기 (px)").grid(row=10, column=0, sticky="w", padx=5, pady=5)
header_fontsize_entry = tk.Entry(main_frame)
header_fontsize_entry.insert(0, "40")  # 기본값은 40px
header_fontsize_entry.grid(row=10, column=1, columnspan=2, sticky="e", padx=5, pady=5)



# 제목 글자색 선택 버튼과 제목 배경색 버튼을 반으로 나누어 배치
tk.Label(main_frame, text="이미지 상단 제목 설정: ").grid(row=11, column=0, sticky="w", padx=5, pady=5)

# 머릿글 표시 범위 입력란
tk.Label(main_frame, text="제목 세로축 조정: ").grid(row=11, column=1, sticky="w", padx=5, pady=5)
header_text_y_entry = tk.Entry(main_frame)
header_text_y_entry.insert(0, "0")  # 기본값은 0px
header_text_y_entry.grid(row=11, column=2, sticky="e", padx=5, pady=5)

# 제목 글자색 선택 버튼
header_text_color_button = tk.Button(main_frame, text="글자색", command=choose_header_text_color)
header_text_color_button.grid(row=12, column=1, columnspan=2, sticky="nsew", padx=5, pady=5)

# 제목 배경색 선택 버튼
header_background_color_button = tk.Button(main_frame, text="배경색", command=choose_header_background_color)
header_background_color_button.grid(row=13, column=1, sticky="nsew", padx=5, pady=5)

# 제목 배경색 투명 체크박스
header_transparent_bg_checkbutton = tk.Checkbutton(main_frame, text="투명", variable=header_transparent_bg_var, command=toggle_header_background_color_button)
header_transparent_bg_checkbutton.grid(row=13, column=2, sticky="w", padx=5, pady=5)

# 제목 배경색 선택 버튼과 배경색 투명 체크박스 초기 상태 업데이트
toggle_header_background_color_button()



# 워드 클라우드 배경색 선택 버튼, 배경색 투명 체크박스 배치
tk.Label(main_frame, text="워드 클라우드 설정: ").grid(row=14, column=0, sticky="w", padx=5, pady=5)

# 워드 클라우드 배경색 선택 버튼
wordcloud_background_color_button = tk.Button(main_frame, text="배경색", command=choose_wordcloud_background_color)
wordcloud_background_color_button.grid(row=14, column=1, sticky="nsew", padx=5, pady=5)

# 워드 클라우드 배경색 투명 체크박스
wordcloud_transparent_bg_checkbutton = tk.Checkbutton(main_frame, text="투명", variable=wordcloud_transparent_bg_var, command=toggle_worldcloud_background_color_button)
wordcloud_transparent_bg_checkbutton.grid(row=14, column=2, sticky="w", padx=5, pady=5)

# 워드 클라우드 배경색 선택 버튼과 배경색 투명 체크박스 초기 상태 업데이트
toggle_worldcloud_background_color_button()



# 리스트 박스를 필터링하는 함수
def filter_listbox(event=None):
    search_term = filter_entry.get().lower()  # 텍스트 박스에서 입력된 텍스트를 소문자로 변환
    filtered_items = [character for character in character_dialogues.keys() if search_term in character.lower()]
    
    # 리스트박스를 업데이트
    character_listbox.delete(0, tk.END)
    for character in filtered_items:
        character_listbox.insert(tk.END, character)

# 리스트 박스 필터링 텍스트 박스 생성
filter_entry = tk.Entry(main_frame)
filter_entry.grid(row=1, column=4, sticky="nsew", padx=5, pady=5)

# 워드 클라우드 항목 표시 리스트
listbox_frame = tk.Frame(main_frame)
listbox_frame.grid(row=2, column=3, rowspan=15, columnspan=2, sticky="nsew", padx=5, pady=5)
character_listbox = tk.Listbox(listbox_frame)
character_listbox.pack(side="left", fill="both", expand=True)
scrollbar = tk.Scrollbar(listbox_frame, orient="vertical", command=character_listbox.yview)
scrollbar.pack(side="right", fill="y")
character_listbox.config(yscrollcommand=scrollbar.set)



# 목록에서 더블클릭 시 워드 클라우드 미리보기 생성
def on_listbox_double_click(event):
    generate_wordcloud_preview()

character_listbox.bind("<Double-1>", on_listbox_double_click)

# 워드 클라우드 저장 버튼
save_button = tk.Button(main_frame, text="워드 클라우드 저장", command=save_wordcloud)
save_button.grid(row=15, column=0, columnspan=3, sticky="nsew", padx=5, pady=5)

# 텍스트 박스에 이벤트 바인딩 (키 입력 시마다 리스트박스를 필터링)
filter_entry.bind("<KeyRelease>", filter_listbox)



# 메인 루프 실행
root.mainloop()
