import os
import sys
import tkinter
import customtkinter
from tkinter.filedialog import askopenfilename
from StockInfo import StockInfo

root = customtkinter.CTk()

customtkinter.set_appearance_mode("Dark")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"

class App(customtkinter.CTk):



    def __init__(self):
        super().__init__()

        self.stock_file_path = ''
        self.order_file_path = ''
        self.stock_info = None

        # configure window
        self.title("Sionny Doumy")
        self.geometry(f"{1220}x{680}")

        # configure grid layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure((0,0), weight=0)
        self.grid_columnconfigure((0,1), weight=1)
        self.grid_rowconfigure(1, weight=0)
        self.grid_columnconfigure((0,0), weight=0)
        self.grid_columnconfigure((0,1), weight=0)

        self.terminate_button = customtkinter.CTkButton(self, width=50, height=30, text="종료", command=self.terminate)
        self.terminate_button.grid(row=1, column=1, padx=(20, 40), pady=(0, 20), sticky="se")

        # create tabview
        self.tabview = customtkinter.CTkTabview(self, width=700, height=600)
        self.tabview.grid(row=0, column=0, padx=(40, 20), pady=(40, 40), sticky="nsew")
        self.tabview.add("재고체크")

        self.tabview.tab("재고체크").grid_rowconfigure((0,1,2,3), weight=1)
        self.tabview.tab("재고체크").grid_columnconfigure((0,1), weight=1)
        # row 0
        self.stock_file_path_text = customtkinter.CTkLabel(self.tabview.tab("재고체크"), width=510, height=30, fg_color=("white","gray75"), text_color="black", text="재고문서 경로를 찾아주세요...")
        self.stock_file_path_text.grid(row=0, column=0, padx=(10, 5), pady=(20, 20))
        self.stock_file_path_search_button = customtkinter.CTkButton(self.tabview.tab("재고체크"), width=30, height=30, text="...", command=self.set_stock_path)
        self.stock_file_path_search_button.grid(row=0, column=1, padx=(5, 10), pady=(10, 5))
        # row 1
        self.order_file_path_text = customtkinter.CTkLabel(self.tabview.tab("재고체크"), width=510, height=30, fg_color=("white","gray75"), text_color="black", text="오더리스트문서 경로를 찾아주세요...")
        self.order_file_path_text.grid(row=1, column=0, padx=(10, 5), pady=(20, 20))
        self.order_file_path_search_button = customtkinter.CTkButton(self.tabview.tab("재고체크"), width=30, height=30, text="...", command=self.set_order_path)
        self.order_file_path_search_button.grid(row=1, column=1, padx=(5, 10), pady=(5, 5))
        
        # row 2
        self.file_check_button = customtkinter.CTkButton(self.tabview.tab("재고체크"), width=30, height=30, text="새주문 & 바잉리스트 템플릿 생성", command=self.execute_stock_order, state='disabled')
        self.file_check_button.grid(row=2, column=0, columnspan=2, padx=(100, 100), pady=(30, 30))
        '''
        # row 3
        self.buttonsFrame = customtkinter.CTkFrame(self.tabview.tab("재고체크"), width=700, height=230)
        self.buttonsFrame.grid_rowconfigure(0, weight=1)
        self.buttonsFrame.grid_columnconfigure((0,1), weight=1)
        self.buttonsFrame.grid(row=3, column=0, padx=(10, 10), pady=(10, 10), columnspan=2, sticky='s')
        self.highlight_maker_button = customtkinter.CTkButton(self.buttonsFrame, width=200, height=100, text="재고 있는 행\n 하이라이트된 문서 생성", command=self.make_new_order_file, state='disabled')
        self.highlight_maker_button.grid(row=0, column=0, padx=(100, 50), pady=(30, 30))
        self.buying_list_maker_button = customtkinter.CTkButton(self.buttonsFrame, width=200, height=100, text="바잉리스트 문서 생성", command=self.make_buying_list_file, state='disabled')
        self.buying_list_maker_button.grid(row=0, column=1, padx=(50, 100), pady=(30, 30))
        '''

        self.tabview.add("영수증입출고")
        self.tabview.tab("영수증입출고").grid_columnconfigure(0, weight=1)

        self.label_tab_2 = customtkinter.CTkLabel(self.tabview.tab("영수증입출고"), text="준비중...")
        self.label_tab_2.grid(row=0, column=0, padx=20, pady=20)

        # create console
        self.console = customtkinter.CTkTextbox(self, width=400, height=600)
        self.console.grid(row=0, column=1, padx=(20, 40), pady=(40, 40), sticky="nsew")

        self.printt("Console\n\n" + "유용한 정보가 출력되는 공간입니다.\n\n")

    def printt(self, text):
        self.console.insert(tkinter.END, f'{text}\n')
        self.console.see(tkinter.END)

    def terminate(self):
        self.destroy()
        sys.exit(0)

    def check_and_enable_file_check_button(self):
        if self.stock_file_path != '' and self.order_file_path != '':
            self.file_check_button.configure(state="normal")

    def set_stock_path(self):
        self.stock_file_path = askopenfilename(filetypes=[('Excel Files','*.xlsx')])
        self.stock_file_path_text.configure(text=os.path.abspath(self.stock_file_path))
        self.check_and_enable_file_check_button()

    def set_order_path(self):
        self.order_file_path = askopenfilename(filetypes=[('Excel Files','*.xlsx')])
        self.order_file_path_text.configure(text=os.path.abspath(self.order_file_path))
        self.check_and_enable_file_check_button()

    def execute_stock_order(self):

        try:
            self.stock_info = StockInfo(self.stock_file_path, self.order_file_path)
            self.stock_info.execute()

            self.printt("\n새주문 & 바잉리스트 완료.")

            #self.highlight_maker_button.configure(state="normal")
            #self.buying_list_maker_button.configure(state="normal")
        except Exception as e:
            #self.highlight_maker_button.configure(state="disabled")
            #self.buying_list_maker_button.configure(state="disabled")
            self.printt(str(e))

    '''
    def make_new_order_file(self):
        self.printt("재고체크: 하이라이트된 문서 생성중입니다...")
        try:
            self.stock_info.make_highlighted_new_order_sheet()
            self.printt("재고체크: 하이라이트된 문서가 생성되었습니다!")
        except Exception as e:
            self.printt(f'오류: {str(e)}')

    def make_buying_list_file(self):
        self.printt("재고체크: 바잉리스트 문서 생성은 아직 지원되지 않습니다.")
    '''

if __name__ == "__main__":
    app = App()
    app.mainloop()