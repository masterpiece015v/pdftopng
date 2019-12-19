import subprocess
import os
from PIL import Image
from PIL import ImageEnhance
import tkinter
from tkinter import ttk
from tkinter import filedialog
from tkinter import StringVar
from tkinter import messagebox
import threading

class MyFram:
    def __init__(self):
        # フォント
        font = ("",12)
        # カレントディレクトリ
        self.cur_dir = os.getcwd()
        self.home_dir = os.getcwd()
        # メインウィンドウ
        self.main_win = tkinter.Tk()
        self.main_win.title('PDFをPNGに変換する')
        self.main_win.geometry('640x480')
        # スタイルの設定
        style = ttk.Style()
        style.configure(".",font=("",12))
        # メインフレーム
        self.main_frm = ttk.Frame(self.main_win)
        # 文字列のバインド
        self.file_path = StringVar()
        self.folder_path = StringVar()
        self.file_list_var = StringVar(value=())
        self.file_name_var = StringVar()
        # PDFファイル選択
        self.file_label = ttk.Label(self.main_frm, text=u'PDFファイル')
        self.file_box = ttk.Entry(self.main_frm, textvariable=self.file_path, width=50)
        self.file_btn = ttk.Button(self.main_frm, text=u'参照', command=self.file_btn_click)
        # 保存先選択
        self.folder_label = ttk.Label(self.main_frm, text=u'保存先フォルダ')
        self.folder_box = ttk.Entry(self.main_frm, textvariable=self.folder_path, width=50)
        self.folder_btn = ttk.Button(self.main_frm, text=u'参照', command=self.folder_btn_click)
        # 保存先のファイルを表示するリスト
        # self.file_list = tkinter.Listbox(self.main_frm,listvariable=self.file_list_var,height=8)
        self.file_list = ttk.Treeview(self.main_frm)
        self.file_list.bind('<Double-1>', lambda event: self.file_list_double_click())
        # スクロールバー
        self.file_list_scr = ttk.Scrollbar( self.main_frm,orient=tkinter.VERTICAL,command=self.file_list.yview)
        self.file_list['yscrollcommand'] = self.file_list_scr.set
        # ファイルを削除するボタン
        self.file_select_del_btn = ttk.Button(self.main_frm,text=u'選択削除',command=self.file_select_del_btn_click)
        # ファイルを全て削除するボタン
        self.file_all_del_btn = ttk.Button(self.main_frm, text=u'全削除', command=self.file_all_del_btn_click)
        # コントラスト
        self.cont_box_var = StringVar()
        self.cont_label = ttk.Label(self.main_frm, text=u'コントラスト')
        self.cont_box = ttk.Entry(self.main_frm, textvariable=self.cont_box_var ,width=5)
        self.cont_box.bind('<KeyPress-Return>',lambda event:self.cont_box_callback())
        self.cont_box.insert(tkinter.END, '1.5')
        # コントラストのスケール
        self.cont_scale_var = tkinter.DoubleVar()
        self.cont_scale_var.set(1.5)
        self.cont_scale = ttk.Scale(self.main_frm,from_=0,to=5,variable=self.cont_scale_var,command=self.cont_scale_callback )
        # ファイル名
        self.file_name_label = ttk.Label(self.main_frm,text=u'ファイル名')
        self.file_name_box = ttk.Entry(self.main_frm,width=10,textvariable=self.file_name_var)
        # プログレスバー
        self.prog_bar = ttk.Progressbar(self.main_frm, orient=tkinter.HORIZONTAL, length=200, mode='indeterminate')
        self.prog_bar.configure(maximum=100, value=0)
        # PDFをPNGに変換するボタン
        self.convert_btn = ttk.Button(self.main_frm, text=u'変換', command=self.convert_btn_pdf)

        # メインフレームのレイアウト
        self.main_frm.place(x=5,y=5,width=630,height=470)
        # PDFファイル選択のレイアウト
        self.file_label.place(x=5,y=5,width=130,height=30)
        self.file_box.place(x=155,y=5,width=300,height=30)
        self.file_btn.place(x=475,y=5,width=80,height=30)
        # 保存先選択のレイアウト
        self.folder_label.place(x=5,y=45,width=130,height=30)
        self.folder_box.place(x=155,y=45,width=300,height=30)
        self.folder_btn.place(x=475,y=45,width=80,height=30)
        # 保存先ファイルのリストのレイアウト
        self.file_list.place(x=5,y=85,width=250,height=200)
        # スクロールバーのレイアウト
        self.file_list_scr.place(x=255,y=85,width=10,height=200)
        # ファイルを削除するボタンのレイアウト
        self.file_select_del_btn.place(x=275,y=85,width=100,height=30)
        # ファイルを全て削除するボタンのレイアウト
        self.file_all_del_btn.place(x=275,y=125,width=100,height=30)
        # コントラストのレイアウト
        self.cont_label.place(x=5,y=295,width=130,height=30)
        self.cont_box.place(x=145,y=295,width=60,height=30)
        # コントラストスケールのレイアウト
        self.cont_scale.place(x=215,y=295,width=200,height=30)
        # ファイル名のレイアウト
        self.file_name_label.place(x=5,y=335,width=130,height=30)
        self.file_name_box.place(x=145,y=335,width=120,height=30)
        # PDFをPNGに変換するボタンのレイアウト
        self.convert_btn.place(x=275,y=335,width=80,height=30)
        # プログレスバーのレイアウト
        self.prog_bar.place(x=5,y=375,width=260,height=30)

    def mainloop(self):
        self.main_win.mainloop()

    # pdfファイルを選択するボタンのイベント
    def file_btn_click(self):
        file_type = [(u'PDFファイル', '*.pdf')]
        path = filedialog.askopenfilename(filetypes=file_type, initialdir=self.cur_dir)
        file=path.split("/")

        p = file[0] + "\\"
        i = 1
        for i in range( len(file)-1):
            p = os.path.join(p,file[i])
        print( "cur_dir:" + p )
        self.file_path.set(path)
        self.file_name_var.set(file[len(file)-1].split(".")[0] )
        self.cur_dir = p

    # pngファイルを保存するフォルダを選択するボタンのイベント
    def folder_btn_click(self):
        path = filedialog.askdirectory(initialdir=self.cur_dir)
        self.file_list_show(path)
        self.folder_path.set(path)

    # パラメータ（パス）内のファイルをfile_listに表示する
    def file_list_show(self,path):
        # リスト内をクリアする
        self.file_list.delete( *self.file_list.get_children() )
        for file in os.listdir(path):
            sub_path = os.path.join( path , file )
            if( os.path.isdir(sub_path) ):
                #フォルダの場合
                folder = self.file_list.insert("","end",text=file)
                for file2 in os.listdir(sub_path):
                    if '.png' in file2:
                        self.file_list.insert(folder,"end",text=file2)

    # コンバートボタンのイベント
    def convert_btn_pdf(self):
        os.makedirs(os.path.join(self.folder_box.get(),self.file_name_box.get()),exist_ok=True)
        self.prog_bar.start(interval=10)
        th = threading.Thread(target=self.pdftopng, args=(self.file_box.get(), os.path.join(self.folder_box.get(),self.file_name_box.get()), self.cont_box.get()))
        th.start()

    # 選択したリスト番号を返却する
    def file_list_selection(self):
        for i in self.file_list.curselection():
            return i

    # 選択したファイルを削除するボタンのイベント
    def file_select_del_btn_click(self):
        file_name = self.get_treeview_file_path(self.file_list)
        #print(os.path.join(self.folder_box.get(),file_name ) )
        file_path = os.path.join(self.folder_box.get(),file_name)
        if '.png' in file_name:
            os.remove( file_path )
            self.file_list.delete(self.file_list.focus())
            #self.file_list.delete( tkinter.ANCHOR ,tkinter.ACTIVE )
        elif os.path.isdir(os.path.join(self.folder_box.get(),file_name)):
            #shutil.rmtree(os.path.join(self.folder_box.get(),file_name))
            try:
                print( file_path )
                os.rmdir( file_path )
                self.file_list.delete(self.file_list.focus())
            except OSError:
                messagebox.showerror("削除できません","フォルダが空ではないため削除できません。")

    # ファイルを全て削除する
    def file_all_del_btn_click(self):
        for child in self.file_list.get_children(self.file_list.focus()):
            file = self.file_list.item( child )['text']
            if '.png' in file:
                parent = self.file_list.item( self.file_list.focus() )['text']
                file_path = os.path.join( self.folder_box.get(),parent,file )
                os.remove( file_path )
                self.file_list.delete( child )

    #リストボックスをダブルクリックすると画像を表示する
    def file_list_double_click(self):
        focus_file_path = self.get_treeview_file_path(self.file_list)
        file_path = os.path.join(self.folder_box.get(), focus_file_path )
        if '.png' in file_path:
            im = Image.open(file_path)
            im.show()

    # フォーカスされたツリービューのパスを返す
    def get_treeview_file_path(self,tv):
        file_name = tv.item(tv.focus())['text']
        parent_name = tv.item(tv.parent( tv.focus()))['text']
        return os.path.join( parent_name,file_name)

    # コントラストスライダーのコールバック
    def cont_scale_callback(self,value):
        self.cont_box_var.set( round(float(value)*10) / 10 )

    # コントラストボックスのコールバック
    def cont_box_callback(self):
        self.cont_scale_var.set( float( self.cont_box.get()) )

    # pdfをpngに変換するメソッド
    def pdftopng(self , pdf_file_path , image_file_path , contrast=1):
        poppler_dir = os.path.join(self.home_dir , 'poppler-0.51','bin','pdftocairo')
        src_path = pdf_file_path
        dst_path = os.path.join( image_file_path , self.file_name_box.get() )
        print( poppler_dir )
        print( src_path )
        print( dst_path )
        cmd = '%s -png %s %s'%(poppler_dir, src_path ,dst_path)

        returncode = subprocess.Popen( cmd , shell=True)
        returncode.wait()

        print( returncode )

        for file in os.listdir(image_file_path):
            if '.png' in file:
                im1 = Image.open( os.path.join( image_file_path , file ) )
                con = ImageEnhance.Contrast( im1 )
                im2 = con.enhance( float(contrast) )

                im1.close()

                os.remove( os.path.join( image_file_path , file ))

                file_name = file[0:len(file)-4] + "-conv.png"
                im2.save( os.path.join( image_file_path,file_name ))
                im2.close()

        self.prog_bar.stop()
        messagebox.showinfo("終了", "変換が終了しました。")
        # ファイル名を表示する
        self.file_list_show(self.folder_box.get() )

if __name__ == "__main__":
    frame = MyFram()
    frame.mainloop()