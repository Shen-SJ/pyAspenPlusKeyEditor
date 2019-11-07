# -*- mode: python ; coding: utf-8 -*-
"""
Description:
    將我自己做的 AspenPlus 機碼給新增與修改上去，讓apw檔案有一下拉是選單可以選擇開啟的aspen版本

Editor: Shen_SJ
Date:   2019.11.04
"""
# %%
import regobj as r
import pickle, re
import sys, getopt
import logging

# 設置我的 logger
handler_s = logging.StreamHandler()
handler_f = logging.FileHandler(filename="RunningMessage.log",
                                mode='a',
                                encoding="utf-8",
                                delay=True
                                )

handler_s.setLevel(logging.ERROR)
handler_f.setLevel(logging.ERROR)

handler_f.setFormatter(logging.Formatter("%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s"))

logging.basicConfig(level=logging.DEBUG,
                    # format="%(asctime)s - %(pathname)s[line:%(lineno)d] - %(levelname)s: %(message)s",
                    handlers=[handler_s, handler_f]
                    )


class AspenPlusKeyBuilder:
    """
    用法: AspenPlusKeyEdit [選項]
    本程式主要使用對象為電腦中有數個 AspenPlus 版本，透過該程式可以修改 AspenPlus 相關的 Registry Key，使得用戶在 apw 檔案點擊右鍵會有選單來選擇 AspenPlus 開啟版本。
    由於本程式會修改電腦 Registry Key，因此必須以系統管理員權限執行。

    本程式所需之選項在以下：
        -m, --modified-key      執行修改 AspenPlus 相關 Registry key
        -r, --restored-key      將修改過的 AspenPlus 相關 Registry key 復原回原來樣貌
        -h, --help              印出本幫助文件
    """
    __CLSID = None
    __default_icon = None
    __default_command = None
    __version_list = None
    __version_label_dict = None
    __version_exe_dict = None

    def __init__(self):
        # 確認你的電腦有沒有安裝 AspenPlus
        self.check_aspen_exist()

        # 看看你的電腦裝了什麼版本的 AspenPlus
        self.aspen_version_list()
        self.aspen_version_label_dict()
        self.aspen_version_exe_dit()

        # 將以下三種屬性存起來，等等複製機碼會用到
        self.__CLSID = r.HKEY_CLASSES_ROOT(r"Apwn.Document\CLSID")[''].data
        self.__default_icon = r.HKEY_CLASSES_ROOT(r"Apwn.Document\DefaultIcon")[''].data
        self.__default_command = r.HKEY_CLASSES_ROOT(r"Apwn.Document\shell\Open\command")[''].data

    def check_aspen_exist(self):
        """
        Check whether the AspenPlus install or not.

        :return: True for installed AspenPlus
        """
        try:
            r.HKEY_CLASSES_ROOT("Apwn.Document")
            return True
        except AttributeError as e:
            if e == "subkey 'Apwn.Document' does not exist":
                raise ValueError("There is no AspenPlus in the Computer !!")
            else:
                raise Exception('Unexpect Error, Please Connect to the developer !!')

    def aspen_version_list(self):
        """
        Find AspenPlus versions you have to a list.

        :return: None
        """
        # 把 HKEY_CLASSES_ROOT 下的 Apwn.Document.XX.X 找出來，確認電腦安裝的 Aspen 版本
        version_list = [k.name for k in r.HKEY_CLASSES_ROOT if re.match("Apwn.Document.\d+", k.name)]
        version_list.sort()  # 我喜歡按照版本順序排列~~
        self.__version_list = version_list

    def aspen_version_label_dict(self):
        """
        Correspond the version label with version name.

        :return: dict.
        """
        # 先以 version name 找到對應的 version label
        l_list = []
        for version_name in self.__version_list:
            h_key = r.HKEY_CLASSES_ROOT(fr"{version_name}\shell")
            for item in h_key.subkeys():
                if re.match("Open with Aspen Plus V\d+.", item.name): l_list.append(item.name)

        # 然後再將 name 與 label 變成字典
        self.__version_label_dict = dict(zip(self.__version_list, l_list))

    def aspen_version_exe_dit(self):
        """
        Correspond the executive path with version name.

        :return: dict.
        """
        # 先以 version name 找到對應的 path
        e_list = []
        for version_name in self.__version_list:
            h_key = r.HKEY_CLASSES_ROOT(f"{version_name}\shell\{self.__version_label_dict[version_name]}\command")
            e_list.append(h_key[''].data)

        # 再將 name 與 exe path 變成字典
        self.__version_exe_dict = dict(zip(self.__version_list, e_list))

    def save_curver(self):
        """
        Save the default value of HKEY_CLASSES_ROOT\Apwn.Document\CurVer.

        :return: None
        """
        curver = r.HKEY_CLASSES_ROOT("Apwn.Document").CurVer[''].data

        with open('data', 'wb') as f:
            pickle.dump(curver, f)

    def create_aspen_key(self):
        """
        Create the Apwn.Document.UserDefine Key which contains bunch of keys to build the secondary list.

        :return: None
        """
        # 新增一個自創Aspen機碼
        r.HKEY_CLASSES_ROOT.set_subkey('Apwn.Document.UserDefine', r.Key)

        # 在這個新增的機碼新增 CLSID 機碼與設定其值
        r.HKEY_CLASSES_ROOT('Apwn.Document.UserDefine').set_subkey('CLSID', {'': self.__CLSID})

        # 在這個新增的機碼新增 DefaultIcon 機碼與其設定值
        r.HKEY_CLASSES_ROOT('Apwn.Document.UserDefine').set_subkey('DefaultIcon', {
            '': self.__default_icon})  # 在這個新增的機碼新增 DefaultIcon 機碼與其設定值

        # 在這個新增的機碼新增 shell\Open\Comman 機碼與其設定值
        r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine').set_subkey('shell', {'': 'Open'})
        r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine\shell').set_subkey('Open', {'': '&Open'})
        r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine\shell\Open').set_subkey('command', {'': self.__default_command})

        # 在這個新增的機碼新增 shell\openaspen\shell\[open1, open2...] 機碼與其設定值
        r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine\shell').set_subkey('openaspen',
                                                                          {'': '',
                                                                           "MUIVerb": "Open with Aspen Plus",
                                                                           "subcommands": ""}
                                                                          )
        r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine\shell\openaspen').set_subkey('shell', r.Key)
        h_key = r.HKEY_CLASSES_ROOT(r'Apwn.Document.UserDefine\shell\openaspen\shell')

        # 把次級選單的名子、執行檔路徑新增進去！
        for num_version, version in enumerate(self.__version_list):
            h_key.set_subkey(f'open{num_version}', {'': self.__version_label_dict[version]})
            h_key(f'open{num_version}').set_subkey('command', {'': self.__version_exe_dict[version]})

    def delete_aspen_key(self):
        """
        Delete the Apwn.Document.UserDefine Key.

        :return: None
        """
        r.HKEY_CLASSES_ROOT.del_subkey('Apwn.Document.UserDefine')

    def modified_aspen_key(self):
        """
        Modified the AspenPlus key to build the secondary list.

        :return: None
        """
        self.create_aspen_key()
        r.HKEY_CLASSES_ROOT(".apw")[''] = 'Apwn.Document.UserDefine'

    def restored_aspen_key(self):
        """
        Restore the AspenPlus key to Default.

        :return: None
        """
        self.delete_aspen_key()     # 如果發現沒有 Apwn.Document.UserDefine 機碼，會直接報錯誤
        r.HKEY_CLASSES_ROOT(".apw")[''] = 'Apwn.Document'

    def run_script(self):
        """
        將本程式能以 CMD 執行的寫法，還能輸入執行相關參數

        :return: None
        """
        try:
            # 一定要以下指定的這幾個參數才能運作，不然就會 GetoptError
            opts, args = getopt.getopt(sys.argv[1:],
                                       shortopts="mrh",
                                       longopts=["modified-key", "restored-key", "help"]
                                       )
            # Debug 用，印出接收的變數是甚麼
            logging.debug(f"opts: {opts}")
            logging.debug(f"args: {args}")

        except getopt.GetoptError:
            # 如果參數格式給錯了，印一下怎麼給
            print(self.__doc__)
            sys.exit()

        if not opts:    # 如果我的程式沒有給定參數或選項，就列印出使用方式
            print(self.__doc__)
            sys.exit()

        for opt, arg in opts:
            if opt in ("-h", "--help"):  # 進入 help 的功能
                # 介紹該程式的參數要怎麼下
                print('以下是該程式的幫助文件......')
                print(self.__doc__)
            elif opt in ('-m', '--modified-key'):  # 進入 modified-key 功能
                # 修改 AspenPlus 機碼
                self.modified_aspen_key()
                print('Modify Success !!!')
            elif opt in ('-r', '--restored-key'):  # 進入 restored-key 功能
                # 將修改過的 AspenPlus 機碼改回原貌
                self.restored_aspen_key()
                print('Restored Success !!!')
            for item in args:
                logging.debug(item)


# %%
if __name__ == "__main__":
    # 把 AspenPlusKeyBuilder 類建立起來，而且也會檢查你的電腦有沒有安裝 AspenPlus
    areg = AspenPlusKeyBuilder()

    logging.debug(areg.check_aspen_exist())
    logging.debug('印出debud了')

    try:
        areg.run_script()
    except PermissionError as mes:
        logging.error(mes)
        print("你必須以系統管理員權限執行該程式 !!!")
        print(areg.__doc__)

