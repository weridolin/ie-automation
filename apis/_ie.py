
import comtypes.client,time
import ctypes,comtypes,comtypes.hresult
import re
import uiautomation as uia

def _create_comtype_client(proid:str):
    try:
        client = comtypes.client.GetActiveObject(proid)
    except OSError as getActiveError:
        client = comtypes.client.CreateObject(proid) 
    return client    


class IWebBrowser2(object):

    def __init__(self,ie_object=None,handle=None,manager=None) -> None:
        if not isinstance(ie_object, comtypes.gen.SHDocVw.IWebBrowser2):
            raise TypeError(u'ie_object must be a type of comtypes.gen.SHDocVw.IWebBrowser2')
        self.ie_object = ie_object            
        if not isinstance(handle, int):
            self.handle = self.get_handle()
        else:
            self.handle = handle
        
        self.manager = manager
        manager.register(self)

    @classmethod
    def create(cls,handle=None,manager=None)->"IWebBrowser2":
        ie_object:comtypes.gen.SHDocVw.IWebBrowser2 = comtypes.client.CreateObject("InternetExplorer.Application") 
        return cls(ie_object=ie_object,handle=handle,manager=manager)

    @classmethod
    def from_opening_windows(cls,search_codnition)->"IWebBrowser2":
        if isinstance(search_codnition,int):
            try:
                return single_ie_manager.already_exist_ie_browser_list[search_codnition]
            except IndexError:
                raise IndexError(f"search_codnition over index:{len(single_ie_manager.already_exist_ie_browser_list)}")
        if isinstance(search_codnition,str) and search_codnition.startswith("http"):
            # get windows by url
            client = None
            for ie in single_ie_manager.ie_browser_list:
                if re.match(search_codnition,ie.url):
                    client =  ie
            if not client:
                raise AttributeError(f"can not find ie browser match url:{search_codnition}")
            return client

    @property
    def full_screen(self)->bool:
        """ if is full screen """
        is_full = comtypes.automation.VARIANT_BOOL()
        res = self.ie_object._IWebBrowserApp__com__get_FullScreen(is_full)
        if res==comtypes.hresult.S_OK:
            return is_full.value
        return False       

    @full_screen.setter
    def full_screen(self,is_full):
        """ set if is full screen """
        self.ie_object._IWebBrowserApp__com__set_FullScreen(
            comtypes.automation.VARIANT(is_full),
        )

    def open(self,
            url="http://www.baidu.com/",
            is_visible=True,
            headers = None,
            is_max=False,
            is_new_tab=False):
        """open a new page 
            @https://docs.microsoft.com/en-us/previous-versions/aa752133(v=vs.85)
        
        """
        self.ie_object.visible = is_visible
        if not headers:
            headers =  comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY)
        else:
            assert isinstance(headers,dict),"header can only be map"
            headers = comtypes.automation.VARIANT(headers)
        if is_max:
            uia.ShowWindow(handle=self.handle,cmdShow=uia.SW.Maximize)
        if not is_new_tab:
            self.ie_object._IWebBrowser2__com_Navigate2(
                url,
                comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY),
                comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY),
                comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY),
                headers)
        else:
            self.ie_object._IWebBrowser2__com_Navigate2(
                url,
                # @https://docs.microsoft.com/en-us/previous-versions/dd565688(v=vs.85)
                comtypes.automation.VARIANT(2048),
                comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY),
                comtypes.automation.VARIANT(comtypes.automation.VT_EMPTY),
                headers)            
    
    def close(self)->bool:
        if self.ie_object._IWebBrowserApp__com_Quit()==comtypes.hresult.S_OK:
            return True
        return False
    
    def go_forword(self)->bool:
        if self.ie_object._IWebBrowser__com_GoForward()==comtypes.hresult.S_OK:
            return True
        return False

    def go_back(self)->bool:
        if self.ie_object._IWebBrowser__com_GoBack()==comtypes.hresult.S_OK:
            return True
        return False       

    def get_handle(self)->int:
        return self.ie_object.HWND

    def go_home(self)->bool:
        if self.ie_object._IWebBrowser__com_GoHome()==comtypes.hresult.S_OK:
            return True
        return False  

    @property
    def width(self)->int:
        return self.ie_object.Width

    @width.setter
    def width(self,width)->bool:
        self.ie_object.Width = width

    @property
    def height(self)->int:
        return self.ie_object.Height

    @height.setter
    def height(self,height)->bool:
        self.ie_object.Height = height

    def refresh2(self,level=3)->bool:
        """
            level:@https://docs.microsoft.com/en-us/previous-versions//aa768363(v=vs.85)
        """
        level = comtypes.automation.VARIANT(level)
        level.vt = comtypes.automation.VT_I4
        if self.ie_object._IWebBrowser__com_Refresh2(level)==comtypes.hresult.S_OK:
            return True
        return False  

    @property
    def url(self)->str:
        url = comtypes.automation.BSTR()
        res = self.ie_object._IWebBrowser__com__get_LocationURL(ctypes.byref(url))
        if res == comtypes.hresult.S_OK:
            return url.value
        else:
            return ""
    
    def set_size(self,width,height)->bool:
        self.width,self.height = width,height

    def open_page(self,url,is_new_tab)->bool:
        return self.open(url=url,is_new_tab=is_new_tab)

    
    def get_ready_state(self)->int:
        """ 
            @https://docs.microsoft.com/en-us/previous-versions//aa752141(v=vs.85)?redirectedfrom=MSDN
            @https://docs.microsoft.com/en-us/previous-versions//aa768362(v=vs.85)
            typedef enum tagREADYSTATE {
                READYSTATE_UNINITIALIZED = 0,
                READYSTATE_LOADING = 1,
                READYSTATE_LOADED = 2,
                READYSTATE_INTERACTIVE = 3,
                READYSTATE_COMPLETE = 4
            } READYSTATE;
        """
        state = comtypes.automation.LONG()
        state.vt = comtypes.automation.VT_I4
        res = self.ie_object._IWebBrowser2__com__get_ReadyState(state)
        if res==comtypes.hresult.S_OK:
            return state.value
        return  -1 
    
    def stop(self)->bool:
        """ 
            stop downloading  
            @https://docs.microsoft.com/en-us/previous-versions//aa768272(v=vs.85)?redirectedfrom=MSDN
            # todo check if it can stop page loading?
        """
        if self.ie_object._IWebBrowser__com_Stop()==comtypes.hresult.S_OK:
            return True
        return False      

    @property
    def document(self):
        ## TODO @https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752541(v=vs.85)?redirectedfrom=MSDN
        # doc = comtypes.automation.IDispatch()
        # doc.vt =comtypes.automation.VT_DISPATCH
        # res = self.ie_object._IWebBrowser__com__get_Document(comtypes.POINTER(doc))
        # if res==comtypes.hresult.S_OK:
        #     return doc
        # return None      



class IeItem(object):
    def __init__(self,ie_object:IWebBrowser2) -> None:
        self.ie_object = ie_object
        self.handle = ie_object.handle
        self.url = ie_object.url

class IWebBrowerManager(object):

    ie_browser_list = []
    already_exist_ie_browser_list = []

    __instance = None
    def __new__(cls, *args, **kwargs):
        if cls.__instance is None:
            cls.__instance = super().__new__(cls)
        return cls.__instance

    def __init__(self) -> None:
        ## shellwindows: file:/// / http://
        self.shellWindows = _create_comtype_client("{9BA05972-F6A8-11CF-A442-00A0C90A8F39}")
        for i in range(self.shellWindows.Count):
            if self.shellWindows[i]:
                print(self.shellWindows[i],type(self.shellWindows[i]),self.shellWindows[i].LocationURL)      
                IWebBrowerManager.already_exist_ie_browser_list.append(self.shellWindows[i])

    @staticmethod
    def register(ie_browser:IWebBrowser2):
        if ie_browser not in [item.handle for item in  IWebBrowerManager.ie_browser_list]:
            IWebBrowerManager.ie_browser_list.append(IeItem(ie_browser))


single_ie_manager = IWebBrowerManager()


if  __name__ == "__main__":
    ie = IWebBrowser2.create(manager=single_ie_manager)
    ie.open(is_max=True)
    print(ie.get_ready_state())
    # ie.open_page(url="www.baidu.com",is_new_tab=False)
    time.sleep(5)
    print(ie.get_ready_state())
    # ie.refresh2()
    print(ie.document)
    # ie.go_back()
    # time.sleep(5)
    # ie.go_forword()
    # time.sleep(5)
    # ie.go_home()
    # time.sleep(5)
    ie.open_page(url="https://www.baidu.com",is_new_tab=True)
    ie.set_size(1200,900)
    time.sleep(5)
    # ie.close()
    # time.sleep(10)