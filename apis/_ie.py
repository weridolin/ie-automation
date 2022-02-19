from datetime import date, datetime,timedelta
from shutil import which
from tkinter.font import BOLD
import comtypes.client,time
import ctypes,comtypes,comtypes.hresult
import re
import uiautomation as uia


class ReadyState(object):
    """ 
        @https://docs.microsoft.com/en-us/previous-versions//aa768362(v=vs.85)
    """
    READYSTATE_UNINITIALIZED = 0
    READYSTATE_LOADING = 1
    READYSTATE_LOADED = 2
    READYSTATE_INTERACTIVE = 3
    READYSTATE_COMPLETE = 4

    @classmethod
    def contains(cls,state):
        return state in [getattr(cls,name) for name in ReadyState.__dict__ if not name.startswith("__")]
    

def _create_comtype_client(proid:str):
    try:
        client = comtypes.client.GetActiveObject(proid)
    except OSError as getActiveError:
        client = comtypes.client.CreateObject(proid) 
    return client    


class IWebBrowser2(object):

    def __init__(self,ie_object=None,handle=None) -> None:
        if not isinstance(ie_object, comtypes.gen.SHDocVw.IWebBrowser2):
            raise TypeError(u'ie_object must be a type of comtypes.gen.SHDocVw.IWebBrowser2')
        self.ie_object = ie_object            
        if not isinstance(handle, int):
            self.handle = self.get_handle()
        else:
            self.handle = handle
        
        self.manager = IWebBrowerManager.register(self)
        self.iHtmlDocumentInterface = None

    @classmethod
    def create(cls)->"IWebBrowser2":
        ie_object:comtypes.gen.SHDocVw.IWebBrowser2 = comtypes.client.CreateObject("InternetExplorer.Application") 
        return cls(ie_object=ie_object)

    @classmethod
    def from_opening_windows(cls,search_codnition)->"IWebBrowser2":
        if not IWebBrowerManager.__instance:
            IWebBrowerManager()
        if isinstance(search_codnition,int):
            try:
                return IWebBrowerManager.already_exist_ie_browser_list[search_codnition]
            except IndexError:
                raise IndexError(f"search_codnition over index:{len(IWebBrowerManager.already_exist_ie_browser_list)}")
        if isinstance(search_codnition,str) and search_codnition.startswith("http"):
            # get windows by url
            client = None
            for ie in IWebBrowerManager.ie_browser_list:
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
        self._wait()          
        self.iHtmlDocumentInterface = IHtmlDocumentInterface.from_ie_browser(self) # todo register interface through event?

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
    def iHtmlDocument(self):
        ## TODO @https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752541(v=vs.85)?redirectedfrom=MSDN
        # html = comtypes.automation.POINTER(comtypes.automation.IDispatch)()
        ...
        # if res==comtypes.hresult.S_OK:
        #     return html.value
        # return None    
    
    @iHtmlDocument.getter
    def iHtmlDocument(self):
        try:
            return self.ie_object.Document
        except comtypes.COMError:
            print("please open a url first!")
            raise
    
    def _wait(self,timeout = 120):
        """ wait until the page load finish """
        start = datetime.now()
        deadline  = start + timedelta(seconds=timeout)
        while datetime.now() < deadline:
            if self.get_ready_state() == ReadyState.READYSTATE_COMPLETE:
                return True
        raise TimeoutError(f"loading page timeout:{timeout}") # todo close or stop loading page?


    # def query_selector(self,select_str:str,many=False):
    #     if not many:
    #         return self.iHtmlDocument.querySelector(select_str)
    #     return self.iHtmlDocument.querySelectorAll(select_str)

    # def get_element_by_id(self,id:str):
    #     return self.iHtmlDocument.getElementById(id)

     
    # def get_elements_by_name(self,name:str):
    #     return self.iHtmlDocument.getElementsByName(name)

    # def getElementsByTagName(self,tag_name:str):
    #     return self.iHtmlDocument.getElementsByTagName(tag_name)

class IHtmlDocumentInterface(object):

    def __init__(self,iHtmlDocument=None) -> None:
        self.__iHtmlDocument=iHtmlDocument

    @classmethod
    def from_ie_browser(cls,ie_browser)->"IHtmlDocumentInterface":
        return cls(iHtmlDocument=ie_browser.iHtmlDocument)

    def query_selector(self,select_str:str,many=False):
        if not many:
            return self.__iHtmlDocument.querySelector(select_str)
        return self.__iHtmlDocument.querySelectorAll(select_str)

    def get_element_by_id(self,id:str):
        return self.__iHtmlDocument.getElementById(id)

    def get_elements_by_name(self,name:str):
        return self.__iHtmlDocument.getElementsByName(name)

    def getElementsByTagName(self,tag_name:str):
        return self.__iHtmlDocument.getElementsByTagName(tag_name)

    def __getattribute__(self, __name: str):
        print(__name)
        return super().__getattribute__(__name)


class IHTMLElement(object):
    
    def __init__(self,_IHTMLElement) -> None:
        self.___IHTMLElement = _IHTMLElement # todo support list

    def get_text(self,include_tag = False):
        """
        @https://docs.microsoft.com/en-us/previous-versions//hh869989(v=vs.85)?redirectedfrom=MSDN
        """

        text = comtypes.automation.BSTR()
        if not include_tag: 
            res = self.___IHTMLElement._IHTMLElement__com__get_innerText(ctypes.byref(text))
        else:
            res = self.___IHTMLElement._IHTMLElement__com__get_outerHTML(ctypes.byref(text))
        if res == comtypes.hresult.S_OK:
            return text.value
        else:
            return ""

    def set_text(self,value) -> bool:
        """
        @https://docs.microsoft.com/en-us/previous-versions//hh869989(v=vs.85)?redirectedfrom=MSDN 
        """ 
        text = comtypes.automation.BSTR()
        text.value = value
        res = self.___IHTMLElement._IHTMLElement__com__set_innerText(text)
        if res == comtypes.hresult.S_OK:
            return True
        else:
            return False

    def click(self):
        return self.___IHTMLElement.click()

    
    def get_attr(self,attr_name,flag = 0):
        attr = comtypes.automation.BSTR()
        attr.value  = attr_name
        flag = comtypes.automation.LONG()
        flag.value = flag
        res = self.___IHTMLElement._IHTMLElement__com_getAttribute(attr,flag)
        return res

    
    def set_attr(self,attr_name,value,flag =1)->bool:
        """
        @https://docs.microsoft.com/en-us/previous-versions/windows/internet-explorer/ie-developer/platform-apis/aa752330(v=vs.85)
        flag = 1: case sensitive  
        flag = 0: no case sensitive 
        """
        attr = comtypes.automation.BSTR()
        attr.value  = attr_name
        value = comtypes.automation.VARIANT(value)
        flag = comtypes.automation.LONG()
        flag.value = flag        
        res = self.___IHTMLElement._IHTMLElement__com_setAttribute(attr,value,flag)
        if res == comtypes.hresult.S_OK:
            return True
        else:
            return False        



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
        if not IWebBrowerManager.__instance: 
            IWebBrowerManager()
        if ie_browser not in [item.handle for item in  IWebBrowerManager.ie_browser_list]:
            IWebBrowerManager.ie_browser_list.append(IeItem(ie_browser))



if  __name__ == "__main__":

    ie = IWebBrowser2.create()
    ie.open(is_max=True)
    # ie.open_page(url="https://www.baidu.com",is_new_tab=True)
    print(ie.get_ready_state())
    # ie.open_page(url="www.baidu.com",is_new_tab=False)
    time.sleep(5)
    # ie.refresh2()
    # print(ie.iHtmlDocument,type(ie.iHtmlDocument),dir(ie.iHtmlDocument))
    el = ie.iHtmlDocumentInterface.get_element_by_id("su")
    # ie.go_back()
    # time.sleep(5)
    # ie.go_forword()
    # time.sleep(5)
    # ie.go_home()
    # time.sleep(5)
    ie.set_size(1200,900)
    time.sleep(5)
    # ie.close()
    time.sleep(10)

