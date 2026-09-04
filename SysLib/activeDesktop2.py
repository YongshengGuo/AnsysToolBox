#--- coding=utf-8
#--- @author: yongsheng.guo@ansys.com
#--- @Time: 20240215



from asyncio import log
import os,sys
import re
import time
isIronpython = "IronPython" in sys.version

appPath = os.path.realpath(__file__)
appDir = os.path.split(appPath)[0] 
sys.path.append(appDir)
sys.path.append(os.path.join(appDir,r"..\site-packages"))

# isIronpython = "IronPython" in sys.version
# isPython = not isIronpython
# is_linux = "posix" in os.name
# global_clr = None
# def initClr():
#     """
#     Initialize CLR. 
#     Note: In multiprocessing spawn mode, this might fail in child processes if called globally.
#     It is recommended to call this only in the main process or handle exceptions in child processes.
#     """
#     global global_clr
#     if global_clr:
#         return global_clr

#     # 设置环境变量以抑制 CLR 重复初始化警告（跨平台支持）
#     # 这应该在首次尝试初始化 CLR 时设置，避免子进程中的重复初始化警告
#     if 'ANSOFT_SUPPRESS_CLR_WARNING' not in os.environ:
#         os.environ['ANSOFT_SUPPRESS_CLR_WARNING'] = '1'

#     if isIronpython:
#         import clr as _clr
#         global_clr = _clr
#         return _clr
    
#     elif is_linux:
#         try:
#             from ansys.aedt.core.generic.clr_module import _clr
#         except Exception as e:
#             # Linux 平台也可能遇到 CLR 初始化问题
#             if os.environ.get('ANSOFT_SUPPRESS_CLR_WARNING') == '1':
#                 print("CLR initialization failed in child process on Linux (suppressed). Error: %s" % str(e))
#                 return None
#             else:
#                 print("CLR initialization failed on Linux: %s" % str(e))
#                 return None
        
#         global_clr = _clr
#         return _clr
    
#     else:
#         # Windows
#         try:
            
#             try:
#                 import clr as _clr
#             except ImportError:
#                 print("Note: for Python environment pythonnet must install, use command: pip install pythonnet")

#             global_clr = _clr
#             return _clr
#         except RuntimeError as e:
#             # 如果在子进程中遇到初始化错误，记录警告并返回 None 或重新抛出
#             # 这通常发生在 spawn 模式下子进程重复导入时
#             if "Failed to initialize Python.Runtime.dll" in str(e):
#                 if os.environ.get('ANSOFT_SUPPRESS_CLR_WARNING') == '1':
#                     print("CLR initialization failed in child process on Windows (suppressed).")
#                 else:
#                     print("CLR initialization failed in child process (likely duplicate init). Ignoring if not needed.")
#                 # 返回 None 或根据调用者逻辑处理
#                 return None 
#             raise

# try:
#     _clr = initClr()
#     from System import String
# except:
#     print("CLR initialization failed in child process (likely duplicate init). Ignoring if not needed.")

try:
    import clr
except:
    if not isIronpython:
        print("Note: for Python environment pythonnet must install, use command: pip install pythonnet")
    exit()

try:
    sys.path.append(r'C:\work\Study\Script\cols\VSRepos\simpleWinAPI\simpleWinAPI2\bin\Release')
    clr.AddReference("tidyWinAPI")
    import tidyWinAPI
    from tidyWinAPI import winAPI
except:
    print("import tidyWinAPI error.")
    
from System import IntPtr

'''
    BOOL CALLBACK EnumWindowsProc(
      _In_ HWND   hwnd,
      _In_ LPARAM lParam
    );
'''



def getAllInstances(progID):
    #ironpython environment
    if isIronpython:
        print("Running in Ironpython environment")
        # import System.Collections.ICollection
        apps = winAPI.GetRunningInstances(progID)
        return [apps[i] for i in range(apps.Count)]
    else:
        #python environment
        print("Running in Python environment")
        import pythoncom
        import win32com.client
        pythoncom.CoInitialize()
        context = pythoncom.CreateBindCtx(0)
        RunningObjectTable = pythoncom.GetRunningObjectTable()
        monikiers = RunningObjectTable.EnumRunning()
        apps = []
        print("Get Runing Aedt instance ....")
        for monikier in monikiers:
            name = monikier.GetDisplayName(context, monikier)
#             print(name)
            if progID in name:
                print(name)
                obj = RunningObjectTable.GetObject(monikier)
                app = win32com.client.Dispatch(obj.QueryInterface(pythoncom.IID_IDispatch))
                apps.append(app)
            else:
                continue
        
        return apps
    
    
def GetTopWindow(hWnd=IntPtr(0)):
    '''
    要检查其子窗口的父窗口的句柄。 如果此参数为 NULL，则该函数将返回 Z 顺序顶部窗口的句柄。
    '''
    return winAPI.GetTopWindow(hWnd)


def GetWidowText(hWnd):
    return winAPI.GetWidowText(hWnd) 

def FindWindowEx(hwndParent=IntPtr(0),hwndChildAfter=None,lpszClass=None,lpszWindow=None):   
    return winAPI.FindWindowEx(hwndParent,hwndChildAfter,lpszClass,lpszWindow)

def GetWindowThreadProcessId(hWnd):
    return winAPI.GetWindowThreadProcessId(hWnd)

def getoDesktopByProcessID(processID):
    apps = getAllInstances("ElectronicsDesktop")
    if not apps:
        print("No Aedt application running ...")
        return None
    
    for app in apps:
        if app.GetProcessID() == processID:
            return app
    
    print("not found Aedt with process_id:%s"%processID)
    return None

def getSiwaveoAppByTitle(title):
    apps = getAllInstances("SIwave.Application")
    if not apps:
        print("No SIwave application running ...")
        return None
    
    for app in apps:
        if app.GetProcessID() == processID:
            return app
    
    print("not found SIwave with process_id:%s"%processID)
    return None

def getActiveDesktop():
    
    apps = getAllInstances("ElectronicsDesktop")
    if not apps:
        print("No Aedt application running ...")
        return None
    
    try:
        current_window = GetTopWindow(IntPtr(0))
        while current_window:
            window_title = GetWidowText(current_window)
            
            if window_title and "Electronics Desktop" in window_title:
                process_id = GetWindowThreadProcessId(current_window)
                for app in apps:
                    if process_id == app.GetProcessID():
                        print("Get current actived AEDT: %s"%window_title)
                        print("Aedt install Dir:%s"%app.GetExeDir())
                        
                        # 将活动桌面实例保存到主模块
                        main_module = sys.modules['__main__']
                        main_module.oDesktop = app
                        return app
            
            current_window = FindWindowEx(hwndChildAfter=current_window)
            
        return None

    except Exception as e:
        print("Error in get_active_desktop:%s"%str(e))
        return None



def _releaseComObjects():
    '''
    Release all COM object references.
    This allows other applications to access AEDT without conflicts.
    The AEDT application itself remains open.
    '''
    module = sys.modules["__main__"]
    
    # Release oDesktop COM object
    released_count = 0
    try:
        if hasattr(module, "_oDesktop"):
            oDesktop = getattr(module, "_oDesktop")
            if oDesktop:
                try:
                    if isIronpython:
                        # For IronPython: use Marshal.ReleaseComObject
                        try:
                            from System.Runtime.InteropServices import Marshal
                            while Marshal.ReleaseComObject(oDesktop) > 0:
                                pass
                            print("Released _oDesktop COM object (IronPython)")
                            released_count += 1
                        except Exception as e:
                            print("Marshal.ReleaseComObject not available: %s" % str(e))
                    else:
                        # For CPython: just delete the reference
                        print("Released _oDesktop reference")
                        released_count += 1
                except Exception as e:
                    print("Error releasing _oDesktop: %s" % str(e))
        
        if hasattr(module, "oDesktop"):
            oDesktop = getattr(module, "oDesktop")
            if oDesktop and oDesktop != getattr(module, "_oDesktop", None):
                try:
                    if isIronpython:
                        try:
                            from System.Runtime.InteropServices import Marshal
                            while Marshal.ReleaseComObject(oDesktop) > 0:
                                pass
                            print("Released oDesktop COM object (IronPython)")
                            released_count += 1
                        except Exception as e:
                            print("Marshal.ReleaseComObject not available: %s" % str(e))
                    else:
                        print("Released oDesktop reference")
                        released_count += 1
                except Exception as e:
                    print("Error releasing oDesktop: %s" % str(e))
    
    except Exception as e:
        print("_releaseComObjects error: %s" % str(e))
    
    if released_count > 0:
        print("COM objects released: %d" % released_count)
    
    return released_count


def _delete_objects():
    '''
    Delete all object references from the main module.
    This should be called after _releaseComObjects() to clean up references.
    '''
    module = sys.modules["__main__"]
    
    # First release COM objects
    _releaseComObjects()
    
    # Then delete all references
    try:
        del module.COMUtil
    except (AttributeError, KeyError):
        pass

    try:
        del module.oDesktop
    except (AttributeError, KeyError):
        pass
    
    try:
        del module._oDesktop
    except (AttributeError, KeyError):
        pass
    
    try:
        del module.pyaedt_initialized
    except (AttributeError, KeyError):
        pass
    try:
        del module.oAnsoftApplication
    except (AttributeError, KeyError):
        pass
    try:
        del module.desktop
    except (AttributeError, KeyError):
        pass
    try:
        del module.sDesktopinstallDirectory
    except (AttributeError, KeyError):
        pass
    try:
        del module.isoutsideDesktop
    except (AttributeError, KeyError):
        pass
    try:
        del module.AEDTVersion
    except (AttributeError, KeyError):
        pass
    try:
        del sys.modules["glob"]
    except Exception:
        pass
    
    # Force garbage collection
    import gc
    gc.collect()


def releaseDesktop():
    '''
    Release COM resources and relinquish AEDT window control.
    
    This function releases all COM object references held by this Python process,
    allowing other applications or scripts to access AEDT without conflicts.
    The AEDT application remains open, and projects/windows are not affected.
    
    Returns:
        bool: True if successfully released, False if exception occurred
        
    Note:
        - AEDT application remains open after calling this function
        - Only COM references are released
        - Projects and windows remain open
        - To fully close AEDT, use Circuit.quitAedt() instead
    
    Examples:
        >>> releaseDesktop()  # Release COM resources
    '''
    module = sys.modules['__main__']
    try:
        # Release all COM objects and delete references
        print("Releasing COM resources...")
        _delete_objects()
        print("COM resources released. AEDT window control relinquished.")
        return True
        
    except Exception as e:
        print("Exception in releaseDesktop: %s" % str(e))
        try:
            _delete_objects()
        except Exception:
            pass
        return False


if __name__ == "__main__":
    oDesktop = getActiveDesktop()
    print(oDesktop.GetExeDir())
    pass
