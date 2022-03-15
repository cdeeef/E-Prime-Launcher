import winreg as wg
import time
import logging
import ctypes,sys
import os


def get_location():
    try:
        logging.info('正在尝试访问E-Prime注册表')
        keyPath = wg.OpenKey(wg.HKEY_LOCAL_MACHINE,r"SOFTWARE\\Wow6432Node\\Psychology Software Tools\\E-Prime\\3.0\\Common")
        logging.info('成功打开注册表')
        value,tp = wg.QueryValueEx(keyPath,"RootFolder")
        logging.warning(f'找到程序路径：{value}')
    except Exception as e:
        logging.error('操作错误')
        logging.error(e)
    else:
        path = f'"{value}Program\\E-Studio.exe"'
        wg.CloseKey(keyPath)
        return path
    
    
def is_admin():
    try:
        return ctypes.windll.shell32.IsUserAnAdmin()
    except:
        return False

def set_count():
    try:
        logging.info('正在尝试访问E-Prime注册表')
        keyCount = wg.OpenKey(wg.HKEY_CURRENT_USER,r"Software\\Psychology Software Tools\\E-Prime\\3.0\\E-Studio\\Options",access=wg.KEY_ALL_ACCESS)
        logging.info('成功打开注册表')
        value,tp = wg.QueryValueEx(keyCount,"OpenCount")
        logging.warning(f'当前值为：{value}')
    except Exception as e:
        logging.error('操作错误')
        logging.error(e)
    else:
        try:
            logging.info('正在尝试修改打开次数')
            wg.SetValueEx(keyCount,'OpenCount','',wg.REG_DWORD,0)
        except WindowsError as e :
            logging.info(f'修改失败:{e}')
        except Exception as e:
            logging.error(f'操作错误{type(e)}')
            logging.error(e)
        else:
            logging.info('修改成功')
        finally:
            logging.info('关闭注册表')
            wg.CloseKey(keyCount)
    finally:
        logging.info('正在启动E-Sudio程序...')
        try:
            os.system(get_location())
            time.sleep(3)
        except Exception as e:
            logging.error(f'未能成功打开E-Prime客户端:{e}')
        


if __name__ == '__main__':
    print('hello')
    
    logging.basicConfig(level=logging.DEBUG, format="[%(levelname)s] - [%(asctime)s] - %(message)s")
    logging.info(f'程序启动')
    logging.info(f'Copyright SY Chen. Ver:2.0')
    
    if not is_admin():
        logging.warning('未在管理员身份运行')
        ctypes.windll.shell32.ShellExecuteW(None,"runas",sys.executable,__file__,None,1)
    else:
        logging.info('当前用户身份为管理员')

    

    set_count()


    
    


