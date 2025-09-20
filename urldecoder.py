import ctypes
from ctypes import wintypes

# 加载 wininet.dll
wininet = ctypes.windll.wininet

# 定义常量
ICU_DECODE = 0x10000000  # 解码百分比转义序列
INTERNET_SCHEME_FTP = 2  # FTP 协议

# 定义 INTERNET_PORT 类型（WORD/unsigned short）
INTERNET_PORT = ctypes.c_ushort

# 定义 URL_COMPONENTS 结构
class URL_COMPONENTS(ctypes.Structure):
    _fields_ = [
        ("dwStructSize", wintypes.DWORD),
        ("lpszScheme", wintypes.LPWSTR),
        ("dwSchemeLength", wintypes.DWORD),
        ("nScheme", wintypes.DWORD),
        ("lpszHostName", wintypes.LPWSTR),
        ("dwHostNameLength", wintypes.DWORD),
        ("nPort", INTERNET_PORT),
        ("lpszUserName", wintypes.LPWSTR),
        ("dwUserNameLength", wintypes.DWORD),
        ("lpszPassword", wintypes.LPWSTR),
        ("dwPasswordLength", wintypes.DWORD),
        ("lpszUrlPath", wintypes.LPWSTR),
        ("dwUrlPathLength", wintypes.DWORD),
        ("lpszExtraInfo", wintypes.LPWSTR),
        ("dwExtraInfoLength", wintypes.DWORD),
    ]

# 初始化 InternetCrackUrl 函数原型
wininet.InternetCrackUrlW.argtypes = [
    wintypes.LPWSTR,        # lpszUrl
    wintypes.DWORD,         # dwUrlLength
    wintypes.DWORD,         # dwFlags
    ctypes.POINTER(URL_COMPONENTS)  # lpUrlComponents
]
wininet.InternetCrackUrlW.restype = wintypes.BOOL

def decode_ftp_url(encoded_url):
    # 初始化 URL_COMPONENTS 结构
    url_components = URL_COMPONENTS()
    url_components.dwStructSize = ctypes.sizeof(URL_COMPONENTS)
    
    # 为各个字段分配缓冲区
    buffer_size = 2048  # 缓冲区大小
    url_components.lpszScheme = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwSchemeLength = buffer_size
    url_components.lpszHostName = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwHostNameLength = buffer_size
    url_components.lpszUserName = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwUserNameLength = buffer_size
    url_components.lpszPassword = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwPasswordLength = buffer_size
    url_components.lpszUrlPath = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwUrlPathLength = buffer_size
    url_components.lpszExtraInfo = ctypes.cast(ctypes.create_unicode_buffer(buffer_size), wintypes.LPWSTR)
    url_components.dwExtraInfoLength = buffer_size

    # 调用 InternetCrackUrl 函数
    if not wininet.InternetCrackUrlW(encoded_url, 0, ICU_DECODE, ctypes.byref(url_components)):
        raise ctypes.WinError()

    # 提取解码后的组件
    decoded_url = {
        "Scheme": url_components.lpszScheme[:url_components.dwSchemeLength],
        "HostName": url_components.lpszHostName[:url_components.dwHostNameLength],
        "Port": url_components.nPort,
        "UserName": url_components.lpszUserName[:url_components.dwUserNameLength],
        "Password": url_components.lpszPassword[:url_components.dwPasswordLength],
        "UrlPath": url_components.lpszUrlPath[:url_components.dwUrlPathLength],
        "ExtraInfo": url_components.lpszExtraInfo[:url_components.dwExtraInfoLength]
    }

    return decoded_url

# 示例使用
if __name__ == "__main__":
    # 示例编码的 FTP URL
    encoded_url = "ftp://%75%73%65%72%3A%70%61%73%73%40%66%74%70%2E%65%78%61%6D%70%6C%65%2E%63%6F%6D/%70%61%74%68%2F%74%6F%2F%66%69%6C%65%2E%74%78%74"
    
    try:
        decoded = decode_ftp_url(encoded_url)
        print(f"Encoded URL: {encoded_url}")
        print(f"Decoded URL: {decoded}")
    except Exception as e:
        print(f"Error: {e}")
        # 打印详细的错误信息
        error_code = ctypes.GetLastError()
        print(f"Error code: {error_code}")