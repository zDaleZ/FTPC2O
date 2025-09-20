def make_it_winui(root):
    from sys import getwindowsversion
    from tkinter import ttk
    import sv_ttk
    import darkdetect
    import ctypes

    theme = darkdetect.theme()

    sv_ttk.set_theme(theme)

    version = getwindowsversion().build

    DWMW_USE_IMMERSIVE_DARK_MODE = 19 if version < 22000 else 20

    is_enable_dark = 1 if darkdetect.isDark() else 0

    def refresh_more():
        hwnd_id = ctypes.windll.user32.GetParent(root.winfo_id())

        bg_color = ttk.Style().lookup(".", "background")

        '''
        if version <= 22621:
            ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd_id, DWMW_USE_IMMERSIVE_DARK_MODE, ctypes.byref(ctypes.c_int(ThemeToSet)), ctypes.sizeof(ctypes.c_int))

            return
        '''

        ctypes.windll.dwmapi.DwmSetWindowAttribute(hwnd_id, DWMW_USE_IMMERSIVE_DARK_MODE, ctypes.byref(ctypes.c_int(is_enable_dark)), ctypes.sizeof(ctypes.c_int))
        return

        # Bug：点击会导致穿透
        '''
        if getwindowsversion().build >= 22621:
            from win32mica import MicaTheme, ApplyMica

            root.wm_attributes("-transparent", bg_color)
            root.update()
            if theme == "Dark":
                ApplyMica(
                    HWND=hwnd_id,
                    Theme=MicaTheme.DARK
                )
            else:
                ApplyMica(
                    HWND=hwnd_id,
                    Theme=MicaTheme.LIGHT
                )

      root.after(100, refresh_more)
      '''

    root.after(100, refresh_more)