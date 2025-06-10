import flet as ft
from ui import MainView

class Router:
    def __init__(self, page: ft.Page):
        self.page = page
        self.setup_page()
    
    def setup_page(self):
        self.page.title = "Form 16 Generator"
        self.page.theme_mode = ft.ThemeMode.DARK
        self.page.auto_scroll = True
        self.page.window.width = 1200
        self.page.window.height = 800
        # self.page.padding = 20
        # self.page.expand = True
    
    def setup_main_route(self):
        main_view = MainView(self.page)
        self.page.add(main_view.build())
        self.page.update()