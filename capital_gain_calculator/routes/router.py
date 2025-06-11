import flet as ft
from ui import MainView

class Router:
    def __init__(self, page: ft.Page):
        self.page = page
        self.setup_page()
    
    def setup_page(self):
        self.page.title = "Capital Gain Calculator"
        self.page.theme_mode = ft.ThemeMode.LIGHT
        # self.page.window_width = 800
        # self.page.window_height = 600
        self.page.padding = 20
        self.page.bgcolor = ft.Colors.BLACK87
    
    def setup_main_route(self):
        main_view = MainView(self.page)
        self.page.add(main_view.build())
        self.page.update()