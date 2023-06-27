from kivy.app import App
from kivy.lang import Builder
from kivy.uix.boxlayout import BoxLayout
from kivy.properties import StringProperty
from kivy.uix.bubble import Button
from kivy.uix.textinput import TextInput

KV = """
HelperENGBL:
    orientation: "vertical"
    size_hint: (0.9, 0.9)
    pos_hint: {"center_x": 0.5, "center_y": 0.5}
    background: '#232946'
    
    Label:
        font_size: "30sp"
        multiline: True
        texst_size: self.width*0.98, None
        size_hint_x: 1.0
        size_hint_y: None
        height: self.texture_size[1] + 15
        text: root.data_label
        
    TextInput:
        id: Input
        multiline: False
        padding_y: (5,5)
        size_hint: (1, 0.15)
            
    Button:
        text: "Button 1"
        bold: True
        background_color: '#007bbc'
        size_hint: (1,0.2)
        on_press: root.check_inf1()
        
    Button:
        text: "Button 2"
        bold: True
        background_color: '#007bbc'
        size_hint: (1,0.2)
        on_press: root.check_inf2()
        
    Button:
        text: "Button 3"
        bold: True
        background_color: '#007bbc'
        size_hint: (1,0.2)
        on_press: root.check_inf3()
"""


class HelperENGBL(BoxLayout):
    data_label = StringProperty("Engineer assistant")


    def check_inf1(self):
        print("Button 1 check")

    def check_inf2(self):
        print("Button 2 check")

    def check_inf3(self):
        print("Button 3 check")




class EngAssistApp(App):
    running = True

    def build(self):
        return Builder.load_string(KV)

    def on_stop(self):
        self.running = False



if __name__ == "__main__":
    EngAssistApp().run()
