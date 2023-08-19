from kivy.uix.popup import Popup
from kivy.metrics import dp
import os
import win32com.client
from tkinter import filedialog
from tkinter import messagebox
from kivy.lang import Builder
from kivymd.app import MDApp
import win32ui
from kivymd.uix.button import MDRaisedButton
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.label import Label
from kivy.core.text import LabelBase
from kivy.core.window import Window

LabelBase.register(name="Montserrat", fn_regular="cust.ttf")
Window.size=(700,500)

class FileConverterApp(MDApp):
    def build(self):
        self.theme_cls.primary_palette = 'Blue'
        return Builder.load_file("main.kv")
             

    def file_manager_open(self):
        open_dialog = win32ui.CreateFileDialog(1)  # 1 stands for open file dialog
        open_dialog.SetOFNTitle("Select a File")
        open_dialog.DoModal()
        
        selected_file = open_dialog.GetPathName()
        if selected_file:
            self.selected_file = selected_file
            if self.selected_file.lower().endswith(".pdf"):
                self.root.ids.convert_button.text = "Convert to Word"
            elif self.selected_file.lower().endswith(".docx"):
                self.root.ids.convert_button.text = "Convert to PDF"
            else:
                self.root.ids.convert_button.text = "Convert"
    
    def convert_file(self):
        if hasattr(self, 'selected_file'):
            if self.selected_file.lower().endswith(".pdf"):
                self.convert_to_word()
            elif self.selected_file.lower().endswith(".docx"):
                self.convert_to_pdf()
            else:
                self.show_message("File Error", "Unsupported File Format")
        else:
            self.show_message("No file selected", "Please Select a file!!!")

    def convert_to_word(self):
        word = win32com.client.Dispatch("Word.Application")
        word.visible = 0

        selected_file = self.selected_file
        if selected_file.lower().endswith('.pdf'):
            filename = os.path.basename(selected_file)
            in_file = os.path.abspath(selected_file)

            try:
                wb = word.Documents.Open(in_file)
                output_file = filedialog.asksaveasfilename(
                    defaultextension=".docx",
                    filetypes=[("Word Documents", "*.docx")],
                    title="Save Converted Word Document"
                )
                if os.path.exists(output_file):
                        self.show_message("Conversion Error", "The file already exists. \n Save with a different name.")
                        return
                else:
                    wb.SaveAs2(output_file, FileFormat=16)
                self.show_message("Conversion Success", f"File converted to Word\n Check {output_file}")
            except Exception as e:
                self.show_message("Conversion Failed", f"Conversion failed \n {e}")
            finally:
                wb.Close()

            word.Quit()
        else:
            self.show_message("Invalid File", "Selected file is not a PDF.")


    def convert_to_pdf(self):
        selected_file = self.selected_file
        if selected_file.lower().endswith(".docx"):
            word = win32com.client.Dispatch("Word.Application")
            word.visible = 0
            
            try:
                doc = word.Documents.Open(selected_file)
                output_file = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Documents", "*.pdf")],
                    title="Save Converted PDF Document"
                )

                if output_file:
                    wdFormatPDF = 17  # File format code for PDF
                    if os.path.exists(output_file):
                        self.show_message("Conversion Error", "The file already exists. \n Save with a different name.")
                        return
                    doc.SaveAs(output_file, FileFormat=wdFormatPDF)
                    doc.Close()
                    word.Quit()
                    self.show_message("Conversion Success", f"File converted to PDF\n Check {output_file}")
            except Exception as e:
                print(e)
                self.show_message("Conversion Error", "Conversion failed")
                word.Quit()

    
    def show_message(self, title, content):
        # Create the popup content using KivyMD widgets
        content_layout = BoxLayout(orientation='vertical', padding=dp(10))

        # Add a label with the success message
        message_label = Label(text=content, font_name='Montserrat')
        content_layout.add_widget(message_label)

        # Add a button to close the popup
        ok_button = MDRaisedButton(text='OK', on_release=lambda x: popup.dismiss(), font_name='Montserrat')
        content_layout.add_widget(ok_button)

        # Create the popup
        popup = Popup(title=title, content=content_layout, size_hint=(None, None), size=(500, 300), title_font='Montserrat')

        # Open the popup
        popup.open()


if __name__ == '__main__':
    FileConverterApp().run()
