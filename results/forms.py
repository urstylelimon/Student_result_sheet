# results/forms.py

from django import forms

class UploadFileForm(forms.Form):
    excel_file = forms.FileField(label='Upload Excel File')
    word_template = forms.FileField(label='Upload Word Template')
