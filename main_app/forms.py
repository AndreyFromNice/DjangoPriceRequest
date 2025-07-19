from django import forms

class ExcelUploadForm(forms.Form):
    excel_file = forms.FileField(label='Выберите Excel файл', help_text='Поддерживаются форматы .xlsx')

class CommercialRequestForm(forms.Form):
    excel_file = forms.FileField(label='Excel-файл с данными', required=True)
    company_name = forms.CharField(label='Название компании', initial='[НАЗВАНИЕ КОМПАНИИ]', required=False)
    project_name = forms.CharField(label='Название проекта', initial='[НАЗВАНИЕ ПРОЕКТА]', required=False)
    delivery_address = forms.CharField(label='Адрес доставки', required=False)
    estimate_section = forms.CharField(label='Раздел сметы', initial='[РАЗДЕЛ СМЕТЫ]', required=False)
    company_details = forms.CharField(label='Реквизиты компании', widget=forms.Textarea(attrs={'rows': 3}), initial='[РЕКВИЗИТЫ КОМПАНИИ]', required=False)
    contact_person = forms.CharField(label='Контактное лицо и телефон', initial='[КОНТАКТНОЕ ЛИЦО]', required=False)
    project_docs_link = forms.URLField(label='Ссылка на проектную документацию', required=False)

class MailingListUploadForm(forms.Form):
    csv_file = forms.FileField(label="Выберите CSV файл со списком рассылки", help_text="Загрузите файл CSV с колонкой 'Электронная почта'.")

class SmtpCredentialsForm(forms.Form):
    email_account = forms.EmailField(label='Ваш Email (логин SMTP и адрес отправителя)', required=True)

    smtp_password = forms.CharField(label='Пароль SMTP', widget=forms.PasswordInput, required=True)
    subject_prefix = forms.CharField(label='Префикс темы письма', max_length=100, required=False,
                                     initial='Коммерческий запрос')

