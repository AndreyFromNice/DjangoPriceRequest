from django.db import models
from django.utils import timezone


class Project(models.Model):
    """
    Модель для хранения информации об общем проекте.
    """
    name = models.CharField(max_length=255, verbose_name="Название проекта")
    description = models.TextField(blank=True, verbose_name="Описание проекта")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Дата последнего обновления")

    class Meta:
        verbose_name = "Проект"
        verbose_name_plural = "Проекты"
        ordering = ['-created_at']

    def __str__(self):
        return self.name


class PriceRequest(models.Model):
    """
    Модель для хранения данных одного элемента коммерческого запроса,
    парсированных из Excel.
    """
    project = models.ForeignKey(Project, on_delete=models.CASCADE, related_name='price_requests', verbose_name="Проект")
    section = models.CharField(max_length=255, verbose_name="Раздел сметы")
    justification = models.CharField(max_length=255, verbose_name="Обоснование (ТЦ/ФССЦ)")
    item_number = models.CharField(max_length=50, blank=True, null=True, verbose_name="Номер по порядку")
    item_name = models.TextField(verbose_name="Наименование")
    unit = models.CharField(max_length=50, verbose_name="Ед. изм.")
    quantity = models.DecimalField(max_digits=10, decimal_places=2, verbose_name="Количество")
    price_per_unit = models.DecimalField(max_digits=15, decimal_places=2, verbose_name="Цена за ед.")
    total_price = models.DecimalField(max_digits=15, decimal_places=2, verbose_name="Всего")
    note = models.TextField(blank=True, verbose_name="Примечание")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания записи")

    class Meta:
        verbose_name = "Элемент коммерческого запроса"
        verbose_name_plural = "Элементы коммерческих запросов"
        ordering = ['project', 'section', 'item_number']

    def __str__(self):
        return f"{self.project.name} - {self.item_name}"


class Client(models.Model):
    """
    Модель для хранения информации о клиентах.
    """
    name = models.CharField(max_length=255, verbose_name="Имя клиента/Название компании")
    contact_person = models.CharField(max_length=255, blank=True, verbose_name="Контактное лицо")
    email = models.EmailField(unique=True, verbose_name="Email")
    phone = models.CharField(max_length=50, blank=True, verbose_name="Телефон")
    address = models.TextField(blank=True, verbose_name="Адрес")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата добавления")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Дата обновления")

    class Meta:
        verbose_name = "Клиент"
        verbose_name_plural = "Клиенты"
        ordering = ['name']

    def __str__(self):
        return self.name


class EmailTemplate(models.Model):
    """
    Модель для хранения шаблонов Email.
    """
    name = models.CharField(max_length=255, unique=True, verbose_name="Название шаблона")
    subject = models.CharField(max_length=255, verbose_name="Тема письма")
    body = models.TextField(verbose_name="Тело письма (HTML/Plain text)")
    created_at = models.DateTimeField(auto_now_add=True, verbose_name="Дата создания")
    updated_at = models.DateTimeField(auto_now=True, verbose_name="Дата последнего обновления")

    class Meta:
        verbose_name = "Шаблон Email"
        verbose_name_plural = "Шаблоны Email"
        ordering = ['name']

    def __str__(self):
        return self.name


class EmailCampaign(models.Model):
    """
    Модель для отслеживания массовых Email-рассылок.
    """
    name = models.CharField(max_length=255, verbose_name="Название рассылки")
    project = models.ForeignKey(Project, on_delete=models.CASCADE, related_name='email_campaigns',
                                verbose_name="Проект", null=True, blank=True)
    template = models.ForeignKey(EmailTemplate, on_delete=models.SET_NULL, null=True, blank=True,
                                 verbose_name="Использованный шаблон")
    recipients = models.ManyToManyField(Client, verbose_name="Получатели")
    sent_at = models.DateTimeField(default=timezone.now, verbose_name="Дата отправки")
    status = models.CharField(max_length=50, default='Draft', verbose_name="Статус", choices=[
        ('Draft', 'Черновик'),
        ('Sending', 'Отправляется'),
        ('Sent', 'Отправлено'),
        ('Failed', 'Ошибка')
    ])
    notes = models.TextField(blank=True, verbose_name="Заметки")

    # Поле для хранения пути к сгенерированному файлу Word, который был прикреплен
    generated_word_file = models.FileField(upload_to='generated_documents/', blank=True, null=True,
                                           verbose_name="Сгенерированный документ Word")
    # Поле для хранения пути к обработанному Excel файлу, который был прикреплен
    processed_excel_file = models.FileField(upload_to='processed_excels/', blank=True, null=True,
                                            verbose_name="Обработанный Excel файл")

    class Meta:
        verbose_name = "Email рассылка"
        verbose_name_plural = "Email рассылки"
        ordering = ['-sent_at']

    def __str__(self):
        return f"Рассылка '{self.name}' от {self.sent_at.strftime('%Y-%m-%d')}"