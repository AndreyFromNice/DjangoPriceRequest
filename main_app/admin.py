from django.contrib import admin
from .models import Project, PriceRequest, Client, EmailTemplate, EmailCampaign

# Регистрация моделей в админ-панели
admin.site.register(Project)
admin.site.register(PriceRequest)
admin.site.register(Client)
admin.site.register(EmailTemplate)
admin.site.register(EmailCampaign)