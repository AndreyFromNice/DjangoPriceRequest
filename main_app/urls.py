from django.urls import path
from . import views

urlpatterns = [
    path('', views.home, name='home'),
    path('upload-excel/', views.upload_excel, name='upload_excel'),
    path('download-processed-excel/', views.download_processed_excel, name='download_processed_excel'),
    path('send-emails/', views.send_emails_view, name='send_emails'), # New URL
]