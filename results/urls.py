# results/urls.py

from django.urls import path
from . import views

app_name = 'results'

urlpatterns = [
    path('upload/', views.upload_files, name='upload_files'),
    path('', views.student_list, name='student_list'),
    path('student/<str:student_id>/', views.student_result, name='student_result'),
]
